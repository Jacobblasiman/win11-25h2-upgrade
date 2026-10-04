#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    In-place upgrade of Windows 11 23H2 (Enterprise/Education) to 25H2 from ISO media.

.DESCRIPTION
    Uses setup.exe from the Windows 11 25H2 Business Editions ISO to perform a full
    in-place upgrade (apps, data and settings kept). The upgrade runs from install media,
    so it does not depend on Windows Update, WSUS, the WU agent, the servicing stack or
    a healthy component store.

    Steps:
      1. Checks the OS build, architecture, and any setup that is already running
      2. Checks for a pending reboot (setup refuses to run when one is pending)
      3. Checks Windows 11 24H2+ hardware requirements (TPM 2.0, UEFI/Secure Boot, SSE4.2/POPCNT, RAM)
      4. Fixes problems it finds: removes stale upgrade folders, frees disk space,
         repairs the component store if it is flagged as corrupt, warns about known
         causes of rollback. Each slow fix is capped at 1 minute.
      5. Deletes any earlier local copy, copies the media fresh, mounts it, and checks that
         the media's language, architecture and build match this machine
      6. Suspends BitLocker for 3 reboots
      7. Starts setup.exe with no arguments (the interactive Windows Setup UI). The tech
         finishes the upgrade in the UI, and setup handles the restart.

    Run it in Windows PowerShell 5.1 (not PowerShell 7) as Administrator:
        powershell.exe -ExecutionPolicy Bypass -File .\Upgrade-Win11-25H2.ps1 -IsoPath "\\server\share\Win11_25H2_Business_x64.iso"

.PARAMETER IsoPath
    UNC or local path to the Windows 11 25H2 Business Editions ISO, OR a folder that contains
    extracted media (setup.exe at its root).

.PARAMETER ShareCredential
    Optional credential for the file share, for when the tech's account can't read it.

.PARAMETER WorkingDirectory
    Local folder for the media copy and logs. Default: C:\ProgramData\Win11Upgrade

.PARAMETER MinFreeSpaceGB
    Free space required on the system drive, not counting the local ISO copy. Default: 30

.PARAMETER DynamicUpdate
    Setup /DynamicUpdate mode for -CompatScanOnly. Default: Disable (no internet or WU dependency).
    'Enable' pulls the latest setup/compat fixes if the machine can reach Microsoft.

.PARAMETER DeepHealthScan
    Run DISM /ScanHealth instead of the quick /CheckHealth (capped at 1 min, like the other step 4 checks).

.PARAMETER CompatScanOnly
    Run only setup's compatibility scan (/Compat ScanOnly), report the result, and exit. Changes nothing.

.PARAMETER Force
    Skip the tech's confirmation prompts (warnings are still logged).

.NOTES
    Exit codes:
       0  Windows Setup was started and has closed, OR already on 25H2+, OR compat scan passed
       1  Unexpected error
      11  Unsupported architecture
      12  Hardware does not meet Windows 11 24H2+ requirements
      13  Not enough disk space after cleanup
      14  Reboot pending, reboot and run again
      15  Media problem (copy, mount, language, architecture or build mismatch)
      16  Compat scan found blockers (see logs)
      17  Canceled by the tech
      18  Windows Setup is already running
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = 'UNC/local path to the Windows 11 25H2 ISO or extracted media folder')]
    [string]$IsoPath,

    [System.Management.Automation.PSCredential]$ShareCredential,

    [string]$WorkingDirectory = (Join-Path $env:ProgramData 'Win11Upgrade'),

    [ValidateRange(20, 500)]
    [int]$MinFreeSpaceGB = 30,

    [ValidateSet('Disable', 'Enable', 'NoDrivers', 'NoLCU', 'NoDriversNoLCU')]
    [string]$DynamicUpdate = 'Disable',

    [switch]$DeepHealthScan,
    [switch]$CompatScanOnly,
    [switch]$Force
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'Continue'

#region ---------- Constants / state ----------
$TargetMinBuild   = 26200          # 25H2 = 26200, 24H2 = 26100, 23H2 = 22631
$ExpectedSource   = 22631          # 23H2

$SysDrive   = $env:SystemDrive
$BTFolder   = Join-Path $SysDrive '$WINDOWS.~BT'
$WSFolder   = Join-Path $SysDrive '$Windows.~WS'
$MediaDir   = Join-Path $WorkingDirectory 'Media'
$LogDir     = Join-Path $WorkingDirectory 'Logs'
$Stamp      = Get-Date -Format 'yyyyMMdd-HHmmss'
$LogFile    = Join-Path $LogDir "Upgrade-$Stamp.log"

$script:MountedIso      = $null
$script:MappedDrive     = $null
$script:BitLockerPaused = $false
$script:TranscriptOn    = $false
#endregion

#region ---------- Helpers ----------
function Invoke-Native {
    # Runs a native exe without PS 5.1 turning stderr output into terminating errors (EAP=Stop)
    param([Parameter(Mandatory)][string]$FilePath, [string[]]$ArgumentList = @())
    $old = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $out = $null; $code = -1
    try {
        $out  = & $FilePath @ArgumentList 2>&1 | ForEach-Object { "$_" }
        $code = $LASTEXITCODE
    } catch {
        $out = "Failed to run ${FilePath}: $($_.Exception.Message)"
    } finally { $ErrorActionPreference = $old }
    [pscustomobject]@{ ExitCode = $code; Output = ($out -join "`n") }
}

function Invoke-NativeWithTimeout {
    # Runs a native exe with a time limit, so a hung DISM/SFC/takeown can't stall the script.
    # Shows elapsed time while it runs; on timeout the process tree is killed and TimedOut is set.
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string]$Arguments,
        [Parameter(Mandatory)][int]$TimeoutMinutes,
        [string]$Activity = $FilePath
    )
    $p = Start-Process -FilePath $FilePath -ArgumentList $Arguments -PassThru -WindowStyle Hidden
    $null = $p.Handle   # needed so ExitCode is available after the process exits
    $sw = [Diagnostics.Stopwatch]::StartNew()
    try {
        while (-not $p.WaitForExit(2000)) {
            if ($sw.Elapsed.TotalMinutes -ge $TimeoutMinutes) {
                Invoke-Native taskkill.exe @('/PID', "$($p.Id)", '/T', '/F') | Out-Null
                Write-Log "$Activity did not finish within $TimeoutMinutes minute(s). Stopped it and continuing." 'WARN'
                return [pscustomobject]@{ ExitCode = $null; TimedOut = $true }
            }
            Write-Progress -Id 5 -Activity $Activity -Status ('Running, {0:hh\:mm\:ss} elapsed (limit {1} min)' -f $sw.Elapsed, $TimeoutMinutes)
        }
        return [pscustomobject]@{ ExitCode = $p.ExitCode; TimedOut = $false }
    } finally { Write-Progress -Id 5 -Activity $Activity -Completed }
}

function Get-ComponentStoreState {
    # Runs Repair-WindowsImage in a background job so it can be time-limited (RestoreHealth can sit
    # for a very long time waiting on WU/WSUS or another servicing operation). Returns the
    # ImageHealthState as a string, or $null on timeout. Errors from the cmdlet are rethrown.
    param(
        [Parameter(Mandatory)][ValidateSet('CheckHealth', 'ScanHealth', 'RestoreHealth')][string]$Mode,
        [Parameter(Mandatory)][int]$TimeoutMinutes
    )
    $job = Start-Job -ArgumentList $Mode -ScriptBlock {
        param($m)
        $a = @{ Online = $true; NoRestart = $true; ErrorAction = 'Stop' }
        $a[$m] = $true
        "$((Repair-WindowsImage @a).ImageHealthState)"
    }
    $sw = [Diagnostics.Stopwatch]::StartNew()
    try {
        while (-not (Wait-Job -Job $job -Timeout 2)) {
            if ($sw.Elapsed.TotalMinutes -ge $TimeoutMinutes) {
                Stop-Job -Job $job
                Get-Process -Name 'DismHost' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                Write-Log "DISM $Mode did not finish within $TimeoutMinutes minute(s). Stopped it and continuing." 'WARN'
                return $null
            }
            Write-Progress -Id 5 -Activity "DISM $Mode" -Status ('Running, {0:hh\:mm\:ss} elapsed (limit {1} min)' -f $sw.Elapsed, $TimeoutMinutes)
        }
        return (Receive-Job -Job $job -ErrorAction Stop | Select-Object -Last 1)
    } finally {
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        Write-Progress -Id 5 -Activity "DISM $Mode" -Completed
    }
}

function Copy-FileWithProgress {
    # Chunked copy with a progress bar (MB/s, ETA). Always starts fresh (any existing destination
    # is deleted); within a run, network errors are retried by reopening at the last good offset.
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$Destination,
        [string]$Activity = 'Copying install media',
        [int]$ProgressId = 1,
        [int]$ParentId = -1,
        [int]$MaxRetries = 10
    )
    $total  = (Get-Item -LiteralPath $Source).Length
    $buffer = New-Object byte[] (4MB)
    $done   = 0L
    if (Test-Path -LiteralPath $Destination) { Remove-Item -LiteralPath $Destination -Force }
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $startBytes = $done
    $lastUpdate = 0
    $retries = 0
    while ($done -lt $total) {
        $in = $null; $out = $null
        try {
            $in  = [IO.File]::Open($Source, 'Open', 'Read', 'Read')
            $out = [IO.File]::Open($Destination, 'OpenOrCreate', 'Write', 'None')
            $out.SetLength($done)   # drop anything past the last good offset
            [void]$in.Seek($done, 'Begin'); [void]$out.Seek($done, 'Begin')
            while (($read = $in.Read($buffer, 0, $buffer.Length)) -gt 0) {
                $out.Write($buffer, 0, $read)
                $done += $read
                $retries = 0
                if ($sw.ElapsedMilliseconds - $lastUpdate -ge 500 -or $done -eq $total) {
                    $lastUpdate = $sw.ElapsedMilliseconds
                    $secs = [math]::Max($sw.Elapsed.TotalSeconds, 0.1)
                    $rate = ($done - $startBytes) / $secs
                    $eta  = if ($rate -gt 0) { [TimeSpan]::FromSeconds([math]::Round(($total - $done) / $rate)) } else { [TimeSpan]::Zero }
                    $p = @{
                        Id               = $ProgressId
                        Activity         = $Activity
                        Status           = ('{0:N0} / {1:N0} MB  ({2:N1} MB/s, ETA {3:hh\:mm\:ss})' -f ($done/1MB), ($total/1MB), ($rate/1MB), $eta)
                        PercentComplete  = [int](($done / [double]$total) * 100)
                        CurrentOperation = (Split-Path $Source -Leaf)
                    }
                    if ($ParentId -ge 0) { $p.ParentId = $ParentId }
                    Write-Progress @p
                }
            }
        } catch {
            $retries++
            if ($retries -gt $MaxRetries) { throw "Copy failed after $MaxRetries retries: $($_.Exception.Message)" }
            Write-Log "Copy interrupted at $([math]::Round($done/1MB)) MB ($($_.Exception.Message)). Retry $retries/$MaxRetries in 10s..." 'WARN'
            Start-Sleep -Seconds 10
        } finally {
            if ($out) { $out.Flush(); $out.Dispose() }
            if ($in)  { $in.Dispose() }
        }
    }
    Write-Progress -Id $ProgressId -Activity $Activity -Completed
}

function Copy-FolderWithProgress {
    param([Parameter(Mandatory)][string]$Source, [Parameter(Mandatory)][string]$Destination)
    $files = @(Get-ChildItem -LiteralPath $Source -Recurse -File -Force)
    $total = ($files | Measure-Object -Property Length -Sum).Sum
    $done  = 0L
    $i = 0
    foreach ($f in $files) {
        $i++
        $rel  = $f.FullName.Substring((Resolve-Path -LiteralPath $Source).ProviderPath.TrimEnd('\','/').Length).TrimStart('\','/')
        $dest = Join-Path $Destination $rel
        New-Item -ItemType Directory -Path (Split-Path $dest -Parent) -Force | Out-Null
        Write-Progress -Id 1 -Activity 'Copying extracted media' -Status ('File {0} of {1}  ({2:N0} / {3:N0} MB)' -f $i, $files.Count, ($done/1MB), ($total/1MB)) -PercentComplete ([int](($done / [double][math]::Max($total,1)) * 100))
        if ($f.Length -gt 50MB) {
            Copy-FileWithProgress -Source $f.FullName -Destination $dest -Activity $rel -ProgressId 2 -ParentId 1
        } else {
            Copy-Item -LiteralPath $f.FullName -Destination $dest -Force
        }
        $done += $f.Length
    }
    Write-Progress -Id 1 -Activity 'Copying extracted media' -Completed
}

function Write-Log {
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Message,
        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR', 'STEP')][string]$Level = 'INFO'
    )
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level.PadRight(5), $Message
    $color = switch ($Level) { 'OK' { 'Green' } 'WARN' { 'Yellow' } 'ERROR' { 'Red' } 'STEP' { 'Cyan' } default { 'Gray' } }
    if ($Level -eq 'STEP') { Write-Host '' }
    Write-Host $line -ForegroundColor $color
    try { Add-Content -Path $LogFile -Value $line -Encoding UTF8 } catch { }
}

function Invoke-Cleanup {
    if ($script:MountedIso) {
        try { Dismount-DiskImage -ImagePath $script:MountedIso -ErrorAction Stop | Out-Null; Write-Log "Dismounted $($script:MountedIso)" }
        catch { Write-Log "Could not dismount ISO: $($_.Exception.Message)" 'WARN' }
        $script:MountedIso = $null
    }
    if ($script:MappedDrive) {
        try { Remove-PSDrive -Name $script:MappedDrive -Force -ErrorAction Stop } catch { }
        $script:MappedDrive = $null
    }
    [Win11Upg.Native]::SetThreadExecutionState([uint32]'0x80000000') | Out-Null   # ES_CONTINUOUS: allow sleep again
}

function Exit-Script {
    param([int]$Code, [string]$Message)
    if ($Message) {
        $lvl = 'ERROR'
        if ($Code -eq 0) { $lvl = 'OK' }
        Write-Log $Message $lvl
    }
    if ($Code -ne 0 -and $script:BitLockerPaused) { Resume-OSBitLocker }
    Invoke-Cleanup
    Write-Log "Script finished with exit code $Code. Log: $LogFile"
    if ($script:TranscriptOn) { try { Stop-Transcript | Out-Null } catch { } }
    exit $Code
}

function Confirm-Continue {
    param([string]$Prompt)
    if ($Force) { Write-Log "(-Force) auto-continuing: $Prompt" 'WARN'; return $true }
    do { $a = Read-Host "$Prompt [Y/N]" } until ($a -match '^[YyNn]')
    return ($a -match '^[Yy]')
}

function ConvertTo-HexCode([int]$Code) { '0x{0:X8}' -f $Code }

function Get-FreeGB {
    [math]::Round(((Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$SysDrive'").FreeSpace / 1GB), 1)
}

function Get-PendingRebootReasons {
    # Only the states that block setup. PendingFileRenameOperations is often set by agents
    # and doesn't block setup, so it's logged as info.
    $r = @()
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $r += 'CBS (component servicing) reboot pending' }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending') { $r += 'CBS packages pending' }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $r += 'Windows Update reboot required' }
    $cn  = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName').ComputerName
    $cnp = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName').ComputerName
    if ($cn -ne $cnp) { $r += 'Computer rename pending' }
    $pfro = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
    if ($pfro -and $pfro.PendingFileRenameOperations) { Write-Log 'Info: PendingFileRenameOperations present (usually harmless; not blocking).' }
    return $r
}

function Remove-FolderHard([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return }
    Write-Log "Removing $Path"
    # Plain delete first. Taking ownership of every file is slow on a large folder, so only do it if needed.
    Invoke-NativeWithTimeout cmd.exe "/c rd /s /q `"$Path`"" -TimeoutMinutes 1 -Activity "Removing $Path" | Out-Null
    if (Test-Path -LiteralPath $Path) {
        Write-Log 'Some files are protected. Taking ownership and retrying...'
        Invoke-NativeWithTimeout takeown.exe "/F `"$Path`" /R /A /D Y" -TimeoutMinutes 1 -Activity "Taking ownership of $Path" | Out-Null
        Invoke-NativeWithTimeout icacls.exe "`"$Path`" /grant *S-1-5-32-544:F /T /C /Q" -TimeoutMinutes 1 -Activity "Granting access to $Path" | Out-Null
        Invoke-NativeWithTimeout cmd.exe "/c rd /s /q `"$Path`"" -TimeoutMinutes 1 -Activity "Removing $Path" | Out-Null
    }
    if (Test-Path -LiteralPath $Path) { Write-Log "Could not fully remove $Path (files in use?)" 'WARN' } else { Write-Log "Removed $Path" 'OK' }
}

function Suspend-OSBitLocker {
    try {
        $v = Get-BitLockerVolume -MountPoint $SysDrive -ErrorAction Stop
        if ($v.ProtectionStatus -eq 'On') {
            Suspend-BitLocker -MountPoint $SysDrive -RebootCount 3 -ErrorAction Stop | Out-Null
            $script:BitLockerPaused = $true
            Write-Log "BitLocker suspended on $SysDrive for 3 reboots" 'OK'
        } else {
            Write-Log "BitLocker protection is not on for $SysDrive (status: $($v.ProtectionStatus)). Nothing to suspend."
        }
    } catch {
        Write-Log "BitLocker cmdlets unavailable ($($_.Exception.Message)). Trying manage-bde." 'WARN'
        $st = (Invoke-Native manage-bde.exe @('-status', $SysDrive)).Output
        if ($st -match 'Protection On') {
            $mb = Invoke-Native manage-bde.exe @('-protectors', '-disable', $SysDrive, '-RebootCount', '3')
            if ($mb.ExitCode -eq 0) { $script:BitLockerPaused = $true; Write-Log 'BitLocker suspended with manage-bde (3 reboots)' 'OK' }
            else { Write-Log 'manage-bde could not suspend BitLocker. Setup /BitLocker AlwaysSuspend will still try.' 'WARN' }
        } else { Write-Log 'BitLocker protection not on (manage-bde).' }
    }
}

function Resume-OSBitLocker {
    try { Resume-BitLocker -MountPoint $SysDrive -ErrorAction Stop | Out-Null; Write-Log "BitLocker protection resumed on $SysDrive" 'OK' }
    catch { Invoke-Native manage-bde.exe @('-protectors', '-enable', $SysDrive) | Out-Null; Write-Log 'Tried to resume BitLocker with manage-bde. Verify with: manage-bde -status' 'WARN' }
    $script:BitLockerPaused = $false
}

function Get-SetupCodeMeaning([string]$Hex) {
    switch ($Hex) {
        '0x00000000' { 'Success' }
        '0x00000003' { 'Success, reboot required' }
        '0xC1900210' { 'Compat scan passed: no compatibility issues found' }
        '0xC1900208' { 'Blocked by an incompatible app or driver (see the CompatData XML hard blocks listed below)' }
        '0xC1900204' { 'Migration choice not available: edition, language or architecture of the media does not match this OS' }
        '0xC1900200' { 'Machine does not meet the minimum Windows 11 hardware requirements' }
        '0xC190020E' { 'Not enough free disk space' }
        '0xC190010E' { 'EULA not accepted (/EULA accept missing)' }
        '0xC1900215' { 'No matching install image in the media for this edition' }
        '0xC1900107' { 'Cleanup from a previous install attempt is pending. Reboot and run again.' }
        '0xC1900101' { 'Driver-related failure caused rollback. Update or remove problem drivers (storage, AV/EDR filter, network, display).' }
        '0x80070070' { 'Not enough disk space' }
        '0x800F0922' { 'System Reserved/EFI partition too small or unreachable' }
        default      { 'Unknown or uncommon code. Check setuperr.log and the SetupDiag results.' }
    }
}

function Get-CompatHardBlocks {
    $panther = Join-Path $BTFolder 'Sources\Panther'
    $files = Get-ChildItem -Path $panther -Filter 'CompatData*.xml' -ErrorAction SilentlyContinue
    foreach ($f in $files) {
        try {
            [xml]$x = Get-Content -LiteralPath $f.FullName -Raw
            $nodes = $x.SelectNodes("//*[*[local-name()='CompatibilityInfo' and @BlockingType='Hard']]")
            foreach ($n in $nodes) {
                $name = $n.GetAttribute('Name'); if (-not $name) { $name = $n.GetAttribute('Id') }
                Write-Log "HARD BLOCK [$($n.LocalName)]: $name" 'ERROR'
            }
        } catch { }
    }
}

function Save-SetupLogs {
    $dest = Join-Path $LogDir "SetupLogs-$Stamp"
    New-Item -ItemType Directory -Path $dest -Force | Out-Null
    $sources = @(
        (Join-Path $BTFolder 'Sources\Panther'),
        (Join-Path $BTFolder 'Sources\Rollback'),
        (Join-Path $env:windir 'Panther')
    )
    foreach ($s in $sources) {
        if (Test-Path -LiteralPath $s) {
            $leaf = ($s -replace '[:\\\$~]', '_')
            Invoke-Native robocopy.exe @($s, (Join-Path $dest $leaf), '*.log', '*.xml', '*.txt', '/S', '/R:1', '/W:1', '/NP', '/NFL', '/NDL', '/NJH', '/NJS') | Out-Null
        }
    }
    Write-Log "Setup logs copied to $dest"

    # SetupDiag: Microsoft's tool that analyzes failed upgrades (best effort, needs internet)
    $sd = Join-Path $WorkingDirectory 'SetupDiag.exe'
    try {
        if (-not (Test-Path $sd)) {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri 'https://go.microsoft.com/fwlink/?linkid=870142' -OutFile $sd -UseBasicParsing -TimeoutSec 60
        }
        $sdOut = Join-Path $dest 'SetupDiagResults.log'
        Write-Log 'Running SetupDiag (this can take a few minutes)...'
        Start-Process -FilePath $sd -ArgumentList "/Output:`"$sdOut`"" -Wait -WindowStyle Hidden
        if (Test-Path $sdOut) {
            Write-Log "SetupDiag results: $sdOut" 'OK'
            Get-Content $sdOut -TotalCount 40 | ForEach-Object { Write-Log "  $_" }
        }
    } catch { Write-Log "SetupDiag not run: $($_.Exception.Message)" 'WARN' }
}

Add-Type -Namespace Win11Upg -Name Native -MemberDefinition @'
[DllImport("kernel32.dll")] public static extern bool IsProcessorFeaturePresent(uint feature);
[DllImport("kernel32.dll")] public static extern uint SetThreadExecutionState(uint esFlags);
'@
#endregion

#region ---------- Init ----------
foreach ($d in @($WorkingDirectory, $MediaDir, $LogDir)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
try { Start-Transcript -Path (Join-Path $LogDir "Transcript-$Stamp.txt") -Force | Out-Null; $script:TranscriptOn = $true } catch { }

# Keep the machine awake while the script runs (ES_CONTINUOUS | ES_SYSTEM_REQUIRED)
[Win11Upg.Native]::SetThreadExecutionState([uint32]'0x80000001') | Out-Null

trap {
    Write-Log "Unhandled error: $($_.Exception.Message) (line $($_.InvocationInfo.ScriptLineNumber))" 'ERROR'
    Exit-Script 1
}

if ($PSVersionTable.PSEdition -eq 'Core') {
    Write-Log 'Running in PowerShell 7. Windows PowerShell 5.1 is recommended (DISM/BitLocker modules).' 'WARN'
}
Write-Log "Windows 11 feature upgrade script started on $env:COMPUTERNAME by $env:USERDOMAIN\$env:USERNAME" 'STEP'
Write-Log "Log file: $LogFile"
#endregion

#region ---------- 1. OS / edition / architecture ----------
Write-Log '1/7  Checking current OS' 'STEP'
$cv       = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
$curBuild = [int]$cv.CurrentBuildNumber
$curUBR   = $cv.UBR
$curVer   = $cv.DisplayVersion
$edition  = $cv.EditionID
$osArch   = $env:PROCESSOR_ARCHITECTURE
if ($env:PROCESSOR_ARCHITEW6432) { $osArch = $env:PROCESSOR_ARCHITEW6432 }
Write-Log "OS: build $curBuild.$curUBR ($curVer), edition $edition, arch $osArch"

if ($curBuild -ge $TargetMinBuild) { Exit-Script 0 "Already on build $curBuild (25H2 or later). Nothing to do." }
if ($curBuild -lt 22000) {
    Write-Log "This is Windows 10 (build $curBuild). The same method works, but this script was built for 23H2." 'WARN'
    if (-not (Confirm-Continue 'Continue anyway?')) { Exit-Script 17 'Canceled by tech.' }
} elseif ($curBuild -eq 26100) {
    Write-Log 'This machine is on 24H2. The 25H2 enablement package (KB5054156) is much faster than a full media upgrade.' 'WARN'
    if (-not (Confirm-Continue 'Do the full media upgrade anyway?')) { Exit-Script 17 'Canceled by tech.' }
} elseif ($curBuild -ne $ExpectedSource) {
    Write-Log "Source build $curBuild is not 23H2 ($ExpectedSource). A media upgrade is still supported." 'WARN'
}
if ($osArch -notin @('AMD64', 'ARM64')) { Exit-Script 11 "Unsupported architecture $osArch." }

$running = Get-Process -Name 'SetupHost', 'SetupPrep' -ErrorAction SilentlyContinue
if ($running) { Exit-Script 18 "Windows Setup is already running (PID $($running.Id -join ', ')). Let it finish or reboot first." }
#endregion

#region ---------- 2. Pending reboot ----------
Write-Log '2/7  Checking for pending reboot' 'STEP'
$pending = @(Get-PendingRebootReasons)
if ($pending.Count -gt 0) {
    $pending | ForEach-Object { Write-Log "Pending: $_" 'WARN' }
    Write-Log 'Setup will not run while a reboot is pending.' 'WARN'
    if (-not $Force -and (Confirm-Continue 'Reboot now? (Run this script again after the reboot)')) {
        Write-Log 'Rebooting at the tech''s request...' 'WARN'
        Invoke-Cleanup
        if ($script:TranscriptOn) { try { Stop-Transcript | Out-Null } catch { } }
        Restart-Computer -Force
        exit 14
    }
    Exit-Script 14 'Reboot pending. Reboot and run the script again.'
}
Write-Log 'No pending reboot.' 'OK'
#endregion

#region ---------- 3. Hardware ----------
Write-Log '3/7  Checking Windows 11 24H2+ hardware requirements' 'STEP'
$hwFail = @()

# TPM 2.0
try {
    $tpm = Get-CimInstance -Namespace 'root/cimv2/Security/MicrosoftTpm' -ClassName Win32_Tpm -ErrorAction Stop
    if (-not $tpm) { $hwFail += 'No TPM detected' }
    else {
        $spec = ($tpm.SpecVersion -split ',')[0].Trim()
        Write-Log "TPM spec $spec, enabled=$($tpm.IsEnabled_InitialValue), activated=$($tpm.IsActivated_InitialValue)"
        if ([version]$spec -lt [version]'2.0') { $hwFail += "TPM version $spec (2.0 required)" }
        if (-not $tpm.IsEnabled_InitialValue) { $hwFail += 'TPM is disabled in firmware' }
    }
} catch { $hwFail += "TPM query failed: $($_.Exception.Message)" }

# UEFI + Secure Boot capable
try {
    $sb = Confirm-SecureBootUEFI -ErrorAction Stop
    Write-Log "UEFI firmware: yes. Secure Boot enabled: $sb"
    if (-not $sb) { Write-Log 'Secure Boot is off. Windows 11 needs it to be supported, not necessarily enabled. Consider enabling it.' 'WARN' }
} catch [System.PlatformNotSupportedException] {
    $hwFail += 'Legacy BIOS boot (UEFI required)'
} catch {
    Write-Log "Secure Boot state unknown: $($_.Exception.Message)" 'WARN'
}

# CPU: 24H2+ requires SSE4.2 and POPCNT (PF_SSE4_2_INSTRUCTIONS_AVAILABLE = 38)
if ($osArch -eq 'AMD64') {
    if (-not [Win11Upg.Native]::IsProcessorFeaturePresent(38)) { $hwFail += 'CPU lacks SSE4.2/POPCNT (required from 24H2)' }
    else { Write-Log 'CPU supports SSE4.2/POPCNT.' }
}
$cpu = (Get-CimInstance Win32_Processor | Select-Object -First 1).Name
Write-Log "CPU: $cpu"

# RAM
$ramGB = [math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 1)
Write-Log "RAM: $ramGB GB"
if ($ramGB -lt 3.8) { $hwFail += "RAM $ramGB GB (4 GB required)" }

if ($hwFail.Count -gt 0) {
    $hwFail | ForEach-Object { Write-Log "HW FAIL: $_" 'ERROR' }
    Exit-Script 12 'This machine does not meet Windows 11 24H2+ hardware requirements.'
}
Write-Log 'Hardware requirements met.' 'OK'

# Power (laptops)
try {
    Add-Type -AssemblyName System.Windows.Forms
    if ([System.Windows.Forms.SystemInformation]::PowerStatus.PowerLineStatus -eq 'Offline') {
        Write-Log 'Running on battery. Plug in AC power before upgrading.' 'WARN'
        if (-not (Confirm-Continue 'Continue on battery?')) { Exit-Script 17 'Canceled by tech (on battery).' }
    }
} catch { }
#endregion

#region ---------- 4. Pre-flight remediation ----------
Write-Log '4/7  Pre-flight checks and fixes' 'STEP'

# 4a. Stale upgrade folders from earlier failed attempts are a common cause of failure. Old media
#     copies are deleted too; step 5 always copies the media fresh.
foreach ($f in @($BTFolder, $WSFolder)) { Remove-FolderHard $f }
Get-ChildItem -LiteralPath $MediaDir -Filter '*.iso' -File -ErrorAction SilentlyContinue | ForEach-Object {
    try { Dismount-DiskImage -ImagePath $_.FullName -ErrorAction Stop | Out-Null } catch { }   # left mounted by an earlier run
}
foreach ($old in @(Get-ChildItem -LiteralPath $MediaDir -Force -ErrorAction SilentlyContinue)) {
    Write-Log "Removing old media copy $($old.FullName)"
    try { Remove-Item -LiteralPath $old.FullName -Recurse -Force -ErrorAction Stop }
    catch { Exit-Script 15 "Could not delete old media copy $($old.FullName): $($_.Exception.Message)" }
}

# 4b. Disk space
$isoSizeGB = 8   # assumed size of the local media copy (ISO or extracted folder)
if ($IsoPath -match '\.iso$' -and (Test-Path -LiteralPath $IsoPath -ErrorAction SilentlyContinue)) {
    $isoSizeGB = [math]::Ceiling((Get-Item -LiteralPath $IsoPath).Length / 1GB)
}
$needGB = $MinFreeSpaceGB + $isoSizeGB
$freeGB = Get-FreeGB
Write-Log "Free space on ${SysDrive}: $freeGB GB (need $needGB GB = $MinFreeSpaceGB for setup + $isoSizeGB for the ISO copy)"

if ($freeGB -lt $needGB) {
    Write-Log 'Low disk space. Running cleanup...' 'WARN'
    # Temp folders
    foreach ($t in @("$env:windir\Temp", $env:TEMP)) {
        Get-ChildItem -LiteralPath $t -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }
    # Windows Update download cache
    foreach ($svc in 'wuauserv', 'bits') {
        # Stop-Service -Force can wait forever on a service that won't stop
        try {
            Stop-Service $svc -Force -NoWait -ErrorAction Stop
            (Get-Service $svc).WaitForStatus('Stopped', [TimeSpan]::FromSeconds(60))
        } catch { Write-Log "Could not stop $svc within 60s. Its cache may only be partly cleared." 'WARN' }
    }
    Get-ChildItem "$env:windir\SoftwareDistribution\Download" -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    foreach ($svc in 'bits', 'wuauserv') { Start-Service $svc -ErrorAction SilentlyContinue }
    # Delivery Optimization cache
    try { Delete-DeliveryOptimizationCache -Force -ErrorAction Stop | Out-Null } catch { }
    # Recycle bin
    try { Clear-RecycleBin -DriveLetter $SysDrive.TrimEnd(':') -Force -ErrorAction Stop } catch { }
    # WinSxS superseded components
    Write-Log 'Running DISM /StartComponentCleanup (limit 1 min)...'
    Invoke-NativeWithTimeout dism.exe '/Online /Cleanup-Image /StartComponentCleanup /Quiet' -TimeoutMinutes 1 -Activity 'DISM StartComponentCleanup' | Out-Null
    $freeGB = Get-FreeGB
    Write-Log "Free space after cleanup: $freeGB GB"
    if ($freeGB -lt $needGB) {
        if (Test-Path "$SysDrive\Windows.old") { Write-Log 'Windows.old exists and uses space. Remove it with Disk Cleanup if you no longer need a rollback.' 'WARN' }
        Exit-Script 13 "Not enough disk space ($freeGB GB free, $needGB GB needed). Free up space and run again."
    }
}
Write-Log 'Disk space OK.' 'OK'

# 4c. Component store health. Setup from media replaces the OS, so a failed repair only
#     produces a warning, but a quick repair helps avoid migration problems.
try {
    if ($DeepHealthScan) {
        Write-Log 'Running DISM ScanHealth (limit 1 min)...'
        $state = Get-ComponentStoreState -Mode ScanHealth -TimeoutMinutes 1
    } else {
        Write-Log 'Running DISM CheckHealth...'
        $state = Get-ComponentStoreState -Mode CheckHealth -TimeoutMinutes 1
    }
    if ($null -eq $state) { throw 'the health check timed out' }
    Write-Log "Component store state: $state"
    if ($state -ne 'Healthy') {
        Write-Log 'Component store is flagged. Running DISM RestoreHealth + SFC...' 'WARN'
        try {
            Write-Log 'Running DISM RestoreHealth (limit 1 min)...'
            $r = Get-ComponentStoreState -Mode RestoreHealth -TimeoutMinutes 1
            if ($null -ne $r) { Write-Log "After RestoreHealth: $r" }
        } catch { Write-Log "RestoreHealth failed ($($_.Exception.Message)). Continuing, since the media upgrade replaces the component store." 'WARN' }
        Write-Log 'Running SFC /scannow (limit 1 min)...'
        $sfc = Invoke-NativeWithTimeout sfc.exe '/scannow' -TimeoutMinutes 1 -Activity 'SFC /scannow'
        if (-not $sfc.TimedOut) { Write-Log "SFC exit code: $($sfc.ExitCode) (details in %windir%\Logs\CBS\CBS.log)" }
        $pending2 = @(Get-PendingRebootReasons)
        if ($pending2.Count -gt 0) { Exit-Script 14 'The repairs need a reboot. Reboot and run the script again.' }
    } else { Write-Log 'Component store healthy.' 'OK' }
} catch { Write-Log "Health check skipped: $($_.Exception.Message)" 'WARN' }

# 4d. Known folders redirected to another local drive are a known cause of 23H2 -> 24H2/25H2 rollbacks
$kfWarn = @()
$kfNames = @('Desktop', 'Personal', 'My Pictures', 'My Music', 'My Video', '{374DE290-123F-4565-9164-39C4925E467B}')
Get-ChildItem 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue |
    Where-Object { $_.PSChildName -match '^S-1-5-21-[\d-]+$' } | ForEach-Object {
        $k = "Registry::HKEY_USERS\$($_.PSChildName)\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
        $p = Get-ItemProperty -Path $k -ErrorAction SilentlyContinue
        if ($p) {
            foreach ($n in $kfNames) {
                $val = $p.PSObject.Properties[$n]
                if ($val -and $val.Value -match '^([A-Za-z]):' -and ($Matches[1] + ':') -ne $SysDrive) {
                    $kfWarn += "$($_.PSChildName): '$n' -> $($val.Value)"
                }
            }
        }
    }
if ($kfWarn.Count -gt 0) {
    $kfWarn | ForEach-Object { Write-Log "Known folder on another drive: $_" 'WARN' }
    Write-Log 'Redirecting known folders to a non-system local drive often causes rollback (0x800701DE). Consider pointing them back to C: before upgrading.' 'WARN'
    if (-not (Confirm-Continue 'Continue anyway?')) { Exit-Script 17 'Canceled by tech (known folder redirection).' }
}

# 4e. SentinelOne: left running, logged for troubleshooting
$s1 = Get-Service -Name 'SentinelAgent' -ErrorAction SilentlyContinue
if ($s1) {
    $s1Exe = Get-ChildItem "$env:ProgramFiles\SentinelOne\*\SentinelAgent.exe" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $s1Ver = 'unknown'; if ($s1Exe) { $s1Ver = $s1Exe.VersionInfo.ProductVersion }
    Write-Log "SentinelOne agent: status $($s1.Status), version $s1Ver (left running)"
}
#endregion

#region ---------- 5. Media ----------
Write-Log '5/7  Preparing install media' 'STEP'
$src = $IsoPath

if ($ShareCredential -and $IsoPath -like '\\*') {
    $shareRoot = $IsoPath
    if ($IsoPath -match '\.iso$') { $shareRoot = Split-Path $IsoPath -Parent }
    $free = [char[]](90..68) | Where-Object { -not (Test-Path "$($_):\") } | Select-Object -First 1
    New-PSDrive -Name $free -PSProvider FileSystem -Root $shareRoot -Credential $ShareCredential -Persist -ErrorAction Stop | Out-Null
    $script:MappedDrive = [string]$free
    if ($IsoPath -match '\.iso$') { $src = "$($free):\" + (Split-Path $IsoPath -Leaf) } else { $src = "$($free):\" }
    Write-Log "Mapped $shareRoot to $($free): with the supplied credential"
}

if (-not (Test-Path -LiteralPath $src)) { Exit-Script 15 "Media path not reachable: $IsoPath" }

$setupRoot = $null
if ((Get-Item -LiteralPath $src).PSIsContainer) {
    # Extracted media folder
    if (-not (Test-Path (Join-Path $src 'setup.exe'))) { Exit-Script 15 "No setup.exe in $src" }
    $localMedia = Join-Path $MediaDir 'Extracted'
    Write-Log "Copying extracted media to $localMedia ..."
    try { Copy-FolderWithProgress -Source $src -Destination $localMedia }
    catch { Exit-Script 15 "Copying media failed: $($_.Exception.Message)" }
    $setupRoot = $localMedia
} else {
    $isoName  = Split-Path $src -Leaf
    $localIso = Join-Path $MediaDir $isoName
    $srcSize  = (Get-Item -LiteralPath $src).Length
    Write-Log "Copying $isoName ($([math]::Round($srcSize/1GB,2)) GB) to $MediaDir ..."
    try { Copy-FileWithProgress -Source $src -Destination $localIso -Activity "Copying $isoName" }
    catch { Exit-Script 15 "ISO copy failed: $($_.Exception.Message). Run again." }
    if ((Get-Item -LiteralPath $localIso).Length -ne $srcSize) { Exit-Script 15 'ISO copy is incomplete. Run again.' }
    Write-Log 'ISO copied.' 'OK'

    # Mount; fall back to 7-Zip extraction if mounting is broken or blocked
    try {
        $img = Mount-DiskImage -ImagePath $localIso -StorageType ISO -PassThru -ErrorAction Stop
        $script:MountedIso = $localIso
        $letter = $null
        for ($i = 0; $i -lt 15 -and -not $letter; $i++) {
            Start-Sleep -Seconds 1
            $letter = ($img | Get-Volume -ErrorAction SilentlyContinue).DriveLetter
        }
        if (-not $letter) { throw 'The mounted ISO did not get a drive letter.' }
        $setupRoot = "$($letter):\"
        Write-Log "ISO mounted at $setupRoot" 'OK'
    } catch {
        Write-Log "ISO mount failed: $($_.Exception.Message). Trying 7-Zip extraction." 'WARN'
        if ($script:MountedIso) { try { Dismount-DiskImage -ImagePath $localIso | Out-Null } catch { }; $script:MountedIso = $null }
        $7z = @("$env:ProgramFiles\7-Zip\7z.exe", "${env:ProgramFiles(x86)}\7-Zip\7z.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
        if (-not $7z) { Exit-Script 15 'Cannot mount the ISO and 7-Zip is not installed. Run again with -IsoPath pointing to an extracted media folder.' }
        $extract = Join-Path $MediaDir 'Extracted'
        $z = Invoke-Native $7z @('x', $localIso, "-o$extract", '-y')
        if ($z.ExitCode -ne 0) { Exit-Script 15 "7-Zip extraction failed (code $($z.ExitCode))." }
        $setupRoot = $extract
    }
}

$setupExe = Join-Path $setupRoot 'setup.exe'
$wim = @('sources\install.wim', 'sources\install.esd') | ForEach-Object { Join-Path $setupRoot $_ } | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not (Test-Path $setupExe) -or -not $wim) { Exit-Script 15 'The media is missing setup.exe or sources\install.wim/esd.' }

# Check that the media matches this machine's edition, build, language and architecture
Write-Log 'Reading image metadata...'
$sysLang = [System.Globalization.CultureInfo]::InstalledUICulture.Name
$image = $null
foreach ($i in (Get-WindowsImage -ImagePath $wim)) {
    $d = Get-WindowsImage -ImagePath $wim -Index $i.ImageIndex
    Write-Log ("  Index {0}: {1} | EditionId {2} | {3} | {4}" -f $d.ImageIndex, $d.ImageName, $d.EditionId, $d.Version, (($d.Languages) -join ','))
    if ($d.EditionId -eq $edition -and -not $image) { $image = $d }
}
if (-not $image) {
    # Not fatal: setup does its own edition matching. Use the first image for the build/arch/language checks.
    $image = Get-WindowsImage -ImagePath $wim -Index 1
    Write-Log "No image in the media has EditionId '$edition'. Continuing; setup will pick the edition." 'WARN'
}

$mediaBuild = ([version]$image.Version).Build
$archRaw    = "$($image.Architecture)"
$mediaArch  = $null
if     ($archRaw -match '^(9|x64|amd64)$') { $mediaArch = 'AMD64' }
elseif ($archRaw -match '^(12|arm64)$')    { $mediaArch = 'ARM64' }
elseif ($archRaw -match '^(0|x86)$')       { $mediaArch = 'x86' }
$mediaLangs = @($image.Languages | ForEach-Object { ($_ -replace '\s*\(.*\)', '').Trim() })
Write-Log "Selected image: $($image.ImageName) | build $($image.Version) | arch $archRaw | languages $($mediaLangs -join ',')"

if ($mediaArch -and $mediaArch -ne $osArch) { Exit-Script 15 "Architecture mismatch: media $mediaArch, OS $osArch." }
if ($mediaBuild -le $curBuild) { Exit-Script 15 "The media build ($mediaBuild) is not newer than the current build ($curBuild)." }
if ($mediaBuild -lt $TargetMinBuild) {
    Write-Log "The media build $mediaBuild is older than 25H2 ($TargetMinBuild). Is this a 24H2 ISO?" 'WARN'
    if (-not (Confirm-Continue 'Upgrade to this build anyway?')) { Exit-Script 17 'Canceled by tech (media older than 25H2).' }
}
if ($mediaLangs -notcontains $sysLang) {
    Exit-Script 15 "Language mismatch: OS default UI language is $sysLang, media has $($mediaLangs -join ','). /auto upgrade requires the same language. Get the $sysLang ISO."
}
Write-Log 'Media matches this machine.' 'OK'
#endregion

#region ---------- 6. Compat scan only (optional) ----------
$setupLogCopy = Join-Path $LogDir "SetupCopyLogs-$Stamp"
if ($CompatScanOnly) {
    Write-Log '6/7  Running compatibility scan only (no changes)' 'STEP'
    $scanArgs = "/Auto Upgrade /Quiet /EULA Accept /Compat ScanOnly /DynamicUpdate $DynamicUpdate /Telemetry Disable /CopyLogs `"$setupLogCopy`""
    Write-Log "setup.exe $scanArgs"
    Write-Log 'The scan usually takes 5-20 minutes...'
    $p = Start-Process -FilePath $setupExe -ArgumentList $scanArgs -PassThru -Wait
    $hex = ConvertTo-HexCode $p.ExitCode
    Write-Log "Compat scan result $hex : $(Get-SetupCodeMeaning $hex)"
    if ($hex -eq '0xC1900210') { Exit-Script 0 'Compatibility scan PASSED. Run again without -CompatScanOnly to upgrade.' }
    Get-CompatHardBlocks
    Save-SetupLogs
    Exit-Script 16 "Compatibility scan found issues ($hex)."
}
#endregion

#region ---------- 7. Run the upgrade ----------
Write-Log '6/7  Ready to upgrade' 'STEP'
Write-Log "  $env:COMPUTERNAME : build $curBuild ($curVer) $edition  ->  build $($image.Version) $($image.ImageName)"
Write-Log '  Windows Setup opens interactively. Finish the upgrade in the Setup window; setup restarts the computer itself.'
if (-not (Confirm-Continue 'Start Windows Setup now?')) { Exit-Script 17 'Canceled by tech before setup started.' }

Write-Log '7/7  Suspending BitLocker and starting Windows Setup' 'STEP'
Suspend-OSBitLocker
Write-Log "Starting $setupExe (interactive). Leave this window open until Setup finishes."
$proc = Start-Process -FilePath $setupExe -PassThru

# setup.exe hands off to SetupHost.exe; keep the media mounted until both have closed
while (-not $proc.HasExited -or (Get-Process -Name 'SetupHost' -ErrorAction SilentlyContinue)) { Start-Sleep -Seconds 5 }
Write-Log "Windows Setup closed. If the upgrade was canceled, BitLocker protection resumes by itself after 3 reboots, or run: manage-bde -protectors -enable $SysDrive"
Exit-Script 0 'Done.'
#endregion
