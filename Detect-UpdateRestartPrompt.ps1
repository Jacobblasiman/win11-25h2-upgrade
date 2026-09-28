<#
.SYNOPSIS
    Intune Remediation - DETECTION script.
    Is Windows Update waiting on a restart that the signed-in user hasn't planned for yet?

.DESCRIPTION
    Pairs with Remediate-UpdateRestartPrompt.ps1, which shows the "pick a time to restart" dialog.

    Exit 0 = nothing to do: no update restart pending, the user already scheduled one,
             or the reminder is snoozed.
    Exit 1 = a restart is pending and the user has no plan yet, so Intune runs the remediation.

    Intune settings (Devices > Scripts and remediations > Create script package):
      Run this script using the logged-on credentials : Yes  (state lives in HKCU; the dialog needs the user's desktop)
      Enforce script signature check                  : No   (or Yes if you sign both scripts)
      Run script in 64-bit PowerShell                 : Yes
      Schedule                                        : Hourly, repeat every 1 hour
#>

# --- Keep in sync with the remediation script --------------------------------
$StateKey = 'HKCU:\Software\UpdateRestartPrompt'
$TaskName = "Update restart - $env:USERNAME"
# ------------------------------------------------------------------------------

function Test-UpdateRestartPending {
    if (Test-Path -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') {
        return $true
    }
    try {
        $systemInfo = New-Object -ComObject Microsoft.Update.SystemInfo
        return [bool]$systemInfo.RebootRequired
    } catch {
        return $false
    }
}

function Get-StateValue([string]$Name) {
    try { (Get-ItemProperty -LiteralPath $StateKey -Name $Name -ErrorAction Stop).$Name } catch { $null }
}

function Set-StateValue([string]$Name, [string]$Value) {
    if (-not (Test-Path -LiteralPath $StateKey)) { New-Item -Path $StateKey -Force | Out-Null }
    Set-ItemProperty -LiteralPath $StateKey -Name $Name -Value $Value
}

function Remove-StateValue([string]$Name) {
    Remove-ItemProperty -LiteralPath $StateKey -Name $Name -ErrorAction SilentlyContinue
}

function ConvertFrom-StateDate($Value) {
    $parsed = [datetime]::MinValue
    if ($Value -and [datetime]::TryParse([string]$Value, [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::RoundtripKind, [ref]$parsed)) {
        if ($parsed.Kind -eq [DateTimeKind]::Unspecified) { return [datetime]::SpecifyKind($parsed, [DateTimeKind]::Local) }
        return $parsed.ToLocalTime()
    }
    return $null
}

function Get-RestartTask {
    try {
        $service = New-Object -ComObject Schedule.Service
        $service.Connect()
        return $service.GetFolder('\').GetTask($TaskName)
    } catch {
        return $null
    }
}

function Remove-RestartTask {
    try {
        $service = New-Object -ComObject Schedule.Service
        $service.Connect()
        $service.GetFolder('\').DeleteTask($TaskName, 0)
    } catch { }
}

$now = Get-Date

if (-not (Test-UpdateRestartPending)) {
    # Nothing pending, or the restart already happened: clear leftovers so the next update starts clean.
    Remove-RestartTask
    Remove-Item -LiteralPath $StateKey -Recurse -Force -ErrorAction SilentlyContinue
    Write-Output 'OK: no update restart pending'
    exit 0
}

# Record when this pending restart was first seen (the remediation uses it to estimate the deadline).
try {
    $lastBoot = (Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop).LastBootUpTime
} catch {
    $lastBoot = [datetime]::MinValue
}
$firstSeen = ConvertFrom-StateDate (Get-StateValue 'PendingSince')
if (-not $firstSeen -or $firstSeen -lt $lastBoot) {
    Set-StateValue 'PendingSince' $now.ToString('o')
    # Choices made during an earlier restart cycle don't carry over to this one.
    Remove-StateValue 'SnoozeUntil'
    Remove-StateValue 'SnoozeReason'
    Remove-StateValue 'ScheduledFor'
}

# Has the user already scheduled a restart? The task fires a few minutes before the chosen
# time to start the warning countdown, so keep treating it as planned until 30 minutes after.
$task = Get-RestartTask
if ($task) {
    $scheduledFor = ConvertFrom-StateDate (Get-StateValue 'ScheduledFor')
    if (($task.NextRunTime -gt $now) -or ($scheduledFor -and $now -lt $scheduledFor.AddMinutes(30))) {
        if ($scheduledFor) {
            Write-Output ('OK: user scheduled the restart for {0:yyyy-MM-dd HH:mm}' -f $scheduledFor)
        } else {
            Write-Output 'OK: user scheduled the restart'
        }
        exit 0
    }
    # The chosen time passed without a restart (device asleep, or the countdown was cancelled). Ask again.
    Remove-RestartTask
    Remove-StateValue 'ScheduledFor'
}

$snoozeUntil = ConvertFrom-StateDate (Get-StateValue 'SnoozeUntil')
if ($snoozeUntil -and $now -lt $snoozeUntil) {
    Write-Output ('OK: reminder snoozed ({0})' -f (Get-StateValue 'SnoozeReason'))
    exit 0
}

Write-Output 'Restart pending: user has not picked a restart time'
exit 1
