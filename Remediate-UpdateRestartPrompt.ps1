<#
.SYNOPSIS
    Intune Remediation - REMEDIATION script.
    Shows a centered "Restart required - pick a time" dialog when Windows Update is waiting on a restart.

.DESCRIPTION
    The user can:
      Schedule restart  Creates a one-time task in the user's own context. It fires $WarningMinutes before the
                        chosen time, re-checks that an update restart is still pending, then starts a shutdown.exe
                        countdown (Windows shows its own restart warning). If the device is asleep at that time,
                        the restart is skipped and the user is asked again on the next run.
      Restart now       Restarts right away. Apps with unsaved work can still hold the restart so the user can save.
      Remind me later   Snoozes the dialog for $SnoozeHours.

    An ignored dialog closes itself after $DialogTimeoutMinutes and comes back after $NoResponseSnoozeHours.
    If the user is presenting, in a full-screen app, or locked, the dialog waits $BusyRetryMinutes.

    Windows still enforces the update ring deadline on its own. This only gives users an easier way
    to restart before it. Microsoft's Remediations guidance says not to put reboot commands in these
    scripts; here a restart only ever happens because the user clicked Restart now or picked a time.

    Intune settings: same as the detection script (logged-on credentials = Yes, 64-bit PowerShell = Yes).

.PARAMETER Preview
    Test mode for running by hand in a normal PowerShell window: shows the dialog even when no restart is
    pending, and doesn't schedule, restart, or save anything. Intune never passes it.

.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File .\Remediate-UpdateRestartPrompt.ps1 -Preview
#>
param([switch]$Preview)

# --- Settings -----------------------------------------------------------------
$OrgName               = 'IT Department'   # shown under the dialog title
$GracePeriodDays       = 4                 # match "Grace period" in the update ring
$SnoozeHours           = 4                 # "Remind me later"
$DialogTimeoutMinutes  = 15                # the Intune run waits while the dialog is open, so close an ignored one after this long...
$NoResponseSnoozeHours = 1                 # ...and show it again after this long
$BusyRetryMinutes      = 30                # user presenting / full screen / locked: try again after this long
$WarningMinutes        = 15                # countdown before a scheduled restart
$MaxScheduleDays       = 7                 # furthest out a user can schedule (the deadline also caps it)

# --- Keep in sync with the detection script ---------------------------------
$StateKey = 'HKCU:\Software\UpdateRestartPrompt'
$TaskName = "Update restart - $env:USERNAME"
# ------------------------------------------------------------------------------

#region Helpers
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

function Set-Snooze([datetime]$Until, [string]$Reason) {
    Set-StateValue 'SnoozeUntil' $Until.ToString('o')
    Set-StateValue 'SnoozeReason' $Reason
}

function Test-UserBusy {
    # SHQueryUserNotificationState: 1 locked or screen saver, 2 full-screen app, 3 full-screen Direct3D,
    # 4 presentation mode, 5 normal, 6 quiet time after first sign-in, 7 full-screen Store app.
    # Compiles a one-line P/Invoke; if your EDR objects to that, delete this check.
    try {
        if (-not ('UpdateRestartPrompt.NativeMethods' -as [type])) {
            Add-Type -Namespace UpdateRestartPrompt -Name NativeMethods -MemberDefinition @'
[DllImport("shell32.dll")]
public static extern int SHQueryUserNotificationState(out int state);
'@
        }
        $state = 0
        if ([UpdateRestartPrompt.NativeMethods]::SHQueryUserNotificationState([ref]$state) -eq 0) {
            return ($state -in 1, 2, 3, 4, 7)
        }
    } catch { }
    return $false
}

function Get-RestartPendingSince {
    # Best estimate of when the device started waiting for a restart. It errs early, so the
    # deadline shown to the user is never later than the one Windows enforces.
    try {
        $lastBoot = (Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop).LastBootUpTime
    } catch {
        $lastBoot = [datetime]::MinValue
    }
    $candidates = New-Object System.Collections.Generic.List[datetime]

    # Earliest update installed since the last boot. Defender definitions/platform and MSRT are
    # skipped because they install often and never need a restart.
    try {
        $searcher = (New-Object -ComObject Microsoft.Update.Session).CreateUpdateSearcher()
        $count = $searcher.GetTotalHistoryCount()
        if ($count -gt 0) {
            foreach ($entry in $searcher.QueryHistory(0, [Math]::Min($count, 100))) {
                if ($entry.Operation -ne 1) { continue }              # 1 = installation
                if ($entry.ResultCode -notin 1, 2, 3) { continue }    # in progress, succeeded, succeeded with errors
                if ($entry.Title -match 'KB2267602|KB4052623|KB890830|Defender Antivirus|Malicious Software Removal') { continue }
                $installed = [datetime]::SpecifyKind($entry.Date, [DateTimeKind]::Utc).ToLocalTime()   # history is in UTC
                if ($installed -gt $lastBoot) { $candidates.Add($installed) }
            }
        }
    } catch { }

    # When the detection script first saw the pending restart.
    $firstSeen = ConvertFrom-StateDate (Get-StateValue 'PendingSince')
    if ($firstSeen -and $firstSeen -gt $lastBoot) { $candidates.Add($firstSeen) }

    if ($candidates.Count -eq 0) { return $null }
    return ($candidates | Sort-Object | Select-Object -First 1)
}

function Get-ActiveHoursEnd {
    # Update ring value first, then the device's own setting; 6 PM if neither can be read.
    foreach ($path in 'HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Update',
                      'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings') {
        try {
            $value = (Get-ItemProperty -LiteralPath $path -Name 'ActiveHoursEnd' -ErrorAction Stop).ActiveHoursEnd
            if ($null -ne $value -and [int]$value -ge 0 -and [int]$value -le 23) { return [int]$value }
        } catch { }
    }
    return 18
}

function Register-RestartTask([datetime]$RestartAt) {
    $fireAt  = $RestartAt.AddMinutes(-$WarningMinutes)
    $message = "This device will restart in $WarningMinutes minutes to finish installing updates. Save your work now."

    # The task re-checks at run time, so a manual restart in the meantime cancels this one.
    $command = ('$p = $false; try { $p = (New-Object -ComObject Microsoft.Update.SystemInfo).RebootRequired } catch { }; ' +
                'if ($p -or (Test-Path -LiteralPath ''HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'')) ' +
                '{ shutdown.exe /r /t __SECONDS__ /d p:2:17 /c ''__MESSAGE__'' }'
               ).Replace('__SECONDS__', [string]($WarningMinutes * 60)).Replace('__MESSAGE__', $message)

    $service = New-Object -ComObject Schedule.Service
    $service.Connect()
    $definition = $service.NewTask(0)
    $definition.RegistrationInfo.Description = 'Restart picked by the user to finish installing updates. Created by the Intune update restart prompt and removed automatically.'
    $definition.Principal.LogonType                 = 3        # interactive token: runs as this user, only while signed in
    $definition.Settings.DisallowStartIfOnBatteries = $false
    $definition.Settings.StopIfGoingOnBatteries     = $false
    $definition.Settings.StartWhenAvailable         = $false   # missed (device asleep) = skipped; the user is asked again
    $definition.Settings.ExecutionTimeLimit         = 'PT5M'
    $definition.Settings.DeleteExpiredTaskAfter     = 'PT1H'

    $trigger = $definition.Triggers.Create(1)                  # 1 = one time
    $trigger.StartBoundary = $fireAt.ToString('s')
    $trigger.EndBoundary   = $RestartAt.AddHours(1).ToString('s')

    $action = $definition.Actions.Create(0)                    # 0 = start a program
    $action.Path      = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    $action.Arguments = '-NoProfile -NonInteractive -WindowStyle Hidden -Command "' + $command + '"'

    # 6 = create or update, 3 = interactive token for the current user
    [void]$service.GetFolder('\').RegisterTaskDefinition($TaskName, $definition, 6, $null, $null, 3)
}

function Format-DayLabel([datetime]$Day) {
    $today = (Get-Date).Date
    if ($Day.Date -eq $today) { return 'Today' }
    if ($Day.Date -eq $today.AddDays(1)) { return 'Tomorrow' }
    return $Day.ToString('dddd, MMMM d')
}

function Format-When([datetime]$When) {
    $today = (Get-Date).Date
    if ($When.Date -eq $today) { $day = 'today' }
    elseif ($When.Date -eq $today.AddDays(1)) { $day = 'tomorrow' }
    else { $day = 'on ' + $When.ToString('dddd, MMMM d') }
    return ('{0} at {1}' -f $day, $When.ToString('t'))
}
#endregion

#region Dialog helpers (use $ui, $slots and $deadline from the script scope)
function Update-TimeBox($Day, $Preferred) {
    $ui.TimeBox.Items.Clear()
    $selectIndex = 0
    foreach ($slotTime in $slots) {
        if ($slotTime.Date -ne $Day) { continue }
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = $slotTime.ToString('t')
        $item.Tag = $slotTime
        $index = $ui.TimeBox.Items.Add($item)
        if ($Preferred -and $slotTime -eq $Preferred) { $selectIndex = $index }
    }
    $ui.TimeBox.SelectedIndex = $selectIndex
}

function Set-DeadlineText {
    $ui.DeadlineText.Inlines.Clear()
    if ($deadline -gt (Get-Date)) {
        $when = [System.Windows.Documents.Run]::new(('{0} at {1}' -f $deadline.ToString('dddd, MMMM d'), $deadline.ToString('t')))
        $when.FontWeight = [System.Windows.FontWeights]::SemiBold
        $ui.DeadlineText.Inlines.Add([System.Windows.Documents.Run]::new("If it hasn't restarted by "))
        $ui.DeadlineText.Inlines.Add($when)
        $ui.DeadlineText.Inlines.Add([System.Windows.Documents.Run]::new(', Windows will restart it automatically, even during work hours.'))
    } else {
        $ui.DeadlineText.Inlines.Add([System.Windows.Documents.Run]::new('Windows can restart it automatically at any time now.'))
    }
}

function Show-Error([string]$Message) {
    $ui.ErrorText.Text = $Message
    $ui.ErrorText.Visibility = 'Visible'
}

function Show-Confirmation([datetime]$When) {
    $ui.TitleText.Text = 'Restart scheduled'
    $ui.BodyText.Text  = "Your device will restart $(Format-When $When). You'll get a $WarningMinutes-minute warning first, so save your work before then."
    foreach ($name in 'DeadlineText', 'PickerPanel', 'ErrorText', 'ScheduleBtn', 'NowBtn', 'LaterBtn') {
        $ui[$name].Visibility = 'Collapsed'
    }
    $ui.ScheduleBtn.IsDefault = $false
    $ui.LaterBtn.IsCancel     = $false
    $ui.OkBtn.Visibility      = 'Visible'
    $ui.OkBtn.IsDefault       = $true
    $ui.OkBtn.IsCancel        = $true
    [void]$ui.OkBtn.Focus()
}
#endregion

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Restart required" Width="560" SizeToContent="Height"
        WindowStartupLocation="CenterScreen" WindowStyle="None" ResizeMode="NoResize"
        AllowsTransparency="True" Background="Transparent" Topmost="True" ShowInTaskbar="True"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="14" Foreground="#1A1A1A"
        UseLayoutRounding="True">
  <Window.Resources>
    <Style x:Key="DialogButton" TargetType="Button">
      <Setter Property="Background" Value="#FBFBFB"/>
      <Setter Property="Foreground" Value="#1A1A1A"/>
      <Setter Property="BorderBrush" Value="#CCCCCC"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="16,7"/>
      <Setter Property="MinWidth" Value="128"/>
      <Setter Property="Margin" Value="8,0,0,0"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Chrome" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Chrome" Property="Opacity" Value="0.88"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="Chrome" Property="Opacity" Value="0.72"/>
              </Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True">
                <Setter TargetName="Chrome" Property="BorderBrush" Value="#1A1A1A"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="AccentButton" TargetType="Button" BasedOn="{StaticResource DialogButton}">
      <Setter Property="Background" Value="#005FB8"/>
      <Setter Property="Foreground" Value="#FFFFFF"/>
      <Setter Property="BorderBrush" Value="#005FB8"/>
    </Style>
  </Window.Resources>

  <Grid Margin="20">
    <!-- Shadow sits on its own layer so the text stays crisp -->
    <Border Background="#FFFFFF" CornerRadius="8">
      <Border.Effect>
        <DropShadowEffect BlurRadius="28" ShadowDepth="6" Direction="270" Opacity="0.28"/>
      </Border.Effect>
    </Border>
    <Border Background="#FFFFFF" CornerRadius="8" BorderBrush="#E0E0E0" BorderThickness="1">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Margin="28,24,28,22">
          <Grid Margin="0,0,0,14">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock Grid.Column="0" Text="&#xE895;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="30"
                       Foreground="#005FB8" VerticalAlignment="Center" Margin="0,0,16,0"/>
            <StackPanel Grid.Column="1" VerticalAlignment="Center">
              <TextBlock x:Name="TitleText" Text="Restart required" FontSize="20" FontWeight="SemiBold"/>
              <TextBlock x:Name="OrgText" FontSize="12" Foreground="#616161" Margin="0,2,0,0"/>
            </StackPanel>
          </Grid>

          <TextBlock x:Name="BodyText" TextWrapping="Wrap" LineHeight="20"/>
          <TextBlock x:Name="DeadlineText" TextWrapping="Wrap" LineHeight="20" Margin="0,10,0,0"/>

          <StackPanel x:Name="PickerPanel" Margin="0,18,0,0">
            <TextBlock Text="Choose when to restart" FontWeight="SemiBold" Margin="0,0,0,8"/>
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="12"/>
                <ColumnDefinition Width="150"/>
              </Grid.ColumnDefinitions>
              <ComboBox x:Name="DayBox" Grid.Column="0" Height="34" Padding="10,0" VerticalContentAlignment="Center"/>
              <ComboBox x:Name="TimeBox" Grid.Column="2" Height="34" Padding="10,0" VerticalContentAlignment="Center" MaxDropDownHeight="260"/>
            </Grid>
          </StackPanel>

          <TextBlock x:Name="ErrorText" Foreground="#C42B1C" TextWrapping="Wrap" Margin="0,10,0,0" Visibility="Collapsed"/>
        </StackPanel>

        <Border Grid.Row="1" Background="#F3F3F3" BorderBrush="#E5E5E5" BorderThickness="0,1,0,0" CornerRadius="0,0,7,7" Padding="28,16">
          <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="ScheduleBtn" Content="Schedule restart" Style="{StaticResource AccentButton}" IsDefault="True"/>
            <Button x:Name="NowBtn" Content="Restart now" Style="{StaticResource DialogButton}"/>
            <Button x:Name="LaterBtn" Content="Remind me later" Style="{StaticResource DialogButton}" IsCancel="True"/>
            <Button x:Name="OkBtn" Content="OK" Style="{StaticResource AccentButton}" Visibility="Collapsed"/>
          </StackPanel>
        </Border>
      </Grid>
    </Border>
  </Grid>
</Window>
'@

# --- Pre-flight -----------------------------------------------------------------
if ([Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA) {
    Write-Output 'Error: the dialog needs an STA thread (the Windows PowerShell 5.1 default)'
    exit 1
}

if (-not $Preview) {
    if (-not (Test-UpdateRestartPending)) {
        Write-Output 'OK: no update restart pending'
        exit 0
    }
    if (Test-UserBusy) {
        $retryAt = (Get-Date).AddMinutes($BusyRetryMinutes)
        Set-Snooze $retryAt 'Busy'
        Write-Output 'Busy: user is presenting, in a full-screen app, or locked; will ask again'
        exit 0
    }
}

$mutex = [System.Threading.Mutex]::new($false, 'Local\UpdateRestartPrompt')
$haveLock = $false
try { $haveLock = $mutex.WaitOne(0) } catch [System.Threading.AbandonedMutexException] { $haveLock = $true }
if (-not $haveLock) {
    Write-Output 'Skipped: the restart dialog is already open'
    exit 0
}

try {
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

    # --- Deadline and restart times the user can pick ---------------------------------
    $now          = Get-Date
    $pendingSince = Get-RestartPendingSince
    if (-not $pendingSince) { $pendingSince = $now }
    $deadline = $pendingSince.AddDays($GracePeriodDays)
    $deadline = $deadline.Date.AddHours($deadline.Hour)   # round down to the hour

    # 30-minute steps, starting far enough out for the warning countdown and ending
    # 30 minutes before the deadline (or $MaxScheduleDays out, whichever comes first).
    $earliest = $now.AddMinutes($WarningMinutes + 5)
    $latest   = $deadline.AddMinutes(-30)
    if ($latest -gt $now.AddDays($MaxScheduleDays)) { $latest = $now.AddDays($MaxScheduleDays) }

    $slots = New-Object System.Collections.Generic.List[datetime]
    $t = $earliest.Date.AddMinutes([Math]::Ceiling($earliest.TimeOfDay.TotalMinutes / 30) * 30)
    while ($t -le $latest) { $slots.Add($t); $t = $t.AddMinutes(30) }

    $days = New-Object System.Collections.Generic.List[datetime]
    foreach ($s in $slots) { if (-not $days.Contains($s.Date)) { $days.Add($s.Date) } }

    # Pre-select "tonight", like the old dialog: the end of active hours today, the soonest slot if
    # that's already past, or tomorrow evening if it's after midnight.
    $default = $null
    if ($slots.Count -gt 0) {
        $default = $now.Date.AddHours((Get-ActiveHoursEnd))
        if ($default -lt $earliest) {
            if ($earliest.Date -eq $now.Date) { $default = $slots[0] } else { $default = $default.AddDays(1) }
        }
        if ($default -gt $latest) { $default = $slots[0] }
    }

    # --- Build the dialog ----------------------------------------------------------
    $window = [Windows.Markup.XamlReader]::Parse($xaml)
    $ui = @{}
    foreach ($name in 'TitleText', 'OrgText', 'BodyText', 'DeadlineText', 'PickerPanel', 'DayBox', 'TimeBox',
                      'ErrorText', 'ScheduleBtn', 'NowBtn', 'LaterBtn', 'OkBtn') {
        $ui[$name] = $window.FindName($name)
    }

    $ui.OrgText.Text  = $OrgName
    $ui.BodyText.Text = 'Updates were installed on this device, and it needs to restart to finish installing them. Pick a time that works for you, or restart now.'
    Set-DeadlineText

    if ($slots.Count -gt 0) {
        foreach ($d in $days) {
            $item = New-Object System.Windows.Controls.ComboBoxItem
            $item.Content = Format-DayLabel $d
            $item.Tag = $d
            [void]$ui.DayBox.Items.Add($item)
        }
        $ui.DayBox.SelectedIndex = $days.IndexOf($default.Date)
        Update-TimeBox $default.Date $default

        # Changing the day keeps the same time of day when that slot exists.
        $ui.DayBox.Add_SelectionChanged({
            $preferred = $null
            if ($ui.TimeBox.SelectedItem) {
                $preferred = ([datetime]$ui.DayBox.SelectedItem.Tag).Add(([datetime]$ui.TimeBox.SelectedItem.Tag).TimeOfDay)
            }
            Update-TimeBox $ui.DayBox.SelectedItem.Tag $preferred
        })
    } else {
        # Too close to the deadline to offer a later time.
        $ui.PickerPanel.Visibility = 'Collapsed'
        $ui.ScheduleBtn.Visibility = 'Collapsed'
        $ui.ScheduleBtn.IsDefault  = $false
        $ui.NowBtn.Style           = $window.FindResource('AccentButton')
        $ui.NowBtn.IsDefault       = $true
        $ui.BodyText.Text = 'Updates were installed on this device, and it needs to restart soon to finish installing them. Save your work and restart now so it happens on your schedule.'
    }

    $script:Choice     = $null
    $script:ChosenTime = $null

    $ui.ScheduleBtn.Add_Click({
        $ui.ErrorText.Visibility = 'Collapsed'
        $pick = $null
        if ($ui.TimeBox.SelectedItem) { $pick = [datetime]$ui.TimeBox.SelectedItem.Tag }
        if (-not $pick) { Show-Error 'Choose a day and time.'; return }
        if ($pick -lt (Get-Date).AddMinutes($WarningMinutes + 1)) {
            Show-Error 'That time is too soon. Choose a later time, or restart now.'
            return
        }
        try {
            if (-not $Preview) {
                Register-RestartTask -RestartAt $pick
                Set-StateValue 'ScheduledFor' $pick.ToString('o')
                Remove-StateValue 'SnoozeUntil'
                Remove-StateValue 'SnoozeReason'
            }
            $script:Choice     = 'Scheduled'
            $script:ChosenTime = $pick
            Show-Confirmation $pick
        } catch {
            Show-Error ("Couldn't schedule the restart ({0}). Try another time, or restart now." -f $_.Exception.Message)
        }
    })
    $ui.NowBtn.Add_Click({ $script:Choice = 'RestartNow'; $window.Close() })
    $ui.LaterBtn.Add_Click({ $script:Choice = 'Snooze'; $window.Close() })
    $ui.OkBtn.Add_Click({ $window.Close() })

    # Close an ignored dialog so the Intune run doesn't hang; it comes back later.
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMinutes($DialogTimeoutMinutes)
    $timer.Add_Tick({
        $timer.Stop()
        if (-not $script:Choice) { $script:Choice = 'NoResponse' }
        $window.Close()
    })

    $window.Add_MouseLeftButtonDown({ try { $window.DragMove() } catch { } })
    $window.Add_Loaded({ $timer.Start(); [void]$window.Activate() })

    [void]$window.ShowDialog()
    $timer.Stop()

    # --- Act on the choice ---------------------------------------------------------
    if (-not $script:Choice) { $script:Choice = 'Snooze' }   # closed some other way (Alt+F4)
    $prefix = ''
    if ($Preview) { $prefix = 'PREVIEW, nothing changed - ' }

    switch ($script:Choice) {
        'Scheduled' {
            Write-Output ('{0}Scheduled: user picked {1:yyyy-MM-dd HH:mm}' -f $prefix, $script:ChosenTime)
        }
        'RestartNow' {
            if ($Preview) {
                Write-Output "${prefix}RestartNow: would restart immediately"
            } else {
                & "$env:SystemRoot\System32\shutdown.exe" /r /t 0 /d p:2:17
                if ($LASTEXITCODE -ne 0) {
                    Write-Output "Error: shutdown.exe exited with $LASTEXITCODE"
                    exit 1
                }
                Write-Output 'RestartNow: user restarted from the dialog'
            }
        }
        'NoResponse' {
            if (-not $Preview) { Set-Snooze ((Get-Date).AddHours($NoResponseSnoozeHours)) 'NoResponse' }
            Write-Output "${prefix}NoResponse: dialog was ignored; will ask again"
        }
        default {
            if (-not $Preview) { Set-Snooze ((Get-Date).AddHours($SnoozeHours)) 'User' }
            Write-Output "${prefix}Snoozed: user chose Remind me later"
        }
    }
    exit 0
} catch {
    Write-Output ('Error: {0}' -f $_.Exception.Message)
    exit 1
} finally {
    if ($haveLock) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
