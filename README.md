# Windows 11 25H2 In-Place Upgrade

`Upgrade-Win11-25H2.ps1` upgrades Windows 11 23H2 (Enterprise/Education) to 25H2 in place from the Business Editions ISO, keeping apps, data and settings. It runs from install media, so it works without Windows Update or WSUS.

Before setup starts, it checks the OS, looks for a pending reboot, checks the hardware requirements (TPM 2.0, UEFI, SSE4.2/POPCNT, RAM), frees disk space and repairs the component store if it's flagged as corrupt (each of those fixes is capped at 1 minute). It deletes any earlier local copy of the media, copies the ISO fresh, checks that its language, architecture and build match the machine, then suspends BitLocker. It then opens the interactive Windows Setup UI (`setup.exe` with no arguments); the tech finishes the upgrade there and Setup handles the restart.

## Usage

Run in Windows PowerShell 5.1 as Administrator:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Upgrade-Win11-25H2.ps1 -IsoPath "\\server\share\Win11_25H2_Business_x64.iso"
```

To run only the compatibility scan (makes no changes):

```powershell
.\Upgrade-Win11-25H2.ps1 -IsoPath "\\server\share\Win11_25H2_Business_x64.iso" -CompatScanOnly
```

Run `Get-Help .\Upgrade-Win11-25H2.ps1 -Full` to see every parameter and exit code.
