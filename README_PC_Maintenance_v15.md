# PC Maintenance v15

A Windows PowerShell maintenance, health-check, cleanup, repair, and reporting script designed for monthly desktop and laptop servicing.

The script can run in **Audit** mode for report-only checks, or **Remediate** mode for active cleanup and repair. It creates branded HTML reports, JSON/CSV summaries, transcripts, optional ZIP packages, and can send notifications by email or Microsoft Teams webhook.

> **Recommended path:** `C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1`  
> **Recommended logo path:** `C:\PCMaintenance\Silicon Beach.png` or `C:\PCMaintenance\logo.png`  
> **Default reports path:** `C:\PCMaintenance\Reports\yyyyMMdd_HHmmss-Audit\` or `C:\PCMaintenance\Reports\yyyyMMdd_HHmmss-Cleanup\`

---

## Table of Contents

- [Overview](#overview)
- [Major Features](#major-features)
- [Folder Layout](#folder-layout)
- [Requirements](#requirements)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Run Modes](#run-modes)
- [Remediate Presets: Lite and Full](#remediate-presets-lite-and-full)
- [Switch Reference](#switch-reference)
- [JSON Configuration Guide](#json-configuration-guide)
- [Report Outputs](#report-outputs)
- [Email Reporting](#email-reporting)
- [Microsoft Teams Webhook Reporting](#microsoft-teams-webhook-reporting)
- [Remote and RMM Usage](#remote-and-rmm-usage)
- [Restore Points and VSS Shadow Storage](#restore-points-and-vss-shadow-storage)
- [Safety Notes](#safety-notes)
- [Example Commands](#example-commands)
- [Version Notes](#version-notes)
- [Troubleshooting](#troubleshooting)
- [Suggested Operational Workflow](#suggested-operational-workflow)

---

## Overview

`pc_cleanup_maintenance_v15.ps1` performs structured monthly maintenance on Windows endpoints. It collects system health data, performs optional cleanup/repair actions, creates before/after metrics, and generates a branded report suitable for technician review, client reporting, or RMM ingestion.

The script is designed for:

- Windows desktop PCs
- Windows laptops
- MSP/RMM support workflows
- Monthly maintenance tasks
- Pre/post cleanup reporting
- Local or remote PowerShell execution

---

## Major Features

### Cleanup and Repair

- Temporary file cleanup
- Browser cache cleanup for Edge, Chrome, and Firefox when enabled
- Recycle Bin cleanup when enabled
- Delivery Optimization cleanup
- Windows component store cleanup
- Optional `Windows.old` cleanup
- Optional crash dump cleanup
- Optional stale scheduled task removal
- Optional stale local profile removal
- CHKDSK smart scan and optional `/f` queueing
- SFC scan with captured output
- DISM RestoreHealth with captured output

### Health and Diagnostics

- Disk free space before and after
- Per-step disk delta and duration tracking
- VSS/shadow storage before and after
- Pending reboot checks
- Windows Update diagnostics
- BitLocker health
- Power/sleep health
- Defender status and optional quick scan
- Disk/SMART/storage reliability checks where available
- Device Manager problem devices
- Automatic services not running
- Startup item audit
- Teams and OneDrive cache/sync checks
- Battery report for laptops
- Crash dump summary
- Event log summary and reliability-style event interpretation
- Local user profile sizes
- Storage hotspots and large files
- Windows component store analysis

### Reporting and Automation

- Branded HTML report
- JSON summary output
- CSV summary output
- PowerShell transcript
- Optional ZIP packaging of the run folder
- Optional email report delivery
- Optional Microsoft Teams webhook summary
- Optional RMM-style exit codes
- Optional remote execution wrapper
- JSON config file support
- Report retention cleanup
- `-ValidateOnly` parse check mode

---

## Folder Layout

Recommended layout:

```text
C:\PCMaintenance\
├── pc_cleanup_maintenance_v15.ps1
├── pc_cleanup_config_template_v15.json
├── Silicon Beach.png
└── Reports\
    ├── 20260510_091500-Audit\
    └── 20260510_103200-Cleanup\
```

Each script run creates a unique timestamped folder under `Reports`:

```text
C:\PCMaintenance\Reports\yyyyMMdd_HHmmss-Audit\
C:\PCMaintenance\Reports\yyyyMMdd_HHmmss-Cleanup\
```

Examples:

```text
C:\PCMaintenance\Reports\20260510_091500-Audit\
C:\PCMaintenance\Reports\20260510_103200-Cleanup\
```

The script places the HTML, JSON, CSV, transcript, event log exports, battery reports, and optional ZIP output in the run folder.

---

## Requirements

### Required

- Windows PowerShell 5.1+
- Local administrator privileges for best results
- Script execution permitted through `-ExecutionPolicy Bypass`, local policy, or signing

### Recommended

- Run from an elevated PowerShell session
- Keep the script in `C:\PCMaintenance`
- Keep the logo in the same folder as the script
- Run `-ValidateOnly` before using a new version
- Test with `-Mode Audit` before running remediation

### Optional Dependencies / Conditions

Some sections only run or populate if supported by the local PC:

- Defender cmdlets for Defender status and quick scan
- BitLocker cmdlets for BitLocker health
- Battery hardware for battery report
- Storage reliability counters for deeper SMART/storage details
- WindowsUpdateProvider or PSWindowsUpdate-compatible availability for update actions
- PowerShell Remoting for `-ComputerName` remote execution

---

## Installation

Create the working folder:

```powershell
New-Item -ItemType Directory -Path C:\PCMaintenance -Force
New-Item -ItemType Directory -Path C:\PCMaintenance\Reports -Force
```

Copy the script, config, and logo:

```powershell
Copy-Item .\pc_cleanup_maintenance_v15.ps1 C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -Force
Copy-Item .\pc_cleanup_config_template_v15.json C:\PCMaintenance\pc_cleanup_config_template_v15.json -Force
Copy-Item ".\Silicon Beach.png" "C:\PCMaintenance\Silicon Beach.png" -Force
```

Unblock the script if it was downloaded from the internet:

```powershell
Unblock-File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1
```

Validate the script before running it:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -ValidateOnly
```

---

## Quick Start

### Show help

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 /?
```

or:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -ShowHelp
```

### Audit-only test run

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -Mode Audit -SkipWindowsUpdate -SkipSfc -SkipDismRestoreHealth
```

### Lite cleanup run

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -Mode Remediate -Lite
```

### Full cleanup run

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -Mode Remediate -Full
```

---

## Run Modes

### `-Mode Audit`

Report-only mode.

The script gathers system health information and estimates cleanup potential, but avoids destructive cleanup/remediation actions.

Use Audit mode when:

- testing the script on a new device
- checking what would be cleaned
- gathering endpoint health data without making changes
- running scheduled assessments

Example:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -Mode Audit
```

### `-Mode Remediate`

Active maintenance mode.

The script performs cleanup and repair actions according to the switches or presets supplied.

Example:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -Mode Remediate -Lite
```

### `-AuditOnly`

Shortcut for Audit mode.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -AuditOnly
```

---

## Remediate Presets: Lite and Full

### `-Lite`

Light cleanup preset for routine monthly runs.

Enables:

```powershell
-RemoveStaleTasks -ArchiveEventLogs -ClearEventLogs
```

Example:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -Mode Remediate -Lite
```

### `-Full`

Full cleanup preset for deeper monthly maintenance.

Enables:

```powershell
-RemoveStaleTasks -ArchiveEventLogs -ClearEventLogs -EmptyRecycleBin -ClearBrowserCache -CleanCrashDumps -CleanWindowsOld
```

Example:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -Mode Remediate -Full
```

### Default Remediate Behaviour

If you run Remediate mode without `-Full`, `-Lite`, or explicit cleanup switches, the script defaults to **Lite**.

This:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -Mode Remediate
```

behaves like:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -Mode Remediate -Lite
```

---

## Switch Reference

### Help and Validation

| Switch | Description |
|---|---|
| `/?` | Shows script help. |
| `-ShowHelp` | Shows script help. |
| `-Help` | Alias for help. |
| `-h` | Alias for help. |
| `-ValidateOnly` | Runs a self-parse check and exits without maintenance. |

### Run Mode and Presets

| Switch | Description |
|---|---|
| `-Mode Audit` | Report-only mode. |
| `-Mode Remediate` | Active maintenance mode. |
| `-AuditOnly` | Shortcut for Audit mode. |
| `-Lite` | Enables the Lite remediation preset. |
| `-Full` | Enables the Full remediation preset. |
| `-AggressiveCleanup` | Enables more assertive cleanup logic where implemented. Does not remove user profiles automatically. |

### Core Settings

| Switch | Default | Description |
|---|---:|---|
| `-DriveLetter` | `C` | Target drive for disk checks and CHKDSK. |
| `-ReportRoot` | Script folder `\Reports` | Base output folder. A timestamped run folder is created under it. |
| `-TempFileMinAgeDays` | `2` | Minimum file age for temp cleanup. |
| `-EventLookbackDays` | `14` | How far back to scan event logs. |
| `-StaleTaskDays` | `90` | Age threshold for stale scheduled tasks. |
| `-ProfileStaleDays` | `90` | Age threshold for stale profile reporting. |
| `-TopProcessCount` | `5` | Number of top CPU/memory processes to report. |
| `-TopLargeFileCount` | `20` | Number of large files/storage hotspots to report. |
| `-CpuWarnPercent` | `85` | CPU warning threshold. |
| `-MemoryWarnPercent` | `85` | Memory warning threshold. |
| `-DiskQueueWarn` | `2.0` | Disk queue warning threshold. |
| `-LowDiskWarnPercent` | `15` | Low disk warning threshold. |

### Cleanup Switches

| Switch | Description |
|---|---|
| `-RemoveStaleTasks` | Removes eligible stale scheduled tasks. Microsoft task paths are avoided. |
| `-ClearEventLogs` | Clears Application/System logs after notable events are captured. |
| `-ArchiveEventLogs` | Exports Application/System logs to the run folder before clearing. |
| `-EmptyRecycleBin` | Clears the Recycle Bin. |
| `-ClearBrowserCache` | Cleans Edge, Chrome, and Firefox cache paths. Does not target cookies, passwords, or bookmarks. |
| `-CleanCrashDumps` | Deletes `C:\Windows\Minidump` files and `C:\Windows\MEMORY.DMP` in Remediate mode. |
| `-CleanWindowsOld` | Deletes `C:\Windows.old` in Remediate mode. Removes rollback files. |
| `-RemoveStaleProfiles` | Enables stale profile removal logic. Requires `-ForceProfileRemoval`. |
| `-ForceProfileRemoval` | Required with `-RemoveStaleProfiles` to actually remove old profiles. |
| `-RemoveProfilesOlderThanDays` | Age threshold for stale profile removal. Default: `180`. |

### Skip Switches

| Switch | Description |
|---|---|
| `-SkipRecycleBin` | Skips Recycle Bin handling. |
| `-SkipComponentCleanup` | Skips DISM component store cleanup. |
| `-SkipDeliveryOptimizationCleanup` | Skips Delivery Optimization cleanup. |
| `-SkipWindowsUpdate` | Skips Windows Update scan/install. |
| `-SkipRestorePoints` | Skips restore point creation. |
| `-SkipChkdsk` | Skips CHKDSK scan/queue logic. |
| `-SkipSfc` | Skips `sfc /scannow`. |
| `-SkipDismRestoreHealth` | Skips DISM RestoreHealth. |
| `-SkipUserProfileSizes` | Skips local user profile size scanning. |
| `-SkipStorageHotspots` | Skips large file/storage hotspot scanning. |
| `-SkipBatteryReport` | Skips battery report generation. |
| `-SkipTranscript` | Skips PowerShell transcript creation. |
| `-SkipComponentStoreAnalysis` | Skips DISM AnalyzeComponentStore. |

### Windows Update Switches

| Switch | Description |
|---|---|
| `-SkipWindowsUpdate` | Skips Windows Update. |
| `-DownloadUpdatesOnly` | Downloads updates where supported, but does not install them. |

### Repair and Health Switches

| Switch | Description |
|---|---|
| `-AlwaysQueueChkdskFix` | Queues `chkdsk C: /f` regardless of `/scan` result. |
| `-RunDefenderQuickScan` | Runs a Microsoft Defender quick scan if Defender cmdlets are available. |
| `-CreateAfterRestorePoint` | Creates an AFTER restore point. v15 creates the BEFORE restore point by default. |
| `-MaxShadowStorage` | Caps System Restore/VSS storage, e.g. `10GB` or `5%`. |

### Reporting and Retention Switches

| Switch | Description |
|---|---|
| `-CleanOldReports` | Deletes old report folders under the report root. |
| `-ReportRetentionDays` | Retention period for old report folders. Default: `90`. |
| `-ZipReportFolder` | Compresses the current run folder into a ZIP at the end. |
| `-UseRmmExitCode` | Uses RMM-friendly process exit codes. |

### Config, Email, Webhook and Remote Switches

| Switch | Description |
|---|---|
| `-ConfigPath` | Loads defaults from a JSON config file. |
| `-TeamsWebhookUrl` | Posts a summary to a Teams incoming webhook. |
| `-EmailReportTo` | Sends report files by email. Requires `-SmtpServer` and `-EmailFrom`. |
| `-EmailFrom` | Sender address for SMTP email. |
| `-SmtpServer` | SMTP server. |
| `-SmtpPort` | SMTP port. Default: `587`. |
| `-SmtpUseSsl` | Enables SSL/TLS for SMTP. |
| `-NotificationSubjectPrefix` | Prefix for email/Teams notification subjects. Default: `PC Maintenance`. |
| `-ComputerName` | Runs the script remotely through PowerShell Remoting. |
| `-RemoteScriptPath` | Script path on remote machine. Default: `C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1`. |
| `-RemotePassthruArgs` | Additional arguments passed to the remote run. |

---

## JSON Configuration Guide

The JSON config file lets you define defaults without typing a long command every time.

The script only uses the JSON config when you pass `-ConfigPath`.

Example:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -ConfigPath C:\PCMaintenance\pc_cleanup_config_template_v15.json
```

### Example JSON: Audit Default

```json
{
  "Mode": "Audit",
  "Lite": true,
  "Full": false,
  "ReportRoot": "C:\\PCMaintenance\\Reports",
  "DriveLetter": "C",
  "SkipWindowsUpdate": true,
  "SkipSfc": true,
  "SkipDismRestoreHealth": true,
  "SkipRestorePoints": true
}
```

### Example JSON: Lite Monthly Cleanup

```json
{
  "Mode": "Remediate",
  "Lite": true,
  "Full": false,
  "ReportRoot": "C:\\PCMaintenance\\Reports",
  "DriveLetter": "C",
  "CreateAfterRestorePoint": false,
  "CleanOldReports": true,
  "ReportRetentionDays": 90,
  "ZipReportFolder": false
}
```

### Example JSON: Full Monthly Cleanup

```json
{
  "Mode": "Remediate",
  "Full": true,
  "Lite": false,
  "ReportRoot": "C:\\PCMaintenance\\Reports",
  "DriveLetter": "C",
  "CreateAfterRestorePoint": false,
  "CleanOldReports": true,
  "ReportRetentionDays": 90,
  "ZipReportFolder": true,
  "MaxShadowStorage": "10GB"
}
```

### Example JSON: Email and Teams Reporting

```json
{
  "Mode": "Remediate",
  "Lite": true,
  "ReportRoot": "C:\\PCMaintenance\\Reports",
  "EmailReportTo": "support@example.com",
  "EmailFrom": "pcmaintenance@example.com",
  "SmtpServer": "smtp.office365.com",
  "SmtpPort": 587,
  "SmtpUseSsl": true,
  "TeamsWebhookUrl": "https://example.webhook.office.com/webhookb2/...",
  "NotificationSubjectPrefix": "PC Maintenance"
}
```

### Config Notes

- JSON property names must match script parameter names.
- Command-line switches can still be used alongside config.
- If a value is missing from the config, the script uses its built-in default.
- Keep secrets out of GitHub where possible. Avoid committing real webhook URLs or SMTP credentials.
- The script uses standard SMTP parameters; if your mail platform requires modern OAuth/Graph authentication, use a relay or RMM email delivery instead.

---

## Report Outputs

Each run creates a folder like:

```text
C:\PCMaintenance\Reports\20260510_103200-Cleanup\
```

Typical files include:

| File | Description |
|---|---|
| `PC_Maintenance_<Computer>_<timestamp>.html` | Main branded report. |
| `PC_Maintenance_<Computer>_<timestamp>.json` | Machine-readable summary. |
| `PC_Maintenance_<Computer>_<timestamp>.csv` | CSV summary for dashboards/RMM import. |
| `PC_Maintenance_Transcript_<Computer>_<timestamp>.txt` | PowerShell transcript. |
| `.evtx` files | Archived event logs when `-ArchiveEventLogs` is used. |
| Battery report | Generated when supported and not skipped. |
| `.zip` package | Created when `-ZipReportFolder` is used. |

### Main HTML Report Sections

The HTML report contains sections such as:

- Technician summary
- Health score
- Disk space before/after
- Per-step duration and disk delta
- VSS/shadow storage before/after
- Pre-flight checks
- Windows Update diagnostics
- BitLocker health
- Power/sleep health
- Component store analysis
- Cleanup category metrics
- System repair results
- Local user profile sizes
- Storage health
- Battery health
- Defender status
- Crash dumps
- Teams and OneDrive health
- Event interpretation
- Scheduled task audit
- Actions performed

---

## Email Reporting

Email reporting is controlled by these switches:

```powershell
-EmailReportTo
-EmailFrom
-SmtpServer
-SmtpPort
-SmtpUseSsl
-NotificationSubjectPrefix
```

### Basic SMTP Example

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 `
  -Mode Remediate -Lite `
  -EmailReportTo "support@example.com" `
  -EmailFrom "pcmaintenance@example.com" `
  -SmtpServer "smtp.office365.com" `
  -SmtpPort 587 `
  -SmtpUseSsl
```

### Email Requirements

- `-EmailReportTo` requires `-EmailFrom` and `-SmtpServer`.
- The script attaches the HTML, JSON, and CSV files where available.
- SMTP authentication requirements depend on your environment.
- For Microsoft 365 tenants with SMTP AUTH disabled, consider using:
  - an authenticated SMTP relay
  - a connector
  - RMM-native email reporting
  - a future Graph API mail implementation

### Email Security Notes

- Do not store SMTP credentials directly in the GitHub config.
- Prefer a secure relay, RMM credential store, or environment-specific secret management.
- Do not commit real email credentials or webhook URLs to a public repo.

---

## Microsoft Teams Webhook Reporting

Teams webhook reporting is controlled by:

```powershell
-TeamsWebhookUrl
```

Example:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 `
  -Mode Remediate -Lite `
  -TeamsWebhookUrl "https://example.webhook.office.com/webhookb2/..."
```

The Teams message includes a concise summary such as:

- computer name
- health score
- health status
- pending reboot state
- disk free after cleanup
- report folder path

### Teams Webhook Notes

- Treat webhook URLs as secrets.
- Do not commit real webhook URLs to GitHub.
- Some Microsoft Teams webhook connector experiences may vary depending on tenant configuration.
- If incoming webhooks are not allowed in your tenant, use email reporting or RMM-native notifications.

---

## Remote and RMM Usage

### Remote Execution with `-ComputerName`

The script includes a remote wrapper using PowerShell Remoting.

Example:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 `
  -ComputerName PC001,PC002 `
  -Mode Remediate -Lite
```

The remote wrapper:

- copies the current script to the remote machine via admin share
- uses the remote script path, defaulting to `C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1`
- invokes the remote script with passed arguments

### Remote Requirements

- PowerShell Remoting enabled
- firewall rules allowing remoting
- local/domain admin permissions
- admin share access, e.g. `\\PC001\C$`

### RMM-Friendly Exit Codes

Use:

```powershell
-UseRmmExitCode
```

Exit codes:

| Exit Code | Meaning |
|---:|---|
| `0` | Success / no major issue. |
| `1` | Completed with warnings. |
| `2` | Reboot required. |
| `3` | Critical issue found. |
| `4` | Script or report failure. |

### RMM Example

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -Mode Remediate -Lite -UseRmmExitCode
```

### RMM JSON Output

The JSON output is useful for custom fields and dashboards. It includes details such as:

- computer name
- user name
- mode
- health score
- disk before/after
- pending reboot
- CHKDSK/SFC/DISM summaries
- report path
- VSS before/after
- step metrics
- notification results

---

## Restore Points and VSS Shadow Storage

v15 changes restore point behaviour to reduce unexpected disk usage.

### Default Behaviour

- Creates a **BEFORE** restore point by default.
- Does **not** create an AFTER restore point unless `-CreateAfterRestorePoint` is supplied.
- Reports VSS/shadow storage before and after the run.

### Optional AFTER Restore Point

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -Mode Remediate -Lite -CreateAfterRestorePoint
```

### Optional Shadow Storage Limit

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -Mode Remediate -Lite -MaxShadowStorage 10GB
```

or:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -Mode Remediate -Lite -MaxShadowStorage 5%
```

### Why This Matters

Restore points can consume significant disk space. If disk space drops after cleanup, check the VSS section in the report or run:

```powershell
vssadmin list shadowstorage
```

---

## Safety Notes

### Run as Administrator

Many actions require admin rights, including:

- restore point creation
- event log archive/clear
- CHKDSK queueing
- SFC/DISM
- Windows Update
- profile removal
- Windows.old cleanup
- crash dump cleanup

### Destructive Actions

The following are intentionally switch-gated:

| Action | Required Switches |
|---|---|
| Clear event logs | `-ClearEventLogs` |
| Empty Recycle Bin | `-EmptyRecycleBin` |
| Clear browser cache | `-ClearBrowserCache` |
| Delete crash dumps | `-CleanCrashDumps` |
| Delete Windows.old | `-CleanWindowsOld` |
| Remove stale profiles | `-RemoveStaleProfiles -ForceProfileRemoval` |

### Windows.old Warning

`-CleanWindowsOld` can remove rollback files after a Windows feature update. Only use it when rollback is no longer required.

### Stale Profile Removal Warning

Profile removal is deliberately locked behind both:

```powershell
-RemoveStaleProfiles -ForceProfileRemoval
```

Only use after reviewing the profile report and confirming the profiles are no longer required.

### Restore Point Warning

System Restore must be enabled for restore point creation. If System Protection is disabled, the script records the failure in the report rather than stopping the run.

---

## Example Commands

### Parse-only validation

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -ValidateOnly
```

### Audit with no repair/update tasks

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -Mode Audit -SkipWindowsUpdate -SkipSfc -SkipDismRestoreHealth
```

### Audit everything possible

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -Mode Audit
```

### Lite cleanup

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -Mode Remediate -Lite
```

### Full cleanup

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -Mode Remediate -Full
```

### Full cleanup with report ZIP

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -Mode Remediate -Full -ZipReportFolder
```

### Lite cleanup and email report

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 `
  -Mode Remediate -Lite `
  -EmailReportTo "support@example.com" `
  -EmailFrom "pcmaintenance@example.com" `
  -SmtpServer "smtp.office365.com" `
  -SmtpPort 587 `
  -SmtpUseSsl
```

### Lite cleanup and Teams webhook report

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 `
  -Mode Remediate -Lite `
  -TeamsWebhookUrl "https://example.webhook.office.com/webhookb2/..."
```

### Run using JSON config

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -ConfigPath C:\PCMaintenance\pc_cleanup_config_template_v15.json
```

### Full cleanup with capped restore storage

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -Mode Remediate -Full -MaxShadowStorage 10GB
```

### Full cleanup with AFTER restore point

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -Mode Remediate -Full -CreateAfterRestorePoint
```

### Dangerous: remove old profiles

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 `
  -Mode Remediate -Lite `
  -RemoveStaleProfiles `
  -ForceProfileRemoval `
  -RemoveProfilesOlderThanDays 180
```

### Remote run

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 `
  -ComputerName PC001,PC002 `
  -Mode Remediate -Lite
```

---

## Version Notes

### v1

Initial monthly cleanup concept:

- temp cleanup
- disk space before/after
- event log review
- Windows Update attempt
- top processes
- HTML report

### v2-v3

Stability improvements:

- fixed Unicode/non-ASCII hyphen issues
- improved runtime error handling
- safer report generation
- better handling when individual sections fail

### v4

Added repair and restore actions:

- `chkdsk C: /f` queueing
- `sfc /scannow`
- DISM RestoreHealth
- local user profile size reporting
- BEFORE and AFTER restore points

### v5

Added endpoint health-check features:

- Audit/Remediate modes
- pending reboot detection
- smarter CHKDSK scan before queueing `/f`
- component cleanup
- Delivery Optimization cleanup
- browser cache cleanup
- Recycle Bin cleanup
- battery health
- storage health
- device problem checks
- Defender status
- startup/service checks
- crash dump summary
- storage hotspots
- health score
- JSON/CSV export
- RMM exit codes

### v6

Stability fixes:

- safer crash dump size calculations
- safer top CPU process handling
- safer technician summary counts
- reduced strict-mode sensitivity

### v7

Branding and output path improvements:

- logo support in top-right report header
- report title on the left
- default reports under `C:\PCMaintenance\Reports`
- logo embedded in HTML

### v8

Run folder organisation:

- created timestamped per-run folders
- folder suffix `-Audit` or `-Cleanup`
- report files grouped by run
- branded report styling refinements

### v9

Advanced automation/reporting features:

- email report delivery
- Teams webhook summary
- JSON config support
- deeper SMART/storage checks
- Teams/OneDrive health checks
- reliability/event interpretation
- before/after cleanup metrics by category
- optional crash dump cleanup
- optional Windows.old cleanup
- optional stale profile removal
- remote/RMM wrapper support

### v10

Remediation presets:

- added `-Full`
- added `-Lite`
- default Remediate behaviour applies Lite when no cleanup preset/switches are supplied

### v11-v14

Parser stability fixes:

- safer report string formatting
- safer health-score string handling
- corrected Teams webhook string issue
- added guidance for local parse testing

### v15

Professional maintenance/reporting improvements:

- VSS/shadow storage before/after reporting
- BEFORE restore point by default
- AFTER restore point only with `-CreateAfterRestorePoint`
- optional `-MaxShadowStorage`
- per-step disk delta tracking
- per-step duration tracking
- long-running watchdog output for SFC/DISM/component cleanup
- `-ValidateOnly` parse check
- pre-flight checks
- report retention cleanup
- optional report ZIP
- Windows Update diagnostics
- BitLocker health
- power/sleep health
- component store analysis

---

## Troubleshooting

### Script will not run due to execution policy

Run with:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -ValidateOnly
```

If downloaded from the internet:

```powershell
Unblock-File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1
```

### Validate syntax without running maintenance

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -ValidateOnly
```

### SFC or DISM appears stuck

Check running processes:

```powershell
Get-Process sfc, dism, TiWorker, TrustedInstaller -ErrorAction SilentlyContinue |
Select-Object Name, Id, CPU, WorkingSet, StartTime
```

Check CBS log activity:

```powershell
Get-Item C:\Windows\Logs\CBS\CBS.log |
Select-Object FullName, LastWriteTime, Length
```

Tail recent logs:

```powershell
Get-Content C:\Windows\Logs\CBS\CBS.log -Tail 30
Get-Content C:\Windows\Logs\DISM\dism.log -Tail 30
```

### Disk space dropped after cleanup

Check the VSS/shadow storage section in the report.

You can also run:

```powershell
vssadmin list shadowstorage
```

Restore points and VSS snapshots can consume multiple GB, especially after repair/component cleanup work.

### Logo does not appear in the report

Ensure the logo is in the same folder as the script.

Recommended names:

```text
C:\PCMaintenance\Silicon Beach.png
C:\PCMaintenance\logo.png
C:\PCMaintenance\SiliconBeach_logo.png
```

### Email did not send

Check:

- `-EmailReportTo` is supplied
- `-EmailFrom` is supplied
- `-SmtpServer` is supplied
- SMTP port is correct
- SSL setting is correct
- firewall allows SMTP outbound
- tenant allows SMTP AUTH or relay

### Teams webhook did not send

Check:

- webhook URL is valid
- the Teams connector/webhook is still enabled
- tenant allows incoming webhooks
- the URL has not been revoked
- outbound HTTPS is allowed

### Remote run failed

Check:

- PowerShell Remoting is enabled
- admin share is available
- DNS/name resolution works
- credentials have admin rights
- firewall allows WinRM
- remote machine can access `C:\PCMaintenance`

---

## Suggested Operational Workflow

### First-time testing on a machine

1. Copy script, config, and logo to `C:\PCMaintenance`.
2. Run `-ValidateOnly`.
3. Run Audit mode with SFC/DISM/updates skipped.
4. Review HTML report.
5. Run Lite remediation.
6. Review report and reboot if needed.

Example:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -ValidateOnly
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -Mode Audit -SkipWindowsUpdate -SkipSfc -SkipDismRestoreHealth
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -Mode Remediate -Lite
```

### Monthly maintenance run

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -Mode Remediate -Lite -CleanOldReports
```

### Quarterly deeper maintenance run

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -Mode Remediate -Full -CleanOldReports -ZipReportFolder
```

### MSP/RMM run

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1 -Mode Remediate -Lite -UseRmmExitCode -TeamsWebhookUrl "https://example.webhook.office.com/webhookb2/..."
```

---

## GitHub Notes

Recommended repository structure:

```text
.
├── README.md
├── pc_cleanup_maintenance_v15.ps1
├── pc_cleanup_config_template_v15.json
├── assets\
│   └── logo-example.png
└── examples\
    ├── audit-config.example.json
    ├── lite-cleanup-config.example.json
    └── full-cleanup-config.example.json
```

Do not commit:

- real Teams webhook URLs
- SMTP credentials
- customer machine reports
- logs containing private/customer data
- live customer config files with secrets

Use `.gitignore` for generated reports:

```gitignore
Reports/
*.evtx
*.zip
*.log
*_Transcript_*.txt
```

---

## Disclaimer

This script performs system maintenance and can make changes to local Windows devices when run in Remediate mode. Review the switches carefully, test in Audit mode first, and avoid destructive cleanup actions unless you understand their impact.

