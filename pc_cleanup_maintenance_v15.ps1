<#
.SYNOPSIS
    Monthly Windows PC maintenance, health-check, cleanup, repair and branded HTML report script.

.DESCRIPTION
    Designed for desktop PCs and laptops. It can run in Audit mode or Remediate mode.
    It checks disk space, cleans temp/cache files, audits or removes stale scheduled tasks,
    reviews event logs, optionally archives/clears event logs, checks Windows Update,
    runs CHKDSK/SFC/DISM, creates restore points, checks reboot state, hardware health,
    Defender status, startup items, service issues, crash dumps, user profile sizes,
    storage hotspots, Teams/OneDrive health, reliability interpretation, cleanup metrics, notifications, and creates HTML/JSON/CSV outputs.

    Use /? or -ShowHelp for switch and function guidance.

.NOTES
    Tested syntax target: Windows PowerShell 5.1+
    Run as Administrator for best results.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Alias('?','Help','h')]
    [switch]$ShowHelp,

    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ExtraArgs,

    [ValidateSet('Audit','Remediate')]
    [string]$Mode = 'Remediate',

    [switch]$AuditOnly,
    [switch]$Full,
    [switch]$Lite,
    [switch]$AggressiveCleanup,

    [ValidatePattern('^[A-Za-z]:?$')]
    [string]$DriveLetter = 'C',

    [string]$ReportRoot = '',

    [int]$TempFileMinAgeDays = 2,
    [int]$EventLookbackDays = 14,
    [int]$StaleTaskDays = 90,
    [int]$ProfileStaleDays = 90,
    [int]$TopProcessCount = 5,
    [int]$TopLargeFileCount = 20,

    [int]$CpuWarnPercent = 85,
    [int]$MemoryWarnPercent = 85,
    [double]$DiskQueueWarn = 2.0,
    [double]$LowDiskWarnPercent = 15,

    [switch]$RemoveStaleTasks,
    [switch]$ClearEventLogs,
    [switch]$ArchiveEventLogs,
    [switch]$EmptyRecycleBin,
    [switch]$ClearBrowserCache,
    [switch]$CleanWindowsOld,

    [switch]$SkipRecycleBin,
    [switch]$SkipComponentCleanup,
    [switch]$SkipDeliveryOptimizationCleanup,
    [switch]$SkipWindowsUpdate,
    [switch]$DownloadUpdatesOnly,
    [switch]$SkipRestorePoints,
    [switch]$SkipChkdsk,
    [switch]$AlwaysQueueChkdskFix,
    [switch]$SkipSfc,
    [switch]$SkipDismRestoreHealth,
    [switch]$SkipUserProfileSizes,
    [switch]$SkipStorageHotspots,
    [switch]$SkipBatteryReport,
    [switch]$SkipTranscript,

    [switch]$RunDefenderQuickScan,
    [switch]$UseRmmExitCode,

    [string]$ConfigPath = '',

    [string[]]$ComputerName = @(),
    [string]$RemoteScriptPath = 'C:\PCMaintenance\pc_cleanup_maintenance_v15.ps1',
    [string[]]$RemotePassthruArgs = @(),

    [string]$TeamsWebhookUrl = '',
    [string]$EmailReportTo = '',
    [string]$EmailFrom = '',
    [string]$SmtpServer = '',
    [int]$SmtpPort = 587,
    [switch]$SmtpUseSsl,
    [string]$NotificationSubjectPrefix = 'PC Maintenance',

    [switch]$CleanCrashDumps,
    [switch]$RemoveStaleProfiles,
    [switch]$ForceProfileRemoval,
    [int]$RemoveProfilesOlderThanDays = 180,

    [switch]$ValidateOnly,
    [switch]$CreateAfterRestorePoint,
    [string]$MaxShadowStorage = '',
    [int]$ReportRetentionDays = 90,
    [switch]$CleanOldReports,
    [switch]$ZipReportFolder,
    [int]$ComponentCleanupTimeoutMinutes = 90,
    [int]$SfcTimeoutMinutes = 120,
    [int]$DismTimeoutMinutes = 180,
    [int]$ChkdskScanTimeoutMinutes = 60,
    [switch]$SkipComponentStoreAnalysis
)

Set-StrictMode -Off
$ErrorActionPreference = 'Continue'
$script:RunStart = Get-Date
$script:ActionLog = @()
$script:IsAuditOnly = $false
$script:TranscriptPath = $null
$script:ReportFiles = @{}
$script:StepMetrics = @()
$script:VssBefore = @()
$script:VssAfter = @()
$script:PreflightChecks = @()
$script:WindowsUpdateDiagnostics = @()
$script:BitLockerHealth = @()
$script:PowerHealth = @()
$script:ComponentStoreAnalysis = $null

function Show-PCMaintenanceHelp {
    $help = @'
PC Cleanup Maintenance v15 - Help

BASIC USAGE
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\pc_cleanup_maintenance_v15.ps1
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\pc_cleanup_maintenance_v15.ps1 /?
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\pc_cleanup_maintenance_v15.ps1 -ShowHelp
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\pc_cleanup_maintenance_v15.ps1 -ValidateOnly

RUN MODES
  -Mode Audit
      Report-only mode. The script gathers data and estimates cleanup potential, but avoids destructive actions.

  -Mode Remediate
      Active maintenance mode. This is the default to preserve the previous monthly cleanup behaviour.

  -AuditOnly
      Shortcut for -Mode Audit.

REMEDIATE PRESETS
  -Lite
      Light cleanup preset for Remediate mode. Enables:
        -RemoveStaleTasks -ArchiveEventLogs -ClearEventLogs

  -Full
      Full cleanup preset for Remediate mode. Enables:
        -RemoveStaleTasks -ArchiveEventLogs -ClearEventLogs -EmptyRecycleBin -ClearBrowserCache -CleanCrashDumps -CleanWindowsOld

  Default Remediate behaviour
      If Remediate mode is used without -Full, -Lite or any explicit cleanup switches, the script defaults to the Lite preset.

COMMON MAINTENANCE SWITCHES
  -RemoveStaleTasks
      Removes eligible stale scheduled tasks. By default stale tasks are audited only.
      The script avoids Microsoft task paths and focuses on disabled non-Microsoft tasks.

  -ClearEventLogs
      Clears Application/System logs after notable events have been captured. Off by default.

  -ArchiveEventLogs
      Exports Application/System logs to the report folder before clearing. Use with -ClearEventLogs.

  -EmptyRecycleBin
      Clears the Recycle Bin. Off by default.

  -ClearBrowserCache
      Includes Edge, Chrome and Firefox cache paths in cleanup. Does not target cookies, passwords or bookmarks.

  -CleanWindowsOld
      Attempts Windows.old cleanup through DISM component cleanup reporting. Kept conservative by default.

  -AggressiveCleanup
      Enables more assertive cleanup choices. It does not remove user profiles automatically.

WINDOWS UPDATE SWITCHES
  -SkipWindowsUpdate
      Skips update scan/install.

  -DownloadUpdatesOnly
      Downloads available updates where the provider supports it, but does not install.

REPAIR / HEALTH SWITCHES
  -SkipChkdsk
      Skips CHKDSK.

  -AlwaysQueueChkdskFix
      Queues chkdsk C: /f regardless of the online scan result.
      Without this, v5 runs chkdsk C: /scan first and queues /f only if repair appears necessary.

  -SkipSfc
      Skips sfc /scannow.

  -SkipDismRestoreHealth
      Skips dism /online /cleanup-image /restorehealth /norestart.

  -SkipRestorePoints
      Skips restore point creation.

  -CreateAfterRestorePoint
      Creates the AFTER restore point. v15 creates the BEFORE restore point by default and makes AFTER optional.

  -MaxShadowStorage <size>
      Optionally caps System Restore/VSS storage, e.g. 10GB or 5%.

  -RunDefenderQuickScan
      Runs a Defender quick scan if Defender cmdlets are available.

REPORT / OUTPUT SWITCHES
  -ReportRoot <path>
      Base output folder. Default: .\Reports under the script folder (for example C:\PCMaintenance\Reports). Each run creates a timestamped child folder ending in -Audit or -Cleanup.

  -SkipTranscript
      Does not create a PowerShell transcript.

  -ValidateOnly
      Runs a self-parse check and exits without maintenance.

  -CleanOldReports
      Removes old run folders under the report root using -ReportRetentionDays.

  -ZipReportFolder
      Compresses the run folder into a ZIP at the end.

  -UseRmmExitCode
      Exits with RMM-friendly codes:
        0 = success/no major issue
        1 = completed with warnings
        2 = reboot required
        3 = critical issue found
        4 = script/report failure

V9 CONFIG / NOTIFICATION / ADVANCED CLEANUP SWITCHES
  -ConfigPath <path>
      Loads defaults from a JSON config file. Matching property names are applied before the run.

  -TeamsWebhookUrl <url>
      Sends a summary card to an incoming Teams webhook after report generation.

  -EmailReportTo <address>
      Sends the HTML report by email. Requires -SmtpServer and -EmailFrom.

  -CleanCrashDumps
      Deletes C:\Windows\Minidump files and C:\Windows\MEMORY.DMP in Remediate mode.

  -CleanWindowsOld
      Deletes C:\Windows.old in Remediate mode. This removes rollback files.

  -RemoveStaleProfiles -ForceProfileRemoval
      Removes unloaded, non-special local profiles older than -RemoveProfilesOlderThanDays.
      This is deliberately locked behind both switches.

  -ComputerName PC1,PC2
      Runs this script remotely through PowerShell Remoting. Requires admin access and remoting.

KEY FUNCTIONS INSIDE THE SCRIPT
  Get-DiskSpace                  Disk free/used/percent before and after cleanup.
  Invoke-TempCleanup             Temp/cache cleanup with audit/remediate awareness.
  Get-PendingRebootStatus        Windows Update/CBS/PendingFileRename/computer rename checks.
  Invoke-ChkdskSmart             Runs chkdsk /scan and conditionally queues chkdsk /f.
  Invoke-SfcScan                 Runs sfc /scannow and captures output.
  Invoke-DismRestoreHealth       Runs DISM RestoreHealth and captures output.
  Get-LocalUserProfileSizes      Calculates all local profile sizes and stale-profile warnings.
  Get-StorageHealth              Disk/physical disk health and reliability counters.
  Get-BatteryHealth              Laptop battery presence and battery report generation.
  Get-ProblemDevices             Device Manager devices with non-OK status.
  Get-DefenderStatus             Microsoft Defender status and optional quick scan.
  Get-ServiceHealthIssues        Auto-start services that are stopped.
  Get-CrashDumpSummary           Minidump and MEMORY.DMP summary.
  Get-StorageHotspots            Large files in Downloads/Desktop/Documents/Videos/ISO-like locations.
  Get-TeamsOneDriveHealth        Teams/OneDrive cache and sync state checks.
  Get-EventInterpretation        Reliability-style explanation of notable events.
  Get-CleanupCategorySnapshot    Before/after cleanup metrics per category.
  Send-MaintenanceNotifications  Email and Teams webhook report delivery.
  Invoke-RemoteMaintenance       Remote/RMM wrapper support.
  New-MaintenanceReport          Creates branded HTML, JSON and CSV summary outputs.
  New-MaintenanceRunFolder       Creates one timestamped output subfolder per run.

SAFER TESTING EXAMPLES
  .\pc_cleanup_maintenance_v15.ps1 -Mode Audit -SkipWindowsUpdate -SkipSfc -SkipDismRestoreHealth
  .\pc_cleanup_maintenance_v15.ps1 -Mode Remediate -Lite -SkipWindowsUpdate

AGGRESSIVE MONTHLY EXAMPLE
  .\pc_cleanup_maintenance_v15.ps1 -Mode Remediate -Full -AlwaysQueueChkdskFix
'@
    Write-Host $help
}

if ($ShowHelp -or ($ExtraArgs -contains '/?') -or ($ExtraArgs -contains '-?') -or ($ExtraArgs -contains 'help') -or ($ExtraArgs -contains '--help')) {
    Show-PCMaintenanceHelp
    return
}


function Import-PCMaintenanceConfig {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Warning ("Config file not found: {0}" -f $Path)
        return
    }
    try {
        $config = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        foreach ($prop in @($config.PSObject.Properties)) {
            $name = $prop.Name
            $value = $prop.Value
            if (Get-Variable -Name $name -Scope Script -ErrorAction SilentlyContinue) {
                Set-Variable -Name $name -Value $value -Scope Script -ErrorAction SilentlyContinue
            }
        }
        Write-Host ("Loaded config: {0}" -f $Path) -ForegroundColor Cyan
    } catch {
        Write-Warning ("Failed to load config file '{0}': {1}" -f $Path, $_.Exception.Message)
    }
}

function Add-RemoteArgValue {
    param([string[]]$Args,[string]$Name,[object]$Value)
    if ($null -eq $Value) { return $Args }
    if ($Value -is [System.Array]) {
        foreach ($v in $Value) { if ($null -ne $v -and [string]$v -ne '') { $Args += $Name; $Args += [string]$v } }
    } elseif ([string]$Value -ne '') {
        $Args += $Name
        $Args += [string]$Value
    }
    return $Args
}

function New-RemoteArgumentArray {
    $argsOut = @('-Mode', $Mode, '-DriveLetter', $DriveLetter, '-TempFileMinAgeDays', [string]$TempFileMinAgeDays, '-EventLookbackDays', [string]$EventLookbackDays, '-StaleTaskDays', [string]$StaleTaskDays, '-ProfileStaleDays', [string]$ProfileStaleDays)
    foreach ($sw in @(
        'AuditOnly','Full','Lite','AggressiveCleanup','RemoveStaleTasks','ClearEventLogs','ArchiveEventLogs','EmptyRecycleBin','ClearBrowserCache','CleanWindowsOld',
        'SkipRecycleBin','SkipComponentCleanup','SkipDeliveryOptimizationCleanup','SkipWindowsUpdate','DownloadUpdatesOnly','SkipRestorePoints','SkipChkdsk','AlwaysQueueChkdskFix','SkipSfc','SkipDismRestoreHealth','SkipUserProfileSizes','SkipStorageHotspots','SkipBatteryReport','SkipTranscript','RunDefenderQuickScan','UseRmmExitCode','CleanCrashDumps','RemoveStaleProfiles','ForceProfileRemoval','ValidateOnly','CreateAfterRestorePoint','CleanOldReports','ZipReportFolder','SkipComponentStoreAnalysis'
    )) {
        try { if ((Get-Variable -Name $sw -Scope Script -ErrorAction SilentlyContinue).Value) { $argsOut += ('-' + $sw) } } catch { }
    }
    if (-not [string]::IsNullOrWhiteSpace($TeamsWebhookUrl)) { $argsOut += '-TeamsWebhookUrl'; $argsOut += $TeamsWebhookUrl }
    if (-not [string]::IsNullOrWhiteSpace($EmailReportTo)) { $argsOut += '-EmailReportTo'; $argsOut += $EmailReportTo }
    if (-not [string]::IsNullOrWhiteSpace($EmailFrom)) { $argsOut += '-EmailFrom'; $argsOut += $EmailFrom }
    if (-not [string]::IsNullOrWhiteSpace($SmtpServer)) { $argsOut += '-SmtpServer'; $argsOut += $SmtpServer; $argsOut += '-SmtpPort'; $argsOut += [string]$SmtpPort }
    if ($SmtpUseSsl) { $argsOut += '-SmtpUseSsl' }
    if (-not [string]::IsNullOrWhiteSpace($NotificationSubjectPrefix)) { $argsOut += '-NotificationSubjectPrefix'; $argsOut += $NotificationSubjectPrefix }
    if ($RemotePassthruArgs) { $argsOut += $RemotePassthruArgs }
    return [string[]]$argsOut
}

function Invoke-RemoteMaintenance {
    param([string[]]$Targets)
    $self = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
    if ([string]::IsNullOrWhiteSpace($self) -or -not (Test-Path -LiteralPath $self)) { throw 'Unable to determine current script path for remote copy.' }
    $remoteArgs = New-RemoteArgumentArray
    foreach ($target in @($Targets)) {
        if ([string]::IsNullOrWhiteSpace($target)) { continue }
        Write-Host ("Preparing remote maintenance on {0}" -f $target) -ForegroundColor Cyan
        try {
            $adminSharePath = ('\\{0}\C$\PCMaintenance' -f $target)
            if (-not (Test-Path -LiteralPath $adminSharePath)) { New-Item -Path $adminSharePath -ItemType Directory -Force | Out-Null }
            $dest = Join-Path $adminSharePath (Split-Path $RemoteScriptPath -Leaf)
            Copy-Item -LiteralPath $self -Destination $dest -Force
            Invoke-Command -ComputerName $target -ScriptBlock {
                param([string]$ScriptPath,[string[]]$RunArgs)
                if (-not (Test-Path -LiteralPath (Split-Path $ScriptPath -Parent))) { New-Item -Path (Split-Path $ScriptPath -Parent) -ItemType Directory -Force | Out-Null }
                & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $ScriptPath @RunArgs
            } -ArgumentList $RemoteScriptPath, $remoteArgs
        } catch {
            Write-Warning ("Remote maintenance failed on {0}: {1}" -f $target, $_.Exception.Message)
        }
    }
}

function Set-ScriptSwitchTrue {
    param([string]$Name)
    Set-Variable -Name $Name -Value $true -Scope Script -ErrorAction SilentlyContinue
}

function Set-MaintenancePreset {
    if ($script:IsAuditOnly) {
        if ($Full -or $Lite) {
            Add-ActionLog -Step 'Cleanup preset' -Status 'Audit' -Details 'Full/Lite presets were supplied but remediation actions are disabled in Audit mode.'
        }
        $script:MaintenancePreset = 'Audit'
        return
    }

    if ($Full -and $Lite) {
        Add-ActionLog -Step 'Cleanup preset' -Status 'Warning' -Details 'Both -Full and -Lite were supplied. Full preset will be used.'
        Set-Variable -Name Lite -Value $false -Scope Script -ErrorAction SilentlyContinue
    }

    $explicitCleanupSwitches = @(
        $RemoveStaleTasks,
        $ArchiveEventLogs,
        $ClearEventLogs,
        $EmptyRecycleBin,
        $ClearBrowserCache,
        $CleanCrashDumps,
        $CleanWindowsOld
    ) | Where-Object { [bool]$_ }

    if ($Full) {
        foreach ($sw in @('RemoveStaleTasks','ArchiveEventLogs','ClearEventLogs','EmptyRecycleBin','ClearBrowserCache','CleanCrashDumps','CleanWindowsOld')) {
            Set-ScriptSwitchTrue -Name $sw
        }
        $script:MaintenancePreset = 'Full'
        Add-ActionLog -Step 'Cleanup preset' -Status 'Full' -Details 'Full preset enabled: stale tasks, event archive/clear, recycle bin, browser cache, crash dumps and Windows.old cleanup.'
        return
    }

    if ($Lite) {
        foreach ($sw in @('RemoveStaleTasks','ArchiveEventLogs','ClearEventLogs')) {
            Set-ScriptSwitchTrue -Name $sw
        }
        $script:MaintenancePreset = 'Lite'
        Add-ActionLog -Step 'Cleanup preset' -Status 'Lite' -Details 'Lite preset enabled: stale tasks and event log archive/clear.'
        return
    }

    if (@($explicitCleanupSwitches).Count -eq 0) {
        foreach ($sw in @('RemoveStaleTasks','ArchiveEventLogs','ClearEventLogs')) {
            Set-ScriptSwitchTrue -Name $sw
        }
        $script:MaintenancePreset = 'Lite(default)'
        Add-ActionLog -Step 'Cleanup preset' -Status 'Lite(default)' -Details 'Remediate mode was used without -Full, -Lite or explicit cleanup switches, so Lite preset was applied.'
        return
    }

    $script:MaintenancePreset = 'Custom'
    Add-ActionLog -Step 'Cleanup preset' -Status 'Custom' -Details 'Explicit cleanup switches were supplied, so no preset was auto-applied.'
}

function Resolve-ReportRoot {
    param([string]$RequestedPath)
    if (-not [string]::IsNullOrWhiteSpace($RequestedPath)) {
        return $RequestedPath
    }
    if ($PSScriptRoot) {
        return (Join-Path $PSScriptRoot 'Reports')
    }
    return 'C:\PCMaintenance\Reports'
}

function New-MaintenanceRunFolder {
    param(
        [string]$BaseReportRoot,
        [bool]$AuditMode
    )

    $runType = if ($AuditMode) { 'Audit' } else { 'Cleanup' }
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $folderName = ('{0}-{1}' -f $timestamp, $runType)
    $runFolder = Join-Path $BaseReportRoot $folderName

    if (-not (Test-Path -LiteralPath $BaseReportRoot)) {
        New-Item -Path $BaseReportRoot -ItemType Directory -Force | Out-Null
    }

    if (-not (Test-Path -LiteralPath $runFolder)) {
        New-Item -Path $runFolder -ItemType Directory -Force | Out-Null
    }

    return [PSCustomObject]@{
        BaseReportRoot = $BaseReportRoot
        RunFolder      = $runFolder
        FolderName     = $folderName
        RunType        = $runType
        Timestamp      = $timestamp
    }
}

function Get-LogoDataUri {
    param([string]$FolderPath)

    if ([string]::IsNullOrWhiteSpace($FolderPath) -or -not (Test-Path -LiteralPath $FolderPath)) {
        return ''
    }

    $preferredNames = @(
        'logo.png','logo.jpg','logo.jpeg','logo.gif','logo.webp',
        'Silicon Beach.png','Silicon Beach.jpg','Silicon Beach.jpeg','Silicon Beach.webp',
        'SiliconBeach.png','SiliconBeach.jpg','SiliconBeach.jpeg','SiliconBeach.webp',
        'siliconbeach_logo.png','siliconbeach_logo.jpg','siliconbeach_logo.jpeg','siliconbeach_logo.webp'
    )

    $logoFile = $null
    foreach ($name in $preferredNames) {
        $candidate = Join-Path $FolderPath $name
        if (Test-Path -LiteralPath $candidate) {
            $logoFile = Get-Item -LiteralPath $candidate -ErrorAction SilentlyContinue
            if ($logoFile) { break }
        }
    }

    if (-not $logoFile) {
        $logoFile = Get-ChildItem -LiteralPath $FolderPath -File -ErrorAction SilentlyContinue |
            Where-Object {
                $_.BaseName -match '(?i)(logo|silicon\s*beach|siliconbeach|^sb$)' -and $_.Extension -match '^\.(png|jpg|jpeg|gif|webp)$'
            } |
            Sort-Object Name |
            Select-Object -First 1
    }

    if (-not $logoFile) {
        return ''
    }

    try {
        $bytes = [System.IO.File]::ReadAllBytes($logoFile.FullName)
        $base64 = [System.Convert]::ToBase64String($bytes)
        switch ($logoFile.Extension.ToLowerInvariant()) {
            '.png'  { $mime = 'image/png' }
            '.jpg'  { $mime = 'image/jpeg' }
            '.jpeg' { $mime = 'image/jpeg' }
            '.gif'  { $mime = 'image/gif' }
            '.webp' { $mime = 'image/webp' }
            default { $mime = 'application/octet-stream' }
        }
        return ('data:{0};base64,{1}' -f $mime, $base64)
    }
    catch {
        Add-ActionLog -Step 'Logo load' -Status 'Warning' -Details ("Failed to embed logo from '{0}': {1}" -f $logoFile.FullName, $_.Exception.Message)
        return ''
    }
}

function Add-ActionLog {
    param(
        [string]$Step,
        [string]$Status,
        [string]$Details = ''
    )

    $script:ActionLog += [PSCustomObject]@{
        Time    = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Step    = $Step
        Status  = $Status
        Details = $Details
    }
}

function Write-Step {
    param([string]$Message)
    Write-Host "[$((Get-Date).ToString('HH:mm:ss'))] $Message" -ForegroundColor Cyan
}

function Test-IsAdmin {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}

function ConvertTo-GB {
    param([object]$Bytes)
    if ($null -eq $Bytes) { return 0 }
    try { return [math]::Round(([double]$Bytes / 1GB), 2) } catch { return 0 }
}

function ConvertTo-MB {
    param([object]$Bytes)
    if ($null -eq $Bytes) { return 0 }
    try { return [math]::Round(([double]$Bytes / 1MB), 2) } catch { return 0 }
}


function Get-ObjectCountSafe {
    param([object]$InputObject)
    if ($null -eq $InputObject) { return 0 }
    try { return @($InputObject).Count } catch { return 0 }
}

function Get-PropertyValueSafe {
    param([object]$Object,[string]$PropertyName,[object]$DefaultValue = $null)
    if ($null -eq $Object -or [string]::IsNullOrWhiteSpace($PropertyName)) { return $DefaultValue }
    try {
        $prop = $Object.PSObject.Properties[$PropertyName]
        if ($null -ne $prop) { return $prop.Value }
    } catch { }
    return $DefaultValue
}

function Limit-Text {
    param([string]$Text, [int]$MaxLength = 20000)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    if ($Text.Length -le $MaxLength) { return $Text }
    return $Text.Substring(0, $MaxLength) + "`r`n... truncated ..."
}

function ConvertTo-HtmlEncodedText {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    try { return [System.Net.WebUtility]::HtmlEncode($Text) } catch { return ($Text -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;') }
}

function ConvertTo-HtmlTable {
    param(
        [object[]]$Data,
        [string]$EmptyMessage = 'None'
    )
    if (-not $Data -or @($Data).Count -eq 0) { return "<p>$EmptyMessage</p>" }
    return (($Data | ConvertTo-Html -Fragment) -join [Environment]::NewLine)
}

function New-BarHtml {
    param([string]$Label, [object]$Percent)
    $num = 0
    if ($null -ne $Percent -and $Percent -ne '') { try { $num = [double]$Percent } catch { $num = 0 } }
    $clamped = [math]::Min([math]::Max($num, 0), 100)
    return ("<div class='metric-label'>{0}: {1}%</div><div class='bar'><div class='bar-fill' style='width:{2}%'></div></div>" -f $Label, $num, $clamped)
}

function Get-DiskSpace {
    param([string]$Letter = 'C')
    $cleanLetter = $Letter.TrimEnd(':')
    $deviceId = '{0}:' -f $cleanLetter
    $filter = "DeviceID='$deviceId'"
    try { $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter $filter -ErrorAction Stop }
    catch { $disk = Get-WmiObject -Class Win32_LogicalDisk -Filter $filter -ErrorAction Stop }
    return [PSCustomObject]@{
        Drive       = $disk.DeviceID
        SizeGB      = ConvertTo-GB $disk.Size
        UsedGB      = ConvertTo-GB ($disk.Size - $disk.FreeSpace)
        FreeGB      = ConvertTo-GB $disk.FreeSpace
        FreePercent = if ([double]$disk.Size -gt 0) { [math]::Round(([double]$disk.FreeSpace / [double]$disk.Size) * 100, 2) } else { 0 }
    }
}

function Get-FolderSizeInfo {
    param([string]$Path)
    $bytes = [int64]0
    $files = 0
    $errors = 0
    $exists = $false
    if (-not [string]::IsNullOrWhiteSpace($Path) -and (Test-Path -LiteralPath $Path)) {
        $exists = $true
        try {
            Get-ChildItem -LiteralPath $Path -Recurse -Force -File -ErrorAction SilentlyContinue | ForEach-Object {
                $files++
                try { $bytes += [int64]$_.Length } catch { $errors++ }
            }
        } catch { $errors++ }
    }
    return [PSCustomObject]@{ Path=$Path; Exists=$exists; SizeBytes=$bytes; SizeGB=(ConvertTo-GB $bytes); SizeMB=(ConvertTo-MB $bytes); FileCount=$files; Errors=$errors }
}

function Get-CleanupPaths {
    $paths = @()
    if ($env:TEMP) { $paths += $env:TEMP }
    if ($env:TMP -and $env:TMP -ne $env:TEMP) { $paths += $env:TMP }
    if ($env:windir) { $paths += (Join-Path $env:windir 'Temp') }
    if ($env:windir) { $paths += (Join-Path $env:windir 'SoftwareDistribution\Download') }

    if ($env:ProgramData) {
        $paths += (Join-Path $env:ProgramData 'Microsoft\Windows\WER\ReportArchive')
        $paths += (Join-Path $env:ProgramData 'Microsoft\Windows\WER\ReportQueue')
    }

    $profileRoot = Join-Path $env:SystemDrive 'Users'
    if (Test-Path -LiteralPath $profileRoot) {
        $profiles = Get-ChildItem -LiteralPath $profileRoot -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -notin @('Default','Default User','Public','All Users') }
        foreach ($profile in @($profiles)) {
            $paths += (Join-Path $profile.FullName 'AppData\Local\Temp')
            $paths += (Join-Path $profile.FullName 'AppData\Local\Microsoft\Windows\INetCache')
            $paths += (Join-Path $profile.FullName 'AppData\Local\Microsoft\Windows\Explorer')
            if ($ClearBrowserCache -or $AggressiveCleanup) {
                $paths += (Join-Path $profile.FullName 'AppData\Local\Microsoft\Edge\User Data\Default\Cache\Cache_Data')
                $paths += (Join-Path $profile.FullName 'AppData\Local\Microsoft\Edge\User Data\Default\Code Cache')
                $paths += (Join-Path $profile.FullName 'AppData\Local\Google\Chrome\User Data\Default\Cache\Cache_Data')
                $paths += (Join-Path $profile.FullName 'AppData\Local\Google\Chrome\User Data\Default\Code Cache')
                $paths += (Join-Path $profile.FullName 'AppData\Local\Mozilla\Firefox\Profiles')
            }
        }
    }
    return @($paths | Where-Object { $_ } | Sort-Object -Unique)
}

function Remove-OldFilesFromPath {
    param([string]$Path, [datetime]$Cutoff)
    $result = [PSCustomObject]@{
        Path            = $Path
        Exists          = $false
        FilesFound      = 0
        FilesRemoved    = 0
        PotentialGB     = 0
        RemovedGB       = 0
        Errors          = 0
        Mode            = if ($script:IsAuditOnly) { 'Audit' } else { 'Remediate' }
    }
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return $result }
    $result.Exists = $true
    $potentialBytes = [int64]0
    $removedBytes = [int64]0
    try {
        $files = Get-ChildItem -LiteralPath $Path -Recurse -Force -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt $Cutoff }
        foreach ($file in @($files)) {
            $result.FilesFound++
            $size = [int64]0
            try { $size = [int64]$file.Length } catch { $size = 0 }
            $potentialBytes += $size
            if (-not $script:IsAuditOnly) {
                try {
                    Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                    $result.FilesRemoved++
                    $removedBytes += $size
                } catch { $result.Errors++ }
            }
        }
        if (-not $script:IsAuditOnly) {
            $folders = Get-ChildItem -LiteralPath $Path -Recurse -Force -Directory -ErrorAction SilentlyContinue | Sort-Object FullName -Descending
            foreach ($folder in @($folders)) {
                try {
                    $children = Get-ChildItem -LiteralPath $folder.FullName -Force -ErrorAction SilentlyContinue
                    if (-not $children) { Remove-Item -LiteralPath $folder.FullName -Force -ErrorAction SilentlyContinue }
                } catch { }
            }
        }
    } catch { $result.Errors++ }
    $result.PotentialGB = ConvertTo-GB $potentialBytes
    $result.RemovedGB = ConvertTo-GB $removedBytes
    return $result
}

function Invoke-TempCleanup {
    param([int]$MinAgeDays = 2)
    $cutoff = (Get-Date).AddDays(-$MinAgeDays)
    $results = @()
    foreach ($path in @(Get-CleanupPaths)) { $results += Remove-OldFilesFromPath -Path $path -Cutoff $cutoff }
    $removedGB = ($results | Measure-Object -Property RemovedGB -Sum).Sum
    $potentialGB = ($results | Measure-Object -Property PotentialGB -Sum).Sum
    if ($null -eq $removedGB) { $removedGB = 0 }
    if ($null -eq $potentialGB) { $potentialGB = 0 }
    if ($script:IsAuditOnly) { Add-ActionLog -Step 'Temp and cache cleanup' -Status 'Audit' -Details ("Potential reclaimable: {0:N2} GB older than {1} day(s)." -f $potentialGB, $MinAgeDays) }
    else { Add-ActionLog -Step 'Temp and cache cleanup' -Status 'Complete' -Details ("Removed {0:N2} GB from files older than {1} day(s)." -f $removedGB, $MinAgeDays) }
    return @($results)
}

function Invoke-RecycleBinCleanup {
    if ($SkipRecycleBin) { Add-ActionLog -Step 'Recycle Bin' -Status 'Skipped' -Details 'SkipRecycleBin was supplied.'; return 'Skipped' }
    if (-not $EmptyRecycleBin -and -not $AggressiveCleanup) { Add-ActionLog -Step 'Recycle Bin' -Status 'Skipped' -Details 'Use -EmptyRecycleBin or -AggressiveCleanup to clear it.'; return 'Skipped; not requested' }
    if ($script:IsAuditOnly) { Add-ActionLog -Step 'Recycle Bin' -Status 'Audit' -Details 'Recycle Bin would be cleared in Remediate mode.'; return 'Audit only' }
    try {
        $cmd = Get-Command Clear-RecycleBin -ErrorAction SilentlyContinue
        if ($cmd) { Clear-RecycleBin -Force -ErrorAction SilentlyContinue; Add-ActionLog -Step 'Recycle Bin' -Status 'Complete' -Details 'Recycle Bin cleared.'; return 'Cleared' }
        Add-ActionLog -Step 'Recycle Bin' -Status 'Skipped' -Details 'Clear-RecycleBin command unavailable.'; return 'Unavailable'
    } catch { Add-ActionLog -Step 'Recycle Bin' -Status 'Failed' -Details $_.Exception.Message; return ('Failed: ' + $_.Exception.Message) }
}

function Invoke-ComponentCleanup {
    if ($SkipComponentCleanup) { Add-ActionLog -Step 'Component store cleanup' -Status 'Skipped' -Details 'SkipComponentCleanup was supplied.'; return 'Skipped' }
    if ($script:IsAuditOnly) { Add-ActionLog -Step 'Component store cleanup' -Status 'Audit' -Details 'DISM StartComponentCleanup would run in Remediate mode.'; return 'Audit only' }
    $dism = Join-Path $env:windir 'System32\dism.exe'
    if (-not (Test-Path -LiteralPath $dism)) { $dism = 'dism.exe' }
    try {
        $arguments = @('/Online','/Cleanup-Image','/StartComponentCleanup','/NoRestart')
        if ($AggressiveCleanup) { $arguments += '/ResetBase' }
        $componentResult = Invoke-ExternalMaintenanceCommand -Name 'DISM_StartComponentCleanup' -FilePath $dism -Arguments $arguments -TimeoutMinutes $ComponentCleanupTimeoutMinutes
        Add-ActionLog -Step 'Component store cleanup' -Status $componentResult.Status -Details ("DISM exit code: {0}" -f $componentResult.ExitCode)
        return ("Status: {0}; Exit code: {1}; Arguments: {2}" -f $componentResult.Status, $componentResult.ExitCode, ($arguments -join ' '))
    } catch { Add-ActionLog -Step 'Component store cleanup' -Status 'Failed' -Details $_.Exception.Message; return $_.Exception.Message }
}

function Invoke-DeliveryOptimizationCleanup {
    if ($SkipDeliveryOptimizationCleanup) { Add-ActionLog -Step 'Delivery Optimization cleanup' -Status 'Skipped' -Details 'SkipDeliveryOptimizationCleanup was supplied.'; return 'Skipped' }
    if ($script:IsAuditOnly) { Add-ActionLog -Step 'Delivery Optimization cleanup' -Status 'Audit' -Details 'Cache cleanup would run in Remediate mode.'; return 'Audit only' }
    try {
        $cmd = Get-Command Delete-DeliveryOptimizationCache -ErrorAction SilentlyContinue
        if ($cmd) { Delete-DeliveryOptimizationCache -Force -ErrorAction SilentlyContinue; Add-ActionLog -Step 'Delivery Optimization cleanup' -Status 'Complete' -Details 'Delivery Optimization cache cleanup requested.'; return 'Requested cleanup' }
        Add-ActionLog -Step 'Delivery Optimization cleanup' -Status 'Skipped' -Details 'Delete-DeliveryOptimizationCache unavailable.'
        return 'Cmdlet unavailable'
    } catch { Add-ActionLog -Step 'Delivery Optimization cleanup' -Status 'Failed' -Details $_.Exception.Message; return ('Failed: ' + $_.Exception.Message) }
}

function Get-NotableEvents {
    param([int]$LookbackDays = 14)
    $start = (Get-Date).AddDays(-$LookbackDays)
    $events = @()
    foreach ($logName in @('System','Application')) {
        try {
            $events += Get-WinEvent -FilterHashtable @{ LogName=$logName; Level=@(1,2,3); StartTime=$start } -ErrorAction SilentlyContinue |
                Select-Object -First 100 TimeCreated, LogName, LevelDisplayName, Id, ProviderName, Message
        } catch { }
    }
    Add-ActionLog -Step 'Event log review' -Status 'Complete' -Details ("Found {0} notable events." -f @($events).Count)
    return @($events | Sort-Object TimeCreated -Descending)
}

function Get-EventSummary {
    param([object[]]$Events)
    return @($Events | Group-Object LogName, ProviderName, Id, LevelDisplayName | Sort-Object Count -Descending | Select-Object -First 25 @{Name='Event';Expression={$_.Name}}, Count)
}

function Invoke-EventLogMaintenance {
    param([switch]$Clear, [switch]$Archive)
    $logs = @('Application','System')
    $results = @()
    foreach ($log in $logs) {
        $result = [PSCustomObject]@{ LogName=$log; Archived=$false; Cleared=$false; ArchivePath=''; Status='Skipped'; Details='' }
        try {
            if ($Archive) {
                $archivePath = Join-Path $ReportRoot ("{0}_{1}_{2}.evtx" -f $env:COMPUTERNAME, $log, (Get-Date -Format 'yyyyMMdd_HHmmss'))
                if ($script:IsAuditOnly) { $result.Details += 'Audit: archive would be created. ' }
                else { wevtutil epl $log $archivePath 2>$null; $result.Archived = Test-Path -LiteralPath $archivePath; $result.ArchivePath = $archivePath }
            }
            if ($Clear) {
                if ($script:IsAuditOnly) { $result.Details += 'Audit: log would be cleared. ' }
                else { Clear-EventLog -LogName $log -ErrorAction SilentlyContinue; $result.Cleared = $true }
            }
            if ($Clear -or $Archive) { $result.Status = 'Complete' } else { $result.Status = 'Not requested' }
        } catch { $result.Status = 'Failed'; $result.Details += $_.Exception.Message }
        $results += $result
    }
    Add-ActionLog -Step 'Event log maintenance' -Status 'Complete' -Details 'Event log archive/clear processing complete.'
    return @($results)
}

function Get-StaleScheduledTasks {
    param([int]$Days = 90)
    $cutoff = (Get-Date).AddDays(-$Days)
    $results = @()
    try { $tasks = Get-ScheduledTask -ErrorAction Stop } catch { Add-ActionLog -Step 'Scheduled task audit' -Status 'Failed' -Details $_.Exception.Message; return @() }
    foreach ($task in @($tasks)) {
        try {
            if ($task.TaskPath -like '\Microsoft\*') { continue }
            $info = $null
            try { $info = Get-ScheduledTaskInfo -TaskName $task.TaskName -TaskPath $task.TaskPath -ErrorAction SilentlyContinue } catch { }
            $lastRun = if ($info) { $info.LastRunTime } else { $null }
            $state = [string]$task.State
            $isDisabled = ($state -eq 'Disabled')
            $isStale = ($null -eq $lastRun -or $lastRun -lt $cutoff)
            if ($isDisabled -and $isStale) {
                $results += [PSCustomObject]@{ TaskName=$task.TaskName; TaskPath=$task.TaskPath; State=$state; LastRunTime=$lastRun; LastTaskResult=if($info){$info.LastTaskResult}else{$null}; Reason=('Disabled and not run in {0}+ days or never run' -f $Days) }
            }
        } catch { }
    }
    Add-ActionLog -Step 'Scheduled task audit' -Status 'Complete' -Details ("Found {0} eligible stale tasks." -f @($results).Count)
    return @($results)
}

function Invoke-StaleTaskCleanup {
    param([object[]]$StaleTasks)
    $removed = @()
    if (-not $RemoveStaleTasks) { Add-ActionLog -Step 'Scheduled task cleanup' -Status 'Audit' -Details 'No tasks removed. Use -RemoveStaleTasks to remove eligible tasks.'; return @() }
    if ($script:IsAuditOnly) { Add-ActionLog -Step 'Scheduled task cleanup' -Status 'Audit' -Details 'Tasks would be removed in Remediate mode.'; return @() }
    foreach ($task in @($StaleTasks)) {
        try {
            Unregister-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath -Confirm:$false -ErrorAction Stop
            $removed += $task
        } catch { Add-ActionLog -Step 'Scheduled task removal' -Status 'Failed' -Details ("{0}: {1}" -f $task.TaskName, $_.Exception.Message) }
    }
    Add-ActionLog -Step 'Scheduled task cleanup' -Status 'Complete' -Details ("Removed {0} scheduled tasks." -f @($removed).Count)
    return @($removed)
}

function Get-ResourceSnapshot {
    param([int]$CpuThreshold=85,[int]$MemoryThreshold=85,[double]$DiskQueueThreshold=2.0)
    $cpu = $null; $mem = $null; $queue = $null; $issues = @()
    try { $cpu = [math]::Round((Get-Counter -Counter '\Processor(_Total)\% Processor Time' -SampleInterval 1 -MaxSamples 1 -ErrorAction Stop).CounterSamples.CookedValue,2) } catch { }
    try { $mem = [math]::Round((Get-Counter -Counter '\Memory\% Committed Bytes In Use' -SampleInterval 1 -MaxSamples 1 -ErrorAction Stop).CounterSamples.CookedValue,2) } catch { }
    try { $queue = [math]::Round((Get-Counter -Counter '\PhysicalDisk(_Total)\Current Disk Queue Length' -SampleInterval 1 -MaxSamples 1 -ErrorAction Stop).CounterSamples.CookedValue,2) } catch { }
    if ($null -ne $cpu -and $cpu -gt $CpuThreshold) { $issues += "High CPU usage: $cpu%" }
    if ($null -ne $mem -and $mem -gt $MemoryThreshold) { $issues += "High memory usage: $mem%" }
    if ($null -ne $queue -and $queue -gt $DiskQueueThreshold) { $issues += "High disk queue length: $queue" }
    return [PSCustomObject]@{ CPUPercent=$cpu; MemoryPercent=$mem; DiskQueueLength=$queue; Issues=$issues }
}

function Get-TopCpuProcesses {
    param([int]$Count = 5)
    $rows = @()
    try {
        foreach ($proc in @(Get-Process -ErrorAction SilentlyContinue)) {
            $cpuSeconds = 0
            $workingSet = 0
            try {
                if ($null -ne $proc.CPU) { $cpuSeconds = [double]$proc.CPU }
            } catch {
                try { $cpuSeconds = [double]$proc.TotalProcessorTime.TotalSeconds } catch { $cpuSeconds = 0 }
            }
            try { $workingSet = [int64]$proc.WorkingSet64 } catch { $workingSet = 0 }
            $rows += [PSCustomObject]@{
                Id           = $proc.Id
                ProcessName  = $proc.ProcessName
                CPUSeconds   = [math]::Round($cpuSeconds, 2)
                WorkingSetMB = ConvertTo-MB $workingSet
            }
        }
        return @($rows | Sort-Object -Property CPUSeconds -Descending | Select-Object -First $Count)
    } catch {
        Add-ActionLog -Step 'Top CPU processes' -Status 'Failed' -Details $_.Exception.Message
        return @()
    }
}

function Get-TopMemoryProcesses {
    param([int]$Count = 5)
    $rows = @()
    try {
        foreach ($proc in @(Get-Process -ErrorAction SilentlyContinue)) {
            $cpuSeconds = 0
            $workingSet = 0
            try { if ($null -ne $proc.CPU) { $cpuSeconds = [double]$proc.CPU } } catch { try { $cpuSeconds = [double]$proc.TotalProcessorTime.TotalSeconds } catch { $cpuSeconds = 0 } }
            try { $workingSet = [int64]$proc.WorkingSet64 } catch { $workingSet = 0 }
            $rows += [PSCustomObject]@{
                Id           = $proc.Id
                ProcessName  = $proc.ProcessName
                WorkingSetMB = ConvertTo-MB $workingSet
                CPUSeconds   = [math]::Round($cpuSeconds, 2)
            }
        }
        return @($rows | Sort-Object -Property WorkingSetMB -Descending | Select-Object -First $Count)
    } catch {
        Add-ActionLog -Step 'Top memory processes' -Status 'Failed' -Details $_.Exception.Message
        return @()
    }
}

function Invoke-WindowsUpdateNoReboot {
    $result = [PSCustomObject]@{ Status='Skipped'; Updates=@(); RebootRequired=$false; Details='' }
    if ($SkipWindowsUpdate) { $result.Details='SkipWindowsUpdate was supplied.'; Add-ActionLog -Step 'Windows Update' -Status 'Skipped' -Details $result.Details; return $result }
    if ($script:IsAuditOnly) { $result.Details='Audit mode: update installation skipped.'; Add-ActionLog -Step 'Windows Update' -Status 'Audit' -Details $result.Details; return $result }
    try {
        if (Get-Command Start-WUScan -ErrorAction SilentlyContinue) {
            $updates = @(Start-WUScan -SearchCriteria "IsInstalled=0 AND Type='Software'" -ErrorAction Stop)
            $result.Updates = $updates | Select-Object -First 50
            if ($DownloadUpdatesOnly) { Install-WUUpdates -Updates $updates -DownloadOnly -ErrorAction SilentlyContinue | Out-Null; $result.Status='Downloaded'; $result.Details=('Downloaded {0} update(s).' -f @($updates).Count) }
            elseif (@($updates).Count -gt 0) { Install-WUUpdates -Updates $updates -ErrorAction SilentlyContinue | Out-Null; $result.Status='Installed'; $result.Details=('Installed/requested {0} update(s) without reboot.' -f @($updates).Count) }
            else { $result.Status='No updates'; $result.Details='No applicable software updates found.' }
        } elseif (Get-Command Get-WindowsUpdate -ErrorAction SilentlyContinue) {
            $updates = @(Get-WindowsUpdate -MicrosoftUpdate -ErrorAction SilentlyContinue)
            $result.Updates = $updates | Select-Object -First 50
            if ($DownloadUpdatesOnly) { Get-WindowsUpdate -MicrosoftUpdate -Download -AcceptAll -IgnoreReboot -ErrorAction SilentlyContinue | Out-Null; $result.Status='Downloaded'; $result.Details=('Downloaded {0} update(s) via PSWindowsUpdate.' -f @($updates).Count) }
            elseif (@($updates).Count -gt 0) { Get-WindowsUpdate -MicrosoftUpdate -Install -AcceptAll -IgnoreReboot -ErrorAction SilentlyContinue | Out-Null; $result.Status='Installed'; $result.Details=('Installed/requested {0} update(s) via PSWindowsUpdate without reboot.' -f @($updates).Count) }
            else { $result.Status='No updates'; $result.Details='No updates found via PSWindowsUpdate.' }
        } else {
            $result.Status='Unavailable'; $result.Details='No supported Windows Update PowerShell provider found.'
        }
    } catch { $result.Status='Failed'; $result.Details=$_.Exception.Message }
    Add-ActionLog -Step 'Windows Update' -Status $result.Status -Details $result.Details
    return $result
}

function Set-RestorePointFrequencyForRun {
    $path = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore'
    $name = 'SystemRestorePointCreationFrequency'
    $backup = [PSCustomObject]@{ Path=$path; Name=$name; Existed=$false; Value=$null; Status='Not changed' }
    try {
        if (-not (Test-Path -LiteralPath $path)) { New-Item -Path $path -Force | Out-Null }
        $prop = Get-ItemProperty -LiteralPath $path -Name $name -ErrorAction SilentlyContinue
        if ($prop -and ($prop.PSObject.Properties.Name -contains $name)) { $backup.Existed=$true; $backup.Value=$prop.$name }
        Set-ItemProperty -LiteralPath $path -Name $name -Value 0 -Type DWord -Force
        $backup.Status='Set to 0 for this run'
    } catch { $backup.Status = 'Failed: ' + $_.Exception.Message }
    return $backup
}

function Restore-RestorePointFrequencySetting {
    param([object]$Backup)
    if ($null -eq $Backup) { return }
    try {
        if ($Backup.Existed) { Set-ItemProperty -LiteralPath $Backup.Path -Name $Backup.Name -Value $Backup.Value -Type DWord -Force }
        else { Remove-ItemProperty -LiteralPath $Backup.Path -Name $Backup.Name -ErrorAction SilentlyContinue }
        Add-ActionLog -Step 'Restore point frequency reset' -Status 'Complete' -Details 'SystemRestorePointCreationFrequency setting restored.'
    } catch { Add-ActionLog -Step 'Restore point frequency reset' -Status 'Failed' -Details $_.Exception.Message }
}

function New-MaintenanceRestorePoint {
    param([ValidateSet('BEFORE','AFTER')][string]$Stage)
    $name = 'PC_Cleanup{0}-{1}' -f (Get-Date -Format 'ddMMyy'), $Stage
    $result = [PSCustomObject]@{ Stage=$Stage; Name=$name; Status='Skipped'; Time=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss'); Details='' }
    if ($SkipRestorePoints) { $result.Details='SkipRestorePoints was supplied.'; return $result }
    if ($script:IsAuditOnly) { $result.Status='Audit'; $result.Details='Restore point would be created in Remediate mode.'; return $result }
    try { Checkpoint-Computer -Description $name -RestorePointType 'MODIFY_SETTINGS' -ErrorAction Stop; $result.Status='Complete'; $result.Details='Restore point created.' }
    catch { $result.Status='Failed'; $result.Details=$_.Exception.Message }
    Add-ActionLog -Step ("Restore point $Stage") -Status $result.Status -Details $result.Details
    return $result
}

function Invoke-ExternalMaintenanceCommand {
    param([string]$Name,[string]$FilePath,[string[]]$Arguments,[int]$TimeoutMinutes = 0)
    $id = [guid]::NewGuid().ToString('N')
    $stdout = Join-Path $env:TEMP ("PCMaint_{0}_{1}.out" -f $Name, $id)
    $stderr = Join-Path $env:TEMP ("PCMaint_{0}_{1}.err" -f $Name, $id)
    $result = [PSCustomObject]@{ Name=$Name; Command=($FilePath + ' ' + ($Arguments -join ' ')); Status='Not run'; ExitCode=$null; Output=''; Errors=''; Summary='' }
    try {
        if ($script:IsAuditOnly -and $Name -notmatch 'CHKDSK_SCAN') { $result.Status='Audit'; $result.Summary='Command would run in Remediate mode.'; return $result }
        $process = Start-Process -FilePath $FilePath -ArgumentList $Arguments -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr -ErrorAction Stop
        $start = Get-Date
        $lastNotice = $start.AddMinutes(-10)
        while (-not $process.HasExited) {
            Start-Sleep -Seconds 15
            $elapsed = New-TimeSpan -Start $start -End (Get-Date)
            if (((Get-Date) - $lastNotice).TotalMinutes -ge 5) {
                $logInfo = ''
                if (Test-Path -LiteralPath 'C:\Windows\Logs\CBS\CBS.log') {
                    try { $item = Get-Item -LiteralPath 'C:\Windows\Logs\CBS\CBS.log'; $logInfo = (' CBS.log last write: {0}' -f $item.LastWriteTime) } catch { }
                }
                Write-Host ("[{0}] {1} still running. Elapsed: {2:N1} minutes.{3}" -f (Get-Date -Format 'HH:mm:ss'), $Name, $elapsed.TotalMinutes, $logInfo) -ForegroundColor DarkYellow
                $lastNotice = Get-Date
            }
            if ($TimeoutMinutes -gt 0 -and $elapsed.TotalMinutes -ge $TimeoutMinutes) {
                try { $process.Kill() } catch { }
                $result.Status = 'Timed out'
                $result.Summary = ('Timed out after {0} minute(s).' -f $TimeoutMinutes)
                break
            }
        }
        if ($process.HasExited) { $result.ExitCode = $process.ExitCode }
        if (Test-Path -LiteralPath $stdout) { $result.Output = Get-Content -LiteralPath $stdout -Raw -ErrorAction SilentlyContinue }
        if (Test-Path -LiteralPath $stderr) { $result.Errors = Get-Content -LiteralPath $stderr -Raw -ErrorAction SilentlyContinue }
        if ($result.Status -ne 'Timed out') {
            if ($process.ExitCode -eq 0) { $result.Status='Complete' } else { $result.Status=('Completed with exit code ' + $process.ExitCode) }
        }
    } catch { $result.Status='Failed'; $result.Errors=$_.Exception.Message; $result.Summary=$_.Exception.Message }
    finally { Remove-Item -LiteralPath $stdout, $stderr -Force -ErrorAction SilentlyContinue }
    return $result
}

function Invoke-ChkdskSmart {
    param([string]$Letter='C')
    $cleanLetter = $Letter.TrimEnd(':')
    $drive = '{0}:' -f $cleanLetter
    $result = [PSCustomObject]@{ Name='CHKDSK smart'; ScanCommand="chkdsk $drive /scan"; QueueCommand="echo Y|chkdsk $drive /f"; Status='Skipped'; ScanExitCode=$null; QueueExitCode=$null; ScanOutput=''; QueueOutput=''; Errors=''; Summary='' }
    if ($SkipChkdsk) { $result.Summary='SkipChkdsk was supplied.'; Add-ActionLog -Step 'CHKDSK' -Status 'Skipped' -Details $result.Summary; return $result }
    try {
        $scanOutput = cmd.exe /c "chkdsk $drive /scan" 2>&1 | Out-String
        $result.ScanExitCode = $LASTEXITCODE
        $result.ScanOutput = $scanOutput
        $needsFix = $false
        if ($AlwaysQueueChkdskFix) { $needsFix = $true }
        if ($scanOutput -match 'found problems|errors found|further action is required|run chkdsk /f|Windows has scanned the file system and found problems') { $needsFix = $true }
        if ($needsFix) {
            if ($script:IsAuditOnly) { $result.Status='Audit'; $result.Summary='CHKDSK /F would be queued in Remediate mode.' }
            else {
                $queueOutput = cmd.exe /c "echo Y|chkdsk $drive /f" 2>&1 | Out-String
                $result.QueueExitCode = $LASTEXITCODE
                $result.QueueOutput = $queueOutput
                if ($queueOutput -match 'scheduled|restart|reboot|next time the system restarts') { $result.Status='Queued'; $result.Summary='CHKDSK /F was queued for the next reboot.' }
                else { $result.Status='Check output'; $result.Summary='CHKDSK /F was requested. Review output to confirm scheduling.' }
            }
        } else { $result.Status='Scan complete'; $result.Summary='CHKDSK /scan did not clearly indicate offline repair was required.' }
    } catch { $result.Status='Failed'; $result.Errors=$_.Exception.Message; $result.Summary=$_.Exception.Message }
    Add-ActionLog -Step 'CHKDSK' -Status $result.Status -Details $result.Summary
    return $result
}

function Invoke-SfcScan {
    $result = [PSCustomObject]@{ Name='SFC scan'; Command='sfc.exe /scannow'; Status='Skipped'; ExitCode=$null; Output=''; Errors=''; Summary='' }
    if ($SkipSfc) { $result.Summary='SkipSfc was supplied.'; Add-ActionLog -Step 'SFC scan' -Status 'Skipped' -Details $result.Summary; return $result }
    $sfc = Join-Path $env:windir 'System32\sfc.exe'; if (-not (Test-Path -LiteralPath $sfc)) { $sfc='sfc.exe' }
    $result = Invoke-ExternalMaintenanceCommand -Name 'SFC' -FilePath $sfc -Arguments @('/scannow') -TimeoutMinutes $SfcTimeoutMinutes
    $result.Name='SFC scan'; $result.Command='sfc.exe /scannow'
    $combined = ($result.Output + "`r`n" + $result.Errors)
    if ($combined -match 'did not find any integrity violations') { $result.Summary='Windows Resource Protection did not find integrity violations.' }
    elseif ($combined -match 'found corrupt files and successfully repaired') { $result.Summary='SFC found corrupt files and repaired them.' }
    elseif ($combined -match 'found corrupt files but was unable to fix') { $result.Summary='SFC found corrupt files but could not repair all of them.' }
    elseif ($result.Status -eq 'Audit') { $result.Summary='SFC would run in Remediate mode.' }
    else { $result.Summary='SFC completed. Review full output.' }
    Add-ActionLog -Step 'SFC scan' -Status $result.Status -Details $result.Summary
    return $result
}

function Invoke-DismRestoreHealth {
    $result = [PSCustomObject]@{ Name='DISM RestoreHealth'; Command='dism.exe /online /cleanup-image /restorehealth /norestart'; Status='Skipped'; ExitCode=$null; Output=''; Errors=''; Summary='' }
    if ($SkipDismRestoreHealth) { $result.Summary='SkipDismRestoreHealth was supplied.'; Add-ActionLog -Step 'DISM RestoreHealth' -Status 'Skipped' -Details $result.Summary; return $result }
    $dism = Join-Path $env:windir 'System32\dism.exe'; if (-not (Test-Path -LiteralPath $dism)) { $dism='dism.exe' }
    $result = Invoke-ExternalMaintenanceCommand -Name 'DISM_RestoreHealth' -FilePath $dism -Arguments @('/online','/cleanup-image','/restorehealth','/norestart') -TimeoutMinutes $DismTimeoutMinutes
    $result.Name='DISM RestoreHealth'; $result.Command='dism.exe /online /cleanup-image /restorehealth /norestart'
    $combined = ($result.Output + "`r`n" + $result.Errors)
    if ($combined -match 'restore operation completed successfully|operation completed successfully') { $result.Summary='DISM RestoreHealth completed successfully.' }
    elseif ($combined -match 'Error:\s*([0-9A-Fa-fx]+)') { $result.Summary='DISM reported an error. Review full output.' }
    elseif ($result.Status -eq 'Audit') { $result.Summary='DISM RestoreHealth would run in Remediate mode.' }
    else { $result.Summary='DISM completed. Review full output.' }
    Add-ActionLog -Step 'DISM RestoreHealth' -Status $result.Status -Details $result.Summary
    return $result
}

function Get-LocalUserProfileSizes {
    if ($SkipUserProfileSizes) { Add-ActionLog -Step 'Local user profile size audit' -Status 'Skipped' -Details 'SkipUserProfileSizes was supplied.'; return @() }
    $profiles = @()
    try { $profiles = Get-CimInstance Win32_UserProfile -ErrorAction Stop | Where-Object { $_.LocalPath -like "$env:SystemDrive\Users\*" } }
    catch { $profiles = @() }
    $results = @()
    foreach ($profile in @($profiles)) {
        $sizeInfo = Get-FolderSizeInfo -Path $profile.LocalPath
        $lastUse = $null
        try {
            if ($profile.LastUseTime) {
                if ($profile.LastUseTime -is [datetime]) { $lastUse = $profile.LastUseTime }
                else { $lastUse = [Management.ManagementDateTimeConverter]::ToDateTime($profile.LastUseTime) }
            }
        } catch { }
        $daysSince = $null
        if ($lastUse) { $daysSince = [math]::Round(((Get-Date) - $lastUse).TotalDays, 0) }
        $stale = ($daysSince -ne $null -and $daysSince -ge $ProfileStaleDays -and -not $profile.Special -and -not $profile.Loaded)
        $results += [PSCustomObject]@{ ProfileName=(Split-Path $profile.LocalPath -Leaf); Path=$profile.LocalPath; SizeGB=$sizeInfo.SizeGB; SizeMB=$sizeInfo.SizeMB; FileCount=$sizeInfo.FileCount; Loaded=$profile.Loaded; Special=$profile.Special; LastUseTime=$lastUse; DaysSinceUse=$daysSince; StaleProfileCandidate=$stale; SID=$profile.SID; ScanErrors=$sizeInfo.Errors }
    }
    Add-ActionLog -Step 'Local user profile size audit' -Status 'Complete' -Details ("Measured {0} profiles." -f @($results).Count)
    return @($results | Sort-Object SizeGB -Descending)
}

function Get-PendingRebootStatus {
    $reasons = @()
    $checks = @(
        @{Path='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'; Reason='CBS servicing reboot pending'},
        @{Path='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'; Reason='Windows Update reboot pending'},
        @{Path='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending'; Reason='CBS packages pending'}
    )
    foreach ($check in $checks) { if (Test-Path -LiteralPath $check.Path) { $reasons += $check.Reason } }
    try {
        $sessionManager = Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
        if ($sessionManager -and $sessionManager.PendingFileRenameOperations) { $reasons += 'Pending file rename operations' }
    } catch { }
    try {
        $active = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName' -ErrorAction SilentlyContinue).ComputerName
        $pending = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName' -ErrorAction SilentlyContinue).ComputerName
        if ($active -and $pending -and $active -ne $pending) { $reasons += 'Computer rename pending' }
    } catch { }
    return [PSCustomObject]@{ RebootRequired=(@($reasons).Count -gt 0); Reasons=($reasons -join '; '); ReasonCount=@($reasons).Count }
}

function Get-StorageHealth {
    $results = @()
    try {
        if (Get-Command Get-PhysicalDisk -ErrorAction SilentlyContinue) {
            foreach ($pd in @(Get-PhysicalDisk -ErrorAction SilentlyContinue)) {
                $results += [PSCustomObject]@{
                    Source='PhysicalDisk'; FriendlyName=$pd.FriendlyName; SerialNumber=$pd.SerialNumber; MediaType=$pd.MediaType; BusType=''; HealthStatus=$pd.HealthStatus; OperationalStatus=($pd.OperationalStatus -join ','); PartitionStyle=''; SizeGB=(ConvertTo-GB $pd.Size); WearPercent=''; TemperatureC=''; ReadErrors=''; WriteErrors=''; PowerOnHours=''; UnsafeShutdowns=''; Details=''
                }
            }
        }
    } catch { }
    try {
        if (Get-Command Get-Disk -ErrorAction SilentlyContinue) {
            foreach ($d in @(Get-Disk -ErrorAction SilentlyContinue)) {
                $results += [PSCustomObject]@{
                    Source='Disk'; FriendlyName=$d.FriendlyName; SerialNumber=$d.SerialNumber; MediaType=''; BusType=$d.BusType; HealthStatus=$d.HealthStatus; OperationalStatus=($d.OperationalStatus -join ','); PartitionStyle=$d.PartitionStyle; SizeGB=(ConvertTo-GB $d.Size); WearPercent=''; TemperatureC=''; ReadErrors=''; WriteErrors=''; PowerOnHours=''; UnsafeShutdowns=''; Details=''
                }
            }
        }
    } catch { }
    try {
        if ((Get-Command Get-StorageReliabilityCounter -ErrorAction SilentlyContinue) -and (Get-Command Get-PhysicalDisk -ErrorAction SilentlyContinue)) {
            foreach ($pd in @(Get-PhysicalDisk -ErrorAction SilentlyContinue)) {
                $rel = $pd | Get-StorageReliabilityCounter -ErrorAction SilentlyContinue
                if ($rel) {
                    $results += [PSCustomObject]@{
                        Source='ReliabilityCounter'; FriendlyName=$pd.FriendlyName; SerialNumber=$pd.SerialNumber; MediaType=$pd.MediaType; BusType=''; HealthStatus=$pd.HealthStatus; OperationalStatus=($pd.OperationalStatus -join ','); PartitionStyle=''; SizeGB=(ConvertTo-GB $pd.Size); WearPercent=(Get-PropertyValueSafe $rel 'Wear' ''); TemperatureC=(Get-PropertyValueSafe $rel 'Temperature' ''); ReadErrors=(Get-PropertyValueSafe $rel 'ReadErrorsTotal' ''); WriteErrors=(Get-PropertyValueSafe $rel 'WriteErrorsTotal' ''); PowerOnHours=(Get-PropertyValueSafe $rel 'PowerOnHours' ''); UnsafeShutdowns=(Get-PropertyValueSafe $rel 'UnsafeShutdowns' ''); Details='StorageReliabilityCounter'
                    }
                }
            }
        }
    } catch { }
    try {
        foreach ($dd in @(Get-CimInstance -Namespace 'root\wmi' -ClassName MSStorageDriver_FailurePredictStatus -ErrorAction SilentlyContinue)) {
            $results += [PSCustomObject]@{ Source='SMARTFailurePredict'; FriendlyName=$dd.InstanceName; SerialNumber=''; MediaType=''; BusType=''; HealthStatus=if($dd.PredictFailure){'PredictFailure'}else{'OK'}; OperationalStatus=''; PartitionStyle=''; SizeGB=''; WearPercent=''; TemperatureC=''; ReadErrors=''; WriteErrors=''; PowerOnHours=''; UnsafeShutdowns=''; Details=('Reason: {0}' -f $dd.Reason) }
        }
    } catch { }
    if (-not $results -or @($results).Count -eq 0) { Add-ActionLog -Step 'Storage health' -Status 'Unavailable' -Details 'No storage health cmdlets returned data.' }
    else { Add-ActionLog -Step 'Storage health' -Status 'Complete' -Details ("Collected {0} disk/storage records." -f @($results).Count) }
    return @($results)
}

function Get-BatteryHealth {
    param([string]$OutputFolder)
    $result = [PSCustomObject]@{ HasBattery=$false; Name=''; Status=''; EstimatedChargeRemaining=$null; EstimatedRunTime=$null; BatteryReportPath=''; Details='' }
    if ($SkipBatteryReport) { $result.Details='SkipBatteryReport was supplied.'; return $result }
    try {
        $battery = Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($battery) {
            $result.HasBattery = $true
            $result.Name = $battery.Name
            $result.Status = $battery.Status
            $result.EstimatedChargeRemaining = $battery.EstimatedChargeRemaining
            $result.EstimatedRunTime = $battery.EstimatedRunTime
            $path = Join-Path $OutputFolder ("BatteryReport_{0}_{1}.html" -f $env:COMPUTERNAME, (Get-Date -Format 'yyyyMMdd_HHmmss'))
            if (-not $script:IsAuditOnly) { powercfg /batteryreport /output $path | Out-Null; if (Test-Path -LiteralPath $path) { $result.BatteryReportPath = $path } }
            else { $result.Details='Audit mode: battery report generation skipped.' }
        } else { $result.Details='No battery detected.' }
    } catch { $result.Details=$_.Exception.Message }
    Add-ActionLog -Step 'Battery health' -Status 'Complete' -Details $result.Details
    return $result
}

function Get-ProblemDevices {
    try {
        if (Get-Command Get-PnpDevice -ErrorAction SilentlyContinue) {
            $devices = @(Get-PnpDevice -ErrorAction SilentlyContinue | Where-Object { $_.Status -ne 'OK' } | Select-Object -First 100 Class, FriendlyName, Status, Problem, InstanceId)
            Add-ActionLog -Step 'Device problem check' -Status 'Complete' -Details ("Found {0} non-OK devices." -f @($devices).Count)
            return $devices
        }
    } catch { }
    Add-ActionLog -Step 'Device problem check' -Status 'Unavailable' -Details 'Get-PnpDevice unavailable or failed.'
    return @()
}

function Get-StartupItems {
    $items = @()
    $regPaths = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run',
        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run',
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
    )
    foreach ($path in $regPaths) {
        try {
            $props = Get-ItemProperty -LiteralPath $path -ErrorAction SilentlyContinue
            if ($props) {
                foreach ($p in $props.PSObject.Properties) {
                    if ($p.Name -notmatch '^PS') { $items += [PSCustomObject]@{ Source=$path; Name=$p.Name; Command=[string]$p.Value } }
                }
            }
        } catch { }
    }
    $startupFolders = @(
        [Environment]::GetFolderPath('Startup'),
        (Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs\Startup')
    )
    foreach ($folder in $startupFolders) {
        if (Test-Path -LiteralPath $folder) {
            Get-ChildItem -LiteralPath $folder -ErrorAction SilentlyContinue | ForEach-Object { $items += [PSCustomObject]@{ Source=$folder; Name=$_.Name; Command=$_.FullName } }
        }
    }
    Add-ActionLog -Step 'Startup item audit' -Status 'Complete' -Details ("Found {0} startup items." -f @($items).Count)
    return @($items)
}

function Get-ServiceHealthIssues {
    try {
        $services = @(Get-CimInstance Win32_Service -ErrorAction SilentlyContinue | Where-Object { $_.StartMode -eq 'Auto' -and $_.State -ne 'Running' } | Select-Object -First 100 Name, DisplayName, State, StartMode, StartName, ExitCode)
        Add-ActionLog -Step 'Service health check' -Status 'Complete' -Details ("Found {0} auto services not running." -f @($services).Count)
        return $services
    } catch { Add-ActionLog -Step 'Service health check' -Status 'Failed' -Details $_.Exception.Message; return @() }
}

function Get-DefenderStatus {
    $result = [PSCustomObject]@{ Available=$false; AMServiceEnabled=$null; RealTimeProtectionEnabled=$null; AntivirusSignatureLastUpdated=$null; QuickScanAgeDays=$null; FullScanAgeDays=$null; QuickScanRequested=$false; Details='' }
    try {
        if (Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue) {
            $mp = Get-MpComputerStatus -ErrorAction Stop
            $result.Available = $true
            $result.AMServiceEnabled = $mp.AMServiceEnabled
            $result.RealTimeProtectionEnabled = $mp.RealTimeProtectionEnabled
            $result.AntivirusSignatureLastUpdated = $mp.AntivirusSignatureLastUpdated
            if ($mp.QuickScanEndTime) { $result.QuickScanAgeDays = [math]::Round(((Get-Date) - $mp.QuickScanEndTime).TotalDays, 1) }
            if ($mp.FullScanEndTime) { $result.FullScanAgeDays = [math]::Round(((Get-Date) - $mp.FullScanEndTime).TotalDays, 1) }
            if ($RunDefenderQuickScan) {
                $result.QuickScanRequested = $true
                if ($script:IsAuditOnly) { $result.Details='Audit mode: quick scan would run in Remediate mode.' }
                elseif (Get-Command Start-MpScan -ErrorAction SilentlyContinue) { Start-MpScan -ScanType QuickScan -ErrorAction SilentlyContinue; $result.Details='Quick scan requested.' }
            }
        } else { $result.Details='Defender PowerShell cmdlets unavailable.' }
    } catch { $result.Details=$_.Exception.Message }
    Add-ActionLog -Step 'Defender status' -Status 'Complete' -Details $result.Details
    return $result
}

function Get-CrashDumpSummary {
    $items = @()
    $paths = @((Join-Path $env:windir 'Minidump'), (Join-Path $env:windir 'MEMORY.DMP'))
    foreach ($path in $paths) {
        try {
            if (Test-Path -LiteralPath $path) {
                $item = Get-Item -LiteralPath $path -ErrorAction SilentlyContinue
                if ($item -and $item.PSIsContainer) {
                    foreach ($dump in @(Get-ChildItem -LiteralPath $path -File -ErrorAction SilentlyContinue)) { $items += $dump }
                } elseif ($item) { $items += $item }
            }
        } catch { }
    }
    $bytes = [int64]0
    foreach ($item in @($items)) {
        try { if ($null -ne $item.Length) { $bytes += [int64]$item.Length } } catch { }
    }
    $latest = @($items | Sort-Object LastWriteTime -Descending | Select-Object -First 1)
    $latestItem = if ((Get-ObjectCountSafe $latest) -gt 0) { $latest[0] } else { $null }
    return [PSCustomObject]@{
        DumpCount      = Get-ObjectCountSafe $items
        TotalSizeGB    = ConvertTo-GB $bytes
        LatestDump     = if ($latestItem) { $latestItem.FullName } else { '' }
        LatestDumpTime = if ($latestItem) { $latestItem.LastWriteTime } else { $null }
    }
}

function Get-StorageHotspots {
    if ($SkipStorageHotspots) { Add-ActionLog -Step 'Storage hotspot scan' -Status 'Skipped' -Details 'SkipStorageHotspots was supplied.'; return @() }
    $folders = @()
    $profileRoot = Join-Path $env:SystemDrive 'Users'
    if (Test-Path -LiteralPath $profileRoot) {
        Get-ChildItem -LiteralPath $profileRoot -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -notin @('Default','Default User','Public','All Users') } | ForEach-Object {
            foreach ($child in @('Downloads','Desktop','Documents','Videos')) { $folders += (Join-Path $_.FullName $child) }
        }
    }
    $files = @()
    foreach ($folder in @($folders | Sort-Object -Unique)) {
        if (Test-Path -LiteralPath $folder) {
            try { $files += Get-ChildItem -LiteralPath $folder -Recurse -Force -File -ErrorAction SilentlyContinue | Where-Object { $_.Length -gt 100MB } } catch { }
        }
    }
    $top = @($files | Sort-Object Length -Descending | Select-Object -First $TopLargeFileCount @{Name='SizeGB';Expression={ConvertTo-GB $_.Length}}, FullName, LastWriteTime, Extension)
    Add-ActionLog -Step 'Storage hotspot scan' -Status 'Complete' -Details ("Found {0} large user files." -f @($top).Count)
    return $top
}

function Get-WindowsOldInfo {
    $path = Join-Path $env:SystemDrive 'Windows.old'
    if (Test-Path -LiteralPath $path) {
        $size = Get-FolderSizeInfo -Path $path
        return [PSCustomObject]@{ Exists=$true; Path=$path; SizeGB=$size.SizeGB; Details='Windows.old found. Use -Mode Remediate -CleanWindowsOld to remove rollback files.' }
    }
    return [PSCustomObject]@{ Exists=$false; Path=$path; SizeGB=0; Details='Windows.old not found.' }
}


function Get-BrowserCachePaths {
    $paths = @()
    $profileRoot = Join-Path $env:SystemDrive 'Users'
    if (Test-Path -LiteralPath $profileRoot) {
        Get-ChildItem -LiteralPath $profileRoot -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -notin @('Default','Default User','Public','All Users') } | ForEach-Object {
            $paths += (Join-Path $_.FullName 'AppData\Local\Microsoft\Edge\User Data\Default\Cache')
            $paths += (Join-Path $_.FullName 'AppData\Local\Microsoft\Edge\User Data\Default\Cache\Cache_Data')
            $paths += (Join-Path $_.FullName 'AppData\Local\Microsoft\Edge\User Data\Default\Code Cache')
            $paths += (Join-Path $_.FullName 'AppData\Local\Google\Chrome\User Data\Default\Cache')
            $paths += (Join-Path $_.FullName 'AppData\Local\Google\Chrome\User Data\Default\Cache\Cache_Data')
            $paths += (Join-Path $_.FullName 'AppData\Local\Google\Chrome\User Data\Default\Code Cache')
            $paths += (Join-Path $_.FullName 'AppData\Local\Mozilla\Firefox\Profiles')
        }
    }
    return @($paths | Sort-Object -Unique)
}

function Get-MultiPathSizeInfo {
    param([string]$Category,[string[]]$Paths)
    $totalBytes = [int64]0
    $files = 0
    $existing = 0
    $errors = 0
    foreach ($path in @($Paths | Where-Object { $_ } | Sort-Object -Unique)) {
        try {
            $info = Get-FolderSizeInfo -Path $path
            if ($info.Exists) { $existing++ }
            $totalBytes += [int64]$info.SizeBytes
            $files += [int]$info.FileCount
            $errors += [int]$info.Errors
        } catch { $errors++ }
    }
    return [PSCustomObject]@{ Category=$Category; ExistingPathCount=$existing; SizeGB=(ConvertTo-GB $totalBytes); SizeMB=(ConvertTo-MB $totalBytes); FileCount=$files; Errors=$errors }
}

function Get-CleanupCategorySnapshot {
    param([string]$Stage)
    $windowsOld = Join-Path $env:SystemDrive 'Windows.old'
    $deliveryPath = Join-Path $env:windir 'ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache'
    $recyclePath = Join-Path $env:SystemDrive '$Recycle.Bin'
    $componentStore = Join-Path $env:windir 'WinSxS'
    $crashPaths = @((Join-Path $env:windir 'Minidump'), (Join-Path $env:windir 'MEMORY.DMP'))
    $data = @()
    $data += Get-MultiPathSizeInfo -Category 'Temp and general cache' -Paths (Get-CleanupPaths)
    $data += Get-MultiPathSizeInfo -Category 'Browser cache' -Paths (Get-BrowserCachePaths)
    $data += Get-MultiPathSizeInfo -Category 'Delivery Optimization cache' -Paths @($deliveryPath)
    $data += Get-MultiPathSizeInfo -Category 'Recycle Bin' -Paths @($recyclePath)
    $data += Get-MultiPathSizeInfo -Category 'Crash dumps' -Paths $crashPaths
    $data += Get-MultiPathSizeInfo -Category 'Windows.old' -Paths @($windowsOld)
    $data += Get-MultiPathSizeInfo -Category 'Component store (WinSxS apparent size)' -Paths @($componentStore)
    foreach ($row in $data) { $row | Add-Member -NotePropertyName Stage -NotePropertyValue $Stage -Force }
    return @($data)
}

function Compare-CleanupCategorySnapshots {
    param([object[]]$Before,[object[]]$After)
    $results = @()
    foreach ($b in @($Before)) {
        $a = @($After | Where-Object { $_.Category -eq $b.Category } | Select-Object -First 1)
        $afterSize = if ($a -and $a.Count -gt 0) { [double]$a[0].SizeGB } else { 0 }
        $afterFiles = if ($a -and $a.Count -gt 0) { [int]$a[0].FileCount } else { 0 }
        $results += [PSCustomObject]@{
            Category = $b.Category
            BeforeGB = [double]$b.SizeGB
            AfterGB = $afterSize
            ChangeGB = [math]::Round(([double]$b.SizeGB - $afterSize), 2)
            BeforeFiles = [int]$b.FileCount
            AfterFiles = $afterFiles
            FileChange = ([int]$b.FileCount - $afterFiles)
        }
    }
    return @($results)
}

function Invoke-CrashDumpCleanup {
    $result = [PSCustomObject]@{ Requested=$false; RemovedCount=0; RemovedGB=0; Status='Skipped'; Details='' }
    if (-not $CleanCrashDumps) { $result.Details='Use -CleanCrashDumps to remove crash dumps.'; Add-ActionLog -Step 'Crash dump cleanup' -Status 'Skipped' -Details $result.Details; return $result }
    $result.Requested = $true
    if ($script:IsAuditOnly) { $result.Status='Audit'; $result.Details='Crash dumps would be removed in Remediate mode.'; Add-ActionLog -Step 'Crash dump cleanup' -Status 'Audit' -Details $result.Details; return $result }
    $paths = @((Join-Path $env:windir 'Minidump'), (Join-Path $env:windir 'MEMORY.DMP'))
    $bytes = [int64]0
    foreach ($path in $paths) {
        try {
            if (Test-Path -LiteralPath $path) {
                $item = Get-Item -LiteralPath $path -ErrorAction SilentlyContinue
                if ($item -and $item.PSIsContainer) {
                    foreach ($dump in @(Get-ChildItem -LiteralPath $path -File -ErrorAction SilentlyContinue)) {
                        $len = [int64]$dump.Length
                        Remove-Item -LiteralPath $dump.FullName -Force -ErrorAction Stop
                        $bytes += $len; $result.RemovedCount++
                    }
                } elseif ($item) {
                    $len = [int64]$item.Length
                    Remove-Item -LiteralPath $item.FullName -Force -ErrorAction Stop
                    $bytes += $len; $result.RemovedCount++
                }
            }
        } catch { $result.Details += ($_.Exception.Message + '; ') }
    }
    $result.RemovedGB = ConvertTo-GB $bytes
    $result.Status = 'Complete'
    Add-ActionLog -Step 'Crash dump cleanup' -Status 'Complete' -Details ("Removed {0} dump file(s), {1} GB." -f $result.RemovedCount, $result.RemovedGB)
    return $result
}

function Invoke-WindowsOldCleanup {
    $path = Join-Path $env:SystemDrive 'Windows.old'
    $result = [PSCustomObject]@{ Requested=$false; Path=$path; BeforeGB=0; AfterGB=0; RemovedGB=0; Status='Skipped'; Details='' }
    if (-not $CleanWindowsOld) { $result.Details='Use -CleanWindowsOld to remove Windows.old.'; Add-ActionLog -Step 'Windows.old cleanup' -Status 'Skipped' -Details $result.Details; return $result }
    $result.Requested = $true
    if (-not (Test-Path -LiteralPath $path)) { $result.Status='Not found'; $result.Details='Windows.old not present.'; Add-ActionLog -Step 'Windows.old cleanup' -Status 'Not found' -Details $result.Details; return $result }
    $before = Get-FolderSizeInfo -Path $path
    $result.BeforeGB = $before.SizeGB
    if ($script:IsAuditOnly) { $result.Status='Audit'; $result.Details='Windows.old would be removed in Remediate mode. This removes rollback files.'; Add-ActionLog -Step 'Windows.old cleanup' -Status 'Audit' -Details $result.Details; return $result }
    try {
        $takeown = Start-Process -FilePath 'takeown.exe' -ArgumentList @('/F', $path, '/R', '/D', 'Y') -Wait -PassThru -WindowStyle Hidden -ErrorAction SilentlyContinue
        $icacls = Start-Process -FilePath 'icacls.exe' -ArgumentList @($path, '/grant', 'Administrators:F', '/T', '/C') -Wait -PassThru -WindowStyle Hidden -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction Stop
        Start-Sleep -Seconds 2
        $after = Get-FolderSizeInfo -Path $path
        $result.AfterGB = $after.SizeGB
        $result.RemovedGB = [math]::Round(([double]$result.BeforeGB - [double]$result.AfterGB), 2)
        $result.Status='Complete'
        $result.Details=("Rollback files removed. takeown exit: {0}; icacls exit: {1}" -f $takeown.ExitCode, $icacls.ExitCode)
        Add-ActionLog -Step 'Windows.old cleanup' -Status 'Complete' -Details $result.Details
    } catch {
        $result.Status='Failed'
        $result.Details=$_.Exception.Message
        Add-ActionLog -Step 'Windows.old cleanup' -Status 'Failed' -Details $result.Details
    }
    return $result
}

function Invoke-StaleProfileRemoval {
    param([object[]]$Profiles)
    $removed = @()
    if (-not $RemoveStaleProfiles) { Add-ActionLog -Step 'Stale profile removal' -Status 'Skipped' -Details 'Use -RemoveStaleProfiles -ForceProfileRemoval to delete stale profiles.'; return @() }
    if (-not $ForceProfileRemoval) { Add-ActionLog -Step 'Stale profile removal' -Status 'Skipped' -Details 'ForceProfileRemoval was not supplied.'; return @() }
    if ($script:IsAuditOnly) { Add-ActionLog -Step 'Stale profile removal' -Status 'Audit' -Details 'Profiles would be removed in Remediate mode.'; return @() }
    try {
        $cimProfiles = @(Get-CimInstance Win32_UserProfile -ErrorAction SilentlyContinue | Where-Object { $_.LocalPath -like "$env:SystemDrive\Users\*" -and -not $_.Special -and -not $_.Loaded })
        foreach ($profile in $cimProfiles) {
            $lastUse = $null
            try { if ($profile.LastUseTime) { $lastUse = [Management.ManagementDateTimeConverter]::ToDateTime($profile.LastUseTime) } } catch { }
            $daysSince = if ($lastUse) { [math]::Round(((Get-Date) - $lastUse).TotalDays, 0) } else { 9999 }
            if ($daysSince -ge $RemoveProfilesOlderThanDays) {
                try {
                    $name = Split-Path $profile.LocalPath -Leaf
                    $size = Get-FolderSizeInfo -Path $profile.LocalPath
                    $profile | Remove-CimInstance -ErrorAction Stop
                    $removed += [PSCustomObject]@{ ProfileName=$name; Path=$profile.LocalPath; SID=$profile.SID; DaysSinceUse=$daysSince; RemovedGB=$size.SizeGB; Status='Removed'; Details='' }
                } catch {
                    $removed += [PSCustomObject]@{ ProfileName=(Split-Path $profile.LocalPath -Leaf); Path=$profile.LocalPath; SID=$profile.SID; DaysSinceUse=$daysSince; RemovedGB=0; Status='Failed'; Details=$_.Exception.Message }
                }
            }
        }
    } catch { Add-ActionLog -Step 'Stale profile removal' -Status 'Failed' -Details $_.Exception.Message }
    Add-ActionLog -Step 'Stale profile removal' -Status 'Complete' -Details ("Removed/attempted {0} stale profile(s)." -f @($removed).Count)
    return @($removed)
}

function Get-EventInterpretation {
    param([object[]]$Events)
    $rules = @(
        @{ Provider='Microsoft-Windows-Kernel-Power'; Id=41; Category='Unexpected shutdown'; Severity='High'; Meaning='The system rebooted without cleanly shutting down first.'; Action='Check power, battery, overheating, crashes, or forced restarts.' },
        @{ Provider='Disk'; Id='7,51,153,157'; Category='Disk or controller issue'; Severity='High'; Meaning='Windows logged disk/controller read/write/reset problems.'; Action='Check SMART/storage health and backup status.' },
        @{ Provider='Ntfs'; Id='55,98,130'; Category='File system issue'; Severity='High'; Meaning='NTFS reported file system corruption or repair activity.'; Action='Review CHKDSK results and consider offline repair.' },
        @{ Provider='Microsoft-Windows-WHEA-Logger'; Id='1,17,18,19,47'; Category='Hardware error'; Severity='High'; Meaning='WHEA detected a CPU, memory, PCIe, or other hardware issue.'; Action='Check firmware, drivers, thermals, RAM, and hardware diagnostics.' },
        @{ Provider='BugCheck'; Id='1001'; Category='BSOD'; Severity='High'; Meaning='Windows bugchecked and likely produced a dump file.'; Action='Review dump files and recent driver/hardware changes.' },
        @{ Provider='Service Control Manager'; Id='7000,7001,7009,7011,7022,7023,7024,7031,7034'; Category='Service failure'; Severity='Medium'; Meaning='One or more services failed to start, stopped, or timed out.'; Action='Review affected service and dependency chain.' },
        @{ Provider='Microsoft-Windows-WindowsUpdateClient'; Id='20,25,31,34'; Category='Windows Update failure'; Severity='Medium'; Meaning='Windows Update reported installation, download, or scan issues.'; Action='Review Windows Update result and pending reboot state.' },
        @{ Provider='Microsoft-Windows-User Profile Service'; Id='1500,1502,1508,1511,1515'; Category='User profile issue'; Severity='Medium'; Meaning='Windows logged temporary profile or user profile load issues.'; Action='Check profile state, disk free space, and profile corruption.' }
    )
    $rows = @()
    foreach ($rule in $rules) {
        $ids = @($rule.Id.ToString().Split(',') | ForEach-Object { [int]$_.Trim() })
        $matches = @($Events | Where-Object { $_.ProviderName -like ('*' + $rule.Provider + '*') -and $ids -contains [int]$_.Id })
        if (@($matches).Count -gt 0) {
            $latest = @($matches | Sort-Object TimeCreated -Descending | Select-Object -First 1)
            $rows += [PSCustomObject]@{ Category=$rule.Category; Severity=$rule.Severity; Provider=$rule.Provider; EventIds=($ids -join ','); Count=@($matches).Count; MostRecent=$latest[0].TimeCreated; Meaning=$rule.Meaning; RecommendedAction=$rule.Action }
        }
    }
    if (@($rows).Count -eq 0) { Add-ActionLog -Step 'Reliability interpretation' -Status 'Complete' -Details 'No mapped reliability events found.' }
    else { Add-ActionLog -Step 'Reliability interpretation' -Status 'Complete' -Details ("Mapped {0} reliability event categories." -f @($rows).Count) }
    return @($rows)
}

function Get-TeamsOneDriveHealth {
    $results = @()
    $profileRoot = Join-Path $env:SystemDrive 'Users'
    if (Test-Path -LiteralPath $profileRoot) {
        foreach ($profile in @(Get-ChildItem -LiteralPath $profileRoot -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -notin @('Default','Default User','Public','All Users') })) {
            $paths = @(
                @{ App='Teams Classic'; Path=(Join-Path $profile.FullName 'AppData\Roaming\Microsoft\Teams') },
                @{ App='New Teams'; Path=(Join-Path $profile.FullName 'AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams') },
                @{ App='OneDrive'; Path=(Join-Path $profile.FullName 'AppData\Local\Microsoft\OneDrive') },
                @{ App='OneDrive Logs'; Path=(Join-Path $profile.FullName 'AppData\Local\Microsoft\OneDrive\logs') }
            )
            foreach ($entry in $paths) {
                $size = Get-FolderSizeInfo -Path $entry.Path
                $results += [PSCustomObject]@{ Scope='ProfileCache'; UserProfile=$profile.Name; App=$entry.App; Path=$entry.Path; Exists=$size.Exists; SizeGB=$size.SizeGB; FileCount=$size.FileCount; Details='' }
            }
        }
    }
    $oneDriveProc = @(Get-Process -Name OneDrive -ErrorAction SilentlyContinue)
    $teamsProc = @(Get-Process -Name @('Teams','ms-teams','msteams') -ErrorAction SilentlyContinue)
    $results += [PSCustomObject]@{ Scope='Process'; UserProfile='System'; App='OneDrive'; Path=''; Exists=(@($oneDriveProc).Count -gt 0); SizeGB=0; FileCount=@($oneDriveProc).Count; Details=('Running process count: {0}' -f @($oneDriveProc).Count) }
    $results += [PSCustomObject]@{ Scope='Process'; UserProfile='System'; App='Teams'; Path=''; Exists=(@($teamsProc).Count -gt 0); SizeGB=0; FileCount=@($teamsProc).Count; Details=('Running process count: {0}' -f @($teamsProc).Count) }
    try {
        $od = Get-Item -LiteralPath (Join-Path $env:LOCALAPPDATA 'Microsoft\OneDrive\OneDrive.exe') -ErrorAction SilentlyContinue
        if ($od) { $results += [PSCustomObject]@{ Scope='Version'; UserProfile=$env:USERNAME; App='OneDrive'; Path=$od.FullName; Exists=$true; SizeGB=0; FileCount=1; Details=$od.VersionInfo.FileVersion } }
    } catch { }
    try {
        $cloudEvents = @(Get-WinEvent -FilterHashtable @{LogName='System'; ProviderName='Microsoft-Windows-CloudFiles'; StartTime=(Get-Date).AddDays(-$EventLookbackDays)} -ErrorAction SilentlyContinue | Select-Object -First 20 TimeCreated, Id, ProviderName, Message)
        if (@($cloudEvents).Count -gt 0) { $results += [PSCustomObject]@{ Scope='Events'; UserProfile='System'; App='OneDrive/CloudFiles'; Path='System event log'; Exists=$true; SizeGB=0; FileCount=@($cloudEvents).Count; Details=('CloudFiles events in lookback: {0}' -f @($cloudEvents).Count) } }
    } catch { }
    Add-ActionLog -Step 'Teams and OneDrive health' -Status 'Complete' -Details ("Collected {0} Teams/OneDrive records." -f @($results).Count)
    return @($results)
}

function Send-MaintenanceNotifications {
    param([object]$Summary,[string]$HtmlReportPath,[string]$JsonReportPath,[string]$CsvReportPath)
    $subject = ("{0}: {1} {2}/100 {3}" -f $NotificationSubjectPrefix, $Summary.ComputerName, $Summary.HealthScore, $Summary.HealthStatus)
    $diskFreeAfterText = ('{0} GB ({1}%)' -f $Summary.DiskFreeAfterGB, $Summary.DiskFreeAfterPercent)
    $body = @"
PC maintenance completed.

Computer: $($Summary.ComputerName)
Mode: $($Summary.Mode)
Health: $($Summary.HealthScore)/100 - $($Summary.HealthStatus)
Pending reboot: $($Summary.PendingReboot)
Disk free after: $diskFreeAfterText
Report: $HtmlReportPath
"@
    if (-not [string]::IsNullOrWhiteSpace($EmailReportTo)) {
        try {
            if ([string]::IsNullOrWhiteSpace($SmtpServer) -or [string]::IsNullOrWhiteSpace($EmailFrom)) { throw 'EmailReportTo requires SmtpServer and EmailFrom.' }
            $attachments = @($HtmlReportPath,$JsonReportPath,$CsvReportPath) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }
            $mailParams = @{ To=$EmailReportTo; From=$EmailFrom; Subject=$subject; Body=$body; SmtpServer=$SmtpServer; Port=$SmtpPort; Attachments=$attachments }
            if ($SmtpUseSsl) { $mailParams.UseSsl = $true }
            Send-MailMessage @mailParams
            Add-ActionLog -Step 'Email notification' -Status 'Complete' -Details ("Sent to {0}." -f $EmailReportTo)
        } catch { Add-ActionLog -Step 'Email notification' -Status 'Failed' -Details $_.Exception.Message }
    }
    if (-not [string]::IsNullOrWhiteSpace($TeamsWebhookUrl)) {
        try {
            $teamsMessage = ("**{0}**`n`nComputer: **{1}**`nHealth: **{2}/100 - {3}**`nPending reboot: **{4}**`nDisk free after: **{5} GB ({6}%)**`nReport folder: {7}" -f $subject,$Summary.ComputerName,$Summary.HealthScore,$Summary.HealthStatus,$Summary.PendingReboot,$Summary.DiskFreeAfterGB,$Summary.DiskFreeAfterPercent,$Summary.RunFolder)
            $payload = @{ text = $teamsMessage } | ConvertTo-Json -Depth 4
            Invoke-RestMethod -Uri $TeamsWebhookUrl -Method Post -ContentType 'application/json' -Body $payload -ErrorAction Stop | Out-Null
            Add-ActionLog -Step 'Teams webhook notification' -Status 'Complete' -Details 'Posted summary to Teams webhook.'
        } catch { Add-ActionLog -Step 'Teams webhook notification' -Status 'Failed' -Details $_.Exception.Message }
    }
}

function Get-HealthScore {
    param(
        [object]$BeforeDisk,
        [object]$AfterDisk,
        [object]$PendingReboot,
        [object]$Resource,
        [object[]]$StorageHealth,
        [object]$Defender,
        [object[]]$ProblemDevices,
        [object[]]$ServiceIssues,
        [object]$Sfc,
        [object]$Dism,
        [object]$Chkdsk,
        [object]$CrashDumps,
        [object[]]$NotableEvents
    )

    $score = 100
    $reasons = @()

    if ($AfterDisk) {
        $afterFreePercentForScore = 0
        try { $afterFreePercentForScore = [double]$AfterDisk.FreePercent } catch { $afterFreePercentForScore = 0 }
        if ($afterFreePercentForScore -lt $LowDiskWarnPercent) {
            $score = $score - 20
            $reasons += ('Low disk free space ({0}%)' -f $AfterDisk.FreePercent)
        }
    }

    if ($PendingReboot -and $PendingReboot.RebootRequired) {
        $score = $score - 10
        $reasons += 'Pending reboot'
    }

    $resourceIssueCountForScore = 0
    if ($Resource -and $Resource.Issues) {
        $resourceIssueCountForScore = Get-ObjectCountSafe $Resource.Issues
    }
    if ($resourceIssueCountForScore -gt 0) {
        $score = $score - 10
        $reasons += ([string]($Resource.Issues -join ', '))
    }

    $badStorage = @()
    if ($StorageHealth) {
        $badStorage = @($StorageHealth | Where-Object { $_.HealthStatus -and $_.HealthStatus -notmatch 'Healthy|OK' })
    }
    $badStorageCountForScore = Get-ObjectCountSafe $badStorage
    if ($badStorageCountForScore -gt 0) {
        $score = $score - 30
        $reasons += 'Storage health warning'
    }

    if ($Defender -and $Defender.Available) {
        if (($Defender.RealTimeProtectionEnabled -eq $false) -or ($Defender.AMServiceEnabled -eq $false)) {
            $score = $score - 20
            $reasons += 'Defender protection disabled'
        }
    }

    $problemDeviceCountForScore = Get-ObjectCountSafe $ProblemDevices
    if ($problemDeviceCountForScore -gt 0) {
        $score = $score - 10
        $reasons += ('Device Manager issues: {0}' -f $problemDeviceCountForScore)
    }

    $serviceIssueCountForScore = Get-ObjectCountSafe $ServiceIssues
    if ($serviceIssueCountForScore -gt 10) {
        $score = $score - 10
        $reasons += ('Many automatic services stopped: {0}' -f $serviceIssueCountForScore)
    }

    if ($Sfc -and $Sfc.Summary -match 'unable|could not|corrupt') {
        $score = $score - 15
        $reasons += 'SFC reported corruption/repair issue'
    }

    if ($Dism) {
        if (($Dism.Status -match 'Failed') -or ($Dism.Summary -match 'error')) {
            $score = $score - 15
            $reasons += 'DISM reported an error'
        }
    }

    if ($Chkdsk -and $Chkdsk.Status -match 'Queued') {
        $score = $score - 10
        $reasons += 'CHKDSK fix queued'
    }

    $crashDumpCountForScore = 0
    if ($CrashDumps -and $null -ne $CrashDumps.DumpCount) {
        try { $crashDumpCountForScore = [int]$CrashDumps.DumpCount } catch { $crashDumpCountForScore = 0 }
    }
    if ($crashDumpCountForScore -gt 0) {
        $score = $score - 10
        $reasons += ('Crash dumps found: {0}' -f $crashDumpCountForScore)
    }

    $criticalEvents = @()
    if ($NotableEvents) {
        $criticalEvents = @($NotableEvents | Where-Object { $_.LevelDisplayName -eq 'Critical' })
    }
    $criticalEventCountForScore = Get-ObjectCountSafe $criticalEvents
    if ($criticalEventCountForScore -gt 0) {
        $score = $score - 10
        $reasons += ('Critical event log entries: {0}' -f $criticalEventCountForScore)
    }

    if ($score -lt 0) { $score = 0 }

    if ($score -ge 85) {
        $status = 'Healthy'
    }
    elseif ($score -ge 65) {
        $status = 'Needs review'
    }
    else {
        $status = 'Needs attention'
    }

    return [PSCustomObject]@{
        Score   = $score
        Status  = $status
        Reasons = ($reasons -join '; ')
    }
}

function Get-TechnicianSummary {
    param([object]$Health,[object]$BeforeDisk,[object]$AfterDisk,[object]$PendingReboot,[object]$Chkdsk,[object]$Sfc,[object]$Dism,[object]$WindowsUpdate,[object[]]$RemovedTasks,[object[]]$Profiles,[object]$WindowsOld)
    $gained = 0
    if ($BeforeDisk -and $AfterDisk) { $gained = [math]::Round([double]$AfterDisk.FreeGB - [double]$BeforeDisk.FreeGB, 2) }
    $parts = @()
    $parts += "Health score: $($Health.Score)/100 ($($Health.Status))."
    if ($AfterDisk) { $parts += ('Disk free changed by {0} GB; now {1} GB free ({2}%).' -f $gained, $AfterDisk.FreeGB, $AfterDisk.FreePercent) }
    if ($PendingReboot -and $PendingReboot.RebootRequired) { $parts += "Reboot required: $($PendingReboot.Reasons)." } else { $parts += 'No pending reboot detected.' }
    if ($Chkdsk) { $parts += "CHKDSK: $($Chkdsk.Summary)" }
    if ($Sfc) { $parts += "SFC: $($Sfc.Summary)" }
    if ($Dism) { $parts += "DISM: $($Dism.Summary)" }
    if ($WindowsUpdate) { $parts += "Windows Update: $($WindowsUpdate.Details)" }
    $removedCount = Get-ObjectCountSafe $RemovedTasks
    if ($removedCount -gt 0) { $parts += "Removed $removedCount stale scheduled task(s)." }
    $staleProfiles = @($Profiles | Where-Object { $_.StaleProfileCandidate -eq $true })
    $staleCount = Get-ObjectCountSafe $staleProfiles
    if ($staleCount -gt 0) { $parts += "Review $staleCount stale local profile candidate(s)." }
    if ($WindowsOld -and $WindowsOld.Exists) { $parts += "Windows.old detected: $($WindowsOld.SizeGB) GB." }
    if ($Health -and $Health.Reasons) { $parts += "Key flags: $($Health.Reasons)" }
    return ($parts -join ' ')
}


function Get-SelfParseResult {
    $targetPath = $PSCommandPath
    if ([string]::IsNullOrWhiteSpace($targetPath)) { $targetPath = $MyInvocation.MyCommand.Path }
    $tokens = $null
    $errors = $null
    try {
        [System.Management.Automation.Language.Parser]::ParseFile($targetPath, [ref]$tokens, [ref]$errors) | Out-Null
        if ($errors -and @($errors).Count -gt 0) {
            return [PSCustomObject]@{ Status='Failed'; ErrorCount=@($errors).Count; Errors=($errors | ForEach-Object { $_.Message }) -join ' | ' }
        }
        return [PSCustomObject]@{ Status='OK'; ErrorCount=0; Errors='' }
    } catch {
        return [PSCustomObject]@{ Status='Failed'; ErrorCount=1; Errors=$_.Exception.Message }
    }
}

function Get-StepDiskFreeGB {
    try {
        $letter = $DriveLetter.TrimEnd(':')
        $deviceId = ('{0}:' -f $letter)
        $disk = Get-CimInstance Win32_LogicalDisk -Filter ("DeviceID='{0}'" -f $deviceId) -ErrorAction Stop
        return [math]::Round([double]$disk.FreeSpace / 1GB, 2)
    } catch { return $null }
}

function Add-StepMetric {
    param([string]$Name,[datetime]$Start,[datetime]$End,[object]$FreeBeforeGB,[object]$FreeAfterGB,[string]$Status,[string]$Details)
    $delta = $null
    try {
        if ($null -ne $FreeBeforeGB -and $null -ne $FreeAfterGB) { $delta = [math]::Round([double]$FreeAfterGB - [double]$FreeBeforeGB, 2) }
    } catch { $delta = $null }
    $script:StepMetrics += [PSCustomObject]@{
        Step = $Name
        StartTime = $Start.ToString('yyyy-MM-dd HH:mm:ss')
        EndTime = $End.ToString('yyyy-MM-dd HH:mm:ss')
        Duration = (New-TimeSpan -Start $Start -End $End).ToString()
        FreeBeforeGB = $FreeBeforeGB
        FreeAfterGB = $FreeAfterGB
        DeltaGB = $delta
        Status = $Status
        Details = $Details
    }
}

function Get-ShadowStorageInfo {
    $items = @()
    try {
        $raw = cmd.exe /c 'vssadmin list shadowstorage' 2>&1 | Out-String
        if ([string]::IsNullOrWhiteSpace($raw)) { return @() }
        $blocks = [regex]::Split($raw, '(?m)^\s*$')
        foreach ($block in $blocks) {
            if ($block -notmatch 'Shadow Copy Storage association') { continue }
            $forVolume = ''
            $onVolume = ''
            $used = ''
            $allocated = ''
            $maximum = ''
            foreach ($line in ($block -split "`r?`n")) {
                if ($line -match 'For volume:\s*(.+)$') { $forVolume = $matches[1].Trim() }
                elseif ($line -match 'Shadow Copy Storage volume:\s*(.+)$') { $onVolume = $matches[1].Trim() }
                elseif ($line -match 'Used Shadow Copy Storage space:\s*(.+)$') { $used = $matches[1].Trim() }
                elseif ($line -match 'Allocated Shadow Copy Storage space:\s*(.+)$') { $allocated = $matches[1].Trim() }
                elseif ($line -match 'Maximum Shadow Copy Storage space:\s*(.+)$') { $maximum = $matches[1].Trim() }
            }
            $items += [PSCustomObject]@{ ForVolume=$forVolume; StorageVolume=$onVolume; Used=$used; Allocated=$allocated; Maximum=$maximum }
        }
        if (@($items).Count -eq 0) { $items += [PSCustomObject]@{ ForVolume=''; StorageVolume=''; Used=''; Allocated=''; Maximum=''; Raw=$raw } }
    } catch {
        $items += [PSCustomObject]@{ ForVolume=''; StorageVolume=''; Used=''; Allocated=''; Maximum=''; Raw=$_.Exception.Message }
    }
    return @($items)
}

function Set-ShadowStorageMaximum {
    param([string]$MaxSize)
    if ([string]::IsNullOrWhiteSpace($MaxSize)) { return 'Not requested' }
    try {
        $letter = $DriveLetter.TrimEnd(':')
        $drive = ('{0}:' -f $letter)
        $output = cmd.exe /c ("vssadmin resize shadowstorage /For={0} /On={0} /MaxSize={1}" -f $drive, $MaxSize) 2>&1 | Out-String
        Add-ActionLog -Step 'Shadow storage limit' -Status 'Complete' -Details $output.Trim()
        return $output.Trim()
    } catch {
        Add-ActionLog -Step 'Shadow storage limit' -Status 'Failed' -Details $_.Exception.Message
        return $_.Exception.Message
    }
}

function Get-PreflightChecks {
    $checks = @()
    $checks += [PSCustomObject]@{ Check='Running as Administrator'; Status=if(Test-IsAdmin){'OK'}else{'Warning'}; Details=if(Test-IsAdmin){'Elevated session.'}else{'Not elevated. Some actions may fail.'} }
    try { $root = Resolve-ReportRoot -RequestedPath $ReportRoot; $checks += [PSCustomObject]@{ Check='Report folder writable'; Status='OK'; Details=$root } } catch { $checks += [PSCustomObject]@{ Check='Report folder writable'; Status='Warning'; Details=$_.Exception.Message } }
    try {
        $disk = Get-DiskSpace -Letter $DriveLetter
        $status = if ($disk.FreePercent -lt 10) { 'Warning' } else { 'OK' }
        $checks += [PSCustomObject]@{ Check='Free disk space'; Status=$status; Details=('{0} GB free ({1}%)' -f $disk.FreeGB, $disk.FreePercent) }
    } catch { $checks += [PSCustomObject]@{ Check='Free disk space'; Status='Warning'; Details=$_.Exception.Message } }
    try {
        $pending = Get-PendingRebootStatus
        $checks += [PSCustomObject]@{ Check='Pending reboot before run'; Status=if($pending.RebootRequired){'Warning'}else{'OK'}; Details=$pending.Reasons }
    } catch { $checks += [PSCustomObject]@{ Check='Pending reboot before run'; Status='Warning'; Details=$_.Exception.Message } }
    try {
        $battery = Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($battery) { $checks += [PSCustomObject]@{ Check='Laptop AC/battery'; Status='Info'; Details=('Battery status: {0}; Charge: {1}%' -f $battery.BatteryStatus, $battery.EstimatedChargeRemaining) } }
        else { $checks += [PSCustomObject]@{ Check='Laptop AC/battery'; Status='OK'; Details='No battery detected.' } }
    } catch { }
    return @($checks)
}

function Get-WindowsUpdateDiagnostics {
    $items = @()
    try {
        $svc = Get-Service wuauserv -ErrorAction SilentlyContinue
        if ($svc) { $items += [PSCustomObject]@{ Category='Service'; Name='wuauserv'; Status=$svc.Status; Details=('StartType: {0}' -f $svc.StartType) } }
    } catch { }
    try {
        $lastHotfix = Get-HotFix -ErrorAction SilentlyContinue | Sort-Object InstalledOn -Descending | Select-Object -First 1
        if ($lastHotfix) { $items += [PSCustomObject]@{ Category='Last hotfix'; Name=$lastHotfix.HotFixID; Status='Info'; Details=('InstalledOn: {0}; Description: {1}' -f $lastHotfix.InstalledOn, $lastHotfix.Description) } }
    } catch { }
    try {
        $failures = Get-WinEvent -FilterHashtable @{LogName='System'; ProviderName='Microsoft-Windows-WindowsUpdateClient'; Level=2; StartTime=(Get-Date).AddDays(-30)} -ErrorAction SilentlyContinue | Select-Object -First 10 TimeCreated, Id, ProviderName, Message
        foreach ($f in @($failures)) { $items += [PSCustomObject]@{ Category='WU event error'; Name=('Event {0}' -f $f.Id); Status='Warning'; Details=([string]$f.Message).Substring(0, [Math]::Min(250, ([string]$f.Message).Length)) } }
    } catch { }
    if (@($items).Count -eq 0) { $items += [PSCustomObject]@{ Category='Windows Update'; Name='Diagnostics'; Status='Info'; Details='No diagnostics returned.' } }
    return @($items)
}

function Get-BitLockerHealth {
    $items = @()
    try {
        $cmd = Get-Command Get-BitLockerVolume -ErrorAction SilentlyContinue
        if ($cmd) {
            foreach ($v in @(Get-BitLockerVolume -ErrorAction SilentlyContinue)) {
                $items += [PSCustomObject]@{ MountPoint=$v.MountPoint; VolumeStatus=$v.VolumeStatus; ProtectionStatus=$v.ProtectionStatus; LockStatus=$v.LockStatus; EncryptionPercentage=$v.EncryptionPercentage; KeyProtectorCount=@($v.KeyProtector).Count }
            }
        } else { $items += [PSCustomObject]@{ MountPoint=''; VolumeStatus='Unavailable'; ProtectionStatus=''; LockStatus=''; EncryptionPercentage=''; KeyProtectorCount=''; Details='Get-BitLockerVolume unavailable.' } }
    } catch { $items += [PSCustomObject]@{ MountPoint=''; VolumeStatus='Failed'; ProtectionStatus=''; LockStatus=''; EncryptionPercentage=''; KeyProtectorCount=''; Details=$_.Exception.Message } }
    return @($items)
}

function Get-PowerHealth {
    $items = @()
    foreach ($cmd in @('/requests','/lastwake','/a')) {
        try {
            $output = powercfg.exe $cmd 2>&1 | Out-String
            $items += [PSCustomObject]@{ Command=('powercfg {0}' -f $cmd); Details=(Limit-Text $output 1200) }
        } catch { $items += [PSCustomObject]@{ Command=('powercfg {0}' -f $cmd); Details=$_.Exception.Message } }
    }
    return @($items)
}

function Invoke-ComponentStoreAnalysis {
    $result = [PSCustomObject]@{ Name='DISM AnalyzeComponentStore'; Command='dism.exe /online /cleanup-image /analyzecomponentstore'; Status='Skipped'; ExitCode=$null; Output=''; Errors=''; Summary='' }
    if ($SkipComponentStoreAnalysis) { $result.Summary='SkipComponentStoreAnalysis was supplied.'; return $result }
    $dism = Join-Path $env:windir 'System32\dism.exe'
    if (-not (Test-Path -LiteralPath $dism)) { $dism='dism.exe' }
    $result = Invoke-ExternalMaintenanceCommand -Name 'DISM_AnalyzeComponentStore' -FilePath $dism -Arguments @('/online','/cleanup-image','/analyzecomponentstore') -TimeoutMinutes $ComponentCleanupTimeoutMinutes
    $result.Name='DISM AnalyzeComponentStore'
    $result.Command='dism.exe /online /cleanup-image /analyzecomponentstore'
    $result.Summary='Review output for component store size and whether cleanup is recommended.'
    return $result
}

function Remove-OldReportFolders {
    param([string]$BasePath,[int]$RetentionDays)
    $results = @()
    if (-not $CleanOldReports) { return @([PSCustomObject]@{ Path=$BasePath; Status='Skipped'; Details='CleanOldReports was not supplied.' }) }
    try {
        $cutoff = (Get-Date).AddDays(-1 * [math]::Abs($RetentionDays))
        foreach ($dir in @(Get-ChildItem -LiteralPath $BasePath -Directory -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt $cutoff })) {
            $full = $dir.FullName
            try { Remove-Item -LiteralPath $full -Recurse -Force -ErrorAction Stop; $results += [PSCustomObject]@{ Path=$full; Status='Removed'; Details=('Older than {0} days.' -f $RetentionDays) } }
            catch { $results += [PSCustomObject]@{ Path=$full; Status='Failed'; Details=$_.Exception.Message } }
        }
        if (@($results).Count -eq 0) { $results += [PSCustomObject]@{ Path=$BasePath; Status='None'; Details='No old report folders matched retention policy.' } }
    } catch { $results += [PSCustomObject]@{ Path=$BasePath; Status='Failed'; Details=$_.Exception.Message } }
    return @($results)
}

function Compress-RunFolder {
    param([string]$FolderPath)
    if (-not $ZipReportFolder) { return '' }
    try {
        $zip = ($FolderPath.TrimEnd('\') + '.zip')
        if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force -ErrorAction SilentlyContinue }
        Compress-Archive -Path (Join-Path $FolderPath '*') -DestinationPath $zip -Force
        Add-ActionLog -Step 'Report ZIP' -Status 'Complete' -Details $zip
        return $zip
    } catch { Add-ActionLog -Step 'Report ZIP' -Status 'Failed' -Details $_.Exception.Message; return '' }
}

function Get-RmmExitCode {
    param([object]$Health,[object]$PendingReboot,[bool]$ReportFailed)
    if ($ReportFailed) { return 4 }
    if ($Health.Score -lt 65) { return 3 }
    if ($PendingReboot.RebootRequired) { return 2 }
    if ($Health.Score -lt 85) { return 1 }
    return 0
}

function Invoke-Safely {
    param([string]$Name,[scriptblock]$ScriptBlock,[object]$DefaultValue = $null)
    $stepStart = Get-Date
    $freeBefore = Get-StepDiskFreeGB
    $status = 'Complete'
    $details = ''
    try {
        Write-Step $Name
        $result = & $ScriptBlock
        $details = 'Completed.'
        return $result
    } catch {
        $status = 'Failed'
        $details = $_.Exception.Message
        Add-ActionLog -Step $Name -Status 'Failed' -Details $details
        Write-Warning ("{0} failed: {1}" -f $Name, $details)
        return $DefaultValue
    } finally {
        $stepEnd = Get-Date
        $freeAfter = Get-StepDiskFreeGB
        Add-StepMetric -Name $Name -Start $stepStart -End $stepEnd -FreeBeforeGB $freeBefore -FreeAfterGB $freeAfter -Status $status -Details $details
    }
}

function New-MaintenanceReport {
    param(
        [object]$BeforeDisk,[object]$AfterDisk,[object[]]$TempCleanup,[object]$RecycleBinResult,[object]$ComponentCleanupResult,[object]$DeliveryOptimizationResult,[object[]]$NotableEvents,[object[]]$EventSummary,[object[]]$EventInterpretation,[object[]]$StaleTasks,[object[]]$RemovedTasks,[object[]]$RemovedProfiles,[object[]]$EventLogMaintenance,[object]$ResourceSnapshot,[object[]]$TopCpuProcesses,[object[]]$TopMemoryProcesses,[object]$WindowsUpdateResult,[object[]]$RestorePointResults,[object]$ChkdskResult,[object]$SfcResult,[object]$DismRestoreHealthResult,[object[]]$UserProfileSizes,[object]$PendingReboot,[object[]]$StorageHealth,[object]$BatteryHealth,[object[]]$ProblemDevices,[object[]]$StartupItems,[object[]]$ServiceIssues,[object]$DefenderStatus,[object]$CrashDumpSummary,[object]$CrashDumpCleanupResult,[object[]]$StorageHotspots,[object]$WindowsOldInfo,[object]$WindowsOldCleanupResult,[object[]]$TeamsOneDriveHealth,[object[]]$CleanupCategoryMetrics,[object]$HealthScore,[string]$TechnicianSummary,[string]$OutputPath
    )
    $diskGained = 0
    if ($BeforeDisk -and $AfterDisk) { $diskGained = [math]::Round([double]$AfterDisk.FreeGB - [double]$BeforeDisk.FreeGB, 2) }
    $removedGB = ($TempCleanup | Measure-Object -Property RemovedGB -Sum).Sum; if ($null -eq $removedGB) { $removedGB = 0 }
    $potentialGB = ($TempCleanup | Measure-Object -Property PotentialGB -Sum).Sum; if ($null -eq $potentialGB) { $potentialGB = 0 }
    $issuesText = if ($ResourceSnapshot -and $ResourceSnapshot.Issues -and (Get-ObjectCountSafe $ResourceSnapshot.Issues) -gt 0) { ($ResourceSnapshot.Issues -join ', ') } else { 'None flagged' }
    $notableTrimmed = @($NotableEvents | Select-Object -First 50 TimeCreated, LogName, LevelDisplayName, Id, ProviderName, @{Name='Message';Expression={ $m=[string]$_.Message; if($m.Length -gt 350){$m.Substring(0,350)+'...'}else{$m} }})
    $actions = @($script:ActionLog)
    $profileData = @($UserProfileSizes | Sort-Object SizeGB -Descending)
    $chkdskOutput = ConvertTo-HtmlEncodedText (Limit-Text (($ChkdskResult.ScanOutput + "`r`n" + $ChkdskResult.QueueOutput + "`r`n" + $ChkdskResult.Errors)) 24000)
    $sfcOutput = ConvertTo-HtmlEncodedText (Limit-Text (($SfcResult.Output + "`r`n" + $SfcResult.Errors)) 24000)
    $dismOutput = ConvertTo-HtmlEncodedText (Limit-Text (($DismRestoreHealthResult.Output + "`r`n" + $DismRestoreHealthResult.Errors)) 24000)
    $modeText = if ($script:IsAuditOnly) { 'Audit' } else { 'Remediate' }
    $presetText = if ($script:MaintenancePreset) { $script:MaintenancePreset } else { 'Custom' }
    $logoDataUri = Get-LogoDataUri -FolderPath $PSScriptRoot
    $logoHtml = if (-not [string]::IsNullOrWhiteSpace($logoDataUri)) { "<div class='header-logo'><img src='$logoDataUri' alt='Company logo'></div>" } else { '' }
    $reportFolder = Split-Path -Parent $OutputPath
    $beforeFreeText = ('{0}% free on {1}' -f $BeforeDisk.FreePercent, $BeforeDisk.Drive)
    $afterFreeText = ('{0}% free on {1}' -f $AfterDisk.FreePercent, $AfterDisk.Drive)

    $html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset='utf-8'>
<title>PC Maintenance Report</title>
<style>
body { font-family: Segoe UI, Arial, sans-serif; background:#f4f6f8; color:#202a33; margin:0; padding:24px; }
.container { max-width: 1280px; margin: 0 auto; }
.header { background:#ffffff; border-radius:16px; padding:24px; box-shadow:0 3px 18px rgba(0,0,0,.10); margin-bottom:18px; border-top:6px solid #ed174c; }
.header-row { display:flex; align-items:center; justify-content:space-between; gap:24px; flex-wrap:wrap; }
.header-copy { flex:1 1 420px; min-width:280px; }
.header-copy h1 { margin-bottom:10px; font-size:34px; letter-spacing:-0.5px; color:#263238; }
.report-kicker { font-size:12px; font-weight:700; text-transform:uppercase; letter-spacing:1.2px; color:#ed174c; margin-bottom:6px; }
.header-logo { flex:0 0 auto; margin-left:auto; text-align:right; }
.header-logo img { max-width:340px; width:100%; height:auto; max-height:120px; object-fit:contain; }
.grid { display:grid; grid-template-columns: repeat(auto-fit, minmax(230px, 1fr)); gap:14px; margin-bottom:18px; }
.card { background:#ffffff; border-radius:14px; padding:18px; box-shadow:0 2px 14px rgba(0,0,0,.08); margin-bottom:18px; border-left:4px solid #eef3f8; }
.summary-card { border-left-color:#ed174c; background:linear-gradient(135deg,#ffffff 0%,#fff5f7 100%); }
.footer { color:#64707d; font-size:12px; text-align:center; padding:18px 8px 4px 8px; }
.big { font-size: 28px; font-weight: 700; margin: 8px 0; }
.small { color:#5b6775; font-size: 13px; }
.status-ok { color:#0b7a3b; font-weight:700; }
.status-warn { color:#a85d00; font-weight:700; }
.status-bad { color:#b00020; font-weight:700; }
table { border-collapse:collapse; width:100%; background:#fff; margin-top:10px; }
th, td { border-bottom:1px solid #e5e9f0; padding:8px; text-align:left; vertical-align:top; font-size:13px; }
th { background:#eef3f8; font-weight:700; color:#263238; }
h1, h2, h3 { margin-top:0; }
.bar { width:100%; height:18px; background:#e9edf3; border-radius:999px; overflow:hidden; margin:5px 0 12px 0; }
.bar-fill { height:100%; background:#ed174c; border-radius:999px; }
.metric-label { font-size:13px; color:#4a5564; margin-top:10px; }
code { background:#eef3f8; padding:2px 5px; border-radius:5px; }
pre { background:#111827; color:#f3f4f6; padding:12px; border-radius:10px; overflow:auto; max-height:420px; white-space:pre-wrap; }
</style>
</head>
<body><div class='container'>
<div class='header'>
<div class='header-row'>
<div class='header-copy'>
<div class='report-kicker'>Silicon Beach Endpoint Maintenance</div>
<h1>PC Maintenance Report</h1>
<div class='small'>Computer: <b>$env:COMPUTERNAME</b> | User: <b>$env:USERNAME</b> | Mode: <b>$modeText</b> | Preset: <b>$presetText</b> | Started: <b>$($script:RunStart.ToString('yyyy-MM-dd HH:mm:ss'))</b> | Finished: <b>$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))</b><br>Output folder: <b>$reportFolder</b></div>
</div>
$logoHtml
</div>
</div>

<div class='grid'>
  <div class='card'><h3>Health Score</h3><div class='big'>$($HealthScore.Score)/100</div><div class='small'>$($HealthScore.Status)</div></div>
  <div class='card'><h3>Disk Free Before</h3><div class='big'>$($BeforeDisk.FreeGB) GB</div><div class='small'>$beforeFreeText</div></div>
  <div class='card'><h3>Disk Free After</h3><div class='big'>$($AfterDisk.FreeGB) GB</div><div class='small'>$afterFreeText</div></div>
  <div class='card'><h3>Disk Change</h3><div class='big'>$diskGained GB</div><div class='small'>Direct cleanup removed $removedGB GB; potential $potentialGB GB.</div></div>
  <div class='card'><h3>Pending Reboot</h3><div class='big'>$($PendingReboot.RebootRequired)</div><div class='small'>$($PendingReboot.Reasons)</div></div>
</div>

<div class='card summary-card'><h2>Technician Summary</h2><p>$TechnicianSummary</p></div>

<div class='card'>
<h2>Visual Metrics</h2>
$(New-BarHtml -Label 'Disk free after cleanup' -Percent $AfterDisk.FreePercent)
$(New-BarHtml -Label 'CPU usage' -Percent $ResourceSnapshot.CPUPercent)
$(New-BarHtml -Label 'Memory usage' -Percent $ResourceSnapshot.MemoryPercent)
<p><b>Disk queue length:</b> $($ResourceSnapshot.DiskQueueLength)</p>
<p><b>Resource issues:</b> $issuesText</p>
</div>

<div class='card'><h2>Health Score Reasons</h2><p>$($HealthScore.Reasons)</p></div>

<div class='card'><h2>v15 Step Duration and Disk Delta</h2><p class='small'>Each major step is measured for duration and disk free-space change. Negative delta usually means restore points, logs, repair staging or cache growth consumed space.</p>$(ConvertTo-HtmlTable -Data $script:StepMetrics -EmptyMessage 'No step metrics captured.')</div>
<div class='card'><h2>VSS / Restore Point Shadow Storage - Before</h2>$(ConvertTo-HtmlTable -Data $script:VssBefore -EmptyMessage 'No VSS data captured before run.')</div>
<div class='card'><h2>VSS / Restore Point Shadow Storage - After</h2>$(ConvertTo-HtmlTable -Data $script:VssAfter -EmptyMessage 'No VSS data captured after run.')</div>
<div class='card'><h2>Pre-flight Checks</h2>$(ConvertTo-HtmlTable -Data $script:PreflightChecks -EmptyMessage 'No pre-flight checks captured.')</div>
<div class='card'><h2>Windows Update Diagnostics</h2>$(ConvertTo-HtmlTable -Data $script:WindowsUpdateDiagnostics -EmptyMessage 'No Windows Update diagnostics captured.')</div>
<div class='card'><h2>Component Store Analysis</h2><p><b>Status:</b> $($script:ComponentStoreAnalysis.Status)</p><p><b>Summary:</b> $($script:ComponentStoreAnalysis.Summary)</p><pre>$(ConvertTo-HtmlEncodedText (Limit-Text ($script:ComponentStoreAnalysis.Output + "`r`n" + $script:ComponentStoreAnalysis.Errors) 18000))</pre></div>
<div class='card'><h2>BitLocker Health</h2>$(ConvertTo-HtmlTable -Data $script:BitLockerHealth -EmptyMessage 'No BitLocker data captured.')</div>
<div class='card'><h2>Power and Sleep Health</h2>$(ConvertTo-HtmlTable -Data $script:PowerHealth -EmptyMessage 'No power health data captured.')</div>

<div class='card'><h2>Restore Points</h2>$(ConvertTo-HtmlTable -Data $RestorePointResults -EmptyMessage 'Restore points skipped or unavailable.')</div>
<div class='card'><h2>Pending Reboot Detail</h2>$(ConvertTo-HtmlTable -Data @($PendingReboot) -EmptyMessage 'No pending reboot data.')</div>
<div class='card'><h2>Disk Space</h2>$(ConvertTo-HtmlTable -Data @($BeforeDisk, $AfterDisk) -EmptyMessage 'No disk data.')</div>
<div class='card'><h2>Temp and Cache Cleanup</h2><p class='small'>Audit mode reports potential cleanup. Remediate mode reports actual deleted files.</p>$(ConvertTo-HtmlTable -Data $TempCleanup -EmptyMessage 'No temp cleanup results.')</div>
<div class='card'><h2>System Repair and Disk Checks</h2>
<h3>CHKDSK</h3><p><b>Status:</b> $($ChkdskResult.Status)</p><p><b>Summary:</b> $($ChkdskResult.Summary)</p><pre>$chkdskOutput</pre>
<h3>SFC /scannow</h3><p><b>Status:</b> $($SfcResult.Status)</p><p><b>Exit code:</b> $($SfcResult.ExitCode)</p><p><b>Summary:</b> $($SfcResult.Summary)</p><pre>$sfcOutput</pre>
<h3>DISM RestoreHealth</h3><p><b>Status:</b> $($DismRestoreHealthResult.Status)</p><p><b>Exit code:</b> $($DismRestoreHealthResult.ExitCode)</p><p><b>Summary:</b> $($DismRestoreHealthResult.Summary)</p><pre>$dismOutput</pre>
</div>
<div class='card'><h2>Windows Update</h2><p><b>Status:</b> $($WindowsUpdateResult.Status)</p><p><b>Details:</b> $($WindowsUpdateResult.Details)</p>$(ConvertTo-HtmlTable -Data @($WindowsUpdateResult.Updates) -EmptyMessage 'No update list returned.')</div>
<div class='card'><h2>Local User Profile Sizes</h2><p class='small'>Stale profile candidates are unloaded, non-special profiles not used for at least $ProfileStaleDays days.</p>$(ConvertTo-HtmlTable -Data $profileData -EmptyMessage 'No profile size data.')<h3>Stale Profiles Removed</h3>$(ConvertTo-HtmlTable -Data $RemovedProfiles -EmptyMessage 'No profiles removed.')</div>
<div class='card'><h2>Storage Health and SMART Reliability</h2>$(ConvertTo-HtmlTable -Data $StorageHealth -EmptyMessage 'No storage health data returned.')</div>
<div class='card'><h2>Battery Health</h2>$(ConvertTo-HtmlTable -Data @($BatteryHealth) -EmptyMessage 'No battery data.')</div>
<div class='card'><h2>Device Manager Problem Devices</h2>$(ConvertTo-HtmlTable -Data $ProblemDevices -EmptyMessage 'No non-OK devices found.')</div>
<div class='card'><h2>Microsoft Defender Status</h2>$(ConvertTo-HtmlTable -Data @($DefenderStatus) -EmptyMessage 'Defender status unavailable.')</div>
<div class='card'><h2>Crash Dumps</h2>$(ConvertTo-HtmlTable -Data @($CrashDumpSummary) -EmptyMessage 'No crash dump data.')<h3>Crash Dump Cleanup</h3>$(ConvertTo-HtmlTable -Data @($CrashDumpCleanupResult) -EmptyMessage 'No crash dump cleanup result.')</div>
<div class='card'><h2>Storage Hotspots</h2><p class='small'>Large files over 100 MB from common user folders.</p>$(ConvertTo-HtmlTable -Data $StorageHotspots -EmptyMessage 'No large user files found or scan skipped.')</div>
<div class='card'><h2>Teams and OneDrive Health</h2>$(ConvertTo-HtmlTable -Data $TeamsOneDriveHealth -EmptyMessage 'No Teams or OneDrive data returned.')</div>
<div class='card'><h2>Before/After Cleanup Metrics by Category</h2>$(ConvertTo-HtmlTable -Data $CleanupCategoryMetrics -EmptyMessage 'No cleanup category metrics returned.')</div>
<div class='card'><h2>Windows.old</h2>$(ConvertTo-HtmlTable -Data @($WindowsOldInfo) -EmptyMessage 'No Windows.old data.')<h3>Windows.old Cleanup</h3>$(ConvertTo-HtmlTable -Data @($WindowsOldCleanupResult) -EmptyMessage 'No Windows.old cleanup result.')</div>
<div class='grid'><div class='card'><h2>Top CPU Processes</h2>$(ConvertTo-HtmlTable -Data $TopCpuProcesses -EmptyMessage 'No CPU process data.')</div><div class='card'><h2>Top Memory Processes</h2>$(ConvertTo-HtmlTable -Data $TopMemoryProcesses -EmptyMessage 'No memory process data.')</div></div>
<div class='card'><h2>Reliability Interpretation</h2>$(ConvertTo-HtmlTable -Data $EventInterpretation -EmptyMessage 'No mapped reliability issues found.')</div>
<div class='card'><h2>Event Summary</h2>$(ConvertTo-HtmlTable -Data $EventSummary -EmptyMessage 'No event summary.')</div>
<div class='card'><h2>Notable Event Viewer Entries</h2>$(ConvertTo-HtmlTable -Data $notableTrimmed -EmptyMessage 'No notable events found.')</div>
<div class='card'><h2>Scheduled Task Audit</h2><h3>Stale Tasks Found</h3>$(ConvertTo-HtmlTable -Data $StaleTasks -EmptyMessage 'No eligible stale tasks found.')<h3>Tasks Removed</h3>$(ConvertTo-HtmlTable -Data $RemovedTasks -EmptyMessage 'No tasks removed.')</div>
<div class='card'><h2>Startup Items</h2>$(ConvertTo-HtmlTable -Data $StartupItems -EmptyMessage 'No startup items found.')</div>
<div class='card'><h2>Automatic Services Not Running</h2>$(ConvertTo-HtmlTable -Data $ServiceIssues -EmptyMessage 'No auto-start stopped services found.')</div>
<div class='card'><h2>Event Log Maintenance</h2>$(ConvertTo-HtmlTable -Data $EventLogMaintenance -EmptyMessage 'Event logs were not cleared or archived.')</div>
<div class='card'><h2>Other Cleanup</h2><p><b>Recycle Bin:</b> $RecycleBinResult</p><p><b>Component cleanup:</b> $ComponentCleanupResult</p><p><b>Delivery Optimization:</b> $DeliveryOptimizationResult</p></div>
<div class='card'><h2>Actions Performed</h2>$(ConvertTo-HtmlTable -Data $actions -EmptyMessage 'No actions logged.')</div>
<div class='footer'>Generated by Silicon Beach PC Maintenance | $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) | Report stored in $reportFolder</div>
</div></body></html>
"@
    $html | Out-File -LiteralPath $OutputPath -Encoding UTF8 -Force
    return $OutputPath
}

# Main run
if ($ValidateOnly) {
    $parseResult = Get-SelfParseResult
    if ($parseResult.Status -eq 'OK') { Write-Host 'ValidateOnly: script parsed successfully.' -ForegroundColor Green; exit 0 }
    Write-Host 'ValidateOnly: parser errors found.' -ForegroundColor Red
    Write-Host $parseResult.Errors
    exit 4
}
Import-PCMaintenanceConfig -Path $ConfigPath
$script:IsAuditOnly = ($AuditOnly -or $Mode -eq 'Audit')
Set-MaintenancePreset
if ($ComputerName -and @($ComputerName).Count -gt 0) { Invoke-RemoteMaintenance -Targets $ComputerName; return }
$BaseReportRoot = Resolve-ReportRoot -RequestedPath $ReportRoot
$RunFolderInfo = $null

try {
    $RunFolderInfo = New-MaintenanceRunFolder -BaseReportRoot $BaseReportRoot -AuditMode $script:IsAuditOnly
    $ReportRoot = $RunFolderInfo.RunFolder
    Add-ActionLog -Step 'Report folder' -Status 'Created' -Details ("Base: {0}; Run folder: {1}" -f $RunFolderInfo.BaseReportRoot, $RunFolderInfo.RunFolder)
} catch {
    $BaseReportRoot = Resolve-ReportRoot -RequestedPath ''
    $RunFolderInfo = New-MaintenanceRunFolder -BaseReportRoot $BaseReportRoot -AuditMode $script:IsAuditOnly
    $ReportRoot = $RunFolderInfo.RunFolder
    Add-ActionLog -Step 'Report folder' -Status 'Warning' -Details ("Primary folder failed. Fallback run folder: {0}. Error: {1}" -f $ReportRoot, $_.Exception.Message)
}

if (-not $SkipTranscript) {
    try {
        $script:TranscriptPath = Join-Path $ReportRoot ("PC_Maintenance_Transcript_{0}_{1}.txt" -f $env:COMPUTERNAME, $RunFolderInfo.Timestamp)
        Start-Transcript -Path $script:TranscriptPath -Force | Out-Null
    } catch { $script:TranscriptPath = $null }
}

if (-not (Test-IsAdmin)) { Add-ActionLog -Step 'Privilege check' -Status 'Warning' -Details 'Not running as Administrator. Some actions may fail.'; Write-Warning 'Not running as Administrator. Some cleanup/repair actions may fail.' }
else { Add-ActionLog -Step 'Privilege check' -Status 'OK' -Details 'Running as Administrator.' }
$runModeStatus = if ($script:IsAuditOnly) { 'Audit' } else { 'Remediate' }
Add-ActionLog -Step 'Run mode' -Status $runModeStatus -Details (('Mode parameter: {0}; Cleanup preset: {1}' -f $Mode, $script:MaintenancePreset))
$reportRetentionResult = Invoke-Safely -Name 'Report retention cleanup' -DefaultValue @() -ScriptBlock { Remove-OldReportFolders -BasePath $BaseReportRoot -RetentionDays $ReportRetentionDays }
$script:PreflightChecks = Invoke-Safely -Name 'Pre-flight checks' -DefaultValue @() -ScriptBlock { Get-PreflightChecks }
$script:VssBefore = Invoke-Safely -Name 'VSS shadow storage before' -DefaultValue @() -ScriptBlock { Get-ShadowStorageInfo }
$shadowLimitResult = Invoke-Safely -Name 'Shadow storage maximum' -DefaultValue 'Not requested' -ScriptBlock { Set-ShadowStorageMaximum -MaxSize $MaxShadowStorage }
$script:WindowsUpdateDiagnostics = Invoke-Safely -Name 'Windows Update diagnostics' -DefaultValue @() -ScriptBlock { Get-WindowsUpdateDiagnostics }
$script:BitLockerHealth = Invoke-Safely -Name 'BitLocker health' -DefaultValue @() -ScriptBlock { Get-BitLockerHealth }
$script:PowerHealth = Invoke-Safely -Name 'Power health' -DefaultValue @() -ScriptBlock { Get-PowerHealth }

$reportPath = Join-Path $ReportRoot ("PC_Maintenance_Report_{0}_{1}.html" -f $env:COMPUTERNAME, $RunFolderInfo.Timestamp)
$jsonPath = [IO.Path]::ChangeExtension($reportPath, '.json')
$csvPath = [IO.Path]::ChangeExtension($reportPath, '.csv')

$restorePointResults = @()
$restorePointFrequencyBackup = $null
if (-not $SkipRestorePoints) {
    $restorePointFrequencyBackup = Invoke-Safely -Name 'Restore point frequency setting' -DefaultValue $null -ScriptBlock { Set-RestorePointFrequencyForRun }
    $restorePointResults += Invoke-Safely -Name 'Restore point BEFORE' -DefaultValue ([PSCustomObject]@{Stage='BEFORE';Name=('PC_Cleanup{0}-BEFORE' -f (Get-Date -Format 'ddMMyy'));Status='Failed';Time=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss');Details='Step failed before returning results.'}) -ScriptBlock { New-MaintenanceRestorePoint -Stage 'BEFORE' }
} else { Add-ActionLog -Step 'Restore points' -Status 'Skipped' -Details 'SkipRestorePoints was supplied.' }

$beforeDisk = Invoke-Safely -Name 'Disk check before' -DefaultValue $null -ScriptBlock { Get-DiskSpace -Letter $DriveLetter }
$cleanupSnapshotBefore = Invoke-Safely -Name 'Cleanup category snapshot before' -DefaultValue @() -ScriptBlock { Get-CleanupCategorySnapshot -Stage 'Before' }
$pendingRebootBefore = Invoke-Safely -Name 'Pending reboot check before' -DefaultValue ([PSCustomObject]@{RebootRequired=$false;Reasons='';ReasonCount=0}) -ScriptBlock { Get-PendingRebootStatus }
$notableEvents = Invoke-Safely -Name 'Event log review' -DefaultValue @() -ScriptBlock { Get-NotableEvents -LookbackDays $EventLookbackDays }
$eventSummary = Invoke-Safely -Name 'Event summary' -DefaultValue @() -ScriptBlock { Get-EventSummary -Events $notableEvents }
$eventInterpretation = Invoke-Safely -Name 'Reliability interpretation' -DefaultValue @() -ScriptBlock { Get-EventInterpretation -Events $notableEvents }
$tempCleanup = Invoke-Safely -Name 'Temp and cache cleanup' -DefaultValue @() -ScriptBlock { Invoke-TempCleanup -MinAgeDays $TempFileMinAgeDays }
$recycleBinResult = Invoke-Safely -Name 'Recycle Bin cleanup' -DefaultValue 'Failed' -ScriptBlock { Invoke-RecycleBinCleanup }
$script:ComponentStoreAnalysis = Invoke-Safely -Name 'Component store analysis' -DefaultValue ([PSCustomObject]@{Name='DISM AnalyzeComponentStore';Command='';Status='Failed';ExitCode=$null;Output='';Errors='Step failed before returning results.';Summary='Failed'}) -ScriptBlock { Invoke-ComponentStoreAnalysis }
$componentCleanupResult = Invoke-Safely -Name 'Component store cleanup' -DefaultValue 'Failed' -ScriptBlock { Invoke-ComponentCleanup }
$deliveryOptimizationResult = Invoke-Safely -Name 'Delivery Optimization cleanup' -DefaultValue 'Failed' -ScriptBlock { Invoke-DeliveryOptimizationCleanup }
$windowsOldCleanupResult = Invoke-Safely -Name 'Windows.old cleanup' -DefaultValue ([PSCustomObject]@{Requested=$CleanWindowsOld;Path='';BeforeGB=0;AfterGB=0;RemovedGB=0;Status='Failed';Details='Step failed before returning results.'}) -ScriptBlock { Invoke-WindowsOldCleanup }
$windowsOldInfo = Invoke-Safely -Name 'Windows.old check' -DefaultValue ([PSCustomObject]@{Exists=$false;Path='';SizeGB=0;Details='Failed'}) -ScriptBlock { Get-WindowsOldInfo }
$chkdskResult = Invoke-Safely -Name 'CHKDSK smart check' -DefaultValue ([PSCustomObject]@{Name='CHKDSK smart';ScanCommand='';QueueCommand='';Status='Failed';ScanExitCode=$null;QueueExitCode=$null;ScanOutput='';QueueOutput='';Errors='Step failed before returning results.';Summary='Step failed before returning results.'}) -ScriptBlock { Invoke-ChkdskSmart -Letter $DriveLetter }
$sfcResult = Invoke-Safely -Name 'SFC scan' -DefaultValue ([PSCustomObject]@{Name='SFC scan';Command='sfc.exe /scannow';Status='Failed';ExitCode=$null;Output='';Errors='Step failed before returning results.';Summary='Step failed before returning results.'}) -ScriptBlock { Invoke-SfcScan }
$dismRestoreHealthResult = Invoke-Safely -Name 'DISM RestoreHealth' -DefaultValue ([PSCustomObject]@{Name='DISM RestoreHealth';Command='dism.exe /online /cleanup-image /restorehealth /norestart';Status='Failed';ExitCode=$null;Output='';Errors='Step failed before returning results.';Summary='Step failed before returning results.'}) -ScriptBlock { Invoke-DismRestoreHealth }
$userProfileSizes = Invoke-Safely -Name 'Local user profile size audit' -DefaultValue @() -ScriptBlock { Get-LocalUserProfileSizes }
$removedProfiles = Invoke-Safely -Name 'Stale profile removal' -DefaultValue @() -ScriptBlock { Invoke-StaleProfileRemoval -Profiles $userProfileSizes }
$storageHealth = Invoke-Safely -Name 'Storage health' -DefaultValue @() -ScriptBlock { Get-StorageHealth }
$batteryHealth = Invoke-Safely -Name 'Battery health' -DefaultValue ([PSCustomObject]@{HasBattery=$false;Name='';Status='';EstimatedChargeRemaining=$null;EstimatedRunTime=$null;BatteryReportPath='';Details='Failed'}) -ScriptBlock { Get-BatteryHealth -OutputFolder $ReportRoot }
$problemDevices = Invoke-Safely -Name 'Device Manager problem check' -DefaultValue @() -ScriptBlock { Get-ProblemDevices }
$startupItems = Invoke-Safely -Name 'Startup item audit' -DefaultValue @() -ScriptBlock { Get-StartupItems }
$serviceIssues = Invoke-Safely -Name 'Service health check' -DefaultValue @() -ScriptBlock { Get-ServiceHealthIssues }
$defenderStatus = Invoke-Safely -Name 'Defender status' -DefaultValue ([PSCustomObject]@{Available=$false;Details='Failed'}) -ScriptBlock { Get-DefenderStatus }
$teamsOneDriveHealth = Invoke-Safely -Name 'Teams and OneDrive health' -DefaultValue @() -ScriptBlock { Get-TeamsOneDriveHealth }
$crashDumpCleanupResult = Invoke-Safely -Name 'Crash dump cleanup' -DefaultValue ([PSCustomObject]@{Requested=$CleanCrashDumps;RemovedCount=0;RemovedGB=0;Status='Failed';Details='Step failed before returning results.'}) -ScriptBlock { Invoke-CrashDumpCleanup }
$crashDumpSummary = Invoke-Safely -Name 'Crash dump summary' -DefaultValue ([PSCustomObject]@{DumpCount=0;TotalSizeGB=0;LatestDump='';LatestDumpTime=$null}) -ScriptBlock { Get-CrashDumpSummary }
$storageHotspots = Invoke-Safely -Name 'Storage hotspot scan' -DefaultValue @() -ScriptBlock { Get-StorageHotspots }
$staleTasks = Invoke-Safely -Name 'Scheduled task audit' -DefaultValue @() -ScriptBlock { Get-StaleScheduledTasks -Days $StaleTaskDays }
$removedTasks = Invoke-Safely -Name 'Scheduled task cleanup' -DefaultValue @() -ScriptBlock { Invoke-StaleTaskCleanup -StaleTasks $staleTasks }
$windowsUpdateResult = Invoke-Safely -Name 'Windows Update' -DefaultValue ([PSCustomObject]@{Status='Failed';Updates=@();RebootRequired=$false;Details='Step failed before returning results.'}) -ScriptBlock { Invoke-WindowsUpdateNoReboot }
$eventLogMaintenance = Invoke-Safely -Name 'Event log maintenance' -DefaultValue @() -ScriptBlock { Invoke-EventLogMaintenance -Clear:$ClearEventLogs -Archive:$ArchiveEventLogs }
$resourceSnapshot = Invoke-Safely -Name 'Resource snapshot' -DefaultValue ([PSCustomObject]@{CPUPercent=$null;MemoryPercent=$null;DiskQueueLength=$null;Issues=@('Resource snapshot failed')}) -ScriptBlock { Get-ResourceSnapshot -CpuThreshold $CpuWarnPercent -MemoryThreshold $MemoryWarnPercent -DiskQueueThreshold $DiskQueueWarn }
$topCpuProcesses = Invoke-Safely -Name 'Top CPU processes' -DefaultValue @() -ScriptBlock { Get-TopCpuProcesses -Count $TopProcessCount }
$topMemoryProcesses = Invoke-Safely -Name 'Top memory processes' -DefaultValue @() -ScriptBlock { Get-TopMemoryProcesses -Count $TopProcessCount }
$afterDisk = Invoke-Safely -Name 'Disk check after' -DefaultValue $beforeDisk -ScriptBlock { Get-DiskSpace -Letter $DriveLetter }
$cleanupSnapshotAfter = Invoke-Safely -Name 'Cleanup category snapshot after' -DefaultValue @() -ScriptBlock { Get-CleanupCategorySnapshot -Stage 'After' }
$cleanupCategoryMetrics = Invoke-Safely -Name 'Cleanup category metrics' -DefaultValue @() -ScriptBlock { Compare-CleanupCategorySnapshots -Before $cleanupSnapshotBefore -After $cleanupSnapshotAfter }
$pendingRebootAfter = Invoke-Safely -Name 'Pending reboot check after' -DefaultValue $pendingRebootBefore -ScriptBlock { Get-PendingRebootStatus }

if (-not $SkipRestorePoints -and $CreateAfterRestorePoint) {
    $restorePointResults += Invoke-Safely -Name 'Restore point AFTER' -DefaultValue ([PSCustomObject]@{Stage='AFTER';Name=('PC_Cleanup{0}-AFTER' -f (Get-Date -Format 'ddMMyy'));Status='Failed';Time=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss');Details='Step failed before returning results.'}) -ScriptBlock { New-MaintenanceRestorePoint -Stage 'AFTER' }
    Invoke-Safely -Name 'Restore point frequency reset' -DefaultValue $null -ScriptBlock { Restore-RestorePointFrequencySetting -Backup $restorePointFrequencyBackup } | Out-Null
} elseif (-not $SkipRestorePoints) {
    Add-ActionLog -Step 'Restore point AFTER' -Status 'Skipped' -Details 'v15 only creates AFTER restore point when -CreateAfterRestorePoint is supplied.'
    Invoke-Safely -Name 'Restore point frequency reset' -DefaultValue $null -ScriptBlock { Restore-RestorePointFrequencySetting -Backup $restorePointFrequencyBackup } | Out-Null
}

$script:VssAfter = Invoke-Safely -Name 'VSS shadow storage after' -DefaultValue @() -ScriptBlock { Get-ShadowStorageInfo }
$healthScore = Invoke-Safely -Name 'Health score' -DefaultValue ([PSCustomObject]@{Score=0;Status='Unknown';Reasons='Failed to calculate health score.'}) -ScriptBlock { Get-HealthScore -BeforeDisk $beforeDisk -AfterDisk $afterDisk -PendingReboot $pendingRebootAfter -Resource $resourceSnapshot -StorageHealth $storageHealth -Defender $defenderStatus -ProblemDevices $problemDevices -ServiceIssues $serviceIssues -Sfc $sfcResult -Dism $dismRestoreHealthResult -Chkdsk $chkdskResult -CrashDumps $crashDumpSummary -NotableEvents $notableEvents }
$technicianSummary = Invoke-Safely -Name 'Technician summary' -DefaultValue 'Summary generation failed.' -ScriptBlock { Get-TechnicianSummary -Health $healthScore -BeforeDisk $beforeDisk -AfterDisk $afterDisk -PendingReboot $pendingRebootAfter -Chkdsk $chkdskResult -Sfc $sfcResult -Dism $dismRestoreHealthResult -WindowsUpdate $windowsUpdateResult -RemovedTasks $removedTasks -Profiles $userProfileSizes -WindowsOld $windowsOldInfo }

$finalReport = Invoke-Safely -Name 'HTML report generation' -DefaultValue $null -ScriptBlock {
    New-MaintenanceReport -BeforeDisk $beforeDisk -AfterDisk $afterDisk -TempCleanup $tempCleanup -RecycleBinResult $recycleBinResult -ComponentCleanupResult $componentCleanupResult -DeliveryOptimizationResult $deliveryOptimizationResult -NotableEvents $notableEvents -EventSummary $eventSummary -EventInterpretation $eventInterpretation -StaleTasks $staleTasks -RemovedTasks $removedTasks -RemovedProfiles $removedProfiles -EventLogMaintenance $eventLogMaintenance -ResourceSnapshot $resourceSnapshot -TopCpuProcesses $topCpuProcesses -TopMemoryProcesses $topMemoryProcesses -WindowsUpdateResult $windowsUpdateResult -RestorePointResults $restorePointResults -ChkdskResult $chkdskResult -SfcResult $sfcResult -DismRestoreHealthResult $dismRestoreHealthResult -UserProfileSizes $userProfileSizes -PendingReboot $pendingRebootAfter -StorageHealth $storageHealth -BatteryHealth $batteryHealth -ProblemDevices $problemDevices -StartupItems $startupItems -ServiceIssues $serviceIssues -DefenderStatus $defenderStatus -CrashDumpSummary $crashDumpSummary -CrashDumpCleanupResult $crashDumpCleanupResult -StorageHotspots $storageHotspots -WindowsOldInfo $windowsOldInfo -WindowsOldCleanupResult $windowsOldCleanupResult -TeamsOneDriveHealth $teamsOneDriveHealth -CleanupCategoryMetrics $cleanupCategoryMetrics -HealthScore $healthScore -TechnicianSummary $technicianSummary -OutputPath $reportPath
}

$summaryObject = [PSCustomObject]@{
    ComputerName = $env:COMPUTERNAME
    UserName = $env:USERNAME
    Mode = if($script:IsAuditOnly){'Audit'}else{'Remediate'}
    StartTime = $script:RunStart
    EndTime = Get-Date
    BaseReportRoot = $BaseReportRoot
    RunFolder = $ReportRoot
    RunFolderName = $RunFolderInfo.FolderName
    ReportPath = $finalReport
    JsonPath = $jsonPath
    CsvPath = $csvPath
    TranscriptPath = $script:TranscriptPath
    HealthScore = $healthScore.Score
    HealthStatus = $healthScore.Status
    HealthReasons = $healthScore.Reasons
    DiskFreeBeforeGB = if($beforeDisk){$beforeDisk.FreeGB}else{$null}
    DiskFreeAfterGB = if($afterDisk){$afterDisk.FreeGB}else{$null}
    DiskFreeAfterPercent = if($afterDisk){$afterDisk.FreePercent}else{$null}
    PendingReboot = $pendingRebootAfter.RebootRequired
    PendingRebootReasons = $pendingRebootAfter.Reasons
    ChkdskStatus = $chkdskResult.Status
    SfcStatus = $sfcResult.Status
    SfcSummary = $sfcResult.Summary
    DismStatus = $dismRestoreHealthResult.Status
    DismSummary = $dismRestoreHealthResult.Summary
    WindowsUpdateStatus = $windowsUpdateResult.Status
    WindowsUpdateDetails = $windowsUpdateResult.Details
    ProblemDeviceCount = @($problemDevices).Count
    StaleProfileCandidateCount = @($userProfileSizes | Where-Object { $_.StaleProfileCandidate -eq $true }).Count
    RemovedTaskCount = @($removedTasks).Count
    RemovedProfileCount = @($removedProfiles).Count
    CleanedCrashDumpGB = $crashDumpCleanupResult.RemovedGB
    CleanedWindowsOldGB = $windowsOldCleanupResult.RemovedGB
    TeamsOneDriveRecordCount = @($teamsOneDriveHealth).Count
    VssBefore = $script:VssBefore
    VssAfter = $script:VssAfter
    StepMetrics = $script:StepMetrics
    PreflightChecks = $script:PreflightChecks
    WindowsUpdateDiagnostics = $script:WindowsUpdateDiagnostics
    BitLockerHealth = $script:BitLockerHealth
    PowerHealth = $script:PowerHealth
    ComponentStoreAnalysis = $script:ComponentStoreAnalysis
    ReportRetention = $reportRetentionResult
    ShadowStorageLimitResult = $shadowLimitResult
    ActionLog = $script:ActionLog
}

try {
    $summaryObject | ConvertTo-Json -Depth 6 | Out-File -LiteralPath $jsonPath -Encoding UTF8 -Force
    $summaryObject | Select-Object ComputerName,UserName,Mode,StartTime,EndTime,HealthScore,HealthStatus,DiskFreeBeforeGB,DiskFreeAfterGB,DiskFreeAfterPercent,PendingReboot,PendingRebootReasons,ChkdskStatus,SfcStatus,DismStatus,WindowsUpdateStatus,ProblemDeviceCount,StaleProfileCandidateCount,RemovedTaskCount,RemovedProfileCount,CleanedCrashDumpGB,CleanedWindowsOldGB,TeamsOneDriveRecordCount,ReportPath | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Force
    Add-ActionLog -Step 'Machine-readable output' -Status 'Complete' -Details ("JSON: {0}; CSV: {1}" -f $jsonPath, $csvPath)
} catch { Add-ActionLog -Step 'Machine-readable output' -Status 'Failed' -Details $_.Exception.Message }

$notificationResult = Invoke-Safely -Name 'Report notifications' -DefaultValue $null -ScriptBlock { Send-MaintenanceNotifications -Summary $summaryObject -HtmlReportPath $finalReport -JsonReportPath $jsonPath -CsvReportPath $csvPath }

if (-not $SkipTranscript -and $script:TranscriptPath) { try { Stop-Transcript | Out-Null } catch { } }
$zipPath = Compress-RunFolder -FolderPath $ReportRoot

if ($finalReport) {
    Write-Host "Maintenance complete." -ForegroundColor Green
    Write-Host "HTML report: $finalReport" -ForegroundColor Green
    Write-Host "JSON summary: $jsonPath" -ForegroundColor Green
    Write-Host "CSV summary: $csvPath" -ForegroundColor Green
    if ($script:TranscriptPath) { Write-Host "Transcript: $script:TranscriptPath" -ForegroundColor Green }
    if ($zipPath) { Write-Host "ZIP package: $zipPath" -ForegroundColor Green }
} else {
    Write-Warning 'Maintenance ran, but report generation failed.'
}

$exitCode = Get-RmmExitCode -Health $healthScore -PendingReboot $pendingRebootAfter -ReportFailed:($null -eq $finalReport)
Write-Host "Suggested RMM exit code: $exitCode" -ForegroundColor Yellow
if ($UseRmmExitCode) { exit $exitCode }
