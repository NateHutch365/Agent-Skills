<#
.SYNOPSIS
    Builds a filterable Excel workbook from an IRM export folder created by Export-IrmConfiguration.ps1.
.DESCRIPTION
    Works offline from IrmPolicies.raw.json and IrmEntityLists.raw.json (plus SensitiveInfoTypes.json, Labels.json,
    IrmRoleGroups.json and ExportInfo.json when present). Uses the ImportExcel module if installed, otherwise Python with openpyxl.
.EXAMPLE
    .\Build-IrmWorkbook.ps1 -ExportFolder .\IRM-Export-Contoso-20261006-1516
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ExportFolder,
    [string]$ClientName,
    [string]$OutputPath,
    [ValidateSet('Auto', 'ImportExcel', 'Python')][string]$Engine = 'Auto'
)

$ErrorActionPreference = 'Stop'
$ExportFolder = (Resolve-Path $ExportFolder).Path
$guidPattern = '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'

#region Lookup tables
$templateText = @{
    TenantSetting               = 'Global settings'
    IntellectualPropertyTheft   = 'Data theft by departing users'
    LeakOfInformation           = 'Data leaks'
    DisgruntledEmployeeDataLeak = 'Data leaks by risky users'
    HighValueEmployeeDataLeak   = 'Data leaks by priority users'
    SecurityPolicyViolation     = 'Security policy violations'
    DepartingEmployeeSPV        = 'Security policy violations by departing users'
    DisgruntledEmployeeSPV      = 'Security policy violations by risky users'
    HighValueEmployeeSPV        = 'Security policy violations by priority users'
    HealthcareDataThreat        = 'Patient data misuse'
    UnacceptableUsage           = 'Risky browser usage'
    RiskyAIUsage                = 'Risky AI usage'
    RiskyAgents                 = 'Risky agents'
    SessionRecording            = 'Forensic evidence'
}
$departingTemplates = 'IntellectualPropertyTheft', 'DepartingEmployeeSPV'

$triggerText = @{
    AggregatedDataLeakTrigger                  = 'Data leak activities'
    AggregatedRiskyAIUsageTrigger              = 'Risky AI usage activities'
    AggregatedUnacceptableUsageBrowsingTrigger = 'Risky browsing activities'
}

$indicatorTypeText = @{ Insight = 'Indicator'; Booster = 'Risk score booster'; ExtensibleInsight = 'Extended indicator' }

# Section, setting label; keys are TenantSettings paths. Unmapped keys are still shown with a generated label.
$globalSettingText = [ordered]@{
    'Region'                                            = 'General', 'Data region'
    'IndicatorVersion'                                  = 'General', 'Indicator schema version'
    'TimeSpan.InScopeTimeSpan'                          = 'Policy timeframes', 'Activation window (days a user stays in scope after triggering)'
    'TimeSpan.HistoricTimeSpan'                         = 'Policy timeframes', 'Past activity detection (days before triggering)'
    'TimeSpan.FutureTerminationWindow'                  = 'Policy timeframes', 'Future termination window (days)'
    'TimeSpan.PastTerminationWindow'                    = 'Policy timeframes', 'Past termination window (days)'
    'TimeSpan.PostTerminationActivity'                  = 'Policy timeframes', 'Detect post-termination activity'
    'IntelligentDetections.AlertVolume'                 = 'Intelligent detections', 'Alert volume'
    'IntelligentDetections.FileVolCutoffLimits'         = 'Intelligent detections', 'Minimum daily events to boost score for unusual activity'
    'IntelligentDetections.OLThresHolds'                = 'Intelligent detections', 'Unusual activity threshold (OLThresholds)'
    'IntelligentDetections.FileExclusion'               = 'Intelligent detections', 'File type exclusions'
    'IntelligentDetections.MDATPTriageStatus'           = 'Intelligent detections', 'Defender for Endpoint alert statuses imported'
    'InterpretedSettings.EmailSignatureExclusionSettings' = 'Intelligent detections', 'Ignore email signature attachments'
    'FeatureSettings.Anonymization'                     = 'Privacy', 'Show pseudonymised user names'
    'FeatureSettings.DLPUserRiskSync'                   = 'Data sharing', 'Share user risk details with other solutions (DLP, Communication Compliance, Defender)'
    'FeatureSettings.OptInIRMDataExport'                = 'Data sharing', 'IRM data export opt-in (OptInIRMDataExport)'
    'FeatureSettings.RaiseAuditAlert'                   = 'Export alerts', 'Export alerts via Office 365 Management Activity API (SIEM)'
    'FeatureSettings.InlineAlertPolicyCustomization'    = 'Inline alert customisation', 'Allow alert reviewers to tune policies from alerts'
    'FeatureSettings.EnableTeam'                        = 'Microsoft Teams', 'Teams integration for cases'
    'FeatureSettings.DisableNotification'               = 'Admin notifications', 'Notifications disabled'
    'FeatureSettings.NotificationDetails'               = 'Admin notifications', 'Email when the first alert is generated for a new policy'
    'InterpretedSettings.IRASettings'                   = 'Analytics', 'Insider risk analytics (tenant-level insights)'
    'InterpretedSettings.UserAnalyticsSettings'         = 'Analytics', 'User-level analytics insights'
    'DynamicRiskPreventionSettings'                     = 'Adaptive Protection', 'Adaptive Protection'
    'AgentAdaptiveProtectionSettings'                   = 'Adaptive Protection', 'Adaptive Protection for agents'
}

$notificationText = @{
    AnalyticsNewInsight = 'Monthly email summarising new analytics insights'
    AnalyticsTurnedOff  = 'Email when analytics is turned off'
    HighSeverityAlerts  = 'Daily email when new high severity alerts are generated'
    PoliciesHealth      = 'Weekly email summarising policies with unresolved warnings'
}

$systemListText = @{
    IrmWhitelistDomains         = 'Intelligent detections - Allowed domains'
    IrmBlacklistDomains         = 'Intelligent detections - Unallowed domains'
    IrmEnterpriseDomains        = 'Intelligent detections - Enterprise domains'
    IrmPublicDomains            = 'Intelligent detections - Public domains'
    IrmPhysicalBadgeAssets      = 'Priority physical assets'
    IrmCustomExWinFilePaths     = 'Global exclusions - File paths'
    IrmDsbldSysExWinFilePaths   = 'Global exclusions - Default file path exclusions turned off'
    IrmCustomExMLClassifiers    = 'Global exclusions - Trainable classifiers'
    IrmDsbldSysExMLClassifiers  = 'Global exclusions - Default classifier exclusions turned off'
    IrmCustomExSensitiveTypes   = 'Global exclusions - Sensitive info types'
    IrmDsbldSysExSensitiveTypes = 'Global exclusions - Default SIT exclusions turned off'
    IrmExcludedSites            = 'Global exclusions - SharePoint sites'
    IrmExcludedKeywords         = 'Global exclusions - Keywords to exclude'
    IrmNotExcludedKeywords      = 'Global exclusions - Keywords to always include'
    IrmXSGSensitiveInfoTypes    = 'Global exclusions - Sensitive info type groups'
    IrmXSGDomains               = 'Global exclusions - Domain groups'
    IrmXSGSites                 = 'Global exclusions - Site groups'
    IrmXSGExcludedKeywords      = 'Global exclusions - Excluded keyword groups'
    IrmXSGExceptionKeywords     = 'Global exclusions - Exception keyword groups'
    IrmXSGFiletypes             = 'Global exclusions - File type groups'
    IrmXSGFilePaths             = 'Global exclusions - File path groups'
    IrmXSGMLClassifierTypes     = 'Global exclusions - Trainable classifier groups'
    IrmXSGApplications          = 'Global exclusions - Application groups'
    IrmXSGPrinters              = 'Global exclusions - Printer groups'
    IrmDlpIndicatorPolicies     = 'DLP policies used by the high severity DLP indicator'
}

$commonNoise = 'Identity', 'Id', 'Guid', 'ImmutableId', 'ExchangeObjectId', 'ObjectVersion', 'DirectoryObjectVersion', 'IsValid',
    'ExchangeVersion', 'DistinguishedName', 'ObjectCategory', 'ObjectClass', 'WhenChanged', 'WhenCreated', 'OrganizationalUnitRoot',
    'OrganizationId', 'OriginatingServer', 'ObjectState', 'ReadOnly', 'ExternalIdentity', 'ErrorMetadata', 'Workload',
    'RunspaceId', 'PSComputerName', 'PSShowComputerName'

$policyNoise = $commonNoise + @('Type', 'IndicatorVersion', 'IsCustom', 'SchemaVersion', 'PolicyCategory', 'UserAdministrativeUnitMembershipMap',
    'ForceValidate', 'PolicyRulesMetaData', 'GlobalListType', 'Locations', 'PolicyConstraints', 'DistributionStatus', 'DistributionSyncStatus',
    'DistributionResults', 'LastStatusUpdateTime', 'ModificationTimeUtc', 'CreationTimeUtc', 'PolicyRBACScopes', 'TenantSettings', 'TenantSetting')

$policyHandled = 'Name', 'InsiderRiskScenario', 'Mode', 'Enabled', 'Priority', 'Comment', 'Triggers', 'TriggerInsightGroups',
    'ExtensibleTriggerInsightGroups', 'DlpPolicy', 'DlpPolicyGuid', 'DlpPoliciesAsTrigger', 'CCPolicyName', 'CCPolicySdsId', 'Indicators',
    'ExtensibleIndicators', 'IsPriorityContentOnlyScoring', 'AdaptiveScopeLocation', 'ExchangeLocation', 'ExchangeLocationException',
    'ModernGroupLocation', 'ModernGroupLocationException', 'SensitivityLabels', 'SharepointSites', 'TeamsSites', 'DlpSensitiveTypes',
    'MLClassifierTypes', 'FileExtensions', 'CustomTags', 'InScopeTimeSpan', 'HistoricTimeSpan', 'FutureTerminationWindow',
    'PastTerminationWindow', 'PostTerminationActivity', 'PolicyHealth', 'CreatedBy', 'LastModifiedBy', 'WhenCreatedUTC', 'WhenChangedUTC'
#endregion

#region Value helpers
function Read-ExportJson([string]$Name) {
    $path = Join-Path $ExportFolder $Name
    if (-not (Test-Path $path)) { return }
    # Raw exports serialise enums as {"value":1,"Value":"Text"}; the case-duplicate key breaks ConvertFrom-Json.
    $text = (Get-Content $path -Raw -Encoding UTF8) -creplace '"value":\s*-?\d+,(\s*)"Value"', '$1"Value"'
    $text | ConvertFrom-Json
}

function Test-Blank($v) {
    if ($null -eq $v) { return $true }
    if ($v -is [string]) { return ([string]::IsNullOrWhiteSpace($v) -or $v -in 'None', 'NotSet', '[]', '{}') }
    if ($v -is [bool]) { return -not $v }
    if ($v -is [ValueType]) { return $false }
    if ($v -is [System.Collections.IEnumerable]) {
        foreach ($i in $v) { if (-not (Test-Blank $i)) { return $false } }
        return $true
    }
    return (@($v.PSObject.Properties).Count -eq 0)
}

function Format-Date($v) {
    if (Test-Blank $v) { return '' }
    $d = if ($v -is [datetime]) { $v } else {
        [datetime]::Parse([string]$v, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind)
    }
    if ($d.Kind -eq [DateTimeKind]::Local) { $d = $d.ToUniversalTime() }
    $d.ToString('yyyy-MM-dd HH:mm')
}

function Format-Val($v) {
    if (Test-Blank $v) { return '' }
    if ($v -is [bool]) { return 'Yes' }
    if ($v -is [datetime]) { return Format-Date $v }
    if ($v -is [string] -or $v -is [ValueType]) { return [string]$v }
    if ($v -is [System.Collections.IEnumerable]) {
        return (@(foreach ($i in $v) { if (-not (Test-Blank $i)) { Format-Val $i } }) -join '; ')
    }
    $props = @($v.PSObject.Properties | Where-Object { -not (Test-Blank $_.Value) })
    if ($props.Count -eq 1 -and $props[0].Name -eq 'Value') { return Format-Val $props[0].Value }
    (@(foreach ($p in $props) { "$($p.Name): $(Format-Val $p.Value)" }) -join ', ')
}

function Get-FriendlyName([string]$Name) {
    if ($Name -notmatch '^[A-Za-z0-9 ,]+$') { return $Name }
    $Name -creplace '(?<=[a-z0-9])(?=[A-Z])', ' ' -creplace '(?<=[A-Z])(?=[A-Z][a-z])', ' '
}

function ConvertFrom-JsonString($v) {
    if ($v -isnot [string]) { return $v }
    $t = $v.TrimStart()
    if (-not ($t.StartsWith('[') -or $t.StartsWith('{'))) { return $v }
    try { $parsed = $v | ConvertFrom-Json; return , $parsed } catch { return $v }
}

# IRM stores most collections as arrays of JSON strings.
function Get-Parsed($v) {
    @(foreach ($i in @($v)) { if (-not (Test-Blank $i)) { ConvertFrom-JsonString $i } })
}

function ConvertTo-CellValue($v) {
    if ($null -eq $v) { return '' }
    if ($v -is [int] -or $v -is [long] -or $v -is [double]) { return $v }
    $s = [string]$v
    if ($s.Length -gt 32000) { $s = $s.Substring(0, 32000) + ' ...[truncated - see raw JSON]' }
    # Prevent tenant-sourced text being interpreted as a formula.
    if ($s -match '^[=+@-]') { $s = ' ' + $s }
    $s
}

function Format-OnOff($v) {
    $s = [string]$v
    if ($s -eq 'True') { return 'On' }
    if ($s -eq 'False') { return 'Off' }
    $s
}
#endregion

#region IRM formatting
$keepCase = 'SharePoint', 'OneDrive', 'Teams', 'Copilot', 'Power', 'BI', 'Box', 'Dropbox', 'Google', 'Drive', 'Azure', 'Fabric', 'Defender',
    'Microsoft', 'Purview', 'Entra', 'Bluetooth'
$tokenFix = @{ Llm = 'LLM'; Ai = 'AI'; Sql = 'SQL'; Dlp = 'DLP'; Usb = 'USB'; Ml = 'ML'; Sit = 'SIT'; Vm = 'VM'; Url = 'URL'; Odb = 'OneDrive' }
$prefixFix = @{ Spo = 'SharePoint'; Epo = 'Device'; Mcas = 'Cloud app'; CC = 'Communication compliance' }

function Get-IndicatorLabel([string]$Name) {
    if ($Name -match $guidPattern) { return "Custom indicator ($Name)" }
    $n = $Name -creplace '(?<=[a-z0-9])(?=[A-Z])', ' ' -creplace '(?<=[A-Z])(?=[A-Z][a-z])', ' ' -creplace '(?<=[a-z])(?=[0-9])', ' '
    $words = @($n -split '\s+' | Where-Object { $_ })
    if ($words.Count -and $prefixFix.ContainsKey($words[0])) { $words[0] = $prefixFix[$words[0]] }
    for ($i = 0; $i -lt $words.Count; $i++) {
        $w = $words[$i]
        if ($tokenFix.ContainsKey($w)) { $words[$i] = $tokenFix[$w]; continue }
        if ($i -gt 0 -and $w -cmatch '^[A-Z][a-z]+$' -and $w -notin $keepCase) { $words[$i] = $w.ToLower() }
    }
    $words -join ' '
}

function Get-IndicatorCategory([string]$Name, [string]$Type) {
    if ($Type -eq 'Booster' -or $Name -match '^(AnomalyDetections|PotentialHighImpactUser|PriorityUserGroupMember|UserHistory)$') { return 'Risk score boosters' }
    switch -Regex ($Name) {
        $guidPattern { return 'Custom indicators' }
        '^EpoBrowse' { return 'Risky browsing' }
        '^(Spo|SPO)' { return 'SharePoint' }
        '^Odb' { return 'OneDrive' }
        '^Teams' { return 'Teams' }
        '^Email' { return 'Email' }
        'Sequence$' { return 'Sequences' }
        'Cumulative' { return 'Cumulative exfiltration' }
        '^(Epo|CopyTo|Print|FilePrint)' { return 'Device' }
        '^Mcas' { return 'Defender for Cloud Apps' }
        '^SecurityAlert' { return 'Defender for Endpoint' }
        '^(Copilot|AIApp|ConnectedAIApp|CCPrompt|CCProtected)' { return 'AI apps' }
        '^CC' { return 'Communication Compliance' }
        '^PowerBI' { return 'Power BI' }
        '^(Lakehouse|Fabric)' { return 'Microsoft Fabric' }
        '^Azure' { return 'Azure' }
        '^AWS' { return 'AWS' }
        '^(Box|Dropbox|GoogleDrive)' { return 'Third-party cloud storage' }
        '^Network' { return 'Network' }
        '^Physical' { return 'Physical access' }
        '^Compromised' { return 'Risky sign-ins / compromised accounts' }
        '^HighSeverityDlp' { return 'DLP' }
        'Patient|EmployeeAccessed' { return 'Health record access' }
        default { return 'Other' }
    }
}

function Get-SubLabel([string]$Parent, [string]$Sub) {
    if ($Sub -match 'Blacklist|Unallowed') { return 'Unallowed domains' }
    if ($Sub -match 'NonPriority') { return 'Non-priority content' }
    if ($Sub -match 'Priority') { return 'Priority content' }
    if ($Sub -match 'Suspicious') { return 'Suspicious attachment' }
    $rest = if ($Sub.StartsWith($Parent, [StringComparison]::OrdinalIgnoreCase)) { $Sub.Substring($Parent.Length) } else { $Sub }
    if (-not $rest) { return 'All' }
    $l = Get-IndicatorLabel $rest
    $l.Substring(0, 1).ToUpper() + $l.Substring(1)
}

function Format-Thresholds($Indicator) {
    @(foreach ($s in @($Indicator.Indicators)) {
            $vals = @()
            if ($s.Thresholds) {
                $vals = @(foreach ($k in 'LowRisk', 'MediumRisk', 'HighRisk') {
                        if ($s.Thresholds.PSObject.Properties[$k]) { '{0} {1}' -f ($k -replace 'Risk$'), $s.Thresholds.$k }
                    })
            }
            elseif ($s.Threshold -and [double]$s.Threshold -ne 0) { $vals = @("Threshold $($s.Threshold)") }
            if ($vals.Count) { "$(Get-SubLabel ([string]$Indicator.Name) ([string]$s.Name)): $($vals -join ', ')" }
        }) -join "`n"
}

function Format-ThresholdMode([string]$Mode) {
    switch ($Mode) { 'Default' { 'Default' } 'Custom' { 'Custom' } '' { '' } default { Get-FriendlyName $Mode } }
}

function Format-Location($Item) {
    if ($Item -is [string]) {
        $parsed = ConvertFrom-JsonString $Item
        if ($parsed -is [string]) { return $Item }
        $Item = $parsed
    }
    if ([string]$Item.Type.Value -eq 'Tenant' -or [string]$Item.Type -eq 'Tenant') { return 'All users' }
    foreach ($n in 'DisplayName', 'PrimarySmtpAddress', 'Name', 'Identity') {
        $val = [string]$Item.$n
        if ($val -and $val -notmatch $guidPattern) { return $val }
    }
    [string]$Item.Name
}

function Get-Names($Items, [hashtable]$Map) {
    @(foreach ($i in (Get-Parsed $Items)) {
            if ($i -is [string]) { $i; continue }
            $g = [string]$i.Guid
            if ($Map -and $g -and $Map.ContainsKey($g)) { $Map[$g] } elseif ($i.Name) { [string]$i.Name } else { Format-Val $i }
        })
}

function Get-PolicyStatus($Policy) {
    if ($Policy.Enabled -eq $false) { return 'Off' }
    switch ([string]$Policy.Mode) { 'Enable' { 'On' } 'Disable' { 'Off' } default { Get-FriendlyName ([string]$Policy.Mode) } }
}

function Format-Health($Health) {
    $h = ConvertFrom-JsonString $Health
    if ($h -is [string] -or $null -eq $h) { return [string]$Health }
    $status = switch ([string]$h.HealthStatus) { 'RecommendReview' { 'Recommend review' } default { Get-FriendlyName ([string]$h.HealthStatus) } }
    $details = @(foreach ($d in @($h.ValidationDetails)) {
            if (-not $d -or [string]$d.ValidationSeverity -eq 'Healthy') { continue }
            "$(Get-FriendlyName ([string]$d.Name)): $(Get-FriendlyName ([string]$d.Status))"
        })
    if ($details.Count) { "$status`n$($details -join "`n")" } else { $status }
}

function Format-TriggerName([string]$Name) {
    if ($triggerText.ContainsKey($Name)) { return $triggerText[$Name] }
    Get-FriendlyName ($Name -replace '^Aggregated' -replace 'Trigger$')
}

function Format-Recipients($v) {
    $parts = @()
    $rg = @(@($v.Rolegroups) + @($v.RoleGroups) | Where-Object { $_ } | Select-Object -Unique)
    $rc = @(@($v.Recepients) + @($v.Recipients) | Where-Object { $_ } | Select-Object -Unique)
    if ($rg.Count) { $parts += "Role groups: $($rg -join ', ')" }
    if ($rc.Count) { $parts += "Recipients: $($rc -join ', ')" }
    if ($parts.Count) { $parts -join '; ' } else { 'No recipients' }
}

function Format-SettingValue($v) {
    $v = ConvertFrom-JsonString $v
    if ($null -eq $v) { return 'Not configured' }
    if ($v -is [string] -or $v -is [ValueType]) { return Format-OnOff $v }
    if ($v -is [System.Collections.IEnumerable]) {
        $text = @($v | ForEach-Object { Format-Val $_ } | Where-Object { $_ }) -join ', '
        return $(if ($text) { $text } else { 'None' })
    }
    $props = $v.PSObject.Properties
    if ($props['Enabled']) {
        $text = Format-OnOff $v.Enabled
        $extra = @()
        if ($props['EnabledBy'] -and $v.EnabledBy) { $extra += "enabled $($v.EnabledBy)" }
        if ($props['ModificationTime'] -and $v.ModificationTime) { $extra += "changed $(Format-Date $v.ModificationTime)" }
        if ($extra.Count) { $text += " ($($extra -join ', '))" }
        return $text
    }
    if ($props['Rolegroups'] -or $props['RoleGroups'] -or $props['Recepients'] -or $props['Recipients']) { return Format-Recipients $v }
    $text = Format-Val $v
    if ($text) { $text } else { 'None' }
}
#endregion

#region Workbook writers
function New-Sheet([string]$Name, $Columns, [int]$FreezeColumns, $Highlights) {
    @{
        Name          = $Name
        Kind          = 'Table'
        TableName     = 'tbl' + ($Name -replace '[^A-Za-z0-9]', '')
        Columns       = [string[]]@($Columns.Keys)
        Widths        = [int[]]@($Columns.Values)
        Rows          = [System.Collections.Generic.List[object]]::new()
        FreezeColumns = $FreezeColumns
        Highlights    = $(if ($Highlights) { $Highlights } else { @{} })
    }
}

function Add-Row($Sheet, $Values) {
    $Sheet.Rows.Add([object[]]@(foreach ($c in $Sheet.Columns) { ConvertTo-CellValue $Values[$c] }))
}

function ConvertTo-Colour([string]$Hex) {
    $h = $Hex.TrimStart('#')
    [System.Drawing.Color]::FromArgb(255, [Convert]::ToInt32($h.Substring(0, 2), 16), [Convert]::ToInt32($h.Substring(2, 2), 16), [Convert]::ToInt32($h.Substring(4, 2), 16))
}

function Write-WorkbookImportExcel($Spec) {
    Import-Module ImportExcel
    if (Test-Path $Spec.Path) { Remove-Item $Spec.Path -Force }
    $pkg = Open-ExcelPackage -Path $Spec.Path -Create
    foreach ($s in $Spec.Sheets) {
        $ws = Add-Worksheet -ExcelPackage $pkg -WorksheetName $s.Name
        if ($s.Kind -eq 'Cover') {
            $ws.Cells[1, 1].Value = $s.Title
            $ws.Cells[1, 1].Style.Font.Size = 16
            $ws.Cells[1, 1].Style.Font.Bold = $true
            $r = 3
            foreach ($row in $s.Rows) {
                $ws.Cells[$r, 1].Value = $row[0]
                $ws.Cells[$r, 1].Style.Font.Bold = $true
                $ws.Cells[$r, 2].Value = $row[1]
                $ws.Cells[$r, 2].Style.WrapText = $true
                $r++
            }
            $ws.Column(1).Width = 30
            $ws.Column(2).Width = 110
            continue
        }
        for ($c = 0; $c -lt $s.Columns.Count; $c++) { $ws.Cells[1, ($c + 1)].Value = $s.Columns[$c] }
        for ($r = 0; $r -lt $s.Rows.Count; $r++) {
            $row = $s.Rows[$r]
            for ($c = 0; $c -lt $row.Count; $c++) { $ws.Cells[($r + 2), ($c + 1)].Value = $row[$c] }
        }
        for ($c = 0; $c -lt $s.Widths.Count; $c++) { $ws.Column($c + 1).Width = $s.Widths[$c] }
        $range = $ws.Cells[1, 1, ($s.Rows.Count + 1), $s.Columns.Count]
        $range.Style.WrapText = $true
        $range.Style.VerticalAlignment = 'Top'
        if ($s.Rows.Count) {
            $table = $ws.Tables.Add($range, $s.TableName)
            $table.TableStyle = 'Medium2'
        }
        $ws.View.FreezePanes(2, $s.FreezeColumns + 1)
        foreach ($col in $s.Highlights.Keys) {
            $idx = [array]::IndexOf($s.Columns, $col)
            if ($idx -lt 0) { continue }
            for ($r = 0; $r -lt $s.Rows.Count; $r++) {
                $v = [string]$s.Rows[$r][$idx]
                foreach ($prefix in $s.Highlights[$col].Keys) {
                    if ($v.StartsWith($prefix)) {
                        $cell = $ws.Cells[($r + 2), ($idx + 1)]
                        $cell.Style.Fill.PatternType = 'Solid'
                        $cell.Style.Fill.BackgroundColor.SetColor((ConvertTo-Colour $s.Highlights[$col][$prefix]))
                        break
                    }
                }
            }
        }
    }
    Close-ExcelPackage $pkg
}

$pythonWriter = @'
import json, sys
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


def unwrap(x):
    # Windows PowerShell 5.1 can serialise arrays as {"value": [...], "Count": n}
    if isinstance(x, dict) and 'value' in x and 'Count' in x:
        return x['value']
    return x


with open(sys.argv[1], encoding='utf-8-sig') as f:
    spec = json.load(f)

wb = Workbook()
wb.remove(wb.active)
top = Alignment(wrap_text=True, vertical='top')

for s in unwrap(spec['Sheets']):
    ws = wb.create_sheet(s['Name'])
    rows = [unwrap(r) for r in unwrap(s['Rows'])]
    if s['Kind'] == 'Cover':
        ws['A1'] = s['Title']
        ws['A1'].font = Font(size=16, bold=True)
        for i, (k, v) in enumerate(rows, start=3):
            ws.cell(row=i, column=1, value=k).font = Font(bold=True)
            ws.cell(row=i, column=2, value=v).alignment = top
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 110
        continue
    cols = unwrap(s['Columns'])
    ws.append(cols)
    for r in rows:
        ws.append(r)
    for i, w in enumerate(unwrap(s['Widths']), start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=1, max_row=len(rows) + 1, max_col=len(cols)):
        for c in row:
            c.alignment = top
    if rows:
        t = Table(displayName=s['TableName'], ref='A1:%s%d' % (get_column_letter(len(cols)), len(rows) + 1))
        t.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
        ws.add_table(t)
    ws.freeze_panes = ws.cell(row=2, column=s['FreezeColumns'] + 1)
    for col, rules in (s.get('Highlights') or {}).items():
        if col not in cols:
            continue
        idx = cols.index(col) + 1
        for r in range(2, len(rows) + 2):
            v = str(ws.cell(row=r, column=idx).value or '')
            for prefix, colour in rules.items():
                if v.startswith(prefix):
                    hexc = colour.lstrip('#')
                    ws.cell(row=r, column=idx).fill = PatternFill('solid', start_color=hexc, end_color=hexc)
                    break

wb.save(spec['Path'])
'@

function Find-Python {
    $ErrorActionPreference = 'Continue'
    $candidates = @('py', '-3'), @('python'), @('python3')
    foreach ($c in $candidates) {
        if (-not (Get-Command $c[0] -CommandType Application -ErrorAction SilentlyContinue)) { continue }
        $pre = @($c | Select-Object -Skip 1)
        try { & $c[0] @pre -c 'import openpyxl' 2>$null | Out-Null } catch { continue }
        if ($LASTEXITCODE -eq 0) { return , $c }
        $script:pythonWithoutOpenpyxl = $c[0]
    }
}

function Write-WorkbookPython($Spec) {
    $python = Find-Python
    if (-not $python) {
        if ($script:pythonWithoutOpenpyxl) { throw "Python was found but openpyxl is missing. Run: $($script:pythonWithoutOpenpyxl) -m pip install openpyxl" }
        throw 'Neither the ImportExcel module nor Python with openpyxl is available. Run "Install-Module ImportExcel -Scope CurrentUser" or install Python and "pip install openpyxl".'
    }
    $tmp = Join-Path ([IO.Path]::GetTempPath()) "irm-workbook-$([guid]::NewGuid())"
    New-Item -ItemType Directory -Path $tmp | Out-Null
    try {
        $specPath = Join-Path $tmp 'spec.json'
        $scriptPath = Join-Path $tmp 'build.py'
        $Spec | ConvertTo-Json -Depth 10 -Compress | Set-Content $specPath -Encoding UTF8
        Set-Content $scriptPath $pythonWriter -Encoding UTF8
        $pre = @($python | Select-Object -Skip 1)
        & $python[0] @pre $scriptPath $specPath
        if ($LASTEXITCODE) { throw "Python workbook build failed (exit code $LASTEXITCODE)." }
    }
    finally {
        Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
    }
}
#endregion

#region Load export
$allPolicies = @(Read-ExportJson 'IrmPolicies.raw.json' | Where-Object { $_ })
if (-not $allPolicies.Count) { throw "IrmPolicies.raw.json not found or empty in $ExportFolder" }
$policies = @($allPolicies | Where-Object InsiderRiskScenario -ne 'TenantSetting' | Sort-Object { [int]$_.Priority }, Name)
$tenantPolicy = $allPolicies | Where-Object InsiderRiskScenario -eq 'TenantSetting' | Select-Object -First 1
$tenantSettings = if ($tenantPolicy) {
    ConvertFrom-JsonString $(if ($tenantPolicy.TenantSettings) { $tenantPolicy.TenantSettings } else { $tenantPolicy.TenantSetting })
}
if (-not $tenantSettings -or $tenantSettings -is [string]) { Write-Warning 'No TenantSetting object found - the Global Settings sheet will be empty.'; $tenantSettings = $null }

$lists = @(Read-ExportJson 'IrmEntityLists.raw.json' | Where-Object { $_ })
if (-not $lists.Count) { Write-Warning 'IrmEntityLists.raw.json not found or empty - no settings lists will be shown.' }
$missingEntries = @($lists | Where-Object { $_.EntitiesCount -gt 0 -and @($_.Entities | Where-Object { $_ }).Count -eq 0 })
if ($missingEntries.Count) {
    Write-Warning "$($missingEntries.Count) list(s) report entries but none were exported (listed by -Type only). Re-run Export-IrmConfiguration.ps1."
}

$exportInfo = Read-ExportJson 'ExportInfo.json'

$sitMap = @{}
$sits = @(Read-ExportJson 'SensitiveInfoTypes.json' | Where-Object { $_ })
foreach ($s in $sits) { $sitMap[[string]$s.Id] = $s }
if (-not $sits.Count) { Write-Warning 'SensitiveInfoTypes.json not found - SIT exclusions will show as GUIDs. Re-run Export-IrmConfiguration.ps1.' }

$labelMap = @{}
$labels = @(Read-ExportJson 'Labels.json' | Where-Object { $_ })
$labelById = @{}
foreach ($l in $labels) { foreach ($k in $l.Id, $l.Guid) { if ($k) { $labelById[[string]$k] = $l } } }
foreach ($l in $labels) {
    $display = [string]$l.DisplayName
    if ($l.ParentId -and $labelById.ContainsKey([string]$l.ParentId)) { $display = "$($labelById[[string]$l.ParentId].DisplayName) / $display" }
    foreach ($k in $l.Id, $l.Guid) { if ($k) { $labelMap[[string]$k] = $display } }
}
if (-not $labels.Count) { Write-Warning 'Labels.json not found - sensitivity labels will show internal names.' }

$roleGroups = @(Read-ExportJson 'IrmRoleGroups.json' | Where-Object { $_ })

$tenant = ([string]$allPolicies[0].OrganizationalUnitRoot -split '/')[-1]
if (-not $ClientName) {
    $ClientName = if ($exportInfo -and $exportInfo.ClientName) { [string]$exportInfo.ClientName } else { (Get-Culture).TextInfo.ToTitleCase(($tenant -split '\.')[0]) }
}
$snapshot = if ($exportInfo -and $exportInfo.ExportedOnUTC) { Format-Date $exportInfo.ExportedOnUTC }
            else { (Get-Item (Join-Path $ExportFolder 'IrmPolicies.raw.json')).LastWriteTimeUtc.ToString('yyyy-MM-dd HH:mm') }
if (-not $OutputPath) {
    $OutputPath = Join-Path $ExportFolder ('{0}-IRM-Configuration-{1}.xlsx' -f ($ClientName -replace '[^\w-]', ''), ($snapshot.Substring(0, 10) -replace '-', ''))
}
$OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
#endregion

#region Build sheets
$statusHighlight = [ordered]@{ 'On' = '#C6EFCE'; 'Off' = '#D9D9D9' }
$policyNames = @($policies | ForEach-Object { [string]$_.Name })

# Global indicator state from the tenant setting
$globalIndicators = [ordered]@{}
if ($tenantSettings) {
    foreach ($i in @($tenantSettings.Indicators) + @($tenantSettings.ExtensibleIndicators)) {
        if ($i -and $i.Name) { $globalIndicators[[string]$i.Name] = [bool]$i.Enabled }
    }
}

$policySheet = New-Sheet 'Policies' ([ordered]@{
        'Policy' = 42; 'Template' = 28; 'Status' = 9; 'Priority' = 8; 'Users In Scope' = 30; 'Excluded Users' = 22
        'Sensitivity Labels (priority content)' = 32; 'Sensitive Info Types (priority content)' = 40; 'Trainable Classifiers (priority content)' = 30
        'SharePoint Sites (priority content)' = 30; 'Teams (priority content)' = 20; 'File Extensions (priority content)' = 16
        'Score Only Priority Content' = 12; 'Triggers' = 26; 'Triggering Events' = 45; 'Indicators Selected' = 10; 'Timeframes' = 24
        'Policy Health' = 30; 'Description' = 60; 'Additional Settings' = 30; 'Created By' = 22; 'Created (UTC)' = 17
        'Last Modified By' = 22; 'Last Modified (UTC)' = 17
    }) 1 ([ordered]@{ 'Status' = $statusHighlight; 'Policy Health' = [ordered]@{ 'Healthy' = '#C6EFCE'; 'Recommend' = '#FFEB9C'; 'Unhealthy' = '#FFC7CE' } })

$indicatorSheet = New-Sheet 'Policy Indicators' ([ordered]@{
        'Policy' = 40; 'Category' = 22; 'Indicator' = 50; 'Indicator Type' = 16; 'Used In Risk Scoring' = 11
        'Threshold Mode' = 16; 'Thresholds (events per day)' = 50; 'Internal Name' = 40
    }) 3 ([ordered]@{ 'Used In Risk Scoring' = [ordered]@{ 'Yes' = '#C6EFCE' } })

$triggerSheet = New-Sheet 'Policy Triggers' ([ordered]@{
        'Policy' = 40; 'Trigger' = 28; 'Triggering Event' = 50; 'Custom Thresholds' = 35; 'Cumulative Window (days)' = 12; 'Internal Name' = 40
    }) 3 $null

$matrix = @{}
$priorityGroups = [System.Collections.Generic.List[string]]::new()
$statusList = @()

foreach ($p in $policies) {
    $name = [string]$p.Name
    $scenario = [string]$p.InsiderRiskScenario
    $statusList += Get-PolicyStatus $p

    # Users in scope
    $scope = @(foreach ($i in @($p.ExchangeLocation) + @($p.ModernGroupLocation)) { if (-not (Test-Blank $i)) { Format-Location $i } })
    $scope += @(foreach ($i in @($p.AdaptiveScopeLocation)) { if (-not (Test-Blank $i)) { "Adaptive scope: $(Format-Location $i)" } })
    $tags = Get-Names $p.CustomTags $null
    foreach ($t in $tags) { $scope += "Priority user group: $t"; if (-not $priorityGroups.Contains($t)) { $priorityGroups.Add($t) } }
    $excluded = @(foreach ($i in @($p.ExchangeLocationException) + @($p.ModernGroupLocationException)) { if (-not (Test-Blank $i)) { Format-Location $i } })

    # Triggers
    $triggerNames = [System.Collections.Generic.List[string]]::new()
    $events = [System.Collections.Generic.List[string]]::new()
    foreach ($g in @(Get-Parsed $p.TriggerInsightGroups) + @(Get-Parsed $p.ExtensibleTriggerInsightGroups)) {
        $tName = Format-TriggerName ([string]$g.Name)
        if (-not $triggerNames.Contains($tName)) { $triggerNames.Add($tName) }
        foreach ($ins in @($g.Insights)) {
            if (-not $ins -or -not $ins.Enabled) { continue }
            $label = Get-IndicatorLabel ([string]$ins.Name)
            $events.Add($label)
            Add-Row $triggerSheet ([ordered]@{
                    'Policy' = $name; 'Trigger' = $tName; 'Triggering Event' = $label
                    'Custom Thresholds' = $(if ($ins.UseCustomThresholds) { Format-Val $ins.EventThresholds } else { '' })
                    'Cumulative Window (days)' = $ins.CumulativeWindowDays; 'Internal Name' = [string]$ins.Name
                })
        }
    }
    foreach ($t in @($p.Triggers)) { $tn = Format-TriggerName ([string]$t); if ($t -and -not $triggerNames.Contains($tn)) { $triggerNames.Add($tn) } }
    foreach ($pair in @(@('DLP policy', $p.DlpPolicy), @('DLP policies', $p.DlpPoliciesAsTrigger), @('Communication Compliance policy', $p.CCPolicyName))) {
        if (Test-Blank $pair[1]) { continue }
        $v = Format-Val (Get-Parsed $pair[1])
        $triggerNames.Add("$($pair[0]): $v")
        Add-Row $triggerSheet ([ordered]@{ 'Policy' = $name; 'Trigger' = $pair[0]; 'Triggering Event' = $v })
    }

    # Indicators
    $selected = 0
    foreach ($ind in @(Get-Parsed $p.Indicators) + @(Get-Parsed $p.ExtensibleIndicators)) {
        if (-not $ind -or -not $ind.Enabled) { continue }
        $selected++
        $iName = [string]$ind.Name
        Add-Row $indicatorSheet ([ordered]@{
                'Policy' = $name; 'Category' = Get-IndicatorCategory $iName ([string]$ind.Type); 'Indicator' = Get-IndicatorLabel $iName
                'Indicator Type' = $(if ($indicatorTypeText.ContainsKey([string]$ind.Type)) { $indicatorTypeText[[string]$ind.Type] } else { [string]$ind.Type })
                'Used In Risk Scoring' = $(if ($ind.UseInScoring) { 'Yes' } else { 'No' })
                'Threshold Mode' = Format-ThresholdMode ([string]$ind.ThresholdMode); 'Thresholds (events per day)' = Format-Thresholds $ind
                'Internal Name' = $iName
            })
        if (-not $matrix.ContainsKey($iName)) { $matrix[$iName] = @{ Type = [string]$ind.Type; Policies = @{} } }
        $matrix[$iName].Policies[$name] = $(if ($ind.UseInScoring) { 'On (scored)' } else { 'On' })
    }

    # Timeframes
    $time = @("Activation window: $($p.InScopeTimeSpan) days", "Past activity: $($p.HistoricTimeSpan) days")
    if ($scenario -in $departingTemplates) {
        $time += "Termination window: $($p.PastTerminationWindow) days before / $($p.FutureTerminationWindow) days after"
        if ($p.PostTerminationActivity -eq $true) { $time += 'Post-termination activity: On' }
    }

    $extra = [System.Collections.Generic.List[string]]::new()
    foreach ($prop in $p.PSObject.Properties) {
        if ($prop.Name -in $policyHandled -or $prop.Name -in $policyNoise -or (Test-Blank $prop.Value)) { continue }
        $v = Format-Val (Get-Parsed $prop.Value)
        if ($v) { $extra.Add("$(Get-FriendlyName $prop.Name): $v") }
    }
    $pc = ConvertFrom-JsonString $p.PolicyConstraints
    if ($pc -and $pc -isnot [string]) {
        foreach ($cp in $pc.PSObject.Properties) { if (-not (Test-Blank $cp.Value)) { $extra.Add("$(Get-FriendlyName $cp.Name): $(Format-Val $cp.Value)") } }
    }

    Add-Row $policySheet ([ordered]@{
            'Policy' = $name
            'Template' = $(if ($templateText.ContainsKey($scenario)) { $templateText[$scenario] } else { Get-FriendlyName $scenario })
            'Status' = Get-PolicyStatus $p; 'Priority' = [int]$p.Priority
            'Users In Scope' = $scope -join "`n"; 'Excluded Users' = $excluded -join "`n"
            'Sensitivity Labels (priority content)' = (Get-Names $p.SensitivityLabels $labelMap) -join "`n"
            'Sensitive Info Types (priority content)' = (Get-Names $p.DlpSensitiveTypes $null) -join "`n"
            'Trainable Classifiers (priority content)' = (Get-Names $p.MLClassifierTypes $null) -join "`n"
            'SharePoint Sites (priority content)' = (@(foreach ($s in (Get-Parsed $p.SharepointSites)) { if ($s -is [string]) { $s } elseif ($s.Url) { $s.Url } else { Format-Val $s } }) -join "`n")
            'Teams (priority content)' = (Get-Names $p.TeamsSites $null) -join "`n"
            'File Extensions (priority content)' = (Get-Names $p.FileExtensions $null) -join ', '
            'Score Only Priority Content' = $(if ($p.IsPriorityContentOnlyScoring) { 'Yes' } else { 'No' })
            'Triggers' = $triggerNames -join "`n"; 'Triggering Events' = ($events | Select-Object -Unique) -join "`n"
            'Indicators Selected' = $selected; 'Timeframes' = $time -join "`n"
            'Policy Health' = Format-Health $p.PolicyHealth; 'Description' = [string]$p.Comment; 'Additional Settings' = $extra -join "`n"
            'Created By' = [string]$p.CreatedBy; 'Created (UTC)' = Format-Date $p.WhenCreatedUTC
            'Last Modified By' = [string]$p.LastModifiedBy; 'Last Modified (UTC)' = Format-Date $p.WhenChangedUTC
        })
}

# Indicator matrix: global state plus each policy
$matrixCols = [ordered]@{ 'Category' = 22; 'Indicator' = 50; 'Global Setting' = 10 }
foreach ($n in $policyNames) { $matrixCols[$n] = 16 }
$matrixCols['Internal Name'] = 40
$matrixHighlights = [ordered]@{ 'Global Setting' = $statusHighlight }
foreach ($n in $policyNames) { $matrixHighlights[$n] = [ordered]@{ 'On (scored)' = '#A9D08E'; 'On' = '#C6EFCE' } }
$matrixSheet = New-Sheet 'Indicator Matrix' $matrixCols 2 $matrixHighlights
foreach ($k in @($globalIndicators.Keys)) { if ($globalIndicators[$k] -and -not $matrix.ContainsKey($k)) { $matrix[$k] = @{ Type = ''; Policies = @{} } } }
$matrixRows = foreach ($k in $matrix.Keys) {
    $row = [ordered]@{
        'Category' = Get-IndicatorCategory $k $matrix[$k].Type; 'Indicator' = Get-IndicatorLabel $k
        'Global Setting' = $(if ($globalIndicators.Contains($k)) { if ($globalIndicators[$k]) { 'On' } else { 'Off' } } else { 'n/a' })
        'Internal Name' = $k
    }
    foreach ($n in $policyNames) { $row[$n] = $matrix[$k].Policies[$n] }
    [pscustomobject]@{ Category = $row.Category; Label = $row.Indicator; Row = $row }
}
foreach ($m in @($matrixRows | Sort-Object Category, Label)) { Add-Row $matrixSheet $m.Row }

# Global settings
$globalSheet = New-Sheet 'Global Settings' ([ordered]@{ 'Section' = 26; 'Setting' = 60; 'Value' = 60; 'Internal Key' = 45 }) 2 $null
if ($tenantSettings) {
    $globalRows = [System.Collections.Generic.List[object]]::new()
    foreach ($prop in $tenantSettings.PSObject.Properties) {
        $n = $prop.Name
        if ($n -in 'Indicators', 'ExtensibleIndicators') { continue }
        if ($n -eq 'NotificationPreferences') {
            foreach ($np in @($prop.Value)) {
                if (-not $np) { continue }
                $t = [string]$np.NotificationType
                $label = if ($notificationText.ContainsKey($t)) { $notificationText[$t] } else { Get-FriendlyName $t }
                $val = "$(Format-OnOff $np.Enabled)$(if ($np.Enabled) { " - $(Format-Recipients $np)" })"
                $globalRows.Add(@('Admin notifications', $label, $val, "NotificationPreferences.$t"))
            }
            continue
        }
        $value = $prop.Value
        $items = [System.Collections.Generic.List[object]]::new()
        if ($value -and $value -isnot [string] -and $value -isnot [ValueType] -and $value -isnot [System.Collections.IEnumerable] -and
            -not $value.PSObject.Properties['Enabled']) {
            foreach ($c in $value.PSObject.Properties) { $items.Add(@("$n.$($c.Name)", $c.Value)) }
        }
        else { $items.Add(@($n, $value)) }
        foreach ($it in $items) {
            $key = $it[0]
            $map = $globalSettingText[$key]
            $section = if ($map) { $map[0] } else { Get-FriendlyName ($key -split '\.')[0] }
            $label = if ($map) { $map[1] } else { Get-FriendlyName ($key -split '\.')[-1] }
            $globalRows.Add(@($section, $label, (Format-SettingValue $it[1]), $key))
        }
    }
    $onCount = @($globalIndicators.Values | Where-Object { $_ }).Count
    $globalRows.Add(@('Policy indicators', 'Indicators turned on for the tenant', "$onCount of $($globalIndicators.Count) (see Indicator Matrix)", 'Indicators / ExtensibleIndicators'))
    if ($tenantPolicy) {
        $globalRows.Add(@('General', 'Settings last modified', "$(Format-Date $tenantPolicy.WhenChangedUTC) by $($tenantPolicy.LastModifiedBy)", 'WhenChangedUTC'))
    }
    $order = @('General', 'Privacy', 'Policy indicators', 'Policy timeframes', 'Intelligent detections', 'Analytics', 'Data sharing', 'Export alerts',
        'Admin notifications', 'Inline alert customisation', 'Microsoft Teams', 'Adaptive Protection')
    $keyOrder = @($globalSettingText.Keys)
    $sortedRows = $globalRows | Sort-Object { $i = [array]::IndexOf($order, $_[0]); if ($i -lt 0) { 99 } else { $i } }, { $_[0] },
        { $i = [array]::IndexOf($keyOrder, $_[3]); if ($i -lt 0) { 999 } else { $i } }, { $_[1] }
    foreach ($r in @($sortedRows)) {
        Add-Row $globalSheet ([ordered]@{ 'Section' = $r[0]; 'Setting' = $r[1]; 'Value' = $r[2]; 'Internal Key' = $r[3] })
    }
}

# Settings lists and entries
$listNameMap = @{}
foreach ($l in $lists) { $listNameMap[[string]$l.Name] = $(if ($l.DisplayName) { [string]$l.DisplayName } else { [string]$l.Name }) }

function Get-ListArea($List) {
    $n = [string]$List.Name
    if ($systemListText.ContainsKey($n)) { return $systemListText[$n] }
    if ([string]$List.ListType -like 'Custom*') { return "Detection group - $(Get-FriendlyName (([string]$List.ListType) -replace '^Custom' -replace 'Lists$'))" }
    Get-FriendlyName ([string]$List.ListType)
}

$listSheet = New-Sheet 'Settings Lists' ([ordered]@{
        'Setting Area' = 50; 'List' = 36; 'List Type' = 30; 'Origin' = 14; 'Entries' = 9; 'Description' = 40; 'Last Modified By' = 22; 'Last Modified (UTC)' = 17; 'Internal Name' = 30
    }) 2 ([ordered]@{ 'Entries' = [ordered]@{ '0' = '#F2F2F2' } })
$entrySheet = New-Sheet 'List Entries' ([ordered]@{ 'Setting Area' = 50; 'List' = 36; 'Entry' = 50; 'Details' = 40 }) 2 $null

$orderedLists = @($lists | Sort-Object @{ e = { if ($_.IsSystemDefined) { 0 } else { 1 } } }, @{ e = { Get-ListArea $_ } })
foreach ($l in $orderedLists) {
    $area = Get-ListArea $l
    $listName = $listNameMap[[string]$l.Name]
    $count = if ($null -ne $l.EntitiesCount) { [int]$l.EntitiesCount } else { @($l.Entities).Count }
    Add-Row $listSheet ([ordered]@{
            'Setting Area' = $area; 'List' = $listName; 'List Type' = Get-FriendlyName ([string]$l.ListType)
            'Origin' = $(if ($l.IsSystemDefined) { 'Built-in' } else { 'Custom' }); 'Entries' = $count; 'Description' = [string]$l.Description
            'Last Modified By' = [string]$l.LastModifiedBy; 'Last Modified (UTC)' = Format-Date $l.WhenChangedUTC; 'Internal Name' = [string]$l.Name
        })
    foreach ($e in (Get-Parsed $l.Entities)) {
        $entry = ''; $details = ''
        if ($e -is [string]) { $entry = $e }
        elseif ($e.PSObject.Properties['Dmn']) { $entry = [string]$e.Dmn; if ($e.isMLSubDmn) { $details = 'Includes multi-level subdomains' } }
        elseif ($e.PSObject.Properties['GroupId']) {
            $gid = [string]$e.GroupId
            $entry = if ($listNameMap.ContainsKey($gid)) { $listNameMap[$gid] } else { $gid }
            $details = 'Detection group applied tenant-wide'
        }
        elseif ($e.PSObject.Properties['Url']) { $entry = [string]$e.Url; $details = [string]$e.Name }
        elseif ($e.PSObject.Properties['FlPthRgx']) { $entry = [string]$e.FlPthRgx }
        elseif ($e.PSObject.Properties['Ext']) { $entry = [string]$e.Ext }
        elseif ($e.PSObject.Properties['Guid']) {
            $g = [string]$e.Guid
            if ($sitMap.ContainsKey($g)) {
                $entry = [string]$sitMap[$g].Name
                $details = if ($sitMap[$g].Publisher -eq 'Microsoft Corporation') { 'Built-in sensitive info type' } else { "Custom sensitive info type ($($sitMap[$g].Publisher))" }
            }
            elseif ($e.Name) { $entry = [string]$e.Name }
            else { $entry = $g; $details = 'GUID not resolved' }
        }
        elseif ($e.PSObject.Properties['Name']) { $entry = [string]$e.Name }
        else { $entry = Format-Val ($e | Select-Object * -ExcludeProperty Operation) }
        Add-Row $entrySheet ([ordered]@{ 'Setting Area' = $area; 'List' = $listName; 'Entry' = $entry; 'Details' = $details })
    }
}
$entryRows = @($entrySheet.Rows)
if ($entryRows.Count) {
    $sorted = @($entryRows | Sort-Object { $_[0] }, { $_[1] }, { [string]$_[3] }, { $_[2] })
    $entrySheet.Rows.Clear()
    foreach ($r in $sorted) { $entrySheet.Rows.Add($r) }
}

# Priority user groups referenced by policies
$pugSheet = New-Sheet 'Priority User Groups' ([ordered]@{ 'Priority User Group' = 30; 'Used By Policies' = 60 }) 1 $null
foreach ($g in $priorityGroups) {
    $users = @($policies | Where-Object { (Get-Names $_.CustomTags $null) -contains $g } | ForEach-Object { [string]$_.Name })
    Add-Row $pugSheet ([ordered]@{ 'Priority User Group' = $g; 'Used By Policies' = $users -join "`n" })
}

# Cover
$listsWithEntries = @($lists | Where-Object { $_.EntitiesCount -gt 0 }).Count
$cover = @{
    Name  = 'Cover'
    Kind  = 'Cover'
    Title = "$ClientName - Microsoft Purview Insider Risk Management Configuration"
    Rows  = [System.Collections.Generic.List[object]]::new()
}
$coverRows = @(
    , @('Client', $ClientName)
    , @('Tenant', $tenant)
    , @('Snapshot taken (UTC)', $snapshot)
    , @('Exported by', $(if ($exportInfo) { [string]$exportInfo.ExportedBy } else { 'Not recorded' }))
    , @('Workbook generated (UTC)', (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm'))
    , @('Policies', ('{0} (On: {1}, Off: {2})' -f $policies.Count, @($statusList -eq 'On').Count, @($statusList -ne 'On').Count))
    , @('Indicators turned on (global)', ('{0} of {1}' -f @($globalIndicators.Values | Where-Object { $_ }).Count, $globalIndicators.Count))
    , @('Settings lists', ('{0} ({1} with entries)' -f $lists.Count, $listsWithEntries))
    , @('Priority user groups in use', $(if ($priorityGroups.Count) { $priorityGroups -join ', ' } else { 'None' }))
    , @('', '')
    , @('Sheet: Policies', 'One row per policy: template, status, users in scope, priority content, triggers, timeframes, health and description.')
    , @('Sheet: Policy Indicators', 'Each indicator selected in each policy, whether it is used in risk scoring, and any threshold settings.')
    , @('Sheet: Indicator Matrix', 'All indicators turned on for the tenant or any policy, showing the global setting and which policies use them. "On (scored)" means the policy uses the indicator in risk scoring; "n/a" means the indicator has no tenant-wide toggle (for example sequences).')
    , @('Sheet: Policy Triggers', 'Triggering events selected for each policy.')
    , @('Sheet: Global Settings', 'Tenant-wide IRM settings: privacy, timeframes, intelligent detections, analytics, data sharing, notifications and Adaptive Protection.')
    , @('Sheet: Settings Lists', 'Built-in settings lists (domains, global exclusions, priority physical assets) and custom detection groups, with entry counts.')
    , @('Sheet: List Entries', 'Every entry in the settings lists and detection groups.')
    , @('Sheet: Priority User Groups', 'Priority user groups referenced by policies.')
    if ($roleGroups.Count) { , @('Sheet: Role Groups', 'Insider Risk Management role groups and their members.') }
    , @('', '')
    , @('Note', 'This is a point-in-time snapshot of the configuration as exported from the tenant. Changes made after the snapshot date are not reflected.')
    , @('Note', 'Exported with the Insider Risk Management PowerShell cmdlets, which Microsoft does not document. Labels for some settings are derived from internal names; the Internal Name / Internal Key columns show the source value.')
    , @('Note', 'Priority user group membership is not available via PowerShell; only group names referenced by policies are shown.')
)
foreach ($c in $coverRows) { $cover.Rows.Add([object[]]@((ConvertTo-CellValue $c[0]), (ConvertTo-CellValue $c[1]))) }

$sheets = @($cover, $policySheet, $indicatorSheet, $matrixSheet, $triggerSheet, $globalSheet, $listSheet, $entrySheet, $pugSheet)
if ($roleGroups.Count) {
    $roleSheet = New-Sheet 'Role Groups' ([ordered]@{ 'Role Group' = 40; 'Member Count' = 10; 'Members' = 60; 'Description' = 60 }) 1 $null
    foreach ($g in $roleGroups) {
        $members = @(foreach ($m in @($g.Members)) { if ($m) { if ($m.Address) { "$($m.Name) <$($m.Address)>" } else { [string]$m.Name } } })
        Add-Row $roleSheet ([ordered]@{
                'Role Group' = $(if ($g.DisplayName) { [string]$g.DisplayName } else { [string]$g.Name }); 'Member Count' = $members.Count
                'Members' = $members -join "`n"; 'Description' = [string]$g.Description
            })
    }
    $sheets += $roleSheet
}
#endregion

$spec = @{ Path = $OutputPath; Sheets = $sheets }
$useImportExcel = $Engine -eq 'ImportExcel' -or ($Engine -eq 'Auto' -and (Get-Module -ListAvailable -Name ImportExcel))
if ($useImportExcel) { Write-WorkbookImportExcel $spec } else { Write-WorkbookPython $spec }
Write-Host "Workbook written to $OutputPath"
