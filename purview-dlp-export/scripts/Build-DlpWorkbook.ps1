<#
.SYNOPSIS
    Builds a single filterable Excel workbook from a DLP export folder created by Export-DlpConfiguration.ps1.
.DESCRIPTION
    Works offline from DlpPolicies.raw.json and DlpRules.raw.json (plus SensitiveInfoTypes.json, Labels.json and
    ExportInfo.json when present). Uses the ImportExcel module if installed, otherwise Python with openpyxl.
.EXAMPLE
    .\Build-DlpWorkbook.ps1 -ExportFolder .\DLP-Export-20261006-1516
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
$scopeMap = [ordered]@{
    'Exchange'    = @(('ExchangeLocation', 'ExchangeSender', 'ExchangeSenderMemberOf', 'ExchangeAdaptiveScopes'),
                      ('ExchangeSenderException', 'ExchangeSenderMemberOfException', 'ExchangeAdaptiveScopesException'))
    'SharePoint'  = @(('SharePointLocation', 'SharePointAdaptiveScopes'),
                      ('SharePointLocationException', 'SharePointAdaptiveScopesException'))
    'OneDrive'    = @(('OneDriveLocation', 'OneDriveSharedBy', 'OneDriveSharedByMemberOf', 'OneDriveAdaptiveScopes'),
                      ('OneDriveLocationException', 'ExceptIfOneDriveSharedBy', 'ExceptIfOneDriveSharedByMemberOf', 'OneDriveAdaptiveScopesException'))
    'Teams'       = @(('TeamsLocation', 'TeamsAdaptiveScopes'), ('TeamsLocationException', 'TeamsAdaptiveScopesException'))
    'Devices'     = @(('EndpointDlpLocation', 'EndpointDlpAdaptiveScopes'), ('EndpointDlpLocationException', 'EndpointDlpAdaptiveScopesException'))
    'Cloud Apps'  = @(('ThirdPartyAppDlpLocation'), ('ThirdPartyAppDlpLocationException'))
    'On-premises' = @(('OnPremisesScannerDlpLocation', 'SharePointOnPremisesLocation', 'ExchangeOnPremisesLocation'),
                      ('OnPremisesScannerDlpLocationException', 'SharePointOnPremisesLocationException'))
    'Power BI'    = @(('PowerBIDlpLocation'), ('PowerBIDlpLocationException'))
}

$accessScopeText = @{
    InOrganization    = 'Content is shared only with people inside the organisation'
    NotInOrganization = 'Content is shared with people outside the organisation'
}

$endpointActivity = @{
    CloudEgress                    = 'Upload to restricted cloud service domain / unallowed browser'
    CopyPaste                      = 'Copy to clipboard'
    Print                          = 'Print'
    RemovableMedia                 = 'Copy to removable USB device'
    NetworkShare                   = 'Copy to network share'
    UnallowedApps                  = 'Access by restricted apps'
    UnallowedBluetoothTransferApps = 'Transfer via unallowed Bluetooth apps'
    UnallowedFtpTransferApps       = 'Transfer via unallowed FTP apps'
    RemoteDesktopServices          = 'Copy to Remote Desktop session'
    PasteToBrowser                 = 'Paste to supported browsers'
    RequireBusinessJustification   = 'Override requires business justification'
}

$conditionText = @{
    HasActivity      = 'Activity is'
    From             = 'Sender is'
    FromMemberOf     = 'Sender is member of'
    SenderDomainIs   = 'Sender domain is'
    SentTo           = 'Recipient is'
    SentToMemberOf   = 'Recipient is member of'
    RecipientDomainIs = 'Recipient domain is'
}

$restrictionLabel = [ordered]@{
    RestrictAccess                   = 'Restrict access'
    MipRestrictAccess                = 'Restrict access (label)'
    EndpointDlpRestrictions          = 'Device'
    EndpointDlpBrowserRestrictions   = 'Device browser'
    ThirdPartyAppDlpRestrictions     = 'Cloud apps'
    OnPremisesScannerDlpRestrictions = 'On-premises'
    PowerBIDlpRestrictions           = 'Power BI'
}

$otherActionProps = 'EncryptRMSTemplate', 'RemoveRMSTemplate', 'ApplyBrandingTemplate', 'Quarantine', 'MoveToQuarantineLocation',
    'SPMoveToQuarantineLocation', 'RedirectMessageTo', 'AddRecipients', 'Moderate', 'ModifySubject', 'PrependSubject', 'SetHeader',
    'RemoveHeader', 'ApplyHtmlDisclaimer', 'BlockDomainsOrUsers', 'BlockDomains', 'BlockDomainsExcept', 'BlockUsers', 'BlockUsersExcept',
    'RestrictWebGrounding', 'RedactSensitiveInfo', 'TriggerPowerAutomateFlow', 'StopPolicyProcessing', 'AccessTimeControl',
    'RestrictBrowserAccess', 'MapRecipients'

$notifyProps = 'NotifyUser', 'NotifyUserType', 'NotifyPolicyTipDisplayOption', 'NotifyEmailCustomSubject', 'NotifyEmailCustomText',
    'NotifyEmailCustomSenderDisplayName', 'NotifyPolicyTipUrl', 'NotifyPolicyTipCustomDialog', 'NotifyEndpointUser',
    'NotifyPolicyTipCustomText', 'NotifyPolicyTipCustomTextTranslations', 'NotifyEmailExchangeIncludeAttachment',
    'NotifyEmailOnedriveRemediationActions', 'NotifyAllowOverride', 'NotifyOverrideRequirements', 'NotifyJustificationCustomText',
    'NotifyJustificationCustomTextTranslations'

$commonNoise = 'Identity', 'Id', 'Guid', 'ImmutableId', 'ExchangeObjectId', 'ObjectVersion', 'DirectoryObjectVersion', 'IsValid',
    'ExchangeVersion', 'DistinguishedName', 'ObjectCategory', 'ObjectClass', 'WhenChanged', 'WhenCreated', 'OrganizationalUnitRoot',
    'OrganizationId', 'OriginatingServer', 'ObjectState', 'ReadOnly', 'ExternalIdentity', 'ErrorMetadata', 'Workload',
    'RunspaceId', 'PSComputerName', 'PSShowComputerName'

$ruleNoise = $commonNoise + @('Policy', 'MaximumBlobRuleLength', 'IsObjectUnderSystemOperation', 'IsSummarizedPsRule', 'RuleXml',
    'ExternalScenarioDependancies', 'ExecutionRuleGuids', 'AdvancedRuleBuilderContext', 'StorageBindings', 'IsAdvancedRule',
    'EnforcePortalAccess', 'Mode', 'SourceType', 'NonBifurcatingAccessScope', 'EffectiveParentFolder', 'ExceptIfEffectiveParentFolder')

$ruleHandled = @('Name', 'ParentPolicyName', 'Priority', 'Disabled', 'ReportSeverityLevel', 'Comment', 'CreatedBy', 'LastModifiedBy',
    'WhenCreatedUTC', 'WhenChangedUTC', 'AdvancedRule', 'ContentContainsSensitiveInformation', 'ExceptIfContentContainsSensitiveInformation',
    'AccessScope', 'ExceptIfAccessScope', 'BlockAccess', 'BlockAccessScope', 'GenerateAlert', 'AlertProperties', 'GenerateIncidentReport',
    'IncidentReportContent') + @($restrictionLabel.Keys) + $otherActionProps + $notifyProps

$policyNoise = $commonNoise + @('Type', 'PolicyCategory', 'PolicyRulesMetaData', 'GlobalListType', 'PolicyConstraints', 'ExpectedLocations',
    'CompletedLocations', 'FailedLocations', 'DistributionSyncStatus', 'DistributionResults', 'LastStatusUpdateTime', 'ModificationTimeUtc',
    'CreationTimeUtc', 'LocationInclusions', 'LocationExclusions', 'Locations', 'EndpointDlpExtendedLocations', 'TopNLocationStatistics',
    'WorkloadStatistics', 'ItemStatistics', 'MatchedItemsCount', 'TotalItemsCount', 'RuleMatchBlob', 'UserAdministrativeUnitMembershipMap',
    'ForceValidate', 'Summary', 'SimulationStatus', 'IsColdDataSimulationPolicy', 'ExtendedProperties', 'LogicalWorkload', 'EnforcementPlanes')

$policyHandled = @('Name', 'DisplayName', 'Mode', 'Enabled', 'Priority', 'Comment', 'CreatedBy', 'LastModifiedBy', 'WhenCreatedUTC',
    'WhenChangedUTC', 'DistributionStatus') + @($scopeMap.Values | ForEach-Object { $_[0]; $_[1] })
#endregion

#region Value helpers
function Read-ExportJson([string]$Name) {
    $path = Join-Path $ExportFolder $Name
    if (-not (Test-Path $path)) { return }
    # Raw exports serialise enums as {"value":1,"Value":"Text"}; the case-duplicate key breaks ConvertFrom-Json.
    $text = (Get-Content $path -Raw -Encoding UTF8) -creplace '"value":\s*-?\d+,(\s*)"Value"', '$1"Value"'
    $data = $text | ConvertFrom-Json
    $data
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
    $Name -creplace '(?<=[a-z0-9])(?=[A-Z])', ' '
}

function ConvertFrom-JsonString($v) {
    if ($v -isnot [string]) { return $v }
    $t = $v.TrimStart()
    if (-not ($t.StartsWith('[') -or $t.StartsWith('{'))) { return $v }
    try { $parsed = $v | ConvertFrom-Json; return , $parsed } catch { return $v }
}

# Builds a GUID -> name map from any object in the tree carrying both an id and a name (e.g. Endpoint DLP groups).
function Add-NamedIds($Node, $Map, [int]$Depth = 0) {
    if ($null -eq $Node -or $Depth -gt 12 -or $Node -is [ValueType]) { return }
    if ($Node -is [string]) {
        $parsed = ConvertFrom-JsonString $Node
        if ($parsed -isnot [string]) { Add-NamedIds $parsed $Map ($Depth + 1) }
        return
    }
    if ($Node -is [System.Collections.IEnumerable]) { foreach ($i in $Node) { Add-NamedIds $i $Map ($Depth + 1) }; return }
    $props = $Node.PSObject.Properties
    $id = foreach ($n in 'id', 'groupId', 'Identity', 'Guid') { if ($props[$n] -and [string]$Node.$n -match $guidPattern) { [string]$Node.$n; break } }
    $name = foreach ($n in 'name', 'groupName', 'displayName') { if ($props[$n] -and [string]$Node.$n -and [string]$Node.$n -notmatch $guidPattern) { [string]$Node.$n; break } }
    if ($id -and $name -and -not $Map.ContainsKey($id)) { $Map[$id] = $name }
    foreach ($p in $props) { Add-NamedIds $p.Value $Map ($Depth + 1) }
}

function ConvertTo-CellValue($v) {
    if ($null -eq $v) { return '' }
    if ($v -is [int] -or $v -is [long] -or $v -is [double]) { return $v }
    $s = [string]$v
    if ($s.Length -gt 32000) { $s = $s.Substring(0, 32000) + ' ...[truncated - see raw JSON]' }
    # Prevent tenant-sourced text being interpreted as a formula.
    if ($s -match '^[=+@]') { $s = ' ' + $s }
    $s
}
#endregion

#region Policy scope
function Format-Location($Item) {
    if ($Item -is [string]) {
        if (-not $Item.TrimStart().StartsWith('{')) { return $Item }
        try { $Item = $Item | ConvertFrom-Json } catch { return $Item }
    }
    foreach ($n in 'DisplayName', 'PrimarySmtpAddress', 'Name', 'Identity') {
        $val = [string]$Item.$n
        if ($val -and $val -notmatch $guidPattern) { return $val }
    }
    [string]$Item.Name
}

function Get-ScopePrefix([string]$Prop) {
    if ($Prop -like '*AdaptiveScopes*') { return 'Adaptive scope: ' }
    if ($Prop -like '*MemberOf*') { return 'Group: ' }
    ''
}

function Add-ScopeItems($Scope, [string]$Workload, $In, $Out) {
    if (-not $Scope.Contains($Workload)) {
        $Scope[$Workload] = @{ In = [System.Collections.Generic.List[string]]::new(); Out = [System.Collections.Generic.List[string]]::new() }
    }
    foreach ($i in $In) { if ($i -and -not $Scope[$Workload].In.Contains($i)) { $Scope[$Workload].In.Add($i) } }
    foreach ($o in $Out) { if ($o -and -not $Scope[$Workload].Out.Contains($o)) { $Scope[$Workload].Out.Add($o) } }
}

function Get-PolicyScope($Policy) {
    $scope = [ordered]@{}
    foreach ($w in $scopeMap.Keys) {
        $in = @(foreach ($prop in $scopeMap[$w][0]) { foreach ($i in @($Policy.$prop)) { if (-not (Test-Blank $i)) { (Get-ScopePrefix $prop) + (Format-Location $i) } } })
        $out = @(foreach ($prop in $scopeMap[$w][1]) { foreach ($i in @($Policy.$prop)) { if (-not (Test-Blank $i)) { (Get-ScopePrefix $prop) + (Format-Location $i) } } })
        if ($in.Count -or $out.Count) { Add-ScopeItems $scope $w $in $out }
    }
    if ($Policy.EndpointDlpExtendedLocations) {
        $ext = $Policy.EndpointDlpExtendedLocations | ConvertFrom-Json
        foreach ($g in $ext) {
            $prefix = if ($g.GroupSet) { "$($g.GroupSet): " } else { '' }
            $in = @(foreach ($i in @($g.Inclusions)) { if ($i -and [string]$i.Type -ne 'Tenant') { $prefix + (Format-Location $i) } })
            $out = @(foreach ($i in @($g.Exclusions)) { if ($i) { $prefix + (Format-Location $i) } })
            if ($in.Count -or $out.Count) { Add-ScopeItems $scope 'Devices' $in $out }
        }
    }
    if ($Policy.Locations -and $Policy.Locations -ne '[]') {
        $locs = $Policy.Locations | ConvertFrom-Json
        foreach ($l in $locs) {
            $name = if ($l.Location -eq 'Copilot.M365') { 'Microsoft 365 Copilot' } elseif ($l.Location) { "$($l.Workload): $($l.Location)" } else { [string]$l.Workload }
            $in = @(foreach ($i in @($l.Inclusions)) { if ($i) { Format-Location $i } })
            $out = @(foreach ($i in @($l.Exclusions)) { if ($i) { Format-Location $i } })
            Add-ScopeItems $scope $name $in $out
        }
    }
    foreach ($w in @($scope.Keys)) {
        $s = $scope[$w]
        if ($s.In.Count -gt 1 -and $s.In.Contains('All')) { [void]$s.In.Remove('All') }
        if ($s.In.Count -eq 0) { $s.In.Add('All') }
    }
    $scope
}

function Format-Scope($Scope, [string]$Part) {
    @(foreach ($w in $Scope.Keys) { $items = $Scope[$w][$Part]; if ($items.Count) { "${w}: $($items -join '; ')" } }) -join "`n"
}

function Get-PolicyMode($Policy) {
    if ($Policy.Enabled -eq $false) { return 'Off' }
    switch ([string]$Policy.Mode) {
        'Enable' { 'On' }
        'TestWithNotifications' { 'Simulation (policy tips shown)' }
        'TestWithoutNotifications' { 'Simulation' }
        'Disable' { 'Off' }
        default { [string]$Policy.Mode }
    }
}
#endregion

#region Rule conditions
function New-ContentMatch([string]$Kind, $Item, [string]$Group, [string]$Target, [bool]$IsException) {
    $id = [string]$Item.id
    $name = [string]$Item.name
    if ($Kind -eq 'SIT') {
        if ($id -and $sitMap.ContainsKey($id)) { $name = $sitMap[$id].Name }
        elseif ($sitMap.ContainsKey($name)) { $name = $sitMap[$name].Name }
    }
    elseif ($id -and $labelMap.ContainsKey($id)) { $name = $labelMap[$id] }
    elseif ($labelMap.ContainsKey($name)) { $name = $labelMap[$name] }
    [pscustomobject]@{
        Kind = $Kind; Name = $name; Id = $id; Group = $Group; Target = $Target; IsException = $IsException
        MinCount = $Item.mincount; MaxCount = $Item.maxcount; Confidence = [string]$Item.confidencelevel
    }
}

# Walks ContentContainsSensitiveInformation values (legacy flat or grouped format) collecting SIT and label references.
function Add-ContentMatches($Node, $Bag, [string]$Group, [string]$Target, [bool]$IsException) {
    if ($null -eq $Node -or $Node -is [string] -or $Node -is [ValueType]) { return }
    if ($Node -is [System.Collections.IEnumerable]) {
        foreach ($i in $Node) { Add-ContentMatches $i $Bag $Group $Target $IsException }
        return
    }
    $props = $Node.PSObject.Properties
    if ($props['groups']) {
        $t = if ($props['target']) { [string]$Node.target } else { $Target }
        Add-ContentMatches $Node.groups $Bag $Group $t $IsException
        return
    }
    if ($props['sensitivetypes'] -or $props['labels']) {
        foreach ($s in @($Node.sensitivetypes)) { if ($s) { $Bag.Add((New-ContentMatch 'SIT' $s ([string]$Node.name) $Target $IsException)) } }
        foreach ($l in @($Node.labels)) { if ($l) { $Bag.Add((New-ContentMatch 'Label' $l ([string]$Node.name) $Target $IsException)) } }
        return
    }
    if ($props['name'] -and ($props['mincount'] -or $props['confidencelevel'])) {
        $Bag.Add((New-ContentMatch 'SIT' $Node $Group $Target $IsException))
        return
    }
    foreach ($p in $props) { Add-ContentMatches $p.Value $Bag $Group $Target $IsException }
}

function Format-Count($Item) {
    $min = if ("$($Item.MinCount)" -match '^\d+$') { [int]$Item.MinCount } else { 1 }
    $max = "$($Item.MaxCount)"
    if ($max -notmatch '^\d+$') { return "$min+" }
    if ([int]$max -eq $min) { return "$min" }
    "$min-$max"
}

function Format-ContentSummary($Items, [bool]$IsException) {
    $Items = @($Items)
    if (-not $Items.Count) { return '' }
    $parts = @()
    $sitNames = @($Items | Where-Object Kind -eq 'SIT' | ForEach-Object Name | Select-Object -Unique)
    $labelNames = @($Items | Where-Object Kind -eq 'Label' | ForEach-Object Name | Select-Object -Unique)
    if ($sitNames.Count) {
        $parts += if ($IsException -or $sitNames.Count -le 3) { "Content contains: $($sitNames -join ', ')" }
                  else { "Content contains any of $($sitNames.Count) sensitive info types" }
    }
    if ($labelNames.Count) { $parts += "Content has label: $($labelNames -join ', ')" }
    $text = $parts -join ' OR '
    $targets = @($Items | Where-Object Target | ForEach-Object Target | Select-Object -Unique)
    if ($targets.Count) { $text += " (in $($targets -join ', '))" }
    $text
}

function Format-Condition($Node, $Bag, [bool]$IsException, [int]$Depth) {
    if ($null -eq $Node) { return '' }
    if ($Node.PSObject.Properties['SubConditions']) {
        $op = [string]$Node.Operator
        $neg = $op -eq 'Not'
        $parts = @(foreach ($s in @($Node.SubConditions)) { $t = Format-Condition $s $Bag ($IsException -or $neg) ($Depth + 1); if ($t) { $t } })
        $text = $parts -join $(if ($op -eq 'Or') { ' OR ' } else { ' AND ' })
        if ($neg) { return "NOT ($text)" }
        if ($parts.Count -gt 1 -and $Depth -gt 1) { return "($text)" }
        return $text
    }
    $name = [string]$Node.ConditionName
    $value = $Node.Value
    if ($name -like '*ContentContainsSensitiveInformation') {
        $start = $Bag.Count
        Add-ContentMatches $value $Bag '' '' ($IsException -or $name -like 'ExceptIf*')
        return Format-ContentSummary @(for ($i = $start; $i -lt $Bag.Count; $i++) { $Bag[$i] }) $IsException
    }
    if ($name -eq 'AccessScope' -and $accessScopeText.Contains([string]$value)) { return $accessScopeText[[string]$value] }
    $label = if ($conditionText.Contains($name)) { $conditionText[$name] } else { Get-FriendlyName $name }
    "${label}: $(Format-Val $value)"
}

function Format-Restriction($Item) {
    if ($Item -is [string]) { return $Item }
    $setting = [string]$Item.setting
    if ($endpointActivity.ContainsKey($setting)) { $setting = $endpointActivity[$setting] }
    $main = if ($setting) { "$setting = $($Item.value)" } else { [string]$Item.value }
    $extras = @(foreach ($p in $Item.PSObject.Properties) {
        if ($p.Name -in 'setting', 'value' -or (Test-Blank $p.Value)) { continue }
        $vals = @(foreach ($g in @($p.Value)) {
            if ($g -isnot [string] -and $g.PSObject.Properties['action']) {
                $gid = [string]$g.id
                "$(if ($groupNames.ContainsKey($gid)) { $groupNames[$gid] } else { $gid }) -> $($g.action)"
            }
            else { Format-Val $g }
        })
        "$($p.Name): $($vals -join ', ')"
    })
    if ($extras.Count) { "$main ($($extras -join '; '))" } else { $main }
}

function Get-RuleName($Rule) {
    if (-not [string]::IsNullOrWhiteSpace([string]$Rule.DisplayName)) { return [string]$Rule.DisplayName }
    [string]$Rule.Name
}

function Get-PolicyName($Policy) {
    if (-not [string]::IsNullOrWhiteSpace([string]$Policy.DisplayName)) { return [string]$Policy.DisplayName }
    [string]$Policy.Name
}

function Get-RuleDetail($Rule) {
    $bag = [System.Collections.Generic.List[object]]::new()
    $cond = [System.Collections.Generic.List[string]]::new()
    $exc = [System.Collections.Generic.List[string]]::new()
    $extra = [System.Collections.Generic.List[string]]::new()
    $topOp = 'AND'
    $isAdvanced = -not [string]::IsNullOrWhiteSpace([string]$Rule.AdvancedRule)

    if ($isAdvanced) {
        $root = ($Rule.AdvancedRule | ConvertFrom-Json).Condition
        if ([string]$root.Operator -eq 'Or') { $topOp = 'OR' }
        $subs = if ($root.PSObject.Properties['SubConditions'] -and [string]$root.Operator -ne 'Not') { @($root.SubConditions) } else { @($root) }
        foreach ($s in $subs) {
            if ([string]$s.Operator -eq 'Not' -and $topOp -eq 'AND') {
                foreach ($x in @($s.SubConditions)) { $t = Format-Condition $x $bag $true 1; if ($t) { $exc.Add($t) } }
            }
            else {
                $t = Format-Condition $s $bag $false 1
                if ($t) { $cond.Add($t) }
            }
        }
    }
    else {
        Add-ContentMatches $Rule.ContentContainsSensitiveInformation $bag '' '' $false
        $t = Format-ContentSummary @($bag) $false
        if ($t) { $cond.Add($t) }
        $start = $bag.Count
        Add-ContentMatches $Rule.ExceptIfContentContainsSensitiveInformation $bag '' '' $true
        $t = Format-ContentSummary @(for ($i = $start; $i -lt $bag.Count; $i++) { $bag[$i] }) $true
        if ($t) { $exc.Add($t) }
        foreach ($pair in @(@('AccessScope', $cond), @('ExceptIfAccessScope', $exc))) {
            $v = [string]$Rule.($pair[0])
            if (-not (Test-Blank $v)) { $pair[1].Add($(if ($accessScopeText.Contains($v)) { $accessScopeText[$v] } else { "Access scope: $v" })) }
        }
    }

    foreach ($prop in $Rule.PSObject.Properties) {
        $n = $prop.Name
        if ($n -in $ruleHandled -or $n -in $ruleNoise -or $n -eq 'DisplayName' -or (Test-Blank $prop.Value)) { continue }
        $v = Format-Val $prop.Value
        if (-not $v) { continue }
        if ($isAdvanced) { $extra.Add("$(Get-FriendlyName $n): $v") }
        elseif ($n -like 'ExceptIf*') { $exc.Add("$(Get-FriendlyName ($n -replace '^ExceptIf', '')): $v") }
        else { $cond.Add("$(Get-FriendlyName $n): $v") }
    }

    # Sensitive info types and labels
    $sits = @($bag | Where-Object { $_.Kind -eq 'SIT' -and -not $_.IsException })
    $groups = @($sits | ForEach-Object Group | Select-Object -Unique)
    $sitLines = @(foreach ($m in $sits) {
        $conf = if ($m.Confidence) { "$($m.Confidence), " } else { '' }
        $line = '{0} ({1}{2})' -f $m.Name, $conf, (Format-Count $m)
        if ($groups.Count -gt 1 -and $m.Group) { "[$($m.Group)] $line" } else { $line }
    }) | Select-Object -Unique
    $labelNames = @($bag | Where-Object { $_.Kind -eq 'Label' -and -not $_.IsException } | ForEach-Object Name | Select-Object -Unique)

    # Actions
    $actions = [System.Collections.Generic.List[string]]::new()
    if ($Rule.BlockAccess -eq $true) {
        $scope = switch ([string]$Rule.BlockAccessScope) {
            'All' { 'everyone' }
            'PerUser' { 'people outside the organisation' }
            'PerAnonymousUser' { 'anyone with the link (anonymous access)' }
            default { [string]$Rule.BlockAccessScope }
        }
        $actions.Add("Block access: $scope")
    }
    foreach ($p in $restrictionLabel.Keys) {
        foreach ($i in @($Rule.$p)) { if (-not (Test-Blank $i)) { $actions.Add("$($restrictionLabel[$p]): $(Format-Restriction $i)") } }
    }
    foreach ($p in $otherActionProps) {
        $v = $Rule.$p
        if (Test-Blank $v) { continue }
        if ($v -is [bool]) { $actions.Add((Get-FriendlyName $p)) } else { $actions.Add("$(Get-FriendlyName $p): $(Format-Val $v)") }
    }
    if (-not $actions.Count) { $actions.Add('Audit only (no restrictive action)') }

    # User notifications and overrides
    $notify = [System.Collections.Generic.List[string]]::new()
    if (-not (Test-Blank $Rule.NotifyUser)) {
        $notify.Add('Notify: ' + ((@($Rule.NotifyUser) | ForEach-Object { Get-FriendlyName ([string]$_) }) -join ', '))
        if (-not (Test-Blank $Rule.NotifyUserType)) { $notify.Add("Method: $($Rule.NotifyUserType)") }
        if (-not (Test-Blank $Rule.NotifyPolicyTipDisplayOption)) { $notify.Add("Policy tip style: $($Rule.NotifyPolicyTipDisplayOption)") }
    }
    foreach ($p in 'NotifyEmailCustomSubject', 'NotifyEmailCustomText', 'NotifyEmailCustomSenderDisplayName', 'NotifyPolicyTipUrl',
        'NotifyPolicyTipCustomDialog', 'NotifyEndpointUser') {
        if (-not (Test-Blank $Rule.$p)) { $notify.Add("$(Get-FriendlyName ($p -replace '^Notify', '')): $(Format-Val $Rule.$p)") }
    }
    $tipText = [string]$Rule.NotifyPolicyTipCustomText
    $translations = @($Rule.NotifyPolicyTipCustomTextTranslations | Where-Object { $_ })
    if ($translations.Count) { $tipText += "`n(+$($translations.Count) translations)" }

    $override = [System.Collections.Generic.List[string]]::new()
    if (-not (Test-Blank $Rule.NotifyAllowOverride)) { $override.Add("Allowed: $(Get-FriendlyName (Format-Val $Rule.NotifyAllowOverride))") }
    if (-not (Test-Blank $Rule.NotifyOverrideRequirements)) { $override.Add("Requirements: $(Format-Val $Rule.NotifyOverrideRequirements)") }
    if (-not (Test-Blank $Rule.NotifyJustificationCustomText)) { $override.Add("Justification options: $(Format-Val $Rule.NotifyJustificationCustomText)") }

    # Alerts and incident reports
    $alertRecipients = @($Rule.GenerateAlert | Where-Object { $_ }) -join "`n"
    $alertSettings = ''
    if ($alertRecipients) {
        $agg = [string]$Rule.AlertProperties.AggregationType
        $alertSettings = if (Test-Blank $agg) { 'Alert every time an activity matches the rule' } else { Format-Val $Rule.AlertProperties }
    }
    $incident = @($Rule.GenerateIncidentReport | Where-Object { $_ }) -join "`n"
    if ($incident -and -not (Test-Blank $Rule.IncidentReportContent)) { $incident += "`nContent: $(Format-Val $Rule.IncidentReportContent)" }

    @{
        Items = $bag
        Row   = [ordered]@{
            'Rule'                                              = Get-RuleName $Rule
            'Rule Status'                                       = $(if ($Rule.Disabled -eq $true) { 'Disabled' } else { 'Enabled' })
            'Severity'                                          = [string]$Rule.ReportSeverityLevel
            'Rule Priority'                                     = [int]$Rule.Priority
            'Conditions'                                        = $cond -join "`n$topOp "
            'SIT Count'                                         = @($sits | ForEach-Object Name | Select-Object -Unique).Count
            'Sensitive Info Types (confidence, instance count)' = $sitLines -join "`n"
            'Sensitivity Labels'                                = $labelNames -join "`n"
            'Exceptions'                                        = $exc -join "`n"
            'Actions'                                           = $actions -join "`n"
            'User Notifications'                                = $notify -join "`n"
            'Policy Tip Text'                                   = $tipText
            'User Overrides'                                    = $override -join "`n"
            'Alert Recipients'                                  = $alertRecipients
            'Alert Settings'                                    = $alertSettings
            'Incident Reports'                                  = $incident
            'Additional Settings'                               = $extra -join "`n"
            'Rule Description'                                  = [string]$Rule.Comment
            'Last Modified By'                                  = [string]$Rule.LastModifiedBy
            'Last Modified (UTC)'                               = Format-Date $Rule.WhenChangedUTC
        }
    }
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
            $ws.Column(1).Width = 28
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
        ws.column_dimensions['A'].width = 28
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
    $tmp = Join-Path ([IO.Path]::GetTempPath()) "dlp-workbook-$([guid]::NewGuid())"
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
$policies = @(Read-ExportJson 'DlpPolicies.raw.json' | Where-Object { $_ })
$rules = @(Read-ExportJson 'DlpRules.raw.json' | Where-Object { $_ })
if (-not $policies.Count) { throw "DlpPolicies.raw.json not found or empty in $ExportFolder" }
$exportInfo = Read-ExportJson 'ExportInfo.json'

$sitMap = @{}
foreach ($s in @(Read-ExportJson 'SensitiveInfoTypes.json' | Where-Object { $_ })) {
    $sitMap[[string]$s.Id] = $s
    $sitMap[[string]$s.Name] = $s
}
if (-not $sitMap.Count) { Write-Warning 'SensitiveInfoTypes.json not found - re-run Export-DlpConfiguration.ps1 to classify built-in vs custom SITs.' }

$labelMap = @{}
$labels = @(Read-ExportJson 'Labels.json' | Where-Object { $_ })
$labelById = @{}
foreach ($l in $labels) { foreach ($k in $l.Id, $l.Guid) { if ($k) { $labelById[[string]$k] = $l } } }
foreach ($l in $labels) {
    $display = [string]$l.DisplayName
    if ($l.ParentId -and $labelById.ContainsKey([string]$l.ParentId)) { $display = "$($labelById[[string]$l.ParentId].DisplayName) / $display" }
    foreach ($k in $l.Id, $l.Guid, $l.Name) { if ($k) { $labelMap[[string]$k] = $display } }
}
if (-not $labels.Count) { Write-Warning 'Labels.json not found - sensitivity labels will show internal names. Re-run Export-DlpConfiguration.ps1 to resolve display names.' }

$policyConfig = @(Read-ExportJson 'PolicyConfig.raw.json' | Where-Object { $_ }) | Select-Object -First 1
$groupNames = @{}
if ($policyConfig) { Add-NamedIds $policyConfig $groupNames }
else { Write-Warning 'PolicyConfig.raw.json not found - no Endpoint Settings sheet, and device groups will show as GUIDs.' }

$tenant = ([string]$policies[0].OrganizationalUnitRoot -split '/')[-1]
if (-not $ClientName) {
    $ClientName = if ($exportInfo -and $exportInfo.ClientName) { [string]$exportInfo.ClientName } else { (Get-Culture).TextInfo.ToTitleCase(($tenant -split '\.')[0]) }
}
$snapshot = if ($exportInfo -and $exportInfo.ExportedOnUTC) { Format-Date $exportInfo.ExportedOnUTC }
            else { (Get-Item (Join-Path $ExportFolder 'DlpPolicies.raw.json')).LastWriteTimeUtc.ToString('yyyy-MM-dd HH:mm') }
if (-not $OutputPath) {
    $OutputPath = Join-Path $ExportFolder ('{0}-DLP-Configuration-{1}.xlsx' -f ($ClientName -replace '[^\w-]', ''), ($snapshot.Substring(0, 10) -replace '-', ''))
}
$OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
#endregion

#region Build sheets
$sortedPolicies = @($policies | Sort-Object { [int]$_.Priority })
$policyInfo = @{}
foreach ($p in $sortedPolicies) {
    $scope = Get-PolicyScope $p
    $policyInfo[[string]$p.Name] = @{
        Mode      = Get-PolicyMode $p
        Priority  = [int]$p.Priority
        Locations = @($scope.Keys) -join ', '
        Included  = Format-Scope $scope 'In'
        Excluded  = Format-Scope $scope 'Out'
    }
}

$modeHighlight = [ordered]@{ 'On' = '#C6EFCE'; 'Simulation' = '#FFEB9C'; 'Off' = '#D9D9D9' }

$mainSheet = New-Sheet 'DLP Rules' ([ordered]@{
        'Policy' = 34; 'Rule' = 40; 'Policy Mode' = 14; 'Rule Status' = 11; 'Severity' = 10; 'Locations' = 18; 'Included Scope' = 30
        'Excluded Scope' = 24; 'Policy Priority' = 9; 'Rule Priority' = 9; 'Conditions' = 50; 'SIT Count' = 8
        'Sensitive Info Types (confidence, instance count)' = 60; 'Sensitivity Labels' = 30; 'Exceptions' = 35; 'Actions' = 45
        'User Notifications' = 34; 'Policy Tip Text' = 40; 'User Overrides' = 26; 'Alert Recipients' = 28; 'Alert Settings' = 24
        'Incident Reports' = 24; 'Additional Settings' = 40; 'Rule Description' = 35; 'Last Modified By' = 22; 'Last Modified (UTC)' = 17
    }) 2 ([ordered]@{
        'Policy Mode' = $modeHighlight
        'Rule Status' = [ordered]@{ 'Enabled' = '#C6EFCE'; 'Disabled' = '#D9D9D9' }
        'Severity'    = [ordered]@{ 'High' = '#FFC7CE'; 'Medium' = '#FFEB9C'; 'Low' = '#DDEBF7' }
    })

$policySheet = New-Sheet 'Policies' ([ordered]@{
        'Policy' = 40; 'Mode' = 14; 'Priority' = 9; 'Locations' = 18; 'Included Scope' = 35; 'Excluded Scope' = 25
        'Rule Count' = 8; 'Rules' = 45; 'Description' = 60; 'Additional Settings' = 35
        'Created By' = 22; 'Created (UTC)' = 17; 'Last Modified By' = 22; 'Last Modified (UTC)' = 17
    }) 1 ([ordered]@{ 'Mode' = $modeHighlight })

$usage = [System.Collections.Generic.List[object]]::new()
$ruleCount = 0; $disabledCount = 0

foreach ($p in $sortedPolicies) {
    $info = $policyInfo[[string]$p.Name]
    $policyRules = @($rules | Where-Object { $_.ParentPolicyName -eq $p.Name } | Sort-Object { [int]$_.Priority })
    $base = [ordered]@{
        'Policy' = Get-PolicyName $p; 'Policy Mode' = $info.Mode; 'Locations' = $info.Locations; 'Included Scope' = $info.Included
        'Excluded Scope' = $info.Excluded; 'Policy Priority' = $info.Priority
    }

    if (-not $policyRules.Count) {
        $row = [ordered]@{}
        foreach ($k in $base.Keys) { $row[$k] = $base[$k] }
        $row['Rule'] = '(no rules)'
        Add-Row $mainSheet $row
    }
    foreach ($r in $policyRules) {
        $detail = Get-RuleDetail $r
        $row = [ordered]@{}
        foreach ($k in $base.Keys) { $row[$k] = $base[$k] }
        foreach ($k in $detail.Row.Keys) { $row[$k] = $detail.Row[$k] }
        Add-Row $mainSheet $row
        $ruleCount++
        if ($r.Disabled -eq $true) { $disabledCount++ }
        foreach ($m in $detail.Items) {
            if ($m.IsException) { continue }
            $usage.Add([pscustomobject]@{
                    Kind = $m.Kind; Name = $m.Name; Policy = Get-PolicyName $p; Rule = Get-RuleName $r
                    Threshold = $(if ($m.Kind -eq 'Label') { 'Yes' } else { Format-Count $m })
                })
        }
    }

    $extra = [System.Collections.Generic.List[string]]::new()
    foreach ($prop in $p.PSObject.Properties) {
        if ($prop.Name -in $policyHandled -or $prop.Name -in $policyNoise -or (Test-Blank $prop.Value)) { continue }
        $v = Format-Val $prop.Value
        if ($v) { $extra.Add("$(Get-FriendlyName $prop.Name): $v") }
    }
    if ($p.PolicyConstraints) {
        try {
            $pc = $p.PolicyConstraints | ConvertFrom-Json
            foreach ($cp in $pc.PSObject.Properties) { if (-not (Test-Blank $cp.Value)) { $extra.Add("Policy constraint - $(Get-FriendlyName $cp.Name): $(Format-Val $cp.Value)") } }
        }
        catch { $extra.Add("Policy constraints: $($p.PolicyConstraints)") }
    }

    Add-Row $policySheet ([ordered]@{
            'Policy' = Get-PolicyName $p; 'Mode' = $info.Mode; 'Priority' = $info.Priority
            'Locations' = $info.Locations; 'Included Scope' = $info.Included; 'Excluded Scope' = $info.Excluded
            'Rule Count' = $policyRules.Count; 'Rules' = (@($policyRules | ForEach-Object { Get-RuleName $_ }) -join "`n")
            'Description' = [string]$p.Comment; 'Additional Settings' = $extra -join "`n"
            'Created By' = [string]$p.CreatedBy; 'Created (UTC)' = Format-Date $p.WhenCreatedUTC
            'Last Modified By' = [string]$p.LastModifiedBy; 'Last Modified (UTC)' = Format-Date $p.WhenChangedUTC
        })
}

$orphans = @($rules | Where-Object { -not $policyInfo.ContainsKey([string]$_.ParentPolicyName) })
foreach ($r in $orphans) {
    $detail = Get-RuleDetail $r
    $row = [ordered]@{ 'Policy' = "$($r.ParentPolicyName) (policy not found)" }
    foreach ($k in $detail.Row.Keys) { $row[$k] = $detail.Row[$k] }
    Add-Row $mainSheet $row
    $ruleCount++
}

# SIT / label usage matrix
$sitCols = [ordered]@{ 'Sensitive Info Type / Label' = 45; 'Type' = 16; 'Policies' = 9; 'Rules' = 9 }
foreach ($p in $sortedPolicies) { $sitCols[(Get-PolicyName $p)] = 18 }
$sitSheet = New-Sheet 'SIT Usage' $sitCols 2 $null
$typeOrder = @{ 'Custom SIT' = 0; 'Built-in SIT' = 1; 'Sensitive info type' = 2; 'Sensitivity label' = 3 }
$matrixRows = foreach ($grp in ($usage | Group-Object Kind, Name)) {
    $first = $grp.Group[0]
    $type = if ($first.Kind -eq 'Label') { 'Sensitivity label' }
            elseif ($sitMap.ContainsKey($first.Name)) { if ($sitMap[$first.Name].Publisher -eq 'Microsoft Corporation') { 'Built-in SIT' } else { 'Custom SIT' } }
            else { 'Sensitive info type' }
    $row = [ordered]@{
        'Sensitive Info Type / Label' = $first.Name
        'Type'                        = $type
        'Policies'                    = @($grp.Group | ForEach-Object Policy | Select-Object -Unique).Count
        'Rules'                       = @($grp.Group | ForEach-Object { "$($_.Policy)|$($_.Rule)" } | Select-Object -Unique).Count
    }
    foreach ($pg in ($grp.Group | Group-Object Policy)) {
        $row[$pg.Name] = (@($pg.Group | ForEach-Object Threshold | Select-Object -Unique | Sort-Object { [int]('0' + ($_ -replace '\D.*$', '')) }) -join ', ')
    }
    [pscustomobject]@{ Order = $typeOrder[$type]; Name = $first.Name; Row = $row }
}
foreach ($m in @($matrixRows | Sort-Object Order, Name)) { Add-Row $sitSheet $m.Row }

# Cover
$modes = @($policyInfo.Values | ForEach-Object Mode)
$cover = @{
    Name  = 'Cover'
    Kind  = 'Cover'
    Title = "$ClientName - Microsoft Purview Data Loss Prevention Configuration"
    Rows  = [System.Collections.Generic.List[object]]::new()
}
$coverRows = @(
    , @('Client', $ClientName)
    , @('Tenant', $tenant)
    , @('Snapshot taken (UTC)', $snapshot)
    , @('Exported by', $(if ($exportInfo) { [string]$exportInfo.ExportedBy } else { 'Not recorded' }))
    , @('Workbook generated (UTC)', (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm'))
    , @('Policies', ('{0} (On: {1}, Simulation: {2}, Off: {3})' -f $modes.Count, @($modes -like 'On*').Count, @($modes -like 'Simulation*').Count, @($modes -like 'Off*').Count))
    , @('Rules', ('{0} (Enabled: {1}, Disabled: {2})' -f $ruleCount, ($ruleCount - $disabledCount), $disabledCount))
    , @('Sensitive info types in use', @($usage | Where-Object Kind -eq 'SIT' | ForEach-Object Name | Select-Object -Unique).Count)
    , @('', '')
    , @('Sheet: DLP Rules', 'One row per rule, with its parent policy mode, locations and scope repeated so you can filter by policy, rule, mode, location, severity or action. Shows conditions, sensitive info types (confidence and instance counts), labels, exceptions, actions, notifications, overrides and alerts.')
    , @('Sheet: Policies', 'One row per policy: mode, priority, locations and scope, rules and description.')
    , @('Sheet: SIT Usage', 'Sensitive info types and sensitivity labels used by each policy, with the instance count thresholds configured in its rules.')
    if ($policyConfig) { , @('Sheet: Endpoint Settings', 'Tenant-wide Endpoint DLP settings (e.g. restricted service domain groups, unallowed apps and browsers) and other DLP policy configuration.') }
    , @('', '')
    , @('Note', 'This is a point-in-time snapshot of the configuration as exported from the tenant. Changes made after the snapshot date are not reflected.')
    , @('Note', 'Instance counts show the number of matches required per item: "1+" means one or more, "1-9" means between 1 and 9. Policies in Simulation mode do not enforce their actions.')
)
foreach ($c in $coverRows) { $cover.Rows.Add([object[]]@((ConvertTo-CellValue $c[0]), (ConvertTo-CellValue $c[1]))) }
#endregion

$sheets = @($cover, $mainSheet, $policySheet, $sitSheet)
if ($policyConfig) {
    $endpointSheet = New-Sheet 'Endpoint Settings' ([ordered]@{ 'Section' = 20; 'Setting' = 40; 'Value' = 100 }) 2 $null
    # Live tenants return this as a JSON string of Setting/Value pairs, with repeated Settings for list values.
    $endpointValues = [ordered]@{}
    foreach ($e in @(ConvertFrom-JsonString $policyConfig.EndpointDlpGlobalSettings | ForEach-Object { $_ })) {
        if (Test-Blank $e) { continue }
        $setting = if ($e.PSObject.Properties['Setting']) { [string]$e.Setting } else { '' }
        $value = if ($setting) { Format-Val (ConvertFrom-JsonString $e.Value) } else { Format-Val $e }
        if (-not $value) { continue }
        if (-not $endpointValues.Contains($setting)) { $endpointValues[$setting] = [System.Collections.Generic.List[string]]::new() }
        $endpointValues[$setting].Add($value)
    }
    foreach ($setting in $endpointValues.Keys) {
        Add-Row $endpointSheet ([ordered]@{ 'Section' = 'Endpoint DLP'; 'Setting' = Get-FriendlyName $setting; 'Value' = ($endpointValues[$setting] -join "`n") })
    }
    foreach ($prop in $policyConfig.PSObject.Properties) {
        if ($prop.Name -in 'EndpointDlpGlobalSettings', 'EndpointDlpGlobalSettingsPsws' -or $prop.Name -in $commonNoise -or (Test-Blank $prop.Value)) { continue }
        $value = Format-Val (ConvertFrom-JsonString $prop.Value)
        if ($value) { Add-Row $endpointSheet ([ordered]@{ 'Section' = 'Policy configuration'; 'Setting' = Get-FriendlyName $prop.Name; 'Value' = $value }) }
    }
    $sheets += $endpointSheet
}

$spec = @{ Path = $OutputPath; Sheets = $sheets }
$useImportExcel = $Engine -eq 'ImportExcel' -or ($Engine -eq 'Auto' -and (Get-Module -ListAvailable -Name ImportExcel))
if ($useImportExcel) { Write-WorkbookImportExcel $spec } else { Write-WorkbookPython $spec }
Write-Host "Workbook written to $OutputPath"
