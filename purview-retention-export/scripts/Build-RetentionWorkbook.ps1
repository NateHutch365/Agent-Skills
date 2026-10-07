<#
.SYNOPSIS
    Builds a filterable Excel workbook from a retention export folder created by Export-RetentionConfiguration.ps1.
.DESCRIPTION
    Works offline from RetentionPolicies.raw.json, RetentionRules.raw.json and RetentionLabels.raw.json (plus app retention
    policies, adaptive scopes, event types, file plan descriptors, MRM policies/tags and ExportInfo.json when present).
    Uses the ImportExcel module if installed, otherwise Python with openpyxl.
.EXAMPLE
    .\Build-RetentionWorkbook.ps1 -ExportFolder .\Retention-Export-Contoso-20261007-1000
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
$locationText = @{
    Exchange     = 'Exchange mailboxes'
    SharePoint   = 'SharePoint sites'
    OneDrive     = 'OneDrive accounts'
    ModernGroup  = 'Microsoft 365 Group mailboxes & sites'
    Skype        = 'Skype for Business'
    PublicFolder = 'Exchange public folders'
    TeamsChannel = 'Teams channel messages'
    TeamsChat    = 'Teams chats'
}

$basisText = @{
    CreationAgeInDays     = 'When items were created'
    ModificationAgeInDays = 'When items were last modified'
    TaggedAgeInDays       = 'When items were labelled'
    EventAgeInDays        = 'When an event occurs'
}

$filePlanColumns = [ordered]@{
    ReferenceId = 'Reference Id'
    Department  = 'Business Function/Department'
    Category    = 'Category'
    SubCategory = 'Sub Category'
    Authority   = 'Authority Type'
    Citation    = 'Provision/Citation'
}

$commonNoise = 'Identity', 'Id', 'Guid', 'ImmutableId', 'ExchangeObjectId', 'ObjectVersion', 'DirectoryObjectVersion', 'IsValid',
    'ExchangeVersion', 'DistinguishedName', 'ObjectCategory', 'ObjectClass', 'WhenChanged', 'WhenCreated', 'OrganizationalUnitRoot',
    'OrganizationId', 'OriginatingServer', 'ObjectState', 'ReadOnly', 'ExternalIdentity', 'ErrorMetadata', 'Workload',
    'RunspaceId', 'PSComputerName', 'PSShowComputerName'

$policyNoise = $commonNoise + @('Type', 'DistributionStatus', 'DistributionSyncStatus', 'DistributionResults', 'LastStatusUpdateTime',
    'ModificationTimeUtc', 'CreationTimeUtc', 'PolicySource', 'Rules', 'ItemsMatchedDuringTestMode', 'CompletedLocations', 'FailedLocations',
    'ExpectedLocations', 'SimulationPolicyVersion', 'VersionCleanupState', 'Priority', 'IsSimulation', 'IsAdaptivePolicy', 'HasRules',
    'TestModeStatus', 'TestModeVerdict', 'ItemStatistics', 'ForceValidate', 'PolicyRulesMetaData', 'GlobalListType', 'PolicyRBACScopes',
    'RetentionRuleTypes', 'Locations', 'PolicyTemplateInfo', 'PolicyApprovalInfo', 'PolicyConstraints', 'UserAdministrativeUnitMembershipMap')

$policyHandled = 'Name', 'DisplayName', 'Enabled', 'Mode', 'Comment', 'CreatedBy', 'LastModifiedBy', 'WhenCreatedUTC', 'WhenChangedUTC',
    'RestrictiveRetention', 'Applications', 'AdaptiveScopeLocation', 'AdaptiveScopeLocationException'

$ruleNoise = $commonNoise + @('Policy', 'ParentPolicyName', 'Priority', 'Mode', 'DistributionStatus', 'ComplianceTagProperty')

$ruleHandled = 'Name', 'DisplayName', 'RetentionDuration', 'RetentionComplianceAction', 'ExpirationDateOption', 'RetentionDurationDisplayHint',
    'ApplyComplianceTag', 'PublishComplianceTag', 'ContentMatchQuery', 'ContentContainsSensitiveInformation', 'MachineLearningModelIDs',
    'ExcludedItemClasses', 'Comment', 'Disabled', 'CreatedBy', 'LastModifiedBy', 'WhenCreatedUTC', 'WhenChangedUTC'

$labelNoise = $commonNoise + @('Policy', 'Mode', 'HasRetentionAction', 'VersionCleanupState', 'ComplianceTagType', 'Priority', 'IsReviewTag', 'IsActiveLabel')

$labelHandled = 'Name', 'DisplayName', 'Comment', 'Notes', 'IsRecordLabel', 'Regulatory', 'RetentionAction', 'RetentionDuration', 'RetentionType',
    'RetentionDurationDisplayHint', 'EventType', 'ReviewerEmail', 'MultiStageReviewProperty', 'MultiStageReviewerMetadata', 'FilePlanMetadata', 'IsRecordUnlockedAsDefault',
    'ComplianceTagForNextStage', 'CreatedBy', 'LastModifiedBy', 'WhenCreatedUTC', 'WhenChangedUTC'
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

# Enum values may arrive as plain strings or as {"Value":"Text"} objects.
function Get-Text($v) {
    if ($null -eq $v) { return '' }
    if ($v -is [string] -or $v -is [ValueType]) { return [string]$v }
    if ($v.PSObject.Properties['Value']) { return [string]$v.Value }
    [string]$v
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

function ConvertTo-CellValue($v) {
    if ($null -eq $v) { return '' }
    if ($v -is [int] -or $v -is [long] -or $v -is [double]) { return $v }
    $s = [string]$v
    if ($s.Length -gt 32000) { $s = $s.Substring(0, 32000) + ' ...[truncated - see raw JSON]' }
    # Prevent tenant-sourced text being interpreted as a formula.
    if ($s -match '^[=+@-]') { $s = ' ' + $s }
    $s
}

function Get-DisplayName($Obj) {
    if ($Obj.DisplayName -and [string]$Obj.DisplayName -notmatch $guidPattern) { return [string]$Obj.DisplayName }
    [string]$Obj.Name
}

function Get-Extras($Obj, $Handled, $Noise, [string]$Prefix) {
    @(foreach ($prop in $Obj.PSObject.Properties) {
            if ($prop.Name -in $Handled -or $prop.Name -in $Noise -or (Test-Blank $prop.Value)) { continue }
            if ($prop.Name -match '^\w+Location(Exception)?$') { continue }
            $v = Format-Val (ConvertFrom-JsonString $prop.Value)
            if ($v) { "$Prefix$(Get-FriendlyName $prop.Name): $v" }
        })
}
#endregion

#region Retention formatting
function Format-Days([int]$Days, [string]$Hint) {
    $unit = $null
    if ($Hint -eq 'Years' -or (-not $Hint -and $Days -ge 365 -and $Days % 365 -eq 0)) { if ($Days % 365 -eq 0) { $unit = 'year'; $n = $Days / 365 } }
    elseif ($Hint -eq 'Months' -and $Days % 30 -eq 0) { $unit = 'month'; $n = $Days / 30 }
    if (-not $unit) { return "$Days days" }
    "$n $unit$(if ($n -ne 1) { 's' }) ($Days days)"
}

function Format-Duration($Duration, $Hint) {
    $d = Get-Text $Duration
    if (-not $d) { return '' }
    if ($d -eq 'Unlimited') { return 'Forever' }
    if ($d -notmatch '^\d+$') { return $d }
    Format-Days ([int]$d) (Get-Text $Hint)
}

function Get-RetentionInfo([string]$Action, $Duration, $Hint, [string]$Basis, [string]$EndOverride) {
    $period = Format-Duration $Duration $Hint
    $basisLabel = if ($basisText.ContainsKey($Basis)) { $basisText[$Basis] } elseif ($Basis) { Get-FriendlyName $Basis } else { '' }
    $end = ''; $summary = ''
    switch ($Action) {
        'Keep' {
            if ($period -eq 'Forever') { $summary = 'Retain forever'; $end = 'n/a' }
            else { $end = $(if ($EndOverride) { $EndOverride } else { 'Do nothing' }); $summary = "Retain for $period, then $($end.ToLower())" }
        }
        'KeepAndDelete' {
            $end = $(if ($EndOverride) { $EndOverride } else { 'Delete' })
            $summary = "Retain for $period, then $($end.Substring(0, 1).ToLower() + $end.Substring(1))"
        }
        'Delete' { $end = 'Delete'; $summary = "Delete items older than $period" }
        '' { $summary = 'No retention settings (classification only)'; $basisLabel = '' }
        default { $summary = "$(Get-FriendlyName $Action) $period".Trim() }
    }
    if ($basisLabel -and $period -ne 'Forever') { $summary += " (based on: $($basisLabel.ToLower()))" }
    [pscustomobject]@{ Summary = $summary; Period = $period; Basis = $(if ($period -eq 'Forever') { '' } else { $basisLabel }); End = $end }
}

function Format-Location($Item) {
    $Item = ConvertFrom-JsonString $Item
    if ($Item -is [string]) { return $Item }
    $name = Get-Text $Item.Name
    if ($name -eq 'All' -or (Get-Text $Item.Type) -eq 'Tenant') { return 'All' }
    if ($name -match '^https?://') { return $name }
    foreach ($n in 'DisplayName', 'PrimarySmtpAddress', 'Name', 'Identity') {
        $val = Get-Text $Item.$n
        if ($val -and $val -notmatch $guidPattern) { return $val }
    }
    $name
}

function Get-PolicyLocations($Policy) {
    $in = [System.Collections.Generic.List[string]]::new()
    $out = [System.Collections.Generic.List[string]]::new()
    foreach ($prop in $Policy.PSObject.Properties) {
        if ($prop.Name -like 'AdaptiveScope*' -or $prop.Name -notmatch '^(?<w>\w+?)Location(?<ex>Exception)?$') { continue }
        $w = $Matches['w']; $isEx = [bool]$Matches['ex']
        $items = @(foreach ($i in @($prop.Value)) { if (-not (Test-Blank $i)) { Format-Location $i } })
        if (-not $items.Count) { continue }
        $label = if ($locationText.ContainsKey($w)) { $locationText[$w] } else { Get-FriendlyName $w }
        $line = "${label}: $($items -join '; ')"
        if ($isEx) { $out.Add($line) } else { $in.Add($line) }
    }
    foreach ($a in @($Policy.Applications)) { if (-not (Test-Blank $a)) { $in.Add("Apps: $(Get-Text $a)") } }
    $scopes = @(foreach ($s in @($Policy.AdaptiveScopeLocation)) { if (-not (Test-Blank $s)) { Format-Location $s } })
    $scopesOut = @(foreach ($s in @($Policy.AdaptiveScopeLocationException)) { if (-not (Test-Blank $s)) { Format-Location $s } })
    [pscustomobject]@{ In = $in -join "`n"; Out = $out -join "`n"; Scopes = $scopes; ScopesOut = $scopesOut }
}

function Get-PolicyStatus($Policy) {
    if ($Policy.Enabled -eq $false) { return 'Off' }
    $mode = Get-Text $Policy.Mode
    switch ($mode) {
        '' { 'On' }
        'Enforce' { 'On' }
        'Enable' { 'On' }
        'Disable' { 'Off' }
        'PendingDeletion' { 'Pending deletion' }
        { $_ -like 'Simulat*' -or $_ -like 'Test*' } { 'Simulation' }
        default { Get-FriendlyName $mode }
    }
}

# Label references may be names, GUIDs or a comma-separated string of names.
function Resolve-LabelNames($v) {
    $out = [System.Collections.Generic.List[string]]::new()
    foreach ($i in @(ConvertFrom-JsonString $v)) {
        if (Test-Blank $i) { continue }
        if ($i -isnot [string]) { $i = if ($i.Name) { [string]$i.Name } else { Get-Text $i } }
        if ($labelNameMap.ContainsKey($i)) { $out.Add($labelNameMap[$i]); continue }
        foreach ($part in $i -split ',') {
            $p = $part.Trim()
            if ($p) { $out.Add($(if ($labelNameMap.ContainsKey($p)) { $labelNameMap[$p] } else { $p })) }
        }
    }
    @($out | Select-Object -Unique)
}

function Get-SitList($Node) {
    $Node = ConvertFrom-JsonString $Node
    if ($null -eq $Node) { return }
    if ($Node -is [string]) { if ($Node.Trim()) { $Node }; return }
    if ($Node -is [ValueType]) { return }
    if ($Node -is [System.Collections.IEnumerable]) { foreach ($i in $Node) { Get-SitList $i }; return }
    $p = $Node.PSObject.Properties
    if ($p['sensitivetypes']) { Get-SitList $Node.sensitivetypes; return }
    if ($p['groups']) { Get-SitList $Node.groups; return }
    if ($p['name']) {
        $extra = @()
        if ($Node.confidencelevel) { $extra += [string]$Node.confidencelevel }
        $min = [string]$Node.mincount; $max = [string]$Node.maxcount
        if ($min -match '^\d+$') { $extra += $(if ($max -match '^\d+$' -and $max -ne '-1') { "$min-$max" } else { "$min+" }) }
        return "$($Node.name)$(if ($extra.Count) { " ($($extra -join ', '))" })"
    }
    Format-Val $Node
}

function Format-Conditions($Rule) {
    $parts = @()
    $sits = @(Get-SitList $Rule.ContentContainsSensitiveInformation)
    if ($sits.Count) { $parts += "Sensitive info types: $($sits -join '; ')" }
    if (-not (Test-Blank $Rule.ContentMatchQuery)) { $parts += "Keyword query (KQL): $($Rule.ContentMatchQuery)" }
    $ml = @($Rule.MachineLearningModelIDs | Where-Object { -not (Test-Blank $_) })
    if ($ml.Count) { $parts += "Trainable classifiers (IDs): $($ml -join '; ')" }
    $parts -join "`n"
}

function Format-FilterQuery($n) {
    $n = ConvertFrom-JsonString $n
    if ($null -eq $n) { return '' }
    if ($n -is [string] -or $n -is [ValueType]) { return [string]$n }
    if ($n -is [System.Collections.IEnumerable]) { return (@(foreach ($i in $n) { Format-FilterQuery $i }) -join '; ') }
    $p = $n.PSObject.Properties
    if ($p['Conditions']) {
        $conj = if ($p['Conjunction'] -and (Get-Text $n.Conjunction)) { " $((Get-Text $n.Conjunction).ToUpper()) " } else { ' AND ' }
        $parts = @(foreach ($c in @($n.Conditions)) {
                $t = Format-FilterQuery $c
                if ($c.PSObject.Properties['Conditions'] -and @($c.Conditions).Count -gt 1) { "($t)" } else { $t }
            })
        return $parts -join $conj
    }
    if ($p['Name'] -and $p['Operator']) { return "$($n.Name) $(Get-Text $n.Operator) '$(Format-Val $n.Value)'" }
    Format-Val $n
}

function Get-MrmTagKind($Tag) {
    $type = Get-Text $Tag.Type
    switch ($type) { 'All' { 'Default policy tag' } 'Personal' { 'Personal tag' } default { "$(Get-FriendlyName $type) folder tag" } }
}

function Get-MrmTagLinkName($v) {
    if ($v -isnot [string]) { $v = if ($v.Name) { [string]$v.Name } else { Get-Text $v } }
    if ($v -match '^CN=([^,]+)') { return $Matches[1] }
    if ($mrmTagById.ContainsKey($v)) { return $mrmTagById[$v] }
    $v
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
    $tmp = Join-Path ([IO.Path]::GetTempPath()) "retention-workbook-$([guid]::NewGuid())"
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
$policies = @(Read-ExportJson 'RetentionPolicies.raw.json' | Where-Object { $_ })
$rules = @(Read-ExportJson 'RetentionRules.raw.json' | Where-Object { $_ })
$appPolicies = @(Read-ExportJson 'AppRetentionPolicies.raw.json' | Where-Object { $_ })
$rules += @(Read-ExportJson 'AppRetentionRules.raw.json' | Where-Object { $_ })
$labels = @(Read-ExportJson 'RetentionLabels.raw.json' | Where-Object { $_ })
$scopes = @(Read-ExportJson 'AdaptiveScopes.raw.json' | Where-Object { $_ })
$eventTypes = @(Read-ExportJson 'RetentionEventTypes.raw.json' | Where-Object { $_ })
$filePlan = @(Read-ExportJson 'FilePlanProperties.raw.json' | Where-Object { $_ })
$mrmPolicies = @(Read-ExportJson 'MrmRetentionPolicies.raw.json' | Where-Object { $_ })
$mrmTags = @(Read-ExportJson 'MrmRetentionTags.raw.json' | Where-Object { $_ })
$mrmCounts = @(Read-ExportJson 'MrmMailboxCounts.json' | Where-Object { $_ })
$exportInfo = Read-ExportJson 'ExportInfo.json'

if (-not (Test-Path (Join-Path $ExportFolder 'RetentionPolicies.raw.json')) -and -not $mrmPolicies.Count) {
    throw "No retention export files (RetentionPolicies.raw.json or MrmRetentionPolicies.raw.json) found in $ExportFolder"
}
foreach ($p in $policies) { $p | Add-Member -NotePropertyName PolicySource -NotePropertyValue 'Retention' -Force }
foreach ($p in $appPolicies) { $p | Add-Member -NotePropertyName PolicySource -NotePropertyValue 'App retention' -Force }
$allPolicies = @($policies) + @($appPolicies)

$labelNameMap = @{}
foreach ($l in $labels) {
    $n = Get-DisplayName $l
    foreach ($k in $l.Name, $l.Guid, $l.ImmutableId, (Get-Text $l.Identity)) { if ($k) { $labelNameMap[[string]$k] = $n } }
}

$eventTypeMap = @{}
foreach ($e in $eventTypes) { foreach ($k in $e.Name, $e.Guid, (Get-Text $e.Identity)) { if ($k) { $eventTypeMap[[string]$k] = [string]$e.Name } } }

$mrmTagById = @{}
foreach ($t in $mrmTags) { foreach ($k in $t.Guid, (Get-Text $t.Identity), $t.RetentionId) { if ($k) { $mrmTagById[[string]$k] = [string]$t.Name } } }

# Link rules to policies by any identifier the rule carries.
$policyByKey = @{}
foreach ($p in $allPolicies) {
    foreach ($k in $p.Guid, (Get-Text $p.Identity), $p.Name, $p.ExchangeObjectId, $p.ImmutableId) { if ($k) { $policyByKey[[string]$k] = $p } }
}
$rulesByPolicy = @{}
$orphanRules = 0
foreach ($r in $rules) {
    $owner = $null
    foreach ($k in (Get-Text $r.Policy), (Get-Text $r.ParentPolicyName)) { if ($k -and $policyByKey.ContainsKey($k)) { $owner = $policyByKey[$k]; break } }
    if (-not $owner) { $orphanRules++; continue }
    $key = [string]$owner.Guid + '|' + $owner.Name
    if (-not $rulesByPolicy.ContainsKey($key)) { $rulesByPolicy[$key] = [System.Collections.Generic.List[object]]::new() }
    $rulesByPolicy[$key].Add($r)
}
if ($orphanRules) { Write-Warning "$orphanRules rule(s) could not be linked to a policy and are not shown." }

$tenantSource = @($allPolicies + $labels + $mrmPolicies | Where-Object { $_.OrganizationalUnitRoot } | Select-Object -First 1)
$tenant = if ($tenantSource.Count) { ([string]$tenantSource[0].OrganizationalUnitRoot -split '/')[-1] }
          elseif ($exportInfo -and $exportInfo.ExportedBy) { ([string]$exportInfo.ExportedBy -split '@')[-1] } else { 'Unknown' }
if (-not $ClientName) {
    $ClientName = if ($exportInfo -and $exportInfo.ClientName) { [string]$exportInfo.ClientName } else { (Get-Culture).TextInfo.ToTitleCase(($tenant -split '\.')[0]) }
}
$snapshotFile = @('RetentionPolicies.raw.json', 'MrmRetentionPolicies.raw.json' | ForEach-Object { Join-Path $ExportFolder $_ } | Where-Object { Test-Path $_ })[0]
$snapshot = if ($exportInfo -and $exportInfo.ExportedOnUTC) { Format-Date $exportInfo.ExportedOnUTC }
            else { (Get-Item $snapshotFile).LastWriteTimeUtc.ToString('yyyy-MM-dd HH:mm') }
if (-not $OutputPath) {
    $OutputPath = Join-Path $ExportFolder ('{0}-Retention-Configuration-{1}.xlsx' -f ($ClientName -replace '[^\w-]', ''), ($snapshot.Substring(0, 10) -replace '-', ''))
}
$OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
#endregion

#region Build sheets
$statusHighlight = [ordered]@{ 'On' = '#C6EFCE'; 'Off' = '#D9D9D9'; 'Simulation' = '#FFEB9C'; 'Pending' = '#FFC7CE' }

$retentionSheet = New-Sheet 'Retention Policies' ([ordered]@{
        'Policy' = 40; 'Status' = 11; 'Policy Type' = 18; 'Scope Type' = 10; 'Locations' = 50; 'Excluded' = 35; 'Adaptive Scopes' = 30
        'Retention Settings' = 50; 'Retention Period' = 20; 'Based On' = 24; 'At End Of Period' = 16; 'Conditions' = 35
        'Preservation Lock' = 11; 'Description' = 50; 'Additional Settings' = 35; 'Created By' = 22; 'Created (UTC)' = 17
        'Last Modified By' = 22; 'Last Modified (UTC)' = 17
    }) 1 ([ordered]@{ 'Status' = $statusHighlight; 'Preservation Lock' = [ordered]@{ 'Yes' = '#F8CBAD' } })

$labelPolicySheet = New-Sheet 'Label Policies' ([ordered]@{
        'Policy' = 40; 'Policy Kind' = 13; 'Status' = 11; 'Labels' = 40; 'Scope Type' = 10; 'Locations' = 50; 'Excluded' = 35
        'Adaptive Scopes' = 30; 'Auto-apply Conditions' = 55; 'Preservation Lock' = 11; 'Description' = 50; 'Additional Settings' = 35
        'Created By' = 22; 'Created (UTC)' = 17; 'Last Modified By' = 22; 'Last Modified (UTC)' = 17
    }) 1 ([ordered]@{ 'Status' = $statusHighlight; 'Preservation Lock' = [ordered]@{ 'Yes' = '#F8CBAD' } })

$publishedIn = @{}
$autoAppliedBy = @{}
$scopeUsedBy = @{}
$retentionCount = 0; $publishCount = 0; $autoCount = 0

foreach ($p in @($allPolicies | Sort-Object { Get-DisplayName $_ })) {
    $name = Get-DisplayName $p
    $key = [string]$p.Guid + '|' + $p.Name
    $pRules = if ($rulesByPolicy.ContainsKey($key)) { @($rulesByPolicy[$key]) } else { @() }
    $loc = Get-PolicyLocations $p
    foreach ($s in @($loc.Scopes) + @($loc.ScopesOut)) {
        if (-not $scopeUsedBy.ContainsKey($s)) { $scopeUsedBy[$s] = [System.Collections.Generic.List[string]]::new() }
        if (-not $scopeUsedBy[$s].Contains($name)) { $scopeUsedBy[$s].Add($name) }
    }
    $scopeText = @($loc.Scopes) + @(foreach ($s in $loc.ScopesOut) { "Excluded: $s" })
    $extras = @(Get-Extras $p $policyHandled $policyNoise '')
    foreach ($r in $pRules) {
        $prefix = if ($pRules.Count -gt 1) { "[$(Get-DisplayName $r)] " } else { '' }
        $extras += @(Get-Extras $r $ruleHandled $ruleNoise $prefix)
        if ($r.Disabled -eq $true) { $extras += "${prefix}Rule disabled" }
        $excl = @($r.ExcludedItemClasses | Where-Object { -not (Test-Blank $_) })
        if ($excl.Count) { $extras += "${prefix}Excluded item types: $($excl -join '; ')" }
    }
    $common = [ordered]@{
        'Policy' = $name; 'Status' = Get-PolicyStatus $p
        'Scope Type' = $(if ($loc.Scopes.Count -or $loc.ScopesOut.Count) { 'Adaptive' } else { 'Static' })
        'Locations' = $loc.In; 'Excluded' = $loc.Out; 'Adaptive Scopes' = $scopeText -join "`n"
        'Preservation Lock' = $(if ($p.RestrictiveRetention -eq $true) { 'Yes' } else { 'No' })
        'Description' = [string]$p.Comment; 'Additional Settings' = $extras -join "`n"
        'Created By' = [string]$p.CreatedBy; 'Created (UTC)' = Format-Date $p.WhenCreatedUTC
        'Last Modified By' = [string]$p.LastModifiedBy; 'Last Modified (UTC)' = Format-Date $p.WhenChangedUTC
    }

    $publishRules = @($pRules | Where-Object { -not (Test-Blank $_.PublishComplianceTag) })
    $applyRules = @($pRules | Where-Object { -not (Test-Blank $_.ApplyComplianceTag) })

    if ($publishRules.Count -or $applyRules.Count) {
        $kind = if ($applyRules.Count) { 'Auto-apply' } else { 'Publish' }
        if ($applyRules.Count) { $autoCount++ } else { $publishCount++ }
        $policyLabels = [System.Collections.Generic.List[string]]::new()
        foreach ($r in $publishRules) { foreach ($l in (Resolve-LabelNames $r.PublishComplianceTag)) { if (-not $policyLabels.Contains($l)) { $policyLabels.Add($l) } } }
        foreach ($r in $applyRules) { foreach ($l in (Resolve-LabelNames $r.ApplyComplianceTag)) { if (-not $policyLabels.Contains($l)) { $policyLabels.Add($l) } } }
        $target = if ($applyRules.Count) { $autoAppliedBy } else { $publishedIn }
        foreach ($l in $policyLabels) {
            if (-not $target.ContainsKey($l)) { $target[$l] = [System.Collections.Generic.List[string]]::new() }
            $target[$l].Add($name)
        }
        $conditions = @(foreach ($r in $applyRules) {
                $c = Format-Conditions $r
                if ($c) { if ($applyRules.Count -gt 1) { "[$(Get-DisplayName $r)]`n$c" } else { $c } }
            })
        $row = [ordered]@{ 'Policy Kind' = $kind; 'Labels' = $policyLabels -join "`n"; 'Auto-apply Conditions' = $conditions -join "`n" }
        foreach ($k in $common.Keys) { $row[$k] = $common[$k] }
        Add-Row $labelPolicySheet $row
        continue
    }

    $retentionCount++
    $infos = @(foreach ($r in $pRules) {
            Get-RetentionInfo (Get-Text $r.RetentionComplianceAction) $r.RetentionDuration $r.RetentionDurationDisplayHint (Get-Text $r.ExpirationDateOption) ''
        })
    $label = { param($i, $r) if ($pRules.Count -gt 1) { "[$(Get-DisplayName $r)] $i" } else { $i } }
    $row = [ordered]@{
        'Policy Type' = $(if ($p.PolicySource -eq 'App retention') { 'App retention policy' } else { 'Retention policy' })
        'Retention Settings' = $(if ($pRules.Count) { (@(for ($i = 0; $i -lt $pRules.Count; $i++) { & $label $infos[$i].Summary $pRules[$i] }) -join "`n") } else { '(no rules)' })
        'Retention Period' = (@($infos | ForEach-Object Period | Where-Object { $_ } | Select-Object -Unique) -join "`n")
        'Based On' = (@($infos | ForEach-Object Basis | Where-Object { $_ } | Select-Object -Unique) -join "`n")
        'At End Of Period' = (@($infos | ForEach-Object End | Where-Object { $_ } | Select-Object -Unique) -join "`n")
        'Conditions' = (@(for ($i = 0; $i -lt $pRules.Count; $i++) { $c = Format-Conditions $pRules[$i]; if ($c) { & $label $c $pRules[$i] } }) -join "`n")
    }
    foreach ($k in $common.Keys) { $row[$k] = $common[$k] }
    Add-Row $retentionSheet $row
}

# Retention labels
$labelColumns = [ordered]@{
    'Label' = 36; 'Label Type' = 17; 'Retention Settings' = 50; 'Retention Period' = 20; 'Based On' = 24; 'Event Type' = 22
    'At End Of Period' = 24; 'Disposition Review' = 35; 'Record Unlocked By Default' = 11; 'Published In' = 35; 'Auto-applied By' = 35
}
foreach ($c in $filePlanColumns.Values) { $labelColumns[$c] = 20 }
$labelColumns['Description'] = 50; $labelColumns['Notes'] = 40; $labelColumns['Additional Settings'] = 35
$labelColumns['Created By'] = 22; $labelColumns['Created (UTC)'] = 17; $labelColumns['Last Modified By'] = 22; $labelColumns['Last Modified (UTC)'] = 17
$labelSheet = New-Sheet 'Retention Labels' $labelColumns 1 ([ordered]@{
        'Label Type' = [ordered]@{ 'Regulatory' = '#FFC7CE'; 'Record' = '#F8CBAD' }
        'Published In' = [ordered]@{ '(not' = '#F2F2F2' }
    })
$recordCount = 0
foreach ($l in @($labels | Sort-Object { Get-DisplayName $_ })) {
    $name = Get-DisplayName $l
    $type = if ($l.Regulatory -eq $true) { 'Regulatory record' } elseif ($l.IsRecordLabel -eq $true) { 'Record' } else { 'Retention label' }
    if ($type -ne 'Retention label') { $recordCount++ }

    $reviewers = @($l.ReviewerEmail | Where-Object { -not (Test-Blank $_) })
    $stages = ConvertFrom-JsonString $l.MultiStageReviewProperty
    $review = @()
    if ($reviewers.Count) { $review += "Reviewers: $($reviewers -join '; ')" }
    if (-not (Test-Blank $stages)) { $review += "Stages: $(Format-Val $stages)" }
    $msr = ConvertFrom-JsonString $l.MultiStageReviewerMetadata
    if ($msr -and $msr -isnot [string] -and $msr.PSObject.Properties['MultiStageReviewSettings']) {
        $stageList = @($msr.MultiStageReviewSettings)
        foreach ($s in $stageList) {
            $who = @($s.Reviewers | Where-Object { -not (Test-Blank $_) }) -join '; '
            $review += $(if ($stageList.Count -gt 1) { "Stage $($s.StageName): $who" } else { "Reviewers: $who" })
        }
    }
    elseif (-not (Test-Blank $msr)) { $review += "Stages: $(Format-Val $msr)" }
    if (-not $review.Count -and $l.IsReviewTag -eq $true) { $review += 'Disposition review (reviewers not exported)' }
    $next = if (-not (Test-Blank $l.ComplianceTagForNextStage)) { (Resolve-LabelNames $l.ComplianceTagForNextStage) -join ', ' }
    $endOverride = if ($review.Count) { 'Start a disposition review' } elseif ($next) { "Relabel to $next" } else { '' }
    $info = Get-RetentionInfo (Get-Text $l.RetentionAction) $l.RetentionDuration $l.RetentionDurationDisplayHint (Get-Text $l.RetentionType) $endOverride

    $eventType = ''
    if (-not (Test-Blank $l.EventType)) {
        $e = Get-Text $l.EventType
        $eventType = if ($eventTypeMap.ContainsKey($e)) { $eventTypeMap[$e] } else { $e }
    }

    $extras = @(Get-Extras $l $labelHandled $labelNoise '')
    $row = [ordered]@{}
    $fp = ConvertFrom-JsonString $l.FilePlanMetadata
    if ($fp -and $fp -isnot [string] -and $fp.PSObject.Properties['Settings']) {
        foreach ($s in @($fp.Settings)) {
            if (-not $s -or (Test-Blank $s.Value)) { continue }
            $k = ([string]$s.Key) -replace '^FilePlanProperty'
            if ($filePlanColumns.Contains($k)) { $row[$filePlanColumns[$k]] = Format-Val $s.Value }
            else { $extras += "File plan $(Get-FriendlyName $k): $(Format-Val $s.Value)" }
        }
    }
    elseif (-not (Test-Blank $fp)) { $extras += "File plan: $(Format-Val $fp)" }

    $row['Label'] = $name; $row['Label Type'] = $type
    $row['Retention Settings'] = $info.Summary; $row['Retention Period'] = $info.Period; $row['Based On'] = $info.Basis
    $row['Event Type'] = $eventType; $row['At End Of Period'] = $info.End; $row['Disposition Review'] = $review -join "`n"
    $row['Record Unlocked By Default'] = $(if ($type -eq 'Retention label') { '' } elseif ($l.IsRecordUnlockedAsDefault -eq $true) { 'Yes' } else { 'No' })
    $row['Published In'] = $(if ($publishedIn.ContainsKey($name)) { $publishedIn[$name] -join "`n" } else { '(not published)' })
    $row['Auto-applied By'] = $(if ($autoAppliedBy.ContainsKey($name)) { $autoAppliedBy[$name] -join "`n" } else { '' })
    $row['Description'] = [string]$l.Comment; $row['Notes'] = [string]$l.Notes; $row['Additional Settings'] = $extras -join "`n"
    $row['Created By'] = [string]$l.CreatedBy; $row['Created (UTC)'] = Format-Date $l.WhenCreatedUTC
    $row['Last Modified By'] = [string]$l.LastModifiedBy; $row['Last Modified (UTC)'] = Format-Date $l.WhenChangedUTC
    Add-Row $labelSheet $row
}

$sheets = [System.Collections.Generic.List[object]]::new()

# Adaptive scopes
$scopeSheet = $null
if ($scopes.Count) {
    $scopeSheet = New-Sheet 'Adaptive Scopes' ([ordered]@{
            'Scope' = 35; 'Scope Type' = 12; 'Query' = 70; 'Administrative Units' = 25; 'Used By Policies' = 40; 'Description' = 40
            'Created By' = 22; 'Created (UTC)' = 17; 'Last Modified (UTC)' = 17
        }) 1 $null
    foreach ($s in @($scopes | Sort-Object Name)) {
        $n = Get-DisplayName $s
        $query = if (-not (Test-Blank $s.RawQuery)) { "Advanced query: $($s.RawQuery)" } else { Format-FilterQuery $s.FilterConditions }
        Add-Row $scopeSheet ([ordered]@{
                'Scope' = $n; 'Scope Type' = Get-Text $s.LocationType; 'Query' = $query
                'Administrative Units' = Format-Val $s.AdministrativeUnits
                'Used By Policies' = $(if ($scopeUsedBy.ContainsKey($n)) { $scopeUsedBy[$n] -join "`n" } elseif ($scopeUsedBy.ContainsKey([string]$s.Name)) { $scopeUsedBy[[string]$s.Name] -join "`n" } else { '' })
                'Description' = [string]$s.Comment; 'Created By' = [string]$s.CreatedBy; 'Created (UTC)' = Format-Date $s.WhenCreatedUTC
                'Last Modified (UTC)' = Format-Date $s.WhenChangedUTC
            })
    }
}

# MRM
$mrmPolicySheet = $null; $mrmTagSheet = $null
if ($mrmPolicies.Count -or $mrmTags.Count) {
    $countMap = @{}
    foreach ($c in $mrmCounts) { $countMap[[string]$c.RetentionPolicy] = $c }

    $tagUsedBy = @{}
    $mrmPolicySheet = New-Sheet 'MRM Retention Policies' ([ordered]@{
            'MRM Policy' = 40; 'Default Policy' = 10; 'Mailboxes Assigned' = 11; 'Mailboxes On Retention Hold' = 11; 'Tag Count' = 8
            'Retention Tags' = 70; 'Last Modified (UTC)' = 17
        }) 1 ([ordered]@{ 'Default Policy' = [ordered]@{ 'Yes' = '#C6EFCE' } })
    $tagByName = @{}
    foreach ($t in $mrmTags) { $tagByName[[string]$t.Name] = $t }
    foreach ($mp in @($mrmPolicies | Sort-Object @{ e = { -not $_.IsDefault } }, Name)) {
        $tagNames = @(foreach ($link in @($mp.RetentionPolicyTagLinks)) { if (-not (Test-Blank $link)) { Get-MrmTagLinkName $link } })
        foreach ($tn in $tagNames) {
            if (-not $tagUsedBy.ContainsKey($tn)) { $tagUsedBy[$tn] = [System.Collections.Generic.List[string]]::new() }
            $tagUsedBy[$tn].Add([string]$mp.Name)
        }
        $tagLines = @(foreach ($tn in $tagNames) {
                $t = $tagByName[$tn]
                if (-not $t) { $tn; continue }
                $age = if ($null -ne $t.AgeLimitDays -and "$($t.AgeLimitDays)" -ne '') { Format-Days ([int]$t.AgeLimitDays) '' } else { 'no age limit' }
                $state = if ($t.RetentionEnabled -eq $false) { ', disabled' } else { '' }
                "$tn ($(Get-MrmTagKind $t), $age, $(Get-FriendlyName (Get-Text $t.RetentionAction))$state)"
            })
        $c = $countMap[[string]$mp.Name]
        Add-Row $mrmPolicySheet ([ordered]@{
                'MRM Policy' = [string]$mp.Name; 'Default Policy' = $(if ($mp.IsDefault -eq $true) { 'Yes' } else { 'No' })
                'Mailboxes Assigned' = $(if ($mrmCounts.Count) { if ($c) { [int]$c.MailboxCount } else { 0 } } else { 'Not exported' })
                'Mailboxes On Retention Hold' = $(if ($mrmCounts.Count) { if ($c) { [int]$c.RetentionHoldCount } else { 0 } } else { 'Not exported' })
                'Tag Count' = $tagNames.Count; 'Retention Tags' = $tagLines -join "`n"; 'Last Modified (UTC)' = Format-Date $mp.WhenChangedUTC
            })
    }

    $mrmTagSheet = New-Sheet 'MRM Retention Tags' ([ordered]@{
            'Retention Tag' = 40; 'Tag Type' = 20; 'Applies To' = 22; 'Enabled' = 9; 'Retention Period' = 22; 'Retention Action' = 26
            'Message Class' = 16; 'Used By MRM Policies' = 40; 'System Tag' = 9; 'Comment' = 50; 'Last Modified (UTC)' = 17
        }) 1 ([ordered]@{ 'Enabled' = [ordered]@{ 'Yes' = '#C6EFCE'; 'No' = '#D9D9D9' }; 'Used By MRM Policies' = [ordered]@{ '(not' = '#F2F2F2' } })
    foreach ($t in @($mrmTags | Sort-Object @{ e = { switch (Get-Text $_.Type) { 'All' { 0 } 'Personal' { 2 } default { 1 } } } }, Name)) {
        $type = Get-Text $t.Type
        $kind = switch ($type) { 'All' { 'Default policy tag' } 'Personal' { 'Personal tag' } default { 'Retention policy tag' } }
        $appliesTo = switch ($type) {
            'All' { if ((Get-Text $t.MessageClass) -match 'voicemail') { 'Voicemail' } else { 'Entire mailbox (default)' } }
            'Personal' { 'Applied by users' }
            default { "$(Get-FriendlyName $type) folder" }
        }
        $age = if ($null -ne $t.AgeLimitDays -and "$($t.AgeLimitDays)" -ne '') { Format-Days ([int]$t.AgeLimitDays) '' } else { 'No age limit' }
        Add-Row $mrmTagSheet ([ordered]@{
                'Retention Tag' = [string]$t.Name; 'Tag Type' = $kind; 'Applies To' = $appliesTo
                'Enabled' = $(if ($t.RetentionEnabled -eq $false) { 'No' } else { 'Yes' }); 'Retention Period' = $age
                'Retention Action' = Get-FriendlyName (Get-Text $t.RetentionAction)
                'Message Class' = $(if ((Get-Text $t.MessageClass) -in '*', '') { 'All' } else { Get-Text $t.MessageClassDisplayName })
                'Used By MRM Policies' = $(if ($tagUsedBy.ContainsKey([string]$t.Name)) { $tagUsedBy[[string]$t.Name] -join "`n" } else { '(not linked)' })
                'System Tag' = $(if ($t.SystemTag -eq $true) { 'Yes' } else { 'No' }); 'Comment' = [string]$t.Comment
                'Last Modified (UTC)' = Format-Date $t.WhenChangedUTC
            })
    }
}

# Reference data: event types and file plan descriptors
$refSheet = $null
if ($eventTypes.Count -or $filePlan.Count) {
    $refSheet = New-Sheet 'Reference Data' ([ordered]@{ 'Type' = 28; 'Name' = 45; 'Details' = 60 }) 2 $null
    foreach ($e in @($eventTypes | Sort-Object Name)) { Add-Row $refSheet ([ordered]@{ 'Type' = 'Retention event type'; 'Name' = [string]$e.Name; 'Details' = [string]$e.Comment }) }
    foreach ($f in @($filePlan | Sort-Object FilePlanType, Name)) {
        $t = [string]$f.FilePlanType
        $typeLabel = if ($filePlanColumns.Contains($t)) { "File plan - $($filePlanColumns[$t])" } else { "File plan - $(Get-FriendlyName $t)" }
        Add-Row $refSheet ([ordered]@{ 'Type' = $typeLabel; 'Name' = Get-DisplayName $f; 'Details' = [string]$f.Comment })
    }
}

# Cover
$cover = @{ Name = 'Cover'; Kind = 'Cover'; Title = "$ClientName - Microsoft Purview Data Lifecycle Management Configuration"; Rows = [System.Collections.Generic.List[object]]::new() }
$noMrm = @($mrmCounts | Where-Object RetentionPolicy -eq '(none)' | Select-Object -First 1)
$coverRows = @(
    , @('Client', $ClientName)
    , @('Tenant', $tenant)
    , @('Snapshot taken (UTC)', $snapshot)
    , @('Exported by', $(if ($exportInfo) { [string]$exportInfo.ExportedBy } else { 'Not recorded' }))
    , @('Workbook generated (UTC)', (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm'))
    , @('Retention policies', [string]$retentionCount)
    , @('Retention labels', ('{0} ({1} record / regulatory record)' -f $labels.Count, $recordCount))
    , @('Retention label policies', ('{0} publish, {1} auto-apply' -f $publishCount, $autoCount))
    , @('Adaptive scopes', [string]$scopes.Count)
    , @('MRM retention policies / tags', $(if ($mrmPolicies.Count -or $mrmTags.Count) { "$($mrmPolicies.Count) / $($mrmTags.Count)" } else { 'Not exported' }))
    if ($noMrm.Count) { , @('Mailboxes with no MRM policy', [string]$noMrm[0].MailboxCount) }
    , @('', '')
    , @('Sheet: Retention Policies', 'Purview retention policies (including app retention policies for Teams private channels, Viva Engage and Copilot): locations, exclusions, adaptive scopes, retention period and end-of-period action.')
    , @('Sheet: Retention Labels', 'Retention labels: label type (retention / record / regulatory record), retention settings, event type, disposition review, file plan descriptors and the policies that publish or auto-apply each label.')
    , @('Sheet: Label Policies', 'Retention label policies: published label policies (users apply labels) and auto-apply policies with their conditions.')
    if ($scopeSheet) { , @('Sheet: Adaptive Scopes', 'Adaptive scope queries and the policies that use them.') }
    if ($mrmPolicySheet) { , @('Sheet: MRM Retention Policies', 'Exchange Messaging Records Management (MRM) retention policies, linked retention tags and mailbox assignment counts (when exported).') }
    if ($mrmTagSheet) { , @('Sheet: MRM Retention Tags', 'Exchange MRM retention tags: default policy tags, retention policy (folder) tags and personal tags.') }
    if ($refSheet) { , @('Sheet: Reference Data', 'Retention event types and file plan descriptor values defined in the tenant.') }
    , @('', '')
    , @('Note', 'This is a point-in-time snapshot of the configuration as exported from the tenant. Changes made after the snapshot date are not reflected.')
    , @('Note', 'Preservation Lock = Yes means the policy is locked and cannot be turned off, deleted or made less restrictive.')
    , @('Note', 'Where Purview retention and MRM retention both apply to a mailbox item, the principles of retention decide the outcome: retention wins over deletion, and the longest retention period wins.')
    , @('Note', 'Not included: retention labels applied by default to SharePoint libraries/folders or Outlook folders, labels applied to individual items, retention events, and disposition review history. Trainable classifiers are shown by ID.')
)
foreach ($c in $coverRows) { $cover.Rows.Add([object[]]@((ConvertTo-CellValue $c[0]), (ConvertTo-CellValue $c[1]))) }

$sheets.Add($cover); $sheets.Add($retentionSheet); $sheets.Add($labelSheet); $sheets.Add($labelPolicySheet)
foreach ($s in $scopeSheet, $mrmPolicySheet, $mrmTagSheet, $refSheet) { if ($s) { $sheets.Add($s) } }
#endregion

$spec = @{ Path = $OutputPath; Sheets = @($sheets) }
$useImportExcel = $Engine -eq 'ImportExcel' -or ($Engine -eq 'Auto' -and (Get-Module -ListAvailable -Name ImportExcel))
if ($useImportExcel) { Write-WorkbookImportExcel $spec } else { Write-WorkbookPython $spec }
Write-Host "Workbook written to $OutputPath"
