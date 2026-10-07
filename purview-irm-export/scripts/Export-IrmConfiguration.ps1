#Requires -Modules ExchangeOnlineManagement
<#
.SYNOPSIS
    Exports Microsoft Purview Insider Risk Management (IRM) policies, global settings, entity lists and IRM role groups
    from a tenant and builds a filterable Excel workbook (via Build-IrmWorkbook.ps1). Outputs the export folder path.
.DESCRIPTION
    Read-only: only Get-* cmdlets are used. The IRM cmdlets are undocumented by Microsoft and may change without notice.
.EXAMPLE
    .\Export-IrmConfiguration.ps1 -UserPrincipalName admin@contoso.onmicrosoft.com -ClientName Contoso -OutputRoot C:\Dev\Contoso
#>
[CmdletBinding()]
param(
    [string]$UserPrincipalName,
    [Parameter(Mandatory)][string]$ClientName,
    [string]$OutputRoot = $PWD,
    [switch]$SkipConnect,
    [switch]$SkipWorkbook
)

$ErrorActionPreference = 'Stop'
if (-not $SkipConnect -and -not $UserPrincipalName) { throw 'Provide -UserPrincipalName, or -SkipConnect if a Security & Compliance session is already open.' }

$safeClient = $ClientName -replace '[^\w-]', ''
$OutputFolder = Join-Path $OutputRoot "IRM-Export-$safeClient-$(Get-Date -Format 'yyyyMMdd-HHmm')"

# Entity list types enumerated by Microsoft365DSC; Get-InsiderRiskEntityList requires -Type or -Identity.
$listTypes = 'HveLists', 'DomainLists', 'CriticalAssetLists', 'WindowsFilePathRegexLists', 'SensitiveTypeLists', 'SiteLists',
    'KeywordLists', 'CustomDomainLists', 'CustomSiteLists', 'CustomKeywordLists', 'CustomFileTypeLists', 'CustomFilePathRegexLists',
    'CustomSensitiveInformationTypeLists', 'CustomMLClassifierTypeLists', 'GlobalExclusionSGMapping', 'DlpPolicyLists'

function Save-Json($Data, [string]$Name) {
    ConvertTo-Json -InputObject @($Data) -Depth 20 | Set-Content (Join-Path $OutputFolder $Name) -Encoding UTF8
}

if (-not $SkipConnect) { Connect-IPPSSession -UserPrincipalName $UserPrincipalName }
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null

Write-Host 'Retrieving insider risk policies and global settings...'
$policies = @(Get-InsiderRiskPolicy)
Save-Json $policies 'IrmPolicies.raw.json'

Write-Host 'Retrieving insider risk entity lists...'
$lists = foreach ($type in $listTypes) {
    try { Get-InsiderRiskEntityList -Type $type -ErrorAction Stop }
    catch { Write-Warning "Entity list type '$type' could not be retrieved: $($_.Exception.Message)" }
}
# Listing by -Type returns EntitiesCount but not the entries themselves.
$lists = @(foreach ($list in $lists) {
        if ($list.EntitiesCount -gt 0) {
            try { Get-InsiderRiskEntityList -Identity $list.Name -ErrorAction Stop }
            catch { Write-Warning "Entries for list '$($list.Name)' could not be retrieved: $($_.Exception.Message)"; $list }
        }
        else { $list }
    })
Save-Json $lists 'IrmEntityLists.raw.json'

try {
    Write-Host 'Retrieving sensitive information types...'
    Save-Json (Get-DlpSensitiveInformationType | Select-Object @{ n = 'Id'; e = { [string]$_.Id } }, Name, Publisher) 'SensitiveInfoTypes.json'
}
catch { Write-Warning "Could not retrieve sensitive information types, so SIT exclusions will show as GUIDs: $($_.Exception.Message)" }

try {
    Write-Host 'Retrieving sensitivity labels...'
    $labels = Get-Label | Select-Object @{ n = 'Id'; e = { [string]$_.ImmutableId } }, @{ n = 'Guid'; e = { [string]$_.Guid } },
        Name, DisplayName, @{ n = 'ParentId'; e = { [string]$_.ParentId } }
    Save-Json $labels 'Labels.json'
}
catch { Write-Warning "Could not retrieve sensitivity labels, so internal label names will be shown: $($_.Exception.Message)" }

try {
    Write-Host 'Retrieving IRM role groups...'
    $roleGroups = foreach ($g in @(Get-RoleGroup | Where-Object Name -like 'InsiderRisk*')) {
        # Identity/DN can match multiple objects in some tenants; the object GUID is unique
        $members = @()
        try {
            $members = @(Get-RoleGroupMember -Identity ([string]$g.ExchangeObjectId) -ResultSize Unlimited -ErrorAction Stop | ForEach-Object {
                    [pscustomobject]@{ Name = [string]$_.DisplayName; Address = [string]$_.PrimarySmtpAddress; Type = [string]$_.RecipientType }
                })
        }
        catch { Write-Warning "Could not retrieve members of role group $($g.Name): $($_.Exception.Message)" }
        [pscustomobject]@{ Name = [string]$g.Name; DisplayName = [string]$g.DisplayName; Description = [string]$g.Description; Members = $members }
    }
    Save-Json $roleGroups 'IrmRoleGroups.json'
}
catch { Write-Warning "Could not retrieve IRM role groups: $($_.Exception.Message)" }

$policyCount = @($policies | Where-Object InsiderRiskScenario -ne 'TenantSetting').Count
[pscustomobject]@{
    ClientName    = $ClientName
    ExportedBy    = $(if ($UserPrincipalName) { $UserPrincipalName } else { (Get-ConnectionInformation | Select-Object -First 1).UserPrincipalName })
    ExportedOnUTC = (Get-Date).ToUniversalTime().ToString('o')
    PolicyCount   = $policyCount
    ListCount     = $lists.Count
} | ConvertTo-Json | Set-Content (Join-Path $OutputFolder 'ExportInfo.json') -Encoding UTF8

Write-Host "Exported $policyCount policies, global settings and $($lists.Count) entity lists to $OutputFolder"

if (-not $SkipWorkbook) {
    & (Join-Path $PSScriptRoot 'Build-IrmWorkbook.ps1') -ExportFolder $OutputFolder -ClientName $ClientName
}
$OutputFolder
