#Requires -Modules ExchangeOnlineManagement
<#
.SYNOPSIS
    Exports Microsoft Purview DLP policies and rules from a tenant and builds a single filterable Excel workbook
    (via Build-DlpWorkbook.ps1) as an offline configuration snapshot. Outputs the export folder path.
.EXAMPLE
    .\Export-DlpConfiguration.ps1 -UserPrincipalName admin@contoso.onmicrosoft.com -ClientName Contoso -OutputRoot C:\Dev\Contoso
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
$OutputFolder = Join-Path $OutputRoot "DLP-Export-$safeClient-$(Get-Date -Format 'yyyyMMdd-HHmm')"

function Save-Json($Data, [string]$Name) {
    ConvertTo-Json -InputObject @($Data) -Depth 20 | Set-Content (Join-Path $OutputFolder $Name) -Encoding UTF8
}

if (-not $SkipConnect) { Connect-IPPSSession -UserPrincipalName $UserPrincipalName }
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null

Write-Host 'Retrieving DLP policies, rules and sensitive information types...'
$policies = @(Get-DlpCompliancePolicy)
$rules = @(Get-DlpComplianceRule)
Save-Json $policies 'DlpPolicies.raw.json'
Save-Json $rules 'DlpRules.raw.json'
Save-Json (Get-DlpSensitiveInformationType | Select-Object @{ n = 'Id'; e = { [string]$_.Id } }, Name, Publisher) 'SensitiveInfoTypes.json'

try {
    Write-Host 'Retrieving sensitivity labels...'
    $labels = Get-Label | Select-Object @{ n = 'Id'; e = { [string]$_.ImmutableId } }, @{ n = 'Guid'; e = { [string]$_.Guid } },
        Name, DisplayName, @{ n = 'ParentId'; e = { [string]$_.ParentId } }, Priority
    Save-Json $labels 'Labels.json'
}
catch {
    Write-Warning "Could not retrieve sensitivity labels, so internal label names will be shown: $($_.Exception.Message)"
}

try {
    Write-Host 'Retrieving Endpoint DLP / policy configuration...'
    Save-Json (Get-PolicyConfig) 'PolicyConfig.raw.json'
}
catch {
    Write-Warning "Could not retrieve policy configuration, so device groups will show as GUIDs: $($_.Exception.Message)"
}

[pscustomobject]@{
    ClientName    = $ClientName
    ExportedBy    = $(if ($UserPrincipalName) { $UserPrincipalName } else { (Get-ConnectionInformation | Select-Object -First 1).UserPrincipalName })
    ExportedOnUTC = (Get-Date).ToUniversalTime().ToString('o')
    PolicyCount   = $policies.Count
    RuleCount     = $rules.Count
} | ConvertTo-Json | Set-Content (Join-Path $OutputFolder 'ExportInfo.json') -Encoding UTF8

Write-Host "Exported $($policies.Count) policies and $($rules.Count) rules to $OutputFolder"

if (-not $SkipWorkbook) {
    & (Join-Path $PSScriptRoot 'Build-DlpWorkbook.ps1') -ExportFolder $OutputFolder -ClientName $ClientName
}
$OutputFolder
