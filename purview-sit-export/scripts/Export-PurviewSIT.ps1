<#
.SYNOPSIS
    Exports custom Microsoft Purview sensitive information types (SITs) for offline documentation.
.DESCRIPTION
    Writes SIT-Summary.csv, RulePack-<id>.xml (raw, re-importable definitions), KeywordDictionaries.csv and
    ExportInfo.json (tenant + export metadata) to a timestamped folder. Feed that folder to Convert-SITExport.py.
.EXAMPLE
    .\Export-PurviewSIT.ps1 -UserPrincipalName admin@contoso.onmicrosoft.com -OutputRoot C:\Dev\Contoso
#>
[CmdletBinding()]
param(
    [string]$UserPrincipalName,
    [string]$OutputRoot = (Get-Location).Path,
    [string]$TenantName,
    [switch]$SkipConnect
)

$ErrorActionPreference = 'Stop'

if (-not $SkipConnect) {
    if (-not (Get-Module -ListAvailable ExchangeOnlineManagement)) {
        throw 'ExchangeOnlineManagement module not found. Install with: Install-Module ExchangeOnlineManagement -Scope CurrentUser'
    }
    Import-Module ExchangeOnlineManagement
    if ($UserPrincipalName) { Connect-IPPSSession -UserPrincipalName $UserPrincipalName } else { Connect-IPPSSession }
}

$connection = Get-ConnectionInformation -ErrorAction SilentlyContinue |
    Where-Object { $_.State -eq 'Connected' -and $_.ConnectionUri -match 'compliance' } |
    Select-Object -First 1
if (-not $connection) { throw 'No Security & Compliance PowerShell session found. Run without -SkipConnect.' }

$customSits = @(Get-DlpSensitiveInformationType | Where-Object { $_.Publisher -ne 'Microsoft Corporation' })
if ($customSits.Count -eq 0) { Write-Warning 'No custom SITs found in this tenant.' }

# Prefer the tenant's own domain; the signed-in UPN may belong to a partner/consultant tenant
if (-not $TenantName) {
    $TenantName = @(
        $connection.DelegatedOrganization
        $connection.Organization
        $customSits.Publisher | Where-Object { $_ -match '\.onmicrosoft\.com$' } | Select-Object -First 1
        ($connection.UserPrincipalName -split '@')[-1]
    ) | Where-Object { $_ } | Select-Object -First 1
    if (-not $TenantName) { $TenantName = 'Tenant' }
}
$folderLabel = ($TenantName -replace '\.onmicrosoft\.com$', '').Split('.')[0] -replace '[^\w-]', ''
$exportedAt = (Get-Date).ToUniversalTime()

$out = Join-Path $OutputRoot ('SIT-Export-{0}-{1:yyyyMMdd-HHmm}' -f $folderLabel, $exportedAt.ToLocalTime())
New-Item $out -ItemType Directory -Force | Out-Null
# .NET file APIs resolve relative paths against the process cwd, not the PowerShell location
$out = (Resolve-Path $out).ProviderPath

[ordered]@{
    Tenant        = $TenantName
    TenantId      = "$($connection.TenantID)"
    ExportedBy    = $connection.UserPrincipalName
    ExportedAtUtc = $exportedAt.ToString('yyyy-MM-ddTHH:mm:ssZ')
} | ConvertTo-Json | Set-Content (Join-Path $out 'ExportInfo.json') -Encoding utf8

$customSits | Select-Object Name, Id, Publisher, RulePackId, Description |
    Export-Csv (Join-Path $out 'SIT-Summary.csv') -NoTypeInformation -Encoding utf8

foreach ($rpId in ($customSits.RulePackId | Sort-Object -Unique)) {
    $rp = Get-DlpSensitiveInformationTypeRulePackage -Identity $rpId
    [System.IO.File]::WriteAllBytes((Join-Path $out "RulePack-$rpId.xml"), $rp.SerializedClassificationRuleCollection)
}

Get-DlpKeywordDictionary | Select-Object Identity, Name, Description, MatchStyle,
    @{ n = 'Terms'; e = { $_.KeywordDictionary } } |
    Export-Csv (Join-Path $out 'KeywordDictionaries.csv') -NoTypeInformation -Encoding utf8

Write-Host "Exported $($customSits.Count) custom SITs from $TenantName to $out"
$out
