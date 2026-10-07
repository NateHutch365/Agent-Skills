#Requires -Modules ExchangeOnlineManagement
<#
.SYNOPSIS
    Exports Microsoft Purview retention policies, retention labels, label policies, adaptive scopes and file plan data
    (Security & Compliance PowerShell) plus Exchange MRM retention policies and tags (Exchange Online PowerShell), then
    builds a filterable Excel workbook (via Build-RetentionWorkbook.ps1). Outputs the export folder path.
.DESCRIPTION
    Read-only: only Get-* cmdlets are used.
.EXAMPLE
    .\Export-RetentionConfiguration.ps1 -UserPrincipalName admin@contoso.onmicrosoft.com -ClientName Contoso -OutputRoot C:\Dev\Contoso
#>
[CmdletBinding()]
param(
    [string]$UserPrincipalName,
    [Parameter(Mandatory)][string]$ClientName,
    [string]$OutputRoot = $PWD,
    [switch]$IncludeMailboxCounts,
    [switch]$SkipMrm,
    [switch]$SkipConnect,
    [switch]$SkipWorkbook
)

$ErrorActionPreference = 'Stop'
if (-not $SkipConnect -and -not $UserPrincipalName) { throw 'Provide -UserPrincipalName, or -SkipConnect if Exchange Online and Security & Compliance sessions are already open.' }

$safeClient = $ClientName -replace '[^\w-]', ''
$OutputFolder = Join-Path $OutputRoot "Retention-Export-$safeClient-$(Get-Date -Format 'yyyyMMdd-HHmm')"

function Save-Json($Data, [string]$Name) {
    ConvertTo-Json -InputObject @($Data) -Depth 20 | Set-Content (Join-Path $OutputFolder $Name) -Encoding UTF8
}

function Invoke-Optional([string]$What, [scriptblock]$Script) {
    try { & $Script }
    catch { Write-Warning "Could not retrieve ${What}: $($_.Exception.Message)" }
}

New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
$info = [ordered]@{ ClientName = $ClientName }

#region Exchange Online - MRM
if (-not $SkipMrm) {
    if (-not $SkipConnect) { Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$false }
    Write-Host 'Retrieving MRM retention policies and retention tags...'
    $mrmPolicies = @(Get-RetentionPolicy)
    $mrmTags = @(Get-RetentionPolicyTag | Select-Object *, @{ n = 'AgeLimitDays'; e = {
                $a = $_.AgeLimitForRetention
                if ($null -ne $a -and "$a") { try { [int][timespan]::Parse("$a").TotalDays } catch { $null } }
            }
        })
    Save-Json $mrmPolicies 'MrmRetentionPolicies.raw.json'
    Save-Json $mrmTags 'MrmRetentionTags.raw.json'
    $info.MrmPolicyCount = $mrmPolicies.Count
    $info.MrmTagCount = $mrmTags.Count

    if ($IncludeMailboxCounts) {
        Invoke-Optional 'mailbox counts per MRM policy' {
            Write-Host 'Counting mailboxes per MRM retention policy (can take a while in large tenants)...'
            # Only counts are saved; no per-mailbox data is written to disk.
            $counts = Get-EXOMailbox -ResultSize Unlimited -Properties RetentionPolicy, RetentionHoldEnabled |
                Group-Object { if ($_.RetentionPolicy) { [string]$_.RetentionPolicy } else { '(none)' } } |
                ForEach-Object {
                    [pscustomobject]@{
                        RetentionPolicy    = $_.Name
                        MailboxCount       = $_.Count
                        RetentionHoldCount = @($_.Group | Where-Object RetentionHoldEnabled).Count
                    }
                }
            Save-Json $counts 'MrmMailboxCounts.json'
        }
    }
}
#endregion

#region Security & Compliance - Purview retention
if (-not $SkipConnect) { Connect-IPPSSession -UserPrincipalName $UserPrincipalName }

Write-Host 'Retrieving retention policies, rules and retention labels...'
# Location properties come back empty without -DistributionDetail.
$policies = @(Get-RetentionCompliancePolicy -DistributionDetail)
$rules = @(Get-RetentionComplianceRule)
$labels = @(Get-ComplianceTag)
Save-Json $policies 'RetentionPolicies.raw.json'
Save-Json $rules 'RetentionRules.raw.json'
Save-Json $labels 'RetentionLabels.raw.json'
$info.PolicyCount = $policies.Count
$info.RuleCount = $rules.Count
$info.LabelCount = $labels.Count

Invoke-Optional 'app retention policies (Teams private channels, Viva Engage, Copilot)' {
    Write-Host 'Retrieving app retention policies...'
    $appPolicies = @(Get-AppRetentionCompliancePolicy -DistributionDetail)
    Save-Json $appPolicies 'AppRetentionPolicies.raw.json'
    Save-Json @(Get-AppRetentionComplianceRule) 'AppRetentionRules.raw.json'
    $info.AppPolicyCount = $appPolicies.Count
}

Invoke-Optional 'adaptive scopes' {
    Write-Host 'Retrieving adaptive scopes...'
    Save-Json @(Get-AdaptiveScope) 'AdaptiveScopes.raw.json'
}

Invoke-Optional 'retention event types' {
    Write-Host 'Retrieving retention event types...'
    Save-Json @(Get-ComplianceRetentionEventType) 'RetentionEventTypes.raw.json'
}

Write-Host 'Retrieving file plan descriptors...'
$filePlan = foreach ($t in 'ReferenceId', 'Department', 'Category', 'SubCategory', 'Authority', 'Citation') {
    Invoke-Optional "file plan $t values" {
        & "Get-FilePlanProperty$t" | Select-Object @{ n = 'FilePlanType'; e = { $t } }, *
    }
}
Save-Json $filePlan 'FilePlanProperties.raw.json'
#endregion

$info.ExportedBy = $(if ($UserPrincipalName) { $UserPrincipalName } else { (Get-ConnectionInformation | Select-Object -First 1).UserPrincipalName })
$info.ExportedOnUTC = (Get-Date).ToUniversalTime().ToString('o')
[pscustomobject]$info | ConvertTo-Json | Set-Content (Join-Path $OutputFolder 'ExportInfo.json') -Encoding UTF8

Write-Host "Exported $($policies.Count) retention/label policies, $($labels.Count) retention labels$(if (-not $SkipMrm) { ", $($info.MrmPolicyCount) MRM policies and $($info.MrmTagCount) MRM tags" }) to $OutputFolder"

if (-not $SkipWorkbook) {
    & (Join-Path $PSScriptRoot 'Build-RetentionWorkbook.ps1') -ExportFolder $OutputFolder -ClientName $ClientName
}
$OutputFolder
