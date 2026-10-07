---
name: purview-retention-export
description: 'Export Microsoft Purview Data Lifecycle Management configuration from a client tenant via PowerShell - retention policies (incl. Teams, app retention), retention labels (incl. records, file plan descriptors, disposition), retention label policies (publish and auto-apply), adaptive scopes, event types - plus Exchange MRM retention policies and retention tags, and build a client-ready filterable Excel workbook. Use for DLM / records management as-built documentation, "export retention policies", "document retention labels", "MRM retention tags", "retention configuration snapshot", offline record of retention config, or rebuilding the workbook from an existing Retention-Export folder (RetentionPolicies.raw.json, RetentionLabels.raw.json, MrmRetentionTags.raw.json).'
argument-hint: 'Client name and admin UPN for the tenant, or path to an existing Retention-Export folder'
---

# Purview Retention & MRM Export

Produces an offline snapshot of Purview retention and Exchange MRM configuration: raw JSON (full fidelity) plus one filterable `<Client>-Retention-Configuration-<yyyyMMdd>.xlsx` for client documentation.

## Prerequisites

| Item | Check | Fix |
|---|---|---|
| `ExchangeOnlineManagement` module | `Get-Module -ListAvailable ExchangeOnlineManagement` | `Install-Module ExchangeOnlineManagement -Scope CurrentUser` (ask first) |
| Excel engine (either) | `Get-Module -ListAvailable ImportExcel`, or `py -3 -c "import openpyxl"` | `Install-Module ImportExcel -Scope CurrentUser` or `py -3 -m pip install --user openpyxl` (ask first) |
| Purview role | Can run `Get-RetentionCompliancePolicy`, `Get-ComplianceTag` | View-Only Retention Management / View-Only Record Management, or Compliance Administrator |
| Exchange role (MRM) | Can run `Get-RetentionPolicyTag` | View-Only Organization Management or Global Reader (skip MRM with `-SkipMrm`) |

## Procedure

### 1. Choose the path

- **Existing export folder** (contains `RetentionPolicies.raw.json`) → step 3.
- **New export** → ask for the **client name** and **admin UPN**, and whether to include **mailbox counts per MRM policy** (`-IncludeMailboxCounts`; enumerates every mailbox, slow in large tenants, saves counts only). Output goes to the workspace root unless told otherwise.

### 2. Export from the tenant and build

```powershell
& "<skill>/scripts/Export-RetentionConfiguration.ps1" -UserPrincipalName admin@client.onmicrosoft.com -ClientName Contoso -OutputRoot "<folder>" [-IncludeMailboxCounts]
```

- Two interactive sign-ins: Exchange Online (MRM) then Security & Compliance. Tell the user to complete them in the pop-up; never ask for, type or handle credentials.
- Read-only: only `Get-*` cmdlets are used.
- Output folder `Retention-Export-<Client>-<yyyyMMdd-HHmm>/`: `RetentionPolicies.raw.json`, `RetentionRules.raw.json`, `RetentionLabels.raw.json`, `AppRetentionPolicies.raw.json`, `AppRetentionRules.raw.json`, `AdaptiveScopes.raw.json`, `RetentionEventTypes.raw.json`, `FilePlanProperties.raw.json`, `MrmRetentionPolicies.raw.json`, `MrmRetentionTags.raw.json`, optional `MrmMailboxCounts.json`, `ExportInfo.json` and the workbook. Last output line is the folder path.
- `-SkipMrm` skips Exchange Online; `-SkipConnect` reuses open sessions; `-SkipWorkbook` exports only.

Script: [Export-RetentionConfiguration.ps1](./scripts/Export-RetentionConfiguration.ps1)

### 3. Build (or rebuild) the workbook offline

```powershell
& "<skill>/scripts/Build-RetentionWorkbook.ps1" -ExportFolder "<export folder>"   # optional: -ClientName, -OutputPath, -Engine ImportExcel|Python
```

Script: [Build-RetentionWorkbook.ps1](./scripts/Build-RetentionWorkbook.ps1)

### 4. Validate before handing over

- Cover counts vs `ExportInfo.json` (`PolicyCount` = retention + label policies combined; `LabelCount`; `MrmPolicyCount`, `MrmTagCount`).
- Warning "rule(s) could not be linked" means rule→policy linking failed — inspect `Policy` / `ParentPolicyName` in `RetentionRules.raw.json`.
- Inspect rows with openpyxl (`py -3`): one static and one adaptive retention policy, one Teams policy, one record label, one auto-apply policy, the Default MRM Policy.
- Review **Additional Settings** columns for noise properties and add them to the noise lists in the builder.
- Spot-check one policy and one label against the Purview portal.

### 5. Deliver

Report the workbook path and caveats (see Limits). Remind the user the export should be handled as confidential.

## Workbook Layout

| Sheet | Content |
|---|---|
| Cover | Client, tenant, snapshot time, exported by, counts, sheet guide, notes (preservation lock, principles of retention, exclusions) |
| Retention Policies | One row per retention / app retention policy: status, static/adaptive, locations, exclusions, adaptive scopes, readable retention settings, period, basis, end action, preservation lock |
| Retention Labels | One row per label: type (retention / record / regulatory record), settings, event type, disposition review, record unlocked, published in / auto-applied by, file plan descriptors |
| Label Policies | Publish and auto-apply label policies: labels, locations, scopes, auto-apply conditions (SITs, KQL, trainable classifier IDs) |
| Adaptive Scopes | Only when present. Scope type, readable query, used by policies |
| MRM Retention Policies | Only when exported. Default flag, linked tags with summary, mailbox counts if exported |
| MRM Retention Tags | Only when exported. DPT / folder RPT / personal tag, applies to, enabled, age limit, action, linked policies (`(not linked)` = orphan) |
| Reference Data | Only when present. Event types and file plan descriptor values |

## Classification rules

- A policy whose rule has `PublishComplianceTag` → **Publish** label policy; `ApplyComplianceTag` → **Auto-apply** label policy; otherwise **Retention policy**.
- Status: `Enabled=false` → Off; Mode `Enforce`/`Enable` → On; `Simulate*`/`Test*` → Simulation; `PendingDeletion` → Pending deletion.
- Durations in days are shown as years/months when they divide evenly (365 / 30) or match `RetentionDurationDisplayHint`; `Unlimited` → Forever.
- Labels with `ReviewerEmail` or `MultiStageReviewProperty` end with "Start a disposition review"; `ComplianceTagForNextStage` → "Relabel to X".
- Distribution status is hidden (same convention as the DLP skill).

## Gotchas

- **First live run 2026-10-07 (Contoso, small tenant)** confirmed:
  - `Get-RetentionCompliancePolicy` returns empty `*Location` properties unless `-DistributionDetail` is used (now in the export).
  - Rule `PublishComplianceTag` is `"<guid>,<label name>"`; `ComplianceTagProperty` duplicates it and is treated as noise.
  - Multi-stage disposition reviewers are in label `MultiStageReviewerMetadata` (a JSON string, `{"MultiStageReviewSettings":[{StageId,StageName,Reviewers[]}]}`), not `MultiStageReviewProperty`. `IsReviewTag=true` marks review labels.
  - MRM policy → tag linking works as built.
  - Still unverified (not present in that tenant): adaptive scope `FilterConditions` shape, `FilePlanMetadata` keys, auto-apply rules, record labels, app retention policies.
- Raw JSON has case-duplicate enum keys (`{"value":1,"Value":"Keep"}`); the builder strips the numeric `value` before parsing.
- `AgeLimitForRetention` (TimeSpan) is converted to `AgeLimitDays` at export time.
- Save fails if the xlsx is open in Excel — close it and re-run.
- In PowerShell, parenthesise Excel cell indices: `$ws.Cells[($r + 2), ($c + 1)]`.

## Limits

- Not included: default labels on SharePoint libraries/folders/Outlook folders, item-level labels, retention events, disposition history, priority cleanup policies, eDiscovery / litigation holds.
- Trainable classifiers show as IDs (no PowerShell lookup).
- Excel cells cap at 32,767 characters; raw JSON is complete.
