---
name: purview-irm-export
description: 'Export Microsoft Purview Insider Risk Management (IRM) policies, global settings, settings lists / detection groups and IRM role groups from a client tenant via Security & Compliance PowerShell and build a client-ready, filterable Excel workbook (policies, indicators and thresholds, indicator matrix, triggers, global settings, global exclusions, detection group entries, priority user groups). Use for IRM as-built documentation, "export insider risk policies", "document IRM settings", "IRM configuration snapshot", offline record of IRM config, or rebuilding the IRM workbook from an existing IRM-Export folder (IrmPolicies.raw.json, IrmEntityLists.raw.json).'
argument-hint: 'Client name and admin UPN for the tenant, or path to an existing IRM-Export folder'
---

# Purview Insider Risk Management Export

Produces an offline snapshot of every IRM policy plus the tenant-wide IRM settings: raw JSON (full fidelity) plus one filterable `<Client>-IRM-Configuration-<yyyyMMdd>.xlsx` for client documentation.

## When to Use

- Documenting a client's IRM policies and global settings as-built, or before/after a change.
- Re-generating the workbook from a previous export folder (no tenant access needed).

## Prerequisites

| Item | Check | Fix |
|---|---|---|
| `ExchangeOnlineManagement` module | `Get-Module -ListAvailable ExchangeOnlineManagement` | `Install-Module ExchangeOnlineManagement -Scope CurrentUser` (ask first) |
| Excel engine (either) | `Get-Module -ListAvailable ImportExcel`, or `py -3 -c "import openpyxl"` | `Install-Module ImportExcel -Scope CurrentUser` or `py -3 -m pip install --user openpyxl` (ask first) |
| IRM role | Account can run `Get-InsiderRiskPolicy` | Insider Risk Management or Insider Risk Management Admins role group. Compliance Administrator alone may return nothing. `Get-Label` needs a label reader role; without it labels show internal names. |

## Procedure

### 1. Choose the path

- **Existing export folder** (contains `IrmPolicies.raw.json`) → step 3.
- **New export** → ask for the **client name** and the **admin UPN**. Output goes to the current workspace root unless the user says otherwise.

### 2. Export from the tenant and build

Run synchronously in the terminal:

```powershell
& "<skill>/scripts/Export-IrmConfiguration.ps1" -UserPrincipalName admin@client.onmicrosoft.com -ClientName Contoso -OutputRoot "<folder>"
```

- Sign-in is interactive (browser / MFA). Tell the user to complete it in the pop-up; never ask for, type or handle credentials.
- Read-only: only `Get-*` cmdlets are used.
- The Security & Compliance session lives in the terminal where `Connect-IPPSSession` ran. If the user is already connected in their own terminal, give them the command with `-SkipConnect` to run there rather than running it in a new agent terminal (which will not be connected).
- Output folder: `IRM-Export-<Client>-<yyyyMMdd-HHmm>/` containing `IrmPolicies.raw.json`, `IrmEntityLists.raw.json`, `SensitiveInfoTypes.json`, `Labels.json`, `IrmRoleGroups.json`, `ExportInfo.json` and the workbook. The script outputs the folder path as its last line.
- `-SkipWorkbook` exports only.

Script: [Export-IrmConfiguration.ps1](./scripts/Export-IrmConfiguration.ps1)

### 3. Build (or rebuild) the workbook offline

```powershell
& "<skill>/scripts/Build-IrmWorkbook.ps1" -ExportFolder "<export folder>"   # optional: -ClientName, -OutputPath, -Engine ImportExcel|Python
```

Script: [Build-IrmWorkbook.ps1](./scripts/Build-IrmWorkbook.ps1)

### 4. Validate before handing over

- Cover policy count matches `ExportInfo.json` `PolicyCount` (the `TenantSetting` object is excluded).
- No warning about lists reporting entries that were not exported. If there is one, the list entries were not re-fetched by `-Identity` — re-run the export.
- Warnings about missing `SensitiveInfoTypes.json` / `Labels.json` mean SIT exclusions show as GUIDs and labels show internal names — re-run the export.
- Inspect a few rows with openpyxl (`py -3`) rather than asking the user to open Excel: Policies (scope, triggers, health), Indicator Matrix (Global Setting column, scored flags), Global Settings, List Entries.
- Spot-check one policy's indicators and thresholds against the Purview portal.

### 5. Deliver

Report the workbook path and caveats (see Gotchas). Remind the user the export reveals the client's insider risk detection logic, priority user groups and exclusions, and must be handled as confidential.

## Workbook Layout

| Sheet | Content |
|---|---|
| Cover | Client, tenant, snapshot time, exported by, policy counts, global indicators on, list counts, priority user groups, sheet guide, notes |
| Policies | One row per policy: template, status, priority, users in scope (incl. priority user groups / adaptive scopes), exclusions, priority content (labels, SITs, classifiers, sites, Teams, file extensions), score-only-priority-content, triggers and triggering events, indicator count, timeframes, policy health, description |
| Policy Indicators | One row per policy x selected indicator: category, type (indicator / booster / extended), used in risk scoring, threshold mode, Low/Medium/High thresholds per variant (non-priority, priority, unallowed domains) |
| Indicator Matrix | Indicators on globally or in any policy; Global Setting (On/Off/n/a) and one column per policy (`On` / `On (scored)`) |
| Policy Triggers | Triggering events selected per policy, custom thresholds, cumulative window; DLP / Communication Compliance policy triggers |
| Global Settings | `TenantSettings` blob flattened: privacy, timeframes, intelligent detections, analytics, data sharing, export alerts, admin notifications, inline alert customisation, Teams, Adaptive Protection; Internal Key column shows the source path |
| Settings Lists | All built-in `Irm*` lists and custom detection groups with entry counts (empty lists are shown so "not configured" is visible) |
| List Entries | Every entry: domains (with subdomain flag), SITs (resolved name, built-in/custom), file paths, keywords, sites, global exclusion groups (resolved to group name) |
| Priority User Groups | Priority user group names referenced by policies, and which policies use them |
| Role Groups | Only when `IrmRoleGroups.json` exists: IRM role groups and members |

## Data Model (learned from live tenant, Contoso 2026-10)

- `Get-InsiderRiskPolicy` returns policies **and** one `InsiderRiskScenario = TenantSetting` object (`IRM_Tenant_Setting_<guid>`). Global settings are a JSON string in its `TenantSettings` property with sections `Indicators`, `ExtensibleIndicators` (Name/Enabled), `TimeSpan`, `IntelligentDetections`, `FeatureSettings`, `NotificationPreferences`, `InterpretedSettings`, `DynamicRiskPreventionSettings` (Adaptive Protection, null when off), `AgentAdaptiveProtectionSettings`.
- Policy `Indicators` / `ExtensibleIndicators` / `TriggerInsightGroups` / `SensitivityLabels` / `DlpSensitiveTypes` / `MLClassifierTypes` / `CustomTags` are arrays of JSON strings. `Enabled` = indicator selected; `UseInScoring` = used in risk scoring; sub-`Indicators[].Thresholds` = `{LowRisk, MediumRisk, HighRisk}`.
- **Priority user groups** appear only as names in policy `CustomTags`; members are not exposed. Policies scoped to a priority user group have an empty `ExchangeLocation`.
- `Get-InsiderRiskEntityList` **requires `-Type` or `-Identity`**. Listing by `-Type` returns `EntitiesCount` but an empty `Entities`; the export re-fetches lists with entries by `-Identity <Name>`. Entries are JSON strings: `{Dmn,isMLSubDmn}`, `{Guid}` (SIT/classifier), `{GroupId}` (global exclusion → detection group Name), `{FlPthRgx}`, `{Ext}`, `{Name}` (keyword), `{Url,Name,Guid}` (site).
- Custom detection groups have base64-like `Name` (e.g. `Z8IF4/...==`) and the friendly `DisplayName`.

## Gotchas

- **IRM cmdlets are undocumented** by Microsoft (no Learn pages); output may change. Labels for settings are partly derived from internal names — the Internal Name / Internal Key columns keep the source value. Uncertain mappings: `OptInIRMDataExport`, `OLThresHolds`, `IrmEnterpriseDomains` / `IrmPublicDomains`, threshold mode `IPTRecommended` (shown as "IPT Recommended").
- **Raw JSON has case-duplicate enum keys** (`{"value":1,"Value":"Tenant"}`) which break `ConvertFrom-Json`; the builder strips the numeric `value` first.
- Sequence indicators and some AI "priority" indicators have no tenant-wide toggle, so the matrix shows `n/a` for Global Setting.
- PowerShell unrolls single-item arrays returned from functions — wrap with `@()` before using `+` (caused `op_Addition` errors during development).
- Hashtable literal keys are case-insensitive (`Spo` and `SPO` clash).
- `Get-RoleGroupMember -Identity <role group Identity/DN>` can fail with "matches multiple entries" (seen at Contoso); the export uses `ExchangeObjectId` instead.
- **Save fails** if the xlsx is open in Excel — close it and re-run.

## Limits

- Priority user group membership, alerts, cases, analytics results and Power Automate flows are not exported.
- Excel cells cap at 32,767 characters; longer values are truncated in the workbook only (raw JSON is complete).
