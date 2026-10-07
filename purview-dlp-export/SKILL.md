---
name: purview-dlp-export
description: 'Export Microsoft Purview Data Loss Prevention (DLP) policies and rules from a client tenant via Security & Compliance PowerShell and build a single client-ready, filterable Excel workbook (policies, rules, conditions, sensitive info types with confidence and instance counts, labels, exceptions, actions, notifications, overrides, alerts, SIT usage matrix). Use for DLP as-built documentation, "export DLP policies", "document DLP rules", "DLP configuration snapshot", offline record of DLP config, or rebuilding the DLP workbook from an existing DLP-Export folder (DlpPolicies.raw.json, DlpRules.raw.json).'
argument-hint: 'Client name and admin UPN for the tenant, or path to an existing DLP-Export folder'
---

# Purview DLP Policy Export

Produces an offline snapshot of every DLP policy and rule in a tenant: raw JSON (full fidelity) plus one filterable `<Client>-DLP-Configuration-<yyyyMMdd>.xlsx` for client documentation.

## When to Use

- Documenting a client's DLP policies and rules as-built, or before/after a change.
- Re-generating the workbook from a previous export folder (no tenant access needed).

## Prerequisites

| Item | Check | Fix |
|---|---|---|
| `ExchangeOnlineManagement` module | `Get-Module -ListAvailable ExchangeOnlineManagement` | `Install-Module ExchangeOnlineManagement -Scope CurrentUser` (ask first) |
| Excel engine (either) | `Get-Module -ListAvailable ImportExcel`, or `py -3 -c "import openpyxl"` | `Install-Module ImportExcel -Scope CurrentUser` or `py -3 -m pip install --user openpyxl` (ask first) |
| Purview role | Account can run `Get-DlpCompliancePolicy` | View-Only DLP Compliance Management (read-only) or Compliance Administrator. `Get-Label` additionally needs a label reader role; without it labels show internal names. |

ImportExcel is preferred; the builder falls back to Python + openpyxl automatically. Excel itself is not required.

## Procedure

### 1. Choose the path

- **Existing export folder** (contains `DlpPolicies.raw.json`) → step 3.
- **New export** → ask for the **client name** (used on the cover and file name) and the **admin UPN**. Output goes to the current workspace root unless the user says otherwise.

### 2. Export from the tenant and build

Run synchronously in the terminal:

```powershell
& "<skill>/scripts/Export-DlpConfiguration.ps1" -UserPrincipalName admin@client.onmicrosoft.com -ClientName Contoso -OutputRoot "<folder>"
```

- Sign-in is interactive (browser / MFA). Tell the user to complete it in the pop-up; never ask for, type or handle credentials.
- Read-only: only `Get-*` cmdlets are used.
- Output folder: `DLP-Export-<Client>-<yyyyMMdd-HHmm>/` containing `DlpPolicies.raw.json`, `DlpRules.raw.json`, `SensitiveInfoTypes.json`, `Labels.json`, `PolicyConfig.raw.json`, `ExportInfo.json` and the workbook. The script outputs the folder path as its last line.
- `-SkipConnect` reuses an open session; `-SkipWorkbook` exports only.

Script: [Export-DlpConfiguration.ps1](./scripts/Export-DlpConfiguration.ps1)

### 3. Build (or rebuild) the workbook offline

```powershell
& "<skill>/scripts/Build-DlpWorkbook.ps1" -ExportFolder "<export folder>"   # optional: -ClientName, -OutputPath, -Engine ImportExcel|Python
```

Client name defaults to `ExportInfo.json`, then the tenant prefix. Script: [Build-DlpWorkbook.ps1](./scripts/Build-DlpWorkbook.ps1)

### 4. Validate before handing over

- Cover sheet policy and rule counts match `ExportInfo.json` (`PolicyCount`, `RuleCount`).
- Warnings about missing `SensitiveInfoTypes.json` / `Labels.json` mean an older export or missing role — re-run the export, or tell the user labels show internal names.
- Inspect a few rows with openpyxl (`py -3`) rather than asking the user to open Excel: check Conditions, Sensitive Info Types and Actions read sensibly for one Exchange, one Devices and one label-based rule.
- Spot-check one rule against the Purview portal.

### 5. Deliver

Report the workbook path and caveats (see Gotchas). Remind the user the export reveals the client's detection and protection logic and should be handled as confidential.

## Workbook Layout

| Sheet | Content |
|---|---|
| Cover | Client, tenant, snapshot time (UTC), exported by, policy/rule counts, sheet guide, notes |
| DLP Rules | One row per rule with policy context repeated (Policy, Rule, Policy Mode, Rule Status, Severity, Locations, Included/Excluded Scope, priorities). Readable Conditions, SIT count, SITs as `Name (Confidence, instances)`, Sensitivity Labels, Exceptions, Actions, User Notifications, Policy Tip Text, User Overrides, Alert Recipients/Settings, Incident Reports, Additional Settings, description, last modified. Policies with no rules get a `(no rules)` row. |
| Policies | One row per policy: name (display name), mode, priority, locations, scope, rule list, description, additional settings, created/modified |
| SIT Usage | Each SIT/label x each policy with the instance thresholds used; Type = Custom SIT / Built-in SIT / Sensitivity label |
| Endpoint Settings | Only when `PolicyConfig.raw.json` exists. Endpoint DLP global settings (`EndpointDlpGlobalSettings`, JSON values expanded) plus other non-default `Get-PolicyConfig` properties. |

All data sheets are Excel tables (filters, banding), wrapped text, frozen header and key columns. Mode/Status/Severity are colour-coded.

## Display Conventions (agreed with users)

- **Policies and rules show the display name only** (`DisplayName`, falling back to `Name`) on every sheet, matching the Purview portal. Internal names are often stale after renames and are hidden; they are used only to link rules to policies.
- **Distribution status is not shown.**
- Locations are derived from the location properties, not `Workload` (which lists Exchange/SharePoint/OneDrive for every policy).
- Policy modes: `Enable` → On, `TestWithNotifications` → Simulation (policy tips shown), `TestWithoutNotifications` → Simulation, `Disable` → Off.
- Instance counts: `1+` = one or more, `1-9` = range, from `MinCount`/`MaxCount` (`-1` = any).
- Internal IDs, default/False flags and system properties are hidden; anything else non-default lands in **Additional Settings** so nothing configured is lost.

## Gotchas

- **Raw JSON has case-duplicate enum keys** (`{"value":1,"Value":"Tenant"}`) which break `ConvertFrom-Json`; the builder strips the numeric `value` before parsing.
- **Most portal-created rules store logic in `AdvancedRule`** (JSON string with nested `SubConditions`, `Operator` And/Or/Not, `Groups` of `sensitivetypes` and `labels`). Top-level `Not` groups are exceptions.
- **Device action groups** (e.g. `cloudEgressGroup`) are resolved to names by scanning `PolicyConfig.raw.json` for any object carrying both a GUID id and a name. First live run (Contoso, 2026-10): `SiteGroups` was `[]` and the `cloudEgressGroup` GUID appeared nowhere in `Get-PolicyConfig`, so it stays a GUID — tell the user to look up the sensitive service domain group name in the portal (Endpoint DLP settings). `EndpointDlpGlobalSettings` is a JSON string of `Setting`/`Value` pairs (repeated Settings = list values) and `EndpointDlpGlobalSettingsPsws` duplicates it; the builder handles both.
- Trainable classifiers appear in SIT lists without a confidence level.
- **Save fails** if the xlsx is open in Excel — close it and re-run.
- When editing the builder: in PowerShell, `$ws.Cells[$r + 2, $c + 1]` parses wrongly (comma binds tighter than `+`); parenthesise each index.

## Limits

- Adaptive scope definitions and SIT definitions are not included. Use the `purview-sit-export` skill for custom SIT patterns.
- Excel cells cap at 32,767 characters; longer values are truncated in the workbook only (raw JSON is complete).
