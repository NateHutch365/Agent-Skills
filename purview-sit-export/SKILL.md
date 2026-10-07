---
name: purview-sit-export
description: 'Export custom Microsoft Purview sensitive information types (SITs) from a client tenant via Security & Compliance PowerShell and build a client-ready Excel workbook of their patterns, confidence levels (High/Medium/Low), regexes, keyword lists, keyword dictionaries and document fingerprints. Use for Purview / DLP as-built documentation, "export custom SITs", "document sensitive info types", "SIT configuration snapshot", "extract SIT patterns or regex", offline record of SIT config, or converting an existing SIT export folder (RulePack-*.xml, SIT-Summary.csv) into xlsx.'
argument-hint: 'Admin UPN for the tenant, or path to an existing SIT export folder'
---

# Purview Custom SIT Export

Produces an offline snapshot of every custom sensitive information type in a tenant: the raw rule package XML (re-importable) plus a readable `SIT-Configuration.xlsx` for client documentation.

## When to Use

- Documenting a client's custom SITs as-built, or before/after a change.
- Extracting SIT regexes, keyword lists and confidence logic for review.
- Re-generating the workbook from a previously exported folder (no tenant access needed).

## Prerequisites

| Item | Check | Fix |
|---|---|---|
| `ExchangeOnlineManagement` module | `Get-Module -ListAvailable ExchangeOnlineManagement` | `Install-Module ExchangeOnlineManagement -Scope CurrentUser` |
| Python 3.9+ with `openpyxl` | `python -c "import openpyxl"` | `python -m pip install --user openpyxl` (ask the user first) |
| Purview role | Account can run `Get-DlpSensitiveInformationType` | e.g. Compliance Administrator, or a role group with SIT/DLP read rights |

## Procedure

### 1. Choose the path

- **Existing export folder** (contains `SIT-Summary.csv` and `RulePack-*.xml`) → skip to step 3.
- **New export** → step 2. Ask for the admin UPN and where to write output (default: workspace root).

### 2. Export from the tenant

```powershell
& "<skill>/scripts/Export-PurviewSIT.ps1" -UserPrincipalName admin@client.onmicrosoft.com -OutputRoot "<folder>"
```

- Sign-in is interactive (browser / MFA). Tell the user to complete it; never ask for or handle credentials.
- Output: `SIT-Export-<tenant>-<yyyyMMdd-HHmm>/` containing `SIT-Summary.csv`, `RulePack-<id>.xml`, `KeywordDictionaries.csv`, `ExportInfo.json`. The script returns the folder path.
- Tenant name is detected from the connection (`Get-ConnectionInformation` Organization, then the tenant's `*.onmicrosoft.com` SIT publisher, then the UPN domain). `Organization` is usually blank for interactive sign-in, so the publisher or UPN normally supplies it. `ExportInfo.json` records tenant, tenant ID, exporting account and UTC time; the workbook title uses it.
- Use `-SkipConnect` if a session is already open; `-TenantName` only to override detection.

### 3. Build the workbook

```powershell
python "<skill>/scripts/Convert-SITExport.py" "<export folder>"   # optional: -o <file.xlsx>
```

Writes `<export folder>/SIT-Configuration.xlsx` and prints `N SITs, M patterns`.

### 4. Validate before handing over

- `N` must equal the row count of `SIT-Summary.csv`. A shortfall means a SIT type the converter does not parse (see Limits).
- Open the **Patterns** sheet and spot-check one multi-pattern SIT against the Purview portal: pattern numbers, confidence and primary element should match.
- Check for **Keyword dictionary** rows showing *Referenced by ID only* — see Gotchas.
- Check the Overview title shows the correct tenant. If detection fell back to the UPN domain (consultant account), re-run the export with `-TenantName`.

### 5. Deliver

Report the output paths and any caveats (unresolved dictionaries, fingerprint limits). Remind the user the files reveal the client's detection logic and should be handled as confidential.

## Workbook Layout

| Sheet | Content |
|---|---|
| Overview | Title with tenant name; tenant ID, export time and exporting account. One row per SIT: type, description, recommended confidence, default proximity, High/Medium/Low pattern counts, last modified. Name links to its patterns. |
| Patterns | One row per element, grouped by SIT then pattern. Columns: Pattern, Confidence, Proximity, Role (Primary/Supporting), Requirement (`Required` / `Group N - at least X of`), Element, Type, Definition, Match style, Case. Excel table with filters. |
| Keyword Dictionaries | Name, match style, term count, terms, dictionary ID. |
| Notes | Legend for confidence mapping, roles, requirements, limits. |

Confidence mapping: `>=85` High, `75-84` Medium, `<75` Low (Purview defaults are 85/75/65).

## Gotchas

- **Never re-sort patterns.** Numbering follows rule package XML order, which matches the portal. Sorting by confidence renumbers them and the client will not recognise "Pattern 1".
- **Rule pack XML is UTF-16LE without a BOM.** Default readers mis-decode it; the converter handles this.
- **Patterns can be nested** inside `<Version minEngineVersion=...>` wrappers; always search descendants.
- **Fingerprint SITs** live in the *Document Fingerprint Rule Package* (`00000000-0000-0000-0001-000000000001`) as `<Affinity>`/`<Evidence>`, not `<Entity>`. Thresholds come from `<FingerprintConfig>` (typically 30/50/80% → Low/Medium/High).
- **GUID element references** not defined in the rule pack are keyword dictionaries. `Get-DlpKeywordDictionary` `Identity` is the GUID the SIT references (confirmed); the converter resolves it to the dictionary name, terms and match style. *Referenced by ID only* means the dictionary CSV lacks `Identity` (older export) or the dictionary was deleted.
- **`PermissionError` on save** = the xlsx is open in Excel. Close it and re-run.
- Built-in Microsoft SITs are excluded (`Publisher -ne 'Microsoft Corporation'`); portal *copies* of built-ins are custom and included.

## Limits

- Exact Data Match (EDM): schema/SIT definition only; the hashed data is not exportable.
- Document fingerprints: thresholds only; the template document is not recoverable.
- Trainable classifiers are not exported.
