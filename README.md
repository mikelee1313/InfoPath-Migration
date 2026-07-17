# InfoPath XML to SharePoint List Migration (PnP App-Only)

Migrate legacy InfoPath XML-based records from a source SharePoint list/library into a modern target SharePoint list using **PnP.PowerShell** and **app-only authentication**.

This solution is designed for scenarios where InfoPath is retired/disabled, and historical XML form data plus embedded attachments must be preserved in a new OOTB list.

## What This Script Does

Script file: `migrate-infopath-xml-to-list-pnp.ps1`

At a high level, the script:

1. Connects to SharePoint Online using app-only auth (certificate).
2. Reads source items from the source list/library.
3. Resolves XML payloads per source item:
   - If source item is an `.xml` file in a library, reads file content directly.
   - If source is a classic list item, checks item attachment collection for XML.
4. Parses InfoPath XML metadata into key/value pairs.
5. Auto-provisions missing target columns (Text columns) for XML keys not present in target.
6. Creates target list items and maps metadata values to target fields.
7. Extracts embedded InfoPath attachments from XML base64 payloads and uploads them to the newly created target item.
8. Outputs summary statistics (processed/resolved/created/uploaded/skipped/failed).

## Key Behaviors

- **Cross-site support**: source and target can be on different site collections.
- **URL-aware list input**: source/target can be provided as list title, list GUID, or list URL.
- **Automatic retries**: transient throttling/server errors (429/502/503/504) use exponential backoff.
- **Non-interactive execution**: avoids interactive prompts in list attachment resolution.
- **Namespace filtering**: metadata mapping only uses InfoPath `my` namespace nodes to avoid junk columns like `div`, `a`, `br`.
- **Value sanitization**: strips invalid control characters and normalizes text-field values to reduce `Invalid text value` errors.

## Prerequisites

## 1) PowerShell Module

Install PnP.PowerShell:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

## 2) Entra ID App Registration

Configure an app registration for app-only access to SharePoint Online.

Minimum expectations:

- SharePoint application permissions appropriate for read/write list + attachments operations.
- Admin consent granted.
- If using certificate auth: certificate is installed in local cert store and private key available.
- If using client secret auth: secret stored securely (environment variable recommended).

## 3) Source Data Shape

The script expects legacy data where InfoPath XML exists either:

- as XML files in a source document library, or
- as XML attachments on classic list items.

## Authentication Options

## Certificate (Required)

Parameters used:

- `-AuthType Certificate`
- `-TenantId`
- `-ClientId`
- `-Thumbprint`
- `-CertStore` (`LocalMachine` or `CurrentUser`)


Parameters used:

- `-AuthType ClientSecret`
- `-TenantId`
- `-ClientId`
  
```

## Parameters

The script currently includes defaults in the param block. Override at runtime as needed.

- `SiteUrl` (string)
  - Fallback site URL used when source/target list values are titles/IDs rather than URLs.
- `SourceList` (string)
  - Source list identity: title, GUID, or URL.
- `TargetList` (string)
  - Target list identity: title, GUID, or URL.
- `TenantId` (string)
- `ClientId` (string)
- `AuthType` (`Certificate` | `ClientSecret`)
- `Thumbprint` (string)
- `CertStore` (`LocalMachine` | `CurrentUser`)
- `ClientSecret` (string)
- `PageSize` (int, default `200`)
- `MaxItems` (int, default `0` meaning all)
- `TempFolder` (string)
- `FallbackAttachmentName` (string)
- `SkipAttachments` (switch)

## Usage Examples

## 1) Run with script defaults

```powershell
. 'C:\Users\michlee\OneDrive - Microsoft\SfMC_Docs\Customers\jefferies\InfoPath_export_to_List\migrate-infopath-xml-to-list-pnp.ps1'
```

## 2) Override key values explicitly

```powershell
.\migrate-infopath-xml-to-list-pnp.ps1 \
  -SiteUrl 'https://contoso.sharepoint.com/sites/Main' \
  -SourceList 'https://contoso.sharepoint.com/sites/Legacy/Shared%20Documents' \
  -TargetList 'https://contoso.sharepoint.com/sites/Modern/Lists/Subscriptions' \
  -TenantId 'tenant-guid' \
  -ClientId 'app-guid' \
  -AuthType Certificate \
  -Thumbprint 'CERT_THUMBPRINT' \
  -CertStore LocalMachine
```

## 3) Client secret mode

```powershell
.\migrate-infopath-xml-to-list-pnp.ps1 \
  -SiteUrl 'https://contoso.sharepoint.com/sites/Main' \
  -SourceList 'LegacyList' \
  -TargetList 'ModernList' \
  -TenantId 'tenant-guid' \
  -ClientId 'app-guid' \
  -AuthType ClientSecret \
  -ClientSecret $env:PNP_CLIENT_SECRET
```

## 4) Metadata-only migration (no attachment upload)

```powershell
.\migrate-infopath-xml-to-list-pnp.ps1 -SkipAttachments
```

## 5) Test subset first

```powershell
.\migrate-infopath-xml-to-list-pnp.ps1 -MaxItems 20
```

## Migration Flow (Detailed)

1. Resolve source and target contexts (`site + list identity`) from title/GUID/URL input.
2. Connect to source site.
3. Read source list items.
4. For each source item:
   - attempt XML resolution,
   - parse XML document,
   - queue successful payloads.
5. Connect to target site.
6. Read writable target fields.
7. Parse all queued XML maps and collect all keys.
8. Auto-create missing target columns as Text under group `InfoPath Migrated Columns`.
9. Create target items with sanitized mapped values.
10. Extract and upload embedded attachments.
11. Print migration summary.

## Important Notes About Field Creation

- Missing XML keys are created as **single-line Text** fields by default.
- Internal names are sanitized (`A-Z`, `a-z`, `0-9`, `_`) and length-limited.
- Display name is the original XML key.

If you need richer field typing (Date/Number/YesNo/Choice), extend field inference logic before creation.

## Troubleshooting

## "Supply values for the following parameters" during run

Cause:

- Usually from a cmdlet call shape that triggers interactive prompting.

Current status:

- Script is hardened to avoid interactive attachment lookup behavior.

## "Invalid text value"

Cause:

- Input text contains unsupported control chars, long multiline payloads, or incompatible formatting for Text fields.

Mitigation in script:

- Control characters removed.
- Text fields flattened to single-line and truncated to 255 chars.

If still occurring:

- Add field-specific conversions (for dates, choice, boolean, etc.).

## Many items show "Skipped: no InfoPath XML found"

Expected when source library contains non-XML artifacts (folders/docs).

- Only items with resolvable XML payload are migrated.

## Wrong source/target list resolved

Verify `SourceList` and `TargetList` values:

- List title: `Subscriptions`
- List URL: `https://tenant.sharepoint.com/sites/Site/Lists/Subscriptions`
- Library URL: `https://tenant.sharepoint.com/sites/Site/Shared%20Documents`

## Throttling/timeouts

- Retry wrapper already handles transient HTTP codes with backoff.
- Reduce `PageSize` or process in batches (`MaxItems`) if needed.

## Safety and Execution Guidance

- Run against a test target list first.
- Start with small batch (`-MaxItems`) to validate mappings and attachment behavior.
- Keep target list versioning enabled during migration.
- Export source and target snapshots before large runs.

## Suggested GitHub Enhancements

- Add CSV/JSON migration log (source ID -> target ID -> status/error).
- Add idempotency key (skip already migrated source IDs).
- Add field type inference and configurable field mapping overrides.
- Add cleanup helper for legacy auto-created fields.

## Repository Structure (Current Folder)

- `migrate-infopath-xml-to-list-pnp.ps1` - main migration script
- `export-infopath-to-list.ps1` - earlier local extraction helper
- `*.xml` - sample InfoPath XML payloads

## License / Ownership

Update this section in your GitHub repo according to your preferred license and ownership model.
