# InfoPath XML to SharePoint List Migration (PnP App-Only)

Migrate legacy InfoPath XML records from a source SharePoint list/library into a modern target SharePoint list using PnP.PowerShell and app-only authentication.

Script file: `migrate-infopath-xml-to-list-pnp.ps1`

This README reflects the current script behavior, including advanced duplicate detection, full-item replace mode, professional run logging, and large-run performance improvements.

## What the Script Does

1. Connects to source and target SharePoint contexts (can be different sites).
2. Reads source items and resolves InfoPath XML payloads.
3. Parses XML metadata from InfoPath namespace leaf nodes.
4. Optionally auto-creates missing target columns (Text type).
5. Maps metadata into target field values with sanitization.
6. Extracts embedded InfoPath attachments from base64 payloads.
7. Handles duplicates via configurable action and detection strategy.
8. Creates or replaces target items, uploads attachments, and logs a structured run summary.

## Major Features

- App-only auth with Certificate (required).
- Source/target input as list title, GUID, or full URL.
- Cross-site migration support.
- Duplicate actions: `Skip`, `Overwrite`, `CreateNew`.
- Duplicate detection modes:
  - `Title`
  - `MetadataAndAttachments` (mapped metadata + attachment file names)
- Overwrite behavior is full replacement (delete old item, create new item).
- Detailed start and end summary blocks in console and log file.
- Robust retry logic for throttling and transient failures.
- Optimized indexing for larger lists (in-memory duplicate indexes).

## Prerequisites

### 1) PowerShell Module

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

### 2) Entra ID App Registration

- SharePoint application permissions for list read/write and attachments.
- Admin consent granted.
- Certificate available for certificate auth.

Note:

- Certificate auth is required for this migration scenario.
- Client secret auth is not supported for this script when calling PnP/SharePoint APIs.

### 3) Source Content Shape

The script supports InfoPath XML from:

- XML files in a document library, or
- XML attachments on classic list items.

## Parameters

### Connection and Auth

- `SiteUrl` (string)
  - Fallback site URL when list inputs are titles/GUIDs.
- `SourceList` (string)
  - Source list title, GUID, or URL.
- `TargetList` (string)
  - Target list title, GUID, or URL.
- `TenantId` (string)
- `ClientId` (string)
- `AuthType` (`Certificate` required)
- `Thumbprint` (string)
- `CertStore` (`LocalMachine` | `CurrentUser`)
- `ClientSecret` (string, present in script but not supported for this scenario)

### Processing Controls

- `PageSize` (int, default `200`)
- `MaxItems` (int, default `0`, meaning all)
- `TempFolder` (string)
- `FallbackAttachmentName` (string)
- `CreateMetadata` (bool)
  - `true`: create missing target text columns from discovered InfoPath XML fields, including fields that are present but empty in every source XML item.
  - `false`: do not create missing columns, unmapped fields are skipped.
- `SkipAttachments` (switch)
  - If set, no attachment upload is attempted.

### Duplicate Handling

- `Duplicate` (`Overwrite` | `Skip` | `CreateNew`)
  - `Skip`: do nothing when a match is found.
  - `Overwrite`: replace matched item by deleting old and creating new.
  - `CreateNew`: always create a new item.
- `DuplicateDetection` (`Title` | `MetadataAndAttachments`)
  - `Title`: matches by `Title` only.
  - `MetadataAndAttachments`: matches by a generated signature from mapped metadata fields and attachment file names.

### Logging

- `EnableLogging` (bool)
- `DebugLogging` (bool)
- `LogFilePath` (string)

## Duplicate Handling Matrix

### `Duplicate = CreateNew`

- Always creates a new target item.
- Duplicate detection mode is ignored.

### `Duplicate = Skip`

- Finds a match using `DuplicateDetection`.
- Skips source item when matched.

### `Duplicate = Overwrite`

- Finds a match using `DuplicateDetection`.
- Replaces whole list item, not just attachments:
  - remove existing target item
  - create new target item from source metadata
  - upload source attachments

## Logging and Run Output

The script writes:

- A startup configuration block with run settings.
- Per-item logs including source ID, create/replace result, title, and item IDs.
- A completion summary block with totals and final status.

If enabled, the same entries are written to `LogFilePath`.

## Performance and Scale Notes

For large lists (thousands of items), the script improves performance by:

- Building an in-memory duplicate index once per run.
- Precomputing source mapped values and attachments before target writes.

When using `MetadataAndAttachments`, index creation is heavier than title mode because it reads target attachment names.

## Throttling and Retry Behavior

The script retries transient conditions including `429`, `502`, `503`, and `504`, plus throttle-like `403` cases.

Retry wait uses:

- `Retry-After`
- `x-ms-retry-after-ms`
- `RateLimit-Reset`
- exponential backoff + jitter

## Usage Examples

### Certificate Auth

```powershell
.\migrate-infopath-xml-to-list-pnp.ps1 `
  -SiteUrl 'https://contoso.sharepoint.com/sites/Main' `
  -SourceList 'https://contoso.sharepoint.com/sites/Legacy/Shared%20Documents' `
  -TargetList 'https://contoso.sharepoint.com/sites/Modern/Lists/Subscriptions' `
  -TenantId 'tenant-guid' `
  -ClientId 'app-guid' `
  -AuthType Certificate `
  -Thumbprint 'CERT_THUMBPRINT' `
  -CertStore LocalMachine
```

### Advanced Duplicate Detection + Skip

```powershell
.\migrate-infopath-xml-to-list-pnp.ps1 `
  -Duplicate Skip `
  -DuplicateDetection MetadataAndAttachments
```

### Advanced Duplicate Detection + Overwrite (Full Replace)

```powershell
.\migrate-infopath-xml-to-list-pnp.ps1 `
  -Duplicate Overwrite `
  -DuplicateDetection MetadataAndAttachments
```

### Do Not Auto-Create Metadata Columns

```powershell
.\migrate-infopath-xml-to-list-pnp.ps1 -CreateMetadata $false
```

### Metadata-Only (No Attachments)

```powershell
.\migrate-infopath-xml-to-list-pnp.ps1 -SkipAttachments
```

### Small Test Batch

```powershell
.\migrate-infopath-xml-to-list-pnp.ps1 -MaxItems 20
```

## Important Implementation Notes

- Metadata column creation is Text-only today.
- Text values are sanitized for control characters.
- Single-line text fields are flattened and truncated to 255 chars.
- Attachment duplicate matching in `MetadataAndAttachments` uses attachment file names, not file content hashes.

## Troubleshooting

### Many items are skipped with "no InfoPath XML found"

Expected when source library contains non-XML artifacts.

### Invalid text value errors

Some fields may require custom transforms (date/choice/boolean/lookup). Extend mapping logic as needed.

### Duplicate matching appears too broad or too narrow

- Try switching `DuplicateDetection` mode.
- Ensure key metadata fields are mapped into target values before duplicate comparison.

### Throttling or long runtime

- Lower `PageSize`.
- Use `MaxItems` to process in batches.

### Client secret auth fails

- Expected for this migration implementation.
- Use certificate auth with `-AuthType Certificate`.

## Recommended Run Strategy

1. Start with `-MaxItems 20` in a test target list.
2. Confirm metadata mappings and attachment outcomes.
3. Choose duplicate strategy for production:
   - Fast: `Title`
   - Safer for repeated titles: `MetadataAndAttachments`
4. Run full migration in monitored batches.

## Repository Contents

- `migrate-infopath-xml-to-list-pnp.ps1` - main migration script
- `export-infopath-to-list.ps1` - earlier helper script
- `*.xml` - sample payload files
