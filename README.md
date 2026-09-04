# InfoPath XML to SharePoint List Migration (PnP App-Only)

Migrate legacy InfoPath XML records from a source SharePoint list/library into a modern target SharePoint list using PnP.PowerShell and app-only authentication.

Script file: `migrate-infopath-xml-to-list-pnp.ps1`

This README reflects the current script behavior, including source-key duplicate detection, external migration state, schema preflight validation, enhanced rich-text support, and large-run memory protections.

## What the Script Does

1. Connects to source and target SharePoint contexts (can be different sites).
2. Reads source items and resolves InfoPath XML payloads.
3. Parses XML metadata from InfoPath namespace leaf nodes.
4. Reads all XML payloads before target item creation and derives the required target schema.
5. Optionally auto-creates missing target columns as single-line Text, multiple-line Note, or enhanced rich text.
6. Maps metadata into target field values with sanitization.
7. Extracts embedded InfoPath attachments from base64 payloads.
8. Handles duplicates via configurable action and detection strategy.
9. Creates or replaces target items, uploads attachments, and logs a structured run summary.

## Major Features

- App-only auth with Certificate (required).
- Source/target input as list title, GUID, or full URL.
- Cross-site migration support.
- Duplicate actions: `Skip`, `Overwrite`, `CreateNew`.
- Default duplicate detection: `SourceItemKey`.
- Duplicate detection modes:
  - `SourceItemKey` (source site, source list, and immutable source item ID stored in external state)
  - `Title`
  - `MetadataAndAttachments` (mapped metadata + attachment file names)
- Overwrite behavior creates and populates the replacement before deleting the old item.
- External, target-list-specific JSON state with source/target scope validation.
- Resumable item state for incomplete migrations.
- Preflight schema validation before target item creation.
- Rich XML fields are created as enhanced rich text Note columns.
- Plain values over 255 characters are created as multiple-line Note columns.
- Detailed start and end summary blocks in console and log file.
- Robust retry logic for throttling and transient failures.
- Bounded source retrieval and disk-backed XML/attachment staging for large runs.

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
- `BatchSize` (int, default `100`)
- `MaxItems` (int, default `0`, meaning all)
- `TempFolder` (string)
- `FallbackAttachmentName` (string)
- `StateFilePath` (string)
  - Optional external JSON state file. By default, the script creates `InfoPathMigrationState-<target-list>.json`.
- `CreateMetadata` (bool)
  - `true`: create missing target columns based on the complete XML scan. XHTML fields become enhanced rich text; plain fields over 255 characters become multiple-line Note fields; other fields become single-line Text.
  - `false`: do not create missing columns, unmapped fields are skipped.
- `SkipAttachments` (switch)
  - If set, no attachment upload is attempted.

### Duplicate Handling

- `Duplicate` (`Overwrite` | `Skip` | `CreateNew`)
  - `Skip`: do nothing when a match is found.
  - `Overwrite`: replace matched item by deleting old and creating new.
  - `CreateNew`: always create a new item.
- `DuplicateDetection` (`SourceItemKey` | `Title` | `MetadataAndAttachments`)
  - `SourceItemKey`: default. Uses the normalized source site, source list, and immutable source item ID in the external state file. Safe when titles repeat.
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
  - create new target item from source metadata
  - upload source attachments
  - remove existing target item only after the replacement is complete

### External Migration State

When `DuplicateDetection` is `SourceItemKey`, the script does not add a tracking column or tracking value to the target list. It stores mappings in an external JSON file containing:

- Source site and list identity
- Target site and list identity
- Source key to target item ID
- Item status (`Created` or `Completed`)
- Attachment names recorded at completion

State entries are validated against the current source and target. A recorded target item is re-queried before it is skipped; stale entries are not trusted.

## Logging and Run Output

The script writes:

- A startup configuration block with run settings.
- Per-item logs including source ID, create/replace result, title, and item IDs.
- A completion summary block with totals and final status.

If enabled, the same entries are written to `LogFilePath`.

## Performance and Scale Notes

For large lists (thousands of items), the script improves performance by:

- Retrieving source items in bounded ID-ordered batches.
- Staging XML payloads and attachments on disk instead of retaining attachment bytes for the full run.
- Parsing and importing one item at a time after the schema is known.
- Keeping the in-memory work queue to source IDs and staged XML paths.

The script intentionally reads and stages all source XML payloads before creating target items. This allows it to determine the complete target schema first while avoiding retention of XML DOMs and attachment byte arrays across the migration.

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

### Source-Key Duplicate Detection + Skip

```powershell
.\migrate-infopath-xml-to-list-pnp.ps1 `
  -Duplicate Skip `
  -DuplicateDetection SourceItemKey
```

### Metadata Duplicate Detection + Overwrite (Full Replace)

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

- InfoPath fields containing XHTML are migrated to enhanced rich text (`Note`) columns with `RichText=TRUE` and `RichTextMode=FullHtml`.
- Other auto-created metadata columns use single-line text.
- Plain-text fields whose source values exceed 255 characters are created as multiple-line Note columns.
- Text values are sanitized for control characters.
- Existing incompatible target columns are reported during preflight and stop the run before target items are written.
- Attachment duplicate matching in `MetadataAndAttachments` uses attachment file names, not file content hashes.

## Troubleshooting

### Many items are skipped with "no InfoPath XML found"

Expected when source library contains non-XML artifacts.

### Invalid text value errors

The script performs schema preflight before writing target items. If an existing target column is incompatible with the XML-derived requirement, correct the column type or use a new target list, then rerun. The script does not silently truncate or omit the incompatible value.

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
3. Use the default `SourceItemKey` duplicate strategy for production. It is safe when titles repeat and does not modify the target list schema.
4. Keep the generated target-specific state JSON file for reruns and back it up.
5. Run the full migration with an appropriate `BatchSize` for the available memory.

## Repository Contents

- `migrate-infopath-xml-to-list-pnp.ps1` - main migration script
- `export-infopath-to-list.ps1` - earlier helper script
- `*.xml` - sample payload files
