<#
.SYNOPSIS
  Migrates InfoPath XML-based items from a legacy SharePoint list/library to a modern SharePoint list.

.DESCRIPTION
  Uses PnP.PowerShell with app-only authentication to:
  1. Read source items from a legacy list or form library.
  2. Resolve each item's InfoPath XML payload.
  3. Parse XML leaf nodes into a metadata hashtable.
  4. Map XML field names to target list internal names (same-name mapping).
  5. Create a new target list item with mapped metadata.
  6. Extract embedded InfoPath attachments from XML and upload them to the new list item.

  This script is designed as a practical starting point for migration and can be extended
  with custom field transforms as needed.

.PREREQUISITES
  - PnP.PowerShell installed: Install-Module PnP.PowerShell -Scope CurrentUser
  - Entra ID app registration with app-only permissions for SharePoint Online
  - App auth configured with either certificate thumbprint or client secret
  - Admin consent granted for required permissions

.NOTES
  - Default mapping uses target field internal names matched to InfoPath XML element names.
  - For rich text nodes (embedded XHTML), text content is extracted.
  - For values that look like embedded attachment payloads, metadata mapping skips those nodes.

  Author: Mike Lee
  Created: 7/17/2026
  
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
  [Parameter()]
  [string]$SiteUrl= 'https://m365cpi13246019.sharepoint.com',

  [Parameter()]
  [string]$SourceList = 'https://m365cpi13246019.sharepoint.com/sites/SPSite1/Shared%20Documents',

  [Parameter()]
  [string]$TargetList = 'https://m365cpi13246019.sharepoint.com/sites/SPSite2/Lists/TargetList',

  [Parameter()]
  [string]$TenantId = '9cfc42cb-51da-4055-87e9-b20a170b6ba3',

  [Parameter()]
  [string]$ClientId = 'abc64618-283f-47ba-a185-50d935d51d57',

  [Parameter()]
  [ValidateSet('Certificate', 'ClientSecret')]
  [string]$AuthType = 'Certificate',

  [Parameter()]
  [string]$Thumbprint = 'B696FDCFE1453F3FBC6031F54DE988DA0ED905A9',

  [Parameter()]
  [ValidateSet('LocalMachine', 'CurrentUser')]
  [string]$CertStore = 'LocalMachine',

  [Parameter()]
  [string]$ClientSecret = $env:PNP_CLIENT_SECRET,

  [Parameter()]
  [int]$PageSize = 200,

  [Parameter()]
  [int]$MaxItems = 0,

  [Parameter()]
  [string]$TempFolder = (Join-Path $env:TEMP 'InfoPathMigrationTemp'),

  [Parameter()]
  [string]$FallbackAttachmentName = 'uploadedFile.bin',

  [Parameter()]
  [switch]$SkipAttachments
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Use explicit validation here to avoid interactive prompts caused by Mandatory parameters.
$requiredParams = @{
  SiteUrl = $SiteUrl
  SourceList = $SourceList
  TargetList = $TargetList
  TenantId = $TenantId
  ClientId = $ClientId
}
foreach ($entry in $requiredParams.GetEnumerator()) {
  if ([string]::IsNullOrWhiteSpace([string]$entry.Value)) {
    throw "Required parameter '$($entry.Key)' is empty. Provide a value in the param block or at runtime."
  }
}

if ($AuthType -eq 'Certificate' -and [string]::IsNullOrWhiteSpace($Thumbprint)) {
  throw "AuthType 'Certificate' requires a non-empty Thumbprint."
}

if ($AuthType -eq 'ClientSecret' -and [string]::IsNullOrWhiteSpace($ClientSecret)) {
  throw "AuthType 'ClientSecret' requires -ClientSecret (or env var PNP_CLIENT_SECRET)."
}

$script:stats = [ordered]@{
  SourceItemsRead = 0
  XmlResolved = 0
  TargetItemsCreated = 0
  AttachmentsUploaded = 0
  Skipped = 0
  Failed = 0
}

function Write-Info {
  param([string]$Message)
  Write-Host $Message -ForegroundColor Cyan
}

function Write-Warn {
  param([string]$Message)
  Write-Host $Message -ForegroundColor Yellow
}

function Write-Err {
  param([string]$Message)
  Write-Host $Message -ForegroundColor Red
}

function Invoke-PnPWithRetry {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [scriptblock]$ScriptBlock,
    [Parameter()] [int]$MaxRetries = 10,
    [Parameter()] [int]$InitialBackoffSeconds = 2
  )

  $retryCount = 0
  $backoffSec = $InitialBackoffSeconds

  while ($retryCount -le $MaxRetries) {
    try {
      return & $ScriptBlock
    }
    catch {
      $statusCode = $null

      # Under StrictMode, some exception types do not expose .Response.
      $responseProp = $_.Exception.PSObject.Properties['Response']
      if ($null -ne $responseProp -and $null -ne $responseProp.Value) {
        try { $statusCode = [int]$responseProp.Value.StatusCode } catch {}
      }

      if (-not $statusCode -and $_.Exception.Message -match '(429|502|503|504)') {
        $statusCode = [int]$Matches[1]
      }

      $isRetryable = $statusCode -in @(429, 502, 503, 504)
      if (-not $isRetryable) { throw }

      if ($retryCount -ge $MaxRetries) { throw }

      $waitSec = $backoffSec
      $retryCount++
      Write-Warn "Throttled ($statusCode). Waiting ${waitSec}s (attempt $retryCount/$MaxRetries)..."
      Start-Sleep -Seconds $waitSec
      $backoffSec = [Math]::Min($backoffSec * 2, 300)
    }
  }
}

function Connect-ToPnPSite {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [string]$Url
  )

  if ($AuthType -eq 'Certificate') {
    if ([string]::IsNullOrWhiteSpace($Thumbprint)) {
      throw 'AuthType=Certificate requires -Thumbprint.'
    }
    if (-not (Test-Path "Cert:\$CertStore\My\$Thumbprint")) {
      throw "Certificate $Thumbprint not found in Cert:\$CertStore\My"
    }

    Invoke-PnPWithRetry {
      Connect-PnPOnline -Url $Url -Tenant $TenantId -ClientId $ClientId -Thumbprint $Thumbprint -ErrorAction Stop
    }
  }
  else {
    if ([string]::IsNullOrWhiteSpace($ClientSecret)) {
      throw 'AuthType=ClientSecret requires -ClientSecret or PNP_CLIENT_SECRET environment variable.'
    }

    Invoke-PnPWithRetry {
      Connect-PnPOnline -Url $Url -ClientId $ClientId -ClientSecret $ClientSecret -ErrorAction Stop
    }
  }

  $web = Invoke-PnPWithRetry { Get-PnPWeb }
  Write-Info "Connected to: $($web.Url)"
}

function Test-IsLikelyBase64 {
  [CmdletBinding()]
  param([Parameter(Mandatory)] [string]$Text)

  if ($Text.Length -lt 100) { return $false }
  if (($Text.Length % 4) -ne 0) { return $false }
  if ($Text.Contains(' ')) { return $false }
  if ($Text -match '^https?://') { return $false }

  try {
    [void][Convert]::FromBase64String($Text)
    return $true
  }
  catch {
    return $false
  }
}

function Resolve-InfoPathXmlFromItem {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)]$Item,
    [Parameter(Mandatory)] [string]$ListTitle
  )

  $fv = $Item.FieldValues

  # Common case for form libraries: item file itself is XML
  if ($fv.ContainsKey('FileRef') -and $fv['FileRef']) {
    $fileRef = [string]$fv['FileRef']
    if ($fileRef.ToLowerInvariant().EndsWith('.xml')) {
      return Invoke-PnPWithRetry {
        Get-PnPFile -Url $fileRef -AsString
      }
    }

    # In document libraries, non-XML items (folders/files) are not InfoPath payloads.
    return $null
  }

  # Fallback: classic list item with attachments containing an XML file
  try {
    # Expand AttachmentFiles on the CSOM list item without any interactive prompts.
    Invoke-PnPWithRetry {
      Get-PnPProperty -ClientObject $Item -Property AttachmentFiles | Out-Null
    }

    $attachments = @($Item.AttachmentFiles)

    $xmlAttachment = $attachments | Where-Object {
      $_.FileName -like '*.xml' -or $_.ServerRelativeUrl -like '*.xml'
    } | Select-Object -First 1

    if ($xmlAttachment) {
      return Invoke-PnPWithRetry {
        Get-PnPFile -Url $xmlAttachment.ServerRelativeUrl -AsString
      }
    }
  }
  catch {
    # Ignore attachment lookup failures and return null below.
  }

  return $null
}

function Convert-InfoPathXmlToFieldMap {
  [CmdletBinding()]
  param([Parameter(Mandatory)] [xml]$XmlDoc)

  $nsMgr = New-Object System.Xml.XmlNamespaceManager($XmlDoc.NameTable)
  $myNamespace = 'http://schemas.microsoft.com/office/infopath/2003/myXSD/2017-07-28T18:04:07'
  $nsMgr.AddNamespace('my', $myNamespace)

  $map = @{}

  # Leaf elements only: we map by local name so matching works against target internal names.
  $leafNodes = $XmlDoc.SelectNodes('//*[not(*)]', $nsMgr)

  foreach ($node in $leafNodes) {
    if (-not $node) { continue }

    # Ignore embedded XHTML nodes (div, a, br, font, strong, etc.).
    if ([string]$node.NamespaceURI -ne $myNamespace) { continue }

    $key = [string]$node.LocalName
    if ([string]::IsNullOrWhiteSpace($key)) { continue }

    $raw = [string]$node.InnerText
    if ([string]::IsNullOrWhiteSpace($raw)) { continue }

    # Skip likely encoded attachment payloads from metadata mapping.
    if (Test-IsLikelyBase64 -Text $raw) { continue }

    $value = $raw.Trim()

    # Keep first non-empty value when duplicates exist (common in repeating sections).
    if (-not $map.ContainsKey($key)) {
      $map[$key] = $value
    }
  }

  return $map
}

function Get-InfoPathAttachmentsFromXml {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [xml]$XmlDoc,
    [Parameter()] [string]$BasicFileName = 'uploadedFile.bin'
  )

  $results = New-Object System.Collections.Generic.List[object]
  $nameCounter = @{}

  foreach ($textNode in $XmlDoc.SelectNodes('//*[text()]')) {
    $text = [string]$textNode.InnerText
    if (-not (Test-IsLikelyBase64 -Text $text)) { continue }

    $bytes = $null
    try {
      $bytes = [Convert]::FromBase64String($text)
    }
    catch {
      continue
    }

    if ($null -eq $bytes -or $bytes.Length -eq 0) { continue }

    $fileName = $BasicFileName
    $payload = $bytes

    # InfoPath attachment signature: C7 49 46 41
    $hasInfoPathHeader = ($bytes.Length -gt 24 -and $bytes[0] -eq 199 -and $bytes[1] -eq 73 -and $bytes[2] -eq 70 -and $bytes[3] -eq 65)
    if ($hasInfoPathHeader) {
      $nameByteLen = [int]$bytes[20] * 2
      $headerLen = 24

      if ($nameByteLen -gt 0 -and ($headerLen + $nameByteLen) -lt $bytes.Length) {
        $nameBytes = $bytes[$headerLen..($headerLen + $nameByteLen - 1)]
        try {
          $decodedName = [System.Text.Encoding]::Unicode.GetString($nameBytes)
          $decodedName = $decodedName.TrimEnd([char]0).Trim()
          if (-not [string]::IsNullOrWhiteSpace($decodedName)) {
            $fileName = $decodedName
          }
        }
        catch {
          # Keep fallback name.
        }

        $contentStart = $headerLen + $nameByteLen
        if ($contentStart -lt $bytes.Length) {
          $payload = $bytes[$contentStart..($bytes.Length - 1)]
        }
      }
    }

    $safeName = [IO.Path]::GetFileName($fileName)
    if ([string]::IsNullOrWhiteSpace($safeName)) {
      $safeName = $BasicFileName
    }

    # Ensure unique names in one item.
    if (-not $nameCounter.ContainsKey($safeName)) {
      $nameCounter[$safeName] = 0
      $finalName = $safeName
    }
    else {
      $nameCounter[$safeName]++
      $base = [IO.Path]::GetFileNameWithoutExtension($safeName)
      $ext = [IO.Path]::GetExtension($safeName)
      $finalName = "{0}-copy{1}{2}" -f $base, $nameCounter[$safeName], $ext
    }

    $results.Add([pscustomobject]@{
      FileName = $finalName
      Bytes = $payload
      NodeName = [string]$textNode.LocalName
    })
  }

  return $results
}

function Get-TargetWritableFields {
  [CmdletBinding()]
  param([Parameter(Mandatory)] [string]$ListTitle)

  $fields = Invoke-PnPWithRetry { Get-PnPField -List $ListTitle }

  # Exclude non-writable/internal system fields.
  $blocked = @(
    'ID', 'GUID', 'Attachments', 'Author', 'Editor', 'Created', 'Modified',
    'ContentType', 'ContentTypeId', 'FileLeafRef', 'FileRef', 'FileDirRef',
    'MetaInfo', '_UIVersionString', '_ModerationStatus'
  )

  return $fields | Where-Object {
    -not $_.ReadOnlyField -and
    -not $_.Hidden -and
    $_.InternalName -notin $blocked -and
    $_.Sealed -ne $true
  }
}

function New-TargetFieldNameMap {
  [CmdletBinding()]
  param([Parameter(Mandatory)] [object[]]$TargetFields)

  $map = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)

  foreach ($field in $TargetFields) {
    $internal = [string]$field.InternalName
    if (-not [string]::IsNullOrWhiteSpace($internal) -and -not $map.ContainsKey($internal)) {
      $map.Add($internal, $internal)
    }

    $display = [string]$field.Title
    if (-not [string]::IsNullOrWhiteSpace($display) -and -not $map.ContainsKey($display)) {
      $map.Add($display, $internal)
    }
  }

  return $map
}

function Convert-ToSafeInternalFieldName {
  [CmdletBinding()]
  param([Parameter(Mandatory)] [string]$InputName)

  $name = $InputName.Trim()
  $name = $name -replace '[^A-Za-z0-9_]', '_'
  $name = $name -replace '_+', '_'
  $name = $name.Trim('_')

  if ([string]::IsNullOrWhiteSpace($name)) {
    $name = "field_$([guid]::NewGuid().ToString('N').Substring(0,8))"
  }

  if ($name -match '^[0-9]') {
    $name = "f_$name"
  }

  if ($name.Length -gt 32) {
    $name = $name.Substring(0, 32)
  }

  return $name
}

function Ensure-TargetFieldsForXmlKeys {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [string]$ListTitle,
    [Parameter(Mandatory)] [string[]]$XmlKeys,
    [Parameter(Mandatory)] [object[]]$TargetFields
  )

  $fieldMap = New-TargetFieldNameMap -TargetFields $TargetFields
  $created = 0

  foreach ($xmlKey in $XmlKeys) {
    if ([string]::IsNullOrWhiteSpace($xmlKey)) { continue }
    if ($fieldMap.ContainsKey($xmlKey)) { continue }

    $baseInternal = Convert-ToSafeInternalFieldName -InputName $xmlKey
    $candidate = $baseInternal
    $suffix = 0

    while ($fieldMap.ContainsKey($candidate)) {
      $suffix++
      $base = if ($baseInternal.Length -gt 28) { $baseInternal.Substring(0, 28) } else { $baseInternal }
      $candidate = "{0}_{1}" -f $base, $suffix
    }

    Write-Warn "Target column '$xmlKey' does not exist. Creating it as Text (internal: $candidate)."
    Invoke-PnPWithRetry {
      Add-PnPField -List $ListTitle -DisplayName $xmlKey -InternalName $candidate -Type Text -Group 'InfoPath Migrated Columns' | Out-Null
    }

    $created++
    $fieldMap[$xmlKey] = $candidate
    if (-not $fieldMap.ContainsKey($candidate)) {
      $fieldMap[$candidate] = $candidate
    }
  }

  if ($created -gt 0) {
    Write-Info "Created $created missing target column(s)."
    $TargetFields = @(Get-TargetWritableFields -ListTitle $ListTitle)
    $fieldMap = New-TargetFieldNameMap -TargetFields $TargetFields
  }

  return [pscustomobject]@{
    TargetFields = $TargetFields
    FieldMap = $fieldMap
    CreatedCount = $created
  }
}

function New-TargetItemValues {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [hashtable]$XmlMap,
    [Parameter(Mandatory)] [System.Collections.Generic.Dictionary[string,string]]$FieldNameMap,
    [Parameter(Mandatory)] [System.Collections.Generic.Dictionary[string,object]]$FieldByInternalName,
    [Parameter(Mandatory)]$SourceItem
  )

  $values = @{}

  foreach ($xmlKey in $XmlMap.Keys) {
    if ($FieldNameMap.ContainsKey([string]$xmlKey)) {
      $targetInternalName = $FieldNameMap[[string]$xmlKey]
      $rawValue = [string]$XmlMap[$xmlKey]
      if (-not [string]::IsNullOrWhiteSpace($rawValue)) {
        $sanitized = $rawValue

        # Remove control chars that SharePoint text fields reject.
        $sanitized = [regex]::Replace($sanitized, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')

        $fieldDef = $null
        if ($FieldByInternalName.ContainsKey($targetInternalName)) {
          $fieldDef = $FieldByInternalName[$targetInternalName]
        }

        $fieldType = if ($null -ne $fieldDef -and $fieldDef.PSObject.Properties['TypeAsString']) {
          [string]$fieldDef.TypeAsString
        }
        else {
          ''
        }

        # Single-line text fields cannot contain line breaks and are max 255 chars.
        if ($fieldType -eq 'Text') {
          $sanitized = $sanitized -replace "`r`n|`n|`r", ' '
          if ($sanitized.Length -gt 255) {
            $sanitized = $sanitized.Substring(0, 255)
          }
        }

        if (-not [string]::IsNullOrWhiteSpace($sanitized)) {
          $values[$targetInternalName] = $sanitized
        }
      }
    }
  }

  # Ensure required Title exists for standard custom lists.
  if (-not $values.ContainsKey('Title') -or [string]::IsNullOrWhiteSpace([string]$values['Title'])) {
    if ($XmlMap.ContainsKey('VendorName') -and -not [string]::IsNullOrWhiteSpace([string]$XmlMap['VendorName'])) {
      $values['Title'] = [string]$XmlMap['VendorName']
    }
    elseif ($XmlMap.ContainsKey('fileName') -and -not [string]::IsNullOrWhiteSpace([string]$XmlMap['fileName'])) {
      $values['Title'] = [string]$XmlMap['fileName']
    }
    else {
      $values['Title'] = "Migrated InfoPath Item $($SourceItem.Id)"
    }
  }

  return $values
}

function Add-ExtractedAttachmentsToItem {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [string]$ListTitle,
    [Parameter(Mandatory)] [int]$ItemId,
    [Parameter(Mandatory)] [object[]]$Attachments,
    [Parameter(Mandatory)] [string]$WorkingFolder
  )

  if (-not (Test-Path -Path $WorkingFolder -PathType Container)) {
    New-Item -Path $WorkingFolder -ItemType Directory -Force | Out-Null
  }

  foreach ($att in $Attachments) {
    $filePath = Join-Path $WorkingFolder $att.FileName
    [IO.File]::WriteAllBytes($filePath, $att.Bytes)

    Invoke-PnPWithRetry {
      Add-PnPListItemAttachment -List $ListTitle -Identity $ItemId -Path $filePath -NewFileName $att.FileName | Out-Null
    }

    $script:stats.AttachmentsUploaded++
    Remove-Item -Path $filePath -Force -ErrorAction SilentlyContinue
  }
}

function Resolve-ListContext {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [string]$ListInput,
    [Parameter(Mandatory)] [string]$FallbackSiteUrl
  )

  $trimmed = $ListInput.Trim()
  if ($trimmed -notmatch '^https?://') {
    return [pscustomobject]@{
      SiteUrl = $FallbackSiteUrl
      ListIdentity = $trimmed
      Original = $ListInput
    }
  }

  $uri = [Uri]$trimmed
  $path = [Uri]::UnescapeDataString($uri.AbsolutePath)

  $sitePath = ''
  $listIdentity = ''

  if ($path -match '^/(sites|teams)/([^/]+)(/.*)?$') {
    $sitePath = "/$($Matches[1])/$($Matches[2])"
    $remaining = if ($Matches[3]) { $Matches[3].Trim('/') } else { '' }

    if ($remaining -match '^Lists/([^/]+)') {
      $listIdentity = [Uri]::UnescapeDataString($Matches[1])
    }
    elseif (-not [string]::IsNullOrWhiteSpace($remaining)) {
      $listIdentity = [Uri]::UnescapeDataString(($remaining -split '/')[0])
    }
  }
  else {
    # Root site path shapes such as /Lists/MyList or /Shared Documents
    $remaining = $path.Trim('/')
    if ($remaining -match '^Lists/([^/]+)') {
      $listIdentity = [Uri]::UnescapeDataString($Matches[1])
    }
    elseif (-not [string]::IsNullOrWhiteSpace($remaining)) {
      $listIdentity = [Uri]::UnescapeDataString(($remaining -split '/')[0])
    }
  }

  if ([string]::IsNullOrWhiteSpace($listIdentity)) {
    throw "Could not derive list identity from '$ListInput'. Use a list title/ID or a list URL."
  }

  $resolvedSiteUrl = "{0}://{1}{2}" -f $uri.Scheme, $uri.Host, $sitePath
  return [pscustomobject]@{
    SiteUrl = $resolvedSiteUrl
    ListIdentity = $listIdentity
    Original = $ListInput
  }
}

try {
  if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    throw 'PnP.PowerShell module not found. Install with: Install-Module PnP.PowerShell -Scope CurrentUser'
  }

  Import-Module PnP.PowerShell -ErrorAction Stop

  $sourceCtx = Resolve-ListContext -ListInput $SourceList -FallbackSiteUrl $SiteUrl
  $targetCtx = Resolve-ListContext -ListInput $TargetList -FallbackSiteUrl $SiteUrl

  Write-Info "Source context -> Site: $($sourceCtx.SiteUrl) | List: $($sourceCtx.ListIdentity)"
  Write-Info "Target context -> Site: $($targetCtx.SiteUrl) | List: $($targetCtx.ListIdentity)"

  Write-Info 'Connecting to source context...'
  Connect-ToPnPSite -Url $sourceCtx.SiteUrl

  $source = Invoke-PnPWithRetry {
    Get-PnPListItem -List $sourceCtx.ListIdentity -PageSize $PageSize
  }

  if ($MaxItems -gt 0) {
    $source = $source | Select-Object -First $MaxItems
  }

  $script:stats.SourceItemsRead = @($source).Count
  Write-Info "Source items to process: $($script:stats.SourceItemsRead)"

  if ($script:stats.SourceItemsRead -eq 0) {
    Write-Warn 'No source items found. Exiting.'
    return
  }

  $migrationQueue = New-Object System.Collections.Generic.List[object]

  $readIndex = 0
  foreach ($item in $source) {
    $readIndex++
    Write-Host "[Source $readIndex/$($script:stats.SourceItemsRead)] Source item ID $($item.Id)" -ForegroundColor White

    try {
      $xmlText = Resolve-InfoPathXmlFromItem -Item $item -ListTitle $sourceCtx.ListIdentity
      if ([string]::IsNullOrWhiteSpace($xmlText)) {
        $script:stats.Skipped++
        Write-Warn "  Skipped: no InfoPath XML found for source item $($item.Id)."
        continue
      }

      $cleanXml = $xmlText.Replace('§', '')
      [xml]$xmlDoc = $cleanXml
      $script:stats.XmlResolved++

      $migrationQueue.Add([pscustomobject]@{
        SourceId = $item.Id
        XmlDoc = $xmlDoc
      })
    }
    catch {
      $script:stats.Failed++
      Write-Err "  Failed reading source item $($item.Id): $($_.Exception.Message)"
    }
  }

  if ($migrationQueue.Count -eq 0) {
    Write-Warn 'No valid XML payloads were resolved from source. Exiting.'
    return
  }

  Write-Info 'Connecting to target context...'
  Connect-ToPnPSite -Url $targetCtx.SiteUrl

  $targetFields = @(Get-TargetWritableFields -ListTitle $targetCtx.ListIdentity)
  Write-Info "Writable target fields discovered: $($targetFields.Count)"

  $allXmlKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($work in $migrationQueue) {
    $xmlMapForItem = Convert-InfoPathXmlToFieldMap -XmlDoc $work.XmlDoc
    Add-Member -InputObject $work -MemberType NoteProperty -Name XmlMap -Value $xmlMapForItem -Force
    foreach ($k in $xmlMapForItem.Keys) {
      $null = $allXmlKeys.Add([string]$k)
    }
  }

  $schemaSync = Ensure-TargetFieldsForXmlKeys -ListTitle $targetCtx.ListIdentity -XmlKeys @($allXmlKeys) -TargetFields $targetFields
  $targetFields = $schemaSync.TargetFields
  $targetFieldMap = $schemaSync.FieldMap

  $targetFieldByInternalName = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($f in $targetFields) {
    $internal = [string]$f.InternalName
    if (-not [string]::IsNullOrWhiteSpace($internal) -and -not $targetFieldByInternalName.ContainsKey($internal)) {
      $targetFieldByInternalName[$internal] = $f
    }
  }

  if (-not (Test-Path -Path $TempFolder -PathType Container)) {
    New-Item -Path $TempFolder -ItemType Directory -Force | Out-Null
  }

  $i = 0
  foreach ($work in $migrationQueue) {
    $i++
    Write-Host "[Target $i/$($migrationQueue.Count)] Source item ID $($work.SourceId)" -ForegroundColor White

    try {
      $values = New-TargetItemValues -XmlMap $work.XmlMap -FieldNameMap $targetFieldMap -FieldByInternalName $targetFieldByInternalName -SourceItem ([pscustomobject]@{ Id = $work.SourceId })

      if (-not $PSCmdlet.ShouldProcess("Target list '$($targetCtx.ListIdentity)'", "Create item from source ID $($work.SourceId)")) {
        continue
      }

      $newItem = Invoke-PnPWithRetry {
        Add-PnPListItem -List $targetCtx.ListIdentity -Values $values
      }
      $script:stats.TargetItemsCreated++

      if (-not $SkipAttachments) {
        $attachments = @(Get-InfoPathAttachmentsFromXml -XmlDoc $work.XmlDoc -BasicFileName $FallbackAttachmentName)
        if ($attachments.Count -gt 0) {
          Add-ExtractedAttachmentsToItem -ListTitle $targetCtx.ListIdentity -ItemId $newItem.Id -Attachments $attachments -WorkingFolder $TempFolder
          Write-Host "  Attachments uploaded: $($attachments.Count)" -ForegroundColor Green
        }
      }

      Write-Host "  Created target item ID $($newItem.Id)" -ForegroundColor Green
    }
    catch {
      $script:stats.Failed++
      Write-Err "  Failed source item $($work.SourceId): $($_.Exception.Message)"
    }
  }

  Write-Host ''
  Write-Host 'Migration complete' -ForegroundColor Green
  Write-Host "  Source items read:     $($script:stats.SourceItemsRead)"
  Write-Host "  XML resolved:          $($script:stats.XmlResolved)"
  Write-Host "  Target items created:  $($script:stats.TargetItemsCreated)"
  Write-Host "  Attachments uploaded:  $($script:stats.AttachmentsUploaded)"
  Write-Host "  Skipped:               $($script:stats.Skipped)"
  Write-Host "  Failed:                $($script:stats.Failed)"
}
catch {
  Write-Err "Script failed: $($_.Exception.Message)"
  throw
}
finally {
  try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch {}
}
