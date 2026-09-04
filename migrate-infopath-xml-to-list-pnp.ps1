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
  - Rich text nodes (embedded XHTML) are preserved as HTML in rich multiple-line text columns.
  - For values that look like embedded attachment payloads, metadata mapping skips those nodes.

  Author: Mike Lee
  Created: 7/17/2026
  Updated: 7/21/2026 - Added duplicate detection and handling, logging, and throttling support.
  Updated: 7/24/2026 - Fixed bug that would not create columns if they were empty.
  Updated: 9/3/2026 - Preserve XHTML fields as rich multiple-line text columns.
  
#>

[CmdletBinding(SupportsShouldProcess = $true)]
#region Parameters
param(
  [Parameter()]
  [string]$SiteUrl= 'https://m365cpi13246019.sharepoint.com',

  [Parameter()]
  [string]$SourceList = 'https://m365cpi13246019.sharepoint.com/sites/SPSite2/InfoPath',

  [Parameter()]
  [string]$TargetList = 'https://m365cpi13246019.sharepoint.com/sites/SPSite2/Lists/TargetList10',

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
  [ValidateRange(1, 1000)]
  [int]$BatchSize = 100,

  [Parameter()]
  [int]$MaxItems = 0,

  [Parameter()]
  [string]$TempFolder = (Join-Path $env:TEMP 'InfoPathMigrationTemp'),

  [Parameter()]
  [string]$FallbackAttachmentName = 'uploadedFile.bin',

  [Parameter()]
  [bool]$CreateMetadata = $true,

  [Parameter()]
  [ValidateSet('Overwrite', 'Skip', 'CreateNew')]
  [string]$Duplicate = 'Skip',

  [Parameter()]
  [ValidateSet('SourceItemKey', 'Title', 'MetadataAndAttachments')]
  [string]$DuplicateDetection = 'SourceItemKey',

  [Parameter()]
  [bool]$EnableLogging = $true,

  [Parameter()]
  [bool]$DebugLogging = $false,

  [Parameter()]
  [string]$LogFilePath = (Join-Path $PSScriptRoot ("InfoPathMigration-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))),

  [Parameter()]
  [string]$StateFilePath = '',

  [Parameter()]
  [switch]$SkipAttachments
)
#endregion Parameters

#region Script Setup
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:RunLog = $null
$script:LoggingEnabled = $false
$script:DebugLoggingEnabled = $DebugLogging
$script:runSucceeded = $false
$script:runStartedUtc = [datetime]::UtcNow
$stagingFolder = $null
$migrationState = @{}

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
#endregion Script Setup

#region Logging Helpers
function Write-Log {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [AllowEmptyString()] [string]$Message,
    [Parameter()] [ValidateSet('INFO', 'SUCCESS', 'WARN', 'ERROR', 'DEBUG')] [string]$Level = 'INFO'
  )

  $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
  $entry = "$ts [$Level] $Message"

  if ($script:LoggingEnabled -and -not [string]::IsNullOrWhiteSpace($script:RunLog)) {
    try {
      Add-Content -Path $script:RunLog -Value $entry -Encoding UTF8
    }
    catch {
      Write-Host "Failed writing to log file '$($script:RunLog)': $($_.Exception.Message)" -ForegroundColor Red
    }
  }

  switch ($Level) {
    'SUCCESS' { Write-Host $Message -ForegroundColor Green }
    'WARN' { Write-Host $Message -ForegroundColor Yellow }
    'ERROR' { Write-Host $Message -ForegroundColor Red }
    'DEBUG' {
      if ($script:DebugLoggingEnabled) {
        Write-Host $Message -ForegroundColor DarkGray
      }
    }
    default { Write-Host $Message -ForegroundColor Cyan }
  }
}

function Write-Info {
  param([string]$Message)
  Write-Log -Message $Message -Level INFO
}

function Write-Warn {
  param([string]$Message)
  Write-Log -Message $Message -Level WARN
}

function Write-Err {
  param([string]$Message)
  Write-Log -Message $Message -Level ERROR
}

function Get-DuplicateModeDescription {
  [CmdletBinding()]
  param([Parameter(Mandatory)] [string]$Mode)

  switch ($Mode.ToLowerInvariant()) {
    'skip' { return 'If a matching item exists, skip source item.' }
    'overwrite' { return 'If a matching item exists, replace existing target item with a new item from source.' }
    'createnew' { return 'Always create a new target item.' }
    default { return 'Unknown duplicate handling behavior.' }
  }
}

function Get-DuplicateDetectionDescription {
  [CmdletBinding()]
  param([Parameter(Mandatory)] [string]$Mode)

  switch ($Mode.ToLowerInvariant()) {
    'sourceitemkey' { return 'Duplicate match by source list identity and immutable source item ID.' }
    'title' { return 'Duplicate match by Title only (fastest).' }
    'metadataandattachments' { return 'Duplicate match by mapped metadata fields + attachment file names.' }
    default { return 'Unknown duplicate detection mode.' }
  }
}

function Normalize-DuplicateMode {
  [CmdletBinding()]
  param([Parameter(Mandatory)] [string]$Mode)

  switch ($Mode.ToLowerInvariant()) {
    'skip' { return 'Skip' }
    'overwrite' { return 'Overwrite' }
    'createnew' { return 'CreateNew' }
    default { return $Mode }
  }
}

function Start-RunLogging {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [bool]$Enabled,
    [Parameter(Mandatory)] [string]$Path
  )

  if (-not $Enabled) {
    $script:LoggingEnabled = $false
    Write-Host 'File logging is disabled for this run.' -ForegroundColor Yellow
    return
  }

  if ([string]::IsNullOrWhiteSpace($Path)) {
    $script:LoggingEnabled = $false
    Write-Host 'LogFilePath is empty. File logging is disabled for this run.' -ForegroundColor Yellow
    return
  }

  $script:RunLog = $Path
  $script:LoggingEnabled = $true
  $script:DebugLoggingEnabled = $DebugLogging

  $logFolder = Split-Path -Path $script:RunLog -Parent
  if (-not [string]::IsNullOrWhiteSpace($logFolder) -and -not (Test-Path -Path $logFolder -PathType Container)) {
    New-Item -Path $logFolder -ItemType Directory -Force | Out-Null
  }

  try {
    if (-not (Test-Path -Path $script:RunLog -PathType Leaf)) {
      New-Item -Path $script:RunLog -ItemType File -Force | Out-Null
    }

    Write-Log -Message "Logging to file: $script:RunLog" -Level INFO
    Write-Log -Message ('-' * 80) -Level INFO
  }
  catch {
    $script:LoggingEnabled = $false
    Write-Host "Could not initialize file logging at '$script:RunLog'. Continuing without file logging. Error: $($_.Exception.Message)" -ForegroundColor Yellow
  }
}

function Write-RunConfigurationSummary {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] $SourceContext,
    [Parameter(Mandatory)] $TargetContext
  )

  $runId = [guid]::NewGuid().ToString('N').Substring(0, 8).ToUpperInvariant()
  $startedLocal = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  $separator = ('=' * 92)

  Write-Info ''
  Write-Info $separator
  Write-Log -Level SUCCESS -Message 'INFO PATH MIGRATION - RUN CONFIGURATION'
  Write-Info $separator
  Write-Info ("  Run ID             : {0}" -f $runId)
  Write-Info ("  Started            : {0}" -f $startedLocal)
  Write-Info ("  Source             : {0} | {1}" -f $SourceContext.SiteUrl, $SourceContext.ListIdentity)
  Write-Info ("  Target             : {0} | {1}" -f $TargetContext.SiteUrl, $TargetContext.ListIdentity)
  Write-Info ("  AuthType           : {0}" -f $AuthType)
  Write-Info ("  PageSize           : {0}" -f $PageSize)
  Write-Info ("  BatchSize          : {0}" -f $BatchSize)
  Write-Info ("  MaxItems           : {0}" -f $MaxItems)
  Write-Info ("  CreateMetadata     : {0}" -f $CreateMetadata)
  Write-Info ("  Duplicate Mode     : {0}" -f $Duplicate)
  Write-Info ("  Duplicate Action   : {0}" -f (Get-DuplicateModeDescription -Mode $Duplicate))
  Write-Info ("  Duplicate Detect   : {0}" -f $DuplicateDetection)
  Write-Info ("  Detect Description : {0}" -f (Get-DuplicateDetectionDescription -Mode $DuplicateDetection))
  Write-Info ("  SkipAttachments    : {0}" -f $SkipAttachments)
  Write-Info ("  EnableLogging      : {0}" -f $EnableLogging)
  Write-Info ("  DebugLogging       : {0}" -f $DebugLogging)
  if ($EnableLogging) {
    Write-Info ("  LogFilePath        : {0}" -f $LogFilePath)
  }
  Write-Info ("  StateFilePath      : {0}" -f $StateFilePath)
  Write-Info $separator
  Write-Info ''
}

function Load-MigrationState {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [string]$Path,
    [Parameter(Mandatory)] [string]$SourceSite,
    [Parameter(Mandatory)] [string]$SourceList,
    [Parameter(Mandatory)] [string]$TargetSite,
    [Parameter(Mandatory)] [string]$TargetList
  )

  $state = [ordered]@{
    SourceSite = $SourceSite
    SourceList = $SourceList
    TargetSite = $TargetSite
    TargetList = $TargetList
    Items = @{}
  }
  if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $state }

  $json = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json
  $jsonProperties = $json.PSObject.Properties
  foreach ($name in @('SourceSite', 'SourceList', 'TargetSite', 'TargetList')) {
    if ($null -eq $jsonProperties[$name]) {
      throw "Migration state '$Path' uses an older format without scope metadata. Rename or remove it before rerunning."
    }
  }
  foreach ($name in @('SourceSite', 'SourceList', 'TargetSite', 'TargetList')) {
    if ([string]$json.$name -ne [string]$state[$name]) {
      throw "Migration state '$Path' belongs to a different source or target. Use a separate StateFilePath."
    }
  }
  if ($null -eq $jsonProperties['Items'] -or $null -eq $json.Items) {
    throw "Migration state '$Path' uses an older format without scope metadata. Rename or remove it before rerunning."
  }
  foreach ($property in $json.Items.PSObject.Properties) {
    $entry = [ordered]@{ TargetId = 0; Status = 'Created'; AttachmentNames = @() }
    if ($null -ne $property.Value.TargetId) { $entry.TargetId = [int]$property.Value.TargetId }
    if (-not [string]::IsNullOrWhiteSpace([string]$property.Value.Status)) { $entry.Status = [string]$property.Value.Status }
    if ($null -ne $property.Value.AttachmentNames) { $entry.AttachmentNames = @($property.Value.AttachmentNames | ForEach-Object { [string]$_ }) }
    if ($entry.TargetId -gt 0) { $state.Items[[string]$property.Name] = $entry }
  }
  return $state
}

function Save-MigrationState {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [string]$Path,
    [Parameter(Mandatory)] $State
  )

  $folder = Split-Path -Path $Path -Parent
  if (-not [string]::IsNullOrWhiteSpace($folder) -and -not (Test-Path -LiteralPath $folder -PathType Container)) {
    New-Item -Path $folder -ItemType Directory -Force | Out-Null
  }
  $tempPath = "$Path.tmp"
  $State | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $tempPath -Encoding UTF8
  Move-Item -LiteralPath $tempPath -Destination $Path -Force
}

function Get-TargetItemById {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [string]$ListTitle,
    [Parameter(Mandatory)] [int]$ItemId
  )

  try {
    return Invoke-PnPWithRetry { Get-PnPListItem -List $ListTitle -Id $ItemId }
  }
  catch {
    return $null
  }
}
#endregion Logging Helpers

#region Throttling Helpers
function Get-HeaderValue {
  [CmdletBinding()]
  param(
    [Parameter()] $Headers,
    [Parameter(Mandatory)] [string]$Name
  )

  if ($null -eq $Headers) { return $null }

  try {
    $value = $Headers[$Name]
    if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
      return [string]$value
    }
  }
  catch {}

  return $null
}

function Get-ThrottleWaitSecondsFromHeaders {
  [CmdletBinding()]
  param(
    [Parameter()] $Headers,
    [Parameter()] [int]$DefaultSeconds = 1
  )

  $retryAfterSec = $null
  $rateResetSec = $null

  $retryAfter = Get-HeaderValue -Headers $Headers -Name 'Retry-After'
  if (-not [string]::IsNullOrWhiteSpace($retryAfter)) {
    $intVal = 0
    if ([int]::TryParse($retryAfter, [ref]$intVal)) {
      $retryAfterSec = [Math]::Max($intVal, 0)
    }
    else {
      $dtVal = [datetime]::MinValue
      if ([datetime]::TryParse($retryAfter, [ref]$dtVal)) {
        $retryAfterSec = [Math]::Max([int][Math]::Ceiling(($dtVal.ToUniversalTime() - [datetime]::UtcNow).TotalSeconds), 0)
      }
    }
  }

  $retryAfterMs = Get-HeaderValue -Headers $Headers -Name 'x-ms-retry-after-ms'
  if (-not [string]::IsNullOrWhiteSpace($retryAfterMs)) {
    $msVal = 0
    if ([int]::TryParse($retryAfterMs, [ref]$msVal)) {
      $msSec = [int][Math]::Ceiling($msVal / 1000.0)
      if ($null -eq $retryAfterSec -or $msSec -gt $retryAfterSec) {
        $retryAfterSec = [Math]::Max($msSec, 0)
      }
    }
  }

  $rateReset = Get-HeaderValue -Headers $Headers -Name 'RateLimit-Reset'
  if (-not [string]::IsNullOrWhiteSpace($rateReset)) {
    $resetVal = 0
    if ([int]::TryParse($rateReset, [ref]$resetVal)) {
      # In SharePoint guidance this is seconds until refill.
      $rateResetSec = [Math]::Max($resetVal, 0)
    }
  }

  $candidates = @()
  if ($null -ne $retryAfterSec) { $candidates += $retryAfterSec }
  if ($null -ne $rateResetSec) { $candidates += $rateResetSec }

  if ($candidates.Count -gt 0) {
    # Microsoft guidance: use the greater of Retry-After and RateLimit-Reset when both exist.
    return ($candidates | Measure-Object -Maximum).Maximum
  }

  return [Math]::Max($DefaultSeconds, 1)
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
      $headers = $null
      $requestId = $null

      # Under StrictMode, some exception types do not expose .Response.
      $responseProp = $_.Exception.PSObject.Properties['Response']
      if ($null -ne $responseProp -and $null -ne $responseProp.Value) {
        try { $statusCode = [int]$responseProp.Value.StatusCode } catch {}
        try { $headers = $responseProp.Value.Headers } catch {}
      }

      if ($null -ne $headers) {
        $requestId = Get-HeaderValue -Headers $headers -Name 'SPRequestGuid'
        if ([string]::IsNullOrWhiteSpace($requestId)) {
          $requestId = Get-HeaderValue -Headers $headers -Name 'request-id'
        }
      }

      if (-not $statusCode -and $_.Exception.Message -match '(429|502|503|504)') {
        $statusCode = [int]$Matches[1]
      }

      $looksThrottleLike = $_.Exception.Message -match '(throttl|too many requests|server too busy|try again)'
      $isRetryable = ($statusCode -in @(429, 502, 503, 504)) -or ($statusCode -eq 403 -and $looksThrottleLike)
      if (-not $isRetryable) { throw }

      if ($retryCount -ge $MaxRetries) { throw }

      $headerWaitSec = Get-ThrottleWaitSecondsFromHeaders -Headers $headers -DefaultSeconds $backoffSec
      $waitSec = [Math]::Max($backoffSec, $headerWaitSec)

      # Add small jitter to avoid synchronized retry storms.
      $jitterMs = Get-Random -Minimum 200 -Maximum 1200
      $waitSec = [Math]::Min($waitSec + ($jitterMs / 1000.0), 900)

      $retryCount++
      $retryAfterHeader = if ($null -ne $headers) { Get-HeaderValue -Headers $headers -Name 'Retry-After' } else { $null }
      $rateResetHeader = if ($null -ne $headers) { Get-HeaderValue -Headers $headers -Name 'RateLimit-Reset' } else { $null }

      Write-Warn ("Throttled (HTTP {0}). Waiting {1:n1}s (attempt {2}/{3}). Retry-After='{4}' RateLimit-Reset='{5}' RequestId='{6}'" -f `
        $statusCode, $waitSec, $retryCount, $MaxRetries, $retryAfterHeader, $rateResetHeader, $requestId)

      Start-Sleep -Milliseconds ([int][Math]::Ceiling($waitSec * 1000.0))
      $backoffSec = [Math]::Min($backoffSec * 2, 300)
    }
  }
}
#endregion Throttling Helpers

#region Connection and Source Helpers
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

function Get-SourceItemBatch {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [string]$ListTitle,
    [Parameter(Mandatory)] [int]$AfterId,
    [Parameter(Mandatory)] [int]$BatchSize
  )

  $caml = @"
<View Scope='RecursiveAll'>
  <Query>
    <Where>
      <Gt><FieldRef Name='ID' /><Value Type='Counter'>$AfterId</Value></Gt>
    </Where>
    <OrderBy><FieldRef Name='ID' Ascending='TRUE' /></OrderBy>
  </Query>
  <ViewFields><FieldRef Name='ID' /><FieldRef Name='FileRef' /><FieldRef Name='Attachments' /></ViewFields>
  <RowLimit>$BatchSize</RowLimit>
</View>
"@

  return @(Invoke-PnPWithRetry {
    Get-PnPListItem -List $ListTitle -Query $caml
  })
}
#endregion Connection and Source Helpers

#region XML and Attachment Parsing
function Convert-InfoPathXmlToFieldMap {
  [CmdletBinding()]
  param([Parameter(Mandatory)] [xml]$XmlDoc)

  $nsMgr = New-Object System.Xml.XmlNamespaceManager($XmlDoc.NameTable)
  $myNamespace = 'http://schemas.microsoft.com/office/infopath/2003/myXSD/2017-07-28T18:04:07'
  $nsMgr.AddNamespace('my', $myNamespace)

  $map = @{}

  # Include simple leaf fields and fields containing XHTML, but exclude structural my:* groups.
  $fieldNodes = $XmlDoc.SelectNodes('//my:*[not(my:*)]', $nsMgr)

  foreach ($node in $fieldNodes) {
    if (-not $node) { continue }

    $key = [string]$node.LocalName
    if ([string]::IsNullOrWhiteSpace($key)) { continue }

    $xhtmlRoot = @($node.ChildNodes | Where-Object {
      $_.NodeType -eq [System.Xml.XmlNodeType]::Element -and
      [string]$_.NamespaceURI -eq 'http://www.w3.org/1999/xhtml'
    }) | Select-Object -First 1

    $raw = if ($null -ne $xhtmlRoot) {
      [string]$xhtmlRoot.InnerXml
    }
    else {
      [string]$node.InnerText
    }
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

function Get-InfoPathXmlMetadataKeys {
  [CmdletBinding()]
  param([Parameter(Mandatory)] [xml]$XmlDoc)

  $nsMgr = New-Object System.Xml.XmlNamespaceManager($XmlDoc.NameTable)
  $myNamespace = 'http://schemas.microsoft.com/office/infopath/2003/myXSD/2017-07-28T18:04:07'
  $nsMgr.AddNamespace('my', $myNamespace)

  $keys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  $fieldNodes = $XmlDoc.SelectNodes('//my:*[not(my:*)]', $nsMgr)

  foreach ($node in $fieldNodes) {
    if (-not $node) { continue }

    $key = [string]$node.LocalName
    if ([string]::IsNullOrWhiteSpace($key)) { continue }

    # AttachmentControl contains the embedded file payload; attachments are uploaded separately.
    if ($key -eq 'AttachmentControl') { continue }

    $null = $keys.Add($key)
  }

  return @($keys)
}

function Get-InfoPathXmlRichTextKeys {
  [CmdletBinding()]
  param([Parameter(Mandatory)] [xml]$XmlDoc)

  $nsMgr = New-Object System.Xml.XmlNamespaceManager($XmlDoc.NameTable)
  $nsMgr.AddNamespace('my', 'http://schemas.microsoft.com/office/infopath/2003/myXSD/2017-07-28T18:04:07')
  $nsMgr.AddNamespace('xhtml', 'http://www.w3.org/1999/xhtml')

  $keys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($node in $XmlDoc.SelectNodes('//my:*[xhtml:*]', $nsMgr)) {
    if ($node -and -not [string]::IsNullOrWhiteSpace([string]$node.LocalName)) {
      $null = $keys.Add([string]$node.LocalName)
    }
  }

  return @($keys)
}

function Get-InfoPathAttachmentsFromXml {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [xml]$XmlDoc,
    [Parameter()] [string]$BasicFileName = 'uploadedFile.bin',
    [Parameter()] [string]$OutputFolder
  )

  $results = New-Object System.Collections.Generic.List[object]
  $nameCounter = @{}

  foreach ($textNode in $XmlDoc.SelectNodes('//*[text() and not(*)]')) {
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
        $nameBytes = New-Object byte[] $nameByteLen
        [Buffer]::BlockCopy($bytes, $headerLen, $nameBytes, 0, $nameByteLen)
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
          $payload = New-Object byte[] ($bytes.Length - $contentStart)
          [Buffer]::BlockCopy($bytes, $contentStart, $payload, 0, $payload.Length)
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

    $filePath = $null
    if (-not [string]::IsNullOrWhiteSpace($OutputFolder)) {
      if (-not (Test-Path -Path $OutputFolder -PathType Container)) {
        New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
      }
      $filePath = Join-Path $OutputFolder $finalName
      [IO.File]::WriteAllBytes($filePath, $payload)
    }

    $results.Add([pscustomobject]@{
      FileName = $finalName
      Bytes = if ($null -eq $filePath) { $payload } else { $null }
      FilePath = $filePath
      NodeName = [string]$textNode.LocalName
    })
  }

  return $results
}
#endregion XML and Attachment Parsing

#region Target Schema and Mapping
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

function Assert-TargetSchemaCompatibility {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [string[]]$XmlKeys,
    [Parameter(Mandatory)] [object[]]$TargetFields,
    [Parameter(Mandatory)] [string[]]$RichTextKeys,
    [Parameter(Mandatory)] [string[]]$LongTextKeys
  )

  $fieldMap = New-TargetFieldNameMap -TargetFields $TargetFields
  $richTextSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  $longTextSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($key in $RichTextKeys) { if (-not [string]::IsNullOrWhiteSpace($key)) { $null = $richTextSet.Add($key) } }
  foreach ($key in $LongTextKeys) { if (-not [string]::IsNullOrWhiteSpace($key)) { $null = $longTextSet.Add($key) } }

  $incompatible = New-Object System.Collections.Generic.List[string]
  foreach ($xmlKey in $XmlKeys) {
    if (-not $fieldMap.ContainsKey($xmlKey)) { continue }
    $internalName = [string]$fieldMap[$xmlKey]
    $field = @($TargetFields | Where-Object { [string]$_.InternalName -eq $internalName }) | Select-Object -First 1
    if ($null -eq $field) { continue }

    $type = [string]$field.TypeAsString
    $richTextProperty = $field.PSObject.Properties['RichText']
    $isRich = $null -ne $richTextProperty -and ([string]$richTextProperty.Value) -match '^(?i:true)$'
    if ($richTextSet.Contains($xmlKey) -and ($type -ne 'Note' -or -not $isRich)) {
      $incompatible.Add("$xmlKey requires enhanced rich text (Note/RichText), but target field '$internalName' is $type.")
    }
    elseif ($longTextSet.Contains($xmlKey) -and $type -eq 'Text') {
      $incompatible.Add("$xmlKey contains values over 255 characters, but target field '$internalName' is single-line Text.")
    }
  }

  if ($incompatible.Count -gt 0) {
    throw "Target schema is incompatible with the XML data. No target items were written:`n$($incompatible -join "`n")"
  }
}

function Ensure-TargetFieldsForXmlKeys {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [string]$ListTitle,
    [Parameter(Mandatory)] [string[]]$XmlKeys,
    [Parameter(Mandatory)] [object[]]$TargetFields,
    [Parameter()] [object[]]$AllTargetFields = @(),
    [Parameter()] [string[]]$RichTextKeys = @(),
    [Parameter()] [string[]]$LongTextKeys = @(),
    [Parameter()] [bool]$CreateMetadata = $true
  )

  $fieldMap = New-TargetFieldNameMap -TargetFields $TargetFields
  $allFieldMap = if (@($AllTargetFields).Count -gt 0) {
    New-TargetFieldNameMap -TargetFields $AllTargetFields
  }
  else {
    New-TargetFieldNameMap -TargetFields $TargetFields
  }
  $created = 0
  $missing = 0
  $richTextKeySet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($richTextKey in $RichTextKeys) {
    if (-not [string]::IsNullOrWhiteSpace($richTextKey)) {
      $null = $richTextKeySet.Add($richTextKey)
    }
  }
  $longTextKeySet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($longTextKey in $LongTextKeys) {
    if (-not [string]::IsNullOrWhiteSpace($longTextKey)) {
      $null = $longTextKeySet.Add($longTextKey)
    }
  }
  $longTextMappings = @{}
  $richTextMappings = @{}

  foreach ($xmlKey in $XmlKeys) {
    if ([string]::IsNullOrWhiteSpace($xmlKey)) { continue }
    if ($fieldMap.ContainsKey($xmlKey)) {
      $existingInternal = [string]$fieldMap[$xmlKey]
      $existingField = @($TargetFields | Where-Object { [string]$_.InternalName -eq $existingInternal }) | Select-Object -First 1
      if ($richTextKeySet.Contains($xmlKey)) {
        $richTextProperty = if ($null -ne $existingField) { $existingField.PSObject.Properties['RichText'] } else { $null }
        if ($null -ne $existingField -and [string]$existingField.TypeAsString -eq 'Note' -and $null -ne $richTextProperty -and ([string]$richTextProperty.Value) -match '^(?i:true)$') { continue }
        Write-Warn "Target field '$xmlKey' already exists but is not enhanced rich text. No alternate column will be created; change that existing column to enhanced rich text before rerunning."
        continue
      }
      elseif (-not $longTextKeySet.Contains($xmlKey)) { continue }
      if ($null -eq $existingField -or [string]$existingField.TypeAsString -ne 'Text') { continue }
      Write-Warn "Target field '$xmlKey' already exists as single-line Text and contains values over 255 characters. No alternate column will be created; change '$xmlKey' to a multiple-line text column before rerunning."
      continue
    }

    if (-not $CreateMetadata) {
      $missing++
      continue
    }

    $candidate = [string]$xmlKey

    if ($richTextKeySet.Contains($xmlKey)) {
      Write-Warn "Target field '$xmlKey' is rich text. Creating enhanced rich text Note field '$candidate'."
      $escapedDisplayName = [System.Security.SecurityElement]::Escape($xmlKey)
      $escapedInternalName = [System.Security.SecurityElement]::Escape($candidate)
      $fieldXml = "<Field Type='Note' DisplayName='$escapedDisplayName' Name='$escapedInternalName' StaticName='$escapedInternalName' NumLines='20' RichText='TRUE' RichTextMode='FullHtml' IsolateStyles='TRUE' AppendOnly='FALSE' Group='InfoPath Migrated Columns' />"
      Invoke-PnPWithRetry {
        Add-PnPFieldFromXml -List $ListTitle -FieldXml $fieldXml | Out-Null
      }
      $richTextMappings[$xmlKey] = $candidate
    }
    elseif ($longTextKeySet.Contains($xmlKey)) {
      Write-Warn "Target field '$xmlKey' contains values over 255 characters. Creating multiple-line Note field '$candidate'."
      $escapedDisplayName = [System.Security.SecurityElement]::Escape($xmlKey)
      $escapedInternalName = [System.Security.SecurityElement]::Escape($candidate)
      $fieldXml = "<Field Type='Note' DisplayName='$escapedDisplayName' Name='$escapedInternalName' StaticName='$escapedInternalName' NumLines='20' RichText='FALSE' AppendOnly='FALSE' Group='InfoPath Migrated Columns' />"
      Invoke-PnPWithRetry {
        Add-PnPFieldFromXml -List $ListTitle -FieldXml $fieldXml | Out-Null
      }
      $longTextMappings[$xmlKey] = $candidate
    }
    else {
      Write-Warn "Target column '$xmlKey' does not exist. Creating it as Text (internal: $candidate)."
      Invoke-PnPWithRetry {
        Add-PnPField -List $ListTitle -DisplayName $xmlKey -InternalName $candidate -Type Text -Group 'InfoPath Migrated Columns' | Out-Null
      }
    }

    $created++
    $fieldMap[$xmlKey] = $candidate
    $fieldMap[$candidate] = $candidate
    $allFieldMap[$xmlKey] = $candidate
    $allFieldMap[$candidate] = $candidate
  }

  if ($missing -gt 0) {
    Write-Warn "Metadata auto-creation disabled. $missing XML field(s) were not mapped because target columns do not exist."
  }

  if ($created -gt 0) {
    Write-Info "Created $created missing target column(s)."
    $TargetFields = @(Get-TargetWritableFields -ListTitle $ListTitle)
    $fieldMap = New-TargetFieldNameMap -TargetFields $TargetFields
    foreach ($mapping in $longTextMappings.GetEnumerator()) {
      $fieldMap[$mapping.Key] = $mapping.Value
    }
    foreach ($mapping in $richTextMappings.GetEnumerator()) {
      $fieldMap[$mapping.Key] = $mapping.Value
    }
  }

  return [pscustomobject]@{
    TargetFields = $TargetFields
    FieldMap = $fieldMap
    CreatedCount = $created
    MissingCount = $missing
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

        # Keep the source value intact; SharePoint will reject an incompatible target field rather than silently losing data.
        if ($fieldType -eq 'Text') {
          $sanitized = $sanitized -replace "`r`n|`n|`r", ' '
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
#endregion Target Schema and Mapping

#region Target Item and List Utilities
function Add-ExtractedAttachmentsToItem {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [string]$ListTitle,
    [Parameter(Mandatory)] [int]$ItemId,
    [Parameter(Mandatory)] [object[]]$Attachments,
    [Parameter(Mandatory)] [string]$WorkingFolder,
    [Parameter()] [ValidateSet('Overwrite', 'Skip', 'CreateNew')] [string]$DuplicateMode = 'CreateNew'
  )

  if (-not (Test-Path -Path $WorkingFolder -PathType Container)) {
    New-Item -Path $WorkingFolder -ItemType Directory -Force | Out-Null
  }

  $existingAttachmentNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  try {
    $item = Invoke-PnPWithRetry {
      Get-PnPListItem -List $ListTitle -Id $ItemId
    }

    Invoke-PnPWithRetry {
      Get-PnPProperty -ClientObject $item -Property AttachmentFiles | Out-Null
    }

    foreach ($existing in @($item.AttachmentFiles)) {
      if ($null -ne $existing -and -not [string]::IsNullOrWhiteSpace([string]$existing.FileName)) {
        [void]$existingAttachmentNames.Add([string]$existing.FileName)
      }
    }
  }
  catch {
    Write-Warn "Could not pre-read existing attachments for item ID $ItemId. Continuing upload. Error: $($_.Exception.Message)"
  }

  foreach ($att in $Attachments) {
    if ($existingAttachmentNames.Contains($att.FileName)) {
      if ($DuplicateMode -eq 'Overwrite') {
        Invoke-PnPWithRetry {
          Remove-PnPListItemAttachment -List $ListTitle -Identity $ItemId -FileName $att.FileName -Force | Out-Null
        }
        [void]$existingAttachmentNames.Remove($att.FileName)
        Write-Info "  Replaced existing attachment '$($att.FileName)' on item ID $ItemId"
      }
      elseif ($DuplicateMode -eq 'Skip') {
        Write-Warn "  Skipped existing attachment '$($att.FileName)' on item ID $ItemId"
        if ($att.PSObject.Properties['FilePath'] -and (Test-Path -LiteralPath $att.FilePath -PathType Leaf)) {
          Remove-Item -LiteralPath $att.FilePath -Force -ErrorAction SilentlyContinue
        }
        continue
      }
    }

    $filePath = $null
    if ($att.PSObject.Properties['FilePath'] -and -not [string]::IsNullOrWhiteSpace([string]$att.FilePath)) {
      $filePath = [string]$att.FilePath
    }
    else {
      $filePath = Join-Path $WorkingFolder $att.FileName
      [IO.File]::WriteAllBytes($filePath, $att.Bytes)
    }

    try {
      Add-PnPListItemAttachment -List $ListTitle -Identity $ItemId -Path $filePath -NewFileName $att.FileName -ErrorAction Stop | Out-Null
      [void]$existingAttachmentNames.Add($att.FileName)
      $script:stats.AttachmentsUploaded++
    }
    finally {
      Remove-Item -LiteralPath $filePath -Force -ErrorAction SilentlyContinue
    }
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

function Escape-CamlValue {
  [CmdletBinding()]
  param([Parameter(Mandatory)] [string]$Value)

  return [System.Security.SecurityElement]::Escape($Value)
}

function Get-TargetTitleIndex {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [string]$ListTitle,
    [Parameter()] [int]$PageSize = 200
  )

  $index = New-Object 'System.Collections.Generic.Dictionary[string,int]' ([System.StringComparer]::OrdinalIgnoreCase)

  $items = Invoke-PnPWithRetry {
    Get-PnPListItem -List $ListTitle -PageSize $PageSize -Fields @('Title')
  }

  foreach ($it in @($items)) {
    $title = ''
    if ($null -ne $it -and $null -ne $it.FieldValues -and $it.FieldValues.ContainsKey('Title')) {
      $title = [string]$it.FieldValues['Title']
    }

    if ([string]::IsNullOrWhiteSpace($title)) { continue }
    if (-not $index.ContainsKey($title)) {
      $index[$title] = [int]$it.Id
    }
  }

  return $index
}

function Normalize-SignatureValue {
  [CmdletBinding()]
  param([Parameter(Mandatory)] [string]$Value)

  $v = $Value.Trim()
  $v = $v -replace '\s+', ' '
  return $v.ToLowerInvariant()
}

function Convert-BytesToHexString {
  [CmdletBinding()]
  param([Parameter(Mandatory)] [byte[]]$Bytes)

  $sb = New-Object System.Text.StringBuilder
  foreach ($b in $Bytes) {
    [void]$sb.Append($b.ToString('x2'))
  }
  return $sb.ToString()
}

function New-DuplicateSignature {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [hashtable]$Values,
    [Parameter()] [string[]]$AttachmentNames = @()
  )

  $fieldParts = New-Object System.Collections.Generic.List[string]
  $keys = @($Values.Keys | ForEach-Object { [string]$_ } | Sort-Object)
  foreach ($k in $keys) {
    $raw = [string]$Values[$k]
    if ([string]::IsNullOrWhiteSpace($raw)) { continue }
    $normalizedFieldName = $k.ToLowerInvariant()
    $normalizedFieldValue = Normalize-SignatureValue -Value $raw
    $fieldParts.Add($normalizedFieldName + '=' + $normalizedFieldValue)
  }

  $attachmentParts = @($AttachmentNames | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { Normalize-SignatureValue -Value ([string]$_) } | Sort-Object -Unique)

  if ($fieldParts.Count -eq 0 -and $attachmentParts.Count -eq 0) {
    return $null
  }

  $payload = "FIELDS:`n" + ($fieldParts -join "`n") + "`nATTACHMENTS:`n" + ($attachmentParts -join "`n")
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
  $sha = [System.Security.Cryptography.SHA256]::Create()
  try {
    $hash = $sha.ComputeHash($bytes)
  }
  finally {
    $sha.Dispose()
  }

  return Convert-BytesToHexString -Bytes $hash
}

function Get-TargetDuplicateSignatureIndex {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [string]$ListTitle,
    [Parameter(Mandatory)] [string[]]$SignatureFieldNames,
    [Parameter()] [int]$PageSize = 200,
    [Parameter()] [bool]$IncludeAttachmentNames = $true
  )

  $index = New-Object 'System.Collections.Generic.Dictionary[string,int]' ([System.StringComparer]::OrdinalIgnoreCase)
  $fieldSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($name in @($SignatureFieldNames)) {
    if (-not [string]::IsNullOrWhiteSpace([string]$name)) {
      $null = $fieldSet.Add([string]$name)
    }
  }

  # Title is always included to improve matching robustness when source/target schemas vary.
  $null = $fieldSet.Add('Title')
  $fieldsToFetch = @($fieldSet)

  $items = Invoke-PnPWithRetry {
    Get-PnPListItem -List $ListTitle -PageSize $PageSize -Fields $fieldsToFetch
  }

  foreach ($it in @($items)) {
    $vals = @{}
    foreach ($fieldName in $fieldsToFetch) {
      if ($null -ne $it -and $null -ne $it.FieldValues -and $it.FieldValues.ContainsKey($fieldName)) {
        $v = [string]$it.FieldValues[$fieldName]
        if (-not [string]::IsNullOrWhiteSpace($v)) {
          $vals[$fieldName] = $v
        }
      }
    }

    $attachmentNames = @()
    if ($IncludeAttachmentNames) {
      try {
        Invoke-PnPWithRetry {
          Get-PnPProperty -ClientObject $it -Property AttachmentFiles | Out-Null
        }
        $attachmentNames = @($it.AttachmentFiles | ForEach-Object { [string]$_.FileName })
      }
      catch {
        Write-Warn "Could not read attachments for target item ID $($it.Id) while building duplicate signature index. Continuing. Error: $($_.Exception.Message)"
      }
    }

    $signature = New-DuplicateSignature -Values $vals -AttachmentNames $attachmentNames
    if ([string]::IsNullOrWhiteSpace($signature)) { continue }
    if (-not $index.ContainsKey($signature)) {
      $index[$signature] = [int]$it.Id
    }
  }

  return $index
}

function Get-ExistingTargetItemByTitle {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [string]$ListTitle,
    [Parameter(Mandatory)] [string]$Title
  )

  if ([string]::IsNullOrWhiteSpace($Title)) {
    return $null
  }

  $safeTitle = Escape-CamlValue -Value $Title
  $caml = @"
<View>
  <Query>
    <Where>
      <Eq>
        <FieldRef Name='Title' />
        <Value Type='Text'>$safeTitle</Value>
      </Eq>
    </Where>
  </Query>
  <RowLimit>1</RowLimit>
</View>
"@

  $existing = Invoke-PnPWithRetry {
    Get-PnPListItem -List $ListTitle -Query $caml
  }

  return @($existing) | Select-Object -First 1
}
#endregion Target Item and List Utilities

#region Main Execution
try {
  if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    throw 'PnP.PowerShell module not found. Install with: Install-Module PnP.PowerShell -Scope CurrentUser'
  }

  Import-Module PnP.PowerShell -ErrorAction Stop

  $sourceCtx = Resolve-ListContext -ListInput $SourceList -FallbackSiteUrl $SiteUrl
  $targetCtx = Resolve-ListContext -ListInput $TargetList -FallbackSiteUrl $SiteUrl
  $Duplicate = Normalize-DuplicateMode -Mode $Duplicate

  if ([string]::IsNullOrWhiteSpace($StateFilePath)) {
    $stateFileName = ($targetCtx.ListIdentity -replace '[^A-Za-z0-9._-]', '_')
    if ([string]::IsNullOrWhiteSpace($stateFileName)) {
      $stateFileName = 'TargetList'
    }
    $StateFilePath = Join-Path $PSScriptRoot ("InfoPathMigrationState-{0}.json" -f $stateFileName)
  }

  Start-RunLogging -Enabled $EnableLogging -Path $LogFilePath
  Write-RunConfigurationSummary -SourceContext $sourceCtx -TargetContext $targetCtx

  Write-Info "Source context -> Site: $($sourceCtx.SiteUrl) | List: $($sourceCtx.ListIdentity)"
  Write-Info "Target context -> Site: $($targetCtx.SiteUrl) | List: $($targetCtx.ListIdentity)"

  Write-Info 'Connecting to source context...'
  Connect-ToPnPSite -Url $sourceCtx.SiteUrl

  if (-not (Test-Path -Path $TempFolder -PathType Container)) {
    New-Item -Path $TempFolder -ItemType Directory -Force | Out-Null
  }
  $stagingFolder = Join-Path $TempFolder ("run-{0}" -f ([guid]::NewGuid().ToString('N')))
  New-Item -Path $stagingFolder -ItemType Directory -Force | Out-Null

  $migrationQueue = New-Object System.Collections.Generic.List[object]
  $allXmlSchemaKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  $allXmlRichTextKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  $allLongTextKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  $lastSourceId = 0
  $sourceBatchNumber = 0

  while ($true) {
    $sourceBatchNumber++
    $sourceBatch = @(Get-SourceItemBatch -ListTitle $sourceCtx.ListIdentity -AfterId $lastSourceId -BatchSize $BatchSize)
    if ($sourceBatch.Count -eq 0) { break }

    foreach ($item in $sourceBatch) {
      $lastSourceId = [int]$item.Id
      if ($MaxItems -gt 0 -and $script:stats.SourceItemsRead -ge $MaxItems) { break }
      $script:stats.SourceItemsRead++
      Write-Info "[Source $($script:stats.SourceItemsRead)] Source item ID $($item.Id) (batch $sourceBatchNumber)"

      try {
        $xmlText = Resolve-InfoPathXmlFromItem -Item $item -ListTitle $sourceCtx.ListIdentity
        if ([string]::IsNullOrWhiteSpace($xmlText)) {
          $script:stats.Skipped++
          Write-Warn "  Skipped: no InfoPath XML found for source item $($item.Id)."
          continue
        }

        $xmlPath = Join-Path $stagingFolder ("source-{0}.xml" -f $item.Id)
        [IO.File]::WriteAllText($xmlPath, $xmlText.Replace('§', ''), [Text.Encoding]::UTF8)
        $migrationQueue.Add([pscustomobject]@{ SourceId = [int]$item.Id; XmlPath = $xmlPath })
        $script:stats.XmlResolved++
      }
      catch {
        $script:stats.Failed++
        Write-Err "  Failed reading source item $($item.Id): $($_.Exception.Message)"
      }
    }

    if (($MaxItems -gt 0 -and $script:stats.SourceItemsRead -ge $MaxItems) -or $sourceBatch.Count -lt $BatchSize) { break }
  }

  Write-Info "Source items to process: $($script:stats.SourceItemsRead)"
  if ($migrationQueue.Count -eq 0) {
    Write-Warn 'No valid XML payloads were resolved from source. Exiting.'
    return
  }

  Write-Info 'Connecting to target context...'
  Connect-ToPnPSite -Url $targetCtx.SiteUrl

  $allTargetFields = @(Invoke-PnPWithRetry { Get-PnPField -List $targetCtx.ListIdentity })
  $targetFields = @(Get-TargetWritableFields -ListTitle $targetCtx.ListIdentity)
  Write-Info "Writable target fields discovered: $($targetFields.Count)"
  $migrationState = Load-MigrationState -Path $StateFilePath -SourceSite $sourceCtx.SiteUrl -SourceList $sourceCtx.ListIdentity -TargetSite $targetCtx.SiteUrl -TargetList $targetCtx.ListIdentity

  foreach ($work in $migrationQueue) {
    [xml]$xmlDoc = [IO.File]::ReadAllText($work.XmlPath)
    $xmlMapForItem = Convert-InfoPathXmlToFieldMap -XmlDoc $xmlDoc
    foreach ($k in (Get-InfoPathXmlMetadataKeys -XmlDoc $xmlDoc)) { $null = $allXmlSchemaKeys.Add([string]$k) }
    foreach ($k in (Get-InfoPathXmlRichTextKeys -XmlDoc $xmlDoc)) { $null = $allXmlRichTextKeys.Add([string]$k) }
    foreach ($entry in $xmlMapForItem.GetEnumerator()) {
      if (-not $allXmlRichTextKeys.Contains([string]$entry.Key) -and [string]$entry.Value -and ([string]$entry.Value).Length -gt 255) { $null = $allLongTextKeys.Add([string]$entry.Key) }
    }
    $xmlDoc = $null
  }

  Write-Info "CreateMetadata is set to: $CreateMetadata"
  Write-Info "Duplicate mode is set to: $Duplicate"
  Write-Info "InfoPath XML schema fields discovered: $($allXmlSchemaKeys.Count)"
  Assert-TargetSchemaCompatibility -XmlKeys @($allXmlSchemaKeys) -TargetFields $targetFields -RichTextKeys @($allXmlRichTextKeys) -LongTextKeys @($allLongTextKeys)
  $schemaSync = Ensure-TargetFieldsForXmlKeys -ListTitle $targetCtx.ListIdentity -XmlKeys @($allXmlSchemaKeys) -TargetFields $targetFields -AllTargetFields $allTargetFields -RichTextKeys @($allXmlRichTextKeys) -LongTextKeys @($allLongTextKeys) -CreateMetadata $CreateMetadata
  $targetFields = $schemaSync.TargetFields
  $targetFieldMap = $schemaSync.FieldMap

  $targetFieldByInternalName = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($f in $targetFields) {
    $internal = [string]$f.InternalName
    if (-not [string]::IsNullOrWhiteSpace($internal) -and -not $targetFieldByInternalName.ContainsKey($internal)) {
      $targetFieldByInternalName[$internal] = $f
    }
  }

  $allSignatureFieldNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($field in $targetFields) {
    if (-not [string]::IsNullOrWhiteSpace([string]$field.InternalName)) {
      $null = $allSignatureFieldNames.Add([string]$field.InternalName)
    }
  }

  $targetDuplicateIndex = $null
  if ($Duplicate -ne 'CreateNew') {
    if ($DuplicateDetection -eq 'SourceItemKey') {
      Write-Info "Loaded migration state entries: $($migrationState.Count)"
      $targetDuplicateIndex = $migrationState
    }
    elseif ($DuplicateDetection -eq 'MetadataAndAttachments') {
      $includeAttachmentNames = -not $SkipAttachments
      Write-Info 'Building in-memory duplicate signature index (metadata + attachments)...'
      $targetDuplicateIndex = Get-TargetDuplicateSignatureIndex -ListTitle $targetCtx.ListIdentity -SignatureFieldNames @($allSignatureFieldNames) -PageSize $PageSize -IncludeAttachmentNames $includeAttachmentNames
      Write-Info "Indexed duplicate signatures: $($targetDuplicateIndex.Count)"
    }
    else {
      Write-Info 'Building in-memory duplicate index by Title...'
      $targetDuplicateIndex = Get-TargetTitleIndex -ListTitle $targetCtx.ListIdentity -PageSize $PageSize
      Write-Info "Indexed target titles: $($targetDuplicateIndex.Count)"
    }
  }

  $i = 0
  foreach ($work in $migrationQueue) {
    $i++
    Write-Info "[Target $i/$($migrationQueue.Count)] Source item ID $($work.SourceId)"

    try {
      [xml]$xmlDoc = [IO.File]::ReadAllText($work.XmlPath)
      $xmlMapForItem = Convert-InfoPathXmlToFieldMap -XmlDoc $xmlDoc
      $sourceKey = "{0}|{1}|{2}" -f $sourceCtx.SiteUrl, $sourceCtx.ListIdentity, $work.SourceId
      $values = New-TargetItemValues -XmlMap $xmlMapForItem -FieldNameMap $targetFieldMap -FieldByInternalName $targetFieldByInternalName -SourceItem ([pscustomobject]@{ Id = $work.SourceId })
      $attachments = @()
      if (-not $SkipAttachments) {
        $attachmentFolder = Join-Path $stagingFolder ("source-{0}" -f $work.SourceId)
        $attachments = @(Get-InfoPathAttachmentsFromXml -XmlDoc $xmlDoc -BasicFileName $FallbackAttachmentName -OutputFolder $attachmentFolder)
      }

      $newItem = $null
      $itemAction = 'Created'
      $existingItemId = $null
      $itemToDeleteId = $null
      $titleValue = if ($values.ContainsKey('Title')) { [string]$values['Title'] } else { '' }
      $sourceDuplicateSignature = $null
      $stateEntry = $null

      if ($Duplicate -ne 'CreateNew' -and $null -ne $targetDuplicateIndex) {
        if ($DuplicateDetection -eq 'SourceItemKey') {
          $sourceKey = "{0}|{1}|{2}" -f $sourceCtx.SiteUrl, $sourceCtx.ListIdentity, $work.SourceId
          if ($targetDuplicateIndex.Items.ContainsKey($sourceKey)) {
            $stateEntry = $targetDuplicateIndex.Items[$sourceKey]
            $verifiedItem = Get-TargetItemById -ListTitle $targetCtx.ListIdentity -ItemId ([int]$stateEntry.TargetId)
            if ($null -ne $verifiedItem) {
              $existingItemId = [int]$stateEntry.TargetId
              if ($stateEntry.Status -ne 'Completed') { $newItem = $verifiedItem }
            }
            else {
              $targetDuplicateIndex.Items.Remove($sourceKey)
            }
          }
        }
        elseif ($DuplicateDetection -eq 'MetadataAndAttachments') {
          $attachmentNames = @($attachments | ForEach-Object { [string]$_.FileName })
          $sourceDuplicateSignature = New-DuplicateSignature -Values $values -AttachmentNames $attachmentNames
          if (-not [string]::IsNullOrWhiteSpace($sourceDuplicateSignature) -and $targetDuplicateIndex.ContainsKey($sourceDuplicateSignature)) {
            $existingItemId = [int]$targetDuplicateIndex[$sourceDuplicateSignature]
          }
        }
        elseif (-not [string]::IsNullOrWhiteSpace($titleValue)) {
          if ($targetDuplicateIndex.ContainsKey($titleValue)) {
            $existingItemId = [int]$targetDuplicateIndex[$titleValue]
          }
        }
      }

      if ($null -ne $existingItemId) {
        if ($Duplicate -eq 'Skip') {
          if ($null -ne $newItem) {
            Write-Warn "  Resuming incomplete migration for source item $($work.SourceId) using target item ID $existingItemId."
          }
          else {
            $script:stats.Skipped++
          if ($DuplicateDetection -eq 'MetadataAndAttachments') {
            Write-Warn "  Skipped duplicate source item $($work.SourceId): metadata+attachment signature matched target item ID $existingItemId (Title '$titleValue')."
          }
          else {
            Write-Warn "  Skipped duplicate source item $($work.SourceId): target item with Title '$titleValue' already exists (ID $existingItemId)."
          }
            continue
          }
        }

        if ($null -eq $newItem -and $Duplicate -eq 'Overwrite') {
          if (-not $PSCmdlet.ShouldProcess("Target list '$($targetCtx.ListIdentity)'", "Replace item ID $existingItemId from source ID $($work.SourceId)")) {
            continue
          }

          $newItem = Add-PnPListItem -List $targetCtx.ListIdentity -Values $values -ErrorAction Stop
          $itemToDeleteId = $existingItemId
          $script:stats.TargetItemsCreated++

          if ($null -ne $targetDuplicateIndex) {
            if ($DuplicateDetection -eq 'SourceItemKey') {
              $targetDuplicateIndex.Items[$sourceKey] = [ordered]@{ TargetId = [int]$newItem.Id; Status = 'Created'; AttachmentNames = @() }
            }
            elseif ($DuplicateDetection -eq 'MetadataAndAttachments') {
              if (-not [string]::IsNullOrWhiteSpace($sourceDuplicateSignature)) {
                $targetDuplicateIndex[$sourceDuplicateSignature] = [int]$newItem.Id
              }
            }
            elseif (-not [string]::IsNullOrWhiteSpace($titleValue)) {
              $targetDuplicateIndex[$titleValue] = [int]$newItem.Id
            }
          }

          $itemAction = 'Replaced'

          Write-Log -Message "  Replaced Title '$titleValue': old item ID $existingItemId -> new item ID $($newItem.Id)" -Level SUCCESS
        }
      }

      if ($null -eq $newItem) {
        if (-not $PSCmdlet.ShouldProcess("Target list '$($targetCtx.ListIdentity)'", "Create item from source ID $($work.SourceId)")) {
          continue
        }

        $newItem = Add-PnPListItem -List $targetCtx.ListIdentity -Values $values -ErrorAction Stop
        $script:stats.TargetItemsCreated++

        # Keep duplicate index current so duplicates within the same run are handled without extra queries.
        if ($Duplicate -ne 'CreateNew' -and $null -ne $targetDuplicateIndex) {
          if ($DuplicateDetection -eq 'SourceItemKey') {
            $targetDuplicateIndex.Items[$sourceKey] = [ordered]@{ TargetId = [int]$newItem.Id; Status = 'Created'; AttachmentNames = @() }
          }
          elseif ($DuplicateDetection -eq 'MetadataAndAttachments') {
            if ([string]::IsNullOrWhiteSpace($sourceDuplicateSignature)) {
              $attachmentNames = @($attachments | ForEach-Object { [string]$_.FileName })
              $sourceDuplicateSignature = New-DuplicateSignature -Values $values -AttachmentNames $attachmentNames
            }

            if (-not [string]::IsNullOrWhiteSpace($sourceDuplicateSignature) -and -not $targetDuplicateIndex.ContainsKey($sourceDuplicateSignature)) {
              $targetDuplicateIndex[$sourceDuplicateSignature] = [int]$newItem.Id
            }
          }
          elseif (-not [string]::IsNullOrWhiteSpace($titleValue) -and -not $targetDuplicateIndex.ContainsKey($titleValue)) {
            $targetDuplicateIndex[$titleValue] = [int]$newItem.Id
          }
        }
      }

      if ($DuplicateDetection -eq 'SourceItemKey' -and ($null -eq $stateEntry -or $stateEntry.Status -ne 'Completed')) {
        $stateEntry = [ordered]@{ TargetId = [int]$newItem.Id; Status = 'Created'; AttachmentNames = @() }
        $migrationState.Items[$sourceKey] = $stateEntry
        Save-MigrationState -Path $StateFilePath -State $migrationState
      }

      if (-not $SkipAttachments) {
        if ($attachments.Count -gt 0) {
          Add-ExtractedAttachmentsToItem -ListTitle $targetCtx.ListIdentity -ItemId $newItem.Id -Attachments $attachments -WorkingFolder $TempFolder -DuplicateMode $Duplicate
          Write-Log -Message "  Attachments uploaded: $($attachments.Count)" -Level SUCCESS
        }
      }

      if ($null -ne $itemToDeleteId) {
        Invoke-PnPWithRetry {
          Remove-PnPListItem -List $targetCtx.ListIdentity -Identity $itemToDeleteId -Force | Out-Null
        }
      }

      if ($DuplicateDetection -eq 'SourceItemKey') {
        $stateEntry = [ordered]@{
          TargetId = [int]$newItem.Id
          Status = 'Completed'
          AttachmentNames = @($attachments | ForEach-Object { [string]$_.FileName })
        }
        $migrationState.Items[$sourceKey] = $stateEntry
        Save-MigrationState -Path $StateFilePath -State $migrationState
      }

      if ($itemAction -eq 'Replaced') {
        Write-Log -Message "  Completed replace for Title '$titleValue' (new item ID $($newItem.Id))" -Level SUCCESS
      }
      elseif ($itemAction -eq 'Created') {
        $logTitle = if ([string]::IsNullOrWhiteSpace($titleValue)) { '<no title>' } else { $titleValue }
        Write-Log -Message "  Created Title '$logTitle' (item ID $($newItem.Id))" -Level SUCCESS
      }
      else {
        Write-Log -Message "  $itemAction target item ID $($newItem.Id)" -Level SUCCESS
      }
    }
    catch {
      $script:stats.Failed++
      Write-Err "  Failed source item $($work.SourceId): $($_.Exception.Message)"
    }
    finally {
      $xmlDoc = $null
      if (Test-Path -LiteralPath $work.XmlPath -PathType Leaf) { Remove-Item -LiteralPath $work.XmlPath -Force -ErrorAction SilentlyContinue }
      $attachmentFolder = Join-Path $stagingFolder ("source-{0}" -f $work.SourceId)
      if (Test-Path -LiteralPath $attachmentFolder -PathType Container) { Remove-Item -LiteralPath $attachmentFolder -Recurse -Force -ErrorAction SilentlyContinue }
    }
  }

  $completionSeparator = ('=' * 92)
  Write-Info ''
  Write-Info $completionSeparator
  Write-Log -Message 'INFO PATH MIGRATION - RUN SUMMARY' -Level SUCCESS
  Write-Info $completionSeparator
  Write-Info ("  Source items read    : {0}" -f $script:stats.SourceItemsRead)
  Write-Info ("  XML resolved         : {0}" -f $script:stats.XmlResolved)
  Write-Info ("  Target items created : {0}" -f $script:stats.TargetItemsCreated)
  Write-Info ("  Attachments uploaded : {0}" -f $script:stats.AttachmentsUploaded)
  Write-Info ("  Skipped              : {0}" -f $script:stats.Skipped)
  Write-Info ("  Failed               : {0}" -f $script:stats.Failed)
  Write-Info $completionSeparator
  Write-Info ''
  $script:runSucceeded = $true
}
catch {
  Write-Err "Script failed: $($_.Exception.Message)"
  throw
}
finally {
  $elapsedSec = [Math]::Round(([datetime]::UtcNow - $script:runStartedUtc).TotalSeconds, 1)
  $finalStatus = if (-not $script:runSucceeded) { 'FAILED' } elseif ($script:stats.Failed -gt 0) { 'PARTIAL' } else { 'SUCCESS' }
  $statusLevel = switch ($finalStatus) {
    'SUCCESS' { 'SUCCESS' }
    'PARTIAL' { 'WARN' }
    default { 'ERROR' }
  }
  Write-Log -Message "Run complete with status: $finalStatus (elapsed ${elapsedSec}s)" -Level $statusLevel
  Write-Log -Message ('=' * 92) -Level INFO
  if ($null -ne $stagingFolder -and (Test-Path -LiteralPath $stagingFolder -PathType Container)) {
    Remove-Item -LiteralPath $stagingFolder -Recurse -Force -ErrorAction SilentlyContinue
  }
  try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch {}
}
#endregion Main Execution
