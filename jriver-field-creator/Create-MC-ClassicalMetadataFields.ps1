<#
Create JRiver fields matching Classical Metadata Picard plugin 0.2.7.
PowerShell 5.1+. JRiver running, Media Network/Library Server enabled.
Default creates 65 core fields, including completeness status, IDs and JSON.
Use -IncludeDiagnostics to also create five operational/debug fields.
Existing fields are NEVER updated or deleted. No media files are changed.
Back up your MC library before applying changes.

MCWS CreateField supports only string/integer. All fields start as String,
per-file, without a calculated expression. Dates MUST remain String.
List/relational/edit/display settings cannot be set by this documented API.
Test multi-value imports before manually changing chosen fields to List.
Keep ordered paths and JSON separate from unordered browsing lists.

Examples:
  .\Create-MC-ClassicalMetadataFields.ps1 -ListOnly
  .\Create-MC-ClassicalMetadataFields.ps1 -DryRun -NoAuthentication
  .\Create-MC-ClassicalMetadataFields.ps1 -NoAuthentication
  .\Create-MC-ClassicalMetadataFields.ps1 -Credential (Get-Credential)
  .\Create-MC-ClassicalMetadataFields.ps1 -IncludeDiagnostics -DryRun

Without -NoAuthentication/-Credential, prompts for username; Enter means no login.
Use HTTPS for credentials on a remote host; HTTP Basic is not encrypted.
Changes reach the currently active MC library. No field rollback/deletion provided.
#>
[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
param(
    [ValidateNotNullOrEmpty()][string]$Server = 'localhost',
    [ValidateRange(1,65535)][int]$Port = 52199,
    [ValidateSet('http','https')][string]$Scheme = 'http',
    [System.Management.Automation.PSCredential]$Credential,
    [switch]$NoAuthentication,
    [switch]$IncludeDiagnostics,
    [switch]$ExcludeDiagnostics,
    [switch]$ListOnly,
    [switch]$DryRun
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$FieldNames = @(
    'CLASSICAL_PLUGIN_STATUS'
    'CLASSICAL_RETRIEVAL_STATUS'
    'CLASSICAL_RETRIEVAL_ERROR_COUNT'
    'CLASSICAL_RETRIEVAL_ELAPSED_SECONDS'
    'CLASSICAL_RETRIEVAL_NETWORK_REQUESTS'
    'CLASSICAL_RETRIEVAL_CACHE_HITS'
    'CLASSICAL_RETRIEVAL_COALESCED_LOOKUPS'
    'CLASSICAL_HIERARCHY_STATUS'
    'CLASSICAL_HIERARCHY_PATH_COUNT'
    'CLASSICAL_WORK_PATHS'
    'CLASSICAL_SOURCE_EVIDENCE'
    'CLASSICAL_SCHEMA_VERSION'
    'CLASSICAL_STATUS'
    'CLASSICAL_RELEASE_ID'
    'CLASSICAL_RECORDING_ID'
    'CLASSICAL_WORK'
    'CLASSICAL_WORK_ID'
    'CLASSICAL_WORK_TYPE'
    'CLASSICAL_CANDIDATE_WORK'
    'CLASSICAL_CANDIDATE_WORK_ID'
    'CLASSICAL_BROADER_WORK_ASSIGNMENT'
    'CLASSICAL_BROADER_WORK_ASSIGNMENT_ID'
    'CLASSICAL_WORK_PATH'
    'CLASSICAL_WORK_PATH_SUMMARY'
    'CLASSICAL_WORK_PATH_ID'
    'CLASSICAL_WORK_DEPTH'
    'CLASSICAL_PARENT_WORK'
    'CLASSICAL_PARENT_WORK_ID'
    'CLASSICAL_PART_ORDER'
    'CLASSICAL_PARTIAL'
    'CLASSICAL_SEGMENT_INDEX'
    'CLASSICAL_SEGMENT_COUNT'
    'CLASSICAL_COMPOSER'
    'CLASSICAL_COMPOSER_ID'
    'CLASSICAL_LIBRETTIST'
    'CLASSICAL_LIBRETTIST_ID'
    'CLASSICAL_LYRICIST'
    'CLASSICAL_LYRICIST_ID'
    'CLASSICAL_WRITER'
    'CLASSICAL_WRITER_ID'
    'CLASSICAL_TRANSLATOR'
    'CLASSICAL_TRANSLATOR_ID'
    'CLASSICAL_ARRANGER'
    'CLASSICAL_ARRANGER_ID'
    'CLASSICAL_ORCHESTRATOR'
    'CLASSICAL_ORCHESTRATOR_ID'
    'CLASSICAL_CATALOG'
    'CLASSICAL_SOURCE_WORK'
    'CLASSICAL_SOURCE_WORK_ID'
    'CLASSICAL_SOURCE_KIND'
    'CLASSICAL_SOURCE_CONFIDENCE'
    'CLASSICAL_CONDUCTOR'
    'CLASSICAL_CONDUCTOR_ID'
    'CLASSICAL_ENSEMBLE'
    'CLASSICAL_ENSEMBLE_ID'
    'CLASSICAL_PERFORMER'
    'CLASSICAL_PERFORMER_ID'
    'CLASSICAL_PERFORMER_ROLE'
    'CLASSICAL_PERFORMER_INSTRUMENT'
    'CLASSICAL_PRODUCER'
    'CLASSICAL_ENGINEER'
    'CLASSICAL_BALANCE'
    'CLASSICAL_EDITOR'
    'CLASSICAL_MIX'
    'CLASSICAL_MASTERING'
    'CLASSICAL_RECORDING_PLACE'
    'CLASSICAL_RECORDING_PLACE_ID'
    'CLASSICAL_RECORDING_AREA'
    'CLASSICAL_RECORDING_DATE_BEGIN'
    'CLASSICAL_RECORDING_DATE_END'
)
$OptionalDiagnosticNames = @(
    'CLASSICAL_PLUGIN_STATUS'
    'CLASSICAL_RETRIEVAL_ELAPSED_SECONDS'
    'CLASSICAL_RETRIEVAL_NETWORK_REQUESTS'
    'CLASSICAL_RETRIEVAL_CACHE_HITS'
    'CLASSICAL_RETRIEVAL_COALESCED_LOOKUPS'
)
if ($IncludeDiagnostics -and $ExcludeDiagnostics) {
    throw 'Choose either -IncludeDiagnostics or the legacy -ExcludeDiagnostics switch, not both.'
}
if (-not $IncludeDiagnostics) {
    $FieldNames = @($FieldNames | Where-Object { $_ -notin $OptionalDiagnosticNames })
}
if ($ListOnly) {
    $FieldNames | ForEach-Object {
        [pscustomobject]@{ Name = $_; DataType = 'String'; Relational = 'Per file' }
    }
    return
}
if ($Credential -and $NoAuthentication) {
    throw 'Choose either -Credential or -NoAuthentication, not both.'
}
$baseUri = [uri]("${Scheme}://${Server}:${Port}/MCWS/v1/")
$headers = @{}
if (-not $NoAuthentication -and -not $Credential) {
    $userName = Read-Host 'Media Network username (Enter for no authentication)'
    if ($userName) {
        $passwordSecure = Read-Host 'Media Network password' -AsSecureString
        $Credential = New-Object System.Management.Automation.PSCredential($userName, $passwordSecure)
    }
}
if ($Credential) {
    if ($Scheme -eq 'http' -and $Server -notin @('localhost','127.0.0.1','::1')) {
        throw 'Remote credentials require -Scheme https. HTTP Basic exposes credentials on the network.'
    }
    $credentialPointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password)
    try {
        $passwordPlain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($credentialPointer)
        $authPair = '{0}:{1}' -f $Credential.UserName, $passwordPlain
        $headers.Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($authPair))
    } finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($credentialPointer)
        $passwordPlain = $null
        $authPair = $null
    }
}
function Invoke-ClassicalMCWS {
    param([string]$Endpoint, [hashtable]$Query = @{})
    $queryString = (@($Query.Keys | Sort-Object | ForEach-Object {
        [uri]::EscapeDataString($_) + '=' + [uri]::EscapeDataString([string]$Query[$_])
    })) -join '&'
    $requestUri = [uri]::new($baseUri, $Endpoint).AbsoluteUri
    if ($queryString) { $requestUri += '?' + $queryString }
    try {
        $response = Invoke-WebRequest -Uri $requestUri -Headers $headers -UseBasicParsing -TimeoutSec 30
    } catch {
        # Do not include response text, URLs with credentials, or authentication headers.
        throw "MCWS request failed at endpoint $Endpoint. Check server, port and credentials."
    }
    try { $xml = [xml]$response.Content } catch { throw "Invalid MCWS XML at $Endpoint." }
    if ($xml.DocumentElement.Name -ne 'Response' -or
        $xml.DocumentElement.GetAttribute('Status') -ne 'OK') {
        throw "MCWS did not report success at endpoint $Endpoint."
    }
    return ,$xml
}
function Get-ClassicalMCFields {
    $responseXml = Invoke-ClassicalMCWS -Endpoint 'Library/Fields'
    $fieldsNode = $responseXml.SelectSingleNode('/Response/Fields')
    if ($null -eq $fieldsNode) { throw 'Cannot safely read field inventory; no fields will be created.' }
    $inventory = @{} # PowerShell string keys are case-insensitive.
    foreach ($node in $fieldsNode.SelectNodes('Field')) {
        $name = $node.GetAttribute('Name')
        if ($name) { $inventory[$name] = $node.GetAttribute('DataType') }
    }
    if ($inventory.Count -eq 0) { throw 'Empty field inventory; no fields will be created.' }
    return $inventory
}
try {
    Write-Host "Target: $baseUri"
    $null = Invoke-ClassicalMCWS -Endpoint 'Alive'
    $existing = Get-ClassicalMCFields
    $created = 0; $skipped = 0; $planned = 0
    foreach ($name in $FieldNames) {
        if ($existing.ContainsKey($name)) {
            $skipped++
            Write-Host "[skip] $name (existing type: $($existing[$name]); unchanged)"
            continue
        }
        if ($DryRun) {
            $planned++
            Write-Host "[dryrun] $name (String)"
            continue
        }
        if (-not $PSCmdlet.ShouldProcess($name, 'Create per-file String field in JRiver')) {
            $planned++
            continue
        }
        $null = Invoke-ClassicalMCWS -Endpoint 'Library/CreateField' -Query @{ Name=$name; Type='string' }
        # Verify creation rather than treating HTTP 200 as proof.
        $verified = Get-ClassicalMCFields
        if (-not $verified.ContainsKey($name)) { throw "Creation not verified for $name; stopping." }
        if ($verified[$name] -ne 'String') { throw "Unexpected type for $name; stopping." }
        $existing = $verified
        $created++
        Write-Host "[created] $name (String)"
    }
    Write-Host "Created: $created; skipped: $skipped; planned/not applied: $planned"
    Write-Host 'No media files were changed. Existing field settings were preserved.'
    Write-Host 'Check Save in file tags and field types in MC, then refresh tags FROM files on test copies.'
    Write-Host 'Use Picard as the tag writer; MC file-tag write-back is outside this supported workflow.'
} finally {
    $headers.Clear()
}
