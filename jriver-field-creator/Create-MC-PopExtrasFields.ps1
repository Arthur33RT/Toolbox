<#
Create the 17 browsable MusicBrainz Pop Extras fields in JRiver Media Center.

PowerShell 5.1+. JRiver must be running with Media Network/Library Server enabled.
Existing fields are never updated or deleted, and no media files are changed.

MCWS CreateField creates these directly as List fields. Existing definitions are
preserved unchanged. MUSICBRAINZ_RECORDING_DATE must not use JRiver's Date type.

Examples:
  .\Create-MC-PopExtrasFields.ps1 -ListOnly
  .\Create-MC-PopExtrasFields.ps1 -DryRun -NoAuthentication
  .\Create-MC-PopExtrasFields.ps1 -Credential (Get-Credential)

Without -NoAuthentication or -Credential, the script prompts for a username;
press Enter for no authentication. Use HTTPS for credentials on a remote host.
This installation requires authentication for operations that write to MCWS;
use -Credential when creating fields. -NoAuthentication remains available only
for JRiver installations explicitly configured to allow anonymous writes.
Back up the MC library before creating or changing fields.
#>
[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
param(
    [ValidateNotNullOrEmpty()][string]$Server = 'localhost',
    [ValidateRange(1,65535)][int]$Port = 52199,
    [ValidateSet('http','https')][string]$Scheme = 'http',
    [System.Management.Automation.PSCredential]$Credential,
    [switch]$NoAuthentication,
    [switch]$ListOnly,
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$FieldNames = @(
    'MUSICBRAINZ_PROGRAMMING'
    'MUSICBRAINZ_INSTRUMENT_ARRANGER'
    'MUSICBRAINZ_VOCAL_ARRANGER'
    'MUSICBRAINZ_QUALIFIED_PRODUCER'
    'MUSICBRAINZ_QUALIFIED_ENGINEER'
    'MUSICBRAINZ_PERFORMANCE_CREDITS'
    'MUSICBRAINZ_PRODUCTION_CREDITS'
    'MUSICBRAINZ_RECORDED_AT'
    'MUSICBRAINZ_MIXED_AT'
    'MUSICBRAINZ_MASTERED_AT'
    'MUSICBRAINZ_RECORDING_DATE'
    'MUSICBRAINZ_WORK_CREDITS'
    'MUSICBRAINZ_PUBLISHERS'
    'MUSICBRAINZ_ISWC'
    'MUSICBRAINZ_WORK_LANGUAGE'
    'MUSICBRAINZ_SOURCE_MATERIAL'
    'MUSICBRAINZ_WORK_GENEALOGY'
)

if ($ListOnly) {
    $FieldNames | ForEach-Object {
        [pscustomobject]@{
            Name = $_
            InstallerDataType = 'List (semicolon delimited)'
            Relational = 'Not relational'
        }
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

function Invoke-PopExtrasMCWS {
    param([string]$Endpoint, [hashtable]$Query = @{})
    $queryString = (@($Query.Keys | Sort-Object | ForEach-Object {
        [uri]::EscapeDataString($_) + '=' + [uri]::EscapeDataString([string]$Query[$_])
    })) -join '&'
    $requestUri = [uri]::new($baseUri, $Endpoint).AbsoluteUri
    if ($queryString) { $requestUri += '?' + $queryString }
    try {
        $response = Invoke-WebRequest -Uri $requestUri -Headers $headers -UseBasicParsing -TimeoutSec 30
    } catch {
        # Never echo response bodies, authorization headers, or credential-bearing URLs.
        throw "MCWS request failed at endpoint $Endpoint. Check server, port and credentials."
    }
    try { $xml = [xml]$response.Content } catch { throw "Invalid MCWS XML at $Endpoint." }
    if ($xml.DocumentElement.Name -ne 'Response' -or
        $xml.DocumentElement.GetAttribute('Status') -ne 'OK') {
        throw "MCWS did not report success at endpoint $Endpoint."
    }
    return ,$xml
}

function Get-PopExtrasMCFields {
    $responseXml = Invoke-PopExtrasMCWS -Endpoint 'Library/Fields'
    $fieldsNode = $responseXml.SelectSingleNode('/Response/Fields')
    if ($null -eq $fieldsNode) {
        throw 'Cannot safely read the field inventory; no fields will be created.'
    }
    $inventory = @{} # PowerShell string keys are case-insensitive.
    foreach ($node in $fieldsNode.SelectNodes('Field')) {
        $name = $node.GetAttribute('Name')
        if ($name) { $inventory[$name] = $node.GetAttribute('DataType') }
    }
    if ($inventory.Count -eq 0) {
        throw 'Empty field inventory; no fields will be created.'
    }
    return $inventory
}

try {
    Write-Host "Target: $baseUri"
    $null = Invoke-PopExtrasMCWS -Endpoint 'Alive'
    $existing = Get-PopExtrasMCFields
    $created = 0
    $skipped = 0
    $planned = 0

    foreach ($name in $FieldNames) {
        if ($existing.ContainsKey($name)) {
            $skipped++
            Write-Host "[skip] $name (existing type: $($existing[$name]); unchanged)"
            continue
        }
        if ($DryRun) {
            $planned++
            Write-Host "[dryrun] $name (List)"
            continue
        }
        if (-not $PSCmdlet.ShouldProcess($name, 'Create per-file List field in JRiver')) {
            $planned++
            continue
        }
        $null = Invoke-PopExtrasMCWS -Endpoint 'Library/CreateField' -Query @{ Name=$name; Type='list' }
        $verified = Get-PopExtrasMCFields
        if (-not $verified.ContainsKey($name)) {
            throw "Creation not verified for $name; stopping."
        }
        if ($verified[$name] -notmatch '^List(?: \(semicolon delimited\))?$') {
            throw "Unexpected type for $name ($($verified[$name])); expected List; stopping."
        }
        $existing = $verified
        $created++
        Write-Host "[created] $name (List)"
    }

    Write-Host "Created: $created; skipped: $skipped; planned/not applied: $planned"
    Write-Host 'No media files were changed and existing field settings were preserved.'
    Write-Host 'New fields were created directly as List fields and verified.'
    Write-Host 'Existing definitions were left unchanged. Keep MUSICBRAINZ_RECORDING_DATE out of Date type.'
    Write-Host 'Afterward, refresh tags FROM files on the validated test albums.'
} finally {
    $headers.Clear()
}
