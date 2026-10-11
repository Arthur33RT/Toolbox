<#
================================================================================
 Create-MC-DiscogsFields.ps1
--------------------------------------------------------------------------------
 One-time setup: creates the custom Discogs library fields in JRiver Media
 Center so that the metadata written by the Discogs Tagger is picked up and 
 displayed by MC.

 Run this ONCE per machine/library. It is safe to re-run: fields that already
 exist are skipped, so nothing is duplicated or overwritten.

 Requirements:
   - JRiver Media Center running with Library Server (Media Network) enabled.
   - If you set a username/password on Media Network, provide them when asked.
   - PowerShell 5.1+ (ships with Windows) or PowerShell 7+.

 Usage (double-click won't work reliably; run from PowerShell):
   Right-click the file > "Run with PowerShell"
   -- or --
   powershell -ExecutionPolicy Bypass -File .\Create-MC-DiscogsFields.ps1

 The script prompts for the Media Network username and password when it runs
 and waits for you to type them in. If your library has no login set, just
 press Enter at both prompts to connect without authentication.

 Optional parameters:
   -Server   MC host        (default: localhost)
   -Port     MCWS port      (default: 52199)
   -DryRun   List what would be created, but create nothing.
================================================================================
#>

[CmdletBinding()]
param(
    [string] $Server   = "localhost",
    [int]    $Port      = 52199,
    [switch] $DryRun
)

# ------------------------------------------------------------------------------
# 1. The fields to create.
#    All are Data Type = String, saved to file tags (MC default for custom
#    string fields). Add or remove names here as your Mp3tag source evolves.
#    Names MUST match the tag names your Mp3tag source writes, exactly.
# ------------------------------------------------------------------------------
$Fields = @(
    "DISCOGS_RELEASE_ID"
    "DISCOGS_LABELS"
    "DISCOGS_COMPANIES"
    "DISCOGS_RECORD_COMPANY"
    "DISCOGS_DISTRIBUTED_BY"
    "DISCOGS_MANUFACTURED_BY"
    "DISCOGS_MADE_BY"
    "DISCOGS_PRESSED_BY"
    "DISCOGS_GLASS_MASTERED_AT"
    "DISCOGS_LACQUER_CUT_AT"
    "DISCOGS_PHONOGRAPHIC_COPYRIGHT"
    "DISCOGS_COPYRIGHT"
    "DISCOGS_PUBLISHED_BY"
    "DISCOGS_PRODUCED_FOR"
    "DISCOGS_LICENSED_FROM"
    "DISCOGS_LICENSED_TO"
    "DISCOGS_MARKETED_BY"
    "DISCOGS_RECORDED_AT"
    "DISCOGS_MIXED_AT"
    "DISCOGS_MASTERED_AT"
    "DISCOGS_DESIGNED_AT"
    "DISCOGS_PRINTED_BY"
    "DISCOGS_IDENTIFIERS"
    "DISCOGS_ALBUM_CREDITS"
    "DISCOGS_CREDIT_PRODUCER"
    "DISCOGS_CREDIT_ENGINEER"
    "DISCOGS_CREDIT_RECORDED_BY"
    "DISCOGS_CREDIT_MIXED_BY"
    "DISCOGS_CREDIT_MASTERED_BY"
    "DISCOGS_CREDIT_LACQUER_CUT_BY"
    "DISCOGS_CREDIT_ARRANGED_BY"
    "DISCOGS_CREDIT_CONDUCTOR"
    "DISCOGS_CREDIT_ORCHESTRA"
    "DISCOGS_CREDIT_CHORUS"
    "DISCOGS_CREDIT_DESIGN"
    "DISCOGS_CREDIT_ARTWORK"
    "DISCOGS_CREDIT_PHOTOGRAPHY"
    "DISCOGS_CREDIT_WRITTEN_BY"
    "DISCOGS_CREDIT_LYRICS_BY"
    "DISCOGS_CREDIT_MUSIC_BY"
    "DISCOGS_CREDIT_COMPOSED_BY"
    "DISCOGS_TRACK_CREDITS"
	"DISCOGS_RELEASE_ID"
	"DISCOGS_RELEASE_URL"
	"DISCOGS_MASTER_ID"
	"DISCOGS_MASTER_URL"
	"ALLMUSIC_RELEASE_URL"
	"ALLMUSIC_RELEASEGROUP_URL"
)

# Data type for every field above. MCWS accepts "string" or "integer" (lowercase).
$FieldType = "string"

# ------------------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------------------
$BaseUrl = "http://${Server}:${Port}/MCWS/v1"

function Get-AuthHeader {
    param([string]$User, [string]$Pass)
    if ([string]::IsNullOrEmpty($User)) { return @{} }
    $pair  = "{0}:{1}" -f $User, $Pass
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    $b64   = [System.Convert]::ToBase64String($bytes)
    return @{ Authorization = "Basic $b64" }
}

function Invoke-MCWS {
    param([string]$Path, [hashtable]$Query, [hashtable]$Headers)
    $qs = ($Query.GetEnumerator() | ForEach-Object {
        "{0}={1}" -f $_.Key, [uri]::EscapeDataString([string]$_.Value)
    }) -join "&"
    $url = "$BaseUrl/$Path"
    if ($qs) { $url = "$url`?$qs" }
    return Invoke-WebRequest -Uri $url -Headers $Headers -UseBasicParsing -TimeoutSec 30
}

# ------------------------------------------------------------------------------
# 2. Prompt for credentials (always) and wait for the user to enter them.
#    Leave both blank to connect without authentication.
# ------------------------------------------------------------------------------
Write-Host ""
Write-Host "  JRiver Media Center - Discogs custom field installer" -ForegroundColor Cyan
Write-Host "  Target: $BaseUrl" -ForegroundColor DarkGray
Write-Host ""
Write-Host "  Enter your Media Network credentials." -ForegroundColor Yellow
Write-Host "  (If your library has no login set, just press Enter at both prompts.)" -ForegroundColor DarkGray

$Username = Read-Host "  Username"
$secure   = Read-Host "  Password" -AsSecureString
$Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure))
$headers  = Get-AuthHeader -User $Username -Pass $Password
Write-Host ""

# ------------------------------------------------------------------------------
# 3. Connectivity check.
# ------------------------------------------------------------------------------
try {
    $probe = Invoke-MCWS -Path "Alive" -Query @{} -Headers $headers
} catch {
    $status = $_.Exception.Response.StatusCode.value__
    if ($status -eq 401) {
        Write-Host "  Authentication failed. Check the username/password and run the script again." -ForegroundColor Red
    } else {
        Write-Host "  Could not reach Media Center at $BaseUrl" -ForegroundColor Red
        Write-Host "  Make sure MC is running and Library Server (Media Network) is enabled," -ForegroundColor Red
        Write-Host "  or pass -Server / -Port if it's on another machine/port." -ForegroundColor Red
    }
    return
}
Write-Host "  Connected to Media Center." -ForegroundColor Green
Write-Host ""

# ------------------------------------------------------------------------------
# 4. Read existing fields so we can skip ones that already exist (idempotent).
# ------------------------------------------------------------------------------
$existing = @{}
try {
    $resp = Invoke-MCWS -Path "Library/Fields" -Query @{} -Headers $headers
    # Format-agnostic: grab every Name="..." attribute in the response.
    [regex]::Matches($resp.Content, 'Name="([^"]+)"') | ForEach-Object {
        $existing[$_.Groups[1].Value.ToLower()] = $true
    }
} catch {
    Write-Host "  (Could not list existing fields; will attempt to create all and" -ForegroundColor DarkYellow
    Write-Host "   let MC skip duplicates.)" -ForegroundColor DarkYellow
}

# ------------------------------------------------------------------------------
# 5. Create fields.
# ------------------------------------------------------------------------------
$created = New-Object System.Collections.Generic.List[string]
$skipped = New-Object System.Collections.Generic.List[string]
$failed  = New-Object System.Collections.Generic.List[string]

foreach ($name in $Fields) {
    if ($existing.ContainsKey($name.ToLower())) {
        $skipped.Add($name)
        Write-Host ("  [skip]   {0}  (already exists)" -f $name) -ForegroundColor DarkGray
        continue
    }

    if ($DryRun) {
        Write-Host ("  [dryrun] {0}  (would create, type={1})" -f $name, $FieldType) -ForegroundColor Yellow
        continue
    }

    try {
        # --- The one MCWS call. Parameter names per MC's /MCWS/v1/doc. ---
        $r = Invoke-MCWS -Path "Library/CreateField" -Headers $headers -Query @{
            Name = $name
            Type = $FieldType
        }
        # MCWS wraps replies in <Response Status="OK"/"Failure">. Treat an
        # explicit Failure as a failure; otherwise a 200 with no Failure is OK.
        $ok = ($r.Content -match 'Status="OK"') -or `
              ($r.StatusCode -eq 200 -and $r.Content -notmatch 'Status="Failure"')
        if ($ok) {
            $created.Add($name)
            Write-Host ("  [ok]     {0}" -f $name) -ForegroundColor Green
        } else {
            $failed.Add($name)
            Write-Host ("  [fail]   {0}  -> {1}" -f $name, ($r.Content -replace '\s+',' ').Trim()) -ForegroundColor Red
        }
    } catch {
        $msg = $_.Exception.Message
        $failed.Add($name)
        Write-Host ("  [fail]   {0}  -> {1}" -f $name, $msg) -ForegroundColor Red
    }
}

# ------------------------------------------------------------------------------
# 6. Summary.
# ------------------------------------------------------------------------------
Write-Host ""
Write-Host "  ------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host ("  Created: {0}   Skipped (already existed): {1}   Failed: {2}" -f `
    $created.Count, $skipped.Count, $failed.Count) -ForegroundColor Cyan
if ($failed.Count -gt 0) {
    Write-Host ""
    Write-Host "  Some fields failed. Copy the [fail] lines above and send them over." -ForegroundColor Yellow
}
if (-not $DryRun -and $created.Count -gt 0) {
    Write-Host ""
    Write-Host "  Done. In MC you can now tag files with the Mp3tag Discogs source" -ForegroundColor Green
    Write-Host "  and these fields will populate. If MC was already open, a library" -ForegroundColor Green
    Write-Host "  refresh (or restart) makes the new fields available in views." -ForegroundColor Green
}
Write-Host ""
if ($Host.Name -eq "ConsoleHost" -and -not $DryRun) {
    Read-Host "  Press Enter to close"
}
