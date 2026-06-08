#requires -Version 5.1
<#
  DriveApi.ps1 - Client mínim de Google Drive (REST API) per al PC, SENSE
  Google Drive d'escriptori. S'autentica amb un "refresh token" (OAuth
  d'aplicació d'escriptori), de manera que tot són crides HTTPS normals.

  L'autorització única es fa amb Authorize-Drive.ps1, que desa les credencials
  a:  %LOCALAPPDATA%\InformesCornella\drive-credencials.json
  (fora del repositori; conté secrets i NO s'ha de pujar mai a GitHub).

  Aquest fitxer NOMÉS defineix funcions; es pot carregar (dot-source) sense
  efectes secundaris. Les rutes/IDs de carpetes les posa config.ps1
  ($DriveEntradaId, $DriveProcessatsId, $DriveDadesId).
#>

# Google rebutja TLS antic; forcem TLS 1.2 (PowerShell 5.1 sovint per defecte
# encara fa servir 1.0). No passa res si ja hi és.
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }

$Script:DriveScope    = 'https://www.googleapis.com/auth/drive'
$Script:_driveToken   = $null
$Script:_driveTokenExp = [datetime]::MinValue

function Get-DriveCredPath {
    $dir = Join-Path $env:LOCALAPPDATA 'InformesCornella'
    return (Join-Path $dir 'drive-credencials.json')
}

function Test-DriveApiConfigured {
    $p = Get-DriveCredPath
    if (-not (Test-Path -LiteralPath $p)) { return $false }
    try {
        $c = Get-Content -LiteralPath $p -Raw -Encoding UTF8 | ConvertFrom-Json
        return [bool]($c.client_id -and $c.client_secret -and $c.refresh_token)
    } catch { return $false }
}

function Get-DriveCredentials {
    $p = Get-DriveCredPath
    if (-not (Test-Path -LiteralPath $p)) { throw "No hi ha credencials de Drive. Executa Authorize-Drive.ps1 primer." }
    return (Get-Content -LiteralPath $p -Raw -Encoding UTF8 | ConvertFrom-Json)
}

function Save-DriveCredentials($cred) {
    $p = Get-DriveCredPath
    $dir = Split-Path -Parent $p
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    ($cred | ConvertTo-Json) | Set-Content -LiteralPath $p -Encoding UTF8
}

# Torna un access token vàlid, renovant-lo amb el refresh token si cal.
function Get-DriveAccessToken {
    if ($Script:_driveToken -and (Get-Date) -lt $Script:_driveTokenExp) {
        return $Script:_driveToken
    }
    $c = Get-DriveCredentials
    $body = @{
        client_id     = $c.client_id
        client_secret = $c.client_secret
        refresh_token = $c.refresh_token
        grant_type    = 'refresh_token'
    }
    $r = Invoke-RestMethod -Method Post -Uri 'https://oauth2.googleapis.com/token' -Body $body
    if (-not $r.access_token) { throw "No s'ha pogut obtenir l'access token de Drive." }
    $Script:_driveToken    = $r.access_token
    $expiresIn = if ($r.expires_in) { [int]$r.expires_in } else { 3600 }
    $Script:_driveTokenExp = (Get-Date).AddSeconds($expiresIn - 60)  # marge de 60s
    return $Script:_driveToken
}

function _DriveAuthHeader {
    return @{ Authorization = ("Bearer " + (Get-DriveAccessToken)) }
}

# Cerca un fitxer per nom dins d'una carpeta. Torna l'id o $null.
function Find-DriveFileId($name, $parentId) {
    $safe = $name -replace "'", "\'"
    $q = "name = '$safe' and '$parentId' in parents and trashed = false"
    $uri = "https://www.googleapis.com/drive/v3/files?fields=files(id,name)&pageSize=1&q=" + [uri]::EscapeDataString($q)
    $r = Invoke-RestMethod -Uri $uri -Headers (_DriveAuthHeader)
    if ($r.files -and $r.files.Count -gt 0) { return $r.files[0].id }
    return $null
}

# Crea o actualitza un fitxer JSON dins d'una carpeta amb el contingut donat.
function Save-DriveJson($name, $parentId, $jsonString) {
    if (-not $parentId) { throw "Falta l'ID de la carpeta de Drive (revisa config.ps1)." }
    $existing = Find-DriveFileId $name $parentId
    $headers = _DriveAuthHeader
    if ($existing) {
        # Actualitza el contingut (mèdia) del fitxer existent.
        $uri = "https://www.googleapis.com/upload/drive/v3/files/" + [uri]::EscapeDataString([string]$existing) + "?uploadType=media"
        Invoke-RestMethod -Method Patch -Uri $uri -Headers $headers -ContentType 'application/json; charset=UTF-8' -Body $jsonString | Out-Null
        return $existing
    } else {
        # Crea un fitxer nou (multipart: metadades + contingut).
        $boundary = [guid]::NewGuid().ToString()
        $meta = @{ name = $name; parents = @($parentId) } | ConvertTo-Json -Compress
        $body = "--$boundary`r`nContent-Type: application/json; charset=UTF-8`r`n`r`n$meta`r`n--$boundary`r`nContent-Type: application/json; charset=UTF-8`r`n`r`n$jsonString`r`n--$boundary--"
        $uri = "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id"
        $r = Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -ContentType "multipart/related; boundary=$boundary" -Body $body
        return $r.id
    }
}

# Llista els fitxers (no carpetes) d'una carpeta. Torna objectes {Id, Name}.
function Get-DriveChildren($parentId, [string]$extension = $null) {
    if (-not $parentId) { throw "Falta l'ID de la carpeta de Drive (revisa config.ps1)." }
    $q = "'$parentId' in parents and trashed = false and mimeType != 'application/vnd.google-apps.folder'"
    $uri = "https://www.googleapis.com/drive/v3/files?orderBy=createdTime&pageSize=100&fields=files(id,name)&q=" + [uri]::EscapeDataString($q)
    $r = Invoke-RestMethod -Uri $uri -Headers (_DriveAuthHeader)
    $out = @()
    foreach ($f in $r.files) {
        if ($extension -and -not ($f.name.ToLower().EndsWith($extension.ToLower()))) { continue }
        $out += [pscustomobject]@{ Id = $f.id; Name = $f.name }
    }
    return $out
}

# Descarrega el contingut (text) d'un fitxer pel seu id. Construim la URL per
# concatenacio (no per interpolacio) perque el separador '?' no s'enganxi mai a
# l'id.
function Get-DriveFileText($id) {
    $encId = [uri]::EscapeDataString([string]$id)
    $uri = "https://www.googleapis.com/drive/v3/files/" + $encId + "?alt=media"
    try {
        $resp = Invoke-WebRequest -Uri $uri -Headers (_DriveAuthHeader) -UseBasicParsing
        return [string]$resp.Content
    } catch {
        throw "Baixant (URL: $uri): $(Get-DriveHttpErrorBody $_)"
    }
}

# Mou un fitxer d'una carpeta a una altra (canvia el pare).
function Move-DriveFile($id, $newParentId, $oldParentId) {
    $encId = [uri]::EscapeDataString([string]$id)
    $uri = "https://www.googleapis.com/drive/v3/files/" + $encId + "?fields=id,parents&addParents=" + $newParentId
    if ($oldParentId) { $uri += "&removeParents=" + $oldParentId }
    try {
        Invoke-RestMethod -Method Patch -Uri $uri -Headers (_DriveAuthHeader) | Out-Null
    } catch {
        throw "Movent (URL: $uri): $(Get-DriveHttpErrorBody $_)"
    }
}

# Extreu el cos de la resposta d'un error HTTP (Drive hi posa el motiu real,
# p. ex. 'File not found' o 'insufficientPermissions'). Si no es pot, torna el
# missatge generic.
function Get-DriveHttpErrorBody($err) {
    $msg = $err.Exception.Message
    try {
        $resp = $err.Exception.Response
        if ($resp) {
            $reader = New-Object System.IO.StreamReader($resp.GetResponseStream())
            $body = $reader.ReadToEnd()
            if ($body) { $msg = "$msg | $body" }
        }
    } catch { }
    return $msg
}
