#requires -Version 5.1
<#
.SYNOPSIS
  Autorització única perquè el PC pugui accedir al teu Google Drive per API
  (sense Google Drive d'escriptori).

.DESCRIPTION
  Demana el Client ID i el Client Secret d'un client OAuth de tipus
  "Aplicació d'escriptori" (creat a Google Cloud Console), obre el navegador
  perquè autoritzis amb el teu compte, captura el codi via un servidor local
  temporal i el bescanvia per un "refresh token". Ho desa tot a:

    %LOCALAPPDATA%\InformesCornella\drive-credencials.json

  Aquest fitxer queda al teu PC (fora del repositori). A partir d'aquí, el
  vigilant i l'exportació d'activitats accedeixen a Drive sols.

  Doble clic NO; executa des de PowerShell:
    powershell -ExecutionPolicy Bypass -File suport\Authorize-Drive.ps1
#>

param(
    # Opcionals: si algun dia regeneres el client/secret, els pots passar aqui
    # sense tocar el codi. Si no, s'usen els valors configurats a sota.
    [string]$ClientId,
    [string]$ClientSecret
)

$ErrorActionPreference = 'Stop'
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $ScriptRoot 'DriveApi.ps1')

try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }

Write-Host "=== Autoritzar el PC a Google Drive ===" -ForegroundColor Cyan

# El Client ID NO es secret: el deixem posat perque no l'hagis d'escriure.
# El Secret (GOCSPX-...) SI que ho es, i GitHub no deixa pujar-lo al repo, aixi
# que el demanem una sola vegada (un sol "enganxa"). Despres ja no cal mai mes.
$DEF_CLIENT_ID = '464628466232-fs34gc9vkhrssjnd73t9rplb5bn8lib1.apps.googleusercontent.com'

$clientId     = if ($ClientId) { $ClientId } else { $DEF_CLIENT_ID }
$clientSecret = $ClientSecret
if ([string]::IsNullOrWhiteSpace($clientSecret)) {
    Write-Host "El Client ID ja esta posat. Nomes falta el Secret del client d'escriptori." -ForegroundColor Yellow
    Write-Host "(El trobes a Cloud Console > Credencials > el teu client d'escriptori; comenca per GOCSPX-)"
    $clientSecret = Read-Host "Enganxa el Client Secret i prem Enter"
}
if ([string]::IsNullOrWhiteSpace($clientId) -or [string]::IsNullOrWhiteSpace($clientSecret)) {
    Write-Host "Falta el Secret. Avorto." -ForegroundColor Red
    exit 1
}
Write-Host ""

# 1) Servidor local temporal (loopback) per rebre el codi. TcpListener a
#    127.0.0.1 no necessita permisos d'administrador.
$port = $null
$listener = $null
foreach ($p in 5599, 5601, 5603, 8723, 8910) {
    try {
        $l = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $p)
        $l.Start()
        $listener = $l; $port = $p; break
    } catch { }
}
if (-not $listener) { Write-Host "No s'ha pogut obrir cap port local. Avorto." -ForegroundColor Red; exit 1 }

$redirect = "http://127.0.0.1:$port"
$scope    = 'https://www.googleapis.com/auth/drive'
$authUrl  = "https://accounts.google.com/o/oauth2/v2/auth?client_id=" + [uri]::EscapeDataString($clientId) +
            "&redirect_uri=" + [uri]::EscapeDataString($redirect) +
            "&response_type=code&access_type=offline&prompt=consent&scope=" + [uri]::EscapeDataString($scope)

Write-Host ""
Write-Host "S'obrira el navegador per autoritzar. Si no s'obre sol, copia aquesta URL:" -ForegroundColor Yellow
Write-Host $authUrl
Start-Process $authUrl

# 2) Esperem la redireccio amb el codi.
Write-Host "Esperant l'autoritzacio al navegador..." -ForegroundColor Yellow
$client = $listener.AcceptTcpClient()
$stream = $client.GetStream()
$reader = New-Object System.IO.StreamReader($stream)
$requestLine = $reader.ReadLine()   # "GET /?code=XYZ&... HTTP/1.1"

$code = $null
if ($requestLine -match 'code=([^&\s]+)') { $code = [uri]::UnescapeDataString($matches[1]) }

# Resposta al navegador
$writer = New-Object System.IO.StreamWriter($stream)
$writer.WriteLine("HTTP/1.1 200 OK")
$writer.WriteLine("Content-Type: text/html; charset=utf-8")
$writer.WriteLine("Connection: close")
$writer.WriteLine()
$ok = if ($code) { "Autoritzat correctament. Ja pots tancar aquesta finestra i tornar a PowerShell." }
      else       { "No s'ha rebut cap codi. Torna a provar." }
$writer.WriteLine("<html><body style='font-family:sans-serif'><h2>Informes Cornella</h2><p>$ok</p></body></html>")
$writer.Flush()
$client.Close()
$listener.Stop()

if (-not $code) { Write-Host "No s'ha rebut el codi d'autoritzacio. Avorto." -ForegroundColor Red; exit 1 }

# 3) Bescanviem el codi per tokens (inclou el refresh_token).
Write-Host "Obtenint el refresh token..." -ForegroundColor Yellow
$body = @{
    code          = $code
    client_id     = $clientId
    client_secret = $clientSecret
    redirect_uri  = $redirect
    grant_type    = 'authorization_code'
}
$tok = Invoke-RestMethod -Method Post -Uri 'https://oauth2.googleapis.com/token' -Body $body
if (-not $tok.refresh_token) {
    Write-Host "Google no ha tornat cap refresh_token. Torna a executar (cal 'prompt=consent')." -ForegroundColor Red
    exit 1
}

Save-DriveCredentials ([pscustomobject]@{
    client_id     = $clientId
    client_secret = $clientSecret
    refresh_token = $tok.refresh_token
})

Write-Host ""
Write-Host "OK. Credencials desades a: $(Get-DriveCredPath)" -ForegroundColor Green
Write-Host "Ara el PC ja pot accedir a Drive per API (vigilant + exportacio)." -ForegroundColor Green
