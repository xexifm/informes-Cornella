#requires -Version 5.1
<#
  EmailQuota.ps1 - Comptador d'enviaments d'EmailJS d'AQUEST PC.

  El pla gratuit d'EmailJS son 200 correus al mes. Aquest comptador els vigila
  i deixa un LIMIT DE SEGURETAT (150 per defecte) perque mai es pugui esgotar
  la quota sencera: els 50 restants queden de reserva.

  Compta TOTS els enviaments del PC (l'eina "Enviar correu" i els recordatoris),
  no nomes els recordatoris: es l'unic que protegeix de debo. Send-EmailJs
  (EnviarCorreu.ps1) hi suma 1 despres de CADA enviament correcte.

  DUES LIMITACIONS QUE CAL SABER (i que es diuen a la interficie):
   1. El mes d'EmailJS es reinicia el dia de FACTURACIO del compte, no l'1. La
      reserva de 50 es justament el coixi d'aquest desfasament.
   2. Els correus enviats DES DEL MOBIL no es poden comptar des del PC.

  Viu a %LOCALAPPDATA% (com controls-cp-email.json): sobreviu a Actualitzar.bat
  i a tornar a clonar, i no va mai al repositori.
#>

# Limit de seguretat per defecte (dels 200 del pla gratuit).
$Script:QuotaLimitPerDefecte = 150
# Quota real del compte, nomes per ensenyar-la a la interficie.
$Script:QuotaLimitCompte = 200

function _QuotaPath {
    $base = [string]$env:LOCALAPPDATA
    if ([string]::IsNullOrWhiteSpace($base)) { $base = [System.IO.Path]::GetTempPath() }
    return [string](Join-Path $base (Join-Path 'InformesCornella' 'emailjs-quota.json'))
}

# Mes actual en format 'yyyy-MM' (clau del comptador). PURA.
function _QuotaMesActual([datetime]$avui) {
    if ($null -eq $avui -or $avui -eq [datetime]::MinValue) { $avui = Get-Date }
    return $avui.ToString('yyyy-MM')
}

# Normalitza un registre de quota contra el mes demanat: si el mes ha canviat,
# el comptador torna a 0. PURA (no toca el disc).
function _QuotaNormalitza($reg, [string]$mes) {
    $limit = $Script:QuotaLimitPerDefecte
    $enviats = 0
    $mesReg = ''
    if ($null -ne $reg) {
        try { if ($reg.PSObject.Properties['limit'])   { $limit   = [int]$reg.limit } }   catch { }
        try { if ($reg.PSObject.Properties['enviats']) { $enviats = [int]$reg.enviats } } catch { }
        try { if ($reg.PSObject.Properties['mes'])     { $mesReg  = [string]$reg.mes } }  catch { }
    }
    if ($limit -le 0) { $limit = $Script:QuotaLimitPerDefecte }
    # Mes nou -> es comenca de zero.
    if ($mesReg -ne $mes) { $enviats = 0 }
    if ($enviats -lt 0) { $enviats = 0 }
    return @{ mes = $mes; enviats = $enviats; limit = $limit }
}

# Correus que encara es poden enviar aquest mes. Mai negatiu. PURA.
function _QuotaRestant($reg) {
    if ($null -eq $reg) { return 0 }
    $r = [int]$reg.limit - [int]$reg.enviats
    if ($r -lt 0) { return 0 }
    return $r
}

# Suma n enviaments al registre (sense tocar el disc). PURA.
function _QuotaSuma($reg, [int]$n) {
    if ($null -eq $reg) { return $null }
    if ($n -lt 0) { $n = 0 }
    return @{ mes = [string]$reg.mes; enviats = ([int]$reg.enviats + $n); limit = [int]$reg.limit }
}

# Llegeix el comptador del disc, ja normalitzat al mes d'avui.
function _QuotaLlegeix {
    $mes = _QuotaMesActual (Get-Date)
    $path = _QuotaPath
    return (_QuotaNormalitza (Read-JsonFile $path) $mes)
}

# Escriu el comptador (UTF-8 sense BOM), creant la carpeta si cal.
function _QuotaDesa($reg) {
    if ($null -eq $reg) { return }
    $path = _QuotaPath
    $dir = Split-Path -Parent $path
    try {
        Write-JsonFile $path ([pscustomobject]$reg) 5
    } catch { }
}

# Apunta n enviaments fets. No llanca mai: un error del comptador no pot fer
# fracassar un correu que JA s'ha enviat.
function _QuotaApunta([int]$n) {
    try {
        $reg = _QuotaLlegeix
        _QuotaDesa (_QuotaSuma $reg $n)
    } catch { }
}
