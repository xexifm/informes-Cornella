#requires -Version 5.1
<#
.SYNOPSIS
  Configuracio LOCAL d'aquest ordinador (rutes de xarxa/carpeta).

.DESCRIPTION
  El programa te diverses rutes hardcodejades amb un valor per defecte pensat
  per a l'ordinador de la feina (unitat I:). suport/config.ps1 permet
  sobreescriure-les, pero es un fitxer VERSIONAT A GIT: editar-lo a casa (amb
  una ruta com F:\...) i despres fer Actualitzar.bat el podria acabar pujant a
  main i trencant la configuracio de l'ordinador de la feina.

  Aquest modul aplica una capa MES per sobre, nomes d'aquest ordinador: un
  settings.json a %LOCALAPPDATA%\InformesCornella\ (fora del repositori, mai
  es puja a git). La pantalla "Configuracio" del programa (Configuracio.ps1)
  hi desa els overrides; GenerarInforme.ps1 i rutes/Ruta.ps1 el llegeixen
  cadascu pel seu compte (son processos/scopes independents) DESPRES de
  carregar el seu config.ps1, aixi l'ordre de prioritat es:

    valor hardcodejat  <  suport/config.ps1 (compartit, git)  <  settings.json (aquest PC)

  Mateix idioma que Save-LastReport/Load-LastReport (GenerarInforme.ps1): la
  mateixa carpeta %LOCALAPPDATA%\InformesCornella, ConvertTo-Json/Set-Content
  -Encoding UTF8 per desar, i try/catch silencios (mai peta) per llegir.
#>

$Script:SettingsDir  = Join-Path $env:LOCALAPPDATA 'InformesCornella'
$Script:SettingsPath = Join-Path $Script:SettingsDir 'settings.json'

# Llegeix els overrides d'aquest ordinador. Retorna sempre un objecte (buit si
# no hi ha fitxer, esta buit o esta corrupte) perque el crider pugui fer
# $AppSettings.InformesDir sense comprovar res abans (PSCustomObject retorna
# $null en una propietat que no existeix, no peta).
function Load-AppSettings {
    if (-not (Test-Path -LiteralPath $Script:SettingsPath)) { return [pscustomobject]@{} }
    try {
        $raw = Get-Content -LiteralPath $Script:SettingsPath -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return [pscustomobject]@{} }
        $obj = $raw | ConvertFrom-Json
        if ($null -eq $obj) { return [pscustomobject]@{} }
        return $obj
    } catch {
        return [pscustomobject]@{}
    }
}

# Desa els overrides d'aquest ordinador (hashtable/objecte amb nomes les claus
# que es volen sobreescriure; vegeu _BuildSettingsOverrides). Crea la carpeta
# si cal. Retorna $true/$false segons si ha pogut desar.
function Save-AppSettings($overrides) {
    try {
        if (-not (Test-Path -LiteralPath $Script:SettingsDir)) {
            New-Item -ItemType Directory -Path $Script:SettingsDir -Force | Out-Null
        }
        ($overrides | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath $Script:SettingsPath -Encoding UTF8
        return $true
    } catch {
        return $false
    }
}

# Valor EFECTIU d'un camp: l'override d'aquest ordinador si n'hi ha (no buit),
# sino el valor per defecte (config.ps1 o hardcodejat). Funcio PURA.
function _ResolveEffectiveValue($override, $default) {
    if ([string]::IsNullOrWhiteSpace([string]$override)) { return [string]$default }
    return [string]$override
}

# Construeix l'objecte a desar a settings.json a partir dels valors que l'usuari
# ha deixat a la pantalla de Configuracio ($values) i dels valors "de
# repositori" ($defaults, capturats ABANS d'aplicar cap override). Nomes
# s'inclouen les claus on l'usuari ha escrit un valor NO BUIT i DIFERENT del
# per defecte -- aixi el JSON nomes conte les diferencies reals d'aquest PC
# (si l'usuari torna a deixar un camp igual al per defecte, deixa d'aparixer-
# hi sol, sense necessitat d'un boto "esborrar" per camp). $values i $defaults
# son hashtables/objectes amb les mateixes claus (InformesDir, ActivitatsDir,
# OutputDir, RutesOutputDir, DriveBaseDir). Funcio PURA.
function _BuildSettingsOverrides($values, $defaults) {
    $out = [ordered]@{}
    foreach ($key in $values.Keys) {
        $v = [string]$values[$key]
        $d = [string]$defaults[$key]
        if (-not [string]::IsNullOrWhiteSpace($v) -and $v -ne $d) {
            $out[$key] = $v
        }
    }
    return $out
}
