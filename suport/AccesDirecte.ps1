#requires -Version 5.1
<#
.SYNOPSIS
  Crea l'acces directe del programa, per poder-lo ANCORAR A LA BARRA DE TASQUES.

.DESCRIPTION
  Windows NO deixa ancorar un .bat (ni un .vbs) a la barra de tasques: nomes hi
  admet accessos directes que apuntin a un EXECUTABLE. Per aixo aqui no es fa un
  acces directe al GenerarInforme.bat, sino un que apunta a:

      wscript.exe  "<clone>\suport\GenerarInforme.vbs"

  ...que es exactament el que ja fa el .bat (el .vbs arrenca el PowerShell sense
  cap finestra de consola). Com que el desti es un .exe, l'acces directe SI que
  es pot ancorar.

  Es deixa a DOS llocs:
    - l'ESCRIPTORI, per tenir-lo a ma;
    - el MENU INICI (%APPDATA%\...\Start Menu\Programs), que es el que fa que
      surti buscant "Generador d'informes" i que es pugui ancorar des d'alli.

  I li posa l'ESCUT (suport\cornella.ico), perque a la barra de tasques es
  reconegui d'un cop d'ull i no surti la icona generica del Windows.

  PER QUE NO S'ANCORA SOL: des del Windows 10, el verb "Ancorar a la barra de
  tasques" ja no es pot invocar per codi (Microsoft el va treure de la interficie
  del shell). El que es troba per internet son trucs que escriuen al registre i
  reinicien l'explorer: son fragils i es poden carregar la barra de tasques de
  l'usuari. Val mes deixar l'acces directe fet i que l'usuari faci un clic dret.

.NOTES
  Nomes te sentit a Windows (WScript.Shell). Les funcions de RUTA son pures i es
  proven a Linux; la creacio de l'acces directe, no.

  CONVENCIO ASCII: el codi no porta accents.
#>

# El nom del fitxer de l'acces directe. Sense accents ni apostrof tipografic:
# es un nom de fitxer i ha de sobreviure a qualsevol codepage.
$Script:AccesDirecteNom = "Generador d'informes Cornella.lnk"

# ON APUNTA l'acces directe, a partir de l'arrel del clone. Funcio PURA: no toca
# el disc, o sigui que es pot provar a qualsevol plataforma.
function Get-AccesDirecteObjectiu([string]$repoRoot, [string]$systemRoot = '') {
    if ([string]::IsNullOrWhiteSpace($systemRoot)) { $systemRoot = 'C:\Windows' }
    $sep = '\'
    $rr = ([string]$repoRoot).TrimEnd('\', '/')
    return @{
        Desti     = ($systemRoot.TrimEnd('\') + $sep + 'System32' + $sep + 'wscript.exe')
        Arguments = ('"' + $rr + $sep + 'suport' + $sep + 'GenerarInforme.vbs"')
        Carpeta   = $rr
        Icona     = ($rr + $sep + 'suport' + $sep + 'cornella.ico')
        Nom       = $Script:AccesDirecteNom
    }
}

# Els llocs on es deixa. Funcio PURA (rep les carpetes ja resoltes).
function Get-AccesDirecteDestins([string]$escriptori, [string]$menuInici) {
    $out = New-Object System.Collections.ArrayList
    foreach ($d in @($escriptori, $menuInici)) {
        if ([string]::IsNullOrWhiteSpace($d)) { continue }
        [void]$out.Add(($d.TrimEnd('\') + '\' + $Script:AccesDirecteNom))
    }
    return $out.ToArray()
}

# Crea (o refresca) l'acces directe. Retorna @{ Ok; Fets; Errors }.
function New-AccesDirecteInformes([string]$repoRoot = '') {
    if ([string]::IsNullOrWhiteSpace($repoRoot)) { $repoRoot = $RepoRoot }
    $obj = Get-AccesDirecteObjectiu $repoRoot $env:SystemRoot
    $fets = New-Object System.Collections.ArrayList
    $errs = New-Object System.Collections.ArrayList

    $vbs = $obj.Arguments.Trim('"')
    if (-not (Test-Path -LiteralPath $vbs)) {
        return @{ Ok = $false; Fets = @(); Errors = @(("No trobo " + $vbs)) }
    }

    $sh = $null
    try { $sh = New-Object -ComObject WScript.Shell } catch {
        return @{ Ok = $false; Fets = @(); Errors = @('No s''ha pogut crear l''acces directe (WScript.Shell).') }
    }
    $destins = Get-AccesDirecteDestins ([Environment]::GetFolderPath('Desktop')) ([Environment]::GetFolderPath('Programs'))
    foreach ($ruta in $destins) {
        try {
            $carpeta = Split-Path -Parent $ruta
            if (-not (Test-Path -LiteralPath $carpeta)) { New-Item -ItemType Directory -Path $carpeta -Force | Out-Null }
            $lnk = $sh.CreateShortcut($ruta)
            $lnk.TargetPath = [string]$obj.Desti
            $lnk.Arguments = [string]$obj.Arguments
            $lnk.WorkingDirectory = [string]$obj.Carpeta
            $lnk.Description = "Generador d'informes - Ajuntament de Cornella de Llobregat"
            if (Test-Path -LiteralPath ([string]$obj.Icona)) { $lnk.IconLocation = [string]$obj.Icona }
            $lnk.Save()
            [void]$fets.Add($ruta)
        } catch {
            [void]$errs.Add(($ruta + ': ' + $_.Exception.Message))
        }
    }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sh) | Out-Null } catch { }
    return @{ Ok = ($fets.Count -gt 0); Fets = $fets.ToArray(); Errors = $errs.ToArray() }
}

# El que crida Crear-acces-directe.bat: crea l'acces directe i ho explica per
# pantalla. Va aqui i no al .bat perque una ordre de PowerShell escrita dins
# d'un .bat, amb cometes i canonades, es un niu d'errors (vegeu CLAUDE.md: dins
# de cometes dobles el cmd NO interpreta el '|', pero si que deixa passar el '^'
# literal, i el que arriba al PowerShell ja no es el que havies escrit).
function Invoke-CrearAccesDirecte([string]$repoRoot = '') {
    if ([string]::IsNullOrWhiteSpace($repoRoot)) { $repoRoot = (Get-Location).Path }
    $r = New-AccesDirecteInformes $repoRoot
    Write-Host ''
    if ([bool]$r.Ok) {
        Write-Host "Acces directe creat a:" -ForegroundColor Green
        foreach ($f in @($r.Fets)) { Write-Host ('   ' + $f) }
    } else {
        Write-Host "No s'ha pogut crear l'acces directe." -ForegroundColor Red
    }
    foreach ($e in @($r.Errors)) { Write-Host ('   ' + $e) -ForegroundColor Yellow }
    return $r
}
