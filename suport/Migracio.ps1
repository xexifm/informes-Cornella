#requires -Version 5.1
<#
.SYNOPSIS
  Mou els fitxers LOCALS del clone a la carpeta 'local\'.

.DESCRIPTION
  Fins ara, tot el que es d'aquest ordinador i NO va al repositori estava
  escampat per l'arrel del clone:

      Informes generats\        Rutes generades\
      BASE DE DADES ACTIVITATS\ BASE DE DADES ACT_EXTR\
      ESTRUCTURALS\*.docx       (vistes generades des dels .json)

  Barrejat amb el que si que es del repositori, i amb un .gitignore que havia
  crescut a dotze regles (una d'elles posada despres d'un ensurt de privadesa:
  el registre de la signatura porta noms i adreces de titulars, i aquest
  repositori es PUBLIC).

  Ara tot aixo viu dins de 'local\', que s'ignora SENCERA. Aixi ja no es pot
  pujar per accident res que hi caigui a dins, i l'arrel del clone nomes te el
  que l'usuari ha de veure.

  NO es mou %LOCALAPPDATA%\InformesCornella\ (settings.json, lastreport.json,
  cache, copies de seguretat dels catalegs, credencials del Drive, running.pid):
  es estat per usuari de Windows, ha de sobreviure a tornar a clonar el
  repositori, i les credencials del Drive no han de viure mai dins d'una
  carpeta que es pugui comprimir i enviar.

  La migracio es IDEMPOTENT (es pot executar mil vegades) i no atura mai res:
  si una carpeta no es pot moure perque hi ha un fitxer obert al Word, avisa i
  la deixa on es; el proper cop ho tornara a provar.

  La crida Motor.ps1 en arrencar (nomes son uns quants Test-Path) i tambe
  Actualitzar.bat despres del 'pull'.
#>

# ----------------------------------------------------------------------------
# NOMS (un sol lloc)
# ----------------------------------------------------------------------------
# Subcarpetes de 'local\'. Qui necessiti una d'aquestes rutes l'ha de demanar
# aqui, no muntar-se-la pel seu compte.
# Arrel del clone, capturada AQUI (en carregar el fitxer) i no dins d'una
# funcio: aixi no depen de com s'hagi cridat ni del directori de treball.
$Script:MigracioRepoRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)

$Script:LocalDirName = 'local'
$Script:LocalSubdirs = [ordered]@{
    Informes       = 'informes-generats'
    Rutes          = 'rutes-generades'
    Activitats     = 'base-dades-activitats'
    ActExtr        = 'base-dades-actextr'
    Vistes         = 'vistes-catalegs'
    Seguiment      = 'seguiment-gia'
    Llicencies     = 'base-dades-llicencies'
    Geocodificacio = 'geocodificacio'
    Planol         = 'planol-activitats'
    Normativa      = 'normativa'
    Revisions      = 'revisions'
}

# ATENCIO al [string] del 'return': Join-Path es un CMDLET, i el que surt d'un
# cmdlet ve embolcallat en un PSObject. Aixo passa desapercebut gairebe sempre
# (PowerShell el desembolcalla sol), pero NO quan el valor s'ha de passar a una
# crida COM per REFERENCIA: $doc.SaveAs([ref]$ruta) peta amb
#   "no se puede convertir el valor ... de tipo psobject al tipo Object"
# Va passar exactament aixo amb les vistes en Word. El cast el desembolcalla.
#
# Nota per a les proves: Join-Path resol la UNITAT, o sigui que fora de Windows
# no se li pot passar 'C:\...' (peta amb "A drive with the name 'C' does not
# exist"). Les proves fan servir una arrel valida a la plataforma on corren.
# Carpeta on viu l'Excel d'activitats a la xarxa de la feina. ES EL VALOR PER
# DEFECTE: config.ps1 i, per sobre, settings.json d'aquest PC el poden canviar.
#
# Viu aqui i NOMES aqui perque el necessiten DOS processos: el programa
# (Motor.ps1) i el planificador de rutes (rutes\Ruta.ps1), que corre a part i no
# carrega Motor.ps1. Tots dos ja carreguen aquest fitxer, que es on el projecte
# ja te els altres noms de carpeta (vegeu $Script:LocalSubdirs aqui sobre).
# Abans la ruta estava escrita a tots dos llocs, i el CLAUDE.md donava aquesta
# duplicacio per resolta quan nomes ho estava per a les subcarpetes de local\.
function Get-ActivitatsDirDefault {
    return 'I:\Activitats_Ordenances\Activitats\5.- Sergi Fadurdo\2_Controls Excels'
}

function Get-LocalDir([string]$repoRoot) {
    return [string](Join-Path $repoRoot $Script:LocalDirName)
}

# Ruta d'una subcarpeta de 'local\' PEL SEU NOM LOGIC (Informes, Rutes...).
# No crea res: nomes calcula la ruta.
function Get-LocalSubdir([string]$repoRoot, [string]$clau) {
    if (-not $Script:LocalSubdirs.Contains($clau)) { throw "Subcarpeta local desconeguda: $clau" }
    return [string](Join-Path (Get-LocalDir $repoRoot) $Script:LocalSubdirs[$clau])
}

# LA BASE D'INFORMES. El nom del fitxer viu AQUI, al costat dels noms de les
# subcarpetes, i pel mateix motiu: estava escrit a CINC llocs
# (InformesEscaneig, Informes, ComprovarExcel, Recordatoris i rutes/Planol),
# cada un amb el seu Join-Path. Els noms de les CARPETES ja es van centralitzar
# un dia; el del fitxer no havia rebut el mateix tracte.
#
# Llegeix l'ambit ($LocalActivitatsDir) i no es calcula des de $RepoRoot a
# posta: ValidarClassificacio.ps1 el desvia a una carpeta temporal per comparar
# un escaneig de debo amb la classificacio feta a ma, i calculant-lo des de
# l'arrel aquell desviament deixaria de funcionar en silenci.
# El defineixen els DOS processos (Activitats.ps1 i rutes/Ruta.ps1), i aquest
# fitxer el carreguen tots dos.
function Get-InformesDbPath {
    return [string](Join-Path $LocalActivitatsDir 'informes-db.json')
}

# ----------------------------------------------------------------------------
# FUNCIO PURA (testejable): que s'ha de moure
# ----------------------------------------------------------------------------
# Parells { Origen; Desti; Tipus } de les CARPETES antigues. Nomes calcula les
# rutes; no mira el disc (aixi es pot provar sense muntar cap arbre de fitxers).
function Get-MigracionsLocal([string]$repoRoot) {
    $out = New-Object System.Collections.ArrayList
    $mapa = @(
        @{ Vell = 'Informes generats';        Clau = 'Informes' }
        @{ Vell = 'Rutes generades';          Clau = 'Rutes' }
        @{ Vell = 'BASE DE DADES ACTIVITATS'; Clau = 'Activitats' }
        @{ Vell = 'BASE DE DADES ACT_EXTR';   Clau = 'ActExtr' }
    )
    foreach ($m in $mapa) {
        [void]$out.Add([pscustomobject]@{
            Origen = (Join-Path $repoRoot $m.Vell)
            Desti  = (Get-LocalSubdir $repoRoot $m.Clau)
            Tipus  = 'carpeta'
        })
    }
    # Array PLA (no ,$ArrayList): aixi @() l'enumera element a element.
    return $out.ToArray()
}

# ----------------------------------------------------------------------------
# APLICACIO
# ----------------------------------------------------------------------------
function _MouContingut([string]$origen, [string]$desti) {
    # Mou el CONTINGUT (no la carpeta): si el desti ja existeix amb alguna cosa
    # a dins, els fitxers nous s'hi afegeixen i els que ja hi son es respecten
    # (mai es trepitja res del desti).
    if (-not (Test-Path -LiteralPath $origen -PathType Container)) { return 0 }
    if (-not (Test-Path -LiteralPath $desti)) {
        New-Item -ItemType Directory -Path $desti -Force -ErrorAction Stop | Out-Null
    }
    $moguts = 0
    foreach ($it in @(Get-ChildItem -LiteralPath $origen -Force -ErrorAction SilentlyContinue)) {
        $dst = Join-Path $desti $it.Name
        if (Test-Path -LiteralPath $dst) { continue }   # ja hi es: no el toquem
        try { Move-Item -LiteralPath $it.FullName -Destination $dst -Force -ErrorAction Stop; $moguts++ }
        catch {
            # Cas tipic: el fitxer esta obert (Word, Acrobat...). No es cap
            # problema: es queda on es i es torna a provar el proper cop.
            Write-Host ("  '{0}' esta obert en un altre programa: el deixo on es i ho tornare a provar." -f $it.Name)
            $Script:MigracioPendents = $true
        }
    }
    # Nomes esborrem la carpeta vella si ha quedat BUIDA del tot.
    if (@(Get-ChildItem -LiteralPath $origen -Force -ErrorAction SilentlyContinue).Count -eq 0) {
        try { Remove-Item -LiteralPath $origen -Force -ErrorAction Stop } catch { }
    }
    return $moguts
}

# Les VISTES en Word dels catalegs. ESTRUCTURALS es queda NOMES amb les fonts:
# els .json i '0 CAPCALERA.docx', que si que es una plantilla de veritat i no es
# pot regenerar. La resta de .docx son derivats i van a local\vistes-catalegs\.
function _MouVistesCatalegs([string]$repoRoot) {
    $estr = Join-Path $repoRoot 'ESTRUCTURALS'
    if (-not (Test-Path -LiteralPath $estr -PathType Container)) { return 0 }
    $desti = Get-LocalSubdir $repoRoot 'Vistes'
    $moguts = 0
    foreach ($f in @(Get-ChildItem -LiteralPath $estr -Filter '*.docx' -File -ErrorAction SilentlyContinue)) {
        if ($f.Name -eq '0 CAPCALERA.docx') { continue }
        if ($f.Name.StartsWith('~$'))       { continue }
        if (-not (Test-Path -LiteralPath $desti)) {
            New-Item -ItemType Directory -Path $desti -Force -ErrorAction Stop | Out-Null
        }
        $dst = Join-Path $desti $f.Name
        try { Move-Item -LiteralPath $f.FullName -Destination $dst -Force -ErrorAction Stop; $moguts++ }
        catch { Write-Host ("  avis: no s'ha pogut moure la vista '{0}' ({1})" -f $f.Name, $_.Exception.Message) }
    }
    return $moguts
}

# CATALEGS QUE HAN CANVIAT DE NOM. MNSTRAS.json es va dir MNSTRANS.json
# (octubre 2026: l'usuari, "l'informe de transmissio s'hauria de dir TRANS").
# Si l'usuari havia editat el vell, Actualitzar.bat el torna a posar a
# ESTRUCTURALS despres del pull ("la versio de l'usuari preval") i hi quedaria
# un cataleg que ja no llegeix ningu. Es MOU (mai s'esborra) a
# local\catalegs-antics\, on la seva feina es pot recuperar.
$Script:CatalegsReanomenats = @(@{ Vell = 'MNSTRAS.json'; Nou = 'MNSTRANS.json' })

function _JubilaCatalegsReanomenats([string]$repoRoot) {
    $estr = Join-Path $repoRoot 'ESTRUCTURALS'
    $moguts = 0
    foreach ($c in $Script:CatalegsReanomenats) {
        $vell = Join-Path $estr ([string]$c.Vell)
        if (-not (Test-Path -LiteralPath $vell -PathType Leaf)) { continue }
        if (-not (Test-Path -LiteralPath (Join-Path $estr ([string]$c.Nou)) -PathType Leaf)) { continue }
        $desti = Join-Path (Get-LocalDir $repoRoot) 'catalegs-antics'
        try {
            if (-not (Test-Path -LiteralPath $desti)) { New-Item -ItemType Directory -Path $desti -Force -ErrorAction Stop | Out-Null }
            $dst = Join-Path $desti ([System.IO.Path]::GetFileNameWithoutExtension([string]$c.Vell) + '-' + (Get-Date).ToString('yyyyMMdd-HHmmss') + '.json')
            Move-Item -LiteralPath $vell -Destination $dst -Force -ErrorAction Stop
            Write-Host ("  El cataleg '{0}' ara es diu '{1}': el teu antic es a '{2}'." -f $c.Vell, $c.Nou, $dst)
            $moguts++
        } catch { Write-Host ("  avis: no s'ha pogut moure '{0}' ({1})" -f $c.Vell, $_.Exception.Message) }
    }
    return $moguts
}

# settings.json pot tenir rutes ABSOLUTES que apuntin a les carpetes velles
# (la pantalla de Configuracio hi desa el que l'usuari hagi triat). Si hi
# apunten, s'hi reescriu la nova; si l'usuari havia triat una carpeta seva de
# fora del clone, no s'hi toca.
function _ActualitzaSettingsLocal([string]$repoRoot) {
    $p = $null
    try { $p = $Script:SettingsPath } catch { }
    if ([string]::IsNullOrWhiteSpace($p) -or -not (Test-Path -LiteralPath $p)) { return $false }
    $o = Read-JsonFile $p
    if ($null -eq $o) { return $false }

    $parells = @(
        @{ Camp = 'OutputDir';      Vell = (Join-Path $repoRoot 'Informes generats'); Nou = (Get-LocalSubdir $repoRoot 'Informes') }
        @{ Camp = 'RutesOutputDir'; Vell = (Join-Path $repoRoot 'Rutes generades');   Nou = (Get-LocalSubdir $repoRoot 'Rutes') }
    )
    $canviat = $false
    foreach ($x in $parells) {
        if (-not $o.PSObject.Properties[$x.Camp]) { continue }
        $val = [string]$o.($x.Camp)
        if ([string]::IsNullOrWhiteSpace($val)) { continue }
        if ($val.TrimEnd('\') -ieq ([string]$x.Vell).TrimEnd('\')) {
            $o.($x.Camp) = [string]$x.Nou
            $canviat = $true
        }
    }
    if ($canviat) {
        try { Write-JsonFile $p $o 6 } catch { return $false }
    }
    return $canviat
}

# Fa la migracio si cal. Retorna el nombre d'elements moguts (0 = no calia res).
# NO llenca mai: si alguna cosa falla, avisa i el programa continua.
#
# $repoRoot buit -> l'arrel del clone deduida d'on viu aquest fitxer (suport\).
# Aixi Actualitzar.bat el pot cridar sense haver de passar-li cap ruta (i sense
# les mil punyetes de les cometes de cmd amb rutes acabades en barra).
# ----------------------------------------------------------------------------
# ELS .BAT DE local\ (per executar al PC el que no es del menu)
# ----------------------------------------------------------------------------
# El repas de contactes es valida contra una referencia feta a ma que porta
# dades personals (local\, mai al repositori), o sigui que nomes es pot fer al
# PC de l'usuari: l'usuari, "posa-hi un .bat a local\ com el de la
# classificacio". Com que local\ no es puja, els .bat no poden anar al
# repositori: els escriu aquesta migracio (en arrencar i a Actualitzar.bat), i
# nomes si no hi son o han canviat. ASCII pur i cap '^' entre cometes (vegeu la
# trampa del cmd a suport/CLAUDE.md). PURA: nom -> contingut.
function _BatsLocal {
    $cap = "@echo off`r`nrem Generat pel programa (suport\Migracio.ps1): es torna a escriure sol.`r`ncd /d `"%~dp0..`"`r`n"
    return [ordered]@{
        'ValidarContactes.bat' = ($cap +
            "echo Validant el repas de contactes contra la referencia (contactes-referencia_*.json)...`r`n" +
            "powershell -NoProfile -ExecutionPolicy Bypass -File `"suport\ValidarContactes.ps1`"`r`n" +
            "pause`r`n")
        'DiagnosticPdf.bat' = ($cap +
            "rem Arrossega un PDF (o un XML de l'e-TRAM) damunt d'aquest fitxer, o obre'l i tria'l.`r`n" +
            "powershell -NoProfile -ExecutionPolicy Bypass -STA -File `"suport\DiagnosticPdf.ps1`" -Pdf `"%~1`"`r`n" +
            "pause`r`n")
    }
}

function _PosaBatsLocal([string]$repoRoot) {
    $dir = Get-LocalDir $repoRoot
    if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
    $bats = _BatsLocal
    foreach ($k in $bats.Keys) {
        $p = Join-Path $dir $k
        $ara = if (Test-Path -LiteralPath $p) { [System.IO.File]::ReadAllText($p) } else { $null }
        if ($ara -ne $bats[$k]) { [System.IO.File]::WriteAllText($p, $bats[$k], [System.Text.Encoding]::ASCII) }
    }
}

function Invoke-MigracioLocal([string]$repoRoot = '') {
    if ([string]::IsNullOrWhiteSpace($repoRoot)) { $repoRoot = $Script:MigracioRepoRoot }
    if ([string]::IsNullOrWhiteSpace($repoRoot)) { return 0 }
    $total = 0
    $Script:MigracioPendents = $false
    try {
        foreach ($m in @(Get-MigracionsLocal $repoRoot)) {
            if (-not (Test-Path -LiteralPath $m.Origen -PathType Container)) { continue }
            $total += (_MouContingut $m.Origen $m.Desti)
        }
        $total += (_MouVistesCatalegs $repoRoot)
        $total += (_JubilaCatalegsReanomenats $repoRoot)
        if ($total -gt 0) {
            [void](_ActualitzaSettingsLocal $repoRoot)
            Write-Host ("Endrecat: {0} elements moguts a '{1}\'." -f $total, $Script:LocalDirName)
        }
        try { _PosaBatsLocal $repoRoot } catch { }
        if ($Script:MigracioPendents) {
            Write-Host "  (queda algun fitxer per moure perque estava obert; tanca'l i torna a fer Actualitzar.bat)"
        }
    } catch {
        Write-Host ("Avis: no s'ha pogut acabar d'endrecar la carpeta 'local' ({0})." -f $_.Exception.Message)
    }
    return $total
}
