#requires -Version 5.1
<#
.SYNOPSIS
  Editor dels TEXTOS del correu que l'app mòbil envia al titular (EmailJS).

.DESCRIPTION
  Només hi ha dos camps: ASSUMPTE i COS. Al cos hi surt TOT; els requeriments
  seleccionats s'insereixen allà on posis la variable {REQUERIMENTS}. Viuen a
  docs\dades\email-textos.json, que l'app mòbil (docs\app.js) llegeix. En desar
  s'escriu el JSON i es publica la propera vegada que es faci Actualitzar.bat
  (que ja puja docs\dades\email-textos.json, pas 2b).

  Variables disponibles al cos i a l'assumpte:
    {CAPCALERA}     les dades de l'activitat, com a la capcalera de l'informe
                    ('0 CAPCALERA'); sola a la seva linia
    {INTRO}         la frase que introdueix els requeriments segons l'origen
                    (introDoc / introDocSenseAnotacio / introInsp del JSON)
    {REQUERIMENTS}  la llista de requeriments (deficiències) seleccionats
    {ID_GIA} {ADRECA} {ACTIVITAT} {TITULAR}   dades de l'activitat
    {DATA}          data d'avui (dd/MM/yyyy)
  El cos admet **negreta** (dos asteriscs) i els enllaços http(s) es tornen
  clicables sols.

  Funcions PURES (rutes, valors per defecte, càrrega/desat) testejables en
  headless; la finestra (WinForms) només a Windows.
#>

# ----------------------------------------------------------------------------
# FUNCIONS PURES
# ----------------------------------------------------------------------------

# Ruta del fitxer de textos (docs\dades\email-textos.json a l'arrel del clone).
function _EmailTextosPath {
    $root = if ($RepoRoot) { $RepoRoot } else { (Get-Location).Path }
    return (Join-Path $root (Join-Path 'docs' (Join-Path 'dades' 'email-textos.json')))
}

# EL JSON ES L'UNIC ORIGEN. NO hi ha cap copia dels textos escrita al codi.
#
# N'hi havia TRES -aqui, a docs\app.js (EMAIL_TEXTOS_DEFAULT) i al propi JSON- i
# el que les mantenia sincronitzades era un comentari. Van divergir: algu va
# afegir l'enllac castella de la seu des de la pantalla "Textos del correu" -que
# escriu el JSON- i les dues copies es van quedar enrere. O sigui que quan el
# mobil no podia llegir el fitxer enviava al titular una versio VELLA del correu,
# sense aquell enllac, i no ho deia ningu.
#
# Ara si el fitxer no hi es o no es valid, PETA amb un missatge clar. Es la
# mateixa regla que el projecte ja aplica als catalegs: val mes petar que enviar
# a un ciutada un text que no es el que toca. El fitxer esta al repositori, o
# sigui que 'git checkout -- docs/dades/email-textos.json' el recupera.


# Text d'ajuda amb les variables disponibles.
function _EmailTextosAjuda {
    return ('Variables: {CAPCALERA} = les dades de l''activitat (com a l''informe)  ' + [char]0x00B7 + '  {INTRO} = la frase segons l''origen  ' + [char]0x00B7 + '  {REQUERIMENTS} = els requeriments  ' + [char]0x00B7 + '  {ID_GIA} {EXP_NUM} {ADRECA} {ACTIVITAT} {TITULAR} {DATA}   ' + [char]0x00B7 + '   **negreta**  //cursiva//  !!vermell!!   ' + [char]0x00B7 + '   els enllacos http es fan clicables')
}

# Les frases que introdueixen els requeriments segons l'origen de l'informe.
$Script:EmailTextosIntroClaus = @('introDoc', 'introDocSenseAnotacio', 'introInsp')

# Llegeix els textos del correu. Torna un ordered hashtable amb assumpte, cos i
# bcc. PETA si el fitxer no hi es, no es valid o li falta alguna clau: no hi ha
# cap valor de reserva a que caure (vegeu el comentari de mes amunt).
function _LoadEmailTextos {
    $path = _EmailTextosPath
    $o = Read-JsonFile $path
    if ($null -eq $o) {
        throw ("No s'han pogut llegir els textos del correu: " + $path + [Environment]::NewLine +
               "El fitxer no hi es o no es un JSON valid. Recupera'l amb:" + [Environment]::NewLine +
               "    git checkout -- docs/dades/email-textos.json")
    }
    $d = [ordered]@{}
    foreach ($k in @('assumpte', 'cos')) {
        if (-not $o.PSObject.Properties[$k] -or [string]::IsNullOrEmpty([string]$o.$k)) {
            throw ("Els textos del correu no tenen la clau '" + $k + "': " + $path)
        }
        $d[$k] = [string]$o.$k
    }
    # LA CCO JA NO ES AQUI (octubre 2026): es a correu-mobil.json, fora del
    # GitHub (Read-CorreuMobilCco). La clau es queda perque els cridadors la
    # llegeixen igual.
    $d['bcc'] = @(Read-CorreuMobilCco)
    # A qui va el correu del mobil (Configuracio -> Correu del mobil).
    $d['dest'] = @(_EmailDestDeJson $o)
    # Les frases de {INTRO} (vegeu _CorreuIntro, CorreuFormat.ps1). No son
    # obligatories: un cos que no porta {INTRO} no les fa servir.
    foreach ($k in $Script:EmailTextosIntroClaus) {
        $d[$k] = if ($o.PSObject.Properties[$k]) { [string]$o.$k } else { '' }
    }
    return ,$d
}

# Llista de Copia Oculta (CCO) del JSON -> array d'objectes @{Addr; Default}.
#
# CONVENCIO: torna un array PLA i el crider hi posa @() (com
# _AutoFirmaCandidatePaths). NO fa 'return ,$out': combinar les dues coses hi
# posa la capa dues vegades i .Count val 1 -es la trampa que aquest projecte ja
# ha vist tres cops-.
# El format ({ addr, def }) el fan servir correu-mobil.json (local i al Drive)
# i, fins a l'octubre de 2026, la clau 'bcc' d'email-textos.json (la migracio
# encara la llegeix).
function _EmailBccDeJson($o) {
    $out = @()
    if ($null -eq $o -or -not $o.PSObject.Properties['bcc']) { return $out }
    foreach ($b in @($o.bcc)) {
        if ($null -eq $b) { continue }
        $addr = [string]$b.addr
        if ([string]::IsNullOrWhiteSpace($addr)) { continue }
        $out += @{ Addr = $addr.Trim(); Default = [bool]$b.def }
    }
    return $out
}

# A qui va el correu del MOBIL: 'titular' i/o 'representant' (octubre 2026).
# Sense la clau, el titular, que es el que el mobil ha fet sempre. Array PLA.
$Script:EmailMobilDestValids = @('titular', 'representant')
function _EmailDestDeJson($o) {
    if ($null -eq $o -or -not $o.PSObject.Properties['dest'] -or $null -eq $o.dest) { return @('titular') }
    return @(@($o.dest) | ForEach-Object { [string]$_ } | Where-Object { $Script:EmailMobilDestValids -contains $_ } | Select-Object -Unique)
}

# ----------------------------------------------------------------------------
# LA CCO DEL MOBIL, FORA DEL GITHUB (octubre 2026)
# ----------------------------------------------------------------------------
# L'usuari: "Podem posar les adreces en local i fora de GitHub perque no siguin
# publiques?". Les quatre adreces de CCO eren a docs\dades\email-textos.json,
# al repositori PUBLIC (la web les llegia d'alla). Ara:
#   - al PC, a %LOCALAPPDATA%\InformesCornella\correu-mobil.json;
#   - al mobil, del Drive PRIVAT (Dades/correu-mobil.json), on ja hi ha
#     activitats.json: Publish-CorreuMobilCco (CorreuMobil.ps1) l'hi puja.
# EL PRIMER COP, el PC se les porta sol (_CorreuMobilCcoMigra): de la clau 'bcc'
# del fitxer si encara hi es (i la hi treu), o de l'historial del git si ja ha
# baixat la versio sense. Aixi l'usuari no les ha de tornar a escriure.
function _CorreuMobilCcoPath {
    $base = [string]$env:LOCALAPPDATA
    if ([string]::IsNullOrWhiteSpace($base)) { $base = [System.IO.Path]::GetTempPath() }
    return [string](Join-Path $base (Join-Path 'InformesCornella' 'correu-mobil.json'))
}

# D'on surten les versions antigues d'email-textos.json (text, la mes nova
# primer). Es una variable perque les proves hi posen les seves; en proves, cap
# (no s'ha de llegir cap adreca de debo de l'historial).
$Script:CorreuMobilCcoHistoria = { _CorreuMobilCcoHistoriaGit }
function _CorreuMobilCcoHistoriaGit {
    if ($Script:HeadlessTest) { return $null }
    $root = if ($RepoRoot) { $RepoRoot } else { (Get-Location).Path }
    $rel = 'docs/dades/email-textos.json'
    $hs = @(& git -C $root log --format=%H -n 40 -- $rel 2>$null)
    if ($LASTEXITCODE -ne 0) { return $null }
    $out = New-Object System.Collections.ArrayList
    foreach ($h in $hs) {
        $t = (@(& git -C $root show ($h.Trim() + ':' + $rel) 2>$null)) -join "`n"
        if ($t) { [void]$out.Add($t) }
    }
    return ,$out.ToArray()
}

function _CorreuMobilCcoAJson($bcc) {
    return [ordered]@{ bcc = @(@($bcc) | Where-Object { $null -ne $_ -and [string]$_.Addr } | ForEach-Object { [pscustomobject]@{ addr = [string]$_.Addr; def = [bool]$_.Default } }) }
}

# LA MIGRACIO. Idempotent i mai peta. $historia: torna els textos antics (o
# $null si no es poden llegir: llavors no es dona per feta i es torna a provar).
# Torna que ha fet: 'fitxer' (les ha tret del fitxer), 'historia', 'res' o ''.
function _CorreuMobilCcoMigra([string]$jsonPath, [string]$localPath, [scriptblock]$historia) {
    $fet = ''
    try {
        $o = Read-JsonFile $jsonPath
        $hiEsLocal = Test-Path -LiteralPath $localPath
        if ($null -ne $o -and $o.PSObject.Properties['bcc']) {
            # Encara al fitxer del repositori: si no n'hi ha de locals, s'hi
            # copien; tant si com no, es treuen del fitxer (Actualitzar el pujara
            # sense).
            if (-not $hiEsLocal) { Write-JsonFile $localPath (_CorreuMobilCcoAJson @(_EmailBccDeJson $o)) 4; $fet = 'fitxer' }
            $sense = [ordered]@{}
            foreach ($p in @($o.PSObject.Properties)) { if ($p.Name -ne 'bcc') { $sense[$p.Name] = $p.Value } }
            Write-JsonFile $jsonPath $sense 5
            if (-not $fet) { $fet = 'res' }
            return $fet
        }
        if ($hiEsLocal) { return '' }
        $versions = $null
        if ($null -ne $historia) { $versions = & $historia }
        if ($null -eq $versions) { return '' }
        foreach ($t in @($versions)) {
            $v = $null
            try { $v = ConvertFrom-Json ([string]$t) } catch { continue }
            if ($null -ne $v -and $v.PSObject.Properties['bcc']) {
                Write-JsonFile $localPath (_CorreuMobilCcoAJson @(_EmailBccDeJson $v)) 4
                return 'historia'
            }
        }
        # Cap versio amb CCO: un fitxer buit, perque no es torni a buscar.
        Write-JsonFile $localPath (_CorreuMobilCcoAJson @()) 4
        return 'res'
    } catch { return $fet }
}

# La CCO del mobil (i la d'"Enviar correu" per defecte): @{ Addr; Default }.
# Array PLA (el cridador hi posa @()).
function Read-CorreuMobilCco {
    $local = _CorreuMobilCcoPath
    [void](_CorreuMobilCcoMigra (_EmailTextosPath) $local $Script:CorreuMobilCcoHistoria)
    $o = $null
    try { $o = Read-JsonFile $local } catch { $o = $null }
    return @(_EmailBccDeJson $o)
}

function Save-CorreuMobilCco($bcc) {
    Write-JsonFile (_CorreuMobilCcoPath) (_CorreuMobilCcoAJson $bcc) 4
}

# Escriu el JSON (UTF-8 sense BOM). $obj = ordered hashtable amb les claus.
function _SaveEmailTextos($obj) {
    Write-JsonFile (_EmailTextosPath) $obj 5
}

# ----------------------------------------------------------------------------
# INTERFICIE (WinForms) - nomes a Windows
# ----------------------------------------------------------------------------
function Invoke-EmailTextos {
    $textos = _LoadEmailTextos

    [void](Show-EditorAssumpteCos `
        -TextFinestra 'Textos del correu (mobil)' `
        -Titol 'Textos del correu' `
        -Subtitol ('El que envia l''app m' + [char]0x00F2 + 'bil al titular ' + [char]0x00B7 + ' es publica amb Actualitzar') `
        -Ajuda (_EmailTextosAjuda) `
        -Assumpte ([string]$textos['assumpte']) `
        -Cos ([string]$textos['cos']) `
        -EtiquetaRestaurar 'Descartar els canvis' `
        -Restaurar {
            # Torna a llegir el FITXER. Abans deia "Restaurar original" i hi
            # posava una copia dels textos escrita al codi; aquella copia ja no
            # existeix -era una de les tres que havien divergit- i el que de
            # debo es vol des d'aqui es desfer el que s'acaba d'escriure.
            $r = [System.Windows.Forms.MessageBox]::Show('Vols tornar als textos que hi ha desats? Perdras els canvis que no hagis desat.', 'Textos del correu', 'YesNo', 'Question')
            if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return $null }
            try { return (_LoadEmailTextos) } catch {
                [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Textos del correu', 'OK', 'Error') | Out-Null
                return $null
            }
        } `
        -Desa {
            param($v)
            if ([string]$v['cos'] -notmatch '\{REQUERIMENTS\}') {
                $r = [System.Windows.Forms.MessageBox]::Show("El cos no conte la variable {REQUERIMENTS}: els requeriments NO sortiran al correu.`n`nVols desar igualment?", 'Textos del correu', 'YesNo', 'Warning')
                if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return $false }
            }
            # La CCO ja no va en aquest fitxer (es local: correu-mobil.json).
            $out = [ordered]@{
                assumpte = [string]$v['assumpte']
                cos      = [string]$v['cos']
            }
            # Les frases de {INTRO} tampoc no s'editen aqui: el mateix motiu.
            foreach ($k in $Script:EmailTextosIntroClaus) { $out[$k] = [string]$textos[$k] }
            # Ni a qui va (Correu del mobil, CorreuMobil.ps1).
            $out['dest'] = [string[]]@($textos['dest'])
            try {
                _SaveEmailTextos $out
                [System.Windows.Forms.MessageBox]::Show("Textos desats.`n`nEs publicaran al mobil la propera vegada que facis Actualitzar.", 'Textos del correu', 'OK', 'Information') | Out-Null
                return $true
            } catch {
                [System.Windows.Forms.MessageBox]::Show("No s'han pogut desar:`n$($_.Exception.Message)", 'Textos del correu', 'OK', 'Error') | Out-Null
                return $false
            }
        })
}
