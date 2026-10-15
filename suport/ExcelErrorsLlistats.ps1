#requires -Version 5.1
<#
.SYNOPSIS
  "Excel Errors": els TRES LLISTATS per correu (Innocues, Bars i Restaurants,
  Annexos II i III), les activitats TRETES dels llistats i el correu.

.DESCRIPTION
  L'usuari (octubre 2026): "Vull automatitzar tambe els errors. Consistira en
  enviar un correu amb 3 Excels":
    - Innocues: Classificacio general annex L18-Proj. i Cert., L18-Cert. o
      Sectorial, i Classificacio general Apartat que NO sigui 561.
    - Bars i Restaurants: annex L18-Proj. i Cert. amb apartat 561, i annex
      Recreativa cataleg.
    - Annexos II i III: annex II i III.
  "Vull poder eliminar activitats d'aquests llistats, no m'importa si tenen
  errors."

  QUI QUEDA FORA DELS TRES (i per aixo la finestra ho diu, columna "Llistat" =
  "(cap)"): l'annex buit, l'annex I (I.1, I.2...), L18-Cert. o Sectorial amb
  l'apartat 561, i qualsevol altre valor que el GIA hi posi. No s'envien a
  ningu; es veuen a la finestra i al registre de l'automatic.

  Els tipus d'error que entren als llistats son els MARCATS a la finestra (la
  tria desada, Read-ExTriats): la mateixa tria per al boto i per a l'automatic.
  Les activitats TRETES es desen a local\base-dades-activitats\
  excel-errors-exclosos.json (porten ID GIA: mai al GitHub) i no surten ni a la
  finestra ni als llistats fins que es tornen a posar.

  El correu va a les adreces fixes de Configuracio -> Correus de cada eina ->
  Excel Errors, NOMES per l'Outlook (porta adjunts). El text, a Textos del
  correu (excel-errors-correu.json, a %LOCALAPPDATA%).

  Es carrega DESPRES de CorreuEines.ps1 (s'hi registra la prova i el text).
#>

# ----------------------------------------------------------------------------
# ELS TRES LLISTATS (l'ordre es el dels adjunts)
# ----------------------------------------------------------------------------
$Script:ExLlistats = [ordered]@{
    innocues = ('Inn' + [char]0x00F2 + 'cues')
    bars     = 'Bars i Restaurants'
    annexos  = 'Annexos II i III'
}

# A quin llistat va una activitat, per la classificacio de la fulla Estes, o ''
# si a cap. Es compara sense accents, majuscules, espais ni signes:
# "L18-Proj. i Cert." -> "l18projicert". L'apartat 561 amb el criteri de
# Controls periodics (561, 5610...; no 1561). PURA.
function _ExLlistatDe([string]$annex, [string]$apartat) {
    $a = (_NormalitzaText $annex) -replace '[^a-z0-9]', ''
    $es561 = [bool](_ControlPeriodicClassify '' $apartat).Is561
    if ($a -eq 'ii' -or $a -eq 'iii' -or $a -eq 'annexii' -or $a -eq 'annexiii') { return 'annexos' }
    if ($a -like 'recreativacatal*') { return 'bars' }
    $l18proj = ($a -like 'l18proj*')
    if ($l18proj -and $es561) { return 'bars' }
    if (($l18proj -or $a -eq 'l18cert' -or $a -eq 'sectorial') -and -not $es561) { return 'innocues' }
    return ''
}

# Els errors dels tipus marcats, sense els exclosos, repartits. Torna
# @{ innocues = @(errors); bars; annexos; SenseLlistat = @(@{ Id; Annex; Apartat }) }.
# PURA.
function _ExPerLlistat($errors, $tipusTriats, $exclosos) {
    $ex = @{}
    foreach ($e in @($exclosos)) { if ($null -ne $e) { $ex[[string]$e.Id] = $true } }
    $out = [ordered]@{}
    foreach ($k in @($Script:ExLlistats.Keys)) { $out[$k] = New-Object System.Collections.ArrayList }
    $sense = [ordered]@{}
    foreach ($e in @(_ExTriats @($errors | Where-Object { $null -ne $_ -and -not $ex.ContainsKey([string]$_.Id) }) $tipusTriats)) {
        $k = _ExLlistatDe ([string]$e.Annex) ([string]$e.Apartat)
        if ($k -eq '') {
            if (-not $sense.Contains([string]$e.Id)) { $sense[[string]$e.Id] = @{ Id = [string]$e.Id; Annex = [string]$e.Annex; Apartat = [string]$e.Apartat } }
            continue
        }
        [void]$out[$k].Add($e)
    }
    $r = @{ SenseLlistat = @($sense.Values) }
    foreach ($k in @($out.Keys)) { $r[$k] = $out[$k].ToArray() }
    return $r
}

# "12 errors (8 activitats)". PURA.
function _ExResumLlistat($errors) {
    $l = @($errors)
    $n = @($l | ForEach-Object { [string]$_.Id } | Sort-Object -Unique).Count
    return ([string]$l.Count + ' errors (' + $n + ' activitats)')
}

# El nom del fitxer de cada llistat, com els que feia l'usuari a ma
# ("Excel_errors_2026-10-15_Bars_i_Restaurants.xlsx"). PURA.
$Script:ExLlistatsFitxer = @{ innocues = 'Innocues'; bars = 'Bars_i_Restaurants'; annexos = 'Annexos_II_i_III' }
function _ExNomFitxerLlistat([string]$clau, [datetime]$data) {
    return ('Excel_errors_' + $data.ToString('yyyy-MM-dd') + '_' + [string]$Script:ExLlistatsFitxer[$clau] + '.xlsx')
}

# ----------------------------------------------------------------------------
# LES ACTIVITATS TRETES DELS LLISTATS
# ----------------------------------------------------------------------------
function _ExExclososPath {
    if ([string]::IsNullOrWhiteSpace($LocalActivitatsDir)) { return '' }
    return [string](Join-Path $LocalActivitatsDir 'excel-errors-exclosos.json')
}

# Llista de @{ Id; Text; El }. Mai peta.
function Read-ExExclosos([string]$path = '') {
    if (-not $path) { $path = _ExExclososPath }
    $out = New-Object System.Collections.ArrayList
    if (-not $path) { return $out.ToArray() }
    $o = $null
    try { $o = Read-JsonFile $path } catch { $o = $null }
    foreach ($d in @(_PropInf $o 'exclosos')) {
        if ($null -eq $d -or -not [string](_PropInf $d 'id')) { continue }
        [void]$out.Add(@{ Id = [string](_PropInf $d 'id'); Text = [string](_PropInf $d 'text'); El = [string](Read-JsonIso (_PropInf $d 'el')) })
    }
    return $out.ToArray()
}

function Save-ExExclosos($llista, [string]$path = '') {
    if (-not $path) { $path = _ExExclososPath }
    if (-not $path) { throw "No hi ha carpeta local per desar les activitats tretes dels llistats." }
    $dir = Split-Path -Parent $path
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $nets = [ordered]@{}
    foreach ($e in @($llista)) { if ($null -ne $e -and [string]$e.Id -and -not $nets.Contains([string]$e.Id)) { $nets[[string]$e.Id] = $e } }
    $ord = @($nets.Values | Sort-Object { _GiaNumeric $_.Id })
    Write-JsonFile $path ([ordered]@{ exclosos = @($ord | ForEach-Object { [ordered]@{ id = [string]$_.Id; text = [string]$_.Text; el = [string]$_.El } }) }) 4
}

# Hi afegeix ($treu = $false) o en treu ($treu = $true) els IDs. Torna la llista
# nova. PURA (amb la llista d'abans).
function _ExExclososAmb($abans, $items, [bool]$treu = $false) {
    $m = [ordered]@{}
    foreach ($e in @($abans)) { if ($null -ne $e) { $m[[string]$e.Id] = $e } }
    $ara = (Get-Date).ToString('o')
    foreach ($i in @($items)) {
        if ($null -eq $i) { continue }
        $id = [string]$i.Id
        if ($treu) { if ($m.Contains($id)) { $m.Remove($id) } }
        elseif (-not $m.Contains($id)) { $m[$id] = @{ Id = $id; Text = [string]$i.Text; El = $ara } }
    }
    return @($m.Values)
}

# ----------------------------------------------------------------------------
# ELS TRES EXCELS
# ----------------------------------------------------------------------------
function _ExLlistatsDir {
    try { return [string](Get-LocalSubdir $RepoRoot 'ExcelErrors') } catch { return '' }
}

# Escriu els tres .xlsx (amb les columnes de l'exportacio de la finestra; un
# llistat sense errors surt igualment, amb la capcalera i prou: tothom rep el
# seu). Torna @{ Ok; Error; Fitxers; Resum (nom -> "n errors (m activitats)");
# SenseLlistat; Dir }.
function New-ExcelErrorsLlistats($errors, $tipusTriats, $exclosos, [string]$dir = '', [datetime]$data = (Get-Date)) {
    if (-not $dir) { $dir = _ExLlistatsDir }
    if (-not $dir) { return @{ Ok = $false; Error = "No hi ha carpeta local per desar els llistats." } }
    try {
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $per = _ExPerLlistat $errors $tipusTriats $exclosos
        $fitxers = New-Object System.Collections.ArrayList
        $resum = [ordered]@{}
        foreach ($k in @($Script:ExLlistats.Keys)) {
            $desti = _GetUniqueOutputPath $dir (_ExNomFitxerLlistat $k $data)
            $tots = @($Script:ExTipus.Keys)
            $bytes = _NormativaXlsxBytes $Script:ExExportCapcalera @(_ExFilesExport @($per[$k]) $tots) $Script:ExExportAmples 'Errors'
            [System.IO.File]::WriteAllBytes($desti, $bytes)
            [void]$fitxers.Add([string]$desti)
            $resum[[string]$Script:ExLlistats[$k]] = _ExResumLlistat @($per[$k])
        }
        return @{ Ok = $true; Error = ''; Fitxers = $fitxers.ToArray(); Resum = $resum; SenseLlistat = @($per.SenseLlistat); Dir = $dir }
    } catch {
        return @{ Ok = $false; Error = $_.Exception.Message }
    }
}

# Llegeix l'Excel (i la base de contactes), en treu els errors i en fa els tres
# llistats amb la tria desada. Sense cap finestra: el fan servir el boto i
# l'automatic. Torna el de New-ExcelErrorsLlistats (+ Origen).
function Invoke-ExcelErrorsLlistats {
    $res = _ReadExcelErrors
    if (-not [bool]$res.Ok) { return @{ Ok = $false; Error = [string]$res.Error } }
    $errors = @(Get-ExcelErrors @($res.Files) $res.Col (_ExContactes))
    $g = New-ExcelErrorsLlistats $errors @(Read-ExTriats) @(Read-ExExclosos)
    if ($g.Ok) { $g.Origen = [string]$res.File }
    return $g
}

# ----------------------------------------------------------------------------
# EL CORREU
# ----------------------------------------------------------------------------
function _ExCorreuPath {
    $base = [string]$env:LOCALAPPDATA
    if ([string]::IsNullOrWhiteSpace($base)) { $base = [System.IO.Path]::GetTempPath() }
    return [string](Join-Path $base (Join-Path 'InformesCornella' 'excel-errors-correu.json'))
}

# El text que va proposar l'usuari, tal qual.
function _ExCorreuDefecte {
    $cos = @(
        'Podeu mirar aquests errors del GIA si us plau?? Us he fet un Excel amb les activitats que li pertoquen a cadascuna de vosaltres.'
        ''
        ('Gr' + [char]0x00E0 + 'cies!!')
    ) -join "`n"
    return [ordered]@{ assumpte = 'Dades incorrectes GIA'; cos = $cos }
}

function _ExCorreuAjuda {
    return ('Variables: {DATA} = la data d''avui  ' + [char]0x00B7 + '  {RESUM} = els errors de cada llistat  ' + [char]0x00B7 +
            '   **negreta**  //cursiva//   ' + [char]0x00B7 + '   els tres Excels van adjunts')
}

function _LoadExCorreu {
    $d = _ExCorreuDefecte
    $o = Read-JsonFile (_ExCorreuPath)
    if ($null -ne $o) {
        foreach ($k in @($d.Keys)) {
            if ($o.PSObject.Properties[$k] -and -not [string]::IsNullOrEmpty([string]$o.$k)) { $d[$k] = [string]$o.$k }
        }
    }
    return ,$d
}

function _SaveExCorreu($obj) { Write-JsonFile (_ExCorreuPath) $obj 5 }

# Les variables. $resum: llistat -> "n errors (m activitats)". PURA.
function _ExFillPh([string]$text, $resum, [datetime]$data) {
    $linies = @()
    if ($null -ne $resum) { $linies = @(@($resum.Keys) | ForEach-Object { '- ' + [string]$_ + ': ' + [string]$resum[$_] }) }
    return (_OmpleVariables $text ([ordered]@{ '{DATA}' = $data.ToString('dd/MM/yyyy'); '{RESUM}' = ($linies -join "`n") }))
}

function Invoke-ExcelErrorsCorreuTextos {
    $textos = _LoadExCorreu
    [void](Show-EditorAssumpteCos `
        -TextFinestra 'Text del correu (Excel Errors)' `
        -Titol 'Text del correu' `
        -Subtitol ('Errors del GIA, tres Excels ' + [char]0x00B7 + ' es desa en aquest ordinador') `
        -Ajuda (_ExCorreuAjuda) `
        -Assumpte ([string]$textos['assumpte']) `
        -Cos ([string]$textos['cos']) `
        -EtiquetaRestaurar 'Restaurar original' `
        -Restaurar {
            $r = [System.Windows.Forms.MessageBox]::Show('Vols recuperar el text original? (No es desa fins que premis Desar.)', 'Text del correu', 'YesNo', 'Question')
            if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return $null }
            return (_ExCorreuDefecte)
        } `
        -Desa {
            param($v)
            try {
                _SaveExCorreu ([ordered]@{ assumpte = [string]$v['assumpte']; cos = [string]$v['cos'] })
                [System.Windows.Forms.MessageBox]::Show('Text desat en aquest ordinador.', 'Text del correu', 'OK', 'Information') | Out-Null
                return $true
            } catch {
                [System.Windows.Forms.MessageBox]::Show("No s'ha pogut desar:`n$($_.Exception.Message)", 'Text del correu', 'OK', 'Error') | Out-Null
                return $false
            }
        })
}

# Envia els tres Excels ($gen: el de New-ExcelErrorsLlistats) als destinataris
# de l'eina 'excelerrors'. $silenci: l'automatic (els esborranys s'apunten
# perque el menu ho avisi; mai la CCO "a triar"). Torna @{ Ok; Text }.
function Send-ExcelErrorsCorreu($gen, [bool]$silenci, $ccoTriades = @()) {
    $cfgE = Get-CorreuEina 'excelerrors'
    $to = [string]$cfgE.Fixes
    if (-not $to) { return @{ Ok = $false; Text = ("No hi ha cap destinatari: posa'l a Configuraci" + [char]0x00F3 + " -> Correus de cada eina -> Excel Errors (adreces fixes).") } }
    $motiu = Test-CorreuViaLlest $cfgE.Via
    if ($motiu) { return @{ Ok = $false; Text = $motiu } }
    $t = _LoadExCorreu
    $ara = Get-Date
    $assumpte = _ExFillPh ([string]$t['assumpte']) $gen.Resum $ara
    $html = _CosAHtml (_ExFillPh ([string]$t['cos']) $gen.Resum $ara)
    $cco = if ($null -ne $cfgE.Cco) { [string]$cfgE.Cco } else { '' }
    if (-not $silenci -and @($ccoTriades).Count -gt 0) { $cco = _CorreuCcoAmb $cco $ccoTriades }
    $ses = $null
    try {
        $ses = Open-CorreuSessio $cfgE.Via
        Send-CorreuSessio $ses $to $cco $assumpte $html '' @($gen.Fitxers)
        $esb = ($cfgE.Via -eq 'outlook-esborrany')
        if ($esb -and $silenci) { Add-CorreuEsborranysPendents 'Excel Errors' @(("Errors del GIA, tres Excels (" + $to + ")")) }
        [void](_ExAutoDesaEstat @{ mode = $(if ($silenci) { 'auto' } else { 'manual' }); enviat_el = $ara.ToString('o') })
        return @{ Ok = $true; Text = ($(if ($esb) { "Correu desat a Esborranys de l'Outlook" } else { 'Correu enviat' }) + ' per a ' + $to + $(if ($cco) { ' (CCO: ' + $cco + ')' } else { '' }) + '.') }
    } catch {
        $txt = if ($null -eq $ses) { [string]$_.Exception.Message } else { _CorreuSessioError $ses $_ }
        return @{ Ok = $false; Text = $txt }
    } finally {
        Close-CorreuSessio $ses
    }
}

# L'estat de l'automatic (l'interruptor i qui va fer l'ultim enviament). Es aqui
# i no a ExcelErrorsAutomatic.ps1 perque l'enviament (tambe el del boto) hi
# apunta qui l'ha fet: si no, els dos fitxers dependrien l'un de l'altre.
$Script:ExEstatPlantilla = [ordered]@{ auto = $false; auto_el = ''; mode = ''; enviat_el = '' }
function _ExAutoStatePath {
    if ([string]::IsNullOrWhiteSpace($LocalActivitatsDir)) { return '' }
    return [string](Join-Path $LocalActivitatsDir 'excel-errors-auto.json')
}
function _ExAutoEstat { return (Read-EstatAuto (_ExAutoStatePath) $Script:ExEstatPlantilla) }
function _ExAutoDesaEstat($canvis) { return (Save-EstatAuto (_ExAutoStatePath) $canvis $Script:ExEstatPlantilla) }

# ----------------------------------------------------------------------------
# ELS REGISTRES DEL CORREU (Configuracio -> Correus de cada eina, Textos)
# ----------------------------------------------------------------------------
# La prova: el text d'aquest correu i, adjunts, els darrers tres Excels que hi
# hagi a la carpeta (sense, va sense adjunts i ho diu).
$Script:CorreuProves['excelerrors'] = {
    param($gia, $cache, $cfgE)
    $t = _LoadExCorreu
    $adj = @()
    try {
        $dir = _ExLlistatsDir
        if ($dir -and (Test-Path -LiteralPath $dir)) { $adj = @(Get-ChildItem -LiteralPath $dir -Filter '*.xlsx' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 3 | ForEach-Object { $_.FullName }) }
    } catch { $adj = @() }
    $ara = Get-Date
    $html = _CosAHtml (_ExFillPh ([string]$t['cos']) $null $ara)
    if ($adj.Count -eq 0) { $html = '<p><i>(Sense adjunts: encara no s''han fet els llistats.)</i></p>' + $html }
    return @{
        Assumpte = (_ExFillPh ([string]$t['assumpte']) $null $ara)
        Html = $html
        Destinataris = @(_CorreuLlistaAdreces ([string]$cfgE.Fixes))
        Adjunts = $adj
        CcoAbans = ''
    }
}
$Script:CorreuCcoAbans['excelerrors'] = { '' }
$Script:CorreuTextosEditors['excelerrors'] = @{ Desc = "El correu amb els tres Excels d'errors del GIA (Innocues, Bars i Restaurants, Annexos II i III)."; Obre = { Invoke-ExcelErrorsCorreuTextos } }
