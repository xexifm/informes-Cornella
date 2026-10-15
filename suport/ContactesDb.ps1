#requires -Version 5.1
<#
  ContactesDb.ps1 - LA BASE DE CONTACTES (local\base-dades-activitats\
  contactes-db.json) i el REPAS que en fa "Actualitzar base".

  L'usuari (octubre 2026): quan es prem "Actualitzar base" (el boto i el mode
  automatic), a mes de l'estat de cada activitat, un repas de les dades de
  CONTACTE de cada ID GIA: completar les que falten a l'Excel d'activitats i
  avisar de les que hi son malament (el cas mes frequent: el TECNIC posat com a
  representant legal, o el seu correu o el seu mobil com si fossin del
  titular). L'Excel es una exportacio del GIA i EL PROGRAMA NO L'HA DE TOCAR
  MAI: el que surt es una base propia, la finestra "Contactes" per revisar-la
  (ContactesPantalla.ps1) i una llista de correccions per entrar-les al GIA.

  QUINS DOCUMENTS: TOT L'ARBRE de cada carpeta d'activitat, subcarpetes
  incloses (Get-FitxersContactes; l'usuari, octubre 2026: "no es llegeixen les
  subcarpetes"), i nomes els que el nom diu que son un dels tres tipus que
  porten les dades (_CtTipusDocument). L'ID GIA de cada document es el de LA
  CARPETA MES PROPERA que ho digui (_CtGiaDocument): el nom, els informes que
  hi ha o el seu expedient; el GIA que diu el document, nomes si cap no en diu.

  INCREMENTAL, com els informes: per cada document es desa la ruta relativa, la
  data de modificacio i el que se n'ha tret, i nomes es tornen a llegir els
  nous o modificats (o tots, si canvia $Script:ContactesVersio).

  ELS PDF, amb el lector propi (PdfText.ps1). Si no en surt text, el Word
  (lent: nomes per a aquests). Si tampoc, son ESCANEJATS: a la llista "no
  llegibles".

  LES CORRECCIONS A MA (correccions, per GIA) es desen a la mateixa base i
  MANEN: cada passada les torna a aplicar damunt de l'automatic
  (_CtManDeCorreccions). Un sol escaneig alhora: el mateix mutex que la base
  d'informes ($Script:BaseMutexNom), i l'automatic escriu al mateix registre.

  Porta dades personals: tot a local\, que no es puja mai.
#>

# Si canvia el que es treu d'un document (els lectors de ContactesExtraccio),
# puja-la: es tornen a llegir tots.
$Script:ContactesVersio = '2026-10-15.1'

function Get-ContactesDbPath {
    if ([string]::IsNullOrWhiteSpace($LocalActivitatsDir)) { return '' }
    return [string](Join-Path $LocalActivitatsDir 'contactes-db.json')
}

# La configuracio del repas (local\base-dades-activitats\contactes-config.json):
#   { "recintes": ["28", "1324"] }   els recintes amb organitzadors d'actes
function Read-ContactesConfig {
    $cfg = @{ Recintes = @($Script:CtRecintesDefecte) }
    if ([string]::IsNullOrWhiteSpace($LocalActivitatsDir)) { return $cfg }
    $p = Join-Path $LocalActivitatsDir 'contactes-config.json'
    if (-not (Test-Path -LiteralPath $p)) { return $cfg }
    $o = Read-JsonFile $p
    if ($null -ne $o -and $null -ne $o.PSObject.Properties['recintes']) { $cfg.Recintes = @(@($o.recintes) | ForEach-Object { [string]$_ } | Where-Object { $_ -ne '' }) }
    return $cfg
}

# ----------------------------------------------------------------------------
# ELS TECNICS CONEGUTS (local\base-dades-activitats\tecnics-coneguts_*.json)
# ----------------------------------------------------------------------------
# El MES RECENT pel nom (porta la data). Format:
#   { "tecnics": [ { "email", "noms": [], "empresa": [], "rol", "activitats": [] } ] }
function Find-TecnicsConeguts {
    if ([string]::IsNullOrWhiteSpace($LocalActivitatsDir)) { return $null }
    $c = @(Get-ChildItem -LiteralPath $LocalActivitatsDir -Filter 'tecnics-coneguts_*.json' -File -ErrorAction SilentlyContinue | Sort-Object Name -Descending)
    if ($c.Count -eq 0) { return $null }
    return $c[0]
}

# Torna @{ Emails = email -> @{ noms; empresa; rol }; Noms; Fitxer; Error }.
# Sense fitxer, buit (no es cap error: el repas va igual, amb menys senyals).
function _CtTecnicsDe($obj) {
    $r = @{ Emails = @{}; Noms = @() }
    if ($null -eq $obj) { return $r }
    $noms = New-Object System.Collections.ArrayList
    foreach ($t in @(_CtV $obj 'tecnics')) {
        if ($null -eq $t) { continue }
        $e = _CtEmailNet ([string](_CtV $t 'email'))
        $ns = @(@(_CtV $t 'noms') | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { [string]$_ })
        foreach ($n in $ns) { if (@(_CtNomParaules $n).Count -ge 2) { [void]$noms.Add($n) } }
        if ($e -eq '') { continue }
        $r.Emails[$e] = @{ noms = $ns; empresa = @(@(_CtV $t 'empresa') | ForEach-Object { [string]$_ }); rol = [string](_CtV $t 'rol') }
    }
    $r.Noms = $noms.ToArray()
    return $r
}

function Read-TecnicsConeguts([string]$path = '') {
    $f = if ($path -ne '') { Get-Item -LiteralPath $path -ErrorAction SilentlyContinue } else { Find-TecnicsConeguts }
    if ($null -eq $f) { $r = _CtTecnicsDe $null; $r.Fitxer = ''; $r.Error = ''; return $r }
    $o = Read-JsonFile $f.FullName
    $r = _CtTecnicsDe $o
    $r.Fitxer = $f.Name
    $r.Error = if ($null -eq $o) { "El fitxer de t" + [char]0x00E8 + "cnics coneguts no es un JSON valid: " + $f.Name } else { '' }
    return $r
}

# ----------------------------------------------------------------------------
# LLEGIR UN DOCUMENT
# ----------------------------------------------------------------------------
# L'XML amb la codificacio que diu la seva capcalera (XmlDocument.Load ho fa).
function _CtTextXml([string]$path) {
    $x = New-Object System.Xml.XmlDocument
    $x.XmlResolver = $null
    $x.Load($path)
    return [string]$x.OuterXml
}

# El text d'un PDF: el lector propi i, si no en surt (xifrat, glifs), el Word.
# $word: @{ App; Provat } compartit per tota la passada (el Word s'obre UN cop,
# nomes si cal, i el tanca qui l'ha creat). Torna @{ Text; Via; Error }.
# Torna @{ Text; Via; Error; PendentWord } (PendentWord: no s'ha pogut provar el
# Word perque la passada no el deixa fer servir; la proxima que si, el
# tornara a llegir).
function _CtTextPdf([string]$path, $word) {
    $err = ''
    $t = ''
    try { $t = Get-PdfText $path } catch { $err = [string]$_.Exception.Message; $t = '' }
    if (_PdfTeText $t) { return @{ Text = $t; Via = 'pdf'; Error = ''; PendentWord = $false } }
    # Un xifrat no va mai al Word: si te contrasenya d'obertura, el Word la
    # demanaria i es quedaria esperant (amagat) per sempre.
    if ($err -match 'xifrat') { return @{ Text = ''; Via = ''; Error = 'PDF xifrat'; PendentWord = $false } }
    if ($null -eq $word) { return @{ Text = ''; Via = ''; Error = $(if ($err -ne '') { $err } else { 'sense text (escanejat?)' }); PendentWord = $true } }
    if ($null -ne $word) {
        if (-not [bool]$word.Provat) { $word.Provat = $true; try { $word.App = New-WordApp -Opcional } catch { $word.App = $null } }
        if ($null -ne $word.App) {
            $tw = _PdfTextWord $word.App $path
            if (_PdfTeText $tw) { return @{ Text = $tw; Via = 'word'; Error = ''; PendentWord = $false } }
        }
    }
    return @{ Text = ''; Via = ''; Error = $(if ($err -ne '') { $err } else { 'sense text (escanejat?)' }); PendentWord = $false }
}

# Llegeix UN document del tipus que diu el nom. Torna @{ Extret; Llegible; Via;
# Error } (Extret $null si no se n'ha tret res).
function _CtLlegeixDocument($file, [string]$tipus, $word) {
    $r = @{ Extret = $null; Llegible = $true; Via = ''; Error = ''; PendentWord = $false }
    try {
        if ($tipus -eq 'etram' -or $tipus -eq 'xmlantic') {
            $txt = _CtTextXml $file.FullName
            $r.Via = 'xml'
            $r.Extret = if ($tipus -eq 'etram') { Read-EtramXml $txt } else { Read-XmlAntic $txt }
            if ($null -eq $r.Extret) { $r.Error = 'XML no valid' }
            return $r
        }
        $p = _CtTextPdf $file.FullName $word
        $r.Via = $p.Via
        $r.PendentWord = [bool]$p.PendentWord
        if ([string]$p.Text -eq '') { $r.Llegible = $false; $r.Error = [string]$p.Error; return $r }
        $r.Extret = if ($tipus -eq 'autoritzacio') { Read-AutoritzacioText $p.Text } else { Read-InstanciaText $p.Text }
        # Una "autoritzacio" que no diu qui autoritza a qui: potser es una
        # instancia anomenada aixi (i al reves).
        if ($null -eq $r.Extret -and $tipus -eq 'autoritzacio') { $r.Extret = Read-InstanciaText $p.Text }
        if ($null -ne $r.Extret -and $tipus -eq 'instancia' -and $null -eq $r.Extret.interessat) {
            $a = Read-AutoritzacioText $p.Text
            if ($null -ne $a) { $r.Extret = $a }
        }
    } catch { $r.Error = [string]$_.Exception.Message; $r.Extret = $null }
    return $r
}

# ----------------------------------------------------------------------------
# QUINS DOCUMENTS: tot l'arbre de cada carpeta d'activitat
# ----------------------------------------------------------------------------
# L'usuari (octubre 2026): "no es llegeixen les subcarpetes" (Firmadoc\, OME\,
# i sobretot la carpeta d'expedient del titular DINS de la d'una queixa). Es
# recorre tot l'arbre de cada carpeta d'activitat (no l'arrel), i es tria PEL
# NOM abans d'obrir res (_CtTipusDocument): dels ~25.000 fitxers de les
# subcarpetes nomes se'n queden uns centenars. Torna @{ F; Tipus } per cada un.
function Get-FitxersContactes([string]$dir) {
    $out = New-Object System.Collections.ArrayList
    foreach ($d in @(Get-ChildItem -LiteralPath $dir -Directory -ErrorAction SilentlyContinue)) {
        foreach ($f in @(Get-ChildItem -LiteralPath $d.FullName -File -Recurse -ErrorAction SilentlyContinue)) {
            $t = _CtTipusDocument $f.Name
            if ($t -ne '') { [void]$out.Add(@{ F = $f; Tipus = $t }) }
        }
    }
    return ,$out.ToArray()
}

# ----------------------------------------------------------------------------
# DE QUIN GIA ES UN DOCUMENT: el de LA CARPETA MES PROPERA que ho digui
# ----------------------------------------------------------------------------
# Carpeta (ruta relativa) -> els GIA dels informes que hi ha (de
# Get-InformeData o de _FlattenInformesDb: Ruta, Gia). PURA.
function _CtGiaPerCarpeta($informes, [string]$arrel) {
    $m = @{}
    foreach ($r in @($informes)) {
        if ($null -eq $r) { continue }
        $g = [string](_CtV $r 'Gia')
        if ($g -eq '') { continue }
        $c = @(_CtCarpetesDe (_ClauInforme ([string](_CtV $r 'Ruta')) $arrel))
        if ($c.Count -eq 0) { continue }
        $k = [string]$c[0].Rel
        if (-not $m.ContainsKey($k)) { $m[$k] = @{} }
        $m[$k][$g] = $true
    }
    return $m
}

# Un expedient -> el GIA de l'Excel. El document l'escriu "2025/3165/16" i la
# carpeta "2025-16-3165": es prova en tots dos ordres. PURA.
function _CtGiaDeExpedient([string]$exp, $expToGia) {
    if ($null -eq $expToGia) { return '' }
    $n = _NormalitzaExpedient $exp
    if ($n -eq '') { return '' }
    if ($expToGia.ContainsKey($n)) { return [string]$expToGia[$n] }
    $p = $n.Split('-')
    if ($p.Count -eq 3) {
        $g = $p[0] + '-' + $p[2] + '-' + $p[1]
        if ($expToGia.ContainsKey($g)) { return [string]$expToGia[$g] }
    }
    return ''
}

# El GIA d'un document. L'usuari (octubre 2026): "el GIA de les esmenes i les
# instancies l'escriu el ciutada, i sovint malament" (carpeta GIA 595, esmena
# "GIA 5"). L'ordre:
#   1. la carpeta MES PROPERA que digui de quina activitat es: el "GIA n" del
#      nom, els informes que hi ha (si tots son del mateix GIA) o el seu
#      expedient a l'Excel; i si no, la de sobre, fins a la d'activitat
#   2. el GIA del document, NOMES si cap carpeta no en diu cap
#   3. l'expedient del document -> Excel
# Si l'expedient del document es el de la carpeta, mana la carpeta (pas 1).
# $conegut: scriptblock (gia) -> bool. PURA.
function _CtGiaDocument([string]$clau, $extret, $perCarpeta, $conegut, $expToGia) {
    foreach ($c in @(_CtCarpetesDe $clau)) {
        $m = [regex]::Match([string]$c.Nom, '(?<![a-z])gia\s*(\d{1,6})(?!\d)', 'IgnoreCase')
        if ($m.Success) { return [string][int]$m.Groups[1].Value }
        if ($null -ne $perCarpeta -and $perCarpeta.ContainsKey($c.Rel) -and $perCarpeta[$c.Rel].Count -eq 1) { return [string]@($perCarpeta[$c.Rel].Keys)[0] }
        $g = _CtGiaDeExpedient (_CtExpedientDeCarpeta $c.Nom) $expToGia
        if ($g -ne '') { return $g }
    }
    $propi = [string](_CtV $extret 'gia')
    if ($propi -ne '' -and (& $conegut $propi)) { return $propi }
    return (_CtGiaDeExpedient ([string](_CtV $extret 'expedient')) $expToGia)
}

# ----------------------------------------------------------------------------
# TOTES LES ACTIVITATS (les regles sobre tota la base). PURA.
# ----------------------------------------------------------------------------
# $documents: clau -> @{ gia; modificat; extret; ... } (el que es desa).
# $excel: gia -> fila de contactes de l'Excel (o @{}). $correccions: gia ->
# llista. Torna gia -> activitat, amb la fila de l'Excel ('excel') al costat
# (la finestra la ensenya i la fa servir per tornar a calcular).
function _CtDocsPerGia($documents) {
    $per = @{}
    foreach ($k in @($documents.Keys)) {
        $d = $documents[$k]
        $g = [string](_CtV $d 'gia')
        if ($g -eq '' -or $null -eq (_CtV $d 'extret')) { continue }
        if (-not $per.ContainsKey($g)) { $per[$g] = New-Object System.Collections.ArrayList }
        [void]$per[$g].Add(@{ ruta = [string]$k; modificat = (Read-JsonIso (_CtV $d 'modificat')); extret = (_CtV $d 'extret') })
    }
    return $per
}

function _CtCalculaActivitats($documents, $excel, $tecnics, $recintes, $correccions) {
    $per = _CtDocsPerGia $documents
    if ($null -eq $excel) { $excel = @{} }
    $ctx = _CtContext $per $excel $tecnics $recintes
    $out = [ordered]@{}
    foreach ($g in @($per.Keys | Sort-Object { _GiaNumeric $_ })) {
        $xl = if ($excel.ContainsKey($g)) { $excel[$g] } else { $null }
        $corrs = if ($null -ne $correccions -and $correccions.ContainsKey($g)) { $correccions[$g] } else { $null }
        $man = if ($null -ne $corrs) { _CtManDeCorreccions $corrs } else { $null }
        $act = Get-ContactesActivitat $g @($per[$g]) $xl $ctx $man
        $act.excel = $xl
        $out[$g] = $act
    }
    return $out
}

# El que es resumeix al final d'"Actualitzar base": quantes activitats tenen
# dades noves (respecte de la passada anterior) i quants avisos de cada tipus.
function _CtResumAvisos($activitats) {
    $per = [ordered]@{ es_el_tecnic = 0; error = 0; falta = 0; diferent = 0; canvi_titular = 0 }
    $n = 0
    foreach ($a in @($activitats.Values)) {
        foreach ($v in @($a.avisos)) { $t = [string](_CtV $v 'tipus'); if ($per.Contains($t)) { $per[$t]++ }; $n++ }
    }
    return @{ Total = $n; PerTipus = $per }
}

# Una "empremta" del que diu una activitat (titular, representant,
# autoritzats, establiment) per saber si te dades noves. PURA.
function _CtEmpremta($a) {
    if ($null -eq $a) { return '' }
    $parts = @()
    foreach ($k in 'titular', 'representant_legal', 'establiment') {
        $p = _CtV $a $k
        $parts += (@('nom', 'nif', 'email', 'telefon', 'mobil', 'nom_comercial') | ForEach-Object { [string](_CtV $p $_) }) -join ','
    }
    $parts += (@(@(_CtV $a 'persones_autoritzades') | ForEach-Object { [string](_CtV $_ 'email') + '/' + (_CtNomNorm ([string](_CtV $_ 'nom'))) } | Sort-Object) -join ',')
    return ($parts -join '|')
}

# El text del resum (una linia, la del missatge del boto i del registre). PURA.
function _CtResumText($res) {
    $pt = $res.PerTipus
    $det = New-Object System.Collections.ArrayList
    $noms = [ordered]@{ es_el_tecnic = "del t$([char]0x00E8)cnic"; error = 'errors'; falta = 'falten'; diferent = 'diferents'; canvi_titular = 'canvis de titular' }
    foreach ($k in $noms.Keys) { if ([int]$pt[$k] -gt 0) { [void]$det.Add(([string]$pt[$k] + ' ' + $noms[$k])) } }
    $t = 'Contactes: ' + $res.NNoves + ' activitats amb dades noves ' + [char]0x00B7 + ' ' + $res.NAvisos + ' avisos'
    if ($det.Count -gt 0) { $t += ' (' + ($det -join ', ') + ')' }
    $t += ' ' + [char]0x00B7 + ' ' + $res.NNoLlegibles + ' documents no llegibles'
    # Sense la llista, el tecnic que posa el correu del client no es veu
    # (l'usuari: a la feina no hi era i "el programa ho hauria de dir").
    # (Sense el simbol: les finestres el pinten com un quadre.)
    if ([string]$res.TecnicsFitxer -eq '') { $t += ' ' + [char]0x00B7 + ' AV' + [char]0x00CD + "S: no hi ha llista de t" + [char]0x00E8 + "cnics coneguts" }
    return $t
}

# Els recomptes de la validacio (ValidarContactes.ps1): documents per tipus i
# quants son de subcarpetes, noms que son etiquetes (ha de ser 0), persones
# autoritzades d'una entitat de control (0) i representants legals que surten
# de la signatura (R:CIF). $docs: clau -> document desat. PURA.
function _CtEstadistiques($acts, $docs) {
    $r = [ordered]@{ PerTipus = [ordered]@{}; NSubcarpetes = 0; NomsEtiqueta = 0; AutEca = 0; RepSignatura = 0 }
    foreach ($k in @($docs.Keys)) {
        $t = [string](_CtV $docs[$k] 'tipus')
        if (-not $r.PerTipus.Contains($t)) { $r.PerTipus[$t] = 0 }
        $r.PerTipus[$t]++
        if (@(_CtCarpetesDe ([string]$k)).Count -gt 1) { $r.NSubcarpetes++ }
    }
    foreach ($a in @($acts.Values)) {
        $ps = @((_CtV $a 'titular'), (_CtV $a 'representant_legal')) + @(_CtV $a 'persones_autoritzades')
        foreach ($p in $ps) { if ($null -ne $p -and (_CtNomEsEtiqueta ([string](_CtV $p 'nom')))) { $r.NomsEtiqueta++ } }
        foreach ($p in @(_CtV $a 'persones_autoritzades')) { if ($null -ne $p -and (_CtEmailExclos ([string](_CtV $p 'email')))) { $r.AutEca++ } }
        if ([string](_CtV (_CtV $a 'representant_legal') 'via') -eq 'signatura') { $r.RepSignatura++ }
    }
    return $r
}

# ----------------------------------------------------------------------------
# LA BASE AL DISC
# ----------------------------------------------------------------------------
function Read-ContactesDb([string]$path = '') {
    if ($path -eq '') { $path = Get-ContactesDbPath }
    if ($path -eq '' -or -not (Test-Path -LiteralPath $path)) { return $null }
    return (Read-JsonFile $path)
}

# Un objecte del JSON (documents, activitats, correccions) -> hashtable.
function _CtMapaDe($o) {
    if ($null -eq $o) { return @{} }
    return (ConvertTo-Mapa $o)
}

# Les correccions de la base, per GIA (llistes).
function _CtCorreccionsDe($db) {
    $m = @{}
    $c = _CtMapaDe (_CtV $db 'correccions')
    foreach ($k in @($c.Keys)) { $m[[string]$k] = @(@($c[$k]) | Where-Object { $null -ne $_ }) }
    return $m
}

# ----------------------------------------------------------------------------
# EL REPAS (el crida Invoke-InformesDbEscaneig, al final)
# ----------------------------------------------------------------------------
# $dir: la carpeta d'informes; $informes: els registres de l'escaneig (Ruta,
# Gia), per l'atribucio; $cache: la de l'Excel (Contactes); $avisa: el
# progres. -SenseCorreccions (ValidarContactes): sense les correccions a ma.
# -NoDesis: no escriu la base (torna el resultat). Torna @{ Ok; Error; NNoves;
# NAvisos; PerTipus; NNoLlegibles; NDocs; Llegits; Text; OutPath; Activitats;
# NoLlegibles; SenseGia }.
# -SenseWord (la passada AUTOMATICA, que no te ningu davant): cap PDF va al Word.
# Si el Word mostra una pregunta, un proces sense finestres s'hi queda esperant
# per sempre (amb el mutex de la base agafat). Els que l'haurien necessitat es
# marquen 'pendent_word' i els torna a llegir el proxim "Actualitzar base" a ma.
function Invoke-ContactesEscaneig([string]$dir, $informes, $cache, $expToGia, [scriptblock]$avisa = $null, [switch]$SenseCorreccions, [switch]$NoDesis, [switch]$SenseWord) {
    $diu = { param($t, $i, $n) if ($null -ne $avisa) { & $avisa $t $i $n } }
    $outPath = Get-ContactesDbPath
    $prev = $null
    try { $prev = Read-ContactesDb $outPath } catch { $prev = $null }
    $prevDocs = @{}; $prevActs = @{}; $corrs = @{}
    if ($null -ne $prev) {
        $corrs = _CtCorreccionsDe $prev
        $prevActs = _CtMapaDe (_CtV $prev 'activitats')
        if ([string](_CtV $prev 'versio') -eq $Script:ContactesVersio) { $prevDocs = _CtMapaDe (_CtV $prev 'documents') }
    }
    if ($SenseCorreccions) { $corrs = @{} }
    $tecnics = Read-TecnicsConeguts
    $cfg = Read-ContactesConfig

    & $diu "Contactes: cercant els documents de les activitats..." 0 0
    $cands = Get-FitxersContactes $dir

    $perCarpeta = _CtGiaPerCarpeta $informes $dir
    $gies = @{}
    foreach ($c in @($perCarpeta.Values)) { foreach ($g in @($c.Keys)) { $gies[[string]$g] = $true } }
    $xlCt = @{}
    if ($null -ne $cache -and $null -ne $cache.PSObject.Properties['Contactes'] -and $null -ne $cache.Contactes) { $xlCt = $cache.Contactes }
    foreach ($g in @($xlCt.Keys)) { $gies[[string]$g] = $true }
    $conegut = { param($g) $gies.ContainsKey([string]$g) }.GetNewClosure()

    $docs = @{}
    $noLlegibles = New-Object System.Collections.ArrayList
    $senseGia = New-Object System.Collections.ArrayList
    $llegits = 0; $i = 0; $total = $cands.Count
    $word = @{ App = $null; Provat = $false }
    try {
        foreach ($c in $cands) {
            $i++
            $f = $c.F
            $clau = _ClauInforme $f.FullName $dir
            $mod = $f.LastWriteTimeUtc.ToString('o')
            $p = if ($prevDocs.ContainsKey($clau)) { $prevDocs[$clau] } else { $null }
            # Read-JsonIso: el pwsh 7 torna les dates ISO del JSON com a
            # [datetime], i en text ja no serien iguals (mesurat a les proves).
            $reprova = (-not $SenseWord) -and [bool](_CtV $p 'pendent_word')
            if ($null -ne $p -and (Read-JsonIso (_CtV $p 'modificat')) -eq $mod -and -not $reprova) {
                $e = [ordered]@{ modificat = $mod; tipus = $c.Tipus; via = [string](_CtV $p 'via'); llegible = [bool](_CtV $p 'llegible'); error = [string](_CtV $p 'error'); extret = (_CtV $p 'extret'); gia = ''; pendent_word = [bool](_CtV $p 'pendent_word') }
            } else {
                $l = _CtLlegeixDocument $f $c.Tipus $(if ($SenseWord) { $null } else { $word })
                $llegits++
                $e = [ordered]@{ modificat = $mod; tipus = $c.Tipus; via = [string]$l.Via; llegible = [bool]$l.Llegible; error = [string]$l.Error; extret = $l.Extret; gia = ''; pendent_word = [bool]$l.PendentWord }
            }
            # L'atribucio es refa sempre: depen dels informes d'ara.
            $e.gia = _CtGiaDocument $clau $e.extret $perCarpeta $conegut $expToGia
            if (-not $e.llegible) { [void]$noLlegibles.Add($clau) }
            elseif ($e.gia -eq '' -and $null -ne $e.extret) { [void]$senseGia.Add($clau) }
            $docs[$clau] = $e
            if (($i % 10) -eq 0 -or $i -eq $total) { & $diu "Contactes: llegint documents... ($i de $total, $llegits de nous/modificats)" $i $total }
        }
    } finally {
        if ($null -ne $word.App) { try { $word.App.Quit() } catch { } }
    }

    $acts = _CtCalculaActivitats $docs $xlCt $tecnics $cfg.Recintes $corrs
    $nNoves = 0
    foreach ($g in @($acts.Keys)) {
        $ant = if ($prevActs.ContainsKey($g)) { $prevActs[$g] } else { $null }
        if ((_CtEmpremta $acts[$g]) -ne (_CtEmpremta $ant)) { $nNoves++ }
    }
    $ra = _CtResumAvisos $acts
    $res = @{
        Ok = $true; Error = ''; NNoves = $nNoves; NAvisos = $ra.Total; PerTipus = $ra.PerTipus
        NNoLlegibles = $noLlegibles.Count; NDocs = $docs.Count; Llegits = $llegits; OutPath = $outPath
        Activitats = $acts; NoLlegibles = $noLlegibles.ToArray(); SenseGia = $senseGia.ToArray()
        TecnicsFitxer = [string]$tecnics.Fitxer; TecnicsError = [string]$tecnics.Error
        Estadistiques = (_CtEstadistiques $acts $docs)
    }
    $res.Text = _CtResumText $res
    if (-not $NoDesis -and $outPath -ne '') {
        $db = [ordered]@{
            versio         = $Script:ContactesVersio
            actualitzat_el = (Get-Date).ToString('o')
            carpeta_arrel  = $dir
            tecnics_fitxer = [string]$tecnics.Fitxer
            n_activitats   = $acts.Count
            activitats     = $acts
            correccions    = $(if ($null -ne $prev) { _CtMapaDe (_CtV $prev 'correccions') } else { @{} })
            no_llegibles   = $noLlegibles.ToArray()
            sense_gia      = $senseGia.ToArray()
            documents      = $docs
        }
        Write-JsonFile $outPath $db 14
        _CtOblidaMemo
    }
    return $res
}

# ----------------------------------------------------------------------------
# LES CORRECCIONS A MA (la finestra Contactes)
# ----------------------------------------------------------------------------
# Torna a calcular les activitats de la base amb les correccions d'ara (sense
# tornar a llegir cap document: els documents i la fila de l'Excel ja hi son).
function Update-ContactesActivitats($db) {
    $docs = _CtMapaDe (_CtV $db 'documents')
    $actsAnt = _CtMapaDe (_CtV $db 'activitats')
    $xl = @{}
    foreach ($g in @($actsAnt.Keys)) { $x = _CtV $actsAnt[$g] 'excel'; if ($null -ne $x) { $xl[[string]$g] = (ConvertTo-Mapa $x) } }
    $cfg = Read-ContactesConfig
    $acts = _CtCalculaActivitats $docs $xl (Read-TecnicsConeguts) $cfg.Recintes (_CtCorreccionsDe $db)
    $db.activitats = $acts
    $db.n_activitats = $acts.Count
    return $db
}

# Afegeix (o treu, amb -Treu i l'index) una correccio d'un GIA i desa. Amb el
# mutex de la base: l'automatic pot estar escrivint. Torna $true si s'ha desat.
function Set-ContactesCorreccio([string]$gia, $correccio, [int]$treuIndex = -1) {
    $caixa = @{ Ok = $false }
    $fet = Invoke-AmbMutexUnic $Script:BaseMutexNom {
        $path = Get-ContactesDbPath
        $o = Read-ContactesDb $path
        if ($null -eq $o) { return }
        $db = ConvertTo-Mapa $o
        $c = _CtMapaDe $db['correccions']
        $llista = New-Object System.Collections.ArrayList
        if ($c.ContainsKey($gia)) { foreach ($x in @($c[$gia])) { if ($null -ne $x) { [void]$llista.Add($x) } } }
        if ($treuIndex -ge 0) { if ($treuIndex -lt $llista.Count) { $llista.RemoveAt($treuIndex) } }
        else {
            $correccio['quan'] = (Get-Date).ToString('o')
            # Una mateixa correccio no es repeteix: la nova substitueix la vella.
            $clau = _CtClauCorreccio $correccio
            for ($k = $llista.Count - 1; $k -ge 0; $k--) { if ((_CtClauCorreccio $llista[$k]) -eq $clau) { $llista.RemoveAt($k) } }
            [void]$llista.Add($correccio)
        }
        if ($llista.Count -gt 0) { $c[$gia] = $llista.ToArray() } else { [void]$c.Remove($gia) }
        $db['correccions'] = $c
        $db = Update-ContactesActivitats ([pscustomobject]$db)
        Write-JsonFile $path $db 14
        _CtOblidaMemo
        $caixa.Ok = $true
    }
    return ($fet -and $caixa.Ok)
}

# Dues correccions del mateix (la mateixa persona, el mateix avis, el mateix
# camp) son la mateixa. PURA.
function _CtClauCorreccio($c) {
    $t = [string](_CtV $c 'tipus')
    switch ($t) {
        'descarta' { return 'descarta|' + [string](_CtV $c 'avis') }
        'edita'    { return 'edita|' + [string](_CtV $c 'camp') }
        default    { return $t + '|' + ((@(_CtV $c 'claus') | Sort-Object) -join ',') }
    }
}

# ----------------------------------------------------------------------------
# FER-LA SERVIR ALS CORREUS (Enviar correu, recordatoris, controls)
# ----------------------------------------------------------------------------
# La base es llegeix UN cop i es guarda mentre el fitxer no canvii (els
# recordatoris en demanen per cada fila).
$Script:CtMemo = @{ Path = ''; Stamp = ''; Db = $null }
function _CtOblidaMemo { $Script:CtMemo.Path = ''; $Script:CtMemo.Stamp = ''; $Script:CtMemo.Db = $null }
function _CtDbMemo {
    $p = Get-ContactesDbPath
    if ($p -eq '' -or -not (Test-Path -LiteralPath $p)) { return $null }
    $st = ''
    try { $st = (Get-Item -LiteralPath $p).LastWriteTimeUtc.Ticks.ToString() } catch { }
    if ($Script:CtMemo.Path -eq $p -and $Script:CtMemo.Stamp -eq $st) { return $Script:CtMemo.Db }
    $db = $null
    try { $db = Read-JsonFile $p } catch { $db = $null }
    $Script:CtMemo.Path = $p; $Script:CtMemo.Stamp = $st; $Script:CtMemo.Db = $db
    return $db
}

function _CtActivitatDe([string]$gia) {
    $db = _CtDbMemo
    if ($null -eq $db -or [string]::IsNullOrWhiteSpace($gia)) { return $null }
    return (_CtV (_CtV $db 'activitats') ([string]$gia))
}

# Els correus de les persones autoritzades (Configuracio -> Correus de cada
# eina -> "Persones autoritzades / tecnic"). Llista plana.
function Get-ContactesAutoritzatsEmails([string]$gia) {
    $a = _CtActivitatDe $gia
    if ($null -eq $a) { return @() }
    return @(@(_CtV $a 'persones_autoritzades') | ForEach-Object { _CtEmailNet ([string](_CtV $_ 'email')) } | Where-Object { $_ -ne '' } | Select-Object -Unique)
}

# Completa els correus de l'Excel amb els dels documents. $emails: @{ titular;
# representant } (els de l'Excel). Torna @{ Emails = @{ titular; representant };
# Notes = textos per ensenyar ("surt dels documents"); AvisTecnic = text o '' (el
# correu de l'Excel es del tecnic: s'ha d'avisar ABANS d'enviar). PURA sobre
# l'activitat ($act, la de la base).
function _CtCompletaEmails($act, $emails) {
    $r = @{ Emails = @{ titular = [string]$emails.titular; representant = [string]$emails.representant }; Notes = @(); AvisTecnic = '' }
    if ($null -eq $act) { return $r }
    $notes = New-Object System.Collections.ArrayList
    foreach ($q in @(@{ K = 'titular'; Qui = 'titular'; Nom = 'del titular' }, @{ K = 'representant'; Qui = 'representant_legal'; Nom = 'del representant legal' })) {
        if ([string]$r.Emails[$q.K] -ne '') { continue }
        $p = _CtV $act $q.Qui
        $e = _CtEmailNet ([string](_CtV $p 'email'))
        if ($e -eq '') { continue }
        $r.Emails[$q.K] = $e
        [void]$notes.Add(("L'Excel no t" + [char]0x00E9 + " el correu " + $q.Nom + ": " + $e + " surt dels documents (" + [string](_CtV $p 'font') + ")."))
    }
    $r.Notes = $notes.ToArray()
    $avisos = New-Object System.Collections.ArrayList
    foreach ($v in @(_CtV $act 'avisos')) {
        if ([string](_CtV $v 'tipus') -ne 'es_el_tecnic') { continue }
        $val = _CtEmailNet ([string](_CtV $v 'valor_excel'))
        if ($val -eq '') { continue }
        if ($val -eq (_CtEmailNet $emails.titular) -or $val -eq (_CtEmailNet $emails.representant)) {
            [void]$avisos.Add(($val + ' (' + [string](_CtV $v 'camp') + '): ' + [string](_CtV $v 'problema')))
        }
    }
    if ($avisos.Count -gt 0) { $r.AvisTecnic = ("Segons els documents, aquest correu de l'Excel " + [char]0x00E9 + "s del t" + [char]0x00E8 + "cnic o la gestoria, no del titular:`n`n" + ($avisos -join "`n")) }
    return $r
}

function Get-ContactesEmailsCompletats([string]$gia, $emails) {
    $act = $null
    try { $act = _CtActivitatDe $gia } catch { $act = $null }
    return (_CtCompletaEmails $act $emails)
}

# ----------------------------------------------------------------------------
# LES CORRECCIONS PER ENTRAR AL GIA (Exportar a Excel)
# ----------------------------------------------------------------------------
$Script:CtExportCapcalera = @('GIA', 'Titular', 'Camp', "Valor a l'Excel", 'Problema', 'Proposta', 'Font', 'Tipus')
$Script:CtExportAmples = @(8, 32, 20, 32, 60, 32, 50, 14)

# Les files: un avis per fila, per GIA. PURA.
function _CtFilesCorreccions($activitats) {
    $out = New-Object System.Collections.ArrayList
    foreach ($g in @($activitats.Keys | Sort-Object { _GiaNumeric $_ })) {
        $a = $activitats[$g]
        $tit = [string](_CtV (_CtV $a 'excel') 'TITULAR')
        if ($tit -eq '') { $tit = [string](_CtV (_CtV $a 'titular') 'nom') }
        foreach ($v in @(_CtV $a 'avisos')) {
            if ($null -eq $v) { continue }
            [void]$out.Add(@([string]$g, $tit, [string](_CtV $v 'camp'), [string](_CtV $v 'valor_excel'), [string](_CtV $v 'problema'), [string](_CtV $v 'proposta'), [string](_CtV $v 'font'), [string](_CtV $v 'tipus')))
        }
    }
    return $out.ToArray()
}

function Export-ContactesCorreccions([string]$desti, $activitats) {
    $bytes = _NormativaXlsxBytes $Script:CtExportCapcalera @(_CtFilesCorreccions $activitats) $Script:CtExportAmples 'Correccions'
    [System.IO.File]::WriteAllBytes($desti, $bytes)
    return $desti
}

# ----------------------------------------------------------------------------
# LA VALIDACIO (ValidarContactes.ps1): el repas contra la referencia feta a ma
# ----------------------------------------------------------------------------
# El nom d'un camp de l'Excel, sigui com sigui escrit ("Rep. Leg. E-mail",
# "Representant legal", "Rao soc. Mobil"...), a una clau: nom, nif, email,
# telefon, mobil, amb 'rep_' al davant si es del representant. PURA.
function _CtCampCanonic([string]$camp) {
    $p = _CtPla $camp
    $rep = ($p -match '\brep\b|rep\.|representant')
    $k = if ($p -match 'mail|correu') { 'email' } elseif ($p -match 'mobil') { 'mobil' } elseif ($p -match 'telef') { 'telefon' } elseif ($p -match '\b(nif|dni|cif|nie)\b') { 'nif' } else { 'nom' }
    if ($rep) { return 'rep_' + $k }
    return $k
}

# Les discrepancies d'UNA activitat ($ref: la de la referencia; $ara: la del
# repas). Torna @{ Que; Ref; Ara; Confianca } per cada una. Es compara (l'usuari):
# l'e-mail i els telefons del titular, el nom del representant legal (sense
# accents ni majuscules), el conjunt d'e-mails de les persones autoritzades i els
# avisos es_el_tecnic. Els telefons, com a CONJUNT: la referencia els pot tenir
# a "telefon" o a "mobil" tal com els va escriure el document, i el repas els
# reparteix pel prefix. PURA.
function _CtComparaActivitat($ref, $ara) {
    $out = New-Object System.Collections.ArrayList
    $conf = { param($p) $c = [string](_CtV $p 'confianca'); if ($c -eq '') { 'segur' } else { $c } }
    $tR = _CtV $ref 'titular'; $tA = _CtV $ara 'titular'
    $eR = _CtEmailNet ([string](_CtV $tR 'email')); $eA = _CtEmailNet ([string](_CtV $tA 'email'))
    if ($eR -ne $eA) { [void]$out.Add(@{ Que = 'e-mail del titular'; Ref = $eR; Ara = $eA; Confianca = (& $conf $tR) }) }
    $telsR = @(@('telefon', 'mobil') | ForEach-Object { _CtTelefonNet ([string](_CtV $tR $_)) } | Where-Object { $_ -ne '' } | Sort-Object -Unique)
    $telsA = @(@('telefon', 'mobil') | ForEach-Object { _CtTelefonNet ([string](_CtV $tA $_)) } | Where-Object { $_ -ne '' } | Sort-Object -Unique)
    if (($telsR -join ',') -ne ($telsA -join ',')) { [void]$out.Add(@{ Que = "tel$([char]0x00E8)fons del titular"; Ref = ($telsR -join ', '); Ara = ($telsA -join ', '); Confianca = (& $conf $tR) }) }
    $rR = _CtV $ref 'representant_legal'; $rA = _CtV $ara 'representant_legal'
    $nR = [string](_CtV $rR 'nom'); $nA = [string](_CtV $rA 'nom')
    $igual = ((_CtNomNorm $nR) -eq (_CtNomNorm $nA)) -or (_CtMateixNom $nR $nA)
    if (-not $igual) { [void]$out.Add(@{ Que = 'representant legal'; Ref = $nR; Ara = $nA; Confianca = (& $conf $rR) }) }
    $autR = @(@(_CtV $ref 'persones_autoritzades') | Where-Object { $null -ne $_ })
    $emR = @($autR | ForEach-Object { _CtEmailNet ([string](_CtV $_ 'email')) } | Where-Object { $_ -ne '' } | Sort-Object -Unique)
    $emA = @(@(_CtV $ara 'persones_autoritzades') | ForEach-Object { _CtEmailNet ([string](_CtV $_ 'email')) } | Where-Object { $_ -ne '' } | Sort-Object -Unique)
    if (($emR -join ',') -ne ($emA -join ',')) {
        # "probable" si alguna de les que no quadren ho es a la referencia.
        $falten = @($emR | Where-Object { $emA -notcontains $_ })
        $c = 'segur'
        foreach ($p in $autR) { if ($falten -contains (_CtEmailNet ([string](_CtV $p 'email'))) -and (& $conf $p) -eq 'probable') { $c = 'probable' } }
        if ($falten.Count -eq 0 -and @($autR | Where-Object { (& $conf $_) -eq 'probable' }).Count -gt 0) { $c = 'probable' }
        [void]$out.Add(@{ Que = 'e-mails de les persones autoritzades'; Ref = ($emR -join ', '); Ara = ($emA -join ', '); Confianca = $c })
    }
    $tecR = @(@(_CtV $ref 'verificacio_excel') | Where-Object { $null -ne $_ -and ((_CtPla ([string](_CtV $_ 'problema') + ' ' + [string](_CtV $_ 'tipus'))) -match 'tecnic') } | ForEach-Object { _CtCampCanonic ([string](_CtV $_ 'camp')) } | Sort-Object -Unique)
    $tecA = @(@(_CtV $ara 'avisos') | Where-Object { [string](_CtV $_ 'tipus') -eq 'es_el_tecnic' } | ForEach-Object { _CtCampCanonic ([string](_CtV $_ 'camp')) } | Sort-Object -Unique)
    if (($tecR -join ',') -ne ($tecA -join ',')) { [void]$out.Add(@{ Que = "avisos es_el_tecnic (camps de l'Excel)"; Ref = ($tecR -join ', '); Ara = ($tecA -join ', '); Confianca = 'segur' }) }
    return $out.ToArray()
}
