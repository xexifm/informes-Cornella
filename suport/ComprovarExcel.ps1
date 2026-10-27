#requires -Version 5.1
<#
.SYNOPSIS
  Eina "Excel Comprovar" (abans "Comprovar Excel"; menu, fila GIA): per cada activitat en estat
  "Precinte / Cessament" a la base d'informes, comprova que l'Excel d'activitats
  hi tingui el Camp Info corresponent amb un valor que comenci per SI.
  Criteri DIFERENT del de Seguiment (SeguimentGia.ps1), a posta: vegeu CLAUDE.md.
  Abans vivia a Informes.ps1 (revisio d'arquitectura, setembre 2026).
#>

# ----------------------------------------------------------------------------
# Excel Comprovar: noms de "Camp Info" que, a l'Excel d'activitats, marquen que
# una activitat esta requerida per decret / precintada. La comprovacio es fa
# sobre el NOM del camp (normalitzat, sense accents/majuscules; el '?' es
# conserva) i, a mes, el VALOR ha de comencar per "SI".
#
# ATENCIO: han de ser els noms EXACTES de les columnes de l'Excel. Aqui hi deia
# 'precinte?' quan a l'Excel el camp es 'PRECINTE ACTIVITAT?', i el resultat era
# que activitats correctament marcades sortien com a DESACTUALITZADES. La
# comparacio es exacta a proposit (un 'comenca per precinte' agafaria coses com
# 'PRECINTE AIXECAT?', que voldria dir justament el contrari). Per aixo el
# llistat del final del resum diu QUINS camps d'aquesta familia hi ha de debo a
# l'Excel: si algun dia es tornen a reanomenar, es veu de seguida.
$Script:ExcelPrecinteCampNoms = @('requerit per decret?', 'precinte activitat?')

# Etiqueta per als missatges: els noms de dalt, en majuscules i entre cometes.
function _ExcelPrecinteCampsText {
    return (($Script:ExcelPrecinteCampNoms | ForEach-Object { "'" + $_.ToUpperInvariant() + "'" }) -join ' o ')
}

# Camps "Camp Info" de l'Excel que parlen de precinte o de decret, siguin els
# que busquem o no. Serveix per DIAGNOSTICAR un canvi de nom: si el camp bo
# passa a dir-se d'una altra manera, aqui es veura. Funcio PURA.
#   $map = la taula GIA -> llista de @{ Nom; Valor } de _ReadExcelCampInfoPerGia.
function _ExcelCampsPrecinteTrobats($map) {
    $vistos = New-Object System.Collections.ArrayList
    if ($null -eq $map) { return $vistos.ToArray() }
    foreach ($gia in $map.Keys) {
        foreach ($p in @($map[$gia])) {
            $nom = [string]$p.Nom
            $n = _NormalitzaText $nom
            if ($n -notmatch 'precinte|decret') { continue }
            if (-not $vistos.Contains($nom)) { [void]$vistos.Add($nom) }
        }
    }
    return (@($vistos.ToArray()) | Sort-Object)
}

# PURA i testejable. $pairs = llista de @{ Nom; Valor } (els Camp Info d'una fila
# de l'Excel). Retorna $true si algun te el Nom entre els objectius I el Valor
# comenca per "SI". _NormalitzaText es de GenerarInforme.ps1 (ja dot-sourcejat).
function _ExcelActivitatActualitzada($pairs) {
    if ($null -eq $pairs) { return $false }
    foreach ($p in $pairs) {
        $nom = _NormalitzaText ([string]$p.Nom)
        if ($Script:ExcelPrecinteCampNoms -contains $nom) {
            $val = _NormalitzaText ([string]$p.Valor)
            if ($val -match '^si\b') { return $true }
        }
    }
    return $false
}

# ============================================================================
# Excel Comprovar — eina INFORMES "Excel Comprovar"
# ============================================================================
# Llegeix els "Camp Info N - Nom/Valor" de la fulla "Estès" i els indexa per GIA
# (columna 1). Retorna @{ Ok; Map = @{ gia -> @(@{Nom;Valor}) }; Error }.
function _ReadExcelCampInfoPerGia {
    $latest = Find-LatestActivitatsExcel
    if ($null -eq $latest) { return @{ Ok = $false; Error = "No s'ha trobat cap Excel d'activitats." } }
    # Read-FullaEstesa LLANCA quan l'Excel no arrenca o la fulla no hi es; aqui
    # el crider espera un @{ Ok = $false; Error }, o sigui que s'hi posa un catch.
    try {
        return (Read-FullaEstesa $latest.File {
            param($x)
            $data = $x.Data; $rows = $x.Rows; $cols = $x.Cols; $headers = $x.Headers
            if ($null -eq $data) { return @{ Ok = $true; Map = @{}; DataControl = @{}; ColControl = $false; DataInspeccio = @{}; ColInspeccio = $false } }
            $pairs = _FindCampInfoPairs $headers
            # La data del control (Excel Comprovar mira que hi consti la de la
            # visita d'inspeccio de l'informe). 0 si la columna no hi es.
            $colCtl = _FindColIndex $data $cols @('data control inicial') $null
            # La de les QUEIXES: 'Inspeccio', EXACTA (n'hi ha d'altres que porten
            # la paraula). 0 si no hi es.
            $colIns = 0
            $hdr = @($headers)
            for ($ci = 0; $ci -lt $hdr.Count; $ci++) { if ((_NormalitzaText ([string]$hdr[$ci])).Trim() -eq 'inspeccio') { $colIns = $ci + 1; break } }
            $map = @{}
            $ctl = @{}
            $ins = @{}
            for ($r = 2; $r -le $rows; $r++) {
                $cell = $data[$r, 1]
                if ($null -eq $cell) { continue }
                $gia = if ($cell -is [double]) {
                    if ([math]::Floor($cell) -eq $cell) { [string][int]$cell } else { [string]$cell }
                } else { [string]$cell }
                $gia = $gia.Trim()
                if ($gia -eq '') { continue }
                $list = @()
                foreach ($p in $pairs) {
                    $nom = [string]$data[$r, $p.NomCol]
                    $val = [string]$data[$r, $p.ValorCol]
                    if (-not [string]::IsNullOrWhiteSpace($nom)) { $list += @{ Nom = $nom.Trim(); Valor = $val.Trim() } }
                }
                $map[$gia] = $list
                if ($colCtl -gt 0) { $ctl[$gia] = @(_DatesDeCella $data[$r, $colCtl]) }
                if ($colIns -gt 0) { $ins[$gia] = @(_DatesDeCella $data[$r, $colIns]) }
            }
            return @{ Ok = $true; Map = $map; DataControl = $ctl; ColControl = ($colCtl -gt 0); DataInspeccio = $ins; ColInspeccio = ($colIns -gt 0) }
        })
    } catch {
        return @{ Ok = $false; Error = $_.Exception.Message }
    }
}

# ----------------------------------------------------------------------------
# LES DATES D'UNA CEL·LA: un Excel les dona com a numero (data d'Excel), com a
# DateTime o com a text (amb una o mes dates). Torna 'dd/MM/yyyy'. PURA.
function _DatesDeCella($v) {
    if ($null -eq $v) { return @() }
    if ($v -is [double]) {
        try { return @([datetime]::FromOADate([double]$v).ToString('dd/MM/yyyy')) } catch { return @() }
    }
    if ($v -is [datetime]) { return @(([datetime]$v).ToString('dd/MM/yyyy')) }
    # L'ExportacioActivitats_*.xls porta el text "2025-07-25 00:00:00.0"
    # (octubre 2026): _DatesDelText nomes enten dd/mm/aaaa i no hi veia cap
    # data, o sigui que totes les visites sortien com a no apuntades.
    $iso = _DataDeCella $v
    if ($null -ne $iso -and ([string]$v).Trim() -match '^\d{4}-') { return @($iso.ToString('dd/MM/yyyy')) }
    return @(_DatesDelText ([string]$v))
}

# LES VISITES PER MOTIU (octubre 2026). L'usuari: la data d'una visita per una
# QUEIXA ha de constar a la columna 'Inspeccio' de l'Excel, i la d'una visita
# del PLA DE VERIFICACIO, a 'Data control inicial/verificacio'; si es per totes
# dues, a totes dues. El motiu surt de l'Objecte de l'informe
# (_ExtractVisitaMotiu, camp 'visita_motiu' de la base). Torna la mes recent de
# cada: @{ Queixa; Pla; Sense } (@{ Data; Fitxer; Dt } o $null). 'Sense' = un
# informe d'abans, que no diu el motiu. PURA.
function _CxVisitesPerMotiu($act) {
    $out = @{ Queixa = $null; Pla = $null; Sense = $null }
    foreach ($inf in @($act.informes)) {
        if ($null -eq $inf -or [bool](_PropInf $inf 'ignorat')) { continue }
        $v = [string](_PropInf $inf 'visita')
        if (-not $v) { continue }
        $dt = [datetime]::MinValue
        if (-not [datetime]::TryParseExact($v, 'dd/MM/yyyy', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) { continue }
        $m = _MotiuVisitaNet ([string](_PropInf $inf 'visita_motiu'))
        $vis = @{ Data = $v; Fitxer = [string]$inf.fitxer; Dt = $dt }
        $quins = @()
        if ($m -match 'queixa') { $quins += 'Queixa' }
        if ($m -match 'pla') { $quins += 'Pla' }
        if ($m -eq '') { $quins += 'Sense' }
        foreach ($q in $quins) { if ($null -eq $out[$q] -or $dt -gt $out[$q].Dt) { $out[$q] = $vis } }
    }
    return $out
}

# TOTS ELS AVISOS de la comprovacio, cadascun amb una CLAU que no canvia d'una
# passada a l'altra (per poder-lo DESCARTAR, vegeu mes avall). $db: la base
# d'informes; $res: el que torna _ReadExcelCampInfoPerGia. Torna una llista de
# @{ Clau; Grup; Text }. PURA.
#   precinte  activitat en Precinte / Cessament sense el Camp Info amb SI
#   notrobat  aquella activitat no es a l'Excel
#   sensegia  aquella activitat no te GIA a la base
#   visita       la darrera visita del PLA DE VERIFICACIO ("Objecte: Visita
#                inspeccio 24/07/2026 - Pla de verificacio 2026") no consta a
#                'Data control inicial/verificacio'
#   visitaqueixa la darrera visita per QUEIXA no consta a 'Inspeccio'
#   visitasense  un informe d'abans (sense motiu), si es la visita mes nova:
#                ha de constar a l'una o a l'altra columna
#   Nomes la DARRERA de cada: les columnes porten la data mes nova; les visites
#   d'abans no hi han de ser per forca. La clau porta la data: una visita nova
#   torna a avisar.
function Get-ComprovacioExcelAvisos($db, $res) {
    $out = New-Object System.Collections.ArrayList
    $map = $res.Map
    $ctl = $res.DataControl
    $ins = if ($null -ne $res.DataInspeccio) { $res.DataInspeccio } else { @{} }
    foreach ($act in @($db.activitats)) {
        if ($null -eq $act) { continue }
        $gia = [string]$act.id_gia
        $etiq = "GIA " + $gia + "  -  " + [string]$act.titular
        if ([string]$act.estat_actual -eq $Script:EstatPrecinte) {
            if ([string]::IsNullOrWhiteSpace($gia)) {
                [void]$out.Add(@{ Clau = ('sensegia|' + [string]$act.carpeta); Grup = 'sensegia'; Text = ([string]$act.titular + "  (carpeta: " + [string]$act.carpeta + ")") })
            } elseif (-not $map.ContainsKey($gia)) {
                [void]$out.Add(@{ Clau = ('notrobat|' + $gia); Grup = 'notrobat'; Text = $etiq })
            } elseif (-not (_ExcelActivitatActualitzada $map[$gia])) {
                $t = $etiq
                $infEstat = _InformeQueDeterminaEstat $act
                $dataEstat = if ($null -ne $infEstat) { _DataInformeDdMmAaaa $infEstat.data } else { '' }
                if ($dataEstat) { $t += " - INFORME ENGINYER " + $dataEstat }
                [void]$out.Add(@{ Clau = ('precinte|' + $gia); Grup = 'precinte'; Text = $t })
            }
        }
        if ([string]::IsNullOrWhiteSpace($gia) -or -not $map.ContainsKey($gia)) { continue }
        $vm = _CxVisitesPerMotiu $act
        $ctlD = @(@($ctl[$gia]) | Where-Object { $_ }); $insD = @(@($ins[$gia]) | Where-Object { $_ })
        $hiDiu = { param($d) if (@($d).Count -gt 0) { @($d) -join ', ' } else { 'buit' } }
        $nomCtl = "'Data control inicial/verificaci" + [char]0x00F3 + "'"
        $nomIns = "'Inspecci" + [char]0x00F3 + "'"
        # Pla de verificacio -> 'Data control inicial/verificacio'.
        if ($null -ne $vm.Pla -and [bool]$res.ColControl -and $ctlD -notcontains [string]$vm.Pla.Data) {
            [void]$out.Add(@{ Clau = ('visita|' + $gia + '|' + [string]$vm.Pla.Data); Grup = 'visita'; Text = ($etiq + " - visita Pla de verificaci" + [char]0x00F3 + " " + [string]$vm.Pla.Data + " (" + [string]$vm.Pla.Fitxer + "); a " + $nomCtl + " hi diu: " + (& $hiDiu $ctlD)) })
        }
        # Queixa -> 'Inspeccio'.
        if ($null -ne $vm.Queixa -and [bool]$res.ColInspeccio -and $insD -notcontains [string]$vm.Queixa.Data) {
            [void]$out.Add(@{ Clau = ('visitaqueixa|' + $gia + '|' + [string]$vm.Queixa.Data); Grup = 'visitaqueixa'; Text = ($etiq + " - visita Queixa " + [string]$vm.Queixa.Data + " (" + [string]$vm.Queixa.Fitxer + "); a " + $nomIns + " hi diu: " + (& $hiDiu $insD)) })
        }
        # Un informe d'ABANS, sense motiu: nomes si es la visita mes nova, i val
        # qualsevol de les dues columnes (no se sap quina li toca). Mateixa clau
        # que abans: els avisos ja descartats segueixen descartats.
        $s = $vm.Sense
        $mesNova = ($null -ne $s) -and ($null -eq $vm.Pla -or $s.Dt -gt $vm.Pla.Dt) -and ($null -eq $vm.Queixa -or $s.Dt -gt $vm.Queixa.Dt)
        if ($mesNova -and ([bool]$res.ColControl -or [bool]$res.ColInspeccio) -and $ctlD -notcontains [string]$s.Data -and $insD -notcontains [string]$s.Data) {
            [void]$out.Add(@{ Clau = ('visita|' + $gia + '|' + [string]$s.Data); Grup = 'visitasense'; Text = ($etiq + " - visita " + [string]$s.Data + " (" + [string]$s.Fitxer + ", sense motiu); a " + $nomCtl + " hi diu: " + (& $hiDiu $ctlD) + $(if ([bool]$res.ColInspeccio) { "; a " + $nomIns + ": " + (& $hiDiu $insD) } else { '' })) })
        }
    }
    return $out.ToArray()
}

# ----------------------------------------------------------------------------
# ELS AVISOS DESCARTATS (octubre 2026). L'usuari: "si surt un avis i jo
# considero que esta arreglat per una altra banda he de poder descartar-lo per
# la propera vegada". Es desen a local\base-dades-activitats (aquest PC; porten
# noms de titulars, mai al GitHub): { descartats: [ { clau; text; el } ] }.
# ----------------------------------------------------------------------------
function _CxDescartatsPath {
    if ([string]::IsNullOrWhiteSpace($LocalActivitatsDir)) { return '' }
    return [string](Join-Path $LocalActivitatsDir 'comprovar-excel-descartats.json')
}

# Llista de @{ Clau; Text; El }. Mai peta.
function Read-CxDescartats([string]$path = '') {
    if (-not $path) { $path = _CxDescartatsPath }
    $out = New-Object System.Collections.ArrayList
    if (-not $path) { return $out.ToArray() }
    $o = $null
    try { $o = Read-JsonFile $path } catch { $o = $null }
    foreach ($d in @(_PropInf $o 'descartats')) {
        if ($null -eq $d -or -not [string]$d.clau) { continue }
        [void]$out.Add(@{ Clau = [string]$d.clau; Text = [string](_PropInf $d 'text'); El = [string](Read-JsonIso (_PropInf $d 'el')) })
    }
    return $out.ToArray()
}

function Save-CxDescartats($llista, [string]$path = '') {
    if (-not $path) { $path = _CxDescartatsPath }
    if (-not $path) { throw "No hi ha carpeta local per desar els avisos descartats." }
    $dir = Split-Path -Parent $path
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    Write-JsonFile $path ([ordered]@{ descartats = @(@($llista) | ForEach-Object { [ordered]@{ clau = [string]$_.Clau; text = [string]$_.Text; el = [string]$_.El } }) }) 4
}

# Separa els avisos en els que s'han d'ensenyar i els descartats. Un descartat
# que ja no surt (l'Excel ja esta be) es treu de la llista: si mai tornes a
# passar, tornara a avisar. Torna @{ Actius; Amagats; Descartats (la llista
# neta per desar) }. PURA.
function _CxFiltraDescartats($avisos, $descartats) {
    $per = @{}
    foreach ($d in @($descartats)) { if ($null -ne $d) { $per[[string]$d.Clau] = $d } }
    $act = New-Object System.Collections.ArrayList
    $ama = New-Object System.Collections.ArrayList
    $nets = New-Object System.Collections.ArrayList
    foreach ($a in @($avisos)) {
        if ($null -eq $a) { continue }
        if ($per.ContainsKey([string]$a.Clau)) { [void]$ama.Add($a); [void]$nets.Add($per[[string]$a.Clau]) } else { [void]$act.Add($a) }
    }
    return @{ Actius = $act.ToArray(); Amagats = $ama.ToArray(); Descartats = $nets.ToArray() }
}

# Els titols de cada grup, en l'ordre en que surten.
$Script:CxGrups = [ordered]@{
    precinte = ("DESACTUALITZADES a l'Excel (precinte sense el Camp Info amb SI)")
    visita   = ("VISITA del PLA DE VERIFICACIO que no consta a 'Data control inicial/verificaci" + [char]0x00F3 + "'")
    visitaqueixa = ("VISITA per una QUEIXA que no consta a 'Inspecci" + [char]0x00F3 + "'")
    visitasense  = ("VISITA d'un informe sense motiu (d'abans) que no consta ni a 'Data control inicial/verificaci" + [char]0x00F3 + "' ni a 'Inspecci" + [char]0x00F3 + "'")
    notrobat = ("NO trobades a l'Excel (GIA inexistent a la fulla Est" + [char]0x00E8 + "s)")
    sensegia = ("NO verificables (activitat sense GIA a la base d'informes)")
}

# El text d'un avis a la llista: el grup al davant. PURA.
function _CxTextLlista($a) {
    $g = switch ([string]$a.Grup) { 'precinte' { 'Precinte' } 'visita' { 'Visita (Pla)' } 'visitaqueixa' { 'Visita (Queixa)' } 'visitasense' { 'Visita' } 'notrobat' { 'No trobada' } default { 'Sense GIA' } }
    return ('[' + $g + '] ' + [string]$a.Text)
}

# La finestra: els avisos amb una casella; "Descartar els marcats" els amaga per
# a les properes vegades i "Recuperar descartats" els torna a ensenyar.
function _CxFinestra($resultat, [string]$peu, [string]$descPath) {
    $st = @{ R = $resultat; Canvis = $false }
    $form = _NewForm
    $form.Text = 'Excel Comprovar'
    $form.ClientSize = New-Object System.Drawing.Size(860, 560)
    $form.MinimumSize = New-Object System.Drawing.Size(560, 420)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(14, 66); $lbl.Size = New-Object System.Drawing.Size(832, 36)
    $lbl.Anchor = 'Top, Left, Right'
    [void]$form.Controls.Add($lbl)
    $clb = New-Object System.Windows.Forms.CheckedListBox
    $clb.Location = New-Object System.Drawing.Point(14, 104); $clb.Size = New-Object System.Drawing.Size(832, 300)
    $clb.Anchor = 'Top, Bottom, Left, Right'
    $clb.CheckOnClick = $true; $clb.IntegralHeight = $false; $clb.HorizontalScrollbar = $true
    [void]$form.Controls.Add($clb)
    $tbPeu = New-Object System.Windows.Forms.TextBox
    $tbPeu.Multiline = $true; $tbPeu.ReadOnly = $true; $tbPeu.ScrollBars = 'Vertical'
    $tbPeu.Location = New-Object System.Drawing.Point(14, 412); $tbPeu.Size = New-Object System.Drawing.Size(832, 84)
    $tbPeu.Anchor = 'Bottom, Left, Right'
    $tbPeu.BackColor = [System.Drawing.Color]::White
    $tbPeu.Text = $peu
    [void]$form.Controls.Add($tbPeu)
    $grups = $Script:CxGrups
    $omple = {
        $clb.Items.Clear()
        $actius = @($st.R.Actius)
        foreach ($g in @($grups.Keys)) { foreach ($a in $actius) { if ([string]$a.Grup -eq $g) { [void]$clb.Items.Add((_CxTextLlista $a)) } } }
        $st.Ordre = @(foreach ($g in @($grups.Keys)) { foreach ($a in $actius) { if ([string]$a.Grup -eq $g) { $a } } })
        $lbl.Text = $(if ($actius.Count -eq 0) { "Cap av" + [char]0x00ED + "s pendent: l'Excel est" + [char]0x00E0 + " al dia." } else { [string]$actius.Count + " av" + [char]0x00ED + "s pendents. Marca els que ja estan arreglats per una altra banda i prem Descartar: no tornaran a sortir." }) +
                    $(if (@($st.R.Amagats).Count -gt 0) { "  (" + @($st.R.Amagats).Count + " descartats, que no es mostren.)" } else { '' })
    }.GetNewClosure()
    $st.Omple = $omple
    & $omple
    $peuB = _AddPeuBotons $form @(@{ Nom = 'Tancar'; Text = 'Tancar'; Resultat = 'Cancel'; Esc = $true }, @{ Nom = 'Recup'; Text = 'Recuperar descartats' }) @(@{ Nom = 'Desc'; Text = 'Descartar els marcats'; Estil = 'primari' }) 510 -Ancorat
    $peuB.Desc.add_Click({
        $idx = @($clb.CheckedIndices | ForEach-Object { [int]$_ })
        if ($idx.Count -eq 0) { return }
        $ara = (Get-Date).ToString('o')
        $nous = @($st.R.Descartats) + @($idx | ForEach-Object { $a = $st.Ordre[$_]; @{ Clau = [string]$a.Clau; Text = [string]$a.Text; El = $ara } })
        try { Save-CxDescartats $nous $descPath } catch {
            [System.Windows.Forms.MessageBox]::Show("No s'han pogut desar:`n" + $_.Exception.Message, 'Excel Comprovar', 'OK', 'Error') | Out-Null
            return
        }
        $tots = @(@($st.R.Actius) + @($st.R.Amagats))
        $st.R = _CxFiltraDescartats $tots $nous
        & $st.Omple
    }.GetNewClosure())
    $peuB.Recup.add_Click({
        if (@($st.R.Amagats).Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No hi ha cap av" + [char]0x00ED + "s descartat.", 'Excel Comprovar', 'OK', 'Information') | Out-Null
            return
        }
        $llista = (@($st.R.Amagats) | ForEach-Object { '  - ' + (_CxTextLlista $_) }) -join "`n"
        $rr = [System.Windows.Forms.MessageBox]::Show(("Avisos descartats:`n`n" + $llista + "`n`nVols que tornin a sortir tots?"), 'Excel Comprovar', 'YesNo', 'Question')
        if ($rr -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        try { Save-CxDescartats @() $descPath } catch { return }
        $tots = @(@($st.R.Actius) + @($st.R.Amagats))
        $st.R = _CxFiltraDescartats $tots @()
        & $st.Omple
    }.GetNewClosure())
    [void](_AddBrandHeader $form ("Comprovaci" + [char]0x00F3 + " de l'Excel") ("Precintes i visites d'inspecci" + [char]0x00F3 + " pendents a l'Excel") 56)
    [void]$form.ShowDialog()
    $form.Dispose()
}

# L'eina: avisos de la base d'informes contra l'Excel (precintes i visites).
#
# Aqui hi havia un _SaveRunTimestamp que escrivia 'comprovat_el' a
# comprovar-excel-state.json. Es va treure: l'unica cosa que en feia servei era
# ensenyar el segell d'"ultima execucio" al menu, i ara aixo ho porta un SOL
# registre per a totes les eines (_MarcaEinaUsada, Seguiment.ps1), apuntat des
# del despatxador. El fitxer d'estat d'aquesta eina ja no fa falta.
function Invoke-ComprovarExcel {
    $outPath = Get-InformesDbPath
    if (-not (Test-Path -LiteralPath $outPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Encara no hi ha cap base d'informes.`n`nExecuta primer 'Actualitzar base'.",
            'Excel Comprovar', 'OK', 'Information') | Out-Null
        return
    }
    try {
        $db = Read-JsonFile $outPath
    } catch {
        [System.Windows.Forms.MessageBox]::Show("No s'ha pogut llegir la base d'informes:`n$($_.Exception.Message)", 'Excel Comprovar', 'OK', 'Error') | Out-Null
        return
    }
    $res = _ReadExcelCampInfoPerGia
    if (-not $res.Ok) {
        [System.Windows.Forms.MessageBox]::Show("No s'ha pogut llegir l'Excel:`n$($res.Error)", 'Excel Comprovar', 'OK', 'Error') | Out-Null
        return
    }
    $avisos = @(Get-ComprovacioExcelAvisos $db $res)
    $descPath = _CxDescartatsPath
    $r = _CxFiltraDescartats $avisos @(Read-CxDescartats $descPath)
    # Els descartats que ja no surten es treuen del fitxer (vegeu _CxFiltraDescartats).
    try { if ($descPath -and (Test-Path -LiteralPath $descPath)) { Save-CxDescartats @($r.Descartats) $descPath } } catch { }

    # El peu: el criteri i el DIAGNOSTIC dels noms dels camps. Si algun dia es
    # reanomenen (va passar: 'PRECINTE?' -> 'PRECINTE ACTIVITAT?') el criteri
    # deixa de trobar-los i tot surt com a desactualitzat; aixi es veu de seguida.
    $nPrec = @($db.activitats | Where-Object { [string]$_.estat_actual -eq $Script:EstatPrecinte }).Count
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine(("Precintes: les {0} activitats en Estat '{1}' han de tenir a l'Excel un Camp Info {2} amb valor que comenci per SI." -f $nPrec, $Script:EstatPrecinte, (_ExcelPrecinteCampsText)))
    [void]$sb.AppendLine("Visites: la darrera visita del Pla de verificaci" + [char]0x00F3 + " ha de constar a 'Data control inicial/verificaci" + [char]0x00F3 + "', i la darrera per una Queixa, a 'Inspecci" + [char]0x00F3 + "' (el motiu surt de l'Objecte de l'informe; els informes d'abans, sense motiu, valen a qualsevol de les dues).")
    if (-not [bool]$res.ColControl) { [void]$sb.AppendLine("   NO s'ha trobat la columna 'Data control inicial/verificaci" + [char]0x00F3 + "' a la fulla Est" + [char]0x00E8 + "s.") }
    if (-not [bool]$res.ColInspeccio) { [void]$sb.AppendLine("   NO s'ha trobat la columna 'Inspecci" + [char]0x00F3 + "' a la fulla Est" + [char]0x00E8 + "s.") }
    $campsExcel = @(_ExcelCampsPrecinteTrobats $res.Map)
    if ($campsExcel.Count -gt 0) {
        [void]$sb.AppendLine("Camps de l'Excel que parlen de precinte o decret (per si s'han reanomenat):")
        foreach ($c in $campsExcel) {
            $marca = if ($Script:ExcelPrecinteCampNoms -contains (_NormalitzaText $c)) { '  <-- es el que es comprova' } else { '' }
            [void]$sb.AppendLine("     - " + $c + $marca)
        }
    }
    _CxFinestra $r ($sb.ToString()) $descPath
}
