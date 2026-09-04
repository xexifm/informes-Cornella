#requires -Version 5.1
<#
.SYNOPSIS
  Eina "Controls periodics": llistat d'activitats amb control periodic.

.DESCRIPTION
  Llegeix la base d'activitats (Excel, fulla "Estes"/"Estès") i llista les
  activitats que:
    - tenen "Classificacio general annex" = II o III, O BE
    - tenen "Classificacio general Apartat" amb un numero que comenca per 561
      (561, 5610, ...).
  Mostra una graella (mateixa estructura que "Editar base d'informes") amb les
  columnes demanades, filtrable per II / III / 561 i ordenable per columna;
  per defecte, ordenada per "Data llicencia/comunicacio" ascendent (mes antic
  primer). Reutilitza helpers de GenerarInforme.ps1 (Read-FullaEstesa,
  _FindColIndex, _FormatDateOnly, Find-LatestActivitatsExcel, _ResolveOutputDir,
  _GetUniqueOutputPath) i de la resta del programa (_NewForm, _AddBrandHeader,
  _StylePrimaryButton/_StyleSecondaryButton).

  Nomes defineix funcions (cap execucio en carregar-se): segur en headless.
#>

# ----------------------------------------------------------------------------
# Funcions PURES (testejables en headless)
# ----------------------------------------------------------------------------

# Classifica una activitat pel valor de "Classificacio general annex" i
# "Classificacio general Apartat". Retorna un objecte amb IsII/IsIII/Is561 i
# Qualifies (= compleix algun dels criteris de control periodic).
#   annex  -> II / III (roma; s'accepta "Annex II", "II", etc. per paraula)
#   apartat-> un numero que COMENCA per 561 (561, 5610, ...; no "1561")
function _ControlPeriodicClassify([string]$annex, [string]$apartat) {
    $a = ([string]$annex).ToUpper()
    $isII  = [bool]($a -match '\bII\b')
    $isIII = [bool]($a -match '\bIII\b')
    # Un numero que comenci per 561: "561" al principi de la cadena o precedit
    # per un caracter que NO es un digit (aixi "1561" queda exclos).
    $is561 = [bool]([string]$apartat -match '(^|\D)561')
    return [pscustomobject]@{
        IsII      = $isII
        IsIII     = $isIII
        Is561     = $is561
        Qualifies = ($isII -or $isIII -or $is561)
    }
}

# Converteix una cel·la de data (double OLE o text) a [datetime], o $null si no
# es pot interpretar. S'usa com a clau d'ordenacio de les columnes de data.
function _ParseCellDate($v) {
    if ($null -eq $v) { return $null }
    if ($v -is [double]) { try { return [DateTime]::FromOADate($v) } catch { return $null } }
    $s = ([string]$v).Trim()
    if ($s -eq '') { return $null }
    $dt = [datetime]::MinValue
    if ([datetime]::TryParse($s, [ref]$dt)) { return $dt }
    foreach ($fmt in 'dd/MM/yyyy', 'd/M/yyyy', 'yyyy-MM-dd', 'dd-MM-yyyy') {
        if ([datetime]::TryParseExact($s, $fmt, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) { return $dt }
    }
    return $null
}

# ----------------------------------------------------------------------------
# Lectura de l'Excel
# ----------------------------------------------------------------------------
# Llegeix la fulla "Estes" i retorna @{ Ok; Rows; File; Error }. Cada fila
# (activitat que compleix el criteri) es un PSCustomObject amb els camps de
# visualitzacio + claus de data per ordenar + flags de classificacio.
function _ReadControlsPeriodics {
    $latest = Find-LatestActivitatsExcel
    if ($null -eq $latest) { return @{ Ok = $false; Error = "No s'ha trobat cap Excel d'activitats." } }
    # Read-FullaEstesa LLANCA quan l'Excel no arrenca o la fulla no hi es; aqui
    # el crider espera un @{ Ok = $false; Error }, o sigui que s'hi posa un catch.
    try {
        return (Read-FullaEstesa $latest.File {
            param($x)
            $data = $x.Data; $rows = $x.Rows; $cols = $x.Cols
            if ($null -eq $data) { return @{ Ok = $true; Rows = @(); File = $latest.File.Name } }

            # Localitzacio de columnes pel TEXT de la capcalera (amb fallback a
            # l'index fix conegut quan cal).
            $cId    = 1
            $cRao   = 10
            $cAct   = 94
            $cVia   = _FindColIndex $data $cols @('emp', 'tipus via') $null; if ($cVia -eq 0) { $cVia = 48 }
            $cCar   = _FindColIndex $data $cols @('emp', 'carrer')    $null; if ($cCar -eq 0) { $cCar = 49 }
            $cNum   = _FindColIndex $data $cols @('emp', 'numero')    $null; if ($cNum -eq 0) { $cNum = 50 }
            $cLle   = _FindColIndex $data $cols @('emp', 'lletra')    $null; if ($cLle -eq 0) { $cLle = 52 }
            $cRaoMail = _FindColIndex $data $cols @('rao soc', 'mail')      $null
            $cRepMail = _FindColIndex $data $cols @('rep', 'leg', 'mail')   $null
            $cLlic    = _FindColIndex $data $cols @('data', 'llicencia')    $null
            $cCtrlIni = _FindColIndex $data $cols @('control', 'inicial')   $null
            $cPeriod  = _FindColIndex $data $cols @('periodicitat')         $null
            $cCtrlPer = _FindColIndex $data $cols @('control', 'periodic')  $null
            $cProper  = _FindColIndex $data $cols @('proper', 'previst')    $null
            $cAnnex   = _FindColIndex $data $cols @('classificacio', 'annex')   $null
            $cApart   = _FindColIndex $data $cols @('classificacio', 'apartat') $null
            $cExp     = _FindColIndex $data $cols @('expedient')                $null

            if ($cAnnex -eq 0 -and $cApart -eq 0) {
                return @{ Ok = $false; Error = "No s'han trobat les columnes 'Classificació general annex' ni 'Classificació general Apartat' a la fulla Estès." }
            }

            $get = $x.Cel   # el lector de cel·la ve amb el context (Excel.ps1)
            $getDate = {
                param($r, $c)
                if ($c -lt 1 -or $c -gt $cols) { return '' }
                return (_FormatDateOnly $data[$r, $c])
            }
            $getKey = {
                param($r, $c)
                if ($c -lt 1 -or $c -gt $cols) { return $null }
                return (_ParseCellDate $data[$r, $c])
            }

            $out = New-Object System.Collections.ArrayList
            for ($r = 2; $r -le $rows; $r++) {
                $annex = & $get $r $cAnnex
                $apart = & $get $r $cApart
                $cl = _ControlPeriodicClassify $annex $apart
                if (-not $cl.Qualifies) { continue }

                $idCell = $data[$r, $cId]
                $id = if ($idCell -is [double]) {
                    if ([math]::Floor($idCell) -eq $idCell) { [string][int]$idCell } else { [string]$idCell }
                } else { ([string]$idCell).Trim() }
                if ([string]::IsNullOrWhiteSpace($id)) { continue }

                $via = & $get $r $cVia; $car = & $get $r $cCar; $num = & $get $r $cNum; $lle = & $get $r $cLle
                $adr = (@($via, $car, $num, $lle) | Where-Object { $_ -ne '' }) -join ' '

                [void]$out.Add([pscustomobject]@{
                    Id              = $id
                    RaoSocial       = (& $get $r $cRao)
                    RaoEmail        = (& $get $r $cRaoMail)
                    RepEmail        = (& $get $r $cRepMail)
                    Adreca          = $adr
                    DataLlic        = (& $getDate $r $cLlic)
                    DataControlIni  = (& $getDate $r $cCtrlIni)
                    PeriodicitatCP  = (& $get $r $cPeriod)
                    DataControlPer  = (& $getDate $r $cCtrlPer)
                    ProperCP        = (& $getDate $r $cProper)
                    Annex           = $annex
                    Apartat         = $apart
                    ActPrincipal    = (& $get $r $cAct)
                    KDataLlic       = (& $getKey $r $cLlic)
                    KDataControlIni = (& $getKey $r $cCtrlIni)
                    KDataControlPer = (& $getKey $r $cCtrlPer)
                    KProperCP       = (& $getKey $r $cProper)
                    IsII            = $cl.IsII
                    IsIII           = $cl.IsIII
                    Is561           = $cl.Is561
                    Exp             = (& $get $r $cExp)
                    Sel             = $false
                })
            }
            return @{ Ok = $true; Rows = @($out); File = $latest.File.Name }
        })
    } catch {
        return @{ Ok = $false; Error = $_.Exception.Message }
    }
}

# ----------------------------------------------------------------------------
# Finestra (graella)
# ----------------------------------------------------------------------------
# Metadades de columnes: capcalera visible, clau al PSCustomObject, amplada i si
# es una data (per ordenar per la clau K* en lloc del text).
$Script:ControlsPeriodicsCols = @(
    @{ H = 'ID Activitat';                 Key = 'Id';             W = 80;  Date = $false }
    @{ H = 'Raó social';                   Key = 'RaoSocial';      W = 200; Date = $false }
    @{ H = 'Raó soc. E-mail';              Key = 'RaoEmail';       W = 170; Date = $false }
    @{ H = 'Rep. Leg. E-mail';             Key = 'RepEmail';       W = 170; Date = $false }
    @{ H = 'Adreça';                       Key = 'Adreca';         W = 220; Date = $false }
    @{ H = 'Data llicència/comunicació';   Key = 'DataLlic';       W = 130; Date = $true;  KKey = 'KDataLlic' }
    @{ H = 'Data control inicial/verif.';  Key = 'DataControlIni'; W = 130; Date = $true;  KKey = 'KDataControlIni' }
    @{ H = 'Periodicitat CP';              Key = 'PeriodicitatCP'; W = 110; Date = $false }
    @{ H = 'Data control periòdic';        Key = 'DataControlPer'; W = 120; Date = $true;  KKey = 'KDataControlPer' }
    @{ H = 'Proper CP previst';            Key = 'ProperCP';       W = 120; Date = $true;  KKey = 'KProperCP' }
    @{ H = 'Classif. annex';               Key = 'Annex';          W = 90;  Date = $false }
    @{ H = 'Classif. Apartat';             Key = 'Apartat';        W = 110; Date = $false }
    @{ H = 'Activitat principal';          Key = 'ActPrincipal';   W = 200; Date = $false }
)

function Show-ControlsPeriodicsWindow($allRows, [string]$fileName) {
    $colsMeta = $Script:ControlsPeriodicsCols

    # La columna 0 de la graella es una casella de seleccio; les columnes de
    # dades comencen a l'index 1. SortColIdx guarda l'index de GRAELLA de la
    # columna d'ordenacio; el mapatge a $colsMeta es (SortColIdx - 1).
    # (SortColIdx/SortAsc son els noms que espera _EnableHeaderSort, UiComuns.ps1.)
    # Ordre per defecte: "Proper CP previst" ascendent (el mes antic primer).
    $defaultSort = 0
    for ($i = 0; $i -lt $colsMeta.Count; $i++) { if ($colsMeta[$i].Key -eq 'ProperCP') { $defaultSort = $i; break } }
    $state = @{ Loading = $false; SortColIdx = ($defaultSort + 1); SortAsc = $true }

    $form = _NewForm
    $form.Text = 'Controls periodics'
    $form.Size = New-Object System.Drawing.Size(1320, 720)
    $form.MinimumSize = New-Object System.Drawing.Size(820, 480)
    $form.StartPosition = 'CenterScreen'

    # ---- Graella --------------------------------------------------------
    $grid = New-Object System.Windows.Forms.DataGridView
    _StyleListGrid $grid
    # Columna 0: casella per triar quines activitats generar (editable).
    $cSel = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $cSel.HeaderText = 'Generar'
    $cSel.Width = 58
    $cSel.SortMode = 'Programmatic'
    [void]$grid.Columns.Add($cSel)
    foreach ($cm in $colsMeta) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.HeaderText = $cm.H
        $col.Width = $cm.W
        $col.SortMode = 'Programmatic'
        $col.ReadOnly = $true
        [void]$grid.Columns.Add($col)
    }

    # ---- Barra superior: filtre de classificacio + cerca ----------------
    $topPanel = New-Object System.Windows.Forms.Panel
    $topPanel.Dock = 'Top'; $topPanel.Height = 42
    $lblF = New-Object System.Windows.Forms.Label
    $lblF.Text = 'Classificació:'; $lblF.AutoSize = $true
    $lblF.Location = New-Object System.Drawing.Point(10, 13)
    $topPanel.Controls.Add($lblF)
    # Filtre de SELECCIO MULTIPLE (cap marcat = totes). L'accio crida $fill
    # (definit mes avall; ja existeix quan l'usuari hi toca).
    $mfClass = _MakeMultiFilter $topPanel 96 10 104 'Tots' @('II', 'III', '561') { & $fill }
    $txtCerca = _AddSearchBox $topPanel 214 10 300 'Cerca:' { & $fill }
    $lblCompte = New-Object System.Windows.Forms.Label
    $lblCompte.AutoSize = $true
    $lblCompte.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
    $lblCompte.Location = New-Object System.Drawing.Point(590, 13)
    $topPanel.Controls.Add($lblCompte)

    # ---- (Re)omple la graella: filtre + cerca + ordre -------------------
    $fill = {
        $state.Loading = $true
        $grid.Rows.Clear()
        $selClass = & $mfClass.GetSelected
        $n = ([string]$txtCerca.Text).Trim().ToLower()

        $sel = foreach ($row in $allRows) {
            if ($selClass.Count -gt 0) {
                $okC = ($selClass -contains 'II'  -and $row.IsII)  -or
                       ($selClass -contains 'III' -and $row.IsIII) -or
                       ($selClass -contains '561' -and $row.Is561)
                if (-not $okC) { continue }
            }
            $hay = $row.Id + ' ' + $row.RaoSocial + ' ' + $row.RaoEmail + ' ' + $row.RepEmail + ' ' + $row.Adreca + ' ' + $row.PeriodicitatCP + ' ' + $row.Annex + ' ' + $row.Apartat + ' ' + $row.ActPrincipal
            if (-not (_TextMatches $hay $n)) { continue }
            $row
        }
        $sel = @($sel)

        # Adjuntem una clau d'ordenacio a cada fila i ordenem per nom de
        # propietat (evita problemes d'abast amb scriptblocks dins de Sort-Object).
        # $state.SortColIdx es index de GRAELLA (col 0 = casella); dades a -1.
        $cm = $colsMeta[[Math]::Max(0, $state.SortColIdx - 1)]
        $isDate = [bool]$cm.Date
        $keyProp = $cm.Key
        $kkeyProp = $cm.KKey
        foreach ($row in $sel) {
            if ($isDate) {
                $v = $row.$kkeyProp
                $sv = if ($null -eq $v) { [datetime]::MaxValue } else { [datetime]$v }
            } else {
                $sv = ([string]$row.$keyProp).ToLower()
            }
            Add-Member -InputObject $row -NotePropertyName '_SortKey' -NotePropertyValue $sv -Force
        }
        $sorted = @($sel | Sort-Object -Property '_SortKey' -Descending:(-not $state.SortAsc))

        foreach ($row in $sorted) {
            $idx = $grid.Rows.Add(@(
                [bool]$row.Sel,
                $row.Id, $row.RaoSocial, $row.RaoEmail, $row.RepEmail, $row.Adreca,
                $row.DataLlic, $row.DataControlIni, $row.PeriodicitatCP, $row.DataControlPer,
                $row.ProperCP, $row.Annex, $row.Apartat, $row.ActPrincipal
            ))
            $grid.Rows[$idx].Tag = $row
        }
        $lblCompte.Text = "$($sorted.Count) activitats"
        $state.Loading = $false
    }.GetNewClosure()

    $txtCerca.add_TextChanged({ & $fill }.GetNewClosure())

    # Persistencia de la casella "Generar" (sobreviu a filtres/ordenacio: es desa
    # a l'objecte fila). Commit immediat del clic.
    $grid.add_CurrentCellDirtyStateChanged({
        if ($grid.IsCurrentCellDirty) { $grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) }
    }.GetNewClosure())
    $grid.add_CellValueChanged({
        param($s, $e)
        if ($state.Loading) { return }
        if ($e.RowIndex -lt 0 -or $e.ColumnIndex -ne 0) { return }
        $row = $s.Rows[$e.RowIndex].Tag
        if ($null -ne $row) { $row.Sel = [bool]$s.Rows[$e.RowIndex].Cells[0].Value }
    }.GetNewClosure())

    # Col 0 = casella "Generar": no s'ordena (minCol = 1).
    _EnableHeaderSort $grid $state 1 @() { & $fill }

    # ---- Barra inferior: Exportar / Tancar ------------------------------
    $botPanel = New-Object System.Windows.Forms.Panel
    $botPanel.Dock = 'Bottom'; $botPanel.Height = 48
    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = 'Exportar (CSV)'; $btnExport.Size = New-Object System.Drawing.Size(150, 30)
    $btnExport.Location = New-Object System.Drawing.Point(10, 9)
    _StyleSecondaryButton $btnExport
    $btnExport.add_Click({ _ExportControlsPeriodics $grid $colsMeta }.GetNewClosure())
    $botPanel.Controls.Add($btnExport)
    $btnGenerar = New-Object System.Windows.Forms.Button
    $btnGenerar.Text = 'Generar informes'; $btnGenerar.Size = New-Object System.Drawing.Size(160, 30)
    $btnGenerar.Location = New-Object System.Drawing.Point(170, 9)
    _StylePrimaryButton $btnGenerar
    $btnGenerar.add_Click({
        $triades = @()
        foreach ($r in $allRows) { if ($r.Sel) { $triades += $r } }
        Invoke-GenerarControlsPeriodics $triades
    }.GetNewClosure())
    $botPanel.Controls.Add($btnGenerar)
    # Enviar correu als titulars de les activitats marcades (esborranys a Outlook).
    $btnCorreu = New-Object System.Windows.Forms.Button
    $btnCorreu.Text = 'Enviar correu (esborranys)'; $btnCorreu.Size = New-Object System.Drawing.Size(200, 30)
    $btnCorreu.Location = New-Object System.Drawing.Point(340, 9)
    _StylePrimaryButton $btnCorreu
    $btnCorreu.add_Click({
        $triades = @()
        foreach ($r in $allRows) { if ($r.Sel) { $triades += $r } }
        Invoke-ControlsCpEmailDrafts $triades
    }.GetNewClosure())
    $botPanel.Controls.Add($btnCorreu)
    # Editar el text del correu (assumpte + cos amb variables).
    $btnText = New-Object System.Windows.Forms.Button
    $btnText.Text = 'Editar text'; $btnText.Size = New-Object System.Drawing.Size(120, 30)
    $btnText.Location = New-Object System.Drawing.Point(550, 9)
    _StyleSecondaryButton $btnText
    $btnText.add_Click({ Invoke-ControlsCpEmailTextos }.GetNewClosure())
    $botPanel.Controls.Add($btnText)
    $btnTancar = New-Object System.Windows.Forms.Button
    $btnTancar.Text = 'Tancar'; $btnTancar.Size = New-Object System.Drawing.Size(120, 30)
    $btnTancar.Location = New-Object System.Drawing.Point(690, 9)
    _StyleSecondaryButton $btnTancar
    $btnTancar.add_Click({ $form.Close() }.GetNewClosure())
    $botPanel.Controls.Add($btnTancar)

    # Ordre d'afegit: Fill (centre), Top, Bottom, i la banda l'ultima (a dalt).
    $form.Controls.Add($grid)
    $form.Controls.Add($topPanel)
    $form.Controls.Add($botPanel)
    $sub = "Activitats amb control periòdic (annex II/III o apartat 561)  ·  $fileName"
    [void](_AddBrandHeader $form 'Controls periòdics' $sub 56)

    # Ordre inicial: per "Proper CP previst" ascendent (el mes antic primer).
    _SetSortGlyph $grid $state.SortColIdx $true
    & $fill
    [void]$form.ShowDialog()
}

# Exporta el contingut ACTUAL de la graella (tal com es veu, filtrat i ordenat)
# a un CSV obrible amb Excel.
function _ExportControlsPeriodics($grid, $colsMeta) {
    if ($grid.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No hi ha cap activitat per exportar.', 'Controls periòdics', 'OK', 'Information') | Out-Null
        return
    }
    $rows = New-Object System.Collections.ArrayList
    foreach ($gr in $grid.Rows) {
        $o = [ordered]@{}
        # Les columnes de dades comencen a l'index 1 (la 0 es la casella Generar).
        for ($c = 0; $c -lt $colsMeta.Count; $c++) {
            $o[$colsMeta[$c].H] = [string]$gr.Cells[$c + 1].Value
        }
        [void]$rows.Add([pscustomobject]$o)
    }
    $dir  = _ResolveOutputDir
    $name = 'Controls periodics ' + (Get-Date).ToString('yyyy-MM-dd') + '.csv'
    $path = _GetUniqueOutputPath $dir $name
    try {
        $rows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    } catch {
        [System.Windows.Forms.MessageBox]::Show("No s'ha pogut escriure el CSV:`n$($_.Exception.Message)", 'Controls periòdics', 'OK', 'Error') | Out-Null
        return
    }
    $r = [System.Windows.Forms.MessageBox]::Show("CSV generat:`n$path`n`nVols obrir-lo ara?", 'Controls periòdics', 'YesNo', 'Information')
    if ($r -eq [System.Windows.Forms.DialogResult]::Yes) {
        try { Start-Process -FilePath $path | Out-Null } catch { }
    }
}

# ----------------------------------------------------------------------------
# Punt d'entrada (menu EINES)
# ----------------------------------------------------------------------------
function Invoke-ControlsPeriodics {
    # Splash mentre s'obre i es llegeix l'Excel (pot trigar uns segons).
    $sp = _NewForm
    $sp.Text = 'Controls periodics'
    $sp.Size = New-Object System.Drawing.Size(420, 130)
    $sp.FormBorderStyle = 'FixedDialog'
    $sp.MaximizeBox = $false; $sp.MinimizeBox = $false
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(20, 18); $lbl.Size = New-Object System.Drawing.Size(370, 40)
    $lbl.Text = "Llegint la base d'activitats (Excel)..."
    $sp.Controls.Add($lbl)
    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Location = New-Object System.Drawing.Point(20, 62); $bar.Size = New-Object System.Drawing.Size(370, 20)
    $bar.Style = 'Marquee'
    $sp.Controls.Add($bar)
    $sp.Show()
    [System.Windows.Forms.Application]::DoEvents()

    $res = $null
    try { $res = _ReadControlsPeriodics } finally { try { $sp.Close() } catch { } }

    if ($null -eq $res -or -not $res.Ok) {
        $msg = if ($res) { $res.Error } else { 'Error desconegut.' }
        [System.Windows.Forms.MessageBox]::Show("No s'ha pogut llegir la base d'activitats:`n$msg", 'Controls periòdics', 'OK', 'Error') | Out-Null
        return
    }
    $rows = @($res.Rows)
    if ($rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No hi ha cap activitat amb control periòdic (annex II/III o apartat 561) a la base actual.",
            'Controls periòdics', 'OK', 'Information') | Out-Null
        return
    }
    Show-ControlsPeriodicsWindow $rows $res.File
}

# ----------------------------------------------------------------------------
# Generació d'informes de requeriment (control periòdic) en lot
# ----------------------------------------------------------------------------
# Regim aplicable a una activitat, amb PRECEDENCIA 561 (Decret 112/2010) > III >
# II. Retorna 'II' | 'III' | '112' | $null. (Funcio PURA, testejable.)
function _ControlCatalegKind($row) {
    if ($row.Is561) { return '112' }
    if ($row.IsIII) { return 'III' }
    if ($row.IsII)  { return 'II' }
    return $null
}

# Titol EXACTE de la deficiencia (Titol 2 del cataleg REQ1) segons el regim.
# NO son catalegs nous: son items dins de REQ1 (grup "Controls periòdics").
function _ControlSectionTitle([string]$kind) {
    switch ($kind) {
        '112' { return 'Decret 112/2010 - control periòdic' }
        'III' { return 'Annex III Llei 20/2009 - control periòdic' }
        'II'  { return 'Annex II Llei 20/2009 - control periòdic' }
    }
    return $null
}

# Dins d'un cataleg ja parsejat, localitza l'ITEM (Titol 2) el text del qual
# coincideix amb $title i retorna les seves claus de seleccio (l'item + els seus
# fills). Buit si no es troba. (Al parser, Titol 1 = seccio, Titol 2 = item.)
function _FindItemKeysByTitle($parsed, [string]$title) {
    $tnorm = _NormalitzaText $title
    foreach ($sec in $parsed.Sections) {
        foreach ($el in $sec.Items) {
            if ($el.Kind -eq 'item' -and (_NormalitzaText $el.Short) -eq $tnorm) {
                $keys = @((_ItemKey $sec.Title $el.Short))
                foreach ($ch in $el.Children) { $keys += (_ItemKey $sec.Title $el.Short $ch.Short) }
                return $keys
            }
        }
    }
    return @()
}

# Genera, per a cada activitat triada, un informe REQUERIMENT (com "Requeriment -
# Nou") amb: capcalera de les dades de l'activitat, TOTES les deficiencies del
# cataleg de control periodic segons la classificacio, i la conclusio
# "Requeriment". Reutilitza el motor no interactiu (Get-ParsedCataleg,
# Build-SelectionFromKeys, Read-Conclusions, Build-ConclusionsFromTitles,
# Build-FieldsFromPaquet, Build-Document).
function Invoke-GenerarControlsPeriodics($rows) {
    $rows = @($rows)
    if ($rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Marca almenys una activitat (columna 'Generar') per generar-ne l'informe.", 'Generar informes', 'OK', 'Information') | Out-Null
        return
    }

    # El cataleg es REQ1; agafem les deficiencies (Titol 2) de control periodic
    # segons la classificacio. Comprovem que el cataleg existeixi i que hi siguin
    # els items que necessitem.
    $catPath = Join-Path $EstructuralsDir 'REQ1.json'
    if (-not (Test-Path -LiteralPath $catPath)) {
        [System.Windows.Forms.MessageBox]::Show("No s'ha trobat el catàleg REQ1.json a ESTRUCTURALS.", 'Generar informes', 'OK', 'Error') | Out-Null
        return
    }

    $rc = [System.Windows.Forms.MessageBox]::Show(
        ("Es generaran $($rows.Count) informes de requeriment (control periòdic), un per activitat triada.`n`n" +
         "Del catàleg REQ1: la deficiència de control periòdic segons la classificació (Decret 112/2010, Annex III o Annex II) i la conclusió 'Requeriment'. Capçalera amb les dades de l'activitat, sense Objecte.`n`nVols continuar?"),
        'Generar informes', 'YesNo', 'Question')
    if ($rc -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    # Finestra de progrés amb Cancel·lar.
    $cancel = @{ Flag = $false; Running = $true }
    $form = _NewForm
    $form.Text = 'Generar informes'
    $form.Size = New-Object System.Drawing.Size(560, 170)
    $form.FormBorderStyle = 'FixedDialog'; $form.MaximizeBox = $false
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(20, 18); $lbl.Size = New-Object System.Drawing.Size(510, 40)
    $lbl.Text = 'Preparant...'
    $form.Controls.Add($lbl)
    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Location = New-Object System.Drawing.Point(20, 66); $bar.Size = New-Object System.Drawing.Size(510, 22)
    $bar.Style = 'Continuous'; $bar.Minimum = 0; $bar.Maximum = [Math]::Max(1, $rows.Count)
    $form.Controls.Add($bar)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel·lar'; $btnCancel.Size = New-Object System.Drawing.Size(120, 30)
    $btnCancel.Location = New-Object System.Drawing.Point(410, 96)
    _StyleSecondaryButton $btnCancel
    $btnCancel.add_Click({ $cancel.Flag = $true }.GetNewClosure())
    $form.Controls.Add($btnCancel)
    $form.add_FormClosing({ param($s, $e) if ($cancel.Running) { $cancel.Flag = $true; $e.Cancel = $true } }.GetNewClosure())
    $form.Show()
    [System.Windows.Forms.Application]::DoEvents()

    $ok = 0; $err = 0; $cancelled = $false
    $errDetalls = New-Object System.Collections.ArrayList
    $word = $null
    try {
        # -Opcional perque aqui es vol el MISSATGE d'aquesta eina, no el quadre
        # generic de New-WordApp seguit d'un exit 1: l'usuari ha de tornar a la
        # llista, no perdre el programa.
        $word = New-WordApp -Opcional
        if ($null -eq $word) {
            [System.Windows.Forms.MessageBox]::Show("No s'ha pogut iniciar Microsoft Word.", 'Generar informes', 'OK', 'Error') | Out-Null
            return
        }

        # Cataleg REQ1 (una sola vegada) + conclusio "Requeriment" (grup REQ1).
        $parsed      = Get-ParsedCataleg -path $catPath
        $conclAll    = Read-Conclusions -path $ConclusionsPath -reportType 'REQ1'
        $conclusions = Build-ConclusionsFromTitles $conclAll.Selectable @('Requeriment')

        $done = 0
        foreach ($r in $rows) {
            if ($cancel.Flag) { $cancelled = $true; break }
            $done++
            $lbl.Text = "Generant informes...  $done de $($rows.Count)`nGIA $($r.Id) - $($r.RaoSocial)"
            if ($bar.Value -lt $bar.Maximum) { $bar.Value = $done }
            [System.Windows.Forms.Application]::DoEvents()

            $kind = _ControlCatalegKind $r
            $title = _ControlSectionTitle $kind
            if (-not $title) { $err++; [void]$errDetalls.Add("GIA $($r.Id): activitat sense classificació de control periòdic"); continue }
            $keys = @(_FindItemKeysByTitle $parsed $title)
            if ($keys.Count -eq 0) { $err++; [void]$errDetalls.Add("GIA $($r.Id): no s'ha trobat la deficiència '$title' al catàleg REQ1"); continue }
            try {
                $selected = Build-SelectionFromKeys $parsed.Sections $keys
                $fields   = Build-FieldsFromPaquet $selected $conclusions $conclAll.Always $null
                # Capçalera de les dades de l'activitat, SENSE línia "Objecte:".
                $header   = @{
                    ID_GIA = [string]$r.Id; EXP_NUM = [string]$r.Exp; ADRECA = [string]$r.Adreca
                    ACTIVITAT = [string]$r.ActPrincipal; TITULAR = [string]$r.RaoSocial
                    ORIGEN_TIPUS = 'cap'
                }
                [void](Build-Document -word $word -header $header `
                                      -selectedSections $selected -fields $fields `
                                      -conclusions $conclusions -alwaysConclusions $conclAll.Always `
                                      -catalegName 'REQ1' -introText $parsed.IntroText `
                                      -conclusionsHeaderText $conclAll.HeaderText `
                                      -isFixedBody $parsed.IsFixedBody -fixedBodyLines $parsed.FixedBodyLines)
                $ok++
            } catch {
                $err++; [void]$errDetalls.Add("GIA $($r.Id): $($_.Exception.Message)")
            }
        }
    } finally {
        $cancel.Running = $false
        try { $form.Close() } catch { }
        Close-WordApp $word
    }

    $titol = if ($cancelled) { 'Generació cancel·lada' } else { 'Generació completada' }
    $dir = _ResolveOutputDir
    $msg = "$titol`n`nInformes generats: $ok`nErrors: $err`n`nCarpeta de sortida:`n$dir"
    if ($errDetalls.Count -gt 0) { $msg += "`n`nDetall d'errors:`n - " + (($errDetalls | Select-Object -First 10) -join "`n - ") }
    $r2 = [System.Windows.Forms.MessageBox]::Show($msg + "`n`nVols obrir la carpeta de sortida?", 'Generar informes', 'YesNo', 'Information')
    if ($r2 -eq [System.Windows.Forms.DialogResult]::Yes) {
        try { Start-Process -FilePath $dir | Out-Null } catch { }
    }
}
