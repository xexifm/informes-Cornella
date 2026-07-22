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
  primer). Reutilitza helpers de GenerarInforme.ps1 (_FindEstesSheet,
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
    $excel = $null
    try { $excel = New-Object -ComObject Excel.Application } catch { $excel = $null }
    if ($null -eq $excel) { return @{ Ok = $false; Error = "No s'ha pogut iniciar Microsoft Excel." } }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
        $wb = $excel.Workbooks.Open($latest.File.FullName, 0, $true)   # ReadOnly
        try {
            $found = _FindEstesSheet $wb
            $sh = $found.Sheet
            if ($null -eq $sh) { return @{ Ok = $false; Error = "No s'ha trobat la fulla 'Estès' a l'Excel." } }
            $data = $sh.UsedRange.Value2
            if ($null -eq $data) { return @{ Ok = $true; Rows = @(); File = $latest.File.Name } }
            $rows = $data.GetLength(0)
            $cols = $data.GetLength(1)

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

            if ($cAnnex -eq 0 -and $cApart -eq 0) {
                return @{ Ok = $false; Error = "No s'han trobat les columnes 'Classificació general annex' ni 'Classificació general Apartat' a la fulla Estès." }
            }

            $get = {
                param($r, $c)
                if ($c -lt 1 -or $c -gt $cols) { return '' }
                $v = $data[$r, $c]
                if ($null -eq $v) { return '' }
                return ([string]$v).Trim()
            }
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
                })
            }
            return @{ Ok = $true; Rows = @($out); File = $latest.File.Name }
        } finally {
            try { $wb.Close($false) } catch { }
        }
    } catch {
        return @{ Ok = $false; Error = $_.Exception.Message }
    } finally {
        try { $excel.Quit() } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
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

    # Ordre per defecte: "Proper CP previst" ascendent (el mes antic primer).
    $defaultSort = 0
    for ($i = 0; $i -lt $colsMeta.Count; $i++) { if ($colsMeta[$i].Key -eq 'ProperCP') { $defaultSort = $i; break } }
    $state = @{ Loading = $false; ColIdx = $defaultSort; Asc = $true }

    $form = _NewForm
    $form.Text = 'Controls periodics'
    $form.Size = New-Object System.Drawing.Size(1320, 720)
    $form.MinimumSize = New-Object System.Drawing.Size(820, 480)
    $form.StartPosition = 'CenterScreen'

    # ---- Graella --------------------------------------------------------
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = 'Fill'
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.RowHeadersVisible = $false
    $grid.ReadOnly = $true
    $grid.SelectionMode = 'FullRowSelect'
    $grid.MultiSelect = $false
    $grid.AutoSizeColumnsMode = 'None'
    $grid.BackgroundColor = [System.Drawing.Color]::White
    foreach ($cm in $colsMeta) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.HeaderText = $cm.H
        $col.Width = $cm.W
        $col.SortMode = 'Programmatic'
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
    $lblC = New-Object System.Windows.Forms.Label
    $lblC.Text = 'Cerca:'; $lblC.AutoSize = $true
    $lblC.Location = New-Object System.Drawing.Point(216, 13)
    $topPanel.Controls.Add($lblC)
    $txtCerca = New-Object System.Windows.Forms.TextBox
    $txtCerca.Location = New-Object System.Drawing.Point(266, 10); $txtCerca.Width = 300
    $topPanel.Controls.Add($txtCerca)
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
            if ($n -ne '') {
                $hay = (($row.Id + ' ' + $row.RaoSocial + ' ' + $row.RaoEmail + ' ' + $row.RepEmail + ' ' + $row.Adreca + ' ' + $row.PeriodicitatCP + ' ' + $row.Annex + ' ' + $row.Apartat + ' ' + $row.ActPrincipal)).ToLower()
                if (-not $hay.Contains($n)) { continue }
            }
            $row
        }
        $sel = @($sel)

        # Adjuntem una clau d'ordenacio a cada fila i ordenem per nom de
        # propietat (evita problemes d'abast amb scriptblocks dins de Sort-Object).
        $cm = $colsMeta[$state.ColIdx]
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
        $sorted = @($sel | Sort-Object -Property '_SortKey' -Descending:(-not $state.Asc))

        foreach ($row in $sorted) {
            $idx = $grid.Rows.Add(@(
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

    $grid.add_ColumnHeaderMouseClick({
        param($s, $e)
        if ($e.ColumnIndex -lt 0) { return }
        if ($state.ColIdx -eq $e.ColumnIndex) { $state.Asc = (-not $state.Asc) }
        else { $state.ColIdx = $e.ColumnIndex; $state.Asc = $true }
        foreach ($c in $grid.Columns) { $c.HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::None }
        $grid.Columns[$e.ColumnIndex].HeaderCell.SortGlyphDirection =
            if ($state.Asc) { [System.Windows.Forms.SortOrder]::Ascending } else { [System.Windows.Forms.SortOrder]::Descending }
        & $fill
    }.GetNewClosure())

    # ---- Barra inferior: Exportar / Tancar ------------------------------
    $botPanel = New-Object System.Windows.Forms.Panel
    $botPanel.Dock = 'Bottom'; $botPanel.Height = 48
    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = 'Exportar (CSV)'; $btnExport.Size = New-Object System.Drawing.Size(150, 30)
    $btnExport.Location = New-Object System.Drawing.Point(10, 9)
    _StyleSecondaryButton $btnExport
    $btnExport.add_Click({ _ExportControlsPeriodics $grid $colsMeta }.GetNewClosure())
    $botPanel.Controls.Add($btnExport)
    $btnTancar = New-Object System.Windows.Forms.Button
    $btnTancar.Text = 'Tancar'; $btnTancar.Size = New-Object System.Drawing.Size(120, 30)
    $btnTancar.Location = New-Object System.Drawing.Point(170, 9)
    _StyleSecondaryButton $btnTancar
    $btnTancar.add_Click({ $form.Close() }.GetNewClosure())
    $botPanel.Controls.Add($btnTancar)

    # Ordre d'afegit: Fill (centre), Top, Bottom, i la banda l'ultima (a dalt).
    $form.Controls.Add($grid)
    $form.Controls.Add($topPanel)
    $form.Controls.Add($botPanel)
    $sub = "Activitats amb control periòdic (annex II/III o apartat 561)  ·  $fileName"
    [void](_AddBrandHeader $form 'Controls periòdics' $sub 56)

    # Ordre inicial: per Data llicència ascendent (mes antic primer).
    $grid.Columns[$state.ColIdx].HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::Ascending
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
        for ($c = 0; $c -lt $colsMeta.Count; $c++) {
            $o[$colsMeta[$c].H] = [string]$gr.Cells[$c].Value
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
