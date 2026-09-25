#requires -Version 5.1
<#
.SYNOPSIS
  Eina "Comprovar Excel" (menu, fila GIA): per cada activitat en estat
  "Precinte / Cessament" a la base d'informes, comprova que l'Excel d'activitats
  hi tingui el Camp Info corresponent amb un valor que comenci per SI.
  Criteri DIFERENT del de Seguiment (SeguimentGia.ps1), a posta: vegeu CLAUDE.md.
  Abans vivia a Informes.ps1 (revisio d'arquitectura, setembre 2026).
#>

# ----------------------------------------------------------------------------
# Comprovar Excel: noms de "Camp Info" que, a l'Excel d'activitats, marquen que
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
# Comprovar Excel — eina INFORMES "Comprovar Excel"
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
            if ($null -eq $data) { return @{ Ok = $true; Map = @{} } }
            $pairs = _FindCampInfoPairs $headers
            $map = @{}
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
            }
            return @{ Ok = $true; Map = $map }
        })
    } catch {
        return @{ Ok = $false; Error = $_.Exception.Message }
    }
}

# Finestra modal senzilla amb un llistat de resultats (només lectura).
function _ShowResultatWindow($titol, $subtitol, $text) {
    $form = _NewForm
    $form.Text = $titol
    $form.Size = New-Object System.Drawing.Size(680, 560)
    $form.MinimumSize = New-Object System.Drawing.Size(480, 360)
    $form.StartPosition = 'CenterScreen'

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Multiline = $true
    $tb.ReadOnly = $true
    $tb.ScrollBars = 'Vertical'
    $tb.Dock = 'Fill'
    $tb.Font = New-Object System.Drawing.Font('Consolas', 9.5)
    $tb.BackColor = [System.Drawing.Color]::White
    $tb.Text = $text

    $bot = New-Object System.Windows.Forms.Panel
    $bot.Dock = 'Bottom'; $bot.Height = 46
    [void](_AddPeuBotons $form @(@{ Nom = 'Tancar'; Text = 'Tancar'; Clic = { $form.Close() }.GetNewClosure() }) @() 7 $bot)

    $form.Controls.Add($tb)
    $form.Controls.Add($bot)
    [void](_AddBrandHeader $form $titol $subtitol 56)
    [void]$form.ShowDialog()
}

# Comprova que les activitats en Estat "Precinte / Cessament" (base d'informes)
# tinguin a l'Excel el Camp Info corresponent amb valor "SI". Llista les que no.
#
# Aqui hi havia un _SaveRunTimestamp que escrivia 'comprovat_el' a
# comprovar-excel-state.json. Es va treure: l'unica cosa que en feia servei era
# ensenyar el segell d'"ultima execucio" al menu, i ara aixo ho porta un SOL
# registre per a totes les eines (_MarcaEinaUsada, Seguiment.ps1), apuntat des
# del despatxador. El fitxer d'estat d'aquesta eina ja no fa falta.
function Invoke-ComprovarExcel {
    $outPath = Join-Path $LocalActivitatsDir 'informes-db.json'
    if (-not (Test-Path -LiteralPath $outPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Encara no hi ha cap base d'informes.`n`nExecuta primer 'Actualitzar base'.",
            'Comprovar Excel', 'OK', 'Information') | Out-Null
        return
    }
    try {
        $db = Read-JsonFile $outPath
    } catch {
        [System.Windows.Forms.MessageBox]::Show("No s'ha pogut llegir la base d'informes:`n$($_.Exception.Message)", 'Comprovar Excel', 'OK', 'Error') | Out-Null
        return
    }
    $targets = @($db.activitats | Where-Object { [string]$_.estat_actual -eq 'Precinte / Cessament' })
    if ($targets.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No hi ha cap activitat en Estat 'Precinte / Cessament' a la base d'informes.", 'Comprovar Excel', 'OK', 'Information') | Out-Null
        return
    }

    $res = _ReadExcelCampInfoPerGia
    if (-not $res.Ok) {
        [System.Windows.Forms.MessageBox]::Show("No s'ha pogut llegir l'Excel:`n$($res.Error)", 'Comprovar Excel', 'OK', 'Error') | Out-Null
        return
    }
    $map = $res.Map

    $desact = New-Object System.Collections.ArrayList
    $noTrob = New-Object System.Collections.ArrayList
    $senseGia = New-Object System.Collections.ArrayList
    foreach ($act in $targets) {
        $gia = [string]$act.id_gia
        if ([string]::IsNullOrWhiteSpace($gia)) {
            [void]$senseGia.Add("     - " + [string]$act.titular + "  (carpeta: " + [string]$act.carpeta + ")")
            continue
        }
        $etiqueta = "GIA " + $gia + "  -  " + [string]$act.titular
        if (-not $map.ContainsKey($gia)) { [void]$noTrob.Add("     - " + $etiqueta); continue }
        if (-not (_ExcelActivitatActualitzada $map[$gia])) {
            # Hi afegim la data de l'informe que va deixar l'activitat en
            # "Precinte / Cessament": es el que has de citar per actualitzar
            # l'Excel, i sense aixo tocava anar a buscar-lo a ma.
            $infEstat = _InformeQueDeterminaEstat $act.informes
            $dataEstat = if ($null -ne $infEstat) { _DataInformeDdMmAaaa $infEstat.data } else { '' }
            if ($dataEstat) { $etiqueta += " - INFORME ENGINYER " + $dataEstat }
            [void]$desact.Add("     - " + $etiqueta)
            continue
        }
    }

    if ($desact.Count -eq 0 -and $noTrob.Count -eq 0 -and $senseGia.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            ("L'Excel està al dia.`n`nTotes les {0} activitats en Estat 'Precinte / Cessament' tenen a l'Excel un Camp Info {1} amb valor SI." -f $targets.Count, (_ExcelPrecinteCampsText)),
            'Comprovar Excel', 'OK', 'Information') | Out-Null
        return
    }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine(("Activitats en Estat 'Precinte / Cessament' a la base d'informes: {0}" -f $targets.Count))
    [void]$sb.AppendLine(("Criteri: a l'Excel han de tenir un Camp Info {0} amb valor que comenci per SI." -f (_ExcelPrecinteCampsText)))
    [void]$sb.AppendLine("")
    if ($desact.Count -gt 0) {
        [void]$sb.AppendLine(("DESACTUALITZADES a l'Excel (sense el Camp Info amb SI): {0}" -f $desact.Count))
        foreach ($l in $desact) { [void]$sb.AppendLine($l) }
        [void]$sb.AppendLine("")
    }
    if ($noTrob.Count -gt 0) {
        [void]$sb.AppendLine(("NO trobades a l'Excel (GIA inexistent a la fulla Estès): {0}" -f $noTrob.Count))
        foreach ($l in $noTrob) { [void]$sb.AppendLine($l) }
        [void]$sb.AppendLine("")
    }
    if ($senseGia.Count -gt 0) {
        [void]$sb.AppendLine(("NO verificables (activitat sense GIA a la base d'informes): {0}" -f $senseGia.Count))
        foreach ($l in $senseGia) { [void]$sb.AppendLine($l) }
        [void]$sb.AppendLine("")
    }

    # DIAGNOSTIC: quins camps d'aquesta familia hi ha DE DEBO a l'Excel. Si algun
    # dia es reanomenen (va passar: 'PRECINTE?' -> 'PRECINTE ACTIVITAT?') el
    # criteri deixa de trobar-los i tot surt com a desactualitzat; aixi es veu de
    # seguida en lloc d'haver-ho d'endevinar.
    $campsExcel = @(_ExcelCampsPrecinteTrobats $map)
    if ($campsExcel.Count -gt 0) {
        [void]$sb.AppendLine("Camps de l'Excel que parlen de precinte o decret (per si s'han reanomenat):")
        foreach ($c in $campsExcel) {
            $marca = if ($Script:ExcelPrecinteCampNoms -contains (_NormalitzaText $c)) { '  <-- es el que es comprova' } else { '' }
            [void]$sb.AppendLine("     - " + $c + $marca)
        }
    }

    _ShowResultatWindow "Comprovació de l'Excel" "Activitats precintades pendents d'actualitzar a l'Excel" ($sb.ToString())
}
