#requires -Version 5.1
<#
.SYNOPSIS
  Eina "Pla verificacio" (menu, fila BASE D'INFORMES): els indicadors de la
  Memoria anual del Pla de verificacio d'activitats comunicades, en el format
  del programa, per a qualsevol any des del 2025.

.DESCRIPTION
  L'usuari (octubre 2026), amb el Programa de verificacio 2026 i la Memoria
  del 2025 al davant. Els indicadors, tal com els va definir:

    AVALUACIO DE L'EXECUCIO DEL PLA
    - Verificacions previstes al Programa anual: 30 el 2025, 45 el 2026 (el
      programa; d'un any nou es posa a la finestra i es desa en aquest PC).
    - Verificacions realitzades (programa de verificacio): activitats amb
      'Data control inicial/verificacio' d'aquell any i classificacio general
      annex que NO es II ni III.
    - Grau d'execucio: realitzades / previstes (%).
    - Verificacions realitzades per denuncies: activitats amb data
      d''Inspeccio' d'aquell any.
    - Verificacions realitzades per una EC-PCA (III): 'Data control
      inicial/verificacio' d'aquell any i annex III.
    AVALUACIO DE LES ACTIVITATS VERIFICADES
    - Activitats amb discrepancies greus entre allo comunicat i la realitat:
      de les del programa, les que tenen a la base d'informes un informe (no
      ignorat) amb la conclusio breu 'Precinte / Cessament' fet el dia del
      control o despres.
    - Percentatge respecte de les verificades: greus / realitzades (%).

  Les dades son les de l'Excel d'activitats mes nou (fulla Estes) i la base
  d'informes (informes-db.json). L'Excel nomes guarda la DARRERA data de
  control de cada activitat: una activitat verificada el 2025 i tornada a
  verificar el 2026 ja nomes compta el 2026.

  Aqui hi ha el que no es pantalla (la finestra es a PlaVerificacioPantalla.ps1).
  NOMES DEFINEIX FUNCIONS.
#>

# El primer any amb programa ("abans del 2025 no hi ha programa de
# verificacio") i les verificacions previstes de cada programa signat.
$Script:PvPrimerAny = 2025
$Script:PvPrevistesDefecte = @{ 2025 = 30; 2026 = 45 }

# Els titols i els textos dels indicadors, tal com surten a la Memoria.
$Script:PvTitolExecucio = ('AVALUACI' + [char]0x00D3 + ' DE L' + [char]0x2019 + 'EXECUCI' + [char]0x00D3 + ' DEL PLA')
$Script:PvTitolActivitats = ('AVALUACI' + [char]0x00D3 + ' DE LES ACTIVITATS VERIFICADES')
$Script:PvIndicadorsExecucio = [ordered]@{
    Previstes = 'Nombre de verificacions previstes al Programa anual'
    Programa  = ('Nombre de verificacions realitzades (programa de verificaci' + [char]0x00F3 + ')')
    Grau      = ('Grau d' + [char]0x2019 + 'execuci' + [char]0x00F3 + ' del Programa anual')
    Denuncies = ('Nombre de verificacions realitzades per den' + [char]0x00FA + 'ncies')
    EcPca     = 'Nombre de verificacions realitzades per una EC-PCA (III)'
}
$Script:PvIndicadorsActivitats = [ordered]@{
    Greus    = ('Nombre d' + [char]0x2019 + 'activitats amb discrep' + [char]0x00E0 + 'ncies greus entre all' + [char]0x00F2 + ' comunicat i la realitat')
    PctGreus = ('Percentatge en relaci' + [char]0x00F3 + ' amb el nombre d' + [char]0x2019 + 'activitats verificades')
}

# Els grups de la llista d'activitats (la de sota de les taules i l'Excel).
$Script:PvGrups = [ordered]@{
    programa  = ('Programa de verificaci' + [char]0x00F3)
    ecpca     = 'EC-PCA (III)'
    denuncies = ('Den' + [char]0x00FA + 'ncia')
}

# ----------------------------------------------------------------------------
# FUNCIONS PURES
# ----------------------------------------------------------------------------
# L'annex II o III de la 'Classificacio general annex' ('II', 'III'; amb
# espais o en minuscules tambe). Torna 'II', 'III' o ''. PURA.
function _PvAnnexIIoIII([string]$annex) {
    $a = ([string]$annex).Trim().ToUpperInvariant()
    if ($a -eq 'II' -or $a -eq 'III') { return $a }
    return ''
}

# Un percentatge com el de la Memoria: '33,3%', '0%', '100%'. Sense
# denominador, '-'. PURA.
function _PvPercent([int]$num, [int]$den) {
    if ($den -le 0) { return '-' }
    $p = [math]::Round((100.0 * $num) / $den, 1)
    if ($p -eq [math]::Floor($p)) { return ([string][int]$p + '%') }
    return ($p.ToString('0.0', [System.Globalization.CultureInfo]::InvariantCulture).Replace('.', ',') + '%')
}

# Les verificacions previstes d'un any: la de la finestra (aquest PC) si n'hi
# ha, si no la del programa signat; 0 si no se'n sap res. PURA.
function _PvPrevistes([int]$any, $locals) {
    $k = [string]$any
    if ($null -ne $locals -and $locals.Contains($k) -and [int]$locals[$k] -gt 0) { return [int]$locals[$k] }
    if ($Script:PvPrevistesDefecte.ContainsKey($any)) { return [int]$Script:PvPrevistesDefecte[$any] }
    return 0
}

# Els anys que es poden triar: del primer programa a l'any d'avui (o al
# darrer amb dades, si l'Excel en porta de mes endavant). PURA.
function _PvAnys([datetime]$avui, $files) {
    $fi = $avui.Year
    foreach ($f in @($files)) {
        foreach ($d in @($f.Control, $f.Inspeccio)) { if ($d -is [datetime] -and $d.Year -gt $fi) { $fi = $d.Year } }
    }
    $out = New-Object System.Collections.ArrayList
    for ($a = $Script:PvPrimerAny; $a -le $fi; $a++) { [void]$out.Add($a) }
    return ,($out.ToArray())
}

# Els informes de la base, per GIA: @{ gia -> @(@{ Data [datetime]; Breu }) },
# sense els ignorats ni els que no tenen data. PURA.
function _PvInformesPerGia($db) {
    $out = @{}
    if ($null -eq $db) { return $out }
    foreach ($act in @(_PropInf $db 'activitats')) {
        if ($null -eq $act) { continue }
        $gia = ([string](_PropInf $act 'id_gia')).Trim()
        if (-not $gia) { continue }
        foreach ($inf in @(_PropInf $act 'informes')) {
            if ($null -eq $inf -or [bool](_PropInf $inf 'ignorat')) { continue }
            $s = [string](_PropInf $inf 'data')
            $dt = [datetime]::MinValue
            if ($s -notmatch '^(\d{4}-\d{2}-\d{2})' -or -not [datetime]::TryParseExact($Matches[1], 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) { continue }
            if (-not $out.ContainsKey($gia)) { $out[$gia] = New-Object System.Collections.ArrayList }
            [void]$out[$gia].Add(@{ Data = $dt; Breu = [string](_PropInf $inf 'conclusio_breu'); Fitxer = [string](_PropInf $inf 'fitxer') })
        }
    }
    return $out
}

# L'informe de precinte d'una verificacio: el primer informe (no ignorat) amb
# la conclusio breu 'Precinte / Cessament' fet el dia del control o despres.
# Torna el registre o $null. PURA.
function _PvInformePrecinte($informes, [datetime]$control) {
    $millor = $null
    foreach ($i in @($informes)) {
        if ($null -eq $i -or [string]$i.Breu -ne $Script:EstatPrecinte) { continue }
        if (([datetime]$i.Data).Date -lt $control.Date) { continue }
        if ($null -eq $millor -or ([datetime]$i.Data) -lt ([datetime]$millor.Data)) { $millor = $i }
    }
    return $millor
}

# ELS INDICADORS D'UN ANY. $files: les de l'Excel (@{ Id; Annex; Control;
# Inspeccio; Activitat; Via; Carrer; Numero }, les dates [datetime] o $null);
# $informes: _PvInformesPerGia ($null = sense base d'informes: les greus no es
# poden saber). Torna @{ Any; Previstes; Programa; Grau; Denuncies; EcPca;
# Greus; PctGreus; SenseBase; Llista (una fila per activitat i grup) }. PURA.
function Get-PlaVerificacioIndicadors($files, [int]$any, [int]$previstes, $informes) {
    $llista = New-Object System.Collections.ArrayList
    $nProg = 0; $nDen = 0; $nEc = 0; $nGreus = 0
    foreach ($f in @($files)) {
        if ($null -eq $f) { continue }
        $ctl = $f.Control
        if ($ctl -is [datetime] -and $ctl.Year -eq $any) {
            $ii = _PvAnnexIIoIII ([string]$f.Annex)
            if ($ii -eq '') {
                $nProg++
                $prec = $null
                if ($null -ne $informes -and $informes.ContainsKey([string]$f.Id)) { $prec = _PvInformePrecinte $informes[[string]$f.Id] $ctl }
                if ($null -ne $prec) { $nGreus++ }
                [void]$llista.Add(@{ Grup = 'programa'; Fila = $f; Data = $ctl; Precinte = $prec })
            } elseif ($ii -eq 'III') {
                $nEc++
                [void]$llista.Add(@{ Grup = 'ecpca'; Fila = $f; Data = $ctl; Precinte = $null })
            }
        }
        $ins = $f.Inspeccio
        if ($ins -is [datetime] -and $ins.Year -eq $any) {
            $nDen++
            [void]$llista.Add(@{ Grup = 'denuncies'; Fila = $f; Data = $ins; Precinte = $null })
        }
    }
    $sense = ($null -eq $informes)
    return @{
        Any = $any; Previstes = $previstes
        Programa = $nProg; Grau = (_PvPercent $nProg $previstes)
        Denuncies = $nDen; EcPca = $nEc
        Greus = $(if ($sense) { $null } else { $nGreus })
        PctGreus = $(if ($sense) { '-' } else { _PvPercent $nGreus $nProg })
        SenseBase = $sense
        Llista = @(_PvOrdenaLlista $llista.ToArray())
    }
}

# La llista, per grup (programa, EC-PCA, denuncies) i dins de cada un per
# data. PURA.
function _PvOrdenaLlista($llista) {
    $ordre = @($Script:PvGrups.Keys)
    return @(@($llista) | Sort-Object -Property @{ Expression = { [array]::IndexOf($ordre, [string]$_.Grup) } }, @{ Expression = { [datetime]$_.Data } }, @{ Expression = { [int]([string]$_.Fila.Id -replace '\D', '0') } })
}

# Les files de les dues taules, en text: @(@(indicador, valor)). PURA.
function _PvFilesExecucio($ind) {
    $n = $Script:PvIndicadorsExecucio
    return @(
        , @($n.Previstes, $(if ([int]$ind.Previstes -gt 0) { [string]$ind.Previstes } else { '-' }))
        , @($n.Programa, [string]$ind.Programa)
        , @($n.Grau, [string]$ind.Grau)
        , @($n.Denuncies, [string]$ind.Denuncies)
        , @($n.EcPca, [string]$ind.EcPca)
    )
}
function _PvFilesActivitats($ind) {
    $n = $Script:PvIndicadorsActivitats
    return @(
        , @($n.Greus, $(if ($null -eq $ind.Greus) { '-' } else { [string]$ind.Greus }))
        , @($n.PctGreus, [string]$ind.PctGreus)
    )
}

# Una fila de la llista d'activitats, en text (la graella i l'Excel). PURA.
$Script:PvLlistaCapcalera = @('Indicador', 'ID Activitat', 'Data', 'Emp. Tipus via', 'Emp. Carrer', ('Emp. N' + [char]0x00FA + 'mero'), 'Activitat principal', ('Classificaci' + [char]0x00F3 + ' general annex'), 'Precinte (informe)')
$Script:PvLlistaAmples = [int[]]@(26, 12, 12, 10, 28, 10, 40, 22, 40)
function _PvFilaLlista($e) {
    $f = $e.Fila
    $prec = if ($null -ne $e.Precinte) { ([datetime]$e.Precinte.Data).ToString('dd/MM/yyyy') + ' ' + [string]$e.Precinte.Fitxer } else { '' }
    return @([string]$Script:PvGrups[[string]$e.Grup], [string]$f.Id, ([datetime]$e.Data).ToString('dd/MM/yyyy'), [string]$f.Via, [string]$f.Carrer, [string]$f.Numero, [string]$f.Activitat, [string]$f.Annex, $prec)
}

# ----------------------------------------------------------------------------
# LES TAULES PER ENGANXAR AL WORD (el format de la Memoria)
# ----------------------------------------------------------------------------
# Tot el que no es ASCII, com a entitat: aixi les posicions del CF_HTML (en
# bytes) son les mateixes que en caracters, i cap Word no s'equivoca de
# codificacio. PURA.
function _PvHtmlEsc([string]$t) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in ([string]$t).ToCharArray()) {
        switch ($ch) {
            '&' { [void]$sb.Append('&amp;') }
            '<' { [void]$sb.Append('&lt;') }
            '>' { [void]$sb.Append('&gt;') }
            '"' { [void]$sb.Append('&quot;') }
            default { if ([int]$ch -gt 126) { [void]$sb.Append('&#' + [int]$ch + ';') } else { [void]$sb.Append($ch) } }
        }
    }
    return $sb.ToString()
}

# Una taula com les de la Memoria: titol a dalt (gris), capcalera
# (Indicador | $col), indicadors en cursiva i valors centrats. PURA.
function _PvTaulaHtml([string]$titol, [string]$col, $files) {
    $vora = 'border:1px solid #000;'
    $gris = 'background:#E6E6E6;'
    $lletra = "font-family:'Bookman Old Style',serif;font-size:11pt;"
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<table style="border-collapse:collapse;width:16cm;' + $lletra + '">')
    [void]$sb.Append('<tr><td colspan="2" style="' + $vora + $gris + 'text-align:center;font-weight:bold;padding:6px;">' + (_PvHtmlEsc $titol) + '</td></tr>')
    [void]$sb.Append('<tr><td style="' + $vora + $gris + 'text-align:center;font-weight:bold;padding:4px;">Indicador</td><td style="' + $vora + $gris + 'text-align:center;font-weight:bold;padding:4px;width:3.2cm;">' + (_PvHtmlEsc $col) + '</td></tr>')
    foreach ($f in @($files)) {
        [void]$sb.Append('<tr><td style="' + $vora + 'font-style:italic;padding:4px;">' + (_PvHtmlEsc ([string]$f[0])) + '</td><td style="' + $vora + 'text-align:center;padding:4px;">' + (_PvHtmlEsc ([string]$f[1])) + '</td></tr>')
    }
    [void]$sb.Append('</table>')
    return $sb.ToString()
}

function _PvTaulesHtml($ind) {
    return ((_PvTaulaHtml $Script:PvTitolExecucio 'Seguiment' (_PvFilesExecucio $ind)) + '<p>&nbsp;</p>' + (_PvTaulaHtml $Script:PvTitolActivitats 'Resultat' (_PvFilesActivitats $ind)))
}

# El mateix en text (per a on no s'enganxa HTML). PURA.
function _PvTaulesText($ind) {
    $nl = "`r`n"
    $t = $Script:PvTitolExecucio + $nl + "Indicador`tSeguiment" + $nl
    foreach ($f in @(_PvFilesExecucio $ind)) { $t += [string]$f[0] + "`t" + [string]$f[1] + $nl }
    $t += $nl + $Script:PvTitolActivitats + $nl + "Indicador`tResultat" + $nl
    foreach ($f in @(_PvFilesActivitats $ind)) { $t += [string]$f[0] + "`t" + [string]$f[1] + $nl }
    return $t
}

# El format del porta-retalls "HTML Format" (CF_HTML): una capcalera amb les
# posicions, en bytes, d'on comenca i acaba el document i el fragment. L'HTML
# ha de ser ASCII (_PvHtmlEsc). PURA.
function _PvCfHtml([string]$fragment) {
    $pre = '<html><body><!--StartFragment-->'
    $post = '<!--EndFragment--></body></html>'
    $plantilla = "Version:0.9`r`nStartHTML:{0:D10}`r`nEndHTML:{1:D10}`r`nStartFragment:{2:D10}`r`nEndFragment:{3:D10}`r`n"
    $lenCap = ($plantilla -f 0, 0, 0, 0).Length
    $iniHtml = $lenCap
    $iniFrag = $iniHtml + $pre.Length
    $fiFrag = $iniFrag + $fragment.Length
    $fiHtml = $fiFrag + $post.Length
    return (($plantilla -f $iniHtml, $fiHtml, $iniFrag, $fiFrag) + $pre + $fragment + $post)
}

# ----------------------------------------------------------------------------
# LES PREVISTES D'UN ANY NOU (aquest PC): local\base-dades-activitats
# ----------------------------------------------------------------------------
function _PvPath {
    if ([string]::IsNullOrWhiteSpace($LocalActivitatsDir)) { return '' }
    return [string](Join-Path $LocalActivitatsDir 'pla-verificacio.json')
}

# @{ 'AAAA' -> n }. Mai peta.
function Read-PvPrevistes([string]$path = '') {
    if (-not $path) { $path = _PvPath }
    $out = @{}
    if (-not $path) { return $out }
    $o = $null
    try { $o = Read-JsonFile $path } catch { $o = $null }
    $p = _PropInf $o 'previstes'
    if ($null -ne $p) {
        foreach ($pr in @($p.PSObject.Properties)) { $n = 0; if ([int]::TryParse([string]$pr.Value, [ref]$n) -and $n -gt 0) { $out[[string]$pr.Name] = $n } }
    }
    return $out
}

function Save-PvPrevistes($previstes, [string]$path = '') {
    if (-not $path) { $path = _PvPath }
    if (-not $path) { throw "No hi ha carpeta local per desar les verificacions previstes." }
    $dir = Split-Path -Parent $path
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $o = [ordered]@{}
    foreach ($k in @(@($previstes.Keys) | Sort-Object)) { $o[[string]$k] = [int]$previstes[$k] }
    Write-JsonFile $path ([ordered]@{ previstes = $o }) 3
}

# ----------------------------------------------------------------------------
# LA LECTURA DE L'EXCEL (fulla Estes)
# ----------------------------------------------------------------------------
# Les columnes que calen. 'Inspeccio' ha de ser EXACTA: n'hi ha d'altres que
# porten la paraula. Torna @{ ID; ANNEX; CONTROL; INSPECCIO; ACT; VIA; CARRER;
# NUMERO } (0 = no hi es). PURA.
function _PvColumnes($headers) {
    $h = @($headers)
    $insp = 0
    for ($i = 0; $i -lt $h.Count; $i++) { if ((_NormalitzaText ([string]$h[$i])).Trim() -eq 'inspeccio') { $insp = $i + 1; break } }
    $col = @{
        ID        = (_ColumnaAmbTermes $h @('id activitat') $null)
        ANNEX     = (_ColumnaAmbTermes $h @('classificacio general annex') $null)
        CONTROL   = (_ColumnaAmbTermes $h @('data control inicial') $null)
        INSPECCIO = $insp
        ACT       = (_ColumnaAmbTermes $h @('activitat principal') $null)
        VIA       = (_ColumnaAmbTermes $h @('emp', 'tipus via') $null)
        CARRER    = (_ColumnaAmbTermes $h @('emp', 'carrer') $null)
        NUMERO    = (_ColumnaAmbTermes $h @('emp', 'numero') $null)
    }
    if ($col.ID -eq 0 -and $h.Count -ge 1) { $col.ID = 1 }
    return $col
}

# Les columnes sense les quals no es pot calcular res, pel seu nom. PURA.
function _PvColumnesQueFalten($col) {
    $noms = [ordered]@{ ANNEX = ('Classificaci' + [char]0x00F3 + ' general annex'); CONTROL = ('Data control inicial/verificaci' + [char]0x00F3); INSPECCIO = ('Inspecci' + [char]0x00F3) }
    return @(@($noms.Keys) | Where-Object { [int]$col[$_] -le 0 } | ForEach-Object { $noms[$_] })
}

# Llegeix l'Excel d'activitats mes nou. Torna @{ Ok; Files; Col; File; Error }.
function _ReadPlaVerificacio {
    $latest = Find-LatestActivitatsExcel
    if ($null -eq $latest) { return @{ Ok = $false; Error = ("No s'ha trobat cap Excel d'activitats (" + (_ExcelBaseNomText) + ").") } }
    try {
        $lectura = Read-FullaEstesa $latest.File {
            param($x)
            $data = $x.Data
            if ($null -eq $data) { return @{ Ok = $true; Files = @(); Col = (_PvColumnes @()); File = $latest.File.Name } }
            $col = _PvColumnes $x.Headers
            $get = $x.Cel
            $files = New-Object System.Collections.ArrayList
            for ($ri = 2; $ri -le $x.Rows; $ri++) {
                $id = Get-IdDeCella $data[$ri, $col.ID]
                if ([string]::IsNullOrWhiteSpace($id)) { continue }
                $txt = { param($k) if ([int]$col[$k] -gt 0) { ([string](& $get $ri $col[$k])).Trim() } else { '' } }
                [void]$files.Add(@{
                    Id        = $id
                    Annex     = (& $txt 'ANNEX')
                    Control   = $(if ($col.CONTROL -gt 0) { _DataDeCella $data[$ri, $col.CONTROL] } else { $null })
                    Inspeccio = $(if ($col.INSPECCIO -gt 0) { _DataDeCella $data[$ri, $col.INSPECCIO] } else { $null })
                    Activitat = (& $txt 'ACT')
                    Via       = (& $txt 'VIA')
                    Carrer    = (& $txt 'CARRER')
                    Numero    = (& $txt 'NUMERO')
                })
            }
            return @{ Ok = $true; Files = $files.ToArray(); Col = $col; File = $latest.File.Name }
        }
        return $lectura
    } catch {
        return @{ Ok = $false; Error = $_.Exception.Message }
    }
}

# La base d'informes, per GIA; $null si no n'hi ha o no es pot llegir.
function _ReadPvInformes {
    $p = Get-InformesDbPath
    if (-not $p -or -not (Test-Path -LiteralPath $p)) { return $null }
    try { return (_PvInformesPerGia (Read-JsonFile $p)) } catch { return $null }
}
