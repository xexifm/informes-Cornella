#requires -Version 5.1
<#
  PlanolGeometria.ps1 - La geometria del Planol activitats: on va l'ID GIA a
  l'ENTRADA de l'establiment i com es junten les parcel.les d'una activitat.

  L'usuari (octubre 2026): "a les parcel.les vull que posis els numeros del ID
  GIA a l'entrada de l'establiment [...] Assegura't que queda dintre la
  parcel.la. En cas que no es trobi l'adreca o quedi fora la parcel.la posa el
  ID GIA com ho estaves fent ara pero escriu-lo en vermell. [...] quan estiguin
  molt a prop (gairebe tocant-se) intenta juntar-los en un mateix poligon."

  Tot en UTM 31N (metres), sobre els anells PLANS del Cadastre
  [x1, y1, x2, y2, ...] (ConvertFrom-CatastroParcelXml): Anells[0] l'exterior,
  la resta forats. Funcions PURES: es proven a Linux sense Cadastre.

  NOMES DEFINEIX FUNCIONS. ASCII pur.
#>

# Fins a quants metres FORA de la parcel.la s'accepta un portal: el Cadastre els
# posa a la facana (sobre la linia o una mica al carrer). Mes enlla, no es
# d'aquella parcel.la i l'ID va en vermell al centre.
$Script:PlanolEntradaMaxForaM = 12.0
# Quants metres CAP A DINS es mou l'etiqueta des de la vora: sobre la linia
# mateixa, mitja etiqueta quedaria al carrer.
$Script:PlanolEntradaMargeM = 2.5
# Dues entrades a menys d'aixo son la MATEIXA porta (una sola etiqueta).
$Script:PlanolMateixaPortaM = 3.0
# Dues entrades de la MATEIXA activitat a menys d'aixo, una sola etiqueta.
$Script:PlanolMateixaActivitatM = 10.0

# ----------------------------------------------------------------------------
# CENTRE I AREA (vivien a PlanolDades.ps1; son geometria i les fa servir
# Get-PuntInterior: alli feien un cicle de dependencies)
# ----------------------------------------------------------------------------
# Centre d'un anell (centroide d'area, formula del cordill); si l'area es nul.la,
# la mitjana dels vertexs. On va l'etiqueta dels ID GIA.
function Get-CentreAnell($anell) {
    $a = @($anell)
    $n = [int]($a.Count / 2)
    if ($n -eq 0) { return $null }
    $area = 0.0; $cx = 0.0; $cy = 0.0; $sx = 0.0; $sy = 0.0
    for ($i = 0; $i -lt $n; $i++) {
        $x0 = [double]$a[2 * $i]; $y0 = [double]$a[2 * $i + 1]
        $j = ($i + 1) % $n
        $x1 = [double]$a[2 * $j]; $y1 = [double]$a[2 * $j + 1]
        $f = $x0 * $y1 - $x1 * $y0
        $area += $f; $cx += ($x0 + $x1) * $f; $cy += ($y0 + $y1) * $f
        $sx += $x0; $sy += $y0
    }
    if ([math]::Abs($area) -lt 1e-9) { return @(($sx / $n), ($sy / $n)) }
    return @(($cx / (3 * $area)), ($cy / (3 * $area)))
}

# Area (absoluta) d'un anell, per triar el poligon mes gran d'una parcel.la.
function Get-AreaAnell($anell) {
    $a = @($anell); $n = [int]($a.Count / 2); $s = 0.0
    for ($i = 0; $i -lt $n; $i++) {
        $j = ($i + 1) % $n
        $s += [double]$a[2 * $i] * [double]$a[2 * $j + 1] - [double]$a[2 * $j] * [double]$a[2 * $i + 1]
    }
    return [math]::Abs($s / 2)
}

# ----------------------------------------------------------------------------
# PUNTS I ANELLS
# ----------------------------------------------------------------------------
# Els vertexs d'un anell pla, sense el darrer si repeteix el primer (el GML el
# tanca). Llista de [double[]] { x, y }.
function _PgPunts($anell) {
    $a = @($anell)
    $out = New-Object System.Collections.Generic.List[double[]]
    for ($i = 0; $i + 1 -lt $a.Count; $i += 2) { $out.Add([double[]]@([double]$a[$i], [double]$a[$i + 1])) }
    if ($out.Count -gt 1 -and $out[0][0] -eq $out[$out.Count - 1][0] -and $out[0][1] -eq $out[$out.Count - 1][1]) { $out.RemoveAt($out.Count - 1) }
    return ,$out
}

function _PgPla($punts) {
    $out = New-Object System.Collections.Generic.List[double]
    foreach ($p in $punts) { $out.Add($p[0]); $out.Add($p[1]) }
    return ,($out.ToArray())
}

# Area AMB SIGNE (positiva si l'anell va en sentit antihorari).
function _PgAreaSignada($punts) {
    $s = 0.0; $n = $punts.Count
    for ($i = 0; $i -lt $n; $i++) {
        $j = ($i + 1) % $n
        $s += $punts[$i][0] * $punts[$j][1] - $punts[$j][0] * $punts[$i][1]
    }
    return ($s / 2)
}

function Test-PuntDinsAnell([double]$x, [double]$y, $anell) {
    $p = _PgPunts $anell
    $dins = $false; $n = $p.Count
    for ($i = 0; $i -lt $n; $i++) {
        $j = if ($i -eq 0) { $n - 1 } else { $i - 1 }
        $xi = $p[$i][0]; $yi = $p[$i][1]; $xj = $p[$j][0]; $yj = $p[$j][1]
        if ((($yi -gt $y) -ne ($yj -gt $y)) -and ($x -lt (($xj - $xi) * ($y - $yi) / ($yj - $yi) + $xi))) { $dins = -not $dins }
    }
    return $dins
}

# Dins d'algun poligon: dins del seu exterior i fora dels seus forats.
function Test-PuntDinsPoligons([double]$x, [double]$y, $polys) {
    foreach ($pl in @($polys)) {
        $an = @($pl.Anells)
        if ($an.Count -eq 0 -or -not (Test-PuntDinsAnell $x $y $an[0])) { continue }
        $foradat = $false
        for ($k = 1; $k -lt $an.Count; $k++) { if (Test-PuntDinsAnell $x $y $an[$k]) { $foradat = $true; break } }
        if (-not $foradat) { return $true }
    }
    return $false
}

# El punt de la VORA (de qualsevol anell) mes proper a (x, y), la distancia i
# el costat on es: { X; Y; D; Ex; Ey } (Ex, Ey: el vector del costat).
function Get-VoraMesPropera([double]$x, [double]$y, $polys) {
    $millor = $null
    foreach ($pl in @($polys)) {
        foreach ($an in @($pl.Anells)) {
            $p = _PgPunts $an; $n = $p.Count
            for ($i = 0; $i -lt $n; $i++) {
                $a = $p[$i]; $b = $p[($i + 1) % $n]
                $ex = $b[0] - $a[0]; $ey = $b[1] - $a[1]
                $l2 = $ex * $ex + $ey * $ey
                $t = if ($l2 -gt 0) { [math]::Max(0.0, [math]::Min(1.0, (($x - $a[0]) * $ex + ($y - $a[1]) * $ey) / $l2)) } else { 0.0 }
                $qx = $a[0] + $t * $ex; $qy = $a[1] + $t * $ey
                $d = [math]::Sqrt(($x - $qx) * ($x - $qx) + ($y - $qy) * ($y - $qy))
                if ($null -eq $millor -or $d -lt $millor.D) { $millor = [pscustomobject]@{ X = $qx; Y = $qy; D = $d; Ex = $ex; Ey = $ey } }
            }
        }
    }
    return $millor
}

# Un punt que SEGUR que es dins (el centroide d'una L pot caure fora): el
# centroide de l'exterior mes gran si hi es; si no, el mig del tram interior mes
# ample de l'horitzontal que passa pel centroide. { X; Y } o $null.
function Get-PuntInterior($polys) {
    $gran = $null; $ag = -1.0
    foreach ($pl in @($polys)) {
        $an = @($pl.Anells)
        if ($an.Count -eq 0) { continue }
        $a = [math]::Abs((_PgAreaSignada (_PgPunts $an[0])))
        if ($a -gt $ag) { $ag = $a; $gran = $pl }
    }
    if ($null -eq $gran) { return $null }
    $c = Get-CentreAnell @($gran.Anells)[0]
    if ($null -eq $c) { return $null }
    if (Test-PuntDinsPoligons $c[0] $c[1] @($gran)) { return [pscustomobject]@{ X = $c[0]; Y = $c[1] } }
    $y = [double]$c[1]
    $xs = New-Object System.Collections.Generic.List[double]
    foreach ($an in @($gran.Anells)) {
        $p = _PgPunts $an; $n = $p.Count
        for ($i = 0; $i -lt $n; $i++) {
            $a = $p[$i]; $b = $p[($i + 1) % $n]
            if ((($a[1] -gt $y) -ne ($b[1] -gt $y))) { $xs.Add($a[0] + ($y - $a[1]) * ($b[0] - $a[0]) / ($b[1] - $a[1])) }
        }
    }
    $xs.Sort()
    $mx = $null; $ample = -1.0
    for ($i = 0; $i + 1 -lt $xs.Count; $i += 2) {
        if (($xs[$i + 1] - $xs[$i]) -gt $ample) { $ample = $xs[$i + 1] - $xs[$i]; $mx = ($xs[$i] + $xs[$i + 1]) / 2 }
    }
    if ($null -eq $mx) { return [pscustomobject]@{ X = $c[0]; Y = $c[1] } }
    return [pscustomobject]@{ X = $mx; Y = $y }
}

# ON VA L'ETIQUETA D'UNA ENTRADA. $px, $py: el portal del Cadastre.
# Torna $null si el portal es massa lluny de la parcel.la ($maxFora); si no,
# { X; Y; Dx; Dy }: un punt DINS de la parcel.la, a $marge de la vora com a
# minim quan es pot, i la direccio cap a dins (perque l'etiqueta creixi cap a
# l'interior i no cap al carrer).
function Resolve-AncoraEntrada([double]$px, [double]$py, $polys, [double]$maxFora = $Script:PlanolEntradaMaxForaM, [double]$marge = $Script:PlanolEntradaMargeM) {
    if (@($polys).Count -eq 0) { return $null }
    $dins = Test-PuntDinsPoligons $px $py $polys
    $v = Get-VoraMesPropera $px $py $polys
    if ($null -eq $v) { return $null }
    if (-not $dins -and $v.D -gt $maxFora) { return $null }
    # La normal del costat que apunta CAP A DINS.
    $l = [math]::Sqrt($v.Ex * $v.Ex + $v.Ey * $v.Ey)
    if ($l -le 0) { return $null }
    $nx = - $v.Ey / $l; $ny = $v.Ex / $l
    if (-not (Test-PuntDinsPoligons ($v.X + $nx * 0.3) ($v.Y + $ny * 0.3) $polys)) { $nx = - $nx; $ny = - $ny }
    if ($dins -and $v.D -ge $marge) { return [pscustomobject]@{ X = $px; Y = $py; Dx = $nx; Dy = $ny } }
    foreach ($m in @($marge, ($marge / 2), 0.5)) {
        $ax = $v.X + $nx * $m; $ay = $v.Y + $ny * $m
        if (Test-PuntDinsPoligons $ax $ay $polys) { return [pscustomobject]@{ X = $ax; Y = $ay; Dx = $nx; Dy = $ny } }
    }
    # Cantonada molt tancada: el punt interior, que segur que hi es.
    $i = Get-PuntInterior $polys
    if ($null -eq $i) { return $null }
    return [pscustomobject]@{ X = $i.X; Y = $i.Y; Dx = 0.0; Dy = 0.0 }
}

# La direccio cap a dins, en el vocabulari de les etiquetes del mapa: 'r'
# (dreta), 'l', 't' (amunt), 'b' o 'c' (centrada, si no n'hi ha).
function Get-DireccioEtiqueta([double]$dx, [double]$dy) {
    if ([math]::Abs($dx) -lt 1e-9 -and [math]::Abs($dy) -lt 1e-9) { return 'c' }
    if ([math]::Abs($dx) -ge [math]::Abs($dy)) { if ($dx -gt 0) { return 'r' } else { return 'l' } }
    if ($dy -gt 0) { return 't' } else { return 'b' }
}

# ----------------------------------------------------------------------------
# JUNTAR PARCEL.LES QUE ES TOQUEN
# ----------------------------------------------------------------------------
# Les parcel.les del Cadastre comparteixen els vertexs dels costats comuns. La
# unio es fa per CANCEL.LACIO D'ARESTES: cada anell en el seu sentit (exteriors
# antihoraris, forats horaris), un costat comu surt una vegada en cada sentit i
# es treu; el que queda, encadenat, es el contorn de la unio. Abans es parteixen
# els costats pels vertexs dels altres anells que hi cauen a sobre (quan una
# parcel.la toca nomes un tros del costat de la veina).
#
# Torna { Polys; Toquen; Unit }: Toquen = hi havia almenys un costat comu;
# Unit = s'ha pogut fer la unio (si no, Polys son els poligons de sempre, tots
# junts, i el mapa els dibuixa com una sola cosa pero amb la linia del mig).
$Script:PgTol = 0.02

function _PgClau($p) { return ('{0:F2}|{1:F2}' -f $p[0], $p[1]).Replace(',', '.') }

function Join-Poligons($polys) {
    $totes = @($polys)
    $orig = [pscustomobject]@{ Polys = $totes; Toquen = $false; Unit = $false }
    if ($totes.Count -lt 2) { return [pscustomobject]@{ Polys = $totes; Toquen = $false; Unit = ($totes.Count -eq 1) } }
    # Els anells orientats.
    $anells = New-Object System.Collections.Generic.List[object]
    foreach ($pl in $totes) {
        $an = @($pl.Anells)
        for ($k = 0; $k -lt $an.Count; $k++) {
            $p = _PgPunts $an[$k]
            if ($p.Count -lt 3) { continue }
            $a = _PgAreaSignada $p
            if (($k -eq 0 -and $a -lt 0) -or ($k -gt 0 -and $a -gt 0)) { $p.Reverse() }
            $anells.Add($p)
        }
    }
    $vertexs = New-Object System.Collections.Generic.List[double[]]
    foreach ($p in $anells) { foreach ($v in $p) { $vertexs.Add($v) } }
    # Les arestes, partides pels vertexs que hi cauen a sobre.
    $arestes = New-Object System.Collections.Generic.List[object]
    foreach ($p in $anells) {
        $n = $p.Count
        for ($i = 0; $i -lt $n; $i++) {
            $a = $p[$i]; $b = $p[($i + 1) % $n]
            $ex = $b[0] - $a[0]; $ey = $b[1] - $a[1]; $l2 = $ex * $ex + $ey * $ey
            $talls = New-Object System.Collections.Generic.List[object]
            if ($l2 -gt 0) {
                foreach ($v in $vertexs) {
                    $t = (($v[0] - $a[0]) * $ex + ($v[1] - $a[1]) * $ey) / $l2
                    if ($t -le 1e-6 -or $t -ge (1 - 1e-6)) { continue }
                    $qx = $a[0] + $t * $ex; $qy = $a[1] + $t * $ey
                    if ([math]::Sqrt(($v[0] - $qx) * ($v[0] - $qx) + ($v[1] - $qy) * ($v[1] - $qy)) -le $Script:PgTol) { $talls.Add(@($t, $v)) }
                }
            }
            # Una LLISTA tipada: amb @($a) + ... un punt [double[]] es desplegaria
            # en dos numeros.
            $seq = New-Object System.Collections.Generic.List[double[]]
            $seq.Add($a)
            foreach ($tl in @($talls | Sort-Object { [double]$_[0] })) { $seq.Add([double[]]$tl[1]) }
            $seq.Add($b)
            for ($s = 0; $s + 1 -lt $seq.Count; $s++) {
                $ka = _PgClau $seq[$s]; $kb = _PgClau $seq[$s + 1]
                if ($ka -ne $kb) { $arestes.Add([pscustomobject]@{ A = $ka; B = $kb; Pa = $seq[$s]; Viva = $true }) }
            }
        }
    }
    # Cancel.lacio: un costat i el seu invers.
    $perClau = @{}
    foreach ($e in $arestes) { $c = $e.A + '>' + $e.B; if (-not $perClau.ContainsKey($c)) { $perClau[$c] = New-Object System.Collections.Generic.List[object] }; $perClau[$c].Add($e) }
    $cancel = 0
    foreach ($e in $arestes) {
        if (-not $e.Viva) { continue }
        $inv = $e.B + '>' + $e.A
        if (-not $perClau.ContainsKey($inv)) { continue }
        $parella = $null
        foreach ($f in $perClau[$inv]) { if ($f.Viva) { $parella = $f; break } }
        if ($null -eq $parella) { continue }
        $e.Viva = $false; $parella.Viva = $false; $cancel++
    }
    if ($cancel -eq 0) { return $orig }
    $orig.Toquen = $true
    # Encadenar el que queda.
    $surten = @{}
    foreach ($e in $arestes) { if ($e.Viva) { if (-not $surten.ContainsKey($e.A)) { $surten[$e.A] = New-Object System.Collections.Generic.List[object] }; $surten[$e.A].Add($e) } }
    $resultats = New-Object System.Collections.Generic.List[object]
    foreach ($e0 in $arestes) {
        if (-not $e0.Viva) { continue }
        $ring = New-Object System.Collections.Generic.List[double[]]
        $e = $e0; $pas = 0
        while ($true) {
            $e.Viva = $false; [void]$surten[$e.A].Remove($e)
            $ring.Add($e.Pa)
            if ($e.B -eq $e0.A) { break }
            if (-not $surten.ContainsKey($e.B) -or $surten[$e.B].Count -eq 0) { return $orig }
            $e = $surten[$e.B][0]
            $pas++; if ($pas -gt 100000) { return $orig }
        }
        if ($ring.Count -ge 3) { $resultats.Add($ring) }
    }
    # Exteriors (antihoraris) i forats (horaris), cada forat al seu exterior.
    $ext = New-Object System.Collections.Generic.List[object]
    $forats = New-Object System.Collections.Generic.List[object]
    foreach ($r in $resultats) { if ((_PgAreaSignada $r) -gt 0) { $ext.Add($r) } else { $forats.Add($r) } }
    if ($ext.Count -eq 0) { return $orig }
    $sortida = New-Object System.Collections.Generic.List[object]
    $grups = @{}
    for ($i = 0; $i -lt $ext.Count; $i++) { $grups[$i] = New-Object System.Collections.Generic.List[object]; $grups[$i].Add((_PgPla $ext[$i])) }
    foreach ($h in $forats) {
        $hp = _PgPla $h
        for ($i = 0; $i -lt $ext.Count; $i++) {
            if (Test-PuntDinsAnell $h[0][0] $h[0][1] (_PgPla $ext[$i])) { $grups[$i].Add($hp); break }
        }
    }
    for ($i = 0; $i -lt $ext.Count; $i++) { $sortida.Add([pscustomobject]@{ Anells = $grups[$i].ToArray() }) }
    return [pscustomobject]@{ Polys = $sortida.ToArray(); Toquen = $true; Unit = $true }
}
