# El REPORTATGE FOTOGRAFIC dels informes de visita d'inspeccio (Fotos.ps1 i
# Format-Fotos): la disposicio, la mida reduida, el gir EXIF, els blocs, la frase
# de les conclusions i com s'escriu al Word (amb un Word simulat).
# Es DOT-SOURCE des de run-tests.ps1 (mateix ambit i comptador d'asserts).

Write-Host "`n--- Fotos: la disposicio ---"
$fV = { param($n) @{ Path = ('v' + $n + '.jpg'); Vertical = $true } }
$fH = { param($n) @{ Path = ('h' + $n + '.jpg'); Vertical = $false } }
$filesF = @(_FotosFiles @((& $fV 1), (& $fV 2), (& $fH 3), (& $fV 4), (& $fH 5), (& $fV 6), (& $fV 7), (& $fV 8)))
$txtF = @($filesF | ForEach-Object { (@($_.Fotos) | ForEach-Object { [string]$_.Num + '@' + [string]$_.AmpleCm }) -join '+' }) -join ' | '
AssertEq $txtF '1@6+2@6 | 3@12 | 4@6 | 5@12 | 6@6+7@6 | 8@6' 'Fotos: dues verticals SEGUIDES de costat (6+6 cm), les horitzontals a 12 cm i una vertical sola a 6 cm; sense reordenar'
AssertEq @(_FotosFiles @()).Count 0 'Fotos: sense fotos, cap fila'

Write-Host "`n--- Fotos: la mida i el gir ---"
AssertEq ((_FotoMidaReduida 4000 3000 1600) -join 'x') '1600x1200' 'Fotos: el costat llarg a 1600 px, sense deformar'
AssertEq ((_FotoMidaReduida 3000 4000 1600) -join 'x') '1200x1600' 'Fotos: tambe una vertical'
AssertEq ((_FotoMidaReduida 800 600 1600) -join 'x') '800x600' 'Fotos: una de petita no s''engrandeix'
AssertEq (_FotoGirExif 6) 'Rotate90FlipNone' 'Fotos: EXIF 6 (el mobil dret) -> girar 90 graus'
AssertEq (_FotoGirExif 1) '' 'Fotos: EXIF 1 -> res'
AssertEq "$(_FotoGiraEixos 6)|$(_FotoGiraEixos 8)|$(_FotoGiraEixos 3)" 'True|True|False' 'Fotos: 5-8 intercanvien ample i alt (90/270), 3 no (180)'

Write-Host "`n--- Fotos: els blocs, la frase i el pas ---"
$blF = @(_FotosBlocs @((& $fV 1), (& $fH 2)))
AssertEq (@($blF | ForEach-Object { $_.T }) -join ',') 'saltpagina,conclusiocap,fotos,fotos' 'Fotos: pagina nova, el titol i una fila per bloc'
AssertEq ([string]$blF[1].Text) ('REPORTATGE FOTOGR' + [char]0x00C0 + 'FIC') 'Fotos: el titol'
AssertEq @(_FotosBlocs @()).Count 0 'Fotos: sense fotos, cap bloc (l''informe queda igual que abans)'
$cF = @(_ConclusionsAmbFotos @([pscustomobject]@{ Title = 'Req'; Body = 'Vist l''anterior...' }) @((& $fV 1)))
AssertEq "$($cF.Count)|$($cF[1])" ('2|S' + [char]0x2019 + 'adjunta reportatge fotogr' + [char]0x00E0 + 'fic.') 'Fotos: la frase va darrere de les conclusions triades'
AssertEq @(_ConclusionsAmbFotos @('A') @()).Count 1 'Fotos: sense fotos, cap frase'
AssertEq "$(_EsVisitaInspeccio ([ordered]@{ ORIGEN_TIPUS = 'insp' }))|$(_EsVisitaInspeccio @{ ORIGEN_TIPUS = 'doc' })|$(_EsVisitaInspeccio $null)" 'True|False|False' 'Fotos: el pas nomes surt amb "Visita d''inspeccio"'
AssertEq (_FotosCarpetaInicial '' 'I:\Informes') 'I:\Informes' 'Fotos: el dialeg s''obre a la carpeta d''informes de la Configuracio'
AssertEq (_FotosCarpetaInicial 'I:\Informes\GIA 1\Fotos' 'I:\Informes') 'I:\Informes\GIA 1\Fotos' 'Fotos: ...i despres, a l''ultima que s''ha fet servir'

Write-Host "`n--- Fotos: Write-Informe (amb els dobles de Format.ps1) ---"
try {
    . (Join-Path $TestsDir 'FormatDoubles.ps1')
    [void](Write-Informe $null @(_FotosBlocs @((& $fV 1), (& $fV 2), (& $fH 3))))
    AssertEq (@($global:emitCalls) -join ' / ') ('CONCLCAP|REPORTATGE FOTOGR' + [char]0x00C0 + 'FIC / FOTOS|1:v1.jpg@6,2:v2.jpg@6 / FOTOS|3:h3.jpg@12') 'Fotos: el titol i les files arriben a Format-Fotos'
} catch {
    Assert $false ('bloc Write-Informe fotos: ' + $_.Exception.Message + ' @ ' + $_.InvocationInfo.ScriptLineNumber)
} finally {
    . (Join-Path (Split-Path -Parent $TestsDir) 'Format.ps1')
}

Write-Host "`n--- Format-Fotos: com s'escriu al Word (Word simulat) ---"
try {
    $wf = @{ Log = (New-Object System.Collections.ArrayList); Pos = 0 }
    $mkRange = { param([int]$s, [int]$e) $r = [pscustomobject]@{ Start = $s; End = $e; Font = [pscustomobject]@{ Bold = 0 } }; $r | Add-Member -MemberType ScriptProperty -Name ListFormat -Value { [pscustomobject]@{} } ; return $r }
    $tabsF = [pscustomobject]@{}
    $tabsF | Add-Member -MemberType ScriptMethod -Name ClearAll -Value { [void]$wf.Log.Add('TABS-') }
    $tabsF | Add-Member -MemberType ScriptMethod -Name Add -Value { param($p, $a) [void]$wf.Log.Add(('TAB ' + [Math]::Round($p, 1) + '/' + $a)) }
    $pfF = [pscustomobject]@{ LeftIndent = 0; FirstLineIndent = 0; Alignment = 0; SpaceBefore = 0; SpaceAfter = 0; LineSpacingRule = 5; TabStops = $tabsF }
    $shapes = [pscustomobject]@{}
    $shapes | Add-Member -MemberType ScriptMethod -Name AddPicture -Value {
        param($path, $link, $save)
        [void]$wf.Log.Add(('PIC ' + [System.IO.Path]::GetFileName([string]$path) + ' link=' + $link + ' save=' + $save))
        $wf.Pos += 1
        $shp = [pscustomobject]@{ LockAspectRatio = 0; Width = 0; Range = [pscustomobject]@{ End = $wf.Pos } }
        $shp | Add-Member -MemberType ScriptProperty -Name W -Value { $this.Width }
        [void]$wf.Log.Add('@shape'); $wf.Ultima = $shp
        return $shp
    }
    $selF = [pscustomobject]@{ ParagraphFormat = $pfF; Font = [pscustomobject]@{ Bold = 0; Italic = 0; Underline = 0; Size = 0; Color = 0; Name = '' }; InlineShapes = $shapes
        PageSetup = [pscustomobject]@{ PageWidth = 595.3; LeftMargin = 70.9; RightMargin = 70.8 }; Document = [pscustomobject]@{} }
    $selF | Add-Member -MemberType ScriptProperty -Name Range -Value { & $mkRange $wf.Pos $wf.Pos }
    $selF | Add-Member -MemberType ScriptMethod -Name TypeParagraph -Value { [void]$wf.Log.Add('PARA') }
    $selF | Add-Member -MemberType ScriptMethod -Name TypeText -Value { param($t) [void]$wf.Log.Add(('TEXT ' + ([string]$t -replace "`t", '<tab>'))); $wf.Pos += ([string]$t).Length }
    $selF | Add-Member -MemberType ScriptMethod -Name SetRange -Value { param($a, $b) [void]$wf.Log.Add(('SET ' + $a)) }
    $selF.Document | Add-Member -MemberType ScriptMethod -Name Range -Value { param($a, $b) & $mkRange $a $b }
    Format-Fotos $selF @{ Fotos = @(@{ Path = 'foto-001.jpg'; Num = 1; AmpleCm = 6 }, @{ Path = 'foto-002.jpg'; Num = 2; AmpleCm = 6 }) }
    $logF = @($wf.Log) -join ' ; '
    Assert ($logF -match 'TAB 134\.7/1 ; TAB 318\.9/1') ('Format-Fotos: dues tabulacions centrades a 3,25 cm de cada costat del mig de la pagina (' + $logF + ')')
    Assert ($logF -match 'TEXT <tab> ; PIC foto-001\.jpg link=False save=True ; @shape ; SET 2 ; TEXT <tab> ; PIC foto-002\.jpg') 'Format-Fotos: tabulacio + foto incrustada (no enllacada), i el cursor darrere de la foto'
    Assert ($logF.EndsWith('TEXT <tab>Foto 1<tab>Foto 2')) 'Format-Fotos: a sota, "Foto 1" i "Foto 2" a les mateixes tabulacions'
    AssertEq "$([Math]::Round($wf.Ultima.Width, 1))|$($wf.Ultima.LockAspectRatio)" '170.1|-1' 'Format-Fotos: 6 cm d''ample amb la proporcio bloquejada'
    AssertEq ([int]$pfF.LineSpacingRule) 0 'Format-Fotos: interlineat senzill (un "exacte" retallaria la foto)'
} catch {
    Assert $false ('bloc Format-Fotos: ' + $_.Exception.Message + ' @ ' + $_.InvocationInfo.ScriptLineNumber)
}

Write-Host "`n--- Fotos: guards de font ---"
$srcWz = _SenseComentaris (Join-Path (Split-Path -Parent $TestsDir) 'Wizard.ps1')
Assert ($srcWz.Contains('Select-FotosInforme $InformesDir $st.Fotos') -and $srcWz.Contains('if (-not (_EsVisitaInspeccio $st.Header))')) 'Wizard: el pas de les fotos, nomes amb visita d''inspeccio, i s''obre a la carpeta d''informes'
Assert ($srcWz.Contains('$prep = Convert-FotosInforme $st.Fotos') -and $srcWz.Contains('-fotos $fotosDoc')) 'Wizard: les fotos van REDUIDES a l''informe'
$srcDoc = _SenseComentaris (Join-Path (Split-Path -Parent $TestsDir) 'Document.ps1')
Assert ($srcDoc.Contains('@(_ConclusionsAmbFotos $conclusions $fotos)') -and $srcDoc.Contains('Write-Informe $sel @(_FotosBlocs $fotos)')) 'Build-Document: la frase a les conclusions i el reportatge darrere'
$srcFo = _SenseComentaris (Join-Path (Split-Path -Parent $TestsDir) 'Fotos.ps1')
Assert (-not ($srcFo -match 'InlineShapes|ParagraphFormat')) 'Fotos.ps1: no toca el Word (el format es de Format.ps1)'
