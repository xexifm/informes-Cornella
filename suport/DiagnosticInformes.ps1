#requires -Version 5.1
<#
.SYNOPSIS
  Diagnostic de la carpeta d'informes: NO envia contingut, nomes ESTRUCTURA.

.DESCRIPTION
  Escaneja $InformesDir al TEU PC i escriu un resum (poques KB) a l'Escriptori:
  quins formats de data tenen els noms, quants .docx / .doc, si l'ID GIA es
  troba al document / a la carpeta / enlloc, i quantes conclusions ("Vist
  l'anterior") es detecten. Serveix perque en Claude afini els parsers sense
  necessitat de pujar els 43 GB.

  Reutilitza EXACTAMENT les mateixes funcions que el programa real (dot-source
  de GenerarInforme.ps1 en mode headless), aixi el diagnostic reflecteix on
  fallaria l'escaner de debo.

  Privadesa: el resum inclou aggregats + uns pocs EXEMPLES (rutes de fitxer i, en
  els casos sense conclusio, les ultimes linies -text legal boilerplate-). Repassa
  el .txt abans d'enviar-lo si vols.

.PARAMETER MaxContent
  Maxim de .docx que s'obren per mirar-ne el contingut (mostra repartida per tot
  l'arbre). Per defecte 600. La resta d'estadistiques (noms, extensions,
  profunditat) es fan sobre TOTS els fitxers.
#>
param(
    [int]$MaxContent = 600
)

$ErrorActionPreference = 'Stop'
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

# Carreguem el motor en headless (nomes per tenir $InformesDir i les funcions de
# lectura/parseig; no obre finestres ni executa res).
$env:GENINFORME_TEST = '1'
. (Join-Path $ScriptRoot 'GenerarInforme.ps1')
Remove-Item Env:\GENINFORME_TEST -ErrorAction SilentlyContinue

$dir = $InformesDir
if ([string]::IsNullOrWhiteSpace($dir) -or -not (Test-Path -LiteralPath $dir)) {
    Write-Host "No s'ha trobat la carpeta d'informes: $dir" -ForegroundColor Red
    Write-Host "Configura \$InformesDir a suport\config.ps1." -ForegroundColor Yellow
    Read-Host "Prem Enter per tancar"
    return
}

Write-Host "Escanejant $dir ..." -ForegroundColor Cyan
Write-Host "(pot trigar una estona; son molts fitxers)" -ForegroundColor DarkGray

# --- Lectura del text d'un Word ---
# .docx: sense Word (descomprimint, rapid). .doc antic (Word 97-2003): via Word
# COM (mes lent; per aixo nomes s'obren els de la MOSTRA). Word es crea una sola
# vegada i es tanca al final.
$script:_wordApp = $null
function _GetWordApp {
    if ($null -ne $script:_wordApp) { return $script:_wordApp }
    try {
        $script:_wordApp = New-Object -ComObject Word.Application
        $script:_wordApp.Visible = $false
        $script:_wordApp.DisplayAlerts = 0
        try { $script:_wordApp.AutomationSecurity = 1 } catch { }
    } catch { $script:_wordApp = $null }
    return $script:_wordApp
}
function _CloseWordApp {
    if ($null -ne $script:_wordApp) {
        try { $script:_wordApp.Quit() } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:_wordApp) | Out-Null } catch { }
        $script:_wordApp = $null
    }
}
function _ReadDocText($path) {
    if ($path -match '(?i)\.docx$') {
        try { return _ReadDocxParagraphs $path } catch { return $null }
    }
    # .doc antic: cal Word.
    $w = _GetWordApp
    if ($null -eq $w) { return $null }
    $doc = $null
    try {
        $doc = $w.Documents.Open($path, $false, $true)   # ReadOnly
        $out = New-Object System.Collections.ArrayList
        foreach ($p in $doc.Paragraphs) {
            [void]$out.Add((($p.Range.Text)).TrimEnd("`r", "`n", "`a", " "))
        }
        return $out.ToArray()
    } catch {
        return $null
    } finally {
        if ($null -ne $doc) { try { $doc.Close($false) } catch { } }
    }
}

# --- Passada 1: TOTS els fitxers (nom, extensio, profunditat) ---
$extCount   = @{}
$dateBucket = [ordered]@{ 'AAAAMMDD' = 0; 'AAAA-sep-MM-DD' = 0; 'AA-sep-MM-DD' = 0; 'altre-format-data' = 0 }
$dateEx     = @{}
$depthCount = @{}
$startDigitNoDate = New-Object System.Collections.ArrayList  # comencen amb xifra pero data NO reconeguda
$docxInforme = New-Object System.Collections.ArrayList        # .docx amb data (candidats a contingut)
$docInforme  = New-Object System.Collections.ArrayList        # .doc amb data (no llegibles sense Word)
$nTot = 0

function _AddEx($tbl, $key, $val, $max = 8) {
    if (-not $tbl.ContainsKey($key)) { $tbl[$key] = New-Object System.Collections.ArrayList }
    if ($tbl[$key].Count -lt $max) { [void]$tbl[$key].Add($val) }
}

Get-ChildItem -LiteralPath $dir -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
    $f = $_
    $nTot++
    if (($nTot % 2000) -eq 0) { Write-Host "  ... $nTot fitxers" -ForegroundColor DarkGray }

    $ext = $f.Extension.ToLower()
    if (-not $extCount.ContainsKey($ext)) { $extCount[$ext] = 0 }
    $extCount[$ext]++

    $rel = $f.FullName.Substring($dir.Length).TrimStart('\', '/')
    $depth = (@($rel -split '[\\/]').Count) - 1
    if (-not $depthCount.ContainsKey($depth)) { $depthCount[$depth] = 0 }
    $depthCount[$depth]++

    $name = $f.Name
    $d = _ParseDataInformeFromName $name
    if ($d) {
        $bucket = if ($name -match '^\d{8}(\D|$)') { 'AAAAMMDD' }
                  elseif ($name -match '^\d{4}[-_.]\d{2}[-_.]\d{2}') { 'AAAA-sep-MM-DD' }
                  elseif ($name -match '^\d{2}[-_.]\d{2}[-_.]\d{2}') { 'AA-sep-MM-DD' }
                  else { 'altre-format-data' }
        $dateBucket[$bucket]++
        _AddEx $dateEx $bucket $name
        if     ($ext -eq '.docx') { [void]$docxInforme.Add($f.FullName) }
        elseif ($ext -eq '.doc')  { [void]$docInforme.Add($f.FullName) }
    }
    elseif ($name -match '^\d') {
        # Comenca amb xifra pero no s'ha reconegut com a data: possible format nou.
        if ($startDigitNoDate.Count -lt 30) { [void]$startDigitNoDate.Add($name) }
    }
}

# --- Passada 2: contingut d'una MOSTRA repartida (.docx + .doc antics amb data) ---
$candidates = @(@($docxInforme) + @($docInforme))
$total = $candidates.Count
$sample = @()
if ($total -le $MaxContent) {
    $sample = @($candidates)
} else {
    $step = [double]$total / $MaxContent
    $sample = for ($i = 0; $i -lt $MaxContent; $i++) { $candidates[[int][math]::Floor($i * $step)] }
}

$giaDoc = 0; $giaFolder = 0; $giaNone = 0; $illegible = 0
$conclOk = 0; $conclNo = 0
$giaNoneEx = New-Object System.Collections.ArrayList
$conclNoEx = New-Object System.Collections.ArrayList
$k = 0
foreach ($path in $sample) {
    $k++
    if (($k % 100) -eq 0) { Write-Host "  ... contingut $k de $($sample.Count)" -ForegroundColor DarkGray }
    $lines = _ReadDocText $path
    if ($null -eq $lines) { $illegible++; continue }

    $gia = _ExtractIdGia $lines
    if (-not [string]::IsNullOrWhiteSpace($gia)) {
        $giaDoc++
    } else {
        $gf = _GiaFromFolderName $path
        if (-not [string]::IsNullOrWhiteSpace($gf)) {
            $giaFolder++
        } else {
            $giaNone++
            if ($giaNoneEx.Count -lt 15) {
                # Nomes ruta + linies de capcalera (GIA/Exp), no titular/adreca.
                $hdr = @($lines | Where-Object { $_ -match '(?i)(ID\s*GIA|Exp)' } | Select-Object -First 4)
                [void]$giaNoneEx.Add(($path + "`n      capcalera: " + ($hdr -join ' | ')))
            }
        }
    }

    $concl = _ExtractConclusio $lines
    if (-not [string]::IsNullOrWhiteSpace($concl)) {
        $conclOk++
    } else {
        $conclNo++
        if ($conclNoEx.Count -lt 15) {
            # Ultimes linies (on solen ser les conclusions): boilerplate legal.
            $tail = @($lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Last 6 |
                      ForEach-Object { if ($_.Length -gt 160) { $_.Substring(0, 160) + '...' } else { $_ } })
            [void]$conclNoEx.Add($path + "`n      final: " + ($tail -join ' / '))
        }
    }
}

# Tanquem Word si l'haviem obert per als .doc.
_CloseWordApp

# --- Escriure el resum ---
$sb = New-Object System.Text.StringBuilder
function _L($t) { [void]$sb.AppendLine($t) }

_L "=== DIAGNOSTIC DE LA CARPETA D'INFORMES ==="
_L ("Data: " + (Get-Date).ToString('yyyy-MM-dd HH:mm'))
_L ("Carpeta: " + $dir)
_L ("Fitxers totals: " + $nTot)
_L ""
_L "--- EXTENSIONS ---"
foreach ($e in ($extCount.GetEnumerator() | Sort-Object Value -Descending)) { _L ("  {0,-8} {1}" -f $e.Key, $e.Value) }
_L ""
_L "--- FORMATS DE DATA al nom (fitxers considerats informe) ---"
foreach ($b in $dateBucket.GetEnumerator()) {
    _L ("  {0,-20} {1}" -f $b.Key, $b.Value)
    if ($dateEx.ContainsKey($b.Key)) { foreach ($x in $dateEx[$b.Key]) { _L ("      ex: " + $x) } }
}
_L ("  .docx amb data: " + $docxInforme.Count + "   |   .doc amb data: " + $docInforme.Count)
_L ""
_L "--- COMENCEN AMB XIFRA pero data NO reconeguda (possibles formats nous) ---"
if ($startDigitNoDate.Count -eq 0) { _L "  (cap)" } else { foreach ($x in $startDigitNoDate) { _L ("  " + $x) } }
_L ""
_L "--- PROFUNDITAT (nivells de carpeta sota Informes) ---"
foreach ($p in ($depthCount.GetEnumerator() | Sort-Object Name)) { _L ("  {0} nivell(s): {1}" -f $p.Key, $p.Value) }
_L ""
_L ("--- CONTINGUT (mostra de " + $sample.Count + " informes .docx/.doc) ---")
_L ("  ID GIA trobat al DOCUMENT:   " + $giaDoc)
_L ("  ID GIA nomes per CARPETA:    " + $giaFolder)
_L ("  ID GIA NO trobat (ni doc ni carpeta; potser si per Excel/expedient): " + $giaNone)
_L ("  .docx il-legibles:           " + $illegible)
_L ("  Conclusio trobada:           " + $conclOk)
_L ("  Conclusio NO trobada:        " + $conclNo)
_L ""
_L "  Exemples SENSE ID GIA (doc/carpeta):"
if ($giaNoneEx.Count -eq 0) { _L "    (cap a la mostra)" } else { foreach ($x in $giaNoneEx) { _L ("    - " + $x) } }
_L ""
_L "  Exemples SENSE conclusio detectada (mira'n el final):"
if ($conclNoEx.Count -eq 0) { _L "    (cap a la mostra)" } else { foreach ($x in $conclNoEx) { _L ("    - " + $x) } }
_L ""
if ($docInforme.Count -gt 0) {
    _L ("--- .doc ANTICS (Word 97-2003): " + $docInforme.Count + " (no es poden llegir sense Word) ---")
    foreach ($x in ($docInforme | Select-Object -First 10)) { _L ("  " + $x) }
    _L ""
}

$desktop = [Environment]::GetFolderPath('Desktop')
if ([string]::IsNullOrWhiteSpace($desktop) -or -not (Test-Path -LiteralPath $desktop)) { $desktop = $env:TEMP }
$outPath = Join-Path $desktop 'informes-diagnostic.txt'
$sb.ToString() | Set-Content -LiteralPath $outPath -Encoding UTF8

Write-Host ""
Write-Host ("Diagnostic escrit a: " + $outPath) -ForegroundColor Green
Write-Host "Obre'l, repassa'l i envia'l a Claude." -ForegroundColor Green
try { Start-Process -FilePath 'notepad.exe' -ArgumentList "`"$outPath`"" | Out-Null } catch { }
