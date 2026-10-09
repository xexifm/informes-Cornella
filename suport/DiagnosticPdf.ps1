#requires -Version 5.1
<#
.SYNOPSIS
  DIAGNOSTIC del lector de PDF del repas de contactes: escriu en un .txt el
  text que se n'extreu i el que en treuen els lectors. Nomes per al PC de
  l'usuari.

.DESCRIPTION
  No s'han pogut veure instancies ni autoritzacions de debo (porten dades
  personals): el lector de text (PdfText.ps1) i els lectors de cada document
  (ContactesExtraccio.ps1) es van fer amb documents INVENTATS. Aixo es el que
  ho comprova amb els de veritat: per al PDF (o l'XML de l'e-TRAM) que es
  trii, escriu
    - el text del lector propi (o l'error: xifrat, sense text...)
    - si no n'ha sortit, el del Word (el recurs del repas)
    - el que en treuen els lectors (interessat, representant, autoritzacio)
  a local\base-dades-activitats\diagnostic-pdf_<nom>.txt, i l'obre.

  Us: arrossega el PDF damunt de local\DiagnosticPdf.bat (o obre'l i tria'l), o:
    powershell -NoProfile -ExecutionPolicy Bypass -STA -File suport\DiagnosticPdf.ps1 -Pdf <fitxer>
#>
param([string]$Pdf = '')

$ErrorActionPreference = 'Stop'
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$MotorSenseGui = $true
. (Join-Path $ScriptRoot 'Motor.ps1')

if ([string]::IsNullOrWhiteSpace($Pdf)) {
    Add-Type -AssemblyName System.Windows.Forms
    $dpDlg = New-Object System.Windows.Forms.OpenFileDialog
    $dpDlg.Title = 'Tria el PDF (o l''XML de l''e-TRAM) per diagnosticar'
    $dpDlg.Filter = 'PDF o XML (*.pdf;*.xml)|*.pdf;*.xml|Tots (*.*)|*.*'
    if (_InformesDirAccessible $InformesDir) { $dpDlg.InitialDirectory = $InformesDir }
    if ($dpDlg.ShowDialog() -ne 'OK') { exit 0 }
    $Pdf = $dpDlg.FileName
}
if (-not (Test-Path -LiteralPath $Pdf)) { throw "No trobo el fitxer: $Pdf" }
$dpF = Get-Item -LiteralPath $Pdf
$dpL = New-Object System.Collections.ArrayList
$dpDiu = { param($t) [void]$dpL.Add([string]$t) }
$dpJson = { param($o) if ($null -eq $o) { '(res)' } else { $o | ConvertTo-Json -Depth 8 } }
& $dpDiu ('Diagnostic del lector de documents (contactes ' + $Script:ContactesVersio + ')  -  ' + (Get-Date).ToString('dd/MM/yyyy HH:mm'))
& $dpDiu ('Fitxer: ' + $dpF.FullName + '   (' + $dpF.Length + ' bytes)')
$dpTipus = _CtTipusDocument $dpF.Name
& $dpDiu ('Tipus segons el nom: ' + $(if ($dpTipus) { $dpTipus } else { '(cap: el repas NO el llegiria pel nom)' }))
& $dpDiu ''

if ($dpF.Extension -ieq '.xml') {
    $dpTxt = _CtTextXml $dpF.FullName
    & $dpDiu '== El que en treu el lector de l''e-TRAM =='
    & $dpDiu (& $dpJson (Read-EtramXml $dpTxt))
    & $dpDiu ''
    & $dpDiu '== El que en treu el lector dels XML antics =='
    & $dpDiu (& $dpJson (Read-XmlAntic $dpTxt))
} else {
    $dpTxt = ''
    & $dpDiu '== Text del lector propi (PdfText.ps1) =='
    try { $dpTxt = Get-PdfText $dpF.FullName; & $dpDiu $dpTxt } catch { & $dpDiu ('ERROR: ' + $_.Exception.Message) }
    & $dpDiu ('(hi ha prou text: ' + (_PdfTeText $dpTxt) + ')')
    if (-not (_PdfTeText $dpTxt)) {
        & $dpDiu ''
        & $dpDiu '== Text del Word (el recurs, nomes si el propi no en treu) =='
        $dpW = $null
        try {
            $dpW = New-WordApp -Opcional
            if ($null -eq $dpW) { & $dpDiu '(no hi ha Word)' }
            else { $dpTxt = _PdfTextWord $dpW $dpF.FullName; & $dpDiu $dpTxt; & $dpDiu ('(hi ha prou text: ' + (_PdfTeText $dpTxt) + ')') }
        } finally { if ($null -ne $dpW) { try { $dpW.Quit() } catch { } } }
    }
    & $dpDiu ''
    & $dpDiu '== El que en treu el lector de la INSTANCIA =='
    & $dpDiu (& $dpJson (Read-InstanciaText $dpTxt))
    & $dpDiu ''
    & $dpDiu '== El que en treu el lector de l''AUTORITZACIO =='
    & $dpDiu (& $dpJson (Read-AutoritzacioText $dpTxt))
}

$dpNom = ([System.IO.Path]::GetFileNameWithoutExtension($dpF.Name) -replace '[^\w\-]+', '_')
$dpSortida = Join-Path $LocalActivitatsDir ('diagnostic-pdf_' + $dpNom + '.txt')
[void](New-Item -ItemType Directory -Path $LocalActivitatsDir -Force)
[System.IO.File]::WriteAllText($dpSortida, ($dpL -join "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
Write-Host ('Desat a: ' + $dpSortida)
try { Start-Process -FilePath 'notepad.exe' -ArgumentList ('"' + $dpSortida + '"') | Out-Null } catch { }
