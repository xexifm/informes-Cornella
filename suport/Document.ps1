#requires -Version 5.1
<#
.SYNOPSIS
  Composicio del document Word final (Pas 6).

.DESCRIPTION
  Copia '0 CAPCALERA.docx', hi substitueix els <<PLACEHOLDERS>>, hi escriu el
  cos (seccions, items, sub-punts, enllacos) i el bloc de conclusions, i el desa
  amb el nom que toca a la carpeta de sortida. Tot el FORMAT (lletra, sangries,
  espaiats) el posa Format.ps1: aqui nomes es decideix QUE s'escriu i EN QUIN
  ORDRE.

  Ve de Motor.ps1. Es dot-sourceja des d'alli: mateix ambit, mateix
  comportament.
#>

# ----------------------------------------------------------------------------
# Step 6 - Compose final document
# ----------------------------------------------------------------------------
# LA CAPCALERA I LA CARPETA DE SORTIDA ja no son aqui: han passat a
# MotorInforme.ps1. Vivien en aquest fitxer nomes perque es on es van estrenar,
# pero no son de REQ1 -les fan servir vuit fitxers mes- i, sobretot,
# Write-InformeDocx (el motor generic) les cridava: el motor depenia del seu
# propi client. El que si que s'ha quedat aqui es _GetOutputFileName, que nomes
# el fa servir Build-Document.

# Calcula el nom de fitxer de sortida: YYYY-MM-DD_<TipusCataleg>_GIA <id>.docx
function _GetOutputFileName($catalegName, $gia) {
    $today = (Get-Date).ToString('yyyy-MM-dd')
    $cat   = $catalegName
    if ($cat) { $cat = $cat.Substring(0,1).ToUpper() + $cat.Substring(1).ToLower() }
    else      { $cat = 'Informe' }
    if ([string]::IsNullOrWhiteSpace($gia)) { $gia = 's_n' }
    # El sanejat el fa _SanejaNomFitxer (MotorInforme.ps1), que es el mateix per
    # a tots. Aqui posava '_' i les altres dues families '-': el mateix caracter
    # dolent donava un nom diferent segons quin informe fessis. La FORMA del nom
    # si que es d'aqui (la inicial en majuscula i el 's_n' quan no hi ha GIA).
    return (_SanejaNomFitxer ("{0}_{1}_GIA {2}.docx" -f $today, $cat, $gia.Trim()))
}

# Escriu el cos del document (intro del cataleg + seccions amb items numerats).
# Retorna el comptador global utilitzat per a la numeracio.
function _WriteCatalegBody($sel, $cfg, $selectedSections, $fields, $introText, $isFixedBody = $false, $fixedBodyLines = @()) {
    # QUE s'escriu ho decideix Build-CatalegBlocs (pura, MotorInforme.ps1) i COM
    # s'escriu, Write-Informe. Abans aquest cos i _VistaCataleg (VistaWord.ps1)
    # eren el MATEIX algorisme escrit dues vegades, i cada un decidia pel seu
    # compte l'aire i quin sub-punt era el primer.
    $blocs = Build-CatalegBlocs $selectedSections $fields $introText $isFixedBody $fixedBodyLines
    [void](Write-Informe $sel $blocs)
}

function Get-TextTancament {
    try { return @((Read-Conclusions $ConclusionsPath).Always) } catch { return @() }
}

function Write-Tancament($sel, $fields = $null) {
    foreach ($l in @(Get-TextTancament)) {
        Format-Conclusion $sel (Apply-Fields -text ([string]$l) -fields $fields)
    }
}

function _WriteConclusionsBlock($sel, $cfg, $headerText, $conclusions, $alwaysConclusions, $fields) {
    $hasBody = ($conclusions.Count -gt 0) -or ($alwaysConclusions.Count -gt 0)
    $hasHead = -not [string]::IsNullOrWhiteSpace($headerText)
    if (-not $hasBody -and -not $hasHead) { return }

    Format-Aire $sel 'conclusions'

    if ($hasHead) {
        Format-ConclusionHeader $sel $headerText
    }

    foreach ($c in $conclusions) {
        $txt = if ($c -is [string]) { $c } else { [string]$c.Body }
        $resolved = Apply-Fields -text $txt -fields $fields
        Format-Conclusion $sel $resolved
    }
    foreach ($a in $alwaysConclusions) {
        $resolved = Apply-Fields -text ([string]$a) -fields $fields
        Format-Conclusion $sel $resolved
    }
}

function Build-Document($word, $header, $selectedSections, $fields, $conclusions, $alwaysConclusions, $catalegName, $introText, $conclusionsHeaderText, $isFixedBody = $false, $fixedBodyLines = @()) {
    $baseName = _GetOutputFileName $catalegName $header['ID_GIA']
    $cfg = $Script:ReportFormatConfig

    # 0 CAPCALERA.docx pot portar tambe els blocs d'ACT_EXTR i de LLIC a sota;
    # ens quedem nomes amb el generic ('' = el primer), que es el de REQ1,
    # TERMINI i qualsevol cataleg nou.
    return Write-InformeDocx $word $baseName '' $header {
        param($sel)
        _WriteCatalegBody $sel $cfg $selectedSections $fields $introText $isFixedBody $fixedBodyLines
        _WriteConclusionsBlock $sel $cfg $conclusionsHeaderText $conclusions $alwaysConclusions $fields
    }
}
