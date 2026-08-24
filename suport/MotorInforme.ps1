#requires -Version 5.1
<#
.SYNOPSIS
  El motor de composicio d'informes: el que TOTS els informes fan igual.

.DESCRIPTION
  Aqui hi va el que abans estava escrit una vegada a cada familia d'informe.
  Format.ps1 diu COM es veu cada paragraf; aquest modul diu COM ES MUNTA un
  document sencer, de manera que canviar-ho en un lloc afecti a tots.

  Punts d'entrada:
    Write-InformeDocx  -> obrir la plantilla, escriure el cos i desar
#>

# ----------------------------------------------------------------------------
# OBRIR, ESCRIURE I DESAR UN INFORME
# ----------------------------------------------------------------------------
# La seqüencia que feien IGUAL les quatre families (REQ1/TERMINI, ACT_EXTR,
# Llicencia i MNS/Traspas), ~20 linies copiades quatre vegades:
#
#   nom unic al directori de sortida -> copia a %TEMP% (si no, el Word obre el
#   fitxer en "Vista protegida" quan el desti es una unitat de xarxa) -> triar
#   el bloc de capcalera -> substituir els <<PLACEHOLDERS>> -> escriure el cos
#   -> desar, tancar i moure al desti.
#
#   $baseName : nom de fitxer ja calculat per la familia (cada una te el seu
#               patro; l'unic que comparteixen es la data al davant).
#   $capBloc  : quin bloc de '0 CAPCALERA.docx' (''=el generic, 'ACT_EXTR',
#               'LLIC'). Si el bloc no hi es, Select-CapcaleraBlock es queda amb
#               el generic i l'informe surt igualment.
#   $cos      : scriptblock que rep la Selection ja col·locada al final del
#               document i hi escriu el cos.
#
# COMPTE: $cos NO ha de portar .GetNewClosure(). Ha de veure els locals del
# Build-* que el crea en TEMPS D'EXECUCIO, i .GetNewClosure() en copiaria els
# VALORS del moment de crear-lo (vegeu CLAUDE.md, "les dues cares de la
# closure").
#
# Si el cos peta, el document es TANCA abans de rellancar l'error: si no, es
# queda una instancia de Word amb un document obert i el %TEMP% brut. Nomes
# ACT_EXTR ho feia; les altres tres no.
function Write-InformeDocx($word, [string]$baseName, [string]$capBloc, $header, [scriptblock]$cos) {
    $targetDir = _ResolveOutputDir
    [string]$outPath = _GetUniqueOutputPath $targetDir $baseName
    $fileName = [System.IO.Path]::GetFileName($outPath)
    $tempPath = Join-Path $env:TEMP $fileName
    $doc = _OpenOutputDocument $word $tempPath
    try {
        Select-CapcaleraBlock $doc $capBloc
        Apply-HeaderReplacements -doc $doc -header $header

        $doc.Activate()
        $sel = $word.Selection
        [void]$sel.EndKey(6)   # wdStory = 6

        & $cos $sel

        $doc.Save()
        $doc.Close($false)
    } catch {
        try { $doc.Close($false) } catch { }
        throw
    }
    # Al desti final (xarxa o local). Si no s'hi pot moure, es queda el temporal
    # i es retorna la seva ruta: val mes un informe a %TEMP% que cap informe.
    try { Move-Item -LiteralPath $tempPath -Destination $outPath -Force } catch { return $tempPath }
    return $outPath
}

# ----------------------------------------------------------------------------
# ESCRIURE UNA LINIA DE CATALEG
# ----------------------------------------------------------------------------
# Una linia del cataleg pot portar text i enllacos barrejats ("[[URL]] ..."). El
# motor els separa: el text va com a cos i CADA enllac com a hipervincle en
# paragraf propi.
#
# N'hi havia TRES copies -$emitLine (Document.ps1), _LlicEmetLinia (Llicencia.ps1)
# i _VLine (VistaWord.ps1)-, i les dues primeres nomes es diferenciaven en si
# deduplicaven els enllacos o no.
#
#   -IsChild : sagnia de sub-nivell (el cos i l'enllac d'un fill).
#   $vistos  : conjunt d'enllacos ja emesos EN AQUEST PUNT, per no repetir-los.
#              A Llicencia cal: el text de REQ1 i el comentari "No es disposa..."
#              solen portar el mateix enllac i sortia dues vegades seguides. Amb
#              $null (REQ1) no es dedupa res, que es el comportament de sempre.
#   $emesos  : on s'apunten els enllacos que s'han arribat a escriure. Serveix a
#              _LlicEscriuPunt per saber quins ha de deixar per despres del
#              comentari (l'enllac va DESPRES de la frase que l'anuncia).
#
# La linia ha d'arribar JA RESOLTA (els [CAMP:]/[OPCIO:] es resolen per BLOC,
# no linia a linia: vegeu Apply-FieldsToLines).
#
# ELS ENLLACOS ES DETECTEN AMB _SplitTextAndUrls, MAI A MA. Llicencia va tenir
# un _EsUrl fet amb -like '[[URL]]*', i en un patro de -like '[[URL]' es una
# CLASSE DE CARACTERS: no coincidia mai, el marcador [[URL]] sortia TAL QUAL a
# l'informe i l'enllac no era hipervincle. (Vegeu CLAUDE.md: aquesta trampa ja
# ha sortit tres vegades en aquest projecte.)
function Write-Linia($sel, [string]$linia, [switch]$IsChild, $vistos = $null, $emesos = $null) {
    if ([string]::IsNullOrWhiteSpace($linia)) { return }
    $parts = _SplitTextAndUrls $linia
    if (-not [string]::IsNullOrWhiteSpace($parts.Text)) {
        if ($IsChild) { Format-Body $sel $parts.Text -IsChild } else { Format-Body $sel $parts.Text }
    }
    foreach ($u in @($parts.Urls)) {
        $clau = ([string]$u).Trim()
        if ($null -ne $vistos -and $vistos.Contains($clau)) { continue }
        if ($null -ne $vistos) { [void]$vistos.Add($clau) }
        if ($null -ne $emesos) { [void]$emesos.Add($clau) }
        if ($IsChild) { Format-Url $sel $u -IsChild } else { Format-Url $sel $u }
    }
}
