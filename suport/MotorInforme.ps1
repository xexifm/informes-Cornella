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
