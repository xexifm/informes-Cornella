#requires -Version 5.1
<#
.SYNOPSIS
  Format del document Word per als informes.

.DESCRIPTION
  Modul de format reutilitzable per a diferents tipus d'informes. Es carrega
  via dot-sourcing (. .\Format.ps1) i ofereix funcions per escriure els
  components estructurats del document a partir d'una Selection de Word
  ja preparada al final del cos.

  Tota la configuracio (sangries, mides, espaiat) es centralitza a la
  variable $ReportFormatConfig. Es pot sobreescriure abans de cridar les
  funcions (per exemple, des d'un altre script base per a un altre tipus
  d'informe).

  Punts d'entrada principals:
    Format-Section $sel "..."        -> titol de seccio (MAJUSCULES, sense negreta)
    Format-Subsection $sel "..."     -> subseccio (subratllat)
    Format-Item $sel "1." "..."      -> item numerat (numero en negreta + text)
    Format-Item $sel "1.1." "..." -IsChild   -> sub-item amb sangria
    Format-Body $sel "..."           -> linia normal d'un item
    Format-Url  $sel "https://..."   -> URL com a hyperlink
    Format-Spacer $sel               -> linia buida per separar blocs
    Format-Conclusion $sel "..."     -> paragraf de conclusio

  Totes les funcions Format-X comencen amb TypeParagraph() i fan reset
  del format de caracter (negreta/cursiva/subratllat/mida) abans d'aplicar
  l'estil propi.

.NOTES
  Les sangries es defineixen en centimetres a $ReportFormatConfig. La
  funcio _CmToPoints fa la conversio a punts que utilitza Word internament
  (1 cm = 28.346 punts).
#>

# Configuracio per defecte. Es pot modificar abans de cridar Build-Document.
$Script:ReportFormatConfig = @{
    # Mides de font
    BodyFontSize         = 11
    UrlFontSize          = 10

    # TIPOGRAFIA BASE del document. A l'informe ve heretada de la plantilla
    # (Build-Document COPIA '0 CAPCALERA.docx', que porta Bookman Old Style i
    # justificat), pero un document NOU de Word faria servir Calibri alineat a
    # l'esquerra. Per aixo es declara aqui: aixi les VISTES dels catalegs
    # (VistaWord.ps1) es veuen igual que l'informe. Els valors son EXACTAMENT
    # els de la plantilla (word/document.xml + sectPr).
    BodyFontName         = 'Bookman Old Style'
    BodyAlignment        = 3        # 3 = wdAlignParagraphJustify (a la plantilla: jc="both")
    BaseLineSpacing      = 1.15     # plantilla: w:line="276" lineRule="auto" -> 1,15 linies
    # Marges de pagina en PUNTS (plantilla, en twips: 1417/849/993/1701; 20 twips = 1 pt)
    PageMarginTopPt      = 70.85
    PageMarginRightPt    = 42.45
    PageMarginBottomPt   = 49.65
    PageMarginLeftPt     = 85.05

    # Sangries (cm) a l'esquerra
    SectionIndentCm      = 0
    SubsectionIndentCm   = 0
    ItemIndentCm         = 0
    ChildIndentCm        = 1
    ConclusionIndentCm   = 0

    # Vinyetes (Format-Bullet): sangria francesa (hanging) i espaiat propi,
    # per reproduir el format de llista de l'informe favorable (pic al primer
    # nivell de sangria i text al segon; separacio per SpaceBefore, sense
    # linies en blanc entre punts). Valors en cm / punts.
    BulletIndentCm       = 1.25   # sangria esquerra del text (1r nivell)
    BulletHangCm         = 0.62   # sangria francesa (el pic queda a l'esquerra)
    # Sub-nivell (els FILLS d'un item numerat). El text s'alinea a 1 cm, igual
    # que ChildIndentCm, que es el que fan servir les sub-linies i els enllacos
    # del fill (Format-Body/-Url -IsChild): aixi tot el bloc del fill queda
    # alineat. Al XML: w:ind left="567" hanging="283".
    BulletChildIndentCm  = 1.0
    BulletChildHangCm    = 0.5
    BulletSpaceBeforePt  = 6      # separacio entre punts (en lloc de linia buida)
    NoteIndentCm         = 1.25   # sub-paragraf sagnat sense pic (Format-Note)
    LabelSpaceAfterPt    = 12     # espai sota una etiqueta de subseccio (Format-Label)

    # Separacio entre un item numerat i el seu PRIMER sub-punt. Sense aixo el
    # sub-punt queda enganxat al text de l'item (12 pt = 240 twips al XML).
    # Els sub-punts entre ells segueixen amb BulletSpaceBeforePt.
    ItemSpaceAfterPt        = 12

    # Anotacions datades de l'informe de SEGUIMENT. Seguiment.ps1 escriu el XML
    # directament (sense Word), pero els valors de FORMAT viuen aqui: el format
    # del document es cosa d'aquest modul.
    #   - L'anotacio d'un SUB-PUNT s'alinea amb els punts de llista
    #     (1,25 cm = 709 twips, el mateix que BulletIndentCm). Cal posar-la
    #     EXPLICITA perque, en forcar numId=0 perque l'anotacio no s'enumeri, es
    #     perd la sagnia que aportava la numeracio.
    #   - L'anotacio d'un requeriment de primer nivell NO se sagna.
    AnnotationIndentCm      = 1.25
    AnnotationSpaceBeforePt = 10
    AnnotationSpaceAfterPt  = 12

    # Espaiat (linia buida entre elements)
    SpacerAfterIntroParagraph     = $true   # despres de la frase intro del cataleg
    SpacerAfterSection            = $true   # entre seccio i el que ve a sota
    SpacerAfterSubsection         = $true   # entre subseccio i el que ve a sota
    SpacerAfterIntro              = $true   # entre intro i el primer item
    SpacerAfterItem               = $true   # despres de cada item complet
    SpacerBeforeConclusionsBlock  = $true

    # Separacio entre conclusions: en lloc d'inserir paragrafs buits
    # (que el Word pot col·lapsar visualment), apliquem "Space After"
    # propi a cada conclusio. Mes robust i sempre visible.
    ConclusionSpaceAfterPt        = 12      # punts despres de cada conclusio (0 = enganxades)
}

function _CmToPoints { param([double]$cm) return ($cm * 28.346456692913385) }

# Conversions a TWIPS (la unitat del XML del Word: w:ind, w:spacing). Les fa
# servir Seguiment.ps1, que escriu el XML directament pero ha de respectar els
# valors d'aquesta configuracio de format.
#   1 polzada = 1440 twips = 2,54 cm = 72 pt  ->  1 cm = 566,93 twips; 1 pt = 20.
function _PtToTwips { param([double]$pt) return [int][Math]::Round($pt * 20) }
function _CmToTwips { param([double]$cm) return [int][Math]::Round($cm * 1440 / 2.54) }

function _Reset-Char($sel) {
    $sel.Font.Bold = 0
    $sel.Font.Italic = 0
    $sel.Font.Underline = 0  # wdUnderlineNone
    $sel.Font.Size = $Script:ReportFormatConfig.BodyFontSize
    # Tipus de lletra EXPLICIT a cada text: mai ha de sortir la Calibri del tema
    # d'un document nou. A l'informe coincideix amb el que ja hereta de la
    # plantilla, o sigui que no en canvia res.
    try { $sel.Font.Name = $Script:ReportFormatConfig.BodyFontName } catch { }
}

function _Apply-Indent($sel, $cm) {
    $sel.ParagraphFormat.LeftIndent = (_CmToPoints $cm)
    $sel.ParagraphFormat.FirstLineIndent = 0
    # Justificat EXPLICIT (com la plantilla). Els qui volen una altra cosa
    # (Format-ConclusionHeader, centrat) l'apliquen DESPRES d'aquesta crida.
    try { $sel.ParagraphFormat.Alignment = $Script:ReportFormatConfig.BodyAlignment } catch { }
    # Reset de l'espaiat propi (SpaceBefore/After) per no heretar el de les
    # vinyetes (Format-Bullet/Note) o el d'una etiqueta (Format-Label) quan ve
    # un paragraf normal a continuacio.
    try { $sel.ParagraphFormat.SpaceBefore = 0 } catch { }
    try { $sel.ParagraphFormat.SpaceAfter  = 0 } catch { }
}

# Deixa un document NOU amb la MATEIXA base que la plantilla de l'informe:
# Bookman Old Style, cos 11, justificat, interlineat 1,15 i els marges de la
# plantilla. Sense aixo, un document creat amb Documents.Add() surt en Calibri
# alineat a l'esquerra i no s'assembla gens a l'informe.
#
# A l'informe NO cal cridar-la (Build-Document copia '0 CAPCALERA.docx' i ja ho
# hereta tot); qui la fa servir es el generador de VISTES dels catalegs.
function Format-ApplyBaseStyle($doc) {
    $cfg = $Script:ReportFormatConfig
    try {
        $normal = $doc.Styles.Item(-1)          # -1 = wdStyleNormal
        $normal.Font.Name = $cfg.BodyFontName
        $normal.Font.Size = $cfg.BodyFontSize
        $normal.ParagraphFormat.Alignment = $cfg.BodyAlignment
        $normal.ParagraphFormat.SpaceBefore = 0
        $normal.ParagraphFormat.SpaceAfter = 0
        # wdLineSpaceMultiple = 5; amb aquesta regla, LineSpacing va en punts on
        # 12 pt = 1 linia (1,15 linies -> 13,8).
        $normal.ParagraphFormat.LineSpacingRule = 5
        $normal.ParagraphFormat.LineSpacing = ([double]$cfg.BaseLineSpacing * 12)
    } catch { }
    # Que la Calibri del TEMA no s'escoli enlloc.
    try {
        $doc.Content.Font.Name = $cfg.BodyFontName
        $doc.Content.Font.Size = $cfg.BodyFontSize
    } catch { }
    try {
        $ps = $doc.PageSetup
        $ps.TopMargin    = [double]$cfg.PageMarginTopPt
        $ps.RightMargin  = [double]$cfg.PageMarginRightPt
        $ps.BottomMargin = [double]$cfg.PageMarginBottomPt
        $ps.LeftMargin   = [double]$cfg.PageMarginLeftPt
    } catch { }
}

function Format-Section {
    param($sel, $text)
    [void]$sel.TypeParagraph()
    _Reset-Char $sel
    _Apply-Indent $sel $Script:ReportFormatConfig.SectionIndentCm
    # Seccions: sense negreta, en MAJUSCULES.
    if ($text) { $sel.TypeText(([string]$text).ToUpper()) }
}

function Format-Subsection {
    param($sel, $text)
    [void]$sel.TypeParagraph()
    _Reset-Char $sel
    _Apply-Indent $sel $Script:ReportFormatConfig.SubsectionIndentCm
    $sel.Font.Underline = 1
    $sel.TypeText([string]$text)
    $sel.Font.Underline = 0
}

function Format-Item {
    param($sel, [string]$number, [string]$text, [switch]$IsChild)
    [void]$sel.TypeParagraph()
    _Reset-Char $sel
    $indent = if ($IsChild) { $Script:ReportFormatConfig.ChildIndentCm }
              else          { $Script:ReportFormatConfig.ItemIndentCm }
    _Apply-Indent $sel $indent
    # El numero s'escriu SENSE negreta i despres es posa en negreta pel RANG.
    #
    # Abans es feia amb $sel.Font.Bold = 1 / TypeText / = 0. El problema: el
    # "Bold = 0" actua sobre el FORMAT D'ESCRIPTURA del punt d'insercio, i el
    # Word no sempre l'hi aplica (depen del que hi hagi al paragraf anterior).
    # Quan no l'aplicava, tot el text de l'item sortia en negreta i el numero i
    # el text quedaven en un sol <w:r> (es veia, p.ex., a tots els items de
    # CONTROLS INICIALS i CONTROLS PERIODICS de la vista de REQ1). Fent-ho pel
    # rang, la negreta NOMES pot tocar el numero: el punt d'insercio no queda
    # mai en negreta i el cos no se la pot encomanar. La negreta INLINE del cos
    # (**...** de Type-RichText) no es toca.
    $numStart = $sel.Range.Start
    $sel.TypeText("$number ")
    $numEnd = $sel.Range.End
    if ($text) { Type-RichText $sel $text }
    if ($numEnd -gt $numStart) {
        try { $sel.Document.Range($numStart, $numEnd).Font.Bold = $true } catch { }
    }
}

# -Bold: negreta a TOT el paragraf. Nomes el fa servir l'informe de LLICENCIA
# (el comentari "No es disposa..." de cada punt). Per defecte no s'hi toca res,
# o sigui que enlloc mes no canvia absolutament res.
#
# S'aplica al RANG que s'acaba d'escriure, no amb $sel.Font.Bold abans de
# teclejar: aixo ultim toca el format del PUNT D'INSERCIO, i el Word no sempre
# l'aplica (es la mateixa trampa que feia sortir en negreta tot un item a
# Format-Item).
function Format-Body {
    param($sel, [string]$text, [switch]$IsChild, [switch]$Bold)
    [void]$sel.TypeParagraph()
    _Reset-Char $sel
    $indent = if ($IsChild) { $Script:ReportFormatConfig.ChildIndentCm }
              else          { $Script:ReportFormatConfig.ItemIndentCm }
    _Apply-Indent $sel $indent
    $ini = $sel.Range.Start
    if ($text) { Type-RichText $sel $text }
    $fi = $sel.Range.End
    if ($fi -gt $ini) {
        try {
            $rng = $sel.Document.Range($ini, $fi)
            if ($Bold) { $rng.Font.Bold = $true }
        } catch { }
    }
}

# Item amb pic (vinyeta) en lloc de numero. S'usa per a llistes que han d'anar
# amb punts i no numerades (p.ex. l'informe favorable d'activitat
# extraordinaria). El pic es escriu com a text (com el numero a Format-Item),
# coherent amb la manera com el motor marca les llistes (sense numeracio
# automatica de Word). -IsChild aplica la sangria de sub-nivell.
#
# -First marca el PRIMER punt d'una llista (el que penja directament d'un item
# numerat): en lloc de la separacio curta entre punts (BulletSpaceBeforePt) hi
# posa ItemSpaceAfterPt, perque el sub-punt no quedi enganxat al text de l'item.
function Format-Bullet {
    param($sel, [string]$text, [switch]$IsChild, [switch]$First)
    [void]$sel.TypeParagraph()
    _Reset-Char $sel
    # Sangria francesa: el text va a LeftIndent i el pic queda a
    # LeftIndent-Hang (a l'esquerra). El tabulador despres del pic salta al
    # LeftIndent (el Word posa una parada de tabulacio implicita alli).
    $left = if ($IsChild) { $Script:ReportFormatConfig.BulletChildIndentCm }
            else          { $Script:ReportFormatConfig.BulletIndentCm }
    $hang = if ($IsChild) { $Script:ReportFormatConfig.BulletChildHangCm }
            else          { $Script:ReportFormatConfig.BulletHangCm }
    $sel.ParagraphFormat.LeftIndent = (_CmToPoints $left)
    $sel.ParagraphFormat.FirstLineIndent = (- (_CmToPoints $hang))
    # Text justificat (com l'estil 'List Paragraph' de la casa, jc=both). Es
    # posa explicit per no heretar un 'center' d'una capcalera anterior.
    try { $sel.ParagraphFormat.Alignment = 3 } catch { }   # 3 = wdAlignParagraphJustify
    # Separacio entre punts amb SpaceBefore (no linies en blanc): aixi la
    # llista surt compacta i amb el mateix aire que el document de referencia.
    # El primer punt de la llista se separa mes (ItemSpaceAfterPt) de l'item.
    $before = if ($First) { $Script:ReportFormatConfig.ItemSpaceAfterPt }
              else        { $Script:ReportFormatConfig.BulletSpaceBeforePt }
    try { $sel.ParagraphFormat.SpaceBefore = [double]$before } catch { }
    # Pic Unicode (U+2022) escrit per codepoint per no dependre de l'encoding.
    $sel.TypeText([string]([char]0x2022) + "`t")
    if ($text) { Type-RichText $sel $text }
}

# Sub-paragraf sagnat SENSE pic (p.ex. la nota del "Dret d'admissio" a
# l'informe favorable): mateixa sangria que un punt pero sense vinyeta.
function Format-Note {
    param($sel, [string]$text)
    [void]$sel.TypeParagraph()
    _Reset-Char $sel
    _Apply-Indent $sel $Script:ReportFormatConfig.NoteIndentCm
    try { $sel.ParagraphFormat.Alignment = 3 } catch { }   # 3 = wdAlignParagraphJustify
    try { $sel.ParagraphFormat.SpaceBefore = [double]$Script:ReportFormatConfig.BulletSpaceBeforePt } catch { }
    if ($text) { Type-RichText $sel $text }
}

# Etiqueta de subseccio dins del cos (p.ex. "RETOLS INFORMATIUS"): text normal
# (no negreta) amb un espai a sota per separar-la del que ve. S'usa quan una
# seccio te un rotul propi seguit del seu contingut sense linia en blanc.
function Format-Label {
    param($sel, [string]$text)
    [void]$sel.TypeParagraph()
    _Reset-Char $sel
    _Apply-Indent $sel 0
    try { $sel.ParagraphFormat.Alignment = 3 } catch { }   # 3 = wdAlignParagraphJustify
    try { $sel.ParagraphFormat.SpaceAfter = [double]$Script:ReportFormatConfig.LabelSpaceAfterPt } catch { }
    if ($text) { Type-RichText $sel $text }
}

function Format-Url {
    param($sel, [string]$url, [switch]$IsChild)
    [void]$sel.TypeParagraph()
    _Reset-Char $sel
    $indent = if ($IsChild) { $Script:ReportFormatConfig.ChildIndentCm }
              else          { $Script:ReportFormatConfig.ItemIndentCm }
    _Apply-Indent $sel $indent
    $sel.Font.Size = $Script:ReportFormatConfig.UrlFontSize
    $startPos = $sel.Range.Start
    $sel.TypeText($url)
    $endPos = $sel.Range.End
    try {
        $doc = $sel.Document
        $rng = $doc.Range($startPos, $endPos)
        [void]$doc.Hyperlinks.Add($rng, $url)
    } catch { }
    $sel.Font.Size = $Script:ReportFormatConfig.BodyFontSize
}

function Format-Spacer {
    param($sel)
    [void]$sel.TypeParagraph()
    _Reset-Char $sel
    _Apply-Indent $sel 0
}

function Format-Conclusion {
    param($sel, [string]$text)
    [void]$sel.TypeParagraph()
    _Reset-Char $sel
    _Apply-Indent $sel $Script:ReportFormatConfig.ConclusionIndentCm
    # Alignment explicit per no heretar el "center" que deixa el
    # Format-ConclusionHeader (CONCLUSIONS centrat). Volem justified.
    try { $sel.ParagraphFormat.Alignment = 3 } catch { }   # 3 = wdAlignParagraphJustify
    # "Space After" propi del paragraf (en punts). Es robust a la
    # compactacio visual del Word que feia que els paragrafs buits no
    # es veiessin entre conclusions.
    $sa = [double]$Script:ReportFormatConfig.ConclusionSpaceAfterPt
    try { $sel.ParagraphFormat.SpaceAfter = $sa } catch { }
    if ($text) { Type-RichText $sel $text }
}

# Titol del bloc de conclusions: text centrat i en negreta.
# NO fem reset d'alignment al final (anava AL MATEIX paragraf i feia que
# 'CONCLUSIONS' sortis alineat a l'esquerra). En lloc d'aixo, Format-Conclusion
# i la resta de Format-* es defineixen amb el seu Alignment explicit.
function Format-ConclusionHeader {
    param($sel, [string]$text)
    [void]$sel.TypeParagraph()
    _Reset-Char $sel
    _Apply-Indent $sel 0
    try { $sel.ParagraphFormat.Alignment = 1 } catch { }   # 1 = wdAlignParagraphCenter
    try { $sel.ParagraphFormat.SpaceAfter = 12 } catch { }
    $sel.Font.Bold = 1
    if ($text) { $sel.TypeText($text) }
    $sel.Font.Bold = 0
}

# Type-RichText: escriu text al document interpretant marcadors inline:
#   **negreta**   -> text en negreta
#   //cursiva//   -> text en cursiva
# La resta s'escriu normal. Es respecta el format de paragraf actual.
function Type-RichText {
    param($sel, [string]$text)
    if ([string]::IsNullOrEmpty($text)) { return }
    # Regex: captura segments alternatius (text normal o marcat).
    # Es no-greedy per als marcadors.
    $pattern = '\*\*(.+?)\*\*|//(.+?)//'
    $rx = [regex]$pattern
    $pos = 0
    foreach ($m in $rx.Matches($text)) {
        if ($m.Index -gt $pos) {
            $sel.TypeText($text.Substring($pos, $m.Index - $pos))
        }
        if ($m.Groups[1].Success) {
            $sel.Font.Bold = 1
            $sel.TypeText($m.Groups[1].Value)
            $sel.Font.Bold = 0
        } elseif ($m.Groups[2].Success) {
            $sel.Font.Italic = 1
            $sel.TypeText($m.Groups[2].Value)
            $sel.Font.Italic = 0
        }
        $pos = $m.Index + $m.Length
    }
    if ($pos -lt $text.Length) {
        $sel.TypeText($text.Substring($pos))
    }
}
