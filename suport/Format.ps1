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
    BulletChildIndentCm  = 2.0    # sangria esquerra del text (sub-nivell)
    BulletHangCm         = 0.62   # sangria francesa (el pic queda a l'esquerra)
    BulletSpaceBeforePt  = 6      # separacio entre punts (en lloc de linia buida)
    NoteIndentCm         = 1.25   # sub-paragraf sagnat sense pic (Format-Note)

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

function _Reset-Char($sel) {
    $sel.Font.Bold = 0
    $sel.Font.Italic = 0
    $sel.Font.Underline = 0  # wdUnderlineNone
    $sel.Font.Size = $Script:ReportFormatConfig.BodyFontSize
}

function _Apply-Indent($sel, $cm) {
    $sel.ParagraphFormat.LeftIndent = (_CmToPoints $cm)
    $sel.ParagraphFormat.FirstLineIndent = 0
    # Reset del SpaceBefore per no heretar l'espaiat propi de les vinyetes
    # (Format-Bullet/Format-Note) quan ve un paragraf normal a continuacio.
    try { $sel.ParagraphFormat.SpaceBefore = 0 } catch { }
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
    $sel.Font.Bold = 1
    $sel.TypeText("$number ")
    $sel.Font.Bold = 0
    if ($text) { Type-RichText $sel $text }
}

function Format-Body {
    param($sel, [string]$text, [switch]$IsChild)
    [void]$sel.TypeParagraph()
    _Reset-Char $sel
    $indent = if ($IsChild) { $Script:ReportFormatConfig.ChildIndentCm }
              else          { $Script:ReportFormatConfig.ItemIndentCm }
    _Apply-Indent $sel $indent
    if ($text) { Type-RichText $sel $text }
}

# Item amb pic (vinyeta) en lloc de numero. S'usa per a llistes que han d'anar
# amb punts i no numerades (p.ex. l'informe favorable d'activitat
# extraordinaria). El pic es escriu com a text (com el numero a Format-Item),
# coherent amb la manera com el motor marca les llistes (sense numeracio
# automatica de Word). -IsChild aplica la sangria de sub-nivell.
function Format-Bullet {
    param($sel, [string]$text, [switch]$IsChild)
    [void]$sel.TypeParagraph()
    _Reset-Char $sel
    # Sangria francesa: el text va a LeftIndent i el pic queda a
    # LeftIndent-Hang (a l'esquerra). El tabulador despres del pic salta al
    # LeftIndent (el Word posa una parada de tabulacio implicita alli).
    $left = if ($IsChild) { $Script:ReportFormatConfig.BulletChildIndentCm }
            else          { $Script:ReportFormatConfig.BulletIndentCm }
    $sel.ParagraphFormat.LeftIndent = (_CmToPoints $left)
    $sel.ParagraphFormat.FirstLineIndent = (- (_CmToPoints $Script:ReportFormatConfig.BulletHangCm))
    # Alineacio explicita a l'esquerra (per no heretar un 'center'/'justify'
    # d'un paragraf anterior, p.ex. una capcalera de conclusions).
    try { $sel.ParagraphFormat.Alignment = 0 } catch { }   # 0 = wdAlignParagraphLeft
    # Separacio entre punts amb SpaceBefore (no linies en blanc): aixi la
    # llista surt compacta i amb el mateix aire que el document de referencia.
    try { $sel.ParagraphFormat.SpaceBefore = [double]$Script:ReportFormatConfig.BulletSpaceBeforePt } catch { }
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
    try { $sel.ParagraphFormat.Alignment = 0 } catch { }   # 0 = wdAlignParagraphLeft
    try { $sel.ParagraphFormat.SpaceBefore = [double]$Script:ReportFormatConfig.BulletSpaceBeforePt } catch { }
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
