#requires -Version 5.1
<#
  SegonPla.ps1 - Llancar un script de suport/ EN SEGON PLA (un proces a part,
  sense finestra). Abans vivia a Motor.ps1; el Planol activitats (procés de
  rutes/, que no pot carregar Motor.ps1) tambe en llanca un i les cometes del
  Start-Process no poden estar escrites dues vegades.

  NOMES DEFINEIX FUNCIONS.
#>

# ----------------------------------------------------------------------------
# Llancar un script d'aquest mateix 'suport' EN SEGON PLA
# ----------------------------------------------------------------------------
# Hi ha feines que no poden fer esperar l'usuari i que no ensenyen res: refer
# les VISTES en Word dels catalegs (l'editor, en tancar-se) i la passada
# AUTOMATICA de "Copiar informes" (el menu, cada dia a les 13:00). Totes dues
# es llancaven -o s'haurien llancat- amb les mateixes cinc linies, i la trampa
# de les cometes es prou fina per no tenir-la escrita dues vegades:
#
#   A PowerShell 5.1, Start-Process -ArgumentList NO enquota els elements, i el
#   clone de l'usuari te espais a la ruta (vegeu _ArgvToCommandLine a
#   PdfSignar.ps1, mateixa trampa). Les cometes les hi posem nosaltres.
#
# -WindowStyle Hidden: el PowerShell 5.1 crea la seva consola i despres l'amaga,
# o sigui que es pot veure una llampada; per a un proces que no torna cap
# resposta a l'usuari es acceptable (el llancador del programa, que si que
# s'obre a ma, va per wscript.exe justament per no ensenyar-la).
#
# Retorna l'objecte Process (per poder saber si encara corre) o $null si no ha
# pogut arrencar. Mai llanca: cap d'aquestes feines no es critica.
#
# $nomScript: relatiu a suport/ (la carpeta d'aquest fitxer) o una ruta sencera.
# $arguments: es passen darrere, CADA UN entre cometes (la mateixa trampa).
function Start-ScriptSegonPla([string]$nomScript, [string[]]$arguments = @()) {
    try {
        $script = if ([System.IO.Path]::IsPathRooted($nomScript)) { $nomScript } else { [System.IO.Path]::Combine($PSScriptRoot, $nomScript) }
        if (-not (Test-Path -LiteralPath $script)) { return $null }
        # $args NO: es una variable AUTOMATICA de PowerShell (els arguments de
        # la funcio) i assignar-la dins d'una funcio es demanar problemes.
        $argv = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $script + '"'))
        foreach ($a in @($arguments)) { if ($null -ne $a) { $argv += ('"' + [string]$a + '"') } }
        return (Start-Process -FilePath 'powershell.exe' -ArgumentList $argv -WindowStyle Hidden -PassThru)
    } catch { return $null }
}
