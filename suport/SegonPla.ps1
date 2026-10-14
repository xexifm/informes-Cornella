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

# ----------------------------------------------------------------------------
# OBRIR UN FITXER O UNA CARPETA PER A L'USUARI, DES D'UN ALTRE PROCES
# ----------------------------------------------------------------------------
# L'usuari (octubre 2026): "no em deixa fer servir finestres fora de
# l'aplicacio: la carpeta de Windows; quan ha generat l'informe no em deixa
# anar a la finestra de Word de l'informe (he de tancar el programa)".
#
# Obrir un document amb Invoke-Item / Start-Process <fitxer> fa el
# ShellExecute DINS d'aquest proces, i per a un .docx aixo es una conversa DDE
# amb el Word que penja del fil de la finestra del programa. Just despres el
# programa obre un dialeg (el d'enviar el correu) o es posa a treballar (llegir
# l'Excel, el Word de l'informe seguent): el Word, i l'explorador que hi parla,
# es queden esperant aquest fil. Ara la conversa la fa un proces a part que
# acaba sol (rundll32 per als fitxers, explorer per a les carpetes): el
# programa ja no hi es pel mig. Hi ha guard: cap Invoke-Item ni Start-Process
# d'un fitxer de l'usuari fora d'aqui.
#
# Torna l'ordre que s'executara: @{ Exe; Args }. PURA.
function _ObrePerUsuariOrdre([string]$ruta, [bool]$esCarpeta, [string]$windir = $env:WINDIR) {
    if ([string]::IsNullOrWhiteSpace($windir)) { $windir = 'C:\Windows' }
    $r = ([string]$ruta).Trim().Trim('"')
    if ($esCarpeta) { return @{ Exe = ($windir.TrimEnd('\') + '\explorer.exe'); Args = ('"' + $r + '"') } }
    return @{ Exe = ($windir.TrimEnd('\') + '\System32\rundll32.exe'); Args = ('shell32.dll,ShellExec_RunDLL "' + $r + '"') }
}

# Obre $ruta (fitxer o carpeta) com si l'usuari hi hagues fet doble clic. Torna
# $true si l'ha pogut llancar. Si el proces a part no arrenca, ho fa com abans.
function Open-PerUsuari([string]$ruta) {
    if ([string]::IsNullOrWhiteSpace($ruta)) { return $false }
    $esCarpeta = $false
    try { $esCarpeta = Test-Path -LiteralPath $ruta -PathType Container } catch { }
    $o = _ObrePerUsuariOrdre $ruta $esCarpeta
    try { Start-Process -FilePath $o.Exe -ArgumentList $o.Args | Out-Null; return $true } catch { }
    try { Invoke-Item -LiteralPath $ruta; return $true } catch { return $false }
}

# ----------------------------------------------------------------------------
# ESPERAR UN PROCES SENSE DEIXAR PENJAT EL WINDOWS
# ----------------------------------------------------------------------------
# $p.StandardOutput.ReadToEnd() bloqueja el fil de les finestres SENSE atendre
# cap missatge: mentre l'AutoFirma signa (5-35 s), qualsevol programa que
# envii un missatge a les finestres d'aquest proces (l'explorador quan obre un
# fitxer, el Word) es queda esperant. Llegir en segon pla (ReadToEndAsync) i
# esperar amb WaitForExit(ms) / Task.Wait(ms), que en un fil STA SI que atenen
# els missatges enviats, ho evita. Mateix resultat: la sortida sencera quan el
# proces (i qui hagi heretat la canonada, el Java de l'AutoFirma) acaba.
# Torna @{ Out; Err }.
function Wait-ProcesCapturat($p) {
    $tOut = $p.StandardOutput.ReadToEndAsync()
    $tErr = $p.StandardError.ReadToEndAsync()
    while (-not $p.WaitForExit(100)) { }
    while (-not $tOut.Wait(100)) { }
    while (-not $tErr.Wait(100)) { }
    $p.WaitForExit()
    return @{ Out = [string]$tOut.Result; Err = [string]$tErr.Result }
}
