' ============================================================================
'  Llancador del generador d'informes SENSE cap finestra de consola.
' ============================================================================
'  Per que un .vbs i no nomes el .bat?
'    Un .bat, en fer-hi doble clic, SEMPRE obre una consola (CMD). I encara que
'    llancem PowerShell amb -WindowStyle Hidden, el PowerShell 5.1 crea la seva
'    consola (blava) i despres l'amaga: es veu una llampada molesta.
'    En canvi, wscript.exe NO mostra cap finestra, i el tercer parametre de
'    Run (0 = amagada) fa que PowerShell no arribi a mostrar mai cap consola.
'    Resultat: nomes es veuen les finestres del propi programa. Els missatges i
'    errors ja es mostren en finestres (MessageBox), aixi que no cal cap consola.
'
'  Aquest .vbs viu a suport\ (al costat de GenerarInforme.ps1). El pot cridar
'  GenerarInforme.bat o s'hi pot fer doble clic directament.
' ============================================================================

Set fso = CreateObject("Scripting.FileSystemObject")
Dim here : here = fso.GetParentFolderName(WScript.ScriptFullName)
Dim ps1  : ps1  = fso.BuildPath(here, "GenerarInforme.ps1")

Set sh = CreateObject("WScript.Shell")
sh.CurrentDirectory = here
sh.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & ps1 & """", 0, False
