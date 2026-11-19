#requires -Version 5.1
<#
.SYNOPSIS
  TANCAR DE DEBO el Word i l'Excel que obre el programa (per COM). Un sol lloc.

.DESCRIPTION
  L'usuari (octubre i novembre 2026, dues vegades): "quan s'ha generat un
  informe o similar no puc treballar a la carpeta de Windows perque me la
  bloqueja el programa [...] no em deixa fer res a la carpeta fins que no tanqui
  el programa"; i abans, "no em deixa anar a la finestra de Word de l'informe".

  LA CAUSA (la que encaixa amb el "fins que no tanqui el programa"): el programa
  feia $app.Quit() i ReleaseComObject NOMES de l'aplicacio. Pero cada document,
  full, rang o seleccio que ha tocat el PowerShell es una referencia COM mes
  (un RCW) que queda viva fins que el recollidor d'escombraries passa, i el
  PowerShell gairebe no el fa passar mai. Amb referencies vives, l'EXCEL.EXE
  (documentat per Microsoft) i el WINWORD.EXE no acaben de morir: es queden
  mig tancats, amagats, fins que el proces del programa s'acaba.
  I un Excel o un Word en aquest estat fa mal a FORA del programa: quan
  l'explorador obre un document (doble clic), el Windows en parla per DDE amb
  TOTES les finestres obertes (SendMessage, sense temps limit) i espera que
  contestin. La que no contesta deixa la carpeta penjada. I quan s'obre un
  .docx, el Word que l'ha d'obrir pot anar a parar a la instancia mig morta.

  EL REMEI, en un sol lloc (Close-AppOffice):
    1. Saber quin PROCES es (abans de tancar-lo): l'Excel ho diu ($app.Hwnd); al
       Word se li posa un titol unic en crear-lo i es busca la seva finestra.
    2. Quit, alliberar l'aplicacio DEL TOT (FinalReleaseComObject) i fer passar
       el recollidor DUES vegades (Collect + WaitForPendingFinalizers): aixo
       allibera tambe les referencies de documents, fulls i rangs.
    3. ESPERAR que el proces surti (amb un limit), atenent els missatges del
       Windows mentre s'espera (WaitForExit(ms) en un fil STA ho fa).
  I tot s'apunta a %LOCALAPPDATA%\InformesCornella\office-log.txt: si tornes a
  veure la carpeta penjada, aquell fitxer diu si un Word o un Excel no va voler
  sortir, i quin.

  NOMES DEFINEIX FUNCIONS (el carrega tambe rutes\Ruta.ps1, que corre en un
  proces a part i llegeix l'Excel). Hi ha guard: cap .Quit() fora d'aqui.
#>

# El titol que es posa al Word del programa per trobar-ne la finestra (i el
# proces). Es nomes de la sessio: el Word no el desa enlloc.
$Script:OfficeTitolPrefix = 'InformesCornella-automatitzacio-'
$Script:OfficeEsperaMs = 20000

function _OfficeWin32 {
    if ('CornellaApp.Office' -as [type]) { return $true }
    try {
        Add-Type -Namespace CornellaApp -Name Office -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
public static extern System.IntPtr FindWindow(string lpClassName, string lpWindowName);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint pid);
'@ -ErrorAction Stop
        return $true
    } catch { return $false }
}

function _OfficeLog([string]$msg) {
    try {
        $base = [string]$env:LOCALAPPDATA
        if ([string]::IsNullOrWhiteSpace($base)) { return }
        $dir = Join-Path $base 'InformesCornella'
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $p = Join-Path $dir 'office-log.txt'
        # Que no creixi sense fi: passat de 256 KB, es comenca de nou.
        if ((Test-Path -LiteralPath $p) -and (Get-Item -LiteralPath $p).Length -gt 262144) { Remove-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue }
        Add-Content -LiteralPath $p -Value ('[' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') + '] [PID ' + $PID + '] ' + $msg) -Encoding UTF8
    } catch { }
}

# El titol unic del Word del programa. PURA.
function _OfficeTitolNou { return ($Script:OfficeTitolPrefix + [guid]::NewGuid().ToString('N')) }

# El proces d'un Word o un Excel obert per COM, o 0 si no se sap. No llanca mai.
#   - Excel: Application.Hwnd (la finestra principal, encara que sigui amagada).
#   - Word: Application.Caption = titol unic (New-WordApp el posa) i la finestra
#     de classe "OpusApp" amb aquell titol.
function Get-PidAppOffice($app) {
    if ($null -eq $app) { return 0 }
    if (-not (_OfficeWin32)) { return 0 }
    $h = [System.IntPtr]::Zero
    try { $h = [System.IntPtr][int64]$app.Hwnd } catch { $h = [System.IntPtr]::Zero }
    if ($h -eq [System.IntPtr]::Zero) {
        $titol = ''
        try { $titol = [string]$app.Caption } catch { }
        if ($titol.StartsWith($Script:OfficeTitolPrefix)) {
            try { $h = [CornellaApp.Office]::FindWindow('OpusApp', $titol) } catch { }
        }
    }
    if ($h -eq [System.IntPtr]::Zero) { return 0 }
    $p = [uint32]0
    try { [void][CornellaApp.Office]::GetWindowThreadProcessId($h, [ref]$p) } catch { return 0 }
    return [int]$p
}

# Allibera DEL TOT un objecte COM (document, llibre, full, aplicacio). No llanca.
function _AlliberaCom($o) {
    if ($null -eq $o) { return }
    try {
        if ($o -is [System.Management.Automation.PSObject]) { $o = $o.psobject.BaseObject }
        if ([System.Runtime.InteropServices.Marshal]::IsComObject($o)) { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) }
    } catch { }
}

# Fa passar el recollidor perque surtin les referencies COM que queden (rangs,
# fulls, documents...). Dues vegades: la primera passada nomes les posa a la
# cua de finalitzacio.
function _RecollirCom {
    try {
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
    } catch { }
}

# TANCA un Word o un Excel obert pel programa i ESPERA que el proces surti.
# $nom: 'Word' / 'Excel' (per al registre). $procId: si el cridador ja el sap
# (si no, es busca ABANS de tancar). Torna $true si el proces ha sortit (o no
# se sabia quin era), $false si encara hi es passat el temps. No llanca mai.
# $esperaMs: quant s'espera com a molt (-1 = $Script:OfficeEsperaMs). Un Word que
# ja no respon (InformesEscaneig) es tanca sense esperar-lo.
function Close-AppOffice($app, [string]$nom = 'Office', [int]$procId = 0, [int]$esperaMs = -1) {
    if ($null -eq $app) { return $true }
    if ($esperaMs -lt 0) { $esperaMs = $Script:OfficeEsperaMs }
    if ($procId -le 0) { $procId = Get-PidAppOffice $app }
    try { [void]$app.Quit() } catch { }
    _AlliberaCom $app
    _RecollirCom
    if ($procId -le 0) {
        _OfficeLog ($nom + ': tancat (no se sap quin proces era; no es pot esperar).')
        return $true
    }
    $p = $null
    try { $p = [System.Diagnostics.Process]::GetProcessById($procId) } catch { $p = $null }
    if ($null -eq $p) { _OfficeLog ($nom + ' (PID ' + $procId + '): tancat.'); return $true }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $fora = $false
    try {
        # WaitForExit(ms) en un fil STA atén els missatges enviats a les
        # finestres del programa mentre espera: no es penja res mes per esperar.
        while (-not ($fora = $p.WaitForExit(200))) {
            if ($sw.ElapsedMilliseconds -ge $esperaMs) { break }
            # A mig cami, un altre cop: hi pot haver referencies que s'han
            # alliberat despres de la primera passada.
            if ($sw.ElapsedMilliseconds -ge 2000 -and $sw.ElapsedMilliseconds -lt 2200) { _RecollirCom }
        }
    } catch { $fora = $true }
    if ($fora) { _OfficeLog ($nom + ' (PID ' + $procId + '): tancat en ' + $sw.ElapsedMilliseconds + ' ms.') }
    else {
        _OfficeLog ($nom + ' (PID ' + $procId + '): NO HA SORTIT en ' + [int]($esperaMs / 1000) + ' s. Si la carpeta o el Word es pengen, es aquest proces (Gestor de tasques).')
    }
    return [bool]$fora
}
