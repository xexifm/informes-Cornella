#requires -Version 5.1
<#
.SYNOPSIS
  MOTOR del generador d'informes de l'Ajuntament de Cornella.

.DESCRIPTION
  Aquest fitxer NOMES DEFINEIX: funcions, rutes i configuracio. Carregar-lo
  (dot-source) no obre cap finestra ni genera cap informe. Qui l'arrenca es
  GenerarInforme.ps1 (el punt d'entrada, que fa doble clic l'usuari); qui el
  reutilitza com a biblioteca son mobil/Vigilant.ps1, mobil/ExportaDades.ps1 i
  les proves de suport/tests/.

  ABANS, motor i punt d'entrada eren el mateix fitxer, i per poder-lo
  reutilitzar sense que s'executes calia fer-lo passar per proves
  ($env:GENINFORME_TEST = '1'). Amb la separacio, la bandera de proves ja
  nomes vol dir "no carreguis WinForms".

  Flux del programa (el condueix Main, definit aqui):
    1. L'usuari escull un cataleg de deficiencies (ESTRUCTURALS\REQ1.json...).
    2. Demana les dades de la capcalera (ID GIA, EXP_NUM, etc.).
    3. Mostra un TreeView amb les seccions/items del cataleg; l'usuari marca
       quines defciencies aplicaran a l'informe. En marcar-ne una, el seu text
       surt al panell de detall i, si conte [OPCIO:]/[CAMP:], s'omplen INLINE
       (alla mateix), sense un pas separat de camps.
    4. Mostra les conclusions disponibles (per tipus d'informe) amb el cos
       sencer; l'usuari en tria i n'omple els [OPCIO:]/[CAMP:] inline.
    5. Composa el document final:
         - Capcalera (ESTRUCTURALS\0 CAPCALERA.docx) amb valors substituits.
         - Per cada seccio escollida: titol Heading 1 + items numerats
           globalment 1..N (o, en informes de cos fix, el text tal qual).
         - Conclusions seleccionades.

.NOTES
  Configuracio: les rutes i constants es defineixen al fitxer config.ps1
  (opcional) al costat del .ps1; si no existeix, s'usen els valors per
  defecte definits a sota.

  Estat a %LOCALAPPDATA%\InformesCornella\ (mai al repositori):
    lastreport.json  dades de l'ULTIM informe generat amb exit, per poder-lo
                     replicar des del Pas 2.
    running.pid      PID de la instancia viva (instancia unica + perque
                     Actualitzar.bat pugui tancar el programa abans d'actualitzar).
    settings.json    rutes d'AQUEST ordinador (Settings.ps1).
  NO hi ha cap "sessio recuperable" pas a pas: si es tanca el programa a mig
  assistent, es torna a comencar (amb l'ultim informe com a plantilla).

  Cache: el resultat del parseig del cataleg .docx es guarda a
  %LOCALAPPDATA%\InformesCornella\cache\<basename>.json amb un hash
  del fitxer com a clau de validesa.

  Convencions del cataleg (REQ1.json i seguents):
    - Heading 1  -> titol de seccio.
    - Heading 2  -> nom curt de l'item (per al TreeView). Si comenca per
                    "::CHILD:: " es tracta d'un sub-bullet (fill de l'item
                    Heading 2 anterior); el prefix es elimina abans de mostrar.
    - Normal     -> cos de l'item: la primera linia es el text principal, les
                    seguents son URLs o complements (es mantenen tal qual).

  Placeholders al cos:
    [CAMP: nom]                 -> demana 'nom'
    [CAMP: nom (hint d'ajuda)]  -> demana 'nom', el hint apareix sota el camp
    Mateix nom = mateix valor (es demana un sol cop).
#>


$ErrorActionPreference = 'Stop'

# IMPORTANT (rendiment): a Windows PowerShell 5.1, Invoke-RestMethod/
# Invoke-WebRequest dibuixen una barra de progres ("Llegint resposta web...")
# que actualitza byte a byte i fa que una descarrega de pocs KB trigui SEGONS.
# Silenciant el progres, les crides a Google Drive (p.ex. llegir activitats.json
# per saber-ne la data) passen a ser gairebe instantanies. Ho posem global
# perque afecti tambe DriveApi.ps1 (que es dot-source a la mateixa sessio).
$global:ProgressPreference = 'SilentlyContinue'

# Mode "sense interficie": NO carreguem WinForms/Drawing. Nomes afecta el
# dibuix; les funcions del motor es defineixen sempre. Dues maneres d'activar-lo:
#   $env:GENINFORME_TEST  -> proves automatiques (poden correr en un Linux sense
#                            Windows/Office, per provar les funcions pures).
#   $MotorSenseGui = $true -> scripts de consola que reutilitzen el motor com a
#                            biblioteca (mobil/Vigilant.ps1, mobil/ExportaDades.ps1).
#                            El posen ABANS del dot-source d'aquest fitxer.
# El motor no executa res per si sol, aixi que aquesta bandera NOMES parla de
# la interficie: ja no serveix per evitar que s'arrenqui el programa.
$Script:HeadlessTest = [bool]$env:GENINFORME_TEST -or [bool]$MotorSenseGui
if (-not $Script:HeadlessTest) {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()
}

$ScriptRoot      = Split-Path -Parent $MyInvocation.MyCommand.Path
# Arrel del clone (un nivell amunt: suport/.. = informes-Cornella/).
# Aixi pots moure la carpeta informes-Cornella on vulguis i tot segueix
# funcionant; nomes la base de dades d'activitats (ActivitatsDir) es una
# ruta absoluta externa que no es mou.
$RepoRoot        = Split-Path -Parent $ScriptRoot

# ----------------------------------------------------------------------------
# Helpers d'interficie compartits (UiComuns.ps1)
# ----------------------------------------------------------------------------
# Es carrega el PRIMER de tots els moduls: nomes depen de WinForms i de
# $ScriptRoot / $Script:HeadlessTest (definits just aqui sobre), i en canvi el
# fan servir gairebe tots els altres (_NewForm, la banda granat, els estils de
# boto, _MakeMultiFilter, _AddConfigRow...). Tenir-los en un modul propi evita
# que un modul hagi de dependre del punt d'entrada per dibuixar una pantalla.
. (Join-Path $ScriptRoot 'UiComuns.ps1')


# ----------------------------------------------------------------------------
# Eines integrades al menu (Pas 1): planificador de rutes i revisio del mobil.
# Es llancen des del menu (botons). La revisio del mobil es una comprovacio
# d'UN SOL COP: mira si han arribat paquets del mobil (via Drive), els genera i
# surt. Ja no hi ha cap vigilant en segon pla ni polling ni interruptor.
# ----------------------------------------------------------------------------

# Revisa UNA vegada si han arribat informes del mobil i els genera. Llanca
# mobil/Vigilant.ps1 (mode d'un sol cop), espera que acabi i mostra un resum.
# Normalment SENSE finestra de consola; nomes visible el primer cop, si cal
# autoritzar Google Drive (mode API), perque l'usuari pugui completar-ho.
function Invoke-RevisarMobil {
    $vig = Join-Path $ScriptRoot (Join-Path 'mobil' 'Vigilant.ps1')
    if (-not (Test-Path -LiteralPath $vig)) {
        [System.Windows.Forms.MessageBox]::Show("No s'ha trobat mobil\Vigilant.ps1.", 'Revisar mobil', 'OK', 'Error') | Out-Null
        return
    }
    # Cal autoritzar Drive (primer cop, mode API)? Aleshores ho fem visible.
    $needsAuth = $false
    try { $needsAuth = ($DriveEntradaId -and -not (Test-DriveApiConfigured)) } catch { $needsAuth = $false }

    $resFile = Join-Path $env:TEMP ("revisarmobil_" + [guid]::NewGuid().ToString() + ".json")
    $psArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$vig`" -ResultFile `"$resFile`""
    try {
        if ($needsAuth) {
            Start-Process -FilePath 'powershell.exe' -ArgumentList $psArgs -Wait | Out-Null
        } else {
            Start-Process -FilePath 'powershell.exe' -ArgumentList $psArgs -WindowStyle Hidden -Wait | Out-Null
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("No s'ha pogut revisar el mobil:`n$($_.Exception.Message)", 'Revisar mobil', 'OK', 'Error') | Out-Null
        return
    }

    # Llegim el resum que ha deixat Vigilant.ps1.
    $ok = 0; $err = 0; $pend = 0
    try {
        if (Test-Path -LiteralPath $resFile) {
            $r = (Get-Content -LiteralPath $resFile -Raw -Encoding UTF8 | ConvertFrom-Json)
            $ok = [int]$r.ok; $err = [int]$r.err; $pend = [int]$r.pending
            Remove-Item -LiteralPath $resFile -Force -ErrorAction SilentlyContinue
        }
    } catch { }

    if ($pend -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No hi havia entrades pendents del mobil.", 'Revisar mobil', 'OK', 'Information') | Out-Null
    } else {
        $icon = if ($err -gt 0) { 'Warning' } else { 'Information' }
        [System.Windows.Forms.MessageBox]::Show(
            "Revisio completada.`n`nInformes generats: $ok`nErrors: $err",
            'Revisar mobil', 'OK', $icon) | Out-Null
    }
}

# Obre el planificador de rutes (Ruta.ps1) EN EL MATEIX PROCES, no en una
# finestra/consola a part. L'operador '&' executa el script en un AMBIT AILLAT:
# aixi les seves variables ($ScriptRoot, etc.) NO contaminen el generador, pero
# la finestra forma part del mateix programa (mateix escut a la barra de
# tasques) i, en acabar o prémer Enrere, es torna al menu. Invoke-RutaMain fa
# servir 'return' (mai 'exit'), aixi que en cancel-lar torna al menu sense
# tancar el generador.
function Start-RutaTool {
    $ruta = Join-Path $ScriptRoot (Join-Path 'rutes' 'Ruta.ps1')
    if (-not (Test-Path -LiteralPath $ruta)) {
        [System.Windows.Forms.MessageBox]::Show("No s'ha trobat Ruta.ps1.", 'Ruta', 'OK', 'Error') | Out-Null
        return
    }
    try {
        & $ruta
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al planificador de rutes:`n$($_.Exception.Message)", 'Ruta', 'OK', 'Error') | Out-Null
    }
}

# Carreguem el modul de format (Format.ps1). Conte les funcions Format-Section,
# Format-Item, etc. i $ReportFormatConfig. Reutilitzable per altres tipus
# d'informes.
. (Join-Path $ScriptRoot 'Format.ps1')

# Carreguem el modul de seguiment (Seguiment.ps1). Conte el mode "Informe de
# seguiment" (afegir anotacions de resolucio sobre un informe anterior). Es
# carrega tambe en mode headless perque les seves funcions pures es puguin
# provar des dels tests.
. (Join-Path $ScriptRoot 'Seguiment.ps1')

# Client de Google Drive per API (per al mode mobil SENSE Drive d'escriptori).
# Nomes defineix funcions; no fa res en carregar-se. Les credencials viuen a
# %LOCALAPPDATA% (fora del repo) i les posa Authorize-Drive.ps1.
. (Join-Path $ScriptRoot 'DriveApi.ps1')

# ESTRUCTURALS viu a l'arrel del clone (al costat dels .bat), no dins
# de suport/, per facilitar que l'usuari edita les plantilles.
$EstructuralsDir = Join-Path $RepoRoot 'ESTRUCTURALS'
$HeaderPath      = Join-Path $EstructuralsDir '0 CAPCALERA.docx'
$ConclusionsPath = Join-Path $EstructuralsDir '0 CONCLUSIONS.json'

# ----------------------------------------------------------------------------
# Configuracio per defecte. Es pot sobreescriure des de config.ps1 (opcional)
# al costat del .ps1 (dins de suport/).
# ----------------------------------------------------------------------------
# OutputDir per defecte: 'Informes generats' AL COSTAT DEL CLONE (no dins
# de suport/), per quedar al nivell que l'usuari veu i pot obrir amb el
# Word sense entrar a carpetes internes. El .gitignore l'exclou.
$OutputDir              = Join-Path $RepoRoot 'Informes generats'
$ActivitatsDir          = 'I:\Activitats_Ordenances\Activitats\5.- Sergi Fadurdo\2_Controls Excels'
$AlwaysConclusionsCount = 2

# Carpeta on l'eina "Copiar informes" (menu INFORMES) fa una copia de seguretat
# PLANA (tots els Word a una sola carpeta) de la carpeta d'informes. Sense valor
# per defecte: l'usuari l'ha de triar a la pantalla Configuracio. Es pot
# sobreescriure a config.ps1 o a settings.json (per ordinador).
$CopiaInformesDir       = ''

# Carpeta ARREL dels informes ja generats (per l'escaner "Actualitzar base
# d'informes" del menu). El valor per defecte es deriva de $ActivitatsDir DESPRES
# de carregar config.ps1 (mes avall), aixi respecta un $ActivitatsDir o un
# $InformesDir personalitzats al config.
$InformesDir            = $null

# Mobil (Google Drive). Carpeta sincronitzada al PC amb el Google Drive
# d'escriptori. Conte tres subcarpetes:
#   Entrada/    -> paquets JSON que arriben del mobil (els llegeix Vigilant.ps1)
#   Processats/ -> paquets ja generats (s'hi mouen despres)
#   Dades/      -> base de dades d'activitats per al mobil (PRIVADA: noms i
#                  adreces de ciutadans NO van mai al GitHub public)
# El default apunta a la ubicacio tipica del Drive d'escriptori. Es pot
# sobreescriure $DriveBaseDir a config.ps1. El guard de $env:USERPROFILE evita
# que el dot-source falli en entorns sense aquesta variable (tests a Linux).
$DriveBaseDir = if ($env:USERPROFILE) {
    Join-Path $env:USERPROFILE (Join-Path 'Google Drive' 'Informes-Cornella')
} else {
    Join-Path $RepoRoot 'Drive-Local'
}

# Mode mobil SENSE Drive d'escriptori (accés a Drive per API). IDs de les
# carpetes de Drive (es treuen de la URL de cada carpeta). Si estan buits i no
# hi ha credencials (Authorize-Drive.ps1), s'usa el mode de carpeta local de
# dalt. Es poden sobreescriure a config.ps1.
$DriveEntradaId    = ''
$DriveProcessatsId = ''
$DriveDadesId      = ''

# Mode "Informe de seguiment": frases que marquen l'inici del bloc de
# conclusions a esborrar de l'informe anterior. La deteccio es insensible a
# accents/majuscules. Es pot sobreescriure des de config.ps1.
$SeguimentConclusionPhrases = @(
    "Vist l'anterior",
    'Ho poso al seu coneixement',
    'Cornella de Llobregat,'
)

$configPath = Join-Path $ScriptRoot 'config.ps1'
if (Test-Path -LiteralPath $configPath) {
    . $configPath
}

# Carpeta arrel dels informes: si el config no l'ha fixat, la derivem de
# $ActivitatsDir (germana '...\Informes'). IMPORTANT: fem servir NOMES operacions
# de cadena (System.IO.Path), NO Split-Path/Join-Path, perque aquests resolen la
# UNITAT del cami i peten si no existeix (p.ex. la I: de la feina quan treballes
# fora). Aixi el programa arrenca igual encara que la I: no hi sigui.
if (-not $InformesDir) {
    try {
        $parentDir = [System.IO.Path]::GetDirectoryName($ActivitatsDir)
        if (-not [string]::IsNullOrWhiteSpace($parentDir)) {
            $InformesDir = [System.IO.Path]::Combine($parentDir, 'Informes')
        }
    } catch { $InformesDir = $null }
}

# Configuracio LOCAL d'aquest ordinador (Settings.ps1 + pantalla Configuracio):
# una capa MES per sobre de config.ps1, nomes d'aquest PC (%LOCALAPPDATA%, mai
# es puja a git). Primer capturem els valors "de repositori" (config.ps1 o
# hardcodejats) ABANS de sobreescriure'ls, perque la pantalla de Configuracio
# els pugui mostrar com a "valor per defecte" i el boto "Restaura" hi torni.
. (Join-Path $ScriptRoot 'Settings.ps1')

$Script:DefaultInformesDir     = $InformesDir
$Script:DefaultActivitatsDir   = $ActivitatsDir
$Script:DefaultOutputDir       = $OutputDir
$Script:DefaultDriveBaseDir    = $DriveBaseDir
$Script:DefaultCopiaInformesDir = $CopiaInformesDir
# $RutesOutputDir no el fa servir aquest script (nomes rutes/Ruta.ps1), pero
# es mostra/edita des de la mateixa pantalla de Configuracio: el calculem amb
# el mateix valor per defecte que fa servir Ruta.ps1.
$Script:DefaultRutesOutputDir = Join-Path $RepoRoot 'Rutes generades'

$Script:AppSettings = Load-AppSettings
$InformesDir      = _ResolveEffectiveValue $AppSettings.InformesDir      $InformesDir
$ActivitatsDir    = _ResolveEffectiveValue $AppSettings.ActivitatsDir    $ActivitatsDir
$OutputDir        = _ResolveEffectiveValue $AppSettings.OutputDir        $OutputDir
$DriveBaseDir     = _ResolveEffectiveValue $AppSettings.DriveBaseDir     $DriveBaseDir
$CopiaInformesDir = _ResolveEffectiveValue $AppSettings.CopiaInformesDir $CopiaInformesDir

# Carreguem el modul ACT_EXTR (ActExtr.ps1): mode "Activitats extraordinaries"
# (Decret 112/2010). Es carrega DESPRES de $RepoRoot, $EstructuralsDir i del
# config (perque pugui calcular les rutes del registre/plantilles i deixar que
# config.ps1 les sobreescrigui). Tambe en headless, per als tests de la logica.
. (Join-Path $ScriptRoot 'ActExtr.ps1')

# Carreguem el modul d'escaneig d'informes (Informes.ps1): construeix la base de
# dades JSON (ID GIA + data + conclusio) a partir de la carpeta d'informes. Es
# carrega tambe en headless perque els tests provin la logica de text pura.
. (Join-Path $ScriptRoot 'Informes.ps1')

# Lector dels ESTRUCTURALS en JSON (format estandard unic). Nomes defineix
# funcions; segur en headless. Es carrega abans que s'usi (Get-ParsedCataleg/
# Read-Conclusions/Parse-ActExtrTemplate), que llegeixen els .json.
. (Join-Path $ScriptRoot 'CatalegJson.ps1')

# Editor visual dels ESTRUCTURALS (Editar catalegs). Funcions pures (model<->JSON)
# testejables; la finestra WinForms nomes s'executa a Windows.
. (Join-Path $ScriptRoot 'EditorCatalegs.ps1')

# Generador de les VISTES en Word dels catalegs (des dels JSON). Els .docx
# d'ESTRUCTURALS ja no son plantilles: son vistes per consultar (excepte
# '0 CAPCALERA.docx'). Funcions pures testejables; Word (COM) nomes a Windows.
. (Join-Path $ScriptRoot 'VistaWord.ps1')

# Eina "Convertir informes a PDF (i signar)". Funcions pures (rutes, arguments
# d'AutoFirma) testejables; Word (COM) i AutoFirma nomes s'executen a Windows.
. (Join-Path $ScriptRoot 'PdfSignar.ps1')

# Editor dels textos del correu del mobil (docs\dades\email-textos.json).
# Funcions pures testejables; la finestra (WinForms) nomes a Windows.
. (Join-Path $ScriptRoot 'EmailTextos.ps1')

# Carreguem l'eina "Controls periodics" (llistat d'activitats amb control
# periodic a partir de l'Excel). Nomes defineix funcions; segur en headless.
. (Join-Path $ScriptRoot 'ControlsPeriodics.ps1')

# Avisos de control periodic per correu (esborranys a Outlook) des de l'eina
# Controls periodics. Funcions pures testejables; Outlook (COM)/WinForms a Windows.
. (Join-Path $ScriptRoot 'ControlsCpEmail.ps1')

# Carreguem la pantalla de Configuracio (rutes d'aquest PC + actualitzar el
# programa). Nomes defineix funcions (WinForms), segur en headless.
. (Join-Path $ScriptRoot 'Configuracio.ps1')

# Subcarpetes de Drive derivades de $DriveBaseDir (despres del config, perque
# n'hi hagi prou amb sobreescriure $DriveBaseDir a config.ps1).
$DriveEntradaDir    = Join-Path $DriveBaseDir 'Entrada'
$DriveProcessatsDir = Join-Path $DriveBaseDir 'Processats'
$DriveDadesDir      = Join-Path $DriveBaseDir 'Dades'

# Estat persistent. Es guarda a %LOCALAPPDATA% per no embrutar el repositori
# i no haver de tocar .gitignore. lastreport.json conte les dades de l'ULTIM
# informe generat amb exit, per poder-lo replicar des del Pas 2.
$AppDataDir     = Join-Path $env:LOCALAPPDATA 'InformesCornella'
$LastReportPath = Join-Path $AppDataDir 'lastreport.json'

function Ensure-AppDataDir {
    if (-not (Test-Path -LiteralPath $AppDataDir)) {
        New-Item -ItemType Directory -Path $AppDataDir -Force | Out-Null
    }
}

# ----------------------------------------------------------------------------
# Una sola instancia del programa
# ----------------------------------------------------------------------------
# Si el programa ja esta obert i es torna a llancar, en lloc d'obrir una segona
# finestra portem al davant la que ja hi ha i sortim. Fem servir un MUTEX amb
# nom (Windows destrueix el mutex automaticament quan el proces propietari mor,
# fins i tot si es tanca amb 'exit' o el mata Actualitzar.bat, aixi la deteccio
# sempre es correcta) i un PIDFILE amb el PID del proces viu (perque la segona
# instancia sapiga quina finestra enfocar i perque Actualitzar.bat pugui tancar
# el programa abans d'actualitzar).
$Script:AppMutex    = $null
$Script:PidFilePath = Join-Path $AppDataDir 'running.pid'

# Retorna $true si som la (unica) instancia i podem continuar; $false si ja n'hi
# ha una d'oberta (l'hem enfocada i el qui crida ha de sortir sense fer res).
function Enter-SingleInstance {
    Ensure-AppDataDir
    $createdNew = $false
    try {
        $Script:AppMutex = New-Object System.Threading.Mutex($true, 'Local\InformesCornellaGenerador', [ref]$createdNew)
    } catch {
        # Si el mutex falla per qualsevol motiu, no bloquegem l'arrencada.
        return $true
    }
    if ($createdNew) {
        # Som la primera instancia: desem el nostre PID i programem la neteja
        # del pidfile en sortir (el mutex el neteja sol el sistema operatiu).
        try { Set-Content -LiteralPath $Script:PidFilePath -Value ([string]$PID) -Encoding ASCII } catch { }
        try {
            $Global:CornellaPidFile = $Script:PidFilePath
            Register-EngineEvent -SourceIdentifier ([System.Management.Automation.PsEngineEvent]::Exiting) -Action {
                try { if ($Global:CornellaPidFile -and (Test-Path -LiteralPath $Global:CornellaPidFile)) { Remove-Item -LiteralPath $Global:CornellaPidFile -Force } } catch { }
            } | Out-Null
        } catch { }
        return $true
    }
    # Ja hi ha una instancia viva: enfoquem la seva finestra i sortim.
    $otherPid = $null
    try { $otherPid = [int]((Get-Content -LiteralPath $Script:PidFilePath -Raw -ErrorAction Stop).Trim()) } catch { $otherPid = $null }
    if ($otherPid) {
        try { [CornellaApp.Win]::FocusProcessWindow($otherPid) | Out-Null } catch { }
    }
    return $false
}

# ----------------------------------------------------------------------------
# Word COM helpers
# ----------------------------------------------------------------------------
function New-WordApp {
    # Crea la instancia de Word. En alguns equips (Word no instal-lat, primera
    # execucio pendent, activacio, o COM trencat) New-Object pot FALLAR o
    # retornar $null sense llançar excepcio. Ho detectem aqui i donem un
    # missatge clar, en lloc de petar 800 linies mes avall amb un
    # "metode sobre NULL" criptic.
    $w = $null
    try { $w = New-Object -ComObject Word.Application } catch { $w = $null }
    if ($null -eq $w) {
        [System.Windows.Forms.MessageBox]::Show(
            "No s'ha pogut iniciar Microsoft Word." + [Environment]::NewLine + [Environment]::NewLine +
            "Comprova que:" + [Environment]::NewLine +
            "  - Word estigui instal-lat en aquest equip." + [Environment]::NewLine +
            "  - L'hagis obert almenys un cop (per completar la primera" + [Environment]::NewLine +
            "    configuracio i l'activacio)." + [Environment]::NewLine +
            "  - No quedi cap finestra de Word bloquejada o demanant accio" + [Environment]::NewLine +
            "    (tanca-les des del Gestor de tasques si cal)." + [Environment]::NewLine + [Environment]::NewLine +
            "Aquest programa necessita Microsoft Word per generar els informes.",
            'Microsoft Word no disponible', 'OK', 'Error') | Out-Null
        exit 1
    }
    $w.Visible = $false
    $w.DisplayAlerts = 0  # wdAlertsNone
    # Evita que Word obri els fitxers de xarxa en "Vista protegida", que
    # bloqueja InsertParagraphAfter i altres operacions de modificacio.
    try { $w.AutomationSecurity = 1 } catch { }  # msoAutomationSecurityLow
    return $w
}

function Close-WordApp($word) {
    if ($null -ne $word) {
        try { $word.Quit() } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}

# ----------------------------------------------------------------------------
# Persistencia de l'ultim informe (per replicar-lo)
# ----------------------------------------------------------------------------
# Format del lastreport.json (versio 1):
#   {
#     "Version": 1,
#     "Timestamp": "<ISO 8601>",
#     "CatalegBaseName": "REQ1",
#     "Header": { "ID_GIA": "...", ... },
#     "SelectedKeys": [ "SectionTitle::ItemShort", ... ],
#     "FieldValues":  { "nom": "valor", ... },
#     "ConclusionTexts": [ "text1", ... ]
#   }

function Save-LastReport($state) {
    try {
        Ensure-AppDataDir
        $state.Version   = 1
        $state.Timestamp = (Get-Date).ToString('o')
        ($state | ConvertTo-Json -Depth 10) | Set-Content -LiteralPath $LastReportPath -Encoding UTF8
    } catch {
        # Si no podem desar, no es un error fatal. Continuem en silenci.
    }
}

function Load-LastReport {
    if (-not (Test-Path -LiteralPath $LastReportPath)) { return $null }
    try {
        $raw = Get-Content -LiteralPath $LastReportPath -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
        return ($raw | ConvertFrom-Json)
    } catch {
        return $null
    }
}

# Construeix la clau "Seccio::Item" o "Seccio::Item::Fill" usada per
# identificar de manera unica un element seleccionat al Pas 3.
function _ItemKey($sectionTitle, $itemShort, $childShort = $null) {
    if ($childShort) { return "$sectionTitle::$itemShort::$childShort" }
    return "$sectionTitle::$itemShort"
}

# ----------------------------------------------------------------------------
# Activitats Excel database - precarrega + validacio
# ----------------------------------------------------------------------------
# Mapeig de columnes Excel (1-based) per la fulla "Estes"/"Estès" del fitxer
# YYYY-MM-DD ACTIVITATS.xls. Es valida pel text de capcalera (fila 1); si no
# es troba el text esperat, es continua amb l'index per defecte pero
# s'afegeix un avis a $script:_activitatsWarnings.
$Script:ActivitatsColumns = @(
    @{ Key='ID';        Col=1;  HeaderHint='ID Activitat' }
    @{ Key='TITULAR';   Col=10; HeaderHint='Rao social' }
    @{ Key='TIPUS_VIA'; Col=48; HeaderHint='Tipus via' }
    @{ Key='CARRER';    Col=49; HeaderHint='Carrer' }
    @{ Key='NUMERO';    Col=50; HeaderHint='Numero' }
    @{ Key='LLETRA';    Col=52; HeaderHint='Lletra' }
    @{ Key='PIS';       Col=55; HeaderHint='Pis' }
    @{ Key='PORTA';     Col=56; HeaderHint='Porta' }
    @{ Key='ACTIVITAT'; Col=94; HeaderHint='Activitat principal' }
)

# Ruta de la base de dades LOCAL al clone. Si l'usuari executa el programa
# fora de la xarxa de la feina, pot copiar una "YYYY-MM-DD ACTIVITATS.xls"
# a aquesta carpeta i el programa la fara servir com a fallback. La carpeta
# es queda al clone (existeix amb un .gitkeep); els .xls/.xlsx de dins NO
# es pugen (estan al .gitignore).
$LocalActivitatsDir = Join-Path $RepoRoot 'BASE DE DADES ACTIVITATS'

# Cerca el fitxer 'YYYY-MM-DD ACTIVITATS.xls/xlsx' mes recent en una carpeta.
# Retorna $null si no se'n troba cap.
function _FindLatestActivitatsIn($dir) {
    if ([string]::IsNullOrWhiteSpace($dir) -or -not (Test-Path -LiteralPath $dir)) { return $null }
    $regex = '^(\d{4}-\d{2}-\d{2})\s+ACTIVITATS\.(xls|xlsx)$'
    $candidates = Get-ChildItem -LiteralPath $dir -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match $regex } |
        ForEach-Object {
            if ($_.Name -match $regex) {
                [pscustomobject]@{
                    File = $_
                    Date = [datetime]::ParseExact($Matches[1], 'yyyy-MM-dd', $null)
                }
            }
        } | Sort-Object Date -Descending
    if ($candidates.Count -eq 0) { return $null }
    return $candidates[0]
}

# Cerca la base de dades en dues ubicacions, en ordre:
#   1. $ActivitatsDir  (xarxa de la feina)
#   2. $LocalActivitatsDir  (carpeta local del clone, fallback per a fora feina)
# Retorna un PSCustomObject amb:
#   File   : System.IO.FileInfo
#   Date   : data parsejada del nom
#   Source : 'primary' (xarxa) o 'fallback' (local del clone)
# Si no se'n troba a cap, retorna $null.
function Find-LatestActivitatsExcel {
    $r = _FindLatestActivitatsIn $ActivitatsDir
    if ($null -ne $r) {
        Add-Member -InputObject $r -NotePropertyName Source -NotePropertyValue 'primary' -Force
        return $r
    }
    $r = _FindLatestActivitatsIn $LocalActivitatsDir
    if ($null -ne $r) {
        Add-Member -InputObject $r -NotePropertyName Source -NotePropertyValue 'fallback' -Force
        return $r
    }
    return $null
}

# Normalitza un text Unicode (sense diacritics, minuscules) per a comparacio.
function _NormalizeText($s) {
    if ($null -eq $s) { return '' }
    $t = ([string]$s).Normalize([System.Text.NormalizationForm]::FormD)
    return (($t -replace '\p{Mn}','').ToLower().Trim())
}

# Cerca l'index (1-based) de la columna a la fila de capcalera (fila 1) el text
# de la qual conte TOTS els termes de $mustContain i CAP dels de $mustNotContain
# (comparacio normalitzada: minuscules, sense accents). Retorna 0 si no es troba.
function _FindColIndex($data, $cols, [string[]]$mustContain, [string[]]$mustNotContain) {
    for ($c = 1; $c -le $cols; $c++) {
        $h = _NormalizeText $data[1, $c]
        if ([string]::IsNullOrWhiteSpace($h)) { continue }
        $ok = $true
        foreach ($m in $mustContain) {
            if (-not $h.Contains((_NormalizeText $m))) { $ok = $false; break }
        }
        if ($ok -and $null -ne $mustNotContain) {
            foreach ($m in $mustNotContain) {
                if ($h.Contains((_NormalizeText $m))) { $ok = $false; break }
            }
        }
        if ($ok) { return $c }
    }
    return 0
}

# Converteix un valor de cel·la a text. Els enters d'Excel arriben com a double;
# els mostrem sense decimals ni notacio cientifica.
function _CellToString($v) {
    if ($null -eq $v) { return '' }
    if ($v -is [double]) {
        if ([math]::Floor($v) -eq $v) { return [string][int64]$v }
        return [string]$v
    }
    return ([string]$v).Trim()
}

# Formata una cel·la de data a "dd/MM/yyyy" descartant l'hora. A l'Excel les
# dates arriben com a double (numero de serie OLE); tambe acceptem text.
function _FormatDateOnly($v) {
    if ($null -eq $v) { return '' }
    if ($v -is [double]) {
        try { return ([DateTime]::FromOADate($v)).ToString('dd/MM/yyyy') } catch { return '' }
    }
    $s = ([string]$v).Trim()
    if ([string]::IsNullOrWhiteSpace($s)) { return '' }
    $dt = [datetime]::MinValue
    if ([datetime]::TryParse($s, [ref]$dt)) { return $dt.ToString('dd/MM/yyyy') }
    return $s.Split(' ')[0]  # si no es pot parsejar, agafem la part abans de l'hora
}

# Localitza la fulla "Estes"/"Estès" del workbook acceptant variants Unicode.
function _FindEstesSheet($wb) {
    $sheetNames = @()
    foreach ($s in $wb.Sheets) {
        $sheetNames += $s.Name
        if ((_NormalizeText $s.Name) -eq 'estes') { return @{ Sheet=$s; Names=$sheetNames } }
    }
    return @{ Sheet=$null; Names=$sheetNames }
}

# Valida la fila de capcalera comparant els textos esperats. Retorna una
# llista (potser buida) d'avisos en text per mostrar a l'usuari.
function _ValidateActivitatsHeaders($data, $rows, $cols) {
    $warnings = New-Object System.Collections.ArrayList
    if ($rows -lt 1) { return $warnings }
    foreach ($col in $Script:ActivitatsColumns) {
        $idx  = $col.Col
        if ($idx -lt 1 -or $idx -gt $cols) {
            [void]$warnings.Add("Columna $idx fora de rang per a '$($col.Key)' (Excel te $cols columnes).")
            continue
        }
        $cell = $data[1, $idx]
        $cellN = _NormalizeText $cell
        $hintN = _NormalizeText $col.HeaderHint
        if ([string]::IsNullOrWhiteSpace($cellN) -or -not $cellN.Contains($hintN.Split(' ')[0])) {
            [void]$warnings.Add("La columna $idx esperava '$($col.HeaderHint)' pero te '$cell'.")
        }
    }
    return $warnings
}

# Precarrega TOTES les activitats de l'Excel en una hashtable indexada per ID.
# Es crida una sola vegada al comencar el Pas 2; despres les cerques son
# instantanies (no calen mes obertures d'Excel encara que l'usuari premi
# "Cercar" diverses vegades).
#
# Retorna un PSCustomObject amb:
#   File        : System.IO.FileInfo del fitxer Excel
#   Date        : data del fitxer (parsejada del nom)
#   ById        : hashtable [string ID] -> hashtable amb TITULAR, ADRECA,
#                 ACTIVITAT, EXP_NUM, NUM_ANOTACIO, DATA_ANOTACIO
#   Warnings    : llista de cadenes amb avisos de validacio de columnes
function Initialize-ActivitatsCache($excelFile) {
    # Igual que amb Word: si Excel no esta disponible, New-Object pot fallar o
    # retornar $null. Donem un missatge clar (aquest error es propaga a
    # Get-HeaderData, que ja mostra "Error llegint l'Excel").
    $excel = $null
    try { $excel = New-Object -ComObject Excel.Application } catch { $excel = $null }
    if ($null -eq $excel) {
        throw "No s'ha pogut iniciar Microsoft Excel. Comprova que estigui instal-lat i obert almenys un cop."
    }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
        $wb = $excel.Workbooks.Open($excelFile.FullName, 0, $true)  # ReadOnly
        try {
            $found = _FindEstesSheet $wb
            $sh = $found.Sheet
            if ($null -eq $sh) {
                throw "No s'ha trobat la fulla 'Estes'/'Estès' al fitxer Excel. Fulles disponibles: $($found.Names -join ', ')"
            }
            $used = $sh.UsedRange
            $data = $used.Value2
            if ($null -eq $data) {
                return [pscustomobject]@{ ById = @{}; Warnings = @("L'Excel sembla buit.") }
            }
            $rows = $data.GetLength(0)
            $cols = $data.GetLength(1)

            # Avisos de validacio. Construim un ArrayList REAL aqui (no podem
            # assignar directament el retorn de _ValidateActivitatsHeaders: en
            # retornar un ArrayList, PowerShell el desempaqueta i, si es buit,
            # $warnings quedaria $null i $warnings.Add()/.ToArray() petarien).
            $warnings = New-Object System.Collections.ArrayList
            foreach ($w in (_ValidateActivitatsHeaders $data $rows $cols)) {
                if ($null -ne $w) { [void]$warnings.Add($w) }
            }

            # Columnes localitzades pel TEXT de la capcalera (mes robust que un
            # index fix). Si l'Excel canvia l'ordre de columnes, segueix
            # funcionant mentre el nom es mantingui.
            $colExp  = _FindColIndex $data $cols @('expedient') $null
            $colNum  = _FindColIndex $data $cols @('registre','entrada') @('data')
            $colData = _FindColIndex $data $cols @('data','registre','entrada') $null
            if ($colExp  -eq 0) { [void]$warnings.Add("No s'ha trobat la columna 'Num. expedient'.") }
            if ($colNum  -eq 0) { [void]$warnings.Add("No s'ha trobat la columna 'Num. registre entrada'.") }
            if ($colData -eq 0) { [void]$warnings.Add("No s'ha trobat la columna 'Data registre entrada'.") }

            # Index per ID (columna 1).
            $byId = @{}
            $get = {
                param($r, $c)
                if ($c -lt 1 -or $c -gt $cols) { return '' }
                $v = $data[$r, $c]
                if ($null -eq $v) { return '' }
                return ([string]$v).Trim()
            }
            for ($r = 2; $r -le $rows; $r++) {
                $cell = $data[$r, 1]
                if ($null -eq $cell) { continue }
                $id = if ($cell -is [double]) {
                    if ([math]::Floor($cell) -eq $cell) { [string][int]$cell } else { [string]$cell }
                } else { [string]$cell }
                if ([string]::IsNullOrWhiteSpace($id)) { continue }

                $tipusVia = & $get $r 48
                $carrer   = & $get $r 49
                $numero   = & $get $r 50
                $lletra   = & $get $r 52
                $pis      = & $get $r 55
                $porta    = & $get $r 56
                $rao      = & $get $r 10
                $raoMobil = & $get $r 23
                $raoEmail = & $get $r 25
                $actPrin  = & $get $r 94
                $parts = @($tipusVia, $carrer, $numero, $lletra, $pis, $porta) |
                    Where-Object { $_ -and $_.Trim() -ne '' }
                # Construim l'accent amb el codepoint Unicode explicit (U+00C0,
                # 'A' amb accent greu) per evitar que la lletra accentuada del
                # literal es corrompi segons l'encoding amb que PowerShell 5.1
                # llegeix aquest fitxer (sortia "CORNELLÃ€").
                $ciutat = "CORNELL$([char]0x00C0) DE LLOBREGAT"
                $adreca = ($parts -join ' ') + ", $ciutat"

                $expNum = if ($colExp  -gt 0) { _CellToString  $data[$r, $colExp] }  else { '' }
                $numAno = if ($colNum  -gt 0) { _CellToString  $data[$r, $colNum] }  else { '' }
                $datAno = if ($colData -gt 0) { _FormatDateOnly $data[$r, $colData] } else { '' }

                $byId[$id] = @{
                    TITULAR       = $rao
                    MOBIL         = $raoMobil
                    EMAIL         = $raoEmail
                    ADRECA        = $adreca
                    ACTIVITAT     = $actPrin
                    EXP_NUM       = $expNum
                    NUM_ANOTACIO  = $numAno
                    DATA_ANOTACIO = $datAno
                }
            }
            return [pscustomobject]@{ ById = $byId; Warnings = $warnings.ToArray() }
        } finally {
            $wb.Close($false)
        }
    } finally {
        try { $excel.Quit() } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

# Cerca una activitat per ID al cache precarregat. Retorna $null si no es troba.
function Get-ActivitatFromCache($cache, $idGia) {
    if ($null -eq $cache -or $null -eq $cache.ById) { return $null }
    $key = [string]$idGia
    if ($cache.ById.ContainsKey($key)) { return $cache.ById[$key] }
    return $null
}

# Treu la data (yyyy-MM-dd) d'un objecte activitats.json ja existent: prefereix
# el camp SourceDate i, si no hi es (versions antigues), la treu del nom Source.
# Retorna [datetime]::MinValue si no en pot deduir cap.
function _ParseActivitatsDate($obj) {
    if ($null -eq $obj) { return [datetime]::MinValue }
    $d = [datetime]::MinValue
    if ($obj.SourceDate -and [datetime]::TryParseExact([string]$obj.SourceDate, 'yyyy-MM-dd', $null, [System.Globalization.DateTimeStyles]::None, [ref]$d)) {
        return $d
    }
    if ($obj.Source -and ([string]$obj.Source) -match '(\d{4}-\d{2}-\d{2})') {
        if ([datetime]::TryParseExact($matches[1], 'yyyy-MM-dd', $null, [System.Globalization.DateTimeStyles]::None, [ref]$d)) { return $d }
    }
    return [datetime]::MinValue
}

# Decideix si cal exportar la base d'activitats al Drive. Nomes cal si la base
# LOCAL es MES NOVA que la que ja hi ha al Drive. Si tenen la mateixa data (o
# la del Drive es mes nova), NO cal: ens estalviem obrir l'Excel i pujar un
# fitxer identic. Funcio pura (provable).
#   $localDate    : data del fitxer Excel local (del nom YYYY-MM-DD).
#   $existingDate : data de l'activitats.json que ja hi ha al Drive.
function Test-ShouldExportActivitats([datetime]$localDate, [datetime]$existingDate) {
    # Si no tenim una data local fiable, exportem (no podem decidir res).
    if ($localDate -le [datetime]::MinValue) { return $true }
    return ($localDate -gt $existingDate)
}

# Llegeix la data (SourceDate) de l'activitats.json que ja hi ha al Drive,
# SENSE obrir l'Excel. Funciona tant en mode API com en mode carpeta local
# sincronitzada. Retorna [datetime]::MinValue si no n'hi ha o no es pot llegir
# (es fail-safe: qualsevol error -> MinValue, que fa que s'exporti).
function Get-DriveActivitatsDate {
    try {
        if (Test-DriveApiConfigured) {
            if (-not $DriveDadesId) { return [datetime]::MinValue }
            $existingId = Find-DriveFileId 'activitats.json' $DriveDadesId
            if (-not $existingId) { return [datetime]::MinValue }
            $existing = (Get-DriveFileText $existingId) | ConvertFrom-Json
            return _ParseActivitatsDate $existing
        }
        $outFile = Join-Path $DriveDadesDir 'activitats.json'
        if (Test-Path -LiteralPath $outFile) {
            $existing = (Get-Content -LiteralPath $outFile -Raw -Encoding UTF8) | ConvertFrom-Json
            return _ParseActivitatsDate $existing
        }
        return [datetime]::MinValue
    } catch {
        return [datetime]::MinValue
    }
}

# Exporta la base de dades d'activitats (nomes els camps de capcalera, per ID
# GIA) a un JSON dins la carpeta PRIVADA de Drive, perque el mobil pugui
# auto-emplenar la capcalera. Aquestes dades son personals i NO van mai al
# GitHub public: nomes a Drive (compte privat de l'usuari). Es fail-safe:
# qualsevol error es registra i es retorna $false sense interrompre el flux.
#
# IMPORTANT: NO sobreescriu la base del Drive si la que ja hi ha es MES NOVA
# (pujada des d'un altre PC). Compara per data del nom del fitxer Excel.
function Export-ActivitatsToDrive($cache, $latest) {
    try {
        if ($null -eq $cache -or $null -eq $cache.ById) { return $false }
        $localDate = if ($latest -and $latest.Date) { $latest.Date } else { [datetime]::MinValue }

        # No tornar a exportar si el Drive ja te una base amb la MATEIXA data
        # (o mes nova, pujada des d'un altre PC). Aixi no es sobreescriu una
        # versio mes nova ni es perd temps pujant una d'identica. La data del
        # Drive es llegeix del propi activitats.json (camp SourceDate).
        $existingDate = Get-DriveActivitatsDate
        if (-not (Test-ShouldExportActivitats $localDate $existingDate)) {
            Write-Host ("  El Drive ja esta al dia ({0}); no cal tornar a exportar (local {1})." -f $existingDate.ToString('yyyy-MM-dd'), $localDate.ToString('yyyy-MM-dd'))
            return $true
        }

        $payload = [ordered]@{
            GeneratedAt = (Get-Date).ToString('o')
            Source      = if ($latest) { $latest.File.Name } else { '' }
            SourceDate  = $localDate.ToString('yyyy-MM-dd')
            Count       = $cache.ById.Count
            ById        = $cache.ById
        }
        $json = ($payload | ConvertTo-Json -Depth 6)

        # Mode API (sense Drive d'escriptori): pugem activitats.json directament
        # a la carpeta Dades de Drive. Si no hi ha credencials, caiem al mode de
        # carpeta local sincronitzada.
        if (Test-DriveApiConfigured) {
            if (-not $DriveDadesId) {
                Write-Host "Avis: hi ha credencials de Drive pero falta \$DriveDadesId a config.ps1. No s'exporten activitats."
                return $false
            }
            Save-DriveJson 'activitats.json' $DriveDadesId $json | Out-Null
            return $true
        }

        if (-not (Test-Path -LiteralPath $DriveDadesDir)) {
            New-Item -ItemType Directory -Path $DriveDadesDir -Force | Out-Null
        }
        $outFile = Join-Path $DriveDadesDir 'activitats.json'
        $json | Set-Content -LiteralPath $outFile -Encoding UTF8
        return $true
    } catch {
        Write-Host "Avis: no s'ha pogut exportar les activitats a Drive ($($_.Exception.Message))."
        return $false
    }
}

# ----------------------------------------------------------------------------
# Step 1 - Cataleg picker
# ----------------------------------------------------------------------------
function Get-Catalegs {
    # Catalegs = fitxers .JSON d'ESTRUCTURALS que NO comencin amb "0 " (fixos:
    # capcalera, conclusions) i que NO siguin plantilles del mode ACT_EXTR
    # (ACT_EXTR_REQ / ACT_EXTR_FAV), que no son catalegs del wizard normal sino
    # que les gestiona el mode "Activitats extraordinaries".
    #
    # ATENCIO: la font de veritat es el .JSON, no el .docx. Abans aixo llistava
    # '*.docx' i, quan l'usuari va apartar els .docx (que ja no servien per
    # generar), van DESAPAREIXER del menu "Requeriment - Nou" i "Ampliacio de
    # termini". Els .docx d'ESTRUCTURALS son ara nomes VISTES generades des dels
    # JSON (vegeu VistaWord.ps1); l'unic .docx que encara es una plantilla de
    # veritat es '0 CAPCALERA.docx'.
    Get-ChildItem -LiteralPath $EstructuralsDir -Filter '*.json' |
        Where-Object {
            $_.Name -notlike '0 *' -and $_.Name -notlike '0_*' -and
            $_.Name -notlike 'ACT_EXTR*' -and -not $_.Name.StartsWith('~$')
        } |
        Sort-Object Name
}

# NOTA: la tria de cataleg ja no es un pas a part. Ara la pantalla inicial
# (Select-Mode, a Seguiment.ps1) fusiona la tria de MODE i la de CATALEG en un
# sol menu (Pas 1), i passa el cataleg triat directament al wizard
# (Invoke-NouWizard). Get-Catalegs (a dalt) segueix sent la font de catalegs.

# ----------------------------------------------------------------------------
# Step 2 - Header data (formulari + precarrega Excel)
# ----------------------------------------------------------------------------
# Construeix el formulari de capcalera (controls + botons), retorna la
# tupla amb el formulari, el diccionari de controls i el boto Cercar perque
# Get-HeaderData hi puga lligar la logica de cerca i validacio.
function _BuildHeaderForm($excelInfo) {
    # $excelInfo: hashtable amb claus Text (string) i Source ('primary'|'fallback').
    # Quan es 'fallback', el rotul es taronja i porta el text "[FALLBACK LOCAL]"
    # davant perque l'usuari sapiga que NO esta usant la base de dades oficial
    # de la xarxa.
    $form = _NewForm
    $form.Text = 'Pas 2 - Dades de la capcalera'
    $form.Size = New-Object System.Drawing.Size(720, 620)
    $form.StartPosition = 'CenterScreen'

    # Espai reservat a dalt per la banda granat (44) + barra de passos (34):
    # tots els controls absoluts es desplacen cap avall aquest offset.
    $topOffset = 74

    $lblBd = New-Object System.Windows.Forms.Label
    $lblBd.Location = New-Object System.Drawing.Point(15, (12 + $topOffset))
    $lblBd.Size = New-Object System.Drawing.Size(680, 22)
    if ($excelInfo.Source -eq 'fallback') {
        $lblBd.Text = "[FALLBACK LOCAL]  " + $excelInfo.Text
        $lblBd.ForeColor = [System.Drawing.Color]::DarkOrange
        $lblBd.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    } else {
        $lblBd.Text = $excelInfo.Text
        $lblBd.ForeColor = [System.Drawing.Color]::DarkBlue
    }
    $form.Controls.Add($lblBd)

    $controls  = @{}
    $rowLabels = @{}
    $addRow = {
        param($label, $y, $tbWidth, $key)
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $label
        $lbl.Location = New-Object System.Drawing.Point(15, $y)
        $lbl.Size = New-Object System.Drawing.Size(200, 22)
        [void]$form.Controls.Add($lbl)

        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Location = New-Object System.Drawing.Point(220, ($y - 2))
        $tb.Size = New-Object System.Drawing.Size($tbWidth, 22)
        [void]$form.Controls.Add($tb)
        $controls[$key]  = $tb
        $rowLabels[$key] = $lbl
    }

    $y = 50 + $topOffset
    & $addRow 'ID GIA' $y 380 'ID_GIA'
    $btnSearch = New-Object System.Windows.Forms.Button
    $btnSearch.Text = 'Cercar'
    $btnSearch.Location = New-Object System.Drawing.Point(605, ($y - 4))
    $btnSearch.Size = New-Object System.Drawing.Size(80, 28)
    _StylePrimaryButton $btnSearch
    [void]$form.Controls.Add($btnSearch)
    $y += 38

    & $addRow "Num. d'expedient (autom., editable)"   $y 460 'EXP_NUM';   $y += 38
    & $addRow 'Titular (autom., editable)'            $y 460 'TITULAR';   $y += 38
    & $addRow 'Adreca (autom., editable)'             $y 460 'ADRECA';    $y += 38
    & $addRow 'Activitat (autom., editable)'          $y 460 'ACTIVITAT'; $y += 38

    # --- Origen de l'informe: documentacio aportada o visita d'inspeccio ---
    # Segons la tria es mostren uns camps o uns altres i canvia la linia
    # "Objecte:" de la capcalera (placeholder <<ORIGEN>>, muntat a
    # Apply-HeaderReplacements).
    $lblOrigen = New-Object System.Windows.Forms.Label
    $lblOrigen.Text = "Origen de l'informe:"
    $lblOrigen.Location = New-Object System.Drawing.Point(15, $y)
    $lblOrigen.Size = New-Object System.Drawing.Size(200, 22)
    [void]$form.Controls.Add($lblOrigen)

    $rbDoc = New-Object System.Windows.Forms.RadioButton
    $rbDoc.Text = 'Documentacio aportada'
    $rbDoc.Location = New-Object System.Drawing.Point(220, ($y - 2))
    $rbDoc.AutoSize = $true
    $rbDoc.Checked = $true
    [void]$form.Controls.Add($rbDoc)

    $rbInsp = New-Object System.Windows.Forms.RadioButton
    $rbInsp.Text = "Visita d'inspeccio"
    $rbInsp.Location = New-Object System.Drawing.Point(440, ($y - 2))
    $rbInsp.AutoSize = $true
    [void]$form.Controls.Add($rbInsp)
    $controls['_ORIGEN_DOC']  = $rbDoc
    $controls['_ORIGEN_INSP'] = $rbInsp
    $y += 34

    # Camps de documentacio aportada (NUM/DATA anotacio) i, superposat al primer,
    # el camp de data d'inspeccio: nomes es veu el joc que toca segons la tria.
    $yAnot = $y
    & $addRow "Num. d'anotacio (autom., editable)" $yAnot 460 'NUM_ANOTACIO'
    & $addRow "Data d'inspeccio"                    $yAnot 460 'DATA_INSPECCIO'
    $y = $yAnot + 38
    & $addRow "Data d'anotacio (autom., editable)"  $y 460 'DATA_ANOTACIO'; $y += 50

    # Mostra/amaga els camps segons l'origen triat.
    $applyOrigen = {
        $isDoc = $rbDoc.Checked
        foreach ($k in 'NUM_ANOTACIO', 'DATA_ANOTACIO') {
            $controls[$k].Visible  = $isDoc
            $rowLabels[$k].Visible = $isDoc
        }
        $controls['DATA_INSPECCIO'].Visible  = -not $isDoc
        $rowLabels['DATA_INSPECCIO'].Visible = -not $isDoc
    }.GetNewClosure()
    $rbDoc.add_CheckedChanged($applyOrigen)
    & $applyOrigen

    $back = New-Object System.Windows.Forms.Button
    $back.Text = 'Enrere'
    $back.Location = New-Object System.Drawing.Point(15, $y)
    $back.Size = New-Object System.Drawing.Size(90, 30)
    $back.DialogResult = 'Retry'
    _StyleSecondaryButton $back
    [void]$form.Controls.Add($back)

    $recover = New-Object System.Windows.Forms.Button
    $recover.Text = ([char]0x21BA + " Recuperar dades ultim informe")
    $recover.Location = New-Object System.Drawing.Point(115, $y)
    $recover.Size = New-Object System.Drawing.Size(250, 30)
    _StyleSecondaryButton $recover
    [void]$form.Controls.Add($recover)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = ('Seg' + [char]0x00FC + 'ent ' + [char]0x2192)   # Següent →
    $ok.Location = New-Object System.Drawing.Point(575, $y)
    $ok.Size = New-Object System.Drawing.Size(110, 30)
    _StylePrimaryButton $ok
    $form.AcceptButton = $ok
    [void]$form.Controls.Add($ok)

    # Barra de passos (pas 2 actiu) + banda granat. La banda s'afegeix DESPRES
    # per quedar a dalt de tot; les dues son Dock=Top.
    [void](_AddStepBar $form 2)
    [void](_AddBrandHeader $form "Dades de l'activitat" $null 44)

    return @{ Form=$form; Controls=$controls; BtnSearch=$btnSearch; BtnOk=$ok; BtnBack=$back; BtnRecover=$recover }
}

# Llegeix els valors dels controls i retorna un hashtable amb la capcalera.
function _ReadHeaderControls($controls) {
    $tipus = if ($controls['_ORIGEN_INSP'] -and $controls['_ORIGEN_INSP'].Checked) { 'insp' } else { 'doc' }
    @{
        ID_GIA         = $controls['ID_GIA'].Text.Trim()
        EXP_NUM        = $controls['EXP_NUM'].Text.Trim()
        TITULAR        = $controls['TITULAR'].Text.Trim()
        ADRECA         = $controls['ADRECA'].Text.Trim()
        ACTIVITAT      = $controls['ACTIVITAT'].Text.Trim()
        ORIGEN_TIPUS   = $tipus
        NUM_ANOTACIO   = $controls['NUM_ANOTACIO'].Text.Trim()
        DATA_ANOTACIO  = $controls['DATA_ANOTACIO'].Text.Trim()
        DATA_INSPECCIO = $controls['DATA_INSPECCIO'].Text.Trim()
    }
}

# Precarrega valors d'una capcalera anterior als controls del formulari.
# $preload pot ser un hashtable (navegacio en memoria) o un PSCustomObject
# (dades de l'ultim informe llegides de JSON).
function _PreloadHeaderControls($controls, $preload) {
    if ($null -eq $preload) { return }
    # Helper per llegir una clau tant si $preload es hashtable com PSCustomObject.
    $getVal = {
        param($k)
        if ($preload -is [System.Collections.IDictionary]) {
            if ($preload.Contains($k)) { return $preload[$k] }
        } elseif ($preload.PSObject.Properties.Name -contains $k) {
            return $preload.$k
        }
        return $null
    }
    foreach ($k in 'ID_GIA','EXP_NUM','TITULAR','ADRECA','ACTIVITAT','NUM_ANOTACIO','DATA_ANOTACIO','DATA_INSPECCIO') {
        $v = & $getVal $k
        if ($null -ne $v) { $controls[$k].Text = [string]$v }
    }
    # Restaura l'origen triat (documentacio aportada / visita d'inspeccio).
    $tipus = & $getVal 'ORIGEN_TIPUS'
    if ($controls['_ORIGEN_DOC'] -and $controls['_ORIGEN_INSP']) {
        if ([string]$tipus -eq 'insp') { $controls['_ORIGEN_INSP'].Checked = $true }
        else                           { $controls['_ORIGEN_DOC'].Checked  = $true }
    }
}

function Get-HeaderData {
    param($preload = $null)

    $latest = Find-LatestActivitatsExcel
    if ($null -eq $latest) {
        [System.Windows.Forms.MessageBox]::Show(
            "No s'ha trobat cap fitxer 'YYYY-MM-DD ACTIVITATS.xls/xlsx' a cap de les ubicacions:`n`n" +
            "  1. $ActivitatsDir`n" +
            "  2. $LocalActivitatsDir  (fallback local)`n`n" +
            "Si estas fora de la feina, copia una base de dades a la carpeta`n" +
            "'BASE DE DADES ACTIVITATS' dins de la carpeta del programa.",
            'Base de dades no trobada', 'OK', 'Error') | Out-Null
        exit 1
    }
    # RENDIMENT: obrir l'Excel i llegir tota la base es la part mes lenta del
    # Pas 2. Ho fem UNA sola vegada per sessio (i per fitxer): si es torna al
    # Pas 2 (p.ex. Enrere des del Pas 3), reaprofitem la cache ja carregada i
    # NO reobrim l'Excel ni refem l'export a Drive. Aixi el pas a pas es immediat.
    $cacheKey = '{0}|{1}|{2}' -f $latest.File.FullName, $latest.File.Length, $latest.File.LastWriteTimeUtc.Ticks
    if ($script:_sessionActKey -eq $cacheKey -and $null -ne $script:_sessionActCache) {
        $actCache = $script:_sessionActCache
    } else {
        if ($latest.Source -eq 'fallback') {
            # Avis explicit (un cop) perque l'usuari sapiga que treballa amb una
            # copia local i no amb la xarxa de la feina.
            [System.Windows.Forms.MessageBox]::Show(
                "No s'ha trobat la base de dades a la xarxa.`n`n" +
                "S'usa la copia LOCAL del clone:`n  $($latest.File.FullName)`n`n" +
                "Comprova que sigui prou recent.",
                'Base de dades: copia local (fallback)', 'OK', 'Warning') | Out-Null
        }
        # Precarrega TOTA la base de dades a memoria. A partir d'aqui les cerques
        # son immediates (no cal reobrir Excel).
        try {
            $actCache = Initialize-ActivitatsCache -excelFile $latest.File
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error llegint l'Excel:`n$($_.Exception.Message)",'Error','OK','Error') | Out-Null
            exit 1
        }
        if ($actCache.Warnings -and $actCache.Warnings.Count -gt 0) {
            $msg = "Avisos de validacio de l'Excel (l'auto-fill podria fallar):`n`n" + ($actCache.Warnings -join "`n")
            [System.Windows.Forms.MessageBox]::Show($msg,'Avisos','OK','Warning') | Out-Null
        }
        $script:_sessionActCache = $actCache
        $script:_sessionActKey   = $cacheKey
    }

    # NOTA: NO refresquem la copia d'activitats al Drive aqui. Comprovar-ho i
    # pujar-ho feia que generar un informe (Pas 2) trigues molt (accedia a la
    # xarxa/Drive cada cop). Aquesta sincronitzacio per al mobil ara nomes es fa
    # amb Actualitzar.bat (ExportaDades.ps1 -> Export-ActivitatsCmd), aixi
    # generar informes es rapid i no depen de la xarxa.

    $labelText = "Base de dades d'activitats: $($latest.File.Name)  (data: $($latest.Date.ToString('yyyy-MM-dd'))) - $($actCache.ById.Count) activitats carregades"
    $f = _BuildHeaderForm @{ Text = $labelText; Source = $latest.Source }
    $form       = $f.Form
    $controls   = $f.Controls
    $btnSearch  = $f.BtnSearch
    $ok         = $f.BtnOk
    $btnRecover = $f.BtnRecover

    _PreloadHeaderControls $controls $preload

    # Boto "Recuperar dades ultim informe": carrega les dades de l'ultim
    # informe generat amb exit i les deixa als formularis perque l'usuari
    # les revisi/modifiqui pas per pas. La resta de passos (seleccio, camps,
    # conclusions) es precarreguen via $script:_recoveredReport, que Main llegeix.
    $script:_recoveredReport = $null
    $btnRecover.add_Click({
        $rep = Load-LastReport
        if ($null -eq $rep) {
            [System.Windows.Forms.MessageBox]::Show("No hi ha cap informe anterior desat.",'Sense dades','OK','Information') | Out-Null
            return
        }
        if ($rep.Header) { _PreloadHeaderControls $controls $rep.Header }
        $script:_recoveredReport = $rep
        [System.Windows.Forms.MessageBox]::Show(
            "Dades de l'ultim informe carregades.`n`nRevisa-les i modifica el que calgui (per exemple, canvia l'ID GIA i prem 'Cercar' per a una activitat nova). En continuar, els passos seguents tambe sortiran precarregats.",
            'Recuperat', 'OK', 'Information') | Out-Null
    })

    # Cerca per ID GIA: instantania des del cache. Omple els camps automatics.
    $doSearch = {
        $idGia = $controls['ID_GIA'].Text.Trim()
        if ([string]::IsNullOrWhiteSpace($idGia)) {
            [System.Windows.Forms.MessageBox]::Show("Has d'introduir un ID GIA.",'Falta ID GIA','OK','Warning') | Out-Null
            return $false
        }
        $act = Get-ActivitatFromCache $actCache $idGia
        if ($null -eq $act) {
            [System.Windows.Forms.MessageBox]::Show(
                "L'ID GIA '$idGia' no s'ha trobat a la base de dades`n($($latest.File.Name)).",
                'Activitat no trobada', 'OK', 'Error') | Out-Null
            foreach ($k in 'TITULAR','ADRECA','ACTIVITAT','EXP_NUM','NUM_ANOTACIO','DATA_ANOTACIO') {
                $controls[$k].Text = ''
            }
            return $false
        }
        foreach ($k in 'TITULAR','ADRECA','ACTIVITAT','EXP_NUM','NUM_ANOTACIO','DATA_ANOTACIO') {
            if ($act.ContainsKey($k)) { $controls[$k].Text = [string]$act[$k] }
        }
        return $true
    }

    $btnSearch.add_Click({ [void](& $doSearch) })

    $script:_headerData = $null
    $ok.add_Click({
        $idGia = $controls['ID_GIA'].Text.Trim()
        if ([string]::IsNullOrWhiteSpace($idGia)) {
            [System.Windows.Forms.MessageBox]::Show("Has d'introduir un ID GIA.",'Falten dades','OK','Warning') | Out-Null
            return
        }
        if ([string]::IsNullOrWhiteSpace($controls['TITULAR'].Text) -or
            [string]::IsNullOrWhiteSpace($controls['ADRECA'].Text) -or
            [string]::IsNullOrWhiteSpace($controls['ACTIVITAT'].Text)) {
            if (-not (& $doSearch)) { return }
        }
        $script:_headerData = _ReadHeaderControls $controls
        $form.DialogResult = 'OK'
        $form.Close()
    })

    $res = $form.ShowDialog()
    if ($res -eq 'Retry') { return [pscustomobject]@{ Nav='back' } }
    if ($res -ne 'OK')    { exit 0 }
    return [pscustomobject]@{ Nav='next'; Data=$script:_headerData; Recovered=$script:_recoveredReport }
}

# ----------------------------------------------------------------------------
# Step 3 - Lectura del cataleg
# ----------------------------------------------------------------------------

# Cache en memoria del cataleg parsejat durant l'execucio del programa.
# Clau = ruta del JSON + data de modificacio + mida: si es genera un segon
# informe del mateix cataleg en la mateixa sessio no cal tornar a llegir-lo, i
# si l'usuari l'edita (l'editor de catalegs reescriu el .json) la clau canvia i
# es torna a llegir sol.
$Script:_parseCache = @{}

# Llegeix un cataleg. La FONT DE VERITAT es el .json d'ESTRUCTURALS.
#
# Abans hi havia un "respatller segur" que, si el JSON no hi era o no es podia
# llegir, obria el .docx del mateix nom amb el Word i el parsejava pels estils
# (Titol 1 / Titol 2). Es va treure, i no per estalviar codi: els .docx ja NO
# son catalegs, son VISTES generades en format d'informe (VistaWord.ps1). El
# respatller, doncs, no hauria fallat: hauria llegit la vista i hauria generat
# un informe silenciosament EQUIVOCAT. Val mes petar aqui, amb un missatge clar.
function Get-ParsedCataleg($path) {
    $jsonPath = if ([System.IO.Path]::GetExtension($path) -ieq '.json') { $path }
                else { [System.IO.Path]::ChangeExtension($path, '.json') }
    if (-not (Test-Path -LiteralPath $jsonPath)) {
        throw "No s'ha trobat el cataleg: $jsonPath"
    }
    $fi = Get-Item -LiteralPath $jsonPath -ErrorAction Stop
    $key = "json|$jsonPath|$($fi.LastWriteTimeUtc.Ticks)|$($fi.Length)"
    if ($Script:_parseCache.ContainsKey($key)) { return $Script:_parseCache[$key] }
    $parsed = Read-CatalegJson $jsonPath
    $Script:_parseCache[$key] = $parsed
    return $parsed
}

# ----------------------------------------------------------------------------
# Step 3 (UI) - TreeView amb filtre + checkboxes
# ----------------------------------------------------------------------------
# Helpers per al filtre del TreeView. Es manté l'estructura $sections a part i
# es reconstrueix el tree quan canvia el filtre, preservant els check states.

function _TextMatches($text, $needle) {
    if ([string]::IsNullOrEmpty($needle)) { return $true }
    if ($null -eq $text) { return $false }
    return $text.ToLower().Contains($needle.ToLower())
}

# Reconstrueix el TreeView segons el text de filtre. Preserva check states
# (passats en una hashtable [key] -> bool) i els actualitza durant la construccio.
# Bloquegem la propagacio automatica de check durant el rebuild perque
# marcar nodes programmaticament dispara l'event AfterCheck.
# Construeix el text del tooltip d'un item (o fill) del TreeView del Pas 3:
# concatena les BodyLines descartant els enllacos (URL-only i URLs incrustats)
# perque l'usuari pugui veure el text de l'item en passar el ratolí. Aplica
# una mica de neteja per fer-lo llegible i el limita a 600 caracters per
# evitar tooltips desmesurats.
function _GetItemTooltip($el) {
    if ($null -eq $el -or $null -eq $el.BodyLines) { return '' }
    $parts = New-Object System.Collections.ArrayList
    foreach ($ln in $el.BodyLines) {
        $s = [string]$ln
        if ([string]::IsNullOrWhiteSpace($s)) { continue }
        # Linies marcades per Cita: tota la linia es URL.
        if ($s.StartsWith('[[URL]] ')) { continue }
        # Linies que son nomes URL.
        if ($s.Trim() -match '^https?://') { continue }
        # Linia mixta: extreu nomes el text (descarta URLs incrustats).
        $p = _SplitTextAndUrls $s
        if (-not [string]::IsNullOrWhiteSpace($p.Text)) { [void]$parts.Add($p.Text) }
    }
    $tip = ($parts -join [Environment]::NewLine)
    if ($tip.Length -gt 600) { $tip = $tip.Substring(0, 600) + '...' }
    return $tip
}

function _RebuildTree($tv, $sections, $needle, $checkStates) {
    $tv.BeginUpdate()
    $script:_propagating = $true
    try {
        $tv.Nodes.Clear()
        foreach ($sec in $sections) {
            $secMatches = _TextMatches $sec.Title $needle

            # Recollim items/children que cal mostrar
            $itemNodesToAdd = New-Object System.Collections.ArrayList
            $currentContainer = $null
            foreach ($el in $sec.Items) {
                if ($el.Kind -eq 'subsection') {
                    $subShow = $secMatches -or (_TextMatches $el.Short $needle)
                    $itemNodesToAdd.Add(@{ Kind='Subsection'; El=$el; ChildShows=@(); ShowMe=$subShow }) | Out-Null
                    continue
                }
                if ($el.Kind -eq 'intro') { continue }  # mai al TreeView

                $itemMatches = _TextMatches $el.Short $needle
                $matchedChildren = New-Object System.Collections.ArrayList
                foreach ($ch in $el.Children) {
                    if ($secMatches -or $itemMatches -or (_TextMatches $ch.Short $needle)) {
                        [void]$matchedChildren.Add($ch)
                    }
                }
                $showItem = $secMatches -or $itemMatches -or ($matchedChildren.Count -gt 0)
                if ($showItem) {
                    $itemNodesToAdd.Add(@{ Kind='Item'; El=$el; ChildShows=$matchedChildren; ShowMe=$true }) | Out-Null
                }
            }

            # Si el filtre no es buit i no hi ha cap item/subsection visible,
            # ometem la seccio del tot (tret que el titol de la seccio matchi).
            $anyChild = $false
            foreach ($n in $itemNodesToAdd) { if ($n.ShowMe) { $anyChild = $true; break } }
            if (-not $secMatches -and -not $anyChild) { continue }

            $secNode = New-Object System.Windows.Forms.TreeNode($sec.Title)
            $secNode.Tag = @{ Kind = 'Section'; Ref = $sec; Key = $sec.Title }
            $secNode.NodeFont = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
            $stKey = "SECT::$($sec.Title)"
            if ($checkStates.ContainsKey($stKey)) { $secNode.Checked = $checkStates[$stKey] }
            [void]$tv.Nodes.Add($secNode)

            $container = $secNode
            foreach ($n in $itemNodesToAdd) {
                if (-not $n.ShowMe) { continue }
                if ($n.Kind -eq 'Subsection') {
                    $subNode = New-Object System.Windows.Forms.TreeNode($n.El.Short)
                    $subNode.Tag = @{ Kind = 'Subsection'; Ref = $n.El }
                    $subNode.NodeFont = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Underline)
                    [void]$secNode.Nodes.Add($subNode)
                    $container = $subNode
                    continue
                }
                # Item. NodeFont explicit (regular) perque el font base del
                # TreeView es negreta (vegeu nota a Select-Items) i no volem
                # que els items surtin en negreta ni que es retallin.
                $itNode = New-Object System.Windows.Forms.TreeNode($n.El.Short)
                $itNode.Tag = @{ Kind = 'Item'; Ref = $n.El; SectionTitle = $sec.Title }
                $itNode.NodeFont = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Regular)
                $itNode.ToolTipText = _GetItemTooltip $n.El
                $itKey = (_ItemKey $sec.Title $n.El.Short)
                if ($checkStates.ContainsKey($itKey)) { $itNode.Checked = $checkStates[$itKey] }
                [void]$container.Nodes.Add($itNode)
                foreach ($ch in $n.ChildShows) {
                    $chNode = New-Object System.Windows.Forms.TreeNode($ch.Short)
                    $chNode.Tag = @{ Kind = 'Child'; Ref = $ch; SectionTitle = $sec.Title; ParentShort = $n.El.Short }
                    $chNode.NodeFont = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Regular)
                    $chNode.ToolTipText = _GetItemTooltip $ch
                    $chKey = (_ItemKey $sec.Title $n.El.Short $ch.Short)
                    if ($checkStates.ContainsKey($chKey)) { $chNode.Checked = $checkStates[$chKey] }
                    [void]$itNode.Nodes.Add($chNode)
                }
            }
            $secNode.ExpandAll()
        }
    } finally {
        $script:_propagating = $false
        $tv.EndUpdate()
    }
}

# Recorre el TreeView i llegeix tots els check states en una hashtable.
function _CollectCheckStates($tv, $checkStates) {
    foreach ($secNode in $tv.Nodes) {
        $secTitle = $secNode.Tag.Ref.Title
        $checkStates["SECT::$secTitle"] = [bool]$secNode.Checked
        foreach ($node in $secNode.Nodes) {
            $kind = $node.Tag.Kind
            if ($kind -eq 'Subsection') {
                foreach ($itNode in $node.Nodes) {
                    $itShort = $itNode.Tag.Ref.Short
                    $checkStates[(_ItemKey $secTitle $itShort)] = [bool]$itNode.Checked
                    foreach ($chNode in $itNode.Nodes) {
                        $checkStates[(_ItemKey $secTitle $itShort $chNode.Tag.Ref.Short)] = [bool]$chNode.Checked
                    }
                }
            } elseif ($kind -eq 'Item') {
                $itShort = $node.Tag.Ref.Short
                $checkStates[(_ItemKey $secTitle $itShort)] = [bool]$node.Checked
                foreach ($chNode in $node.Nodes) {
                    $checkStates[(_ItemKey $secTitle $itShort $chNode.Tag.Ref.Short)] = [bool]$chNode.Checked
                }
            }
        }
    }
}

function Select-Items {
    param($sections, $preloadSelectedKeys = $null, $fields = $null, $preloadValues = $null)
    if ($null -eq $fields) { $fields = [ordered]@{} }

    $form = _NewForm
    $form.Text = 'Pas 3 - Seleccio de deficiencies'
    $form.Size = New-Object System.Drawing.Size(1180, 818)
    $form.StartPosition = 'CenterScreen'

    # Espai reservat a dalt per la banda de marca + barra de passos.
    $topOffset = 78

    # Filtre (textbox al capdamunt). Cada vegada que canvia, es reconstrueix
    # el TreeView amb nomes les coincidencies. Els check states es preserven.
    $lblFilter = New-Object System.Windows.Forms.Label
    $lblFilter.Text = 'Filtre:'
    $lblFilter.Location = New-Object System.Drawing.Point(10, (14 + $topOffset))
    $lblFilter.AutoSize = $true
    $form.Controls.Add($lblFilter)

    $tbFilter = New-Object System.Windows.Forms.TextBox
    $tbFilter.Location = New-Object System.Drawing.Point(60, (10 + $topOffset))
    $tbFilter.Size = New-Object System.Drawing.Size(400, 22)
    $tbFilter.Anchor = 'Top, Left'
    $form.Controls.Add($tbFilter)

    $btnClear = New-Object System.Windows.Forms.Button
    $btnClear.Text = 'Esborra'
    $btnClear.Location = New-Object System.Drawing.Point(465, (9 + $topOffset))
    $btnClear.Size = New-Object System.Drawing.Size(70, 24)
    $btnClear.add_Click({ $tbFilter.Text = '' })
    $form.Controls.Add($btnClear)

    $tv = New-Object System.Windows.Forms.TreeView
    $tv.Location = New-Object System.Drawing.Point(10, (40 + $topOffset))
    $tv.Size = New-Object System.Drawing.Size(560, 610)
    $tv.CheckBoxes = $true
    $tv.HideSelection = $false
    $tv.ShowNodeToolTips = $true
    $tv.Anchor = 'Top, Bottom, Left'
    # El font base es la negreta mes ampla que faran servir les seccions. Aixo
    # evita el bug de WinForms en que un node amb NodeFont mes ample que el
    # font del control surt retallat. Items i fills posen NodeFont regular.
    $tv.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($tv)

    # Panell de detall (dreta): mostra el TEXT de les deficiencies marcades amb
    # els seus desplegables/camps inline per omplir-los aqui mateix.
    $lblDetail = New-Object System.Windows.Forms.Label
    $lblDetail.Text = 'Text i opcions de les deficiencies marcades:'
    $lblDetail.Location = New-Object System.Drawing.Point(585, (14 + $topOffset))
    $lblDetail.AutoSize = $true
    $lblDetail.Anchor = 'Top, Left, Right'
    $form.Controls.Add($lblDetail)

    $detailHost = New-Object System.Windows.Forms.Panel
    $detailHost.Location = New-Object System.Drawing.Point(585, (40 + $topOffset))
    $detailHost.Size = New-Object System.Drawing.Size(575, 610)
    $detailHost.AutoScroll = $true
    $detailHost.BorderStyle = 'FixedSingle'
    $detailHost.Anchor = 'Top, Bottom, Left, Right'
    $form.Controls.Add($detailHost)

    $detailV = New-Object System.Windows.Forms.FlowLayoutPanel
    $detailV.FlowDirection = 'TopDown'
    $detailV.WrapContents = $false
    $detailV.AutoSize = $true
    $detailV.AutoSizeMode = 'GrowAndShrink'
    $detailV.Location = New-Object System.Drawing.Point(0, 0)
    $detailV.Width = 550
    $detailHost.Controls.Add($detailV)

    # Estats de check persistents entre rebuilds. Inicialitzat des de session.
    $checkStates = @{}
    if ($preloadSelectedKeys) {
        foreach ($k in $preloadSelectedKeys) {
            if ([string]::IsNullOrWhiteSpace($k)) { continue }
            $checkStates[[string]$k] = $true
        }
    }

    # Amplada util del panell de detall (sense la barra de desplacament).
    $detailInnerWidth = { [Math]::Max(280, $detailHost.ClientSize.Width - 26) }

    # Crea un FlowLayoutPanel que EMBOLCALLA el text (wrap) a l'ample disponible.
    # Clau: amb AutoSize + WrapContents cal limitar l'amplada amb MaximumSize,
    # si no, el panell creix cap a la dreta i no salta de linia.
    $newWrapFlow = {
        param($leftMargin)
        $f = New-Object System.Windows.Forms.FlowLayoutPanel
        $f.FlowDirection = 'LeftToRight'; $f.WrapContents = $true
        $f.AutoSize = $true; $f.AutoSizeMode = 'GrowAndShrink'
        $f.Margin = New-Object System.Windows.Forms.Padding($leftMargin, 0, 2, 4)
        $iw = & $detailInnerWidth
        $f.MaximumSize = New-Object System.Drawing.Size([Math]::Max(120, ($iw - $leftMargin - 6)), 0)
        return $f
    }

    # Reajusta l'amplada maxima de tots els blocs quan la finestra canvia de mida
    # (sense reconstruir res, per no perdre el focus mentre s'escriu).
    $applyDetailWidths = {
        $iw = & $detailInnerWidth
        $detailV.MaximumSize = New-Object System.Drawing.Size($iw, 0)
        foreach ($child in $detailV.Controls) {
            if ($child -is [System.Windows.Forms.FlowLayoutPanel]) {
                $child.MaximumSize = New-Object System.Drawing.Size([Math]::Max(120, ($iw - $child.Margin.Left - 6)), 0)
            }
        }
    }

    # Reconstrueix el panell de detall a partir de les caselles marcades.
    $refreshDetail = {
        _CollectCheckStates $tv $checkStates
        $detailV.SuspendLayout()
        $detailV.Controls.Clear()
        $registry = _NewFieldRegistry
        $detailV.MaximumSize = New-Object System.Drawing.Size((& $detailInnerWidth), 0)
        $anyShown = $false
        foreach ($sec in $sections) {
            foreach ($el in $sec.Items) {
                if ($el.Kind -ne 'item') { continue }
                $itKey = (_ItemKey $sec.Title $el.Short)
                $itSel = $checkStates.ContainsKey($itKey) -and $checkStates[$itKey]
                $selChildren = New-Object System.Collections.ArrayList
                foreach ($ch in $el.Children) {
                    $chKey = (_ItemKey $sec.Title $el.Short $ch.Short)
                    if ($checkStates.ContainsKey($chKey) -and $checkStates[$chKey]) { [void]$selChildren.Add($ch) }
                }
                if (-not $itSel -and $selChildren.Count -eq 0) { continue }
                $anyShown = $true

                $hdr = New-Object System.Windows.Forms.Label
                $hdr.AutoSize = $true
                $hdr.Text = $el.Short
                $hdr.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
                $hdr.Margin = New-Object System.Windows.Forms.Padding(2, 8, 2, 2)
                [void]$detailV.Controls.Add($hdr)

                if ($itSel) {
                    $txt = _RichTextOfBodyLines $el.BodyLines
                    if ($txt) {
                        $flow = & $newWrapFlow 8
                        _RenderRichInto $flow $txt $fields $preloadValues $registry
                        [void]$detailV.Controls.Add($flow)
                    }
                }
                foreach ($ch in $selChildren) {
                    $ctxt = _RichTextOfBodyLines $ch.BodyLines
                    if (-not $ctxt) { continue }
                    $cflow = & $newWrapFlow 24
                    _RenderRichInto $cflow $ctxt $fields $preloadValues $registry
                    [void]$detailV.Controls.Add($cflow)
                }
            }
        }
        if (-not $anyShown) {
            $empty = New-Object System.Windows.Forms.Label
            $empty.AutoSize = $true
            $empty.ForeColor = [System.Drawing.Color]::DimGray
            $empty.Text = 'Marca deficiencies a l''esquerra per veure-les i omplir-ne les opcions.'
            $empty.Margin = New-Object System.Windows.Forms.Padding(6, 10, 2, 2)
            [void]$detailV.Controls.Add($empty)
        }
        $detailV.ResumeLayout()
    }

    # Build inicial sense filtre
    _RebuildTree $tv $sections '' $checkStates

    # Propagacio recursiva: marcar un node marca tots els descendents.
    $script:_propagating = $false
    $propagate = {
        param($node)
        foreach ($c in $node.Nodes) {
            $c.Checked = $node.Checked
            & $propagate $c
        }
    }
    $tv.add_AfterCheck({
        param($sender, $e)
        if ($script:_propagating) { return }
        $script:_propagating = $true
        try { & $propagate $e.Node } finally { $script:_propagating = $false }
        & $refreshDetail
    })

    # Refilter en temps real (debouncing simple: rebuild a cada keystroke;
    # amb 131 items va fluid)
    $tbFilter.add_TextChanged({
        # Guardem l'estat actual ABANS de reconstruir
        _CollectCheckStates $tv $checkStates
        _RebuildTree $tv $sections $tbFilter.Text $checkStates
    })

    # Quan la finestra canvia de mida, reajustem l'amplada dels blocs perque el
    # text es reembolcalli a l'ample nou (sense reconstruir, per no perdre focus).
    $detailHost.add_SizeChanged({ & $applyDetailWidths })

    # Detall inicial (mostra els ja marcats per precarrega).
    & $refreshDetail

    $back = New-Object System.Windows.Forms.Button
    $back.Text = ([char]0x2190 + ' Enrere')
    $back.Location = New-Object System.Drawing.Point(10, 740)
    $back.Size = New-Object System.Drawing.Size(100, 30)
    $back.DialogResult = 'Retry'
    $back.Anchor = 'Bottom, Left'
    _StyleSecondaryButton $back
    $form.Controls.Add($back)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = ('Seg' + [char]0x00FC + 'ent ' + [char]0x2192)
    $ok.Location = New-Object System.Drawing.Point(1050, 740)
    $ok.Size = New-Object System.Drawing.Size(110, 30)
    $ok.DialogResult = 'OK'
    $ok.Anchor = 'Bottom, Right'
    _StylePrimaryButton $ok
    $form.AcceptButton = $ok
    $form.Controls.Add($ok)

    [void](_AddStepBar $form 3)
    [void](_AddBrandHeader $form ("Defici" + [char]0x00E8 + "ncies de l'activitat") $null 44)

    $res = $form.ShowDialog()
    if ($res -eq 'Retry') { return [pscustomobject]@{ Nav='back' } }
    if ($res -ne 'OK')    { exit 0 }

    # Recollim l'estat final i el barregem amb el que tenim memoritzat per
    # items que ara mateix no es mostren (perque hi hagi filtre actiu).
    _CollectCheckStates $tv $checkStates

    # Construim el resultat en ordre del data, preservant subseccions/intros.
    $result = New-Object System.Collections.ArrayList
    foreach ($sec in $sections) {
        $chosen = New-Object System.Collections.ArrayList
        foreach ($el in $sec.Items) {
            if ($el.Kind -in 'subsection','intro') {
                [void]$chosen.Add([pscustomobject]@{
                    Kind      = $el.Kind
                    Short     = $el.Short
                    BodyLines = $el.BodyLines
                    Children  = @()
                    Selected  = $false
                })
                continue
            }
            $itKey = (_ItemKey $sec.Title $el.Short)
            $isSel = $checkStates.ContainsKey($itKey) -and $checkStates[$itKey]
            $chosenChildren = New-Object System.Collections.ArrayList
            foreach ($ch in $el.Children) {
                $chKey = (_ItemKey $sec.Title $el.Short $ch.Short)
                if ($checkStates.ContainsKey($chKey) -and $checkStates[$chKey]) {
                    [void]$chosenChildren.Add($ch)
                }
            }
            if ($isSel -or $chosenChildren.Count -gt 0) {
                [void]$chosen.Add([pscustomobject]@{
                    Kind      = 'item'
                    Short     = $el.Short
                    BodyLines = $el.BodyLines
                    Children  = $chosenChildren
                    Selected  = [bool]$isSel
                })
            }
        }
        $hasRealItem = $false
        foreach ($x in $chosen) { if ($x.Kind -eq 'item') { $hasRealItem = $true; break } }
        if ($hasRealItem) {
            [void]$result.Add([pscustomobject]@{
                Title = $sec.Title
                Items = $chosen
            })
        }
    }
    if ($result.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No s''ha seleccionat cap deficiencia.','Avis','OK','Warning') | Out-Null
        return [pscustomobject]@{ Nav='stay' }   # es torna a mostrar el Pas 3
    }
    return [pscustomobject]@{ Nav='next'; Data=$result }
}

# Extreu les claus "Seccio::Item[::Fill]" del resultat de Select-Items, per
# desar-les a la sessio.
function Get-SelectedKeysFromResult($selectedSections) {
    $keys = New-Object System.Collections.ArrayList
    foreach ($sec in $selectedSections) {
        foreach ($it in $sec.Items) {
            if ($it.Kind -ne 'item') { continue }
            if ($it.Selected) { [void]$keys.Add((_ItemKey $sec.Title $it.Short)) }
            foreach ($ch in $it.Children) {
                [void]$keys.Add((_ItemKey $sec.Title $it.Short $ch.Short))
            }
        }
    }
    return $keys.ToArray()
}

# ----------------------------------------------------------------------------
# Step 4 - Field placeholders
#   [CAMP: nom]                  -> camp de text lliure
#   [CAMP: nom (hint)]           -> camp de text amb ajuda
#   [OPCIO: nom | A | B | C]     -> desplegable; l'usuari tria A, B o C i el
#                                   text triat substitueix el placeholder
# ----------------------------------------------------------------------------
$Script:CampRegex  = [regex]'\[CAMP:\s*([^\]]+?)\s*\]'
$Script:OpcioRegex = [regex]'\[OPCIO:\s*([^\]]+?)\s*\]'

# Analitza el contingut d'un [OPCIO: ...]: "nom | A | B" -> nom + opcions.
function _ParseOpcio($raw) {
    $segs = $raw -split '\|'
    $name = $segs[0].Trim()
    $opts = @()
    for ($i = 1; $i -lt $segs.Count; $i++) {
        $o = $segs[$i].Trim()
        if ($o -ne '') { $opts += $o }
    }
    return @{ Name = $name; Options = $opts }
}

# Detecta [CAMP: ...] i [OPCIO: ...] dins $allText i els afegeix a $fields
# (sense duplicar). Modifica $fields in-place.
function _AddFieldsFromText($fields, $allText) {
    foreach ($m in $Script:CampRegex.Matches($allText)) {
        $raw = $m.Groups[1].Value.Trim()
        $name = $raw
        $hint = ''
        $parenIdx = $raw.IndexOf('(')
        if ($parenIdx -ge 0) {
            $name = $raw.Substring(0, $parenIdx).Trim()
            $hint = $raw.Substring($parenIdx).Trim().TrimStart('(').TrimEnd(')')
        }
        if (-not $fields.Contains($name)) {
            $fields[$name] = [pscustomobject]@{ Name=$name; Type='text'; Hint=$hint; Options=@(); Value='' }
        }
    }
    foreach ($m in $Script:OpcioRegex.Matches($allText)) {
        $parsed = _ParseOpcio $m.Groups[1].Value
        $name = $parsed.Name
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        if (-not $fields.Contains($name)) {
            $val = if ($parsed.Options.Count -gt 0) { [string]$parsed.Options[0] } else { '' }
            $fields[$name] = [pscustomobject]@{ Name=$name; Type='choice'; Hint=''; Options=$parsed.Options; Value=$val }
        }
    }
}

function Get-FieldsFromSelection($selectedSections) {
    $fields = [ordered]@{}
    foreach ($sec in $selectedSections) {
        foreach ($it in $sec.Items) {
            $allText = ($it.BodyLines -join ' ')
            foreach ($ch in $it.Children) {
                $allText += ' ' + ($ch.BodyLines -join ' ')
            }
            _AddFieldsFromText $fields $allText
        }
    }
    return $fields
}

# Afegeix els camps detectats a les conclusions TRIADES i les SEMPRE al
# diccionari $fields existent. Aixi al Pas 4 surten alhora els del REQ1
# i els del CONCLUSIONS.
function Add-FieldsFromConclusions($fields, $selectedConcl, $alwaysConcl) {
    foreach ($c in $selectedConcl) {
        $body = if ($c -is [string]) { $c } else { [string]$c.Body }
        _AddFieldsFromText $fields $body
    }
    foreach ($a in $alwaysConcl) {
        _AddFieldsFromText $fields ([string]$a)
    }
}

function Apply-Fields($text, $fields) {
    # Primer els desplegables [OPCIO: nom | ...] i despres els [CAMP: ...].
    $out = $Script:OpcioRegex.Replace($text, {
        param($m)
        $name = (_ParseOpcio $m.Groups[1].Value).Name
        if ($fields.Contains($name)) { return [string]$fields[$name].Value }
        return ''
    })
    $out = $Script:CampRegex.Replace($out, {
        param($m)
        $raw = $m.Groups[1].Value.Trim()
        $name = $raw
        $parenIdx = $raw.IndexOf('(')
        if ($parenIdx -ge 0) { $name = $raw.Substring(0, $parenIdx).Trim() }
        if ($fields.Contains($name)) { return [string]$fields[$name].Value }
        return ''
    })
    return $out
}

# Extreu els valors dels camps en un hashtable simple per a la sessio.
function Get-FieldValuesForSession($fields) {
    $h = @{}
    foreach ($name in $fields.Keys) { $h[$name] = $fields[$name].Value }
    return $h
}

# ----------------------------------------------------------------------------
# Renderitzat "ric" amb camps inline (Pas 3 i Pas de conclusions)
# ----------------------------------------------------------------------------
# Treu els marcadors **negreta** i //cursiva// d'un text per mostrar-lo net a
# la pantalla (al .docx final SI s'apliquen via Type-RichText). Es "loose":
# elimina TOTS els ** i //, encara que un parell quedi partit per un
# [OPCIO:]/[CAMP:] (p.ex. "**ampliar el termini [OPCIO]**"). Aixo nomes afecta
# la PREVISUALITZACIO; el text original (amb marcadors) es el que es desa i
# s'emet al document.
function _StripMarkers([string]$t) {
    if ([string]::IsNullOrEmpty($t)) { return '' }
    $t = $t -replace '\*\*', ''
    $t = $t -replace '//', ''
    return $t
}

# Segmenta un text amb [OPCIO:]/[CAMP:] en trossos ORDENATS, per renderitzar-lo
# amb controls inline alla on toca. Retorna una llista de hashtables:
#   @{ Kind='text';  Text='...' }                  (marcadors ** // ja retirats)
#   @{ Kind='opcio'; Name='...'; Options=@(...) }
#   @{ Kind='camp';  Name='...'; Hint='...' }
# Es una funcio PURA (provable sense Word/WinForms).
function _SegmentRichText([string]$text) {
    $segments = New-Object System.Collections.ArrayList
    if ([string]::IsNullOrEmpty($text)) { return $segments.ToArray() }
    # CAMP i OPCIO en una sola passada, en ordre d'aparicio.
    $rx = [regex]'\[OPCIO:\s*([^\]]+?)\s*\]|\[CAMP:\s*([^\]]+?)\s*\]'
    $pos = 0
    foreach ($m in $rx.Matches($text)) {
        if ($m.Index -gt $pos) {
            $plain = _StripMarkers $text.Substring($pos, $m.Index - $pos)
            if ($plain.Length -gt 0) { [void]$segments.Add(@{ Kind='text'; Text=$plain }) }
        }
        if ($m.Groups[1].Success) {
            $p = _ParseOpcio $m.Groups[1].Value
            [void]$segments.Add(@{ Kind='opcio'; Name=$p.Name; Options=$p.Options })
        } else {
            $raw = $m.Groups[2].Value.Trim(); $name = $raw; $hint = ''
            $pi = $raw.IndexOf('(')
            if ($pi -ge 0) { $name = $raw.Substring(0, $pi).Trim(); $hint = $raw.Substring($pi).Trim().TrimStart('(').TrimEnd(')') }
            [void]$segments.Add(@{ Kind='camp'; Name=$name; Hint=$hint })
        }
        $pos = $m.Index + $m.Length
    }
    if ($pos -lt $text.Length) {
        $plain = _StripMarkers $text.Substring($pos)
        if ($plain.Length -gt 0) { [void]$segments.Add(@{ Kind='text'; Text=$plain }) }
    }
    return $segments.ToArray()
}

# Llegeix un valor precarregat (sessio anterior) per nom de camp. $preload pot
# ser un hashtable o un PSCustomObject. Retorna $null si no hi es.
function _GetPreloadValue($preload, $name) {
    if ($null -eq $preload) { return $null }
    if ($preload -is [System.Collections.IDictionary]) {
        if ($preload.Contains($name)) { return $preload[$name] }
        return $null
    }
    if ($preload.PSObject.Properties.Name -contains $name) { return $preload.$name }
    return $null
}

# El "registre" relaciona nom de camp -> llista de controls (per sincronitzar
# els duplicats: un mateix nom pot sortir a diversos llocs de la mateixa
# pantalla i tots han de mostrar el mateix valor).
function _NewFieldRegistry { return @{} }
function _RegisterFieldControl($registry, $name, $ctrl) {
    if (-not $registry.ContainsKey($name)) { $registry[$name] = New-Object System.Collections.ArrayList }
    [void]$registry[$name].Add($ctrl)
}

# Renderitza $text dins d'un FlowLayoutPanel ($flow) barrejant etiquetes de
# text (paraula a paraula, per poder fer salt de linia) i controls inline per
# als [OPCIO:]/[CAMP:]. Crea/actualitza les entrades a $fields (mateixa forma
# que _AddFieldsFromText) i registra els controls a $registry per sincronitzar.
function _RenderRichInto($flow, [string]$text, $fields, $preload, $registry) {
    $segs = _SegmentRichText $text
    foreach ($seg in $segs) {
        if ($seg.Kind -eq 'text') {
            foreach ($w in ($seg.Text -split '\s+')) {
                if ($w -eq '') { continue }
                $lbl = New-Object System.Windows.Forms.Label
                $lbl.AutoSize = $true
                $lbl.Text = $w
                $lbl.Margin = New-Object System.Windows.Forms.Padding(0, 5, 4, 0)
                [void]$flow.Controls.Add($lbl)
            }
            continue
        }

        $name = $seg.Name
        if ([string]::IsNullOrWhiteSpace($name)) { continue }

        if ($seg.Kind -eq 'opcio') {
            if (-not $fields.Contains($name)) {
                $val = if ($seg.Options.Count -gt 0) { [string]$seg.Options[0] } else { '' }
                $pv = _GetPreloadValue $preload $name
                if ($null -ne $pv) { $val = [string]$pv }
                $fields[$name] = [pscustomobject]@{ Name=$name; Type='choice'; Hint=''; Options=$seg.Options; Value=$val }
            }
            $cb = New-Object System.Windows.Forms.ComboBox
            $cb.DropDownStyle = 'DropDownList'
            $cb.Margin = New-Object System.Windows.Forms.Padding(0, 1, 4, 0)
            foreach ($o in $seg.Options) { [void]$cb.Items.Add($o) }
            # Amplada segons l'opcio mes llarga (limitada), per llegir-la be.
            $maxLen = 0; foreach ($o in $seg.Options) { if ($o.Length -gt $maxLen) { $maxLen = $o.Length } }
            $cb.Width = [Math]::Min(520, [Math]::Max(90, ($maxLen * 7) + 30))
            $idx = $cb.Items.IndexOf([string]$fields[$name].Value)
            if ($idx -lt 0 -and $cb.Items.Count -gt 0) { $idx = 0 }
            if ($idx -ge 0) { $cb.SelectedIndex = $idx }
            $cb.Tag = $name
            _RegisterFieldControl $registry $name $cb
            $cb.add_SelectedIndexChanged({
                $v = if ($null -ne $cb.SelectedItem) { [string]$cb.SelectedItem } else { '' }
                $fields[$name].Value = $v
                foreach ($other in $registry[$name]) {
                    if ($other -ne $cb -and ($other -is [System.Windows.Forms.ComboBox]) -and ([string]$other.SelectedItem -ne $v)) {
                        $other.SelectedItem = $v
                    }
                }
            }.GetNewClosure())
            [void]$flow.Controls.Add($cb)
        } else {
            if (-not $fields.Contains($name)) {
                $val = ''
                $pv = _GetPreloadValue $preload $name
                if ($null -ne $pv) { $val = [string]$pv }
                $fields[$name] = [pscustomobject]@{ Name=$name; Type='text'; Hint=$seg.Hint; Options=@(); Value=$val }
            }
            $tb = New-Object System.Windows.Forms.TextBox
            $tb.Margin = New-Object System.Windows.Forms.Padding(0, 1, 4, 0)
            $tb.Width = 150
            $tb.Text = [string]$fields[$name].Value
            if ($seg.Hint) { $tb.AccessibleDescription = $seg.Hint }
            $tb.Tag = $name
            _RegisterFieldControl $registry $name $tb
            $tb.add_TextChanged({
                $fields[$name].Value = $tb.Text
                foreach ($other in $registry[$name]) {
                    if ($other -ne $tb -and ($other -is [System.Windows.Forms.TextBox]) -and ($other.Text -ne $tb.Text)) {
                        $other.Text = $tb.Text
                    }
                }
            }.GetNewClosure())
            [void]$flow.Controls.Add($tb)
            if ($seg.Hint) {
                $hl = New-Object System.Windows.Forms.Label
                $hl.AutoSize = $true
                $hl.Text = "($($seg.Hint))"
                $hl.ForeColor = [System.Drawing.Color]::DimGray
                $hl.Margin = New-Object System.Windows.Forms.Padding(0, 5, 4, 0)
                [void]$flow.Controls.Add($hl)
            }
        }
    }
}

# Construeix el text "ric" d'un element (item/fill): nomes la part de TEXT de
# cada BodyLine (descartem els URLs, que aqui no s'editen), unit amb espais.
function _RichTextOfBodyLines($bodyLines) {
    $parts = New-Object System.Collections.ArrayList
    foreach ($ln in $bodyLines) {
        $sp = _SplitTextAndUrls $ln
        if (-not [string]::IsNullOrWhiteSpace($sp.Text)) { [void]$parts.Add($sp.Text) }
    }
    return ($parts -join ' ')
}

# Separa el text d'una linia dels URLs que pugui contenir. Retorna:
#   @{ Text = '<tot el que hi ha abans del primer URL>'; Urls = @(url1, url2...) }
#
# Hi ha dues fonts d'URLs reconegudes:
#   1. Prefix intern '[[URL]] ': l'ha posat el lector del cataleg quan el paragraf
#      del .docx te estil 'Cita' (manera explicita, recomanada al cataleg
#      modern). En aquest cas tota la linia es l'URL.
#   2. Deteccio per contingut: qualsevol token que comenci per 'http://' o
#      'https://' (retrocompatible amb cataleg vell).
function _SplitTextAndUrls($line) {
    if ([string]::IsNullOrWhiteSpace($line)) { return @{ Text=''; Urls=@() } }
    # Cas 1: paragraf marcat com a enllac pel lector del cataleg.
    if ($line.StartsWith('[[URL]] ')) {
        $url = $line.Substring('[[URL]] '.Length).Trim()
        return @{ Text=''; Urls=@($url) }
    }
    # Cas 2: deteccio per contingut.
    $m = [regex]::Match($line, 'https?://')
    if (-not $m.Success) { return @{ Text = $line.Trim(); Urls=@() } }
    $text = $line.Substring(0, $m.Index).Trim()
    $rest = $line.Substring($m.Index)
    $urls = @()
    foreach ($tok in ($rest -split '\s+')) {
        if ($tok -match '^https?://') { $urls += $tok }
    }
    return @{ Text = $text; Urls = $urls }
}

# ----------------------------------------------------------------------------
# Step 5 - Conclusions
# ----------------------------------------------------------------------------
# Llegeix '0 CONCLUSIONS.json' i retorna un PSCustomObject amb:
#
#   HeaderText : titol del bloc de conclusions (sol ser 'CONCLUSIONS').
#   Selectable : conclusions triables al Pas 5. Cada element: { Title; Body }.
#   Always     : frases fixes (tipus 'sempre'), que hi van sempre al final.
#
# Les conclusions depenen del TIPUS D'INFORME: el fitxer s'organitza en seccions
# (una per tipus) i $reportType (el BaseName del cataleg: 'REQ1', 'TERMINI'...)
# tria quina secció va a Selectable. Buit -> totes (ho fa servir l'export del
# mobil).
#
# Nomes JSON: igual que Get-ParsedCataleg, aqui hi havia un respatller que obria
# '0 CONCLUSIONS.docx' amb el Word. Aquell .docx ja no es una font, es una VISTA
# generada, o sigui que el respatller hauria donat conclusions equivocades sense
# dir res. Es va treure.
function Read-Conclusions($path, $reportType = $null) {
    $jsonPath = if ([System.IO.Path]::GetExtension($path) -ieq '.json') { $path }
                else { [System.IO.Path]::ChangeExtension($path, '.json') }
    if (-not (Test-Path -LiteralPath $jsonPath)) {
        return [pscustomobject]@{ HeaderText=''; Selectable=@(); Always=@() }
    }
    return (Read-ConclusionsJson $jsonPath $reportType)
}

function Select-Conclusions {
    # $conclusions   : array d'objectes {Title, Body} (de Read-Conclusions).
    # $always        : array de cadenes ::SEMPRE:: (es mostren com a fixes).
    # $fields        : diccionari de camps compartit (s'hi afegeixen/editen els
    #                  [OPCIO:]/[CAMP:] de les conclusions, inline).
    # $preloadTitles : array de titols preseleccionats (sessio anterior).
    # $preloadValues : valors de camps precarregats (sessio anterior).
    param($conclusions, $always = @(), $fields = $null, $preloadTitles = $null, $preloadValues = $null)
    if ($null -eq $fields) { $fields = [ordered]@{} }
    if ($conclusions.Count -eq 0) { return [pscustomobject]@{ Nav='next'; Data=@() } }

    # Convertim preloadTitles a un HashSet per a comparacio rapida.
    $preloadSet = New-Object System.Collections.Generic.HashSet[string]
    if ($preloadTitles) { foreach ($t in $preloadTitles) { [void]$preloadSet.Add([string]$t) } }

    $form = _NewForm
    $form.Text = 'Pas 4 - Conclusions'
    $form.Size = New-Object System.Drawing.Size(940, 738)
    $form.StartPosition = 'CenterScreen'

    # Espai reservat a dalt per la banda de marca + barra de passos.
    $topOffset = 78

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = 'Marca les conclusions a incloure i omple-hi les opcions/camps:'
    $lbl.Location = New-Object System.Drawing.Point(15, (10 + $topOffset))
    $lbl.AutoSize = $true
    $lbl.Anchor = 'Top, Left, Right'
    $form.Controls.Add($lbl)

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Location = New-Object System.Drawing.Point(15, (35 + $topOffset))
    $panel.Size = New-Object System.Drawing.Size(895, 545)
    $panel.AutoScroll = $true
    $panel.BorderStyle = 'FixedSingle'
    $panel.Anchor = 'Top, Bottom, Left, Right'
    $form.Controls.Add($panel)

    $listV = New-Object System.Windows.Forms.FlowLayoutPanel
    $listV.FlowDirection = 'TopDown'
    $listV.WrapContents = $false
    $listV.AutoSize = $true
    $listV.AutoSizeMode = 'GrowAndShrink'
    $listV.Location = New-Object System.Drawing.Point(0, 0)
    $panel.Controls.Add($listV)

    $registry = _NewFieldRegistry

    # Amplada util (sense barra de desplacament). El cos de cada conclusio
    # s'embolcalla (wrap) a aquesta amplada; cal MaximumSize perque un
    # FlowLayoutPanel amb AutoSize+WrapContents salti de linia en lloc de
    # creixer cap a la dreta.
    $innerW = { [Math]::Max(280, $panel.ClientSize.Width - 26) }
    $mkConclFlow = {
        param($leftMargin)
        $f = New-Object System.Windows.Forms.FlowLayoutPanel
        $f.FlowDirection = 'LeftToRight'; $f.WrapContents = $true
        $f.AutoSize = $true; $f.AutoSizeMode = 'GrowAndShrink'
        $f.Margin = New-Object System.Windows.Forms.Padding($leftMargin, 0, 2, 6)
        $f.MaximumSize = New-Object System.Drawing.Size([Math]::Max(120, ((& $innerW) - $leftMargin - 6)), 0)
        return $f
    }

    $checks = @()
    for ($i = 0; $i -lt $conclusions.Count; $i++) {
        $c = $conclusions[$i]
        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Text = $c.Title
        $cb.AutoSize = $true
        $cb.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
        $cb.Margin = New-Object System.Windows.Forms.Padding(4, 12, 4, 2)
        $cb.Tag = $c
        if ($preloadSet.Contains([string]$c.Title)) { $cb.Checked = $true }
        [void]$listV.Controls.Add($cb)
        $checks += $cb

        $flow = & $mkConclFlow 22
        _RenderRichInto $flow ([string]$c.Body) $fields $preloadValues $registry
        [void]$listV.Controls.Add($flow)
    }

    # Frases fixes (::SEMPRE::): es mostren perque l'usuari les vegi i, si tenen
    # algun [CAMP:]/[OPCIO:], les pugui omplir (al document hi van sempre).
    $alwaysArr = @($always)
    if ($alwaysArr.Count -gt 0) {
        $sep = New-Object System.Windows.Forms.Label
        $sep.Text = 'Es posa sempre al final:'
        $sep.AutoSize = $true
        $sep.ForeColor = [System.Drawing.Color]::DimGray
        $sep.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Italic)
        $sep.Margin = New-Object System.Windows.Forms.Padding(4, 16, 4, 2)
        [void]$listV.Controls.Add($sep)
        foreach ($a in $alwaysArr) {
            $aflow = & $mkConclFlow 22
            _RenderRichInto $aflow ([string]$a) $fields $preloadValues $registry
            [void]$listV.Controls.Add($aflow)
        }
    }

    # En canviar la mida de la finestra, reembolcallem els cossos a l'ample nou.
    $panel.add_SizeChanged({
        $iw = & $innerW
        $listV.MaximumSize = New-Object System.Drawing.Size($iw, 0)
        foreach ($child in $listV.Controls) {
            if ($child -is [System.Windows.Forms.FlowLayoutPanel]) {
                $child.MaximumSize = New-Object System.Drawing.Size([Math]::Max(120, ($iw - $child.Margin.Left - 6)), 0)
            }
        }
    })

    $back = New-Object System.Windows.Forms.Button
    $back.Text = ([char]0x2190 + ' Enrere')
    $back.Location = New-Object System.Drawing.Point(15, 670)
    $back.Size = New-Object System.Drawing.Size(100, 30)
    $back.DialogResult = 'Retry'
    $back.Anchor = 'Bottom, Left'
    _StyleSecondaryButton $back
    $form.Controls.Add($back)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = ('Generar informe ' + [char]0x2192)
    $ok.Location = New-Object System.Drawing.Point(775, 670)
    $ok.Size = New-Object System.Drawing.Size(135, 30)
    $ok.DialogResult = 'OK'
    $ok.Anchor = 'Bottom, Right'
    _StylePrimaryButton $ok
    $form.AcceptButton = $ok
    $form.Controls.Add($ok)

    [void](_AddStepBar $form 4)
    [void](_AddBrandHeader $form "Conclusions de l'informe" $null 44)

    $res = $form.ShowDialog()
    if ($res -eq 'Retry') { return [pscustomobject]@{ Nav='back' } }
    if ($res -ne 'OK')    { exit 0 }

    # Retornem els objectes triats (no nomes el text), per preservar
    # Title i Body per al desat de sessio i l'emissio al document.
    $selected = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $checks.Count; $i++) {
        if ($checks[$i].Checked) { [void]$selected.Add($checks[$i].Tag) }
    }
    return [pscustomobject]@{ Nav='next'; Data=(,$selected.ToArray()) }
}

# ----------------------------------------------------------------------------
# Step 6 - Compose final document
# ----------------------------------------------------------------------------
# Retalla 0 CAPCALERA.docx per quedar-se nomes amb el bloc de capcalera demanat.
# El document pot contenir DUES capcaleres: la de REQ1 (a dalt, l'original) i la
# d'ACT_EXTR (a sota), separades per un paragraf marcador "[[CAP:ACT_EXTR]]".
#   $which = 'REQ1'     -> esborra des del marcador fins al final (i el marcador).
#   $which = 'ACT_EXTR' -> esborra des de l'inici fins al marcador (inclos).
# Si el marcador no existeix (capcalera antiga, nomes REQ1), no fa res. Aixi es
# retrocompatible amb una 0 CAPCALERA.docx que encara no tingui el bloc ACT_EXTR.
function Select-CapcaleraBlock($doc, [string]$which) {
    $marker = $null
    foreach ($p in $doc.Paragraphs) {
        $t = $p.Range.Text.TrimEnd("`r","`n","`a"," ")
        if ($t.Trim() -eq '[[CAP:ACT_EXTR]]') { $marker = $p; break }
    }
    if ($null -eq $marker) { return }   # nomes hi ha la capcalera REQ1: res a fer
    if ($which -eq 'ACT_EXTR') {
        # Esborra tot el que hi ha ABANS del marcador (bloc REQ1 + taula) i el
        # propi marcador.
        $rng = $doc.Range(0, $marker.Range.End)
        $rng.Delete() | Out-Null
    } else {
        # REQ1: esborra des del marcador (inclos) fins al final del document.
        $rng = $doc.Range($marker.Range.Start, $doc.Content.End)
        $rng.Delete() | Out-Null
    }
}

# Munta el text de la linia "Objecte:" (placeholder <<ORIGEN>>) de la capcalera
# generica a partir de l'origen triat al Pas 2. Funcio PURA (nomes llegeix el
# hashtable/objecte $header), testejable en headless. Accents amb codepoint
# per no dependre de l'encoding amb que PowerShell 5.1 llegeix aquest fitxer.
#   'doc'  -> "Doc. aportada amb Num. d'anotacio <NUM_ANOTACIO> del <DATA_ANOTACIO>"
#   'insp' -> "Visita inspeccio <DATA_INSPECCIO>"
#   'cap'  -> '' (sense Objecte; p.ex. requeriments de control periodic)
function _BuildOrigenText($header) {
    $get = {
        param($k)
        if ($null -eq $header) { return '' }
        if ($header -is [System.Collections.IDictionary]) { if ($header.Contains($k)) { return [string]$header[$k] }; return '' }
        if ($header.PSObject.Properties.Name -contains $k) { return [string]$header.$k }
        return ''
    }
    $tipus = (& $get 'ORIGEN_TIPUS'); if ([string]::IsNullOrWhiteSpace($tipus)) { $tipus = 'doc' }
    if ($tipus -eq 'cap')  { return '' }
    if ($tipus -eq 'insp') {
        return 'Visita inspecci' + [char]0x00F3 + ' ' + (& $get 'DATA_INSPECCIO')
    }
    return 'Doc. aportada amb N' + [char]0x00FA + 'm. d' + [char]0x2019 + 'anotaci' + [char]0x00F3 +
           ' ' + (& $get 'NUM_ANOTACIO') + ' del ' + (& $get 'DATA_ANOTACIO')
}

function Apply-HeaderReplacements($doc, $header) {
    # Substituim els placeholders <<NOM>> de la capcalera pels valors del Pas 2.
    # Inclou els d'ACT_EXTR (<<DATES>>, <<AFORAMENT>>); si no apareixen a la
    # capcalera triada, simplement no es substitueix res.
    $get = {
        param($k)
        if ($null -eq $header) { return '' }
        if ($header -is [System.Collections.IDictionary]) { if ($header.Contains($k)) { return [string]$header[$k] }; return '' }
        if ($header.PSObject.Properties.Name -contains $k) { return [string]$header.$k }
        return ''
    }
    # <<ORIGEN>>: la linia "Objecte:" de la capcalera generica (REQ1) es
    # construeix segons el que s'ha triat al Pas 2 (documentacio aportada o
    # visita d'inspeccio). Vegeu _BuildOrigenText (funcio pura, testejable).
    $origen = _BuildOrigenText $header
    $map = @{
        '<<ID_GIA>>'         = (& $get 'ID_GIA')
        '<<EXP_NUM>>'        = (& $get 'EXP_NUM')
        '<<ADRECA>>'         = (& $get 'ADRECA')
        '<<ACTIVITAT>>'      = (& $get 'ACTIVITAT')
        '<<TITULAR>>'        = (& $get 'TITULAR')
        '<<ORIGEN>>'         = $origen
        '<<NUM_ANOTACIO>>'   = (& $get 'NUM_ANOTACIO')
        '<<DATA_ANOTACIO>>'  = (& $get 'DATA_ANOTACIO')
        '<<DATA_INSPECCIO>>' = (& $get 'DATA_INSPECCIO')
        '<<DATES>>'          = (& $get 'DATES')
        '<<AFORAMENT>>'      = (& $get 'AFORAMENT')
    }
    foreach ($k in $map.Keys) {
        $find = $doc.Content.Find
        $find.ClearFormatting()
        $find.Replacement.ClearFormatting()
        $find.Text = $k
        $find.Replacement.Text = [string]$map[$k]
        $find.Forward = $true
        $find.Wrap = 1
        $find.MatchCase = $false
        $find.Execute([ref]$k, $false, $false, $false, $false, $false, $true, 1, $false, [string]$map[$k], 2) | Out-Null
    }
}

# Calcula el nom de fitxer de sortida: YYYY-MM-DD_<TipusCataleg>_GIA <id>.docx
function _GetOutputFileName($catalegName, $gia) {
    $today = (Get-Date).ToString('yyyy-MM-dd')
    $cat   = $catalegName
    if ($cat) { $cat = $cat.Substring(0,1).ToUpper() + $cat.Substring(1).ToLower() }
    else      { $cat = 'Informe' }
    if ([string]::IsNullOrWhiteSpace($gia)) { $gia = 's_n' }
    $gia = ($gia -replace '[\\/:*?"<>|]','_').Trim()
    return ("{0}_{1}_GIA {2}.docx" -f $today, $cat, $gia)
}

# A partir del nom base, retorna la ruta a $targetDir que no col·lisioni amb
# cap fitxer existent. Si el primer ja existeix, prova "_2", "_3"... fins
# trobar-ne un de lliure. Aixi pots fer diversos informes del mateix GIA el
# mateix dia sense haver de tancar Word ni renombrar res manualment.
#
# Ex.: "2026-05-29_Req1_GIA 1379.docx" existeix
#      -> torna "2026-05-29_Req1_GIA 1379_2.docx"
function _GetUniqueOutputPath($targetDir, $baseFileName) {
    $candidate = Join-Path $targetDir $baseFileName
    if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($baseFileName)
    $ext  = [System.IO.Path]::GetExtension($baseFileName)
    for ($i = 2; $i -lt 1000; $i++) {
        $candidate = Join-Path $targetDir ("{0}_{1}{2}" -f $stem, $i, $ext)
        if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
    }
    # Si arribem aqui es que hi ha mes de 999 informes pel mateix GIA i dia;
    # cas extrem, retornem un nom amb timestamp.
    return Join-Path $targetDir ("{0}_{1}{2}" -f $stem, (Get-Date -Format 'HHmmss'), $ext)
}

# Determina el directori de sortida: l'$OutputDir si es accessible, en cas
# contrari una subcarpeta 'Informes generats' a l'arrel del clone (al
# costat dels .bat).
function _ResolveOutputDir {
    $targetDir = $OutputDir
    try {
        if (-not (Test-Path -LiteralPath $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force -ErrorAction Stop | Out-Null
        }
        return $targetDir
    } catch {
        $local = Join-Path $RepoRoot 'Informes generats'
        if (-not (Test-Path -LiteralPath $local)) {
            New-Item -ItemType Directory -Path $local -Force | Out-Null
        }
        return $local
    }
}

# Obre el document Word a partir d'una copia LOCAL de la capcalera (per
# evitar la "Vista protegida" en unitats de xarxa). Retorna el doc obert i
# la ruta temporal.
function _OpenOutputDocument($word, $tempPath) {
    Copy-Item -LiteralPath $HeaderPath -Destination $tempPath -Force
    $doc = $word.Documents.Open($tempPath, $false, $false)
    try {
        if ($doc.ProtectedViewWindow -ne $null) {
            $doc = $doc.ProtectedViewWindow.Edit()
        }
    } catch { }
    return $doc
}

# Escriu el cos del document (intro del cataleg + seccions amb items numerats).
# Retorna el comptador global utilitzat per a la numeracio.
function _WriteCatalegBody($sel, $cfg, $selectedSections, $fields, $introText, $isFixedBody = $false, $fixedBodyLines = @()) {
    # Informe de cos fix (p.ex. TERMINI.docx): no hi ha seccions ni items a
    # numerar; el cos son directament els paragrafs del document, amb els camps
    # [CAMP:]/[OPCIO:] resolts i separant text/URLs com a la resta del motor.
    if ($isFixedBody) {
        $lines = @($fixedBodyLines)
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $resolved = [string](Apply-Fields -text $lines[$i] -fields $fields)
            if ([string]::IsNullOrWhiteSpace($resolved)) { continue }
            $parts = _SplitTextAndUrls $resolved
            if (-not [string]::IsNullOrWhiteSpace($parts.Text)) { Format-Body $sel $parts.Text }
            foreach ($u in $parts.Urls) { Format-Url $sel $u }
            if ($i -lt ($lines.Count - 1)) { Format-Spacer $sel }
        }
        return
    }

    if (-not [string]::IsNullOrWhiteSpace($introText)) {
        Format-Body $sel $introText
        if ($cfg.SpacerAfterIntroParagraph) { Format-Spacer $sel }
    }

    # Resol [CAMP: ...] a cada linia i EMET cada linia resolta al pipeline.
    # Els cridadors fan servir @(& $resolveLines ...) per recollir un array
    # PLA de cadenes. (No retornem ,@(...): combinat amb el @() del cridador
    # provocava un doble embolcall on $itemLines[0] era TOT l'array de linies
    # en comptes de la primera linia -> trencava la separacio text/URL i feia
    # petar Substring.)
    $resolveLines = {
        param($lines)
        foreach ($ln in $lines) {
            [string](Apply-Fields -text $ln -fields $fields)
        }
    }

    # Emet una linia separant text i URLs: el text (si n'hi ha) va com a cos i
    # cada URL com a hipervincle en paragraf propi.
    $emitLine = {
        param($line, $isChild)
        if ([string]::IsNullOrWhiteSpace($line)) { return }
        $parts = _SplitTextAndUrls $line
        if (-not [string]::IsNullOrWhiteSpace($parts.Text)) {
            if ($isChild) { Format-Body $sel $parts.Text -IsChild } else { Format-Body $sel $parts.Text }
        }
        foreach ($u in $parts.Urls) {
            if ($isChild) { Format-Url $sel $u -IsChild } else { Format-Url $sel $u }
        }
    }

    $emitExtras = {
        param($lines, $isChild)
        for ($i = 1; $i -lt $lines.Count; $i++) {
            & $emitLine $lines[$i] $isChild
        }
    }

    $emitIntro = {
        param($introEl)
        $lines = @(& $resolveLines $introEl.BodyLines)
        foreach ($bp in $lines) { & $emitLine $bp $false }
        if ($cfg.SpacerAfterIntro) { Format-Spacer $sel }
    }

    $emitItem = {
        param($it)
        $itemLines = @(& $resolveLines $it.BodyLines)
        $hasChildren = ($it.Children.Count -gt 0)
        $itemWritten = $false

        if ($it.Selected -or $hasChildren) {
            if ($itemLines.Count -gt 0) {
                $script:_buildGlobal++
                # Separem un possible URL enganxat al text principal de l'item.
                $p0 = _SplitTextAndUrls $itemLines[0]
                Format-Item $sel "$($script:_buildGlobal)." $p0.Text
                foreach ($u in $p0.Urls) { Format-Url $sel $u }
                & $emitExtras $itemLines $false
                $itemWritten = $true
            }
        }
        if ($hasChildren) {
            $firstChild = $true
            foreach ($ch in $it.Children) {
                $childLines = @(& $resolveLines $ch.BodyLines)
                if ($childLines.Count -eq 0) { continue }
                if (-not $itemWritten) {
                    $script:_buildGlobal++
                    $itemWritten = $true
                }
                # Els fills (::CHILD::) NO es numeren: s'emeten com a PUNT DE
                # LLISTA amb Format-Bullet (-IsChild = sangria de sub-nivell).
                # Les linies extra del fill (p.ex. un URL) segueixen com a
                # enllac/cos de fill, sense un nou pic.
                $pc = _SplitTextAndUrls $childLines[0]
                if (-not [string]::IsNullOrWhiteSpace($pc.Text)) {
                    # -First: el primer sub-punt se separa de l'item numerat amb
                    # ItemSpaceAfterPt (els seguents, amb BulletSpaceBeforePt).
                    Format-Bullet $sel $pc.Text -IsChild -First:$firstChild
                    $firstChild = $false
                }
                foreach ($u in $pc.Urls) { Format-Url $sel $u -IsChild }
                & $emitExtras $childLines $true
            }
        }
        if ($itemWritten -and $cfg.SpacerAfterItem) { Format-Spacer $sel }
    }

    $script:_buildGlobal = 0
    $lastSectionName = $null

    foreach ($sec in $selectedSections) {
        $parts = $sec.Title -split ' - ', 2
        if ($parts.Count -eq 2) {
            $secName = $parts[0].Trim()
            $subName = $parts[1].Trim()
            if ($secName -ne $lastSectionName) {
                Format-Section $sel $secName
                if ($cfg.SpacerAfterSection) { Format-Spacer $sel }
                $lastSectionName = $secName
            }
            Format-Subsection $sel $subName
            if ($cfg.SpacerAfterSubsection) { Format-Spacer $sel }
        } else {
            Format-Section $sel $sec.Title
            if ($cfg.SpacerAfterSection) { Format-Spacer $sel }
            $lastSectionName = $sec.Title
        }

        # Les subseccions i les intros s'emeten "tard": nomes quan ve un
        # item REAL que les segueix. Si una secció conté 3 ::SUB:: pero
        # nomes s'ha triat un ítem que viu a la 3a subsecció, només
        # surten el títol de la secció i la 3a subsecció (no les 2
        # anteriors buides). Una nova subsecció sobreescriu la pendent.
        $pendingSubsection = $null
        $pendingIntro = $null
        foreach ($el in $sec.Items) {
            if ($el.Kind -eq 'subsection') {
                $pendingSubsection = $el
                $pendingIntro = $null   # una nova subseccio invalida l'intro pendent
                continue
            }
            if ($el.Kind -eq 'intro') {
                $pendingIntro = $el
                continue
            }
            # Item real: emetem primer la subseccio pendent (si en hi ha),
            # despres l'intro pendent, i finalment l'item.
            if ($null -ne $pendingSubsection) {
                Format-Subsection $sel $pendingSubsection.Short
                if ($cfg.SpacerAfterSubsection) { Format-Spacer $sel }
                $pendingSubsection = $null
            }
            if ($null -ne $pendingIntro) {
                & $emitIntro $pendingIntro
                $pendingIntro = $null
            }
            & $emitItem $el
        }
    }
}

# Escriu el bloc de conclusions:
#   - $headerText : titol del bloc (sol ser 'CONCLUSIONS'), centrat-negreta.
#                   '' = no s'emet.
#   - $conclusions : array d'objectes {Title; Body} de les conclusions
#                    TRIADES al Pas 5. Es emet el seu Body.
#   - $alwaysConclusions : array de cadenes ja sense el prefix '::SEMPRE::'
#                          que s'emeten sempre, despres de les triades.
#   - $fields : per resoldre [CAMP:] i [OPCIO:] dins els textos.
# La separacio entre conclusions la posa Format-Conclusion via SpaceAfter
# (ConclusionSpaceAfterPt), aixi que aqui no hi fa falta un Spacer entre.
function _WriteConclusionsBlock($sel, $cfg, $headerText, $conclusions, $alwaysConclusions, $fields) {
    $hasBody = ($conclusions.Count -gt 0) -or ($alwaysConclusions.Count -gt 0)
    $hasHead = -not [string]::IsNullOrWhiteSpace($headerText)
    if (-not $hasBody -and -not $hasHead) { return }

    if ($cfg.SpacerBeforeConclusionsBlock) { Format-Spacer $sel }

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
    $baseName  = _GetOutputFileName $catalegName $header['ID_GIA']
    $targetDir = _ResolveOutputDir
    # Triem el primer nom lliure al directori de sortida (afegim _2, _3...
    # si ja existeix). Aixi pots generar diversos informes del mateix dia/GIA
    # sense que cap es sobreescrigui.
    $outPath  = _GetUniqueOutputPath $targetDir $baseName
    $fileName = [System.IO.Path]::GetFileName($outPath)

    # Treballem amb una copia LOCAL (a %TEMP%) per evitar que Word obri el
    # fitxer en "Vista protegida" quan el desti es una unitat de xarxa.
    # El temp porta el mateix nom (ja unic) que el desti final.
    $tempPath = Join-Path $env:TEMP $fileName
    $doc = _OpenOutputDocument $word $tempPath

    # 0 CAPCALERA.docx pot portar tambe el bloc d'ACT_EXTR a sota; ens quedem
    # nomes amb el bloc de REQ1 (no fa res si el marcador no hi es).
    Select-CapcaleraBlock $doc 'REQ1'
    Apply-HeaderReplacements -doc $doc -header $header

    $doc.Activate()
    $sel = $word.Selection
    [void]$sel.EndKey(6)  # wdStory = 6

    $cfg = $Script:ReportFormatConfig
    _WriteCatalegBody $sel $cfg $selectedSections $fields $introText $isFixedBody $fixedBodyLines
    _WriteConclusionsBlock $sel $cfg $conclusionsHeaderText $conclusions $alwaysConclusions $fields

    $doc.Save()
    $doc.Close($false)

    # Movem el fitxer al desti final (xarxa o local segons disponibilitat).
    try {
        Move-Item -LiteralPath $tempPath -Destination $outPath -Force
    } catch {
        return $tempPath
    }
    return $outPath
}

# ============================================================================
# Mode "des de paquet" (generar des del mobil)
# ----------------------------------------------------------------------------
# Genera un informe a partir d'un paquet JSON (el mateix model que
# lastreport.json) en lloc de l'assistent WinForms. El paquet el prepara el
# formulari web del mobil i el porta fins aqui Vigilant.ps1 (o es pot passar a
# ma). Necessita Word, com el flux normal, pero es 100% no interactiu.
#
# Clau del disseny: REAPROFITA el mateix motor que el flux normal
# (Get-ParsedCataleg, Read-Conclusions, Build-Document). Nomes canvia QUI omple les
# dades: en comptes dels dialegs WinForms, surten del paquet. Les tres funcions
# Build-*FromPaquet son PURES (sense Word/UI) i es proven als tests.
# ============================================================================

# Reconstrueix l'estructura de seleccio que retorna Select-Items (Pas 3) a
# partir d'una llista de claus "Seccio::Item[::Fill]". Es l'invers de
# Get-SelectedKeysFromResult i replica EXACTAMENT la logica de construccio del
# resultat del Pas 3 (Select-Items), pero sense UI.
function Build-SelectionFromKeys($sections, $selectedKeys) {
    $checkStates = @{}
    foreach ($k in $selectedKeys) {
        if ([string]::IsNullOrWhiteSpace($k)) { continue }
        $checkStates[[string]$k] = $true
    }

    $result = New-Object System.Collections.ArrayList
    foreach ($sec in $sections) {
        $chosen = New-Object System.Collections.ArrayList
        foreach ($el in $sec.Items) {
            if ($el.Kind -in 'subsection','intro') {
                [void]$chosen.Add([pscustomobject]@{
                    Kind      = $el.Kind
                    Short     = $el.Short
                    BodyLines = $el.BodyLines
                    Children  = @()
                    Selected  = $false
                })
                continue
            }
            $itKey = (_ItemKey $sec.Title $el.Short)
            $isSel = $checkStates.ContainsKey($itKey) -and $checkStates[$itKey]
            $chosenChildren = New-Object System.Collections.ArrayList
            foreach ($ch in $el.Children) {
                $chKey = (_ItemKey $sec.Title $el.Short $ch.Short)
                if ($checkStates.ContainsKey($chKey) -and $checkStates[$chKey]) {
                    [void]$chosenChildren.Add($ch)
                }
            }
            if ($isSel -or $chosenChildren.Count -gt 0) {
                [void]$chosen.Add([pscustomobject]@{
                    Kind      = 'item'
                    Short     = $el.Short
                    BodyLines = $el.BodyLines
                    Children  = $chosenChildren
                    Selected  = [bool]$isSel
                })
            }
        }
        $hasRealItem = $false
        foreach ($x in $chosen) { if ($x.Kind -eq 'item') { $hasRealItem = $true; break } }
        if ($hasRealItem) {
            [void]$result.Add([pscustomobject]@{
                Title = $sec.Title
                Items = $chosen
            })
        }
    }
    return $result
}

# Reconstrueix la llista de conclusions triades a partir dels seus titols,
# preservant l'ordre del fitxer (com fa Select-Conclusions, que itera les
# checkboxes en ordre del panell). $selectable son els objectes {Title, Body}
# de Read-Conclusions.
function Build-ConclusionsFromTitles($selectable, $titles) {
    $titleSet = New-Object System.Collections.Generic.HashSet[string]
    if ($titles) { foreach ($t in $titles) { [void]$titleSet.Add([string]$t) } }
    $out = New-Object System.Collections.ArrayList
    foreach ($c in $selectable) {
        if ($titleSet.Contains([string]$c.Title)) { [void]$out.Add($c) }
    }
    return $out.ToArray()
}

# Construeix el diccionari de camps (com els passos 4-5) i hi aplica els valors
# del paquet. $fieldValues pot ser un hashtable o un PSCustomObject (tal com el
# torna ConvertFrom-Json). Els camps sense valor al paquet queden amb el seu
# valor per defecte (per als desplegables, la primera opcio).
function Build-FieldsFromPaquet($selectedSections, $conclusions, $alwaysConcl, $fieldValues) {
    $fields = Get-FieldsFromSelection $selectedSections
    Add-FieldsFromConclusions $fields $conclusions $alwaysConcl
    foreach ($name in @($fields.Keys)) {
        $v = $null
        if ($fieldValues -is [System.Collections.IDictionary]) {
            if ($fieldValues.Contains($name)) { $v = $fieldValues[$name] }
        } elseif ($null -ne $fieldValues -and ($fieldValues.PSObject.Properties.Name -contains $name)) {
            $v = $fieldValues.$name
        }
        if ($null -ne $v) { $fields[$name].Value = [string]$v }
    }
    return $fields
}

# Orquestrador del mode paquet. Llegeix el JSON, resol el cataleg, completa la
# capcalera des de l'Excel si cal (el PC si que hi te acces), reconstrueix les
# seleccions i crida Build-Document. Torna la ruta del .docx generat.
function Invoke-GenerateFromPaquet($paquetPath) {
    if (-not (Test-Path -LiteralPath $paquetPath)) {
        throw "No s'ha trobat el paquet: $paquetPath"
    }
    $raw = Get-Content -LiteralPath $paquetPath -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($raw)) { throw "El paquet esta buit: $paquetPath" }
    $pkg = $raw | ConvertFrom-Json

    $catName = [string]$pkg.CatalegBaseName
    if ([string]::IsNullOrWhiteSpace($catName)) { throw "El paquet no indica 'CatalegBaseName'." }
    $catPath = Join-Path $EstructuralsDir ($catName + '.json')
    if (-not (Test-Path -LiteralPath $catPath)) {
        throw "No s'ha trobat el cataleg '$catName' a $EstructuralsDir."
    }

    # Capcalera: hashtable amb les claus que espera Apply-HeaderReplacements.
    $header = @{}
    foreach ($k in 'ID_GIA','EXP_NUM','ADRECA','ACTIVITAT','TITULAR','ORIGEN_TIPUS','NUM_ANOTACIO','DATA_ANOTACIO','DATA_INSPECCIO') {
        $val = ''
        if ($null -ne $pkg.Header -and ($pkg.Header.PSObject.Properties.Name -contains $k)) {
            $val = [string]$pkg.Header.$k
        }
        $header[$k] = $val
    }
    # El mobil no ofereix la tria d'origen: per defecte, documentacio aportada
    # (mateix comportament que abans, la linia "Objecte:" mostra la doc. aportada).
    if ([string]::IsNullOrWhiteSpace($header['ORIGEN_TIPUS'])) { $header['ORIGEN_TIPUS'] = 'doc' }

    # Si el paquet ve del mobil amb (gairebe) nomes l'ID GIA, completem la
    # capcalera des de l'Excel d'activitats. El PC si que hi te acces; el mobil
    # no (les dades personals no surten mai a la web). Nomes omplim els buits.
    if ([string]::IsNullOrWhiteSpace($header['TITULAR']) -and -not [string]::IsNullOrWhiteSpace($header['ID_GIA'])) {
        try {
            $xls = Find-LatestActivitatsExcel
            if ($null -ne $xls) {
                $cache = Initialize-ActivitatsCache $xls.File
                $act = Get-ActivitatFromCache $cache $header['ID_GIA']
                if ($null -ne $act) {
                    foreach ($k in 'TITULAR','ADRECA','ACTIVITAT','EXP_NUM','NUM_ANOTACIO','DATA_ANOTACIO') {
                        if ([string]::IsNullOrWhiteSpace($header[$k]) -and $act.ContainsKey($k)) {
                            $header[$k] = [string]$act[$k]
                        }
                    }
                }
            }
        } catch {
            # Sense Excel/xarxa seguim amb el que porti el paquet. No es fatal.
            Write-Host "Avis: no s'ha pogut completar la capcalera des de l'Excel ($($_.Exception.Message))."
        }
    }

    $selectedKeys     = @(); if ($pkg.SelectedKeys)    { $selectedKeys     = @($pkg.SelectedKeys) }
    $conclusionTitles = @(); if ($pkg.ConclusionTexts) { $conclusionTitles = @($pkg.ConclusionTexts) }
    $fieldValues      = $pkg.FieldValues

    $word = New-WordApp
    try {
        $parsed   = Get-ParsedCataleg -path $catPath
        $selected = Build-SelectionFromKeys $parsed.Sections $selectedKeys
        # Els informes de cos fix (p.ex. TERMINI) no seleccionen deficiencies:
        # nomes els que tenen seccions exigeixen alguna seleccio valida.
        if (-not $parsed.IsFixedBody -and $selected.Count -eq 0) {
            throw "El paquet no selecciona cap deficiencia valida per al cataleg '$catName'."
        }

        $conclAll    = Read-Conclusions -path $ConclusionsPath -reportType $catName
        $conclusions = Build-ConclusionsFromTitles $conclAll.Selectable $conclusionTitles
        $fields      = Build-FieldsFromPaquet $selected $conclusions $conclAll.Always $fieldValues

        $outPath = Build-Document -word $word -header $header `
                                  -selectedSections $selected `
                                  -fields $fields `
                                  -conclusions $conclusions `
                                  -alwaysConclusions $conclAll.Always `
                                  -catalegName $catName `
                                  -introText $parsed.IntroText `
                                  -conclusionsHeaderText $conclAll.HeaderText `
                                  -isFixedBody $parsed.IsFixedBody `
                                  -fixedBodyLines $parsed.FixedBodyLines

        Save-LastReport ([ordered]@{
            Version         = 1
            Timestamp       = (Get-Date).ToString('o')
            CatalegBaseName = $catName
            Header          = $header
            SelectedKeys    = (Get-SelectedKeysFromResult $selected)
            FieldValues     = (Get-FieldValuesForSession $fields)
            ConclusionTexts = $conclusionTitles
        })

        Write-Host "Informe generat: $outPath"
        return $outPath
    } finally {
        Close-WordApp $word
    }
}

# ----------------------------------------------------------------------------
# Main flow
# ----------------------------------------------------------------------------
function Main {
    if (-not (Test-Path $HeaderPath)) {
        [System.Windows.Forms.MessageBox]::Show("No s'ha trobat $HeaderPath",'Error','OK','Error') | Out-Null
        exit 1
    }

    # Pas 1: un sol menu (Select-Mode) que tria alhora el MODE i, per al cas
    # "nou", el CATALEG (ja no hi ha un segon pas de tria). Cada flux SEMPRE
    # torna a aquest menu quan acaba o quan es prem Enrere; aixi el programa
    # ROMAN OBERT. L'UNICA manera de sortir del programa es tancar la finestra
    # (X) d'aquest menu inicial (Select-Mode fa exit 0).
    while ($true) {
        $sel = Select-Mode
        switch ($sel.Action) {
            'seguiment'  { Invoke-SeguimentFlow }
            'actextr'    { Invoke-ActExtrFlow }
            'ruta'       { Start-RutaTool }   # llanca el planificador; torna al menu
            'controlsperiodics' { Invoke-ControlsPeriodics }   # llistat d'activitats amb control periodic (Excel)
            'informesdb'     { Invoke-InformesDbScan }   # escaneja informes -> JSON; torna al menu
            'informesdbedit' { Invoke-InformesDbEdit }   # editor de la base d'informes
            'copiarinformes' { Invoke-CopiarInformes }   # copia incremental dels Word a la carpeta de copia
            'comprovarexcel' { Invoke-ComprovarExcel }   # comprova que l'Excel reflecteix els precintes
            'revisarmobil'   { Invoke-RevisarMobil }     # revisa el mobil un sol cop; torna al menu
            'config'         { Invoke-ConfiguracioScreen }   # rutes d'aquest PC + actualitzar; torna al menu
            'editcataleg'    { Show-CatalegEditor -focusDoc ([string]$sel.Doc) }   # editor dels ESTRUCTURALS (xip del document)
            'convertirpdf'   { Invoke-ConvertirPdf }   # converteix una carpeta de Word a PDF (i signa)
            'emailtextos'    { Invoke-EmailTextos }    # edita els textos del correu del mobil
            'nou'        { [void](Invoke-NouWizard -cataleg $sel.Cataleg) }
            default      { return }
        }
        # ...i es torna a mostrar el menu (Pas 1).
    }
}

# Wizard de "generar informe nou" (Pas 2..5). El cataleg ja ve triat del Pas 1.
#
# Assistent navegable. Cada pas (dialeg) retorna un objecte amb:
#   Nav  = 'next' | 'back' | 'stay'   ·   Data = el resultat del pas (si 'next')
# El boto "Enrere" retorna 'back' i el wizard torna al pas anterior conservant
# les dades (precarrega). Enrere al Pas 2 => torna al menu inicial (retorna
# 'menu'). Tancar una finestra (X) avorta tot el programa (exit 0).
#
# $st  : dades confirmades de cada pas (es mantenen en memoria al navegar).
# $pre : precarregues per a cada pas (de l'estat o de "Recuperar ultim").
# $st.Fields es un diccionari de camps COMPARTIT: els [OPCIO:]/[CAMP:] s'omplen
# inline alla on apareixen (Pas 3 i Pas 4).
#
# Retorna 'menu' (Enrere al Pas 2) o 'done' (informe generat).
function Invoke-NouWizard {
    param($cataleg)

    $st  = @{ Cataleg=$cataleg; Parsed=$null; Header=$null; Selected=$null; Fields=[ordered]@{}; ConclAll=$null; Conclusions=$null }
    $pre = @{ Header=$null; Keys=$null; Fields=$null; Concl=$null }

    # RENDIMENT: NO obrim Word ni parsejem el cataleg aqui. El Pas 2 (dades de
    # la capcalera) es WinForms pur i no necessita Word, aixi que apareix de
    # seguida en clicar el tipus d'informe. Word (arrencada "en fred", lenta el
    # primer cop) i el parseig del cataleg es fan de forma DIFERIDA quan es
    # necessiten per primer cop (Pas 3), mentre l'usuari omple la capcalera.
    $word = $null
    try {
        $step = 2
        $dir  = 'fwd'
        while ($step -ge 2 -and $step -le 5) {
            switch ($step) {

                2 {
                    $r = Get-HeaderData -preload $pre.Header
                    if ($r.Nav -eq 'back') { return 'menu' }   # enrere Pas 2 = tornar al menu
                    else {
                        $st.Header  = $r.Data
                        $pre.Header = $r.Data
                        if ($r.Recovered) {
                            # "Recuperar dades ultim informe": precarreguem la
                            # resta de passos amb les dades de l'ultim informe.
                            $pre.Keys   = $r.Recovered.SelectedKeys
                            $pre.Fields = $r.Recovered.FieldValues
                            $pre.Concl  = $r.Recovered.ConclusionTexts
                        }
                        $step = 3; $dir = 'fwd'
                    }
                }

                3 {
                    # Primer cop que necessitem Word i el cataleg parsejat:
                    # arrenquem Word (diferit) i parsegem ara (amb cache).
                    if ($null -eq $st.Parsed) {
                        if ($null -eq $word) { $word = New-WordApp }
                        $st.Parsed = Get-ParsedCataleg -path $st.Cataleg.FullName
                    }
                    if ($st.Parsed.IsFixedBody) {
                        # Informe de cos fix (p.ex. TERMINI): no hi ha
                        # deficiencies a triar. Saltem el Pas 3.
                        $st.Selected = @()
                        if ($dir -eq 'back') { $step = 2; $dir = 'back' }
                        else                 { $step = 4; $dir = 'fwd' }
                    } else {
                        # Pas 3: triar deficiencies I omplir-ne les opcions/camps
                        # inline (al mateix panell de detall).
                        $r = Select-Items -sections $st.Parsed.Sections -preloadSelectedKeys $pre.Keys -fields $st.Fields -preloadValues $pre.Fields
                        if     ($r.Nav -eq 'back') { $step = 2; $dir = 'back' }
                        elseif ($r.Nav -eq 'stay') { }   # cap seleccio: es torna a mostrar
                        else {
                            $st.Selected = $r.Data
                            $pre.Keys    = Get-SelectedKeysFromResult $st.Selected
                            $pre.Fields  = Get-FieldValuesForSession $st.Fields
                            $step = 4; $dir = 'fwd'
                        }
                    }
                }

                # Pas 4 = CONCLUSIONS. El cos de cada conclusio es mostra sencer i
                # els seus [CAMP:]/[OPCIO:] s'omplen inline aqui mateix (ja no hi
                # ha un pas separat de camps).
                4 {
                    if ($null -eq $st.ConclAll) {
                        # Les conclusions triables depenen del tipus d'informe
                        # (BaseName del cataleg: REQ1, TERMINI...).
                        if ($null -eq $word) { $word = New-WordApp }
                        $st.ConclAll = Read-Conclusions -path $ConclusionsPath -reportType $st.Cataleg.BaseName
                    }
                    if ($st.ConclAll.Selectable.Count -eq 0) {
                        # No hi ha conclusions triables: saltem el pas.
                        $st.Conclusions = @()
                        if ($dir -eq 'back') { $step = 3; $dir = 'back' } else { $step = 5 }
                    } else {
                        $r = Select-Conclusions -conclusions $st.ConclAll.Selectable -always $st.ConclAll.Always -fields $st.Fields -preloadTitles $pre.Concl -preloadValues $pre.Fields
                        if ($r.Nav -eq 'back') { $step = 3; $dir = 'back' }
                        else {
                            $st.Conclusions = $r.Data
                            # $pre.Concl guarda nomes els TITOLS per a la
                            # precarrega de la propera vegada.
                            $pre.Concl  = @($st.Conclusions | ForEach-Object { $_.Title })
                            $pre.Fields = Get-FieldValuesForSession $st.Fields
                            $step = 5; $dir = 'fwd'
                        }
                    }
                }

                5 {
                    if ($null -eq $word) { $word = New-WordApp }
                    $outPath = Build-Document -word $word -header $st.Header `
                                              -selectedSections $st.Selected `
                                              -fields $st.Fields `
                                              -conclusions $st.Conclusions `
                                              -alwaysConclusions $st.ConclAll.Always `
                                              -catalegName $st.Cataleg.BaseName `
                                              -introText $st.Parsed.IntroText `
                                              -conclusionsHeaderText $st.ConclAll.HeaderText `
                                              -isFixedBody $st.Parsed.IsFixedBody `
                                              -fixedBodyLines $st.Parsed.FixedBodyLines

                    # Desem les dades per poder replicar aquest informe mes endavant.
                    # Per a 'ConclusionTexts' guardem els TITOLS triats (es el que
                    # fa servir Select-Conclusions per precarregar).
                    Save-LastReport ([ordered]@{
                        Version         = 1
                        Timestamp       = (Get-Date).ToString('o')
                        CatalegBaseName = $st.Cataleg.BaseName
                        Header          = $st.Header
                        SelectedKeys    = (Get-SelectedKeysFromResult $st.Selected)
                        FieldValues     = (Get-FieldValuesForSession $st.Fields)
                        ConclusionTexts = @($st.Conclusions | ForEach-Object { $_.Title })
                    })

                    [System.Windows.Forms.MessageBox]::Show(
                        "Informe generat:`n$outPath",
                        'Finalitzat', 'OK', 'Information') | Out-Null

                    # Obrim Word en primer pla per a l'usuari
                    $word.Visible = $true
                    $word.Documents.Open($outPath) | Out-Null
                    $step = 99   # surt del bucle
                }
            }
        }
        return 'done'
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)",'Error','OK','Error') | Out-Null
        throw
    }
    finally {
        # Si mai vam arribar a obrir Word (p.ex. Enrere al Pas 2), no hi ha res
        # a tancar. Si el vam obrir pero no es va fer visible (l'usuari va
        # cancel-lar abans de generar), el tanquem. Si es va fer visible (informe
        # generat i obert), el deixem obert per a l'usuari.
        if ($null -ne $word -and -not $word.Visible) { Close-WordApp $word }
    }
}

