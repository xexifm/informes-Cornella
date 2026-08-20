#requires -Version 5.1
<#
.SYNOPSIS
  MOTOR del generador d'informes de l'Ajuntament de Cornella: rutes,
  configuracio i carrega dels moduls.

.DESCRIPTION
  Aquest fitxer NOMES DEFINEIX: rutes, configuracio i unes quantes funcions de
  base. Carregar-lo (dot-source) no obre cap finestra ni genera cap informe. Qui
  l'arrenca es GenerarInforme.ps1 (el punt d'entrada, que fa doble clic
  l'usuari); qui el reutilitza com a biblioteca son mobil/Vigilant.ps1,
  mobil/ExportaDades.ps1 i les proves de suport/tests/.

  MAPA DELS MODULS. Motor.ps1 havia arribat a 3.300 linies i 70 funcions, amb
  tot barrejat. Ara cada concepte te el seu fitxer i aqui nomes queda la base:

    -- el fil de l'aplicacio --
    Wizard.ps1          Main (el menu) i l'assistent de "Requeriment - Nou"
    Seguiment.ps1       pantalla inicial (mode+cataleg) i informe de seguiment
    Paquet.ps1          generar sense assistent, des d'un paquet del mobil

    -- els passos de l'assistent --
    Capcalera.ps1       Pas 2: dades de l'activitat
    SeleccioItems.ps1   Pas 3 i Pas 5: triar deficiencies i conclusions
    Camps.ps1           [CAMP:]/[OPCIO:], text ric i separacio d'enllacos
    Document.ps1        Pas 6: composicio del .docx final
    Format.ps1          COM es veu el document (lletra, sangries, espaiats)

    -- dades --
    CatalegJson.ps1     lectura dels catalegs (ESTRUCTURALS\*.json)
    Activitats.ps1      Excel d'activitats: cache per ID GIA + pujada a Drive
    Informes.ps1        escaneig dels informes ja fets (informes-db.json)
    Migracio.ps1        rutes de local\ i endrec de les carpetes velles

    -- eines del menu --
    EditorCatalegs.ps1  editar els catalegs   VistaWord.ps1     vistes en Word
    PdfSignar.ps1       Word a PDF + signar   ActExtr.ps1       act. extraordinaries
    ControlsPeriodics.ps1 + ControlsCpEmail.ps1                 controls periodics
    EmailTextos.ps1     textos del correu     Configuracio.ps1  rutes d'aquest PC
    rutes\Ruta.ps1      planificador de rutes (proces a part)
    rutes\Coordenades.ps1  mapa per repassar la geolocalitzacio dels establiments

    -- comu --
    UiComuns.ps1        finestres, botons, la banda granat, _AddConfigRow
    Settings.ps1        settings.json d'aquest ordinador
    DriveApi.ps1        client de Google Drive

  Tot va amb DOT-SOURCE al mateix ambit: una funcio d'un modul veu les dels
  altres i les variables d'aqui. Per aixo l'ORDRE de carrega importa nomes per
  als moduls que calculen alguna cosa EN CARREGAR-SE (Activitats.ps1 i
  ActExtr.ps1 en calculen rutes): han d'anar despres del bloc de rutes.

  Flux del programa:
    1. Trias mode i cataleg (Seguiment.ps1, Select-Mode).
    2. Dades de la capcalera (ID GIA, EXP_NUM...), que s'omplen soles amb
       l'Excel d'activitats si el GIA hi es.
    3. Arbre amb les deficiencies del cataleg; en marcar-ne una, el text surt al
       panell de detall i els [OPCIO:]/[CAMP:] s'omplen ALLA MATEIX.
    4. Conclusions del tipus d'informe, tambe amb els camps inline.
    5. Document final: capcalera + items numerats 1..N + conclusions.

.NOTES
  Configuracio: config.ps1 (opcional, al costat d'aquest fitxer) i, per sobre,
  settings.json d'aquest ordinador (pantalla Configuracio).

  Estat a %LOCALAPPDATA%\InformesCornella\ (mai al repositori):
    lastreport.json  dades de l'ULTIM informe generat amb exit, per poder-lo
                     replicar des del Pas 2.
    running.pid      PID de la instancia viva (instancia unica + perque
                     Actualitzar.bat pugui tancar el programa abans d'actualitzar).
    settings.json    rutes d'AQUEST ordinador (Settings.ps1).
  NO hi ha cap "sessio recuperable" pas a pas: si es tanca el programa a mig
  assistent, es torna a comencar (amb l'ultim informe com a plantilla).

  El que es d'aquest ordinador (informes generats, Excel, vistes en Word...) va
  a local\, que el .gitignore exclou sencera. Vegeu local\README.txt.

  Placeholders del cos dels catalegs:
    [CAMP: nom]                 -> demana 'nom'
    [CAMP: nom (hint d'ajuda)]  -> demana 'nom', el hint apareix sota el camp
    [OPCIO: nom | A | B]        -> desplegable
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

    # ELS ERRORS DELS HANDLERS, AMB FITXER I LINIA.
    #
    # Una excepcio dins d'un handler de WinForms (un clic, un CheckedChanged...)
    # NO arriba al try/catch de qui va obrir la finestra: la para el bucle de
    # missatges, que ensenya el seu dialeg gris en l'idioma del Windows i sense
    # dir on ha passat. Amb aixo surt el mateix format que la resta del
    # programa -missatge + fitxer + linia-, que es l'unic que permet arreglar-ho.
    [System.Windows.Forms.Application]::SetUnhandledExceptionMode('CatchException')
    [System.Windows.Forms.Application]::add_ThreadException({
        param($errSender, $errArgs)
        $ex = $errArgs.Exception
        $on = ''
        if ($null -ne $ex.ErrorRecord -and $null -ne $ex.ErrorRecord.InvocationInfo) {
            $ii = $ex.ErrorRecord.InvocationInfo
            $on = "`n`n" + (Split-Path -Leaf ([string]$ii.ScriptName)) + ', linia ' + $ii.ScriptLineNumber
        }
        [System.Windows.Forms.MessageBox]::Show(
            ($ex.Message + $on), 'Error', 'OK', 'Error') | Out-Null
    })
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

# Obre l'eina de COORDENADES (rutes/Coordenades.ps1). Mateixa mecanica que
# Start-RutaTool: '&' l'executa en un AMBIT AILLAT (les seves variables no
# contaminen el generador) pero dins del MATEIX proces, aixi que la finestra
# porta el mateix escut i, en acabar o cancel-lar, es torna al menu.
function Start-CoordenadesTool {
    $eina = Join-Path $ScriptRoot (Join-Path 'rutes' 'Coordenades.ps1')
    if (-not (Test-Path -LiteralPath $eina)) {
        [System.Windows.Forms.MessageBox]::Show("No s'ha trobat Coordenades.ps1.", 'Coordenades', 'OK', 'Error') | Out-Null
        return
    }
    try {
        & $eina
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error a l'eina de coordenades:`n$($_.Exception.Message)", 'Coordenades', 'OK', 'Error') | Out-Null
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

# ESTRUCTURALS viu a l'arrel del clone (al costat dels .bat), no dins de
# suport/. Hi ha NOMES LES FONTS dels catalegs: els .json (que edita l'editor de
# catalegs) i '0 CAPCALERA.docx', l'unica plantilla de Word de veritat. Les
# vistes en Word dels catalegs son derivades i viuen a local\vistes-catalegs\.
$EstructuralsDir = Join-Path $RepoRoot 'ESTRUCTURALS'
$HeaderPath      = Join-Path $EstructuralsDir '0 CAPCALERA.docx'
$ConclusionsPath = Join-Path $EstructuralsDir '0 CONCLUSIONS.json'

# Carpeta 'local\': TOT el que es d'aquest ordinador i no va al repositori.
# Defineix Get-LocalSubdir (les rutes) i Invoke-MigracioLocal (que hi mou el que
# abans estava escampat per l'arrel). Es carrega aqui perque les linies de sota
# ja en necessiten les rutes.
. (Join-Path $ScriptRoot 'Migracio.ps1')

# ----------------------------------------------------------------------------
# Configuracio per defecte. Es pot sobreescriure des de config.ps1 (opcional)
# al costat del .ps1 (dins de suport/).
# ----------------------------------------------------------------------------
# OutputDir per defecte: local\informes-generats\, dins del clone pero fora del
# repositori (local\ s'ignora sencera). Es pot canviar a la Configuracio.
$OutputDir              = Get-LocalSubdir $RepoRoot 'Informes'
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
$Script:DefaultRutesOutputDir = Get-LocalSubdir $RepoRoot 'Rutes'

$Script:AppSettings = Load-AppSettings
$InformesDir      = _ResolveEffectiveValue $AppSettings.InformesDir      $InformesDir
$ActivitatsDir    = _ResolveEffectiveValue $AppSettings.ActivitatsDir    $ActivitatsDir
$OutputDir        = _ResolveEffectiveValue $AppSettings.OutputDir        $OutputDir
$DriveBaseDir     = _ResolveEffectiveValue $AppSettings.DriveBaseDir     $DriveBaseDir
$CopiaInformesDir = _ResolveEffectiveValue $AppSettings.CopiaInformesDir $CopiaInformesDir

# Endrec de la carpeta 'local': si encara hi ha les carpetes velles a l'arrel
# del clone (o vistes .docx a ESTRUCTURALS), s'hi mouen. Va DESPRES de carregar
# el settings.json perque, si l'usuari hi tenia desada una ruta antiga, es pugui
# reescriure. Es idempotent i molt barat: si ja esta fet, nomes son uns quants
# Test-Path. En headless no cal (els tests no volen tocar el disc del clone).
if (-not $Script:HeadlessTest) { [void](Invoke-MigracioLocal $RepoRoot) }

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

# ----------------------------------------------------------------------------
# Els TROSSOS DEL MOTOR (abans, tot dins d'aquest fitxer)
# ----------------------------------------------------------------------------
# Motor.ps1 havia arribat a 3.300 linies i 70 funcions: hi convivien la
# configuracio, l'Excel d'activitats, el Drive, la pantalla de seleccio, els
# camps, les conclusions, la construccio del document, el mode mobil i el menu.
# Ara cada concepte te el seu fitxer i aqui nomes queda el que fa de BASE:
# rutes, carrega de moduls, estat, instancia unica, Word i acces al cataleg.
#
# COM QUE TOT VA AMB DOT-SOURCE AL MATEIX AMBIT, moure una funcio d'un fitxer a
# un altre NO en canvia el comportament. L'unic que compta es l'ORDRE: aquests
# moduls s'han de carregar DESPRES del bloc de rutes de mes amunt, perque algun
# hi calcula una ruta en carregar-se (p. ex. $LocalActivitatsDir a Activitats).
. (Join-Path $ScriptRoot 'Activitats.ps1')      # Excel d'activitats + cache + Drive
. (Join-Path $ScriptRoot 'Capcalera.ps1')       # Pas 2: formulari de la capcalera
. (Join-Path $ScriptRoot 'Camps.ps1')           # [CAMP:]/[OPCIO:] i el text ric
. (Join-Path $ScriptRoot 'SeleccioItems.ps1')   # Pas 3 i Pas 5: les pantalles de tria
. (Join-Path $ScriptRoot 'Document.ps1')        # Pas 6: composicio del .docx
. (Join-Path $ScriptRoot 'Paquet.ps1')          # mode "des de paquet" (mobil)
. (Join-Path $ScriptRoot 'Wizard.ps1')          # Main + assistent de passos

# Editor visual dels ESTRUCTURALS (Editar catalegs). Funcions pures (model<->JSON)
# testejables; la finestra WinForms nomes s'executa a Windows.
. (Join-Path $ScriptRoot 'EditorCatalegs.ps1')

# Generador de les VISTES en Word dels catalegs (des dels JSON). Els .docx
# d'ESTRUCTURALS ja no son plantilles: son vistes per consultar (excepte
# '0 CAPCALERA.docx'). Funcions pures testejables; Word (COM) nomes a Windows.
. (Join-Path $ScriptRoot 'VistaWord.ps1')

# Eina "Convertir informes a PDF (i signar)". Funcions pures (rutes, arguments
# d'AutoFirma) testejables; Word (COM) i AutoFirma nomes s'executen a Windows.
. (Join-Path $ScriptRoot 'PdfCms.ps1')
. (Join-Path $ScriptRoot 'PdfSignar.ps1')

# Eina "Seguiment" (fila GIA): els cinc llistats de seguiment de la base de
# dades d'activitats, en Excel o en PDF. Les funcions de dades son pures i es
# proven en headless; nomes la construccio del llibre fa servir Excel (COM).
. (Join-Path $ScriptRoot 'SeguimentGia.ps1')
. (Join-Path $ScriptRoot 'Llicencia.ps1')

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
    # LLIC.json tampoc: no es un cataleg de deficiencies sino la capa propia de
    # Llicencia (per cada requeriment de REQ1, el "No es disposa", el "Es
    # disposa" i el "Quan:"). Els seus items NO tenen text propi -el treuen de
    # REQ1 en viu-, o sigui que triar-lo aqui donaria un informe buit.
    Get-ChildItem -LiteralPath $EstructuralsDir -Filter '*.json' |
        Where-Object {
            $_.Name -notlike '0 *' -and $_.Name -notlike '0_*' -and
            $_.Name -notlike 'ACT_EXTR*' -and $_.Name -ne 'LLIC.json' -and
            -not $_.Name.StartsWith('~$')
        } |
        Sort-Object Name
}

# NOTA: la tria de cataleg ja no es un pas a part. Ara la pantalla inicial
# (Select-Mode, a Seguiment.ps1) fusiona la tria de MODE i la de CATALEG en un
# sol menu (Pas 1), i passa el cataleg triat directament al wizard
# (Invoke-NouWizard). Get-Catalegs (a dalt) segueix sent la font de catalegs.

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
