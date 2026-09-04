# Configuracio local del generador d'informes.
#
# Aquest fitxer es OPCIONAL. Si existeix AL COSTAT de GenerarInforme.ps1
# (dins de la carpeta suport/), el script l'ha de carregar i fa servir
# aquestes variables en lloc dels valors per defecte.
#
# Si vols personalitzar una ruta nomes en aquest equip, descomenta i posa
# el valor aqui. Aquest fitxer no es comparteix entre maquines.

# Ruta on s'han de desar els informes generats. Per defecte el script els
# desa a local\informes-generats\ (dins del clone pero fora del repositori).
# Si vols una ruta absoluta diferent, descomenta:
#
# $OutputDir = 'D:\Informes\Sortida'

# Directori que conte la base de dades d'activitats en format Excel. El
# fitxer ha de seguir el nom "YYYY-MM-DD ACTIVITATS.xls" o ".xlsx".
$ActivitatsDir = 'I:\Activitats_Ordenances\Activitats\5.- Sergi Fadurdo\2_Controls Excels'

# Carpeta ARREL dels informes ja generats. La fa servir el boto "Actualitzar
# base d'informes" del menu, que la recorre i escriu
# local\base-dades-activitats\informes-db.json (ID GIA + data + conclusio per
# informe, agrupat per activitat). Per defecte es la germana de $ActivitatsDir:
# ...\5.- Sergi Fadurdo\Informes. Descomenta si la tens en una altra ubicacio:
#
# $InformesDir = 'I:\Activitats_Ordenances\Activitats\5.- Sergi Fadurdo\Informes'

# Carpeta on l'eina "Copiar informes" (menu INFORMES) fa una copia de seguretat
# PLANA (tots els Word a una sola carpeta) dels informes. Sense valor per
# defecte: normalment es configura a la pantalla Configuracio (per ordinador),
# pero tambe es pot fixar aqui. Descomenta:
#
# $CopiaInformesDir = 'D:\Copia Informes'

# --- Planificador de rutes (Ruta.bat) -------------------------------------
# Servidor de rutes OSRM que calcula la ruta circular mes rapida. Per defecte
# s'usa el servidor public de demostracio. NOMES s'hi envien coordenades
# (mai noms ni adreces). Si tens un OSRM propi (per privacitat o per anar
# sense internet), posa la seva URL base aqui. Si el deixes buit o no hi ha
# xarxa, el programa calcula una ruta aproximada en linia recta.
#
# $OsrmBaseUrl = 'http://localhost:5000'
#
# Carpeta on es desen els mapes de ruta generats (HTML que pots imprimir a
# PDF). Per defecte local\rutes-generades\ (dins del clone, fora del repositori).
#
# $RutesOutputDir = 'D:\Rutes'
#
# Base de sortida: la ruta comenca SEMPRE per l'activitat mes propera a aquest
# punt (i hi torna al final). Per defecte Carrer de l'Energia, 97 (Cornella),
# en coordenades UTM (mateix sistema que la base de dades). Per canviar la
# base, descomenta i posa unes altres coordenades. Posa-les a 0 si vols que la
# ruta comenci per la primera activitat que escriguis a la llista.
#
# $RutaOrigenUtmX = 424456
# $RutaOrigenUtmY = 4578205
#
# Etiqueta de la base (apareix a la "Parada 0" del mapa quan surts d'allà):
# $RutaOrigenLabel = "Carrer de l'Energia, 97"
#
# Per defecte la casella "Sortir des de la BASE i tornar-hi" del formulari
# surt MARCADA. Posa $false aquí si vols que surti desmarcada per defecte
# (la ruta començarà per l'activitat més propera a la base, sense passar
# expressament per la base).
# $RutaSortirDesDeBaseDefault = $false

# Mobil (Google Drive). Carpeta sincronitzada al PC amb el Google Drive
# d'escriptori, on:
#   - el botó "Revisar entrades del mòbil" del menú llegeix els paquets que
#     arriben del mòbil (comprovació d'un sol cop), i
#   - s'hi exporta la base de dades d'activitats per al mobil (carpeta PRIVADA).
# El default apunta a "%USERPROFILE%\Google Drive\Informes-Cornella". Si el teu
# Drive d'escriptori esta en una altra ruta (per exemple amb la lletra G:),
# descomenta i ajusta. Es creen soles les subcarpetes Entrada/Processats/Dades.
#
# $DriveBaseDir = 'G:\El meu Drive\Informes-Cornella'

# Mode mobil SENSE Google Drive d'escriptori (accés a Drive per API).
#
# ELS IDs DE LES CARPETES DE DRIVE JA NO SÓN AQUÍ: viuen a docs/config.js
# (DRIVE_ENTRADA_FOLDER_ID, DRIVE_PROCESSATS_FOLDER_ID, DRIVE_DADES_FOLDER_ID),
# que és on el navegador ja els necessita. Abans hi eren als dos llocs i el
# manual demanava escriure'ls dues vegades. Si algun dia cal sobreescriure'n un
# només en aquest ordinador, es pot tornar a assignar aquí sota (aquest fitxer
# es carrega DESPRÉS de llegir config.js).
#
# Les credencials (secretes) NO van aquí: les desa Authorize-Drive.ps1 a
# %LOCALAPPDATA%.

# (Obsolet pero conservat per compatibilitat amb el codi vell.) Quantes
# conclusions del final del fitxer 0 CONCLUSIONS.docx s'inclouen sempre
# al document final. Avui els paragrafs fixos es marquen amb '::SEMPRE::'.
$AlwaysConclusionsCount = 2

# Mode "Informe de seguiment": frases que marquen on comenca el bloc de
# conclusions de l'informe anterior (es a dir, on s'ha de tallar i esborrar).
# La comparacio es insensible a accents/majuscules. Descomenta per
# personalitzar-ho si els teus informes antics fan servir un altre tancament:
#
# $SeguimentConclusionPhrases = @(
#     "Vist l'anterior",
#     'Ho poso al seu coneixement',
#     'Cornella de Llobregat,'
# )
