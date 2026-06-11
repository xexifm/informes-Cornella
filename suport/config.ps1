# Configuracio local del generador d'informes.
#
# Aquest fitxer es OPCIONAL. Si existeix AL COSTAT de GenerarInforme.ps1
# (dins de la carpeta suport/), el script l'ha de carregar i fa servir
# aquestes variables en lloc dels valors per defecte.
#
# Si vols personalitzar una ruta nomes en aquest equip, descomenta i posa
# el valor aqui. Aquest fitxer no es comparteix entre maquines.

# Ruta on s'han de desar els informes generats. Per defecte el script els
# desa a la subcarpeta 'Informes generats' a l'arrel del clone (al costat
# dels .bat). Si vols una ruta absoluta diferent, descomenta:
#
# $OutputDir = 'D:\Informes\Sortida'

# Directori que conte la base de dades d'activitats en format Excel. El
# fitxer ha de seguir el nom "YYYY-MM-DD ACTIVITATS.xls" o ".xlsx".
$ActivitatsDir = 'I:\Activitats_Ordenances\Activitats\5.- Sergi Fadurdo\2_Controls Excels'

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
# PDF). Per defecte 'Rutes generades' a l'arrel del clone (ignorada per git).
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

# Mobil (Google Drive). Carpeta sincronitzada al PC amb el Google Drive
# d'escriptori, on:
#   - el vigilant (Vigilant.bat) llegeix els paquets que arriben del mobil, i
#   - s'hi exporta la base de dades d'activitats per al mobil (carpeta PRIVADA).
# El default apunta a "%USERPROFILE%\Google Drive\Informes-Cornella". Si el teu
# Drive d'escriptori esta en una altra ruta (per exemple amb la lletra G:),
# descomenta i ajusta. Es creen soles les subcarpetes Entrada/Processats/Dades.
#
# $DriveBaseDir = 'G:\El meu Drive\Informes-Cornella'

# Mode mobil SENSE Google Drive d'escriptori (accés a Drive per API).
# IDs de les carpetes de Drive (es treuen de la URL: .../folders/<AIXÒ>). Les
# credencials (secretes) NO van aquí: les desa Authorize-Drive.ps1 a %LOCALAPPDATA%.
$DriveEntradaId    = '1EZteYidosVE8iJeli4tNdRg6WxzOk5fn'   # carpeta on arriben els paquets del mòbil
$DriveProcessatsId = '131iyeU0B9eWgDA8E6ncxixNFwx5Yz68X'   # carpeta on es mouen un cop generats
$DriveDadesId      = '1bRof2NwgLYKpzvUD0aGximQzZQkv73od'   # carpeta on es puja activitats.json

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
