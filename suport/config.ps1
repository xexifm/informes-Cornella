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

# (Obsolet pero conservat per compatibilitat amb el codi vell.) Quantes
# conclusions del final del fitxer 0 CONCLUSIONS.docx s'inclouen sempre
# al document final. Avui els paragrafs fixos es marquen amb '::SEMPRE::'.
$AlwaysConclusionsCount = 2
