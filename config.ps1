# Configuracio local del generador d'informes.
#
# Aquest fitxer es OPCIONAL. Si existeix al costat de GenerarInforme.ps1,
# el script l'ha de carregar i fa servir aquestes variables en lloc dels
# valors per defecte. Si no existeix, el script funciona amb els valors
# per defecte definits a GenerarInforme.ps1.
#
# Si vols personalitzar una ruta nomes en aquest equip, posa el valor
# aqui. Aquest fitxer pot quedar fora del control de versions (es a dir,
# no es comparteix entre maquines).

# Ruta on s'han de desar els informes generats. Si no es accessible, el
# script cau automaticament a la subcarpeta 'Informes generats' al costat
# del .ps1.
$OutputDir = 'I:\Activitats_Ordenances\Activitats\5.- Sergi Fadurdo\0_Plantilles\Powershell\Informes generats'

# Directori que conte la base de dades d'activitats en format Excel. El
# fitxer ha de seguir el nom "YYYY-MM-DD ACTIVITATS.xls" o ".xlsx".
$ActivitatsDir = 'I:\Activitats_Ordenances\Activitats\5.- Sergi Fadurdo\2_Controls Excels'

# Quantes conclusions del final del fitxer 0 CONCLUSIONS.docx s'inclouen
# sempre al document final (no apareixen al Pas 5 per ser triades).
$AlwaysConclusionsCount = 2
