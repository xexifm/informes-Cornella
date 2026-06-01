BASE DE DADES D'ACTIVITATS (fallback local)
============================================

Aquesta carpeta serveix per quan executes el programa FORA de la feina
(sense acces a la unitat de xarxa I:).

COM FER-HO SERVIR
-----------------
1. Copia el fitxer "YYYY-MM-DD ACTIVITATS.xls" o ".xlsx" mes recent a
   aquesta carpeta.
2. Executa GenerarInforme.bat normalment.

El programa primer prova la ruta de la feina; si no hi ha acces, fa
servir el fitxer mes recent d'aquesta carpeta i ho indica al Pas 2 amb
l'etiqueta [FALLBACK LOCAL] en taronja.

NOTA
----
Els fitxers .xls/.xlsx d'aquesta carpeta NO es pugen a GitHub
(estan al .gitignore). Es queden nomes al teu PC.
