CARPETA "local" - tot el que es d'AQUEST ordinador
==================================================

Aqui hi ha el que NO va al GitHub: el que generes, el que baixes i les teves
bases de dades. Tot aquest contingut es queda nomes al teu PC.

Aquest fitxer (README.txt) es l'unic de "local" que si que es puja: hi es
perque la carpeta existeixi al repositori i per explicar-te que hi ha.

QUE HI TROBARAS
---------------
informes-generats\      Els informes .docx que fa el programa (i els PDF que
                        en surten amb l'eina "Word a PDF"). Ho pots canviar
                        a la pantalla de Configuracio.

rutes-generades\        Els mapes de ruta (.html) del planificador de rutes.
                        Els obres al navegador i els imprimeixes a PDF.

base-dades-activitats\  La teva copia de l'Excel d'activitats, per treballar
                        FORA de la feina (sense la unitat de xarxa I:), i la
                        base d'informes (informes-db.json).
                        Per fer-la servir: copia-hi el fitxer
                        "AAAA-MM-DD ACTIVITATS.xls" (o .xlsx) mes recent. El
                        programa prova primer la ruta de la feina i, si no hi
                        te acces, agafa el mes recent d'aqui i t'ho indica al
                        Pas 2 amb l'etiqueta [FALLBACK LOCAL] en taronja.
                        Aqui tambe hi ha el registre de l'eina "Word a PDF"
                        (pdf-signar-log.txt i pdf-signar-state.json).

base-dades-actextr\     El registre de les activitats extraordinaries
                        (Decret 112/2010) que portes amb el mode ACT_EXTR:
                        activitats-extraordinaries.json.
                        Si vols tenir el mateix registre a la feina i a casa,
                        copia aquest fitxer a ma entre els dos PCs.

vistes-catalegs\        Copies en Word dels catalegs, per poder-los llegir
                        sencers sense obrir el programa (REQ1.docx,
                        TERMINI.docx...). Es REGENEREN soles cada cop que
                        deses des de l'editor de catalegs o fas
                        Actualitzar.bat, aixi que no les editis: els canvis
                        es perdrien. Per canviar un cataleg, fes servir
                        l'editor de catalegs del programa (edita els .json
                        d'ESTRUCTURALS, que son la font de veritat).

seguiment-gia\          Els llistats de seguiment del GIA que fa l'eina
                        "Seguiment" (fila GIA del menu), en Excel i en PDF:
                        PRECINTES, DENUNCIES, REQUERIT DECRET, SONOMETRIA i
                        ANNEX II, amb una copia de la fulla Estes.

PRIVACITAT
----------
Bona part d'aquests fitxers contenen DADES PERSONALS (noms i adreces de
titulars) i el repositori de GitHub es PUBLIC. Per aixo tota la carpeta "local"
esta al .gitignore: res del que hi hagi a dins es pot pujar per accident.

I QUE NO HI HA?
---------------
La configuracio del programa i els fitxers d'estat no son aqui, sino a

    %LOCALAPPDATA%\InformesCornella\

(settings.json, lastreport.json, les copies de seguretat dels catalegs, les
credencials del Google Drive...). Son d'usuari de Windows, no "documents": aixi
sobreviuen encara que tornis a baixar el programa de zero, i les credencials del
Drive no queden mai dins d'una carpeta que puguis comprimir i enviar a algu.
Per obrir-la, enganxa aquesta ruta a la barra de l'Explorador de Windows.
