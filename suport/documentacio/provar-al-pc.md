# Provar el programa al PC (prompt per a Claude Desktop)

Aquest document té dues parts: **el prompt** que pots enganxar a Claude Desktop
(o a Claude Code al teu ordinador), i **el perquè** de cada comprovació.

Serveix perquè hi ha una part del programa que **no es pot provar en el
contenidor de Linux on es fan els canvis**: tot el que és WinForms (les
finestres) i tot el que va per COM (Word, Excel, Outlook). La suite de 2.253
asserts cobreix la lògica pura; això cobreix la resta.

---

## EL PROMPT (enganxa'l tal qual)

```
Ets al clone de git del programa "Informes Cornellà", al meu PC amb Windows,
Office i accés a la unitat de xarxa I:.

Vull que provis el programa de dalt a baix després d'una tanda de canvis
d'arquitectura. NO has de tocar codi: has de FER SERVIR el programa i dir-me
què falla, amb el missatge d'error exacte i on ha passat.

Regles:
- Si una cosa peta, apunta el text sencer del quadre d'error i segueix amb la
  resta de proves. No t'aturis a la primera.
- Si alguna cosa et demana enviar un correu de debò a un titular, ATURA'T i
  pregunta-m'ho abans.
- No facis "git push" ni "Actualitzar.bat" fins que t'ho digui jo.

Primer:
1. Executa la suite sencera i digues el total:
   GENINFORME_TEST=1 pwsh -NoProfile -File suport/tests/run-tests-all.ps1
   (o amb "powershell" si no tens pwsh). Han de sortir 2.253 asserts i 0
   fallades. Si no, para i digue'm què falla.

Després obre el programa (GenerarInforme.bat) i comprova, en aquest ordre:

A. EL MENÚ PRINCIPAL  [el codi del menú s'ha mogut de fitxer]
   1. S'obre la finestra i es veuen els quatre apartats de rajoles
      (EINES, INFORMES, GIA, MÒBIL) i els botons de tipus d'informe.
   2. Sota cada rajola hi surt la data d'última execució (o "(mai)").
   3. El títol de cada botó no queda tapat pels xips ✏️ / Dades.
   4. El botó 📁 obre la carpeta dels informes i NO tanca el menú.
   5. Els enllaços "Capçalera" i "Conclusions" obren l'editor de catàlegs.

B. ELS TEXTOS DEL CORREU  [les tres pantalles ara són una de sola]
   6. MÒBIL → "Textos del correu": s'obre, es veu l'assumpte i el cos, i la
      línia d'ajuda és A SOBRE del quadre gran.
   7. Prem "Desar" i després obre docs\dades\email-textos.json amb el bloc de
      notes: HA DE SEGUIR TENINT el bloc "bcc" amb les quatre adreces.
      >>> AIXÒ ÉS EL MÉS IMPORTANT DE TOTA LA LLISTA <<<
   8. EINES → "Controls periòdics" → "Editar text": s'obre, prova
      "Restaurar original" i després "Desar".
   9. EINES → "Recordatoris" → "Editar text...": s'obre, es desa, i en
      tornar-hi el text que has desat hi és.

C. L'EXCEL D'ACTIVITATS  [ara el llegeix una sola funció]
   10. EINES → "Generar ruta": genera una ruta i mira que el mapa s'obri amb
       els punts i les pastilles del GIA.
   11. EINES → "Coordenades": genera el mapa, obre'l, arrossega un punt i
       baixa l'Excel; comprova que s'obre bé.
   12. EINES → "Controls periòdics": surt la llista d'activitats.
   13. INFORMES → "Actualitzar base": acaba i diu quants informes ha trobat.
   14. GIA → "Comprovar Excel" i GIA → "Seguiment": funcionen.
   15. Obre el Gestor de tasques i comprova que NO queda cap EXCEL.EXE
       corrent després de tot això.  >>> és el defecte que s'ha arreglat <<<

D. EL WORD  [ara s'obre sempre per la mateixa funció]
   16. "Requeriment - Nou": fes un informe sencer d'una activitat de prova i
       obre'l. Mira la capçalera, la numeració, els enllaços i les conclusions.
   17. INFORMES → "Word a PDF": converteix l'últim informe generat.
   18. Edita un catàleg des del xip ✏️ i desa: en tancar l'editor s'han de
       regenerar les vistes en Word d'ESTRUCTURALS.
   19. MÒBIL → "Revisar mòbil" i "Informe de seguiment" sobre un informe
       anterior: comprova que les anotacions datades surten amb la lletra i
       l'espaiat de sempre (compara'l amb un de fet abans d'aquests canvis).

E. EL CORREU
   20. Genera un requeriment i prova "Enviar correu": comprova que la VISTA
       PRÈVIA es veu bé i que els enllaços són clicables. NO l'enviïs encara.
   21. Prem "Fet" i comença un informe nou: la vista prèvia del correu
       s'ha d'haver BUIDAT (abans es quedava la de l'anterior).
   22. Controls periòdics → "Enviar correu (esborranys)" amb UNA activitat de
       prova: mira l'esborrany a l'Outlook (no l'enviïs). Els enllaços han de
       ser clicables i, si el text porta //cursiva//, ha de sortir en cursiva.

F. EL MÒBIL (al telèfon o al navegador)
   23. Obre la pàgina del mòbil. Al Pas 2 NO hi han de sortir camps de text
       que es diguin ORIGEN, DATES ni CLASSIFICACIO.
   24. El comptador de dalt ha de dir "Pas X / 4" (no "/ 5").
   25. Fes un informe sencer i mira la vista prèvia del correu.

Quan acabis, fes-me un resum: què ha anat bé, què ha fallat (amb l'error
exacte) i què no has pogut provar.
```

---

## Per què cada bloc

| Bloc | Què s'ha canviat i què podria fallar |
|---|---|
| **A** | `Select-Mode` ha passat de `Seguiment.ps1` a `Menu.ps1`. Si alguna cosa hi falla, serà que el menú no s'obre o que una rajola no despatxa. |
| **B** | Les tres pantalles d'«assumpte + cos» són ara una de sola. El punt **7** és el crític: el bloc `bcc` no s'edita des d'aquella pantalla però és del mateix fitxer, i si el desat no el tornés a escriure **es perdrien les quatre adreces de còpia oculta**. |
| **C** | Els set lectors de l'Excel són ara un. El punt **15** comprova el defecte que això va arreglar: tres còpies feien `Close`/`Quit` sense `try` i podien deixar un Excel orfe amb el fitxer agafat. |
| **D** | Les cinc obertures del Word són ara una (`New-WordApp`), que hi afegeix l'`AutomationSecurity` — el que evita la **Vista protegida** amb fitxers de la unitat de xarxa. Si això falla, es veuria com un document que no es deixa modificar. |
| **E** | L'HTML del correu és ara una sola funció, i els URLs s'aparten abans d'aplicar la cursiva (dos enllaços a la mateixa línia es destrossaven). El punt **21** és un defecte que hi havia: el reinici feia `.value` sobre un `<div>`, que no fa res. |
| **F** | El mòbil pintava tres quadres de text que no anaven enlloc, i el comptador deia «/ 5» quan només s'hi arriba a 4. |

## Si alguna cosa falla

Digue-m'ho amb el missatge exacte i el punt de la llista. Tot el que s'ha
tocat en aquesta tanda són **moviments de codi entre fitxers** i **unificacions
de funcions duplicades**: no hi hauria d'haver cap canvi de comportament, i
els 19 fitxers d'or ho confirmen per als documents generats. Si en surt un,
serà una crida que ha quedat apuntant on no toca — i es veurà de seguida.
