# Prompt: revisió i simplificació de l'arquitectura

> Per enganxar tal qual en una sessió nova de Claude Code (web) sobre aquest
> repositori. Es pot tornar a fer servir cada cop que el programa hagi crescut:
> les xifres de sota són de setembre 2026 i la sessió les ha de tornar a mesurar.

---

Revisa l'arquitectura del programa d'informes (PowerShell 5.1 + WinForms + Word
COM, amb una web mòbil a `docs/`) i simplifica-la tant com sigui possible, perquè
les modificacions futures no el facin més pesat i reaprofitin el que ja hi ha.
Respon-me sempre en català.

## 0. Abans de tocar res

1. Llegeix `suport/CLAUDE.md` sencer, i els documents de `suport/documentacio/`
   dels mòduls que pensis tocar. Hi ha regles d'arquitectura amb guards a la
   suite i trampes del PowerShell que ja han costat rondes senceres (la coma que
   lliga més que el `+`, `-like` amb claudàtors, les closures,
   `continue`/`break` dins de `switch`/`ForEach-Object`, el `$()` que desenrotlla).
2. Instal·la el `pwsh` si no hi és (el `tar.gz` de GitHub) i executa la suite:
   `GENINFORME_TEST=1 pwsh -NoProfile -File suport/tests/run-tests-all.ps1`.
   Ha d'estar en verd **abans** de començar; si no ho està, atura't i digues-m'ho.
3. Fes una foto del punt de partida: la llista de noms de funció de tot
   `suport/` (via AST), el nombre de línies per fitxer i els 19 fitxers d'or.
   Són les tres coses que després demostren que no s'ha trencat res.

## 1. Inventari mesurat (no deduït)

Fes-ho amb eines, no a ull, i posa les xifres a l'informe:

- **Mida i responsabilitats per fitxer.** La regla és «cada fitxer, una cosa»: si
  per dir què fa un fitxer necessites la paraula «i», és candidat a partir-se.
  Punts de partida de setembre 2026: `Llicencia.ps1` (~2.550 línies: dades
  pures + composició en blocs + cinc pantalles + l'assistent), `Informes.ps1`
  (~1.800: base d'informes, copiar informes, comprovar l'Excel i exportar),
  `PdfSignar.ps1`, `ActExtr.ps1`, `EditorCatalegs.ps1`, `UiComuns.ps1`.
- **Codi mort**: funcions que ningú no crida. Recorre l'AST de tot `suport/`,
  **inclosos** els processos a part (`rutes/Ruta.ps1`, `rutes/Coordenades.ps1`,
  `mobil/*.ps1`, `RecordatorisAuto.ps1`, `CopiaInformesAuto.ps1`), els `.bat`,
  el `.vbs`, les proves i els noms que es construeixen en una cadena (accions del
  menú, `& $nom`). Una funció que sembla morta i la crida un altre procés ja va
  enganyar un cop: comprova-ho abans d'esborrar-la.
- **Duplicació**: busca blocs de codi quasi iguals (normalitza els espais i
  compara finestres de 6–10 línies) i funcions que fan el mateix amb noms
  diferents. Candidats coneguts per verificar:
  - els convertidors «PSCustomObject del JSON → hashtable» (`_LlicDbAMapa`,
    `_RecHistorialAMapa`, `_RecObjAMapa` i els que hi hagi);
  - el peu de les pantalles («← Enrere» / «Continuar», amb les mateixes mides i
    estils, repetit a gairebé totes les finestres) i les pantalles de llista +
    detall (`Select-LlicDocumentacio`, `Show-LlicenciaDb`, `Select-Items`);
  - les màquines de passos dels assistents (`Invoke-NouWizard`,
    `Invoke-LlicenciaWizard`: `while` + `switch ($step)` + Enrere/Endavant);
  - les vistes en Word d'ACT_EXTR, MNS i conclusions, que encara escriuen amb
    els embolcalls `_V*` en lloc de blocs + `Write-Informe -AmbNivells` (com ja fa
    la de LLIC);
  - la lògica del format del catàleg (`[CAMP:]`, `[OPCIO:]`, `**`, `//`,
    `[[URL]]`) escrita dues vegades: al PC (PowerShell) i al mòbil (`docs/app.js`).
    Aquí no es pot fer una sola implementació; sí que es pot fer un **joc de
    proves compartit** (entrades i sortides esperades en un JSON que llegeixin
    les dues bandes).
- **Dependències**: qui crida qui entre fitxers. Un mòdul genèric no pot
  dependre d'un client seu (regla 3 del `CLAUDE.md`); llista els cicles.
- **La documentació també pesa**: `suport/CLAUDE.md` torna a tenir més de 2.000
  línies (es va partir quan en tenia 2.750 perquè ningú no el llegia sencer).
  Proposa què es queda (regles vigents, com provar, mapa de mòduls) i què passa
  als documents de `documentacio/` (històries de diagnòstic ja tancades).

## 2. Informe de propostes, abans de fer res gros

Escriu `suport/documentacio/arquitectura.md` amb:

1. el mapa actual (mòduls, què fa cadascun, qui depèn de qui);
2. la llista de candidats **ordenada per benefici / risc**, amb les xifres;
3. per a cadascun: què es guanya, què pot trencar, com es demostra que no ha
   canviat res i si el comportament o l'aspecte canvien per a l'usuari;
4. el que **no** s'ha de fer, i per què (una decisió de no unificar també
   s'escriu: el genèric de vuit paràmetres per cobrir tres casos és pitjor que
   les tres còpies).

Després **executa directament** els canvis que siguin mecànics i demostrables
(moure funcions entre fitxers, esborrar codi mort comprovat, unificar còpies
idèntiques, passar una vista a blocs, partir un fitxer). **Pregunta'm abans**
qualsevol cosa que canviï el que es veu o el que fa el programa, o que toqui la
signatura de PDF, `Actualitzar.bat` o la sincronització de catàlegs.

## 3. Com s'executa cada canvi

- **Un canvi per commit**, i abans de cada commit:
  - la suite sencera en verd;
  - els **19 fitxers d'or idèntics**. Si n'ha de canviar algun a posta, explica
    per què i comprova per programa que la diferència és només la que esperaves;
  - la llista de noms de funció idèntica, tret de les que s'han esborrat o
    reanomenat a posta (i ningú més les ha de cridar).
- Moure funcions entre fitxers és neutre (tot va amb dot-source al mateix àmbit),
  però l'**ordre de càrrega** de `Motor.ps1` importa per als mòduls que calculen
  coses en carregar-se. Els fitxers que carrega un procés a part (`Json.ps1`,
  `Excel.ps1`, `UiFinestra.ps1`, `Docx.ps1`) **només poden definir funcions**.
- Tot `.ps1` amb BOM i sense caràcters no ASCII fora dels comentaris (els
  accents del codi van amb `[char]0x00E0`...).
- Un guard nou s'ha de **validar injectant el defecte** i veient-lo en vermell.
- Els comentaris diuen **per què**, quin defecte hi havia i què es va provar i
  descartar; mai no tradueixen el codi a paraules.
- Al final de cada canvi, `git push` a la branca de la sessió **i a `main`**
  (és d'on actualitza l'usuari amb `Actualitzar.bat`).
- El codi que només corre a Windows (WinForms, Word, Excel, AutoFirma) no el
  toca cap prova: digues **exactament** què he de comprovar jo al PC.

## 4. Perquè no torni a créixer

Proposa i, si és barat, implementa **guards** i **plantilles** perquè el codi
nou vagi al lloc bo sense haver-hi de pensar:

- un guard de **mida màxima per fitxer** (amb la llista d'excepcions justificades)
  i un que no deixi aparèixer **funcions amb el mateix nom** en dos fitxers
  (el darrer carregat guanya en silenci);
- una secció curta al `CLAUDE.md`, **«On va cada cosa nova»**: un informe nou =
  un `Build-<Familia>Blocs` pur + `Write-Informe`; una pantalla nova = els helpers
  de `UiComuns.ps1`; una dada nova d'una activitat = la memòria de la base; un
  catàleg nou = el format estàndard de `CatalegJson.ps1`; etc.;
- si surten helpers de pantalla comuns (peu de botons, llista + detall, màquina
  de passos), fes que els assistents i les pantalles existents els facin servir,
  o no serviran de res.

## 5. Què em lliures

Un resum curt, en català: què has canviat (amb els commits), què has deixat
proposat i per què, els números d'abans i després (línies, funcions, fitxers,
proves) i la llista del que he de comprovar al PC.
