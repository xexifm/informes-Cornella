# Prompt: millorar l'eina «Coordenades»

> Per enganxar tal qual en una sessió nova de Claude Code (web) sobre aquest
> repositori. Està escrit l'octubre de 2026, quan l'eina ja funcionava contra
> el Cadastre de debò. Les xifres de sota són d'aquell moment i la sessió les
> ha de tornar a mesurar. Quan una millora quedi feta o descartada, apunta-ho a
> `rutes-i-mapes.md` i treu-la d'aquí, perquè el prompt no proposi dues vegades
> el mateix.

---

Millora l'eina «Coordenades» del programa d'informes (`suport/rutes/Coordenades.ps1`
+ `suport/rutes/Geocodificador.ps1`): l'eina que posa sobre un mapa la
coordenada de l'Excel d'activitats (la parcel·la del Cadastre) i la del portal de
façana, i deixa validar-la o arrossegar-la i baixar-se un `.xlsx` amb les
correccions. Vull que sigui més fiable, que no es perdi mai la feina feta i que el
repàs sigui més ràpid. Respon-me sempre en català.

## 0. Abans de tocar res

1. Llegeix **sencers** `suport/documentacio/rutes-i-mapes.md` (secció
   «Eina Coordenades») i la secció «Coordenades dels establiments» de
   `LLEGEIX-ME.md`. Després, de `suport/CLAUDE.md`, el que val per a tot el
   projecte: com executar les proves, els guards, el BOM, i les trampes del
   PowerShell. N'hi ha tres que ja han mossegat **aquesta mateixa eina**:
   - el `return ,@(...)` combinat amb `@()` a la crida (`$portals.Count` valia 1);
     la convenció d'aquests fitxers és **array pla + `@()` al lloc de la crida**;
   - el guard del JSON d'un sol element: es mira `StartsWith('[')`, no el `Count`;
   - la sigla de via que faltava (`CL`): no falla, **empitjora en silenci**.
2. Instal·la el `pwsh` si no hi és i executa la suite sencera:
   `GENINFORME_TEST=1 pwsh -NoProfile -File suport/tests/run-tests-all.ps1`.
   Ha d'estar en verd **abans** de començar; si no ho està, atura't i digues-m'ho.
3. Mesura el punt de partida: línies de `Coordenades.ps1` (1.458 a l'octubre de
   2026) contra el seu sostre a `tests/proves/06-guards.ps1` (**1.550**), quantes
   d'aquestes línies són l'HTML/JS del here-string de `Build-CoordenadesHtml`, i
   la cobertura de `tests/run-tests-coordenades.ps1`. **No apugis el sostre**:
   si una millora no hi cap, primer es parteix el fitxer (vegeu A1).

## 1. Pregunta'm abans de prioritzar

Fes-me aquestes preguntes (amb opcions, d'un sol cop) abans de proposar res.
La resposta canvia l'ordre de tot:

- **Fins on he arribat amb el repàs?** Quantes zones, quantes validades, i si ja
  he abocat cap tanda de correccions a la base.
- **Què faig amb el `.xlsx` que em baixo?** Si les correccions es posen a mà al
  GIA, una a una, o si hi ha cap manera d'importar-les. D'això depèn el format
  que ha de tenir l'Excel de sortida (D1).
- **Quin navegador obre el mapa a la feina** (Edge, Chrome, Firefox) i si
  `unpkg.com` hi és accessible sempre.
- **Què em fa perdre més temps al mapa**, de la llista de C, o una altra cosa.
- Si em pots passar una o dues respostes reals del Cadastre:
  `local/geocodificacio/resposta-<refcat>.xml` (les desa `Provar-Cadastre.bat`).
  Només hi ha referències cadastrals i adreces de portals, dades públiques.

## 2. Candidats (per verificar i ordenar per benefici / risc, no és un pla tancat)

### A. Fiabilitat i deute tècnic

1. **Treure l'HTML/JS del here-string** cap a un fitxer de plantilla al costat
   (per exemple `rutes/CoordenadesMapa.html`, o `.html` + `.js`), on el
   PowerShell només injecta les dades en JSON en un sol punt. Guanys: desapareix
   la trampa del `$` i del `` ` `` dins de `@"…"@`, el text català deixa de
   dependre del BOM del `.ps1`, s'alliberen centenars de línies sota el sostre i
   el JS es pot provar directament. Compte amb: llegir la plantilla amb
   `[IO.File]::ReadAllText(..., UTF8)` explícit (el 5.1 llegeix ANSI per
   defecte), que la injecció escapi `</script>` i que el resultat sigui
   **byte a byte** el mateix HTML que abans (compara'l per programa amb el que
   genera la versió actual sobre les mateixes dades de prova).
2. **Canviar la fixture muntada a mà** (`tests/dades/wfsAD-exemple.xml`) per una
   resposta **gravada** del servei real, si te la puc passar. Mantén els casos
   difícils que ja prova la fixture (dos portals amb el mateix número, eixos
   girats, adreça sense `<pos>`), afegint-los a part si la resposta real no els
   porta.
3. **Proves del mapa repetibles.** El mapa sencer es va provar un sol cop amb
   Chromium + Playwright i un doble de Leaflet. Converteix-ho en un script de
   prova que es pugui tornar a executar (estat inicial, arrossegar, validar,
   desfer, desat i recuperació, filtre, `.xlsx` rellegit amb `openpyxl`). El SRI
   de Leaflet bloqueja el doble: treu l'`integrity` **només** a la còpia de prova.
   Si no pot anar dins de `run-tests-all.ps1` (no hi ha Node a Windows), que
   quedi com a suite a part i digues-ho al `CLAUDE.md`.
4. **Leaflet des d'un CDN.** Si a la feina `unpkg` falla, el mapa surt en
   blanc i la feina no es pot fer. Valora un CDN de reserva o una còpia local a
   `suport/`, i que el mapa digui **en clar** que no ha carregat en lloc de
   quedar-se buit.

### B. Que no es perdi mai el repàs

El repàs (hores de feina) viu **només** al `localStorage` del navegador, amb
clau = nom de l'Excel d'origen, i cada execució genera un HTML amb un nom nou
(`Coordenades_<data>.html`).

1. **Comprova, no suposis,** que el navegador de la feina comparteix el
   `localStorage` entre dos fitxers `file://` diferents. Si en algun navegador
   no ho fa, cada mapa nou començaria buit i no ho veuria ningú.
2. **Una còpia de seguretat que surti del navegador**: per exemple que el mapa
   pugui tornar a **carregar** un `.xlsx` (o un `.json`) baixat abans i en
   recuperi les correccions. Esborrar dades de navegació o canviar d'ordinador
   no hauria de costar la feina. Pregunta'm abans el format.
3. Si això existeix, la finestra de tria de zones **podria** dir quantes en porto
   de repassades llegint aquell fitxer del disc. Sense una via real de retorn del
   navegador al disc, **no** ho intentis (ja està descartat a `rutes-i-mapes.md`).

### C. Repassar més de pressa (només el que jo triï a la pregunta de l'apartat 1)

- Filtrar la llista per estat: pendents, blancs (sense portal), grocs (dubtosos),
  verds clars (aproximats), moguts, validats. I el recompte de cada color a la
  llegenda.
- Un botó / tecla **«Següent pendent»** que centri el mapa i obri el popup.
- Desfer **només l'últim arrossegament** sense perdre la validació de la resta
  (avui, desfer un punt mogut el torna al portal del Cadastre).
- Ressaltar la parella vermell–verd seleccionada quan n'hi ha moltes al mateix
  lloc.

### D. Tancar el cercle

1. **L'Excel de sortida, al format que necessito** per posar les correccions al
   GIA (columnes, ordre, decimals, separador), segons la resposta de l'apartat 1.
2. **Quan arribi una base nova**: que l'eina pugui comparar el darrer `.xlsx`
   de correccions amb la base nova i dir quines ja hi són, quines no i quines han
   canviat. Sense tocar res: l'eina **només mira i genera fitxers**.
3. **Apilades amb tolerància?** Avui «apilada» vol dir mateixa coordenada
   exacta. Mesura sobre una base real si hi ha activitats a menys d'1 m que no es
   detecten. Si no n'hi ha, no ho canviïs.

### E. Precisió

Mira el resum del final d'una execució real (exacte / dubtós / més proper /
sense portal) i ataca el grup més gran. Per als **blancs** (cap portal) es pot
mirar si la parcel·la té portals a la parcel·la veïna o al mateix carrer. Fer
servir un altre servei (ICGC, Cartociudad...) voldria dir enviar-hi **l'adreça**,
no només la referència cadastral: això canvia la promesa de privacitat del
`LLEGEIX-ME.md` i **m'ho has de preguntar abans**.

## 3. Regles de la casa per a aquesta eina

- **Primer un informe curt** (a `rutes-i-mapes.md`, secció nova «Millores
  pendents») amb els candidats ordenats per benefici / risc, què pot trencar
  cadascun i com es demostra que no ha trencat res. Executa directament el que
  sigui mecànic i demostrable (A1, A2, A3); **pregunta'm abans** qualsevol cosa
  que canviï el que veig o el que faig al mapa.
- **No canviïs el comportament de `Ruta.ps1` ni de `Precintades.ps1`**: segueixen
  amb la coordenada original de l'Excel (petició explícita meva). Si mous una
  funció compartida, executa també `run-tests-ruta.ps1` i
  `run-tests-precintades.ps1`.
- `Coordenades.ps1` corre en un **procés a part** i no pot carregar
  `UiComuns.ps1`. Carrega `Geocodificador.ps1` **abans** de `Ruta.ps1` perquè
  `config.ps1` pugui sobreescriure les variables del geocodificador: no ho
  canviïs d'ordre.
- `Coordenades.ps1` **porta BOM**; `Geocodificador.ps1` és **ASCII pur** sense BOM.
  Mentre l'HTML visqui al here-string, cap `$` ni `` ` `` que no sigui una
  interpolació volguda.
- Res del geocodificador llança mai: si el servei no respon, la parcel·la es
  queda sense portals i l'activitat amb la coordenada de sempre. La guarda dels
  250 m (`$GeoDistanciaMaximaM`) es queda.
- **Mai `. Ruta.ps1` a pèl** per diagnosticar (obre el planificador). Per provar
  el Cadastre, `Provar-Cadastre.bat`. Una comanda de diagnòstic amb cometes
  niuades és un `.bat` que falta.
- **Un canvi per commit**, amb la suite sencera en verd i els fitxers d'or
  idèntics. Una prova nova s'ha de veure en **vermell** injectant el defecte
  abans de donar-la per bona. Una rèplica en Python valida la lògica, mai les
  trampes del PowerShell: les suites s'executen amb `pwsh` de debò.
- Els comentaris diuen **per què**, quin defecte hi havia i què es va descartar.
- Actualitza `LLEGEIX-ME.md` (el que veig jo), `rutes-i-mapes.md` (el que ha de
  saber la propera sessió) i el punt 11 de `provar-al-pc.md` (què s'ha de
  comprovar al PC).
- Al final de cada canvi, `git push` a la branca de la sessió **i a `main`** (és
  d'on actualitzo amb `Actualitzar.bat`).

## 4. Què m'has de donar al final

1. Què ha canviat, en llenguatge d'usuari (què veuré diferent al mapa i a la
   finestra).
2. **Exactament** què he de comprovar jo al PC, pas a pas, per a tot el que no
   toca cap prova (WinForms, Excel per COM, la crida real al Cadastre, el
   navegador de la feina).
3. El que queda pendent o descartat, i per què, ja apuntat a `rutes-i-mapes.md`.
