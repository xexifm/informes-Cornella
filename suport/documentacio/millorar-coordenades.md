# Prompt: millorar l'eina «Coordenades»

> Per enganxar tal qual en una sessió nova de Claude Code (web) sobre aquest
> repositori. Està escrit l'octubre de 2026, quan l'eina ja funcionava contra
> el Cadastre de debò, i s'ha aplicat una vegada (octubre de 2026): els
> candidats que ja es van fer estan tret d'aquí. Les xifres de sota són d'aquell moment i la sessió les
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

El que ja vaig respondre l'octubre de 2026 (no ho tornis a preguntar si no ha
canviat): el mapa s'obre amb **Chrome**; les correccions les **importa una altra
persona**, amb el **mateix format que la base original** i les canviades en
vermell. Pregunta'm, amb opcions i d'un sol cop:

- **Fins on he arribat amb el repàs** i si ja s'ha importat cap tanda.
- Si l'«Excel per importar» va bé a qui l'importa, o si hi falta res.
- **Què em fa perdre més temps al mapa** ara.
- Si em pots passar una o dues respostes reals del Cadastre:
  `local/geocodificacio/resposta-<refcat>.xml` (les desa `Provar-Cadastre.bat`).
  Només hi ha referències cadastrals i adreces de portals, dades públiques.

## 2. Candidats (per verificar i ordenar per benefici / risc, no és un pla tancat)

**Ja fet l'octubre de 2026** (detalls a `rutes-i-mapes.md`; no ho tornis a
proposar): el mapa és una plantilla a part (`rutes/CoordenadesMapa.html`); la
prova al navegador és repetible (`tests/navegador/`); segon CDN i missatge si el
Leaflet no carrega; filtre per estat, «Següent pendent», «Desfer» i la parella
ressaltada; el repàs es recarrega des de l'Excel baixat; i l'«Excel per
importar» (`rutes/CoordenadesImportar.ps1`).

**Pendent:**

1. **Confirmar al PC l'«Excel per importar»** (punt 11 de `provar-al-pc.md`):
   és l'única peça que escriu amb l'Excel de debò i no s'ha pogut provar aquí.
   Si les coordenades de la base són text, mira que la còpia les conservi com a
   text.
2. **Canviar la fixture muntada a mà** (`tests/dades/wfsAD-exemple.xml`) per una
   resposta **gravada** del servei real, si te la puc passar. Mantén els casos
   difícils que ja prova la fixture (dos portals amb el mateix número, eixos
   girats, adreça sense `<pos>`), afegint-los a part si la resposta real no els
   porta.
3. **Quan arribi una base nova**, que l'eina digui quines correccions del repàs
   ja hi són, quines no i quines han canviat. L'«Excel per importar» ja separa
   les que a la base tenen una altra coordenada (`JaCanviades`); falta que es
   pugui consultar sense generar l'Excel.
4. **Progrés per zona a la finestra de triar zones**: ara que el repàs existeix
   com a fitxer, la finestra el podria llegir (`Read-RepasXlsx`) i dir quantes
   en portes de cada zona. Pregunta'm abans si val la pena.
5. **Apilades amb tolerància?** Avui «apilada» vol dir mateixa coordenada
   exacta. Mesura sobre una base real si hi ha activitats a menys d'1 m que no
   es detecten. Si no n'hi ha, no ho canviïs.
6. **Precisió.** Mira el resum del final d'una execució real (exacte / dubtós /
   més proper / sense portal) i ataca el grup més gran. Per als **blancs** (cap
   portal) es pot mirar si la parcel·la té portals a la parcel·la veïna o al
   mateix carrer. Fer servir un altre servei (ICGC, Cartociudad...) voldria dir
   enviar-hi **l'adreça**, no només la referència cadastral: això canvia la
   promesa de privacitat del `LLEGEIX-ME.md` i **m'ho has de preguntar abans**.

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
- Si toques `rutes/CoordenadesMapa.html`, executa també la suite del navegador:
  `cd suport/tests/navegador && npm install && node prova-mapa-coordenades.mjs`.
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
