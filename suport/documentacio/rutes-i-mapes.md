# Rutes, coordenades i el plànol públic

> Ve de `suport/CLAUDE.md`, que s'havia fet massa gros per llegir-lo
> sencer. **Llegeix-lo ABANS de tocar `suport/rutes/`**
>
> Per millorar l'eina «Coordenades» hi ha un prompt per enganxar a
> `millorar-coordenades.md`.

## EL FONS DELS MAPES: res d'OpenStreetMap (`rutes/MapaFons.js`, octubre 2026)

L'usuari va obrir el Plànol activitats i **cada quadre del fons deia «Access
blocked» / «403»**. Els servidors de rajoles d'OpenStreetMap exigeixen que la
petició digui de quina web ve (capçalera *Referer*), i un HTML obert des del
disc (`file://`) **no n'envia cap**. Les dades i les parcel·les estaven bé:
només queia el fons. Passava igual a **Coordenades** i a **Ruta**, que tenien
cadascun la seva còpia de l'adreça d'OpenStreetMap.

- **Un sol fitxer per als tres mapes**: `rutes/MapaFons.js`, que
  `Get-MapaFonsJs` (`MapaHtml.ps1`) posa **tal qual** dins de l'HTML. Les
  plantilles hi porten `{{fonsJs}}`, que omple **`Get-PlantillaHtml` per a
  totes** (no és cosa de cap eina); Ruta, que no fa servir plantilla, el crida
  directament. Cada mapa només fa `var FONS = afegeixFonsMapa(map)`.
- **L'ordre**: *Mapa (ICGC)* → *Ortofoto (ICGC)* → *Mapa (Esri)*. L'ICGC és
  la cartografia oficial de Catalunya i no demana clau ni Referer. Hi ha un
  selector a dalt a la dreta i **la tria es recorda** (`localStorage`, comú a
  tots els mapes oberts des del disc).
- **CARTO es va treure** (octubre 2026): les seves rajoles ara demanen clau i
  **carreguen igualment** amb «API KEY REQUIRED» pintat a sobre. Com que
  carreguen, el recanvi automàtic no ho veu. Abans d'afegir un fons nou, mira
  que no en demani.
- **Sobre l'ortofoto la vora de les parcel·les és blanca i més gruixuda** (i el
  color, més transparent): la fosca es perdia entre teulades i ombres. El mapa
  rep l'event `fonscanviat` cada vegada que canvia el fons.
- **Si un fons no respon, es passa sol al següent** (cap rajola carregada i ja
  `FONS_ERRORS_MAX` errades) i es diu en un requadre a baix a l'esquerra. **No
  es van poder provar les adreces des de l'entorn de desenvolupament** (el
  proxy bloqueja tots els servidors de mapes): el recanvi automàtic és el que
  fa que una adreça equivocada o caiguda no deixi el mapa sense fons. Si mai
  surt l'avís a la feina, és que aquell servei ha canviat i cal revisar
  l'adreça.
- `maxNativeZoom` conservador (18-19) i `maxZoom` 22 a cada capa: si el mapa
  s'apropa més del que el servei dona, la rajola s'amplia en lloc de fallar (i
  de fer saltar el recanvi per error). El zoom màxim del mapa el fixa cada
  mapa (`maxZoom` a `L.map`).
- **Prova al navegador** (`prova-planol.mjs`, «El fons del mapa»): cap petició
  a OpenStreetMap; amb l'ICGC caigut i Esri servit, passa sol a Esri i ho
  diu; la tria del selector es recorda en tornar-lo a obrir; i sense cap fons,
  ho diu i el mapa segueix funcionant.
- El **plànol públic de precintades** (`docs/precintades.html`) es queda amb
  OpenStreetMap: es publica a GitHub Pages, que sí que envia Referer.

## Eina «Plànol activitats» (`rutes/Planol.ps1`, octubre 2026)

### On va l'ID GIA: la COORDENADA UTM de l'Excel d'activitats (`rutes/PlanolGeometria.ps1`)
Tres rondes amb l'usuari (octubre 2026), i **la que mana és l'última**:
1. *«l'ID GIA a l'entrada de l'establiment»*: el portal del Cadastre amb el
   número exacte, dins de la parcel·la; si no es troba, al centre i en vermell.
2. Amb dues captures: en un polígon industrial (**Sant Ferran**) el Cadastre té
   **una sola porta** per a 13 naus i hi sortien 12 ID apilats. Els números de
   nau del fons (1… 13) són part de la **imatge** de l'ICGC, i el Cadastre diu
   «Esc. 1 - Pl. baixa - Pt. 13» **sense coordenades**: no hi ha d'on treure la
   posició de cada nau. Es va fer un «situa'l a mà» (al `localStorage`)…
3. …i l'usuari: *«És igual, dibuixa les etiquetes amb el ID GIA segons les
   coordenades UTM de la base de dades d'activitats [...] Ho modificaré primer
   amb l'eina coordenades i després es veurà reflectit al plànol»*. El situar a
   mà **es va treure**: dues fonts per a la mateixa posició acabarien dient
   coses diferents. **La font és la columna UTM X / UTM Y de la fulla Estès**
   (`ActX`/`ActY` de cada entrada), la que corregeix Coordenades.

Com es decideix (`Get-CasesActivitats` + `Get-EtiquetesGrup`, a `PlanolDades.ps1`):
- Una activitat té **una** coordenada i potser diversos establiments. Primer es
  busca la seva **casa**: el grup (parcel·la o parcel·les juntades) on cau.
  - Si cau **dins** d'una parcel·la dibuixada: la d'aquella si és seva; si és
    d'**una altra activitat**, cap (la coordenada està malament: **vermell**),
    encara que la seva sigui a 2 m. Si no, es diria «bé» d'un punt que
    l'usuari ha d'arreglar.
  - Si no cau dins de cap: la seva parcel·la més propera a menys de
    `PlanolEntradaMaxForaM` (12 m: un punt a la façana o una mica al carrer, com
    el punt verd de Coordenades). Es posa a `PlanolEntradaMargeM` (2,5 m) cap a
    dins (`Resolve-AncoraEntrada`) i l'etiqueta **creix cap a l'interior**.
  - Es busca amb una **graella de 100 m** sobre les capses de les parcel·les
    (`Get-CapsaPoligons`): provar cada punt dins de les ~950 parcel·les seria
    1,3 milions de proves en PowerShell.
- El codi `x` de cada entrada (va al mapa): **1** a la seva coordenada; **2** la
  coordenada és a l'**altre establiment** de l'activitat (al centre, **sense**
  vermell: no és cap error); **0** fora de les seves parcel·les i **3** sense
  coordenada (tots dos al centre, **en vermell**, i al filtre *Per revisar*);
  **-1** la parcel·la no té dibuix (un punt: no se sap).
- Les que cauen al **mateix punt** (menys de `PlanolMateixPuntM`, 3 m) van a la
  mateixa etiqueta. Abans de repassar-les amb Coordenades, les d'una parcel·la
  solen ser la mateixa (el GIA porta la del Cadastre, que és de la parcel·la).
- **Parcel·les juntades**: les que es toquen i tenen **exactament les mateixes
  activitats** es dibuixen com una sola forma (`Join-Poligons`: unió per
  cancel·lació d'arestes, amb els costats partits pels vèrtexs de la veïna). Si
  una porta una activitat més no es junten: s'hi barrejaria el color de l'altra.
- El **centre** és un punt **dins** (`Get-PuntInterior`): el centroide d'una
  parcel·la amb pati hi queia a dins.
- **Al plànol, només els ID** (`textEntrada`), de quatre en quatre: fins a 4 i
  «+N» a mig zoom, fins a 16 de prop. El local/planta/porta va a la **fitxa**:
  al plànol tapava les naus del costat (*«treu els texts aquests del plànol que
  estorben»*).
- **Les dues adreces, diferenciades** (*«no té per què ser exactament la mateixa
  la de la base de dades d'activitats i la del cadastre»*): «Adreça (base
  d'activitats)» (`ad`) i «Adreça (Cadastre)» (`ca`: el portal de la parcel·la
  amb el número de l'establiment, `Get-PortalsExactes` + `Get-AdrecaPortal`); a
  dalt, totes les de la parcel·la (`pa`, `Get-AdrecesCadastre`). **Els portals
  només serveixen per a això**; es segueixen demanant amb la memòria cau de
  Coordenades (`portals.json`).
- **Filtres**: *Allotjaments turístics* (CCAE 552/5520, columna «CCAE Codi»
  de la fulla Estès; per defecte, **sense**), *Classificació (annex)* (una
  casella per cada valor de «Classificació general annex») i *Per revisar*,
  amb recomptes, una ajuda a sota i «Coordenada UTM fora de la parcel·la o sense».
- **L'etiqueta, CENTRADA al seu punt** (`direction: 'center'`): és on acaba la
  línia de punts. Abans creixia cap a dins des del punt i la línia sortia de
  sota.
- **Clic a un ID → la fitxa d'AQUELLA activitat** (`obreFitxaActivitat`, amb tots
  els seus establiments, també d'altres parcel·les, i un enllaç «Totes les
  activitats de…»); **clic a la parcel·la → totes**. Els tooltips del Leaflet no
  reben el ratolí (`pointer-events: none`): `.leaflet-tooltip.ent` el torna a
  activar i el clic s'atura a l'etiqueta (`lligaClicEtiqueta`), si no el mapa
  tancaria la fitxa que s'acaba d'obrir.
- **Una fila per ACTIVITAT** (`perActivitat` + `filaActivitat`): una activitat amb 3
  establiments a la mateixa parcel·la (el 122 de l'usuari, CTRA PRAT 77) surt UN
  cop, amb «3 establiments:» i els seus locals. També al CSV i al cercador.
- **Locals buits DUPLICATS al GIA** (el 1365 de l'usuari): el mateix local hi és
  dues vegades, un establiment amb l'activitat i un altre de buit (mateixa refcat
  de 20, adreça i local/bloc/escala/pis/porta, `_PlanolClauLocal`; una refcat
  sola no basta, un edifici pot ser una sola unitat). A l'Excel de l'octubre:
  37. **No es compten com a buits** (`du`), l'activitat porta l'ID de
  l'establiment buit (`bd`) i surt a *Per revisar: Local buit duplicat al GIA*.
  Decidit amb l'usuari: s'ha d'arreglar al GIA, no amagar-ho.
- **Al mòbil** (*«l'eina s'ha de poder usar des del mòbil»*): el mateix HTML es
  puja al Drive privat (`mobil/PujaPlanol.ps1`, llançat amb
  `Start-ScriptSegonPla`, que per això ha passat de `Motor.ps1` a `SegonPla.ps1`:
  el procés de `rutes/` no pot carregar el Motor) i el mostra `docs/planol.html`.
  A menys de 760 px el lateral és un panell de baix que s'obre amb «Filtres», i
  dins de l'app cada activitat porta «Fer informe» (`DINS_APP`, `postMessage`).
  Detalls a `DESPLEGAMENT-MOBIL.md`.
- **La línia de punts** (l'usuari: *«una lleugera línia de punts que uneixi
  l'etiqueta de l'ID GIA (UTM base de dades activitats) i la UTM de la parcel·la
  cadastral»*): de cada etiqueta posada a la seva coordenada fins al
  **`referencePoint`** de la parcel·la on cau (`Get-PuntReferenciaParcela`, a la
  mateixa resposta del wfsCP; camp `r` de l'etiqueta, `_PlanolPuntDeLaParcela`
  per a les juntades). Les del centre (vermelles) no en porten, ni les que són a
  menys d'1 m (encara no corregides). S'amaga amb l'etiqueta; blanca sobre
  l'ortofoto.
  - **La memòria cau de les parcel·les va a un fitxer NOU, `parceles2.json`**
    (camp `Parcela` = `{ Poligons; Punt }`). Afegir-hi el punt amb un altre camp
    al fitxer vell hauria fet que `Test-CacheCadastreValida` donés per bones
    (buides, 30 dies) les entrades sense aquell camp: parcel·les dibuixades com
    un punt. **El primer plànol després d'això torna a demanar totes les
    parcel·les** (uns minuts, un sol cop).
- **Quan es veu una correcció de Coordenades**: quan l'Excel d'activitats que
  llegeix el plànol (el `AAAA-MM-DD ACTIVITATS.xls` més nou) ja la porta, o
  sigui **després d'importar-la al GIA i tornar-lo a baixar**. L'«Excel per
  importar» (`… - coordenades corregides AAAA-MM-DD HHmm.xls`) **no** el llegeix:
  el nom no casa amb `_RutaFindLatestIn`, i és a posta (és una còpia per a qui
  importa, no la base).

> **El primer plànol de l'usuari: TOTES les activitats «sense establiment».**
> La crida feia `@(Read-EstablimentsExcel …)`, que ja torna la llista amb coma:
> el model rebia **un** establiment que les contenia tots (1.582), no en lligava
> cap amb la seva activitat i no hi havia cap local buit. Les proves no ho
> veien perquè provaven el lector i el model per separat, no la crida que els
> uneix. Ara: la crida sense `@()`, `Build-PlanolModel` desplega una llista
> embolcallada, una prova del lector amb un `Read-FullaEstesa` fals i un guard
> d'AST per a tot el projecte. L'avís «sense establiment: ref. de l'Excel
> d'activitats» vol dir que **aquella** activitat no surt a cap fila de l'Excel
> d'ESTABLIMENTS (columna «ID Activitat») i se situa amb la referència
> cadastral de l'Excel d'activitats; ha de sortir poc, no a tot arreu.

Parcel·les de Cornellà pintades segons l'estat de les activitats que hi ha. Es
**només local** (`local\planol-activitats\`): porta requeriments pendents.

- **Fitxers:** `Planol.ps1` (el flux i la finestra), `PlanolDades.ps1` (tot el
  que és pur: lectura de les fulles a partir de la matriu, colors, local/planta/
  porta, model per parcel·la, parsers del Cadastre, dades del mapa),
  `PlanolMapa.html` (la plantilla). Comparteix amb Coordenades `Cadastre.ps1`
  (xarxa, memòria cau, bucle amb progrés), `MapaHtml.ps1` (plantilla i JSON dins
  d'un `<script>`) i `EinesUi.ps1` (missatge i barra de progrés).
- **Per què l'Excel d'ESTABLIMENTS** (`AAAA-MM-DD ESTABLIMENTS.xls`, fulla
  «Establiments», a la mateixa carpeta que el d'activitats): l'Excel d'activitats
  porta UNA refcat per activitat, i una activitat pot tenir diversos
  establiments. També hi ha «Local buit» i el local/pis/porta. El nom va arribar
  amb `_` entre la data i el nom: `_RutaFindLatestIn` accepta espai o `_`.
  **Les capçaleres d'aquest Excel són rares**: `Emp._Número_`, `Emp. Nº Local`.
  Per això `Find-HeaderColumn`, si no troba la coincidència exacta, compara
  només lletres i números.
- **Els colors** (`Get-EstatPlanol`, decidits amb l'usuari): vermell = precinte
  a l'Excel (`Test-IsPrecintada`, ara a `Excel.ps1`) o darrer informe
  «Precinte / Cessament»; groc = «Requeriment» o «Ampliació termini»; verd =
  «Favorable», «FI Requeriment», «FI Precinte / Cessament»; **blau** = tota la
  resta, sense informes inclòs («no sabem si està legalitzada»). L'estat ve
  d'`estat_actual` de la base d'informes, tal com el deixa *Actualitzar base* /
  *Editar base* (aquest procés no carrega `Informes.ps1`).
- **Una parcel·la = el PITJOR** dels que passen el filtre (vermell > groc > blau
  > verd). Ho decideix el mapa, perquè depèn dels filtres. Els recomptes són
  d'activitats úniques (una amb dos establiments compta un cop).
- **Local/planta/porta** (`Get-SubEstabliment`): primer el de l'Excel
  (local, bloc, escala, pis, porta); si no en porta, el del **Cadastre**
  (`Consulta_DNPRC`, escala/planta/porta de la unitat); si tampoc, el número
  d'**unitat** (caràcters 15-18 de la refcat). Al Cadastre només es pregunten
  les unitats que ho necessiten: refcat de 20, sense res a l'Excel, i en una
  parcel·la amb més d'un establiment (`Get-UnitatsAConsultar`: 292 amb l'Excel
  de l'octubre de 2026). Les paraules de l'Excel van amb etiqueta («Bl. C»,
  «Pl. BXS») tret del local, que ja diu què és («NAU 6»).
- **Geometria**: INSPIRE `wfsCP` (`GetParcel`), en EPSG:25831 com tot. Un
  polígon pot tenir forats, i una parcel·la diversos polígons. Els anells es
  guarden **plans** (`[x1, y1, x2, y2…]`): un array d'arrays en PowerShell es
  desenrotlla a la mínima. Guàrdia d'eixos com als portals. Sense geometria, la
  parcel·la surt com un **punt** a la coordenada de l'Excel.
- **Les fixtures del Cadastre (`wfsCP-exemple.xml`, `dnprc-*.xml`) estan muntades
  a mà**, com la dels portals: el host està bloquejat des d'aquest entorn. Abans
  de fiar-se'n, `suport\rutes\Provar-Planol.bat` a la feina: desa les respostes
  de debò a `local\geocodificacio\resposta-parcela-*.xml` / `resposta-unitat-*.xml`.
- **Memòria cau**: `parceles.json` i `unitats.json` a `local\geocodificacio\`,
  un any (30 dies si no hi ha resultat). La primera vegada són ~950 + ~300
  consultes (uns minuts); després, segons. **Cancel·lar no avorta**: el mapa es
  fa amb el que hi hagi.
- **Mida**: amb l'Excel real, 943 parcel·les i ~360 KB de dades sense geometria.
- Proves: `tests/run-tests-planol.ps1` (142) i `tests/navegador/prova-planol.mjs`
  (66, al Chromium, amb `genera-planol-prova.ps1`).

### Mode automàtic setmanal (`PlanolAutomatic.ps1` + `PlanolAuto.ps1`, octubre 2026)
«Aplica l'automatisme de Copiar informes i Actualitzar base a Plànol activitats,
però un cop a la setmana». Mateix patró que els altres A/M: entrada a
`$Script:ModesAuto['planol']` (requisit: la carpeta de l'Excel d'activitats),
programació `planol` **setmanal, dilluns a les 13:00** per defecte (canviable a
*Configuració → Automatismes*), estat a `local\planol-activitats\planol-auto.json`,
mutex `Global\InformesCornella.PlanolActivitats` i registre a
`%LOCALAPPDATA%\InformesCornella\planol-log.txt`.
- La feina és **`Invoke-PlanolGenera($silenci)`**: el botó (`Invoke-PlanolMain`)
  i el procés a part la comparteixen. En silenci no hi ha cap finestra (ni barra:
  `_PlanolAmbProgres` fa la feina directament) i, si falta l'Excel
  d'ESTABLIMENTS, continua amb un avís al registre en comptes de preguntar.
- **El botó pregunta abans de generar** («que pregunti si es vol fer un de nou o
  consultar l'existent, així no se'n farà un cada vegada que premo»): si hi ha
  algun `Planol_*.html` a `local\planol-activitats\` (`Get-PlanolUltim`),
  `Show-EinaTria` (`EinesUi.ps1`, botons amb `_AddPeuBotons`) ofereix
  *Consultar l'últim* (només l'obre: ni Excel ni pujada al Drive) o *Fer-ne un de
  nou*. Sense cap plànol fet, es genera directament.
- `PlanolAuto.ps1` carrega `rutes/Planol.ps1` amb `$PlanolNomesFuncions = $true`
  (no obre res) i, si va bé, **puja el plànol al Drive** (`Save-ADadesDrive
  'planol.html'`) perquè el mòbil tingui el de la setmana.

## LA TRAMPA DE `$Script:` DINS D'UN `.GetNewClosure()` (mesurada, octubre 2026)

**Dins d'una closure, `$Script:X` NO és la variable de l'script**: llegir-la torna
buit i escriure-hi no arriba enlloc (cada closure té el seu propi mòdul).
Mesurat amb pwsh 7. El Cancel·lar de la barra de progrés de **Coordenades no
aturava mai res** per això: el botó posava `$Script:CoordCancelat` i el bloc de
progrés (amb closure, perquè ha de veure `$prog`) el llegia des de dins.
**Solució**: un hashtable compartit (`New-EinaProgres` torna `Estat`, i el bloc
llegeix `$prog.Estat.Cancelat`). Un hashtable és una referència: la closure i el
botó veuen el mateix. Hi ha altres llocs del programa amb el mateix patró
pendents de revisar (Configuració, Normativa, Revisió, Ruta).

## Eina «Coordenades» — Excel vs façana (`rutes/Coordenades.ps1` + `rutes/Geocodificador.ps1`)
- **El problema, mesurat** (base del 18/08/2026): el GIA porta les coordenades
  del Cadastre, i el Cadastre georeferencia la **parcel·la**, no el local. 1.380
  activitats → **899 punts diferents**; 711 files (52%) en comparteixen un amb
  alguna altra. De les 898 parcel·les, **227 tenen més d'una activitat** i
  concentren aquelles 711 files. Pitjor cas: `4091106DF2749A` (Ctra. de
  l'Hospitalet 147) amb **19 apilades**.
- **La coordenada verda NO surt de geocodificar el text de l'adreça.** Surt de
  la `Ref. cadastral` que ja hi ha a l'Excel: els **14 primers caràcters** són la
  parcel·la, i al Cadastre se li demanen els **portals** d'aquella parcel·la
  (servei INSPIRE d'Adreces, `wfsAD.aspx`, consulta desada `GetadByRefcat`). Cada
  portal és un punt d'**entrada** amb el seu número de carrer. Avantatges: **una
  consulta per parcel·la** (no per activitat) i la resposta ja ve en **EPSG:25831**,
  el mateix sistema que l'Excel — cap reprojecció, cap error de conversió.
- **NO toca Ruta.ps1 ni Precintades.ps1.** Va ser petició explícita de l'usuari:
  aquells segueixen amb la coordenada original. L'eina només MIRA i genera un
  fitxer.
- **Tres xarxes de seguretat, i totes tres hi són a posta:**
  1. `Resolve-CoordEstabliment` descarta qualsevol portal a més de
     `$GeoDistanciaMaximaM` (250 m) de la parcel·la. Val més un marcador imprecís
     que un marcador mentider.
  2. `ConvertFrom-CatastroAdXml` parseja **per `local-name()`**, sense lligar-se a
     cap espai de noms ni nivell de l'arbre, i gira els eixos si venen a l'inrevés
     (en UTM 31N l'est ~420.000 va molt per sota del nord ~4.578.000).
  3. Res del mòdul llança mai: si el servei no respon, la parcel·la queda sense
     portals i cada activitat es queda amb la seva coordenada de sempre.
- **`$GeoCatastroUrlTemplate` és una VARIABLE, no una cadena enterrada al codi**:
  si el Cadastre canvia el nom del paràmetre de la consulta desada, s'arregla des
  de `config.ps1` sense tocar el programa. Per això `Coordenades.ps1` carrega
  `Geocodificador.ps1` **ABANS** de `Ruta.ps1` — que és qui carrega `config.ps1`:
  si es carregués després, els valors per defecte del mòdul trepitjarien el que
  l'usuari hagués posat a `config.ps1`.
- **EL CLIENT DE XARXA NO S'HA POGUT PROVAR CONTRA EL SERVEI REAL.** L'entorn on
  es va escriure tenia `ovc.catastro.meh.es` bloquejat per política de sortida
  (403 al CONNECT; també ICGC, Cartociudad i Nominatim). La fixture
  `tests/dades/wfsAD-exemple.xml` està **muntada a mà** seguint l'esquema INSPIRE,
  no gravada. Abans de fiar-se'n, a la feina:
  ```
  suport\rutes\Provar-Cadastre.bat        (doble clic; accepta una refcat com a argument)
  ```
  Ha de llistar els portals de Cadis i Huelva amb els seus números.
  **NO cridis `. Ruta.ps1` a pèl** per fer-ho: `Ruta.ps1` executa la seva `Main`
  i t'obre el planificador de rutes (va passar). `Coordenades.ps1` en mode
  headless ja ho carrega tot sense obrir res.
  `Test-Geocodificador` compta les adreces de la resposta **crua** (per
  `regex`, sense parsejar) i les compara amb les que ha entès el parseig: així
  es distingeix «el servei no ha tornat res» de «no n'he sabut treure res». I
  desa **sempre** la resposta sencera a `local/geocodificacio/resposta-<rc>.xml`,
  que és l'única cosa que permet arreglar el parseig sense anar a les palpentes.
- **EL MAPA ÉS UNA PLANTILLA A PART: `rutes/CoordenadesMapa.html`** (octubre
  2026). Abans era un here-string de 645 línies dins de `Build-CoordenadesHtml`,
  amb dues trampes permanents: qualsevol `$` o `` ` `` del JavaScript era una
  interpolació de PowerShell, i tot el text català depenia que el `.ps1` no
  perdés el BOM. Ara:
  - la plantilla és HTML/JS normal, amb marques **`{{nom}}`** per a les dades
    (`{{itemsJson}}`, `{{dbEnc}}`...). Ja s'hi pot escriure `$` i template
    literals, però **cap `{{paraula}}`** que no sigui una marca;
  - es llegeix amb **`[IO.File]::ReadAllText(..., UTF8)` explícit**: sense dir-li
    res, el 5.1 la llegiria com a ANSI i sortiria `Ã§`. Es desa en UTF-8 sense BOM;
  - `Expand-CoordPlantilla` l'omple en **una sola passada** (un valor que porti
    `{{x}}` no es torna a substituir) i **llança** si una marca no té valor;
  - els JSON s'injecten amb `</` → `<\/`: un `</script>` a l'adreça tancaria
    l'etiqueta i la pàgina no arrencaria;
  - es va comprovar que l'HTML resultant és **byte a byte** el mateix que el del
    here-string (4, 1 i 0 activitats, amb accents i cometes).
  Amb això `Coordenades.ps1` va baixar de 1.458 a ~870 línies i ja no necessita
  l'excepció de mida de `06-guards.ps1`. `Coordenades.ps1` **segueix portant
  BOM** (hi ha `Cancel·lar`, `Parcel·la`... a les finestres). `Geocodificador.ps1`
  és ASCII pur i no en porta (com `Precintades.ps1`).
- **Si `unpkg` no respon, el Leaflet es demana a jsDelivr** (`cdn.jsdelivr.net/npm/`),
  que serveix el mateix paquet de npm **byte a byte**: el SRI és el mateix i no
  cal afluixar-lo. **No hi posis cdnjs**: no és una còpia garantida de npm i, si
  el hash no hi coincidís, el segon intent fallaria igualment. Si tampoc no
  arriba, la pàgina diu **en clar** que no s'ha pogut carregar el mapa (abans es
  quedava en blanc) i que el repàs no s'ha perdut. Les dues coses les prova la
  suite del navegador.
- **L'`.xlsx` el genera el NAVEGADOR, sense cap biblioteca.** Un `.xlsx` és un ZIP
  amb cinc XML a dins; amb el mètode «sense compressió» només cal el CRC-32 i les
  capçaleres del ZIP (`crc32`/`zipStore`/`buildXlsx` a `CoordenadesMapa.html`). Els textos
  van **inline** (`t="inlineStr"`), així no cal `sharedStrings.xml`. Verificat:
  el fitxer generat el valida `zipfile` i l'obre `openpyxl` **sense avisos**, amb
  números com a números i accents intactes. Sense el `<cellStyles>` a
  `styles.xml`, `openpyxl` es queixa («no default style»).
- **`latLonToUtm31` (al JS) és la INVERSA de `Convert-UtmToLatLon`** i cal perquè
  Leaflet dona graus quan s'arrossega un punt i nosaltres hem d'exportar metres.
  Comprovada d'anada i tornada sobre 525 punts de tot el terme municipal: error
  màxim **0,07 mm**. Una coordenada que **no** s'ha mogut a mà s'exporta amb els
  metres **tal com van arribar** (`utmActual`), sense reprojectar: així no s'hi
  acumula l'error d'anar i tornar.
- **Les eines del repàs (octubre 2026, les va triar l'usuari):**
  - **Filtre per estat** (`#filtreEstat`, `passaEstat`): pendents, validades o
    un color (`origen`). Amaga la fila **i** les tres capes del mapa (verd,
    vermell, línia); «Validar tot el que es veu» valida exactament el que queda.
  - **La seleccionada (`sel`) no s'amaga mai pel filtre**: amb «Pendents», la que
    acabes de validar desapareixeria de sota el ratolí abans de veure-la. Marxa
    quan en selecciones una altra.
  - **Ressaltar** (`pintaSeleccio`): fila, verd (classe `sel`), vermell i línia.
  - **Següent pendent** (botó i tecla N): la pendent **més propera** a la
    seleccionada (o al centre del mapa), no la següent de la llista: es repassa
    un edifici sencer abans de saltar. Recorda les que ja ha ensenyat
    (`vistesSeguent`, **també la d'on surts**: sense això, dues apilades a 0 m es
    passaven la pilota) i torna a començar quan s'acaben.
  - **Desfer** (botó i Ctrl+Z, no dins del cercador): pila en memòria de l'estat
    d'un punt abans de cada arrossegament **o clic** (desvalidar un punt mogut el
    torna al Cadastre: era la manera de perdre la posició sense voler).
  - **Mai `setIcon` al `dragstart`**: refà l'arrossegador de Leaflet i talla el
    drag. Per això el punt es selecciona al `dragend`.
  - **Les JA CORREGIDES** (*«ha de quedar ben clar aquelles coordenades que han
    sigut modificades respecte les coordenades del cadastre»*): el **punt de la
    parcel·la al Cadastre** (`referencePoint`) es demana amb el mòdul comú
    **`CadastreParceles.ps1`** (abans era del Plànol; ara també de Coordenades,
    amb la MATEIXA memòria cau `parceles2.json`). Si l'Excel en és a
    `$Script:CoordCorregidaM` (1 m) o més: `Corregida`, punt de l'Excel **lila**,
    línia de punts fins al punt del Cadastre (un cercle petit), fila marcada,
    filtre «Ja corregides» i la distància a la fitxa i a la targeta. **Supòsit**:
    el GIA porta el `referencePoint` del Cadastre; si a la feina surten totes
    lila, el GIA en fa servir un altre i cal mirar-ho amb `Provar-Planol.bat`.
  - **La tria de zones** (*«també vull poder moure els punts de les activitats que
    no estan duplicades»*): tres opcions (apilades / NO apilades / totes,
    `Get-RegistresPerAbast`), «Amaga les ja corregides» (es mira en generar el
    mapa: la finestra encara no sap el punt del Cadastre) i «Només les marcades
    per revisar» (viuen al navegador: el mapa es fa amb **totes** les zones i
    s'obre amb el filtre «Per revisar», `FILTRE_INICIAL`). Abans, sense cap
    apilada, l'eina plegava.
  - **Per revisar** (l'usuari: *«hi ha punts estranys. Vull poder posar un
    warning per revisar posteriorment»*): a la fitxa del punt, «⚠ Marca per
    revisar» amb una nota opcional. **A part del repàs** (`coordenades-avisos:`
    + base, `avisos`/`teAvis`): marcar no valida ni mou el punt. Surt un «!» al
    punt, ⚠ a la fila, un filtre «Per revisar» i una línia a la llegenda. Va a
    l'Excel del repàs a la columna **«Per revisar»** (la nota, o «Sí»); una fila
    només marcada **no porta coordenada nova** i ni l'«Excel per importar»
    (`Get-CorreccionsDelRepas`, `PerRevisar`) ni «Carregar repàs» la compten
    com a invàlida; aquest últim recupera l'avís (el del navegador guanya).
    «Esborrar el meu repàs» també els esborra.
    - **Trampa del Leaflet (mesurada)**: un enllaç de la fitxa que refà la
      fitxa a mig clic fa que l'enllaç clicat ja no hi sigui, i el Leaflet el
      pren per un clic al **mapa** i la tanca. Els enllaços de l'avís fan la
      feina amb `setTimeout(…, 0)` (`enllacAvis`).
  - **La targeta en passar-hi el ratolí** (*«si una coordenada l'he mogut, al
    posar-me sobre em digui les dades, perquè si tinc dubtes i clico em retorna
    a la posició original»*): `targetaHtml`, un tooltip al verd i al vermell
    amb l'ID, el titular, l'activitat, l'adreça, d'on surt, els metres, si és
    validat, l'avís, i —si és mogut i validat— què faria un clic. Mirar-lo no
    el toca. Es refà a `refrescaItem`.
  - **L'adreça sencera** (*«posa'm tota l'adreça, no només carrer i número»*):
    la de l'activitat porta darrere `Emp. Bloc/Escala/Pis/Porta`, i la del
    titular surt de `Raó soc. Tipus via/Carrer/Número/Escala/Pis/Porta`
    (`Format-AdrecaSencera`, etiquetes del Plànol: Bl., Esc., Pl., Pt.; camp
    `adt` al mapa). Columna que no hi sigui, buida. **Compte**: amb aquestes
    columnes hi ha molts «Raó soc. …» que no són el nom; `Get-ColumnaTitular` les
    descarta (`$Script:CoordNoEsTitular`) i, sense «Rao social» exacte, va
    primer a la 10 (la d'`Activitats.ps1`).
  - **El titular** (l'usuari: *«vull veure el titular de l'activitat a l'eina
    Coordenades»*): la «Raó social» de la fulla Estès (`Get-ColumnaTitular`, per
    nom amb respatller a la columna 10, la que fa servir `Activitats.ps1`). Surt a
    la fitxa del punt, sota l'adreça a la llista, i el cercador el troba. **No va
    a l'Excel del repàs** (no li cal a qui importa) ni surt mai de l'ordinador:
    el mapa és a `local\geocodificacio\` i al Cadastre només hi va la refcat.
  - **El plànol del Cadastre** (WMS oficial, parcel·les i edificis) a sobre
    del fons: casella «plànol del Cadastre» a la barra (l'usuari: *«posa'm el
    plànol del cadastre a Coordenades també»*). La capa és a **`MapaFons.js`**
    (`CADASTRE_WMS`, `lligaCasellaCadastre`) i és **la mateixa** que la casella
    del Plànol activitats i l'entrada «Plànol del Cadastre» del selector de dalt
    a tots els mapes: abans el Plànol en tenia una còpia pròpia. Apagada per
    defecte; si l'encens, es recorda (`informesCornella.planolCadastre`). Guard a
    `run-tests-planol.ps1`: l'URL del WMS només a `MapaFons.js` (validat
    injectant-lo al Plànol).
  - **Zoom fins al 22** (abans 19; l'usuari: *«deixa'm fer més zoom que sinó
    estan molt lluny»*). Les rajoles del fons arriben al 18 i d'allà en amunt
    s'amplien (`maxNativeZoom` a `MapaFons.js`). «Anar a» una activitat (`vesA`)
    apropa fins al 19 però **no allunya** si ja eres més a prop.
  - La llegenda, amb recomptes, va **a dalt** del panell: al final de la llista
    no es veia mai. El cos és una columna flex: el mapa ocupa el que queda (abans
    `calc(100vh - 116px)`, i amb la barra en dues línies la pàgina feia scroll).
- **El repàs té còpia fora del navegador: l'Excel que es baixa** (octubre 2026).
  «Carregar repàs…» (`llegeixXlsx` + `aplicaRepas`) el llegeix tal com surt del
  mapa (ZIP sense comprimir, textos inline) **i** desat de nou amb l'Excel
  (deflate, `sharedStrings.xml`, la fulla amb un altre nom intern): el deflate el
  fa `DecompressionStream('deflate-raw')` del mateix Chrome, sense biblioteques.
  - **El que ja és al navegador guanya**: el fitxer és una còpia i pot ser vella.
  - L'Excel porta la columna **«Base de dades»** (`FONT`): no es carreguen files
    d'una altra base. Els fitxers d'abans no la porten i s'accepten.
  - `utm31ToLatLon` és `Convert-UtmToLatLon` passada a JS (l'Excel només porta
    metres). Anada i tornada sobre el terme: 0,07 mm.
  - **`openpyxl` no serveix per simular «desat amb l'Excel»**: la 3.1 també
    desa els textos inline. Per això hi ha `tests/navegador/simula-excel.py`.
- **«Excel per importar…»** (`rutes/CoordenadesImportar.ps1`, octubre 2026).
  L'usuari no entra les correccions: les **importa una altra persona**, que les
  vol **amb el mateix format que la base original** i les canviades en vermell.
  - No es munta cap Excel: es **copia el fitxer de la base** tal qual
    (`[IO.File]::Copy`) i només s'hi reescriuen les cel·les `UTM X`/`UTM Y` de les
    activitats corregides, amb `Font.Color = 255` (vermell en BGR).
  - Es fa amb **`Read-FullaEstesa -Desa`** (`Excel.ps1`): obrir l'Excel és en un
    sol lloc (hi ha guard), i `-Desa` obre per escriure i desa **només si el cos
    acaba bé**. Provat amb el doble de COM.
  - **Una activitat que a la base ja no té la coordenada de quan es va repassar
    NO es toca** (`JaCanviades`): la base ha canviat, potser ja s'ha corregit, i
    sobreescriure-la seria desfer feina d'algú altre.
  - **El tipus de la cel·la es conserva** (`Format-CoordComOriginal`): número →
    número; text → text amb el mateix separador i **amb apòstrof davant**, perquè
    escriure `'421982,90'` per COM en un Excel en català el convertiria en número.
    **No s'ha pogut provar amb l'Excel de debò**: és al punt 11 de `provar-al-pc.md`.
  - Capes: `Read-RepasXlsx` → `Get-CorreccionsDelRepas` → `Get-EscripturesCoordenades`
    (purs, provats amb fixtures de `tests/dades/repas-*.xlsx`) → `Set-CoordenadesALaBase`
    (l'únic que toca COM). La finestra (`Invoke-CoordExcelImportar`) és a
    `Coordenades.ps1`: el mòdul **no pot cridar el seu client** (guard de cicles;
    per això `Test-CoordPlausible` va passar a `Geocodificador.ps1` i
    `Get-IdDeCella` és al mòdul).
- **`estat[]` va per POSICIÓ dins d'`ITEMS`, no per ID**: si algun dia la base
  portés dos cops el mateix ID Activitat, dues fitxes es trepitjarien. Al
  `localStorage`, en canvi, es desa **per ID**, que és el que ha de sobreviure
  quan es torni a generar el mapa.
- **El mapa es prova en un navegador de debò, i ara de manera repetible:**
  `suport/tests/navegador/prova-mapa-coordenades.mjs` (Chromium + Playwright).
  Genera tres mapes amb les funcions de debò (`genera-mapa-prova.ps1`) i hi fa
  el que fa l'usuari: validar amb un clic, arrossegar, filtrar, baixar l'Excel
  (rellegit amb `openpyxl`, que s'ha d'obrir **sense avisos**), tancar el
  navegador i obrir un mapa **nou**, i esborrar el repàs. Detalls que cal saber:
  - el **Leaflet es serveix des de `node_modules/leaflet`**: és el mateix
    paquet 1.9.4 de npm, byte a byte, i el SRI hi coincideix. **No cal treure
    l'`integrity`** (la primera prova a mà, al setembre, ho havia de fer perquè
    feia servir un doble de Leaflet);
  - el navegador és **persistent** (un perfil en una carpeta temporal): és
    l'única manera de provar que el repàs sobreviu a tancar el Chrome;
  - **comprovat (octubre 2026): a Chromium el `localStorage` de `file://` es
    comparteix entre fitxers de carpetes i noms diferents**, i per això el repàs
    es veu des de cada `Coordenades_<data>.html` nou. L'usuari fa servir Chrome.
    A Firefox **no** s'ha comprovat (cada fitxer `file://` hi pot tenir origen
    propi); si algun dia es canvia de navegador, és el primer que cal mirar;
  - validada injectant el defecte: amb la clau del repàs lligada al nom de
    l'HTML, cinc comprovacions es posen en vermell.
- **LA TRAMPA DEL `return ,@(...)`, TERCERA APARICIÓ — i la primera crida real
  al Cadastre la va destapar.** `ConvertFrom-CatastroAdXml`, `Get-RegistresApilats`
  i `Get-RefcatsAConsultar` acabaven amb `return ,@(...)` **i** totes les crides
  les embolcallaven amb `@()`. Les dues coses juntes hi posen la capa **dues
  vegades**: `$portals.Count` valia **1**, `$p.Numero` feia enumeració de
  membres i el diagnòstic escrivia `numero='System.Object[]'`. El servei
  funcionava perfectament; el que fallava era el consum.
  - **Tria una convenció i escriu-la al costat de la funció.** Aquí és **array
    pla + `@()` al lloc de la crida** (com `_AutoFirmaCandidatePaths`), no la de
    `_FindCampInfoPairs` (`,@(...)` consumit **sense** `@()`). Les dues són
    correctes; barrejar-les no.
  - **Les proves ho haurien enxampat** (`AssertEq $portals.Count 5`,
    `AssertEq @(ConvertFrom-CatastroAdXml '').Count 0`), però no s'havien pogut
    executar: a l'entorn no hi ha `pwsh`, i la rèplica en Python **no modela
    aquesta semàntica** — retorna llistes planes i el problema no hi existeix.
    Lliçó: una rèplica en Python valida la LÒGICA, mai les trampes del llenguatge.
- **EL DIAGNÒSTIC ÉS UN `.bat`, i ho és per dues rascades seguides.**
  `suport/rutes/Provar-Cadastre.bat`, doble clic. Els dos intents d'escriure la
  comanda a mà van fallar tots dos:
  1. `. Ruta.ps1` a pèl → `Ruta.ps1` executa la seva `Main` al final i **obre el
     planificador de rutes**. L'usuari es va trobar la finestra oberta sense
     saber què fer.
  2. `powershell -NoProfile -Command "$env:COORDENADES_TEST=1; …"` llançat
     **des d'un PowerShell** → el shell de FORA expandeix `$env:` (buit) abans
     de passar-ho i al de dins li arriba `=1; …` → *«El término '=1' no se
     reconoce»*. Dins de cometes dobles, el `$` és del shell exterior.
  Al `.bat` la variable la posa el `cmd` i a la línia de PowerShell **no hi ha
  cap `$`**. Si algú ja té un PowerShell obert a l'arrel, el que sí que va és
  `$env:COORDENADES_TEST=1; . .\suport\rutes\Coordenades.ps1; Test-Geocodificador '…'`
  (sense embolcallar-ho en un altre `powershell`).
  Regla general: **una comanda de diagnòstic que s'ha d'escriure a mà amb
  cometes niuades no és una comanda de diagnòstic, és un `.bat` que falta.**
- **`Find-HeaderColumn` ha passat de `Precintades.ps1` a `Ruta.ps1`**: és
  utillatge comú de `rutes/` i ara la fan servir dos fitxers. Tot va a dot-source
  al mateix àmbit, o sigui que Precintades la segueix veient.
- La memòria cau dels portals viu a `local/geocodificacio/portals.json` (clau
  `Geocodificacio` a `$Script:LocalSubdirs`) i **es desa cada 25 parcel·les
  noves**: si es cancel·la a mitja tanda, no es perd el que ja s'ha demanat. Les
  entrades amb portals valen 365 dies; les buides, 30 (per si el servei era
  caigut). Si la crida **falla**, no s'hi escriu res.
- **LES SIGLES DE VIA DEL CADASTRE HI HAN DE SER TOTES.** El Cadastre escriu
  `CL CADIS`, i a `Get-ViaNormalitzada` hi faltava **`CL`** (hi havia `CALLE`,
  `CARRER` i `C`, però la `C` demana un espai al darrere i a `CL` la segueix una
  `L`). Resultat: `'CL CADIS' ≠ 'CADIS'`, **cap adreça no ha coincidit mai per
  carrer** i la tria del portal es feia només pel número. En una illa amb
  entrades per dos carrers això vol dir agafar el número del carrer del costat:
  `C HUELVA 1` va acabar a 129 m d'on tocava. La llista viu a
  `$Script:GeoSiglesVia`; si un dia surten desplaçaments estranys, mira primer
  si hi ha una sigla nova. Va passar desapercebut perquè **no falla, empitjora
  en silenci**: seguia trobant un portal, només que el que no era.
- **Dos portals amb el mateix número existeixen.** A la illa de Cadis n'hi ha
  dos amb l'1, i un d'ells cau exactament al centre de la parcel·la. Quan passa,
  `Select-PortalFacana` retorna `facana-dubtosa`: al mapa surt en ambre perquè
  l'usuari el miri, en lloc de triar-ne un a l'atzar i callar.
- **El guard del JSON ha de mirar la SORTIDA, no el `Count`.** El
  `if ($arr.Count -eq 1) { "[$json]" }` dona per fet que `ConvertTo-Json`
  desembolcalla quan hi ha un sol element — i **en el PowerShell de l'usuari no
  ho fa**: sortia `[[{…}]]` i el mapa d'una sola activitat no arrencava. Ara es
  mira `StartsWith('[')`. **`Ruta.ps1` tenia el mateix patró** i s'ha arreglat
  igual (amb una sola parada hauria petat exactament igual).
- **La graella de zones s'ancora a un origen CONSTANT** (`$CoordZonaX0` /
  `$CoordZonaY0`), no al mínim de les dades. Si sortís de les dades, n'hi hauria
  prou que una activitat nova caigués més a l'oest perquè **totes** les zones es
  desplacessin i «la zona C6» volgués dir una altra cosa que la setmana passada.
  400 m: 40 zones, la més gran de 40 activitats i la mediana de 17.
- **Els noms de les zones surten de les dades, no del codi.** `Get-CarrersDominants`
  els calcula dels carrers de les activitats que hi cauen: al codi no hi ha
  escrit cap nom de cap carrer de Cornellà, i per tant no es pot desfasar.
- **`Test-CoordPlausible` i per què cal.** La base porta coordenades impossibles
  (el GIA 1009 té `X=423,37`: les xifres bones dividides per mil). Si es colen, el
  mapa s'estira fins a l'Atlàntic i la resta de punts queden tots en un píxel. Es
  descarten i es llisten al resum del final, mai en silenci.
- **El progrés del repàs viu al NAVEGADOR, i el PowerShell no hi té accés.** Per
  això la finestra de tria diu quantes activitats té cada zona però **no** quantes
  en portes de repassades: això ho diu el mapa, que sí que pot llegir el
  `localStorage`. No intentis posar-ho a la finestra sense una via real de
  retorn del navegador al disc.
- **El `localStorage` desa la FILA SENCERA**, no només la posició: l'Excel ha de
  poder portar tot el que s'hagi validat d'aquesta base, també el de les zones
  que avui no estan obertes. La clau és el **nom del fitxer d'origen**, de manera
  que en canviar de base d'activitats el repàs es buida sol — que és el que toca,
  perquè les correccions ja seran a dins de la base nova.
- **Proves**: `tests/run-tests-coordenades.ps1` (registrada a `run-tests-all.ps1`).
  A l'entorn de desenvolupament no hi havia `pwsh` (ni paquet ni GitHub), o sigui
  que la lògica es va validar amb una **rèplica en Python** — el mateix recurs
  que ja s'havia fet servir en aquest projecte — i les suites de PowerShell les
  ha d'executar l'usuari a Windows.


### Millores pendents de l'eina «Coordenades»

Després de la tanda d'octubre de 2026 (el prompt és `millorar-coordenades.md`,
que ja només porta el que queda). Per benefici / risc:

1. **Provar l'«Excel per importar» amb l'Excel de debò** (punt 11 de
   `provar-al-pc.md`). Benefici alt, risc nul: és l'única peça que escriu per COM.
2. **Fixture gravada del Cadastre** en lloc de la muntada a mà. Cal que
   l'usuari passi un `resposta-<rc>.xml`. Benefici mitjà, risc nul.
3. **Correccions contra una base nova**, sense haver de generar l'Excel.
   Benefici mitjà quan arribi la primera base corregida.
4. **Progrés per zona a la finestra de triar zones**, llegint l'Excel del
   repàs. Canvia el que veu l'usuari: preguntar-ho.
5. **Apilades amb tolerància** i **precisió dels «sense portal»**: necessiten
   dades reals per decidir. Un altre geocodificador vol dir enviar adreces fora:
   preguntar-ho.

**Descartat:** cdnjs com a segon CDN (no és una còpia garantida de npm i el SRI
podria no coincidir; jsDelivr sí que ho és).

## Plànol públic d'activitats precintades
- `suport/rutes/Precintades.ps1` genera `docs/dades/precintades.json` a partir
  de l'Excel d'activitats (fulla "Estès"): les activitats amb el camp lliure
  "PRECINTE ACTIVITAT?" i valor que comença per "SI". La pàgina pública
  `docs/precintades.html` (GitHub Pages) el llegeix i pinta el mapa (Leaflet).
- Ho refresca i puja a `main` **`Actualitzar.bat`** (pas 7). URL pública:
  `https://xexifm.github.io/informes-Cornella/precintades.html`.
- **Privadesa**: el JSON només conté activitat genèrica (p.ex. "BAR"), adreça de
  l'establiment, ID intern i coordenades — **mai** la raó social ni el text
  lliure del Valor (que conté noms i tràmits interns). No hi afegeixis dades
  personals: aquesta pàgina és pública.
- **Etiquetes del plànol**: tant a `precintades.html` com al mapa de ruta
  (`Ruta.ps1`, `Build-RouteHtml`) els marcadors mostren l'**ID Activitat (GIA)**
  en una "pastilla" (no un número correlatiu; l'amplada s'ajusta als dígits). A
  la ruta, la BASE conserva el "0" i s'afegeixen unes quantes **fletxes de
  sentit** (~9, `addRouteArrows`) orientades al traçat. El GIA ja era públic al
  JSON i al popup, així que no hi ha cap dada nova exposada.
- Reutilitza les funcions de `Ruta.ps1` carregant-lo en mode headless
  (`RUTA_TEST`); si canvies `Ruta.ps1`, executa també
  `run-tests-precintades.ps1`.

