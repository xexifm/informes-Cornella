# Llicència (Annex II / LL Prov), MNS i Traspàs

> Ve de `suport/CLAUDE.md`, que s'havia fet massa gros per llegir-lo
> sencer. **Llegeix-lo ABANS de tocar `suport/Llicencia.ps1`, `LlicenciaDb.ps1`,
> `MnsTraspas.ps1` o `ESTRUCTURALS/LLIC.json`**

## Llicència (Annex II / LL Prov) — `suport/Llicencia.ps1` + `ESTRUCTURALS/LLIC.json`
- **No és un informe, són TRES** (`_LlicFases`): *requeriment* (el que es fa el
  95% dels cops), *favorable-pre* («a l'espera de rebre la citada documentació»)
  i *favorable-post* («es dóna per tancat l'expedient»). **Un sol botó** al menú
  (📜, acció `llicencia`, entre *Activitats extraordinàries* i EINES) i la fase
  es tria a dins, amb la casella **«Llicència provisional»**.
- **LLIC NOMÉS HI AFEGEIX; el text és de REQ1, EN VIU.** Aquesta és la decisió
  que ho explica tot: els punts de Llicència **són** els requeriments de REQ1, i
  duplicar-los voldria dir mantenir dues còpies del mateix text. `LLIC.json`
  desa per punt una **`clau`** (`"Secció::Títol"`, el mateix format que
  `_FindItemKeysByTitle` de `ControlsPeriodics.ps1`) i **`cos` buit**; el cos el
  resol `_LlicPuntsPerBloc` contra el REQ1 parsejat (`_LlicIndexReq1`).
  - **Per què `clau` i no el títol a pèl:** si l'usuari reanomena un requeriment
    a REQ1, el lligam s'ha de trencar **dient-ho**. `_LlicPuntsPerBloc` retorna
    `@{Punts; Orfes}` i l'assistent **avisa amb el llistat de claus òrfenes**.
    Empassar-s'ho voldria dir generar un informe al qual li falta un punt sense
    que ningú se n'adoni.
  - Els punts que **no** són a REQ1 (els dos condicionals de compatibilitat i
    l'ANNEX 1) van a la secció `PROPIS` **amb el text sencer** i sense clau.
- **Vocabulari propi de la família `llicencia`** (`_Ed_TipusOptions`):
  `nodisposa` / `sidisposa` (els dos comentaris de cada punt) i `quan` (el
  «Quan:» del bloc DESPRÉS). Seccions: `ABANS`, `DESPRES`, `PROPIS`, `ANNEX 1`.
- **Els comentaris NO van en verd.** Al Word que feia servir l'usuari sí, però
  aquell color era **una marca seva** per veure què havia de canviar a cada
  informe — no forma part del document. Al generat van amb el color de sempre;
  l'única distinció és **negreta al «No es disposa…»** (falta) i normal al «Es
  disposa…». `Format-Body` va guanyar `-Bold` per això (aplicat **al rang**, mai
  amb `$sel.Font.Bold` abans de teclejar: la trampa de `Format-Item`).
- **Classificació**: només als informes de Llicència. `_ClassificacioText`
  (pura, `Activitats.ps1`) munta `"Llei 20/2009; Annex II; Epígraf 12.25"` de les
  columnes `Classificació general annex` / `... Apartat` de l'Excel. Va al
  `<<CLASSIFICACIO>>` de la capçalera.
- **En copiar un bloc de la capçalera, s'han de copiar els FILLS DE PRIMER
  NIVELL, no els paràgrafs.** El requadre de la «Nota: S'ha publicat
  l'Ordenança…» és una **taula** (`<w:tbl>`), no un paràgraf amb vora. La
  primera versió del bloc `[[CAP:LLIC]]` només va copiar els `<w:p>`, i els
  paràgrafs de dins de la taula van sortir-ne: **l'informe perdia el requadre**.
  Ara es recorren els fills de primer nivell del `<w:body>` (`w:p`, `w:tbl`,
  `w:sectPr`) i es copien sencers.
- **`0 CAPCALERA.docx` NO ES TOCA AMB UN SERIALITZADOR D'XML. MAI.** (Error meu,
  gros, i que va deixar el programa inservible uns dies.) La cirurgia del bloc
  `[[CAP:LLIC]]` es va fer amb `ElementTree`, que **reserialitza tot el
  document**: dels **19** espais de noms de l'arrel en van quedar **3**, va
  inventar prefixos `ns0:`/`ns1:` i va perdre el `mc:Ignorable`. Word deia
  *"El archivo parece estar corrompido"* i **no es podia generar CAP informe**
  —ni REQ1, ni TERMINI, ni ACT_EXTR— perquè tots surten d'aquesta capçalera.
  El símptoma va sortir a `Document.ps1:197` (`$word.Documents.Open`).
  - **Com s'ha de fer**: edicions de **TEXT** sobre `word/document.xml` (buscar
    els `<w:p>…</w:p>` com a trossos de cadena i moure'ls), i tornar a escriure
    el ZIP **conservant `compress_type` i l'ordre** de cada entrada. Així només
    canvia `document.xml` i tota la resta queda byte a byte igual.
  - **Proves que ho vigilen** (i s'ha comprovat que fallen si es reprodueix el
    desastre): `mc:Ignorable` present, cap prefix `ns\d+:`, la declaració XML
    original amb `standalone="yes"`, i **≥15 espais de noms** a l'arrel.
- **`0 CAPCALERA.docx` té ara TRES blocs**: el genèric (REQ1/TERMINI), el
  d'`[[CAP:ACT_EXTR]]` i el nou `[[CAP:LLIC]]`, que és una còpia del genèric amb
  la línia `Classificació: <<CLASSIFICACIO>>`. `_CapMarcador` (pura) reconeix
  **qualsevol** `[[CAP:X]]` i `Select-CapcaleraBlock $doc 'LLIC'` es queda el que
  toca (`''` = el genèric), **esborrant primer el tros de baix** perquè els
  índexs dels paràgrafs de dalt no ballin. Si el bloc no hi és, es queda amb el
  genèric i l'informe surt igualment, sense classificació.
- **L'ANNEX 1 només al REQUERIMENT d'una llicència PROVISIONAL**, i el seu text
  **viu al catàleg** (secció que comença per `ANNEX 1`), no encastat al codi:
  així l'usuari el pot editar com tota la resta.
- L'assistent (`Invoke-LlicenciaWizard`) reaprofita el que ja hi ha:
  `Get-HeaderData` per a la capçalera i **`Select-Items` tal qual** per al bloc
  *Projecte* (són els requeriments normals de REQ1: no s'hi inventa cap pantalla).
  `_LlicPuntsDeSeleccio` converteix el que retorna `Select-Items` als mateixos
  punts que la composició, per no tenir **dos camins** de generació.
- `Get-Catalegs` **exclou `LLIC.json`** (no és un catàleg de deficiències: no ha
  de sortir al menú de "Requeriment - Nou") i `_VistaEsProtegit` també el
  protegeix.
- **EL FORMAT EL MANA `Format.ps1`, i el camí és el de REQ1.** La primera prova
  real de Llicència va destapar que `Build-LlicenciaDocument` s'havia inventat
  mitja composició. Ara segueix `_WriteCatalegBody` (`Document.ps1`) al peu de
  la lletra:
  - **URLs**: `_SplitTextAndUrls` + `Format-Url`, mai un detector propi. N'hi
    havia un (`_EsUrl`) fet amb `-like '[[URL]]*'` — i en un patró de `-like`
    `[[URL]` és una **classe de caràcters**, o sigui que no encertava mai: el
    marcador `[[URL]]` sortia **tal qual** a l'informe i l'enllaç no era
    hipervincle. (Tercera aparició d'aquesta mateixa trampa.)
  - **Un enllaç no es repeteix dins d'un punt**: el text de REQ1 i el comentari
    «No es disposa…» solen portar-lo tots dos i sortia dues vegades seguides.
  - **Espais**: `$cfg.SpacerAfterSection` / `-Subsection` / `-Item`, com REQ1.
    No s'hi posa cap `Format-Spacer` a mà.
  - **Numeració**: `Format-Item $sel "N." $text` — amb el **punt**.
  - **Conclusions**: `Format-ConclusionHeader 'CONCLUSIONS'` (centrat i en
    negreta) + `Format-Conclusion`. A REQ1 la negreta de la conclusió ve del
    `**…**` del catàleg; aquí el text el posem nosaltres, així que l'hi afegim.
  - **`FormatDoubles.ps1` ha de tenir TOTES les `Format-*`**: li faltaven
    `Format-ConclusionHeader`, `Format-Note` i `Format-Label`, i per això la
    suite no podia executar la composició sencera.
  - **Prova de generació COMPLETA** (Word simulat + dobles): hauria enxampat de
    cop el `[[URL]]`, els enllaços repetits, la manca de CONCLUSIONS i de
    negreta i el nom del fitxer. Ara hi és.
- **La CLASSIFICACIÓ surt SOLA i mai es pregunta.** La llei la decideix la
  columna **«Classificació general annex»** de l'Excel (`_ClassificacioText`,
  `Activitats.ps1`): el que comença per **`L18`** (`L18 Cert`,
  `L18 Proj i Cert`) → **`Llei 18/2020; Epígraf …`** *sense annex* — el
  «Cert / Proj i Cert» és el tipus de tràmit, no part de la classificació —;
  `II` / `III` → `Llei 20/2009; Annex …; Epígraf …`.
  - La fitxa de la caché és un **hashtable** (`Activitats.ps1` hi desa `@{…}`).
    S'hi accedia amb `$act.PSObject.Properties[…]` → **sempre buit**, i per
    això el programa preguntava cada vegada. Si no se'n troba cap, es deixa
    buida: aturar l'assistent per això seria pitjor que generar l'informe.
- **CADA BLOC ES PORTA SECCIONS SENCERES DE REQ1, i qui ho decideix és el
  CATÀLEG.** Una entrada de `LLIC.json` amb una **`clau` que no és un ítem** —
  una secció (`Instal·lacions`) o una subsecció (`Incendis::Documentació (ITC
  SP)`) — **s'expandeix**: un punt per cada ítem d'aquella part, amb el text
  **literal** de REQ1 i el mateix «Quan:» per a tots.
  - **ABANS** (`_LlicSeccionsAbans`): *Autoritzacions / Informes preceptius* i
    *Registres* — **36 punts**. Aquí LLIC hi aporta el «No es disposa / Es
    disposa» **per ítem**, per clau.
  - **DESPRÉS**: els 6 punts de text propi més **5 seccions expandides** —
    *Incendis / Documentació (ITC SP)*, *Pla d'Autoprotecció*, *Controls
    inicials*, *Controls periòdics* i *Instal·lacions* (**51 punts**).
  - **PROJECTE**: la resta (**68**). Les exclusions surten de
    `_LlicSeccionsExpandides` (llegeix el catàleg), no d'una llista al codi:
    moure una secció de bloc és editar `LLIC.json`, no tocar el programa.
  - **Cap punt de REQ1 surt a dos blocs.** Abans el PAU i dos controls inicials
    sortien a ABANS i a DESPRÉS alhora; hi ha prova que ho vigila.
  - Un requeriment **nou** d'una secció expandida hi surt sol, sense apuntar-lo
    enlloc — que és tot el motiu de fer-ho així.
  - `_LlicEsSeccioAbans` compara **sense accents ni apòstrof tipogràfic**: el
    catàleg escriu `Pla d'Autoprotecció` amb U+2019 i és fàcil que un dia no
    coincideixi caràcter a caràcter.
- **Cap dada d'una activitat concreta al catàleg.** `LLIC.json` portava **sis**
  expedients i referències cremats (`Expedient: FUE-2023-03018882`,
  `Referència 24/2022/000104`, `NIMA: 0801096860`…) que sortien a l'informe de
  **tothom**. Ara són `[CAMP: …]`.
  - **Convenció**: les dades van al **parèntesi del final** de la línia
    `sidisposa`, en peces `Etiqueta: [CAMP: Nom]` separades per `;`. Hi ha una
    prova que recorre `LLIC.json` i **falla si després dels dos punts hi ha
    text que no és un camp** — també vigila el que s'hi escrigui de nou des de
    l'editor de catàlegs. Els dos punts són obligatoris: és el que fa que la
    comprovació sigui fiable.
- **Tots els punts d'ABANS poden dir que es tenen.** El bloc en té 41 i
  `LLIC.json` només en descrivia 15: als altres 26, triar «Es disposa» no
  ensenyava res i a l'informe no s'hi escrivia res. `_LlicTextosPerDefecte`
  (pura) hi posa `No es disposa del document` i
  `Es disposa del document (Id Firmadoc: [CAMP: Id Firmadoc])` quan el catàleg
  no en diu res; els que tenen redacció pròpia la conserven. Un requeriment
  **nou** de REQ1 ja surt utilitzable sense tocar `LLIC.json`.
- **LA MEMÒRIA D'UNA LLICÈNCIA ÉS UN MAPA PLA, no els objectes de camp.**
  `$st.MemAbans[<clau>]` era `@{ Marcat; Estat; Camps; Subs }`, i `Camps` eren
  els objectes vius que fa `_RenderRichInto` — que **no sobreviuen un pas per
  JSON**. Ara és `@{ Marcat; Estat; Valors; Subs }` amb `Valors` = nom → valor,
  i la pantalla el torna a fer servir com a **`$preload`** de `_RenderRichInto`
  (`_GetPreloadValue`, `Camps.ps1`), que és la via que ja feia servir REQ1.
  - En sortir de la pantalla, `Valors` es reconstrueix **fusionant** el que hi
    havia amb el que hi ha als controls: un punt que l'usuari no arriba a
    clicar no té controls pintats, i els valors recuperats es perdrien.
- **La base de dades de llicències** (`suport/LlicenciaDb.ps1`,
  `local\base-dades-llicencies\llicencies-db.json`, ignorada pel git) és el
  mateix patró que el registre d'activitats extraordinàries. Un informe de
  llicència gairebé mai va sol —requeriment → favorable pre → post— i fins ara
  el segon tornava a demanar-ho **tot**.
  - **Es carrega sola** en sortir del pas 2, quan ja se sap l'ID GIA, i
    **només un cop per sessió** (`$st.DbCarregat`): si l'usuari torna Enrere,
    el que acaba d'editar mana.
  - **No toca la capçalera**: TITULAR, ADREÇA i ACTIVITAT les omple l'Excel per
    ID GIA (`Get-HeaderData`), i les de la base poden ser més velles.
  - **Es desa al pas 9**, després de generar, dins d'un `try`: un error desant
    la memòria no pot fer perdre un informe que ja està fet.
  - `ConvertFrom-Json` torna **PSCustomObjects** i converteix les claus
    numèriques dels sub-punts en **text**; `_LlicDbAMapa` i
    `ConvertFrom-LlicenciaMemoria` ho desfan. Hi ha prova d'anada i tornada
    **amb el JSON pel mig**, que és on es trencava.
- **El xip «Dades» del menú.** La fila de Llicència té ara **dos** xips
  clicables: el `✏️ LLIC` de sempre i un de nou que obre la base de dades
  (`Extra` a l'entrada del menú → `ExtraChipRect`, mateix hit-test i hover que
  `DocChipRect`; acció `llicdb`).
- **`LLIC.json` TÉ vista en Word** (`_VistaLlicencia`, `VistaWord.ps1`): abans
  era l'únic catàleg sense, i no es podia consultar fora del programa. Ensenya
  cada bloc amb **tots** els seus punts ja resolts contra REQ1 i, en cursiva, el
  que hi afegeix (`[No es disposa]`, `[Es disposa]`, `[Quan]`), més el resum del
  PROJECTE i l'ANNEX 1. Surten de **`_LlicPuntsPerBloc`**, la mateixa funció que
  munta l'informe: la vista no pot dir una cosa i el document una altra.
  `_VistaEsProtegit` ara només protegeix `0 CAPCALERA.docx`.
- **La base de dades edita TOT el que es recorda**, en l'ordre de l'informe:
  primer **Documentació PROJECTE** (tècnic redactor, núm. col·legiat, col·legi,
  data i els documents signats amb el seu Id Firmadoc), després els blocs ABANS
  i DESPRÉS i, al final, la resta. Abans la documentació del projecte sortia al
  final i **com a text**: no s'hi podia tocar res. Els camps de cada punt surten
  del **text del catàleg** (`Get-LlicenciaPuntsEditables`), no només dels valors
  ja desats — per això abans només es podia editar «Compatibilitat», l'únic que
  en tenia.
- **La base de dades no ensenyava res** (pantalla en blanc): `.Selected = $true`
  **no mou el `CurrentRow`**, i com que la fila ja sortia seleccionada, clicar-la
  **no disparava `SelectionChanged`** — el detall no es pintava mai. Ara es posa
  el `CurrentCell`, hi ha `add_CellClick` a més del canvi de selecció, i es
  repinta al `Shown` (abans de mostrar la finestra el `CurrentCell` encara pot
  ser `$null`). El detall mostra també els punts del projecte, el tècnic, les
  condicions i l'historial: només amb els blocs semblava buit.
- **La pantalla de documentació és LLISTA + DETALL**, com `Select-Items`, no una
  graella. Els `[CAMP: …]` es pinten **inline amb `_RenderRichInto`**
  (`Camps.ps1`) — la mateixa funció que REQ1 —, amb un **diccionari de camps
  per punt** (l'«Id Firmadoc» val una cosa diferent a cada document). El botó
  «Omplir…» que obria un diàleg s'ha esborrat: no s'assemblava a res de la
  resta del programa. Al bloc DESPRÉS, el detall mostra les **caselles dels
  sub-punts** (els certificats d'inscripció i les inspeccions inicials: no
  totes les activitats els tenen tots).
- **L'ENLLAÇ VA DESPRÉS DE LA FRASE QUE L'ANUNCIA.** El comentari acaba amb
  «…en el següent enllaç:» i el cos de l'ítem (de REQ1) **sol portar el mateix
  enllaç**: sortia abans i la frase quedava penjada. `_LlicEscriuPunt` mira
  **primer** els enllaços del comentari; els que també són al cos **no
  s'emeten amb l'ítem** i surten després del comentari. Cap enllaç es repeteix
  dins d'un punt. (La deduplicació a seques, sense això, ho **empitjorava**:
  esborrava justament el que havia d'anar després de la frase.)
- **La NUMERACIÓ de l'informe va seguida de cap a peus**: el bloc DESPRÉS no
  reinicia el comptador (`$n` no es torna a posar a 0). Sortien dues llistes que
  començaven per 1 totes dues.
- **Els sub-punts d'INSTAL·LACIONS surten de REQ1, no d'una còpia a mà.**
  «Certificats d'inscripció al RITSIC» i «Inspecció inicial» tenien 10 i 4
  subítems escrits a `LLIC.json` mentre la llista de debò (17 i 5) viu a REQ1.
  Ara la seva **`clau` apunta a una SUBSECCIÓ** (`Instal·lacions::Legalitzacions`,
  `Instal·lacions::Inspeccions inicials`) i `_LlicPuntsPerBloc` en treu els
  sub-punts amb `_LlicItemsDeSubseccio` (pura).
  - **Una clau pot apuntar a un ítem O a una subsecció.** Si és una subsecció, el
    **cos** és el del punt de LLIC (la frase que encapçala la llista) i els
    **sub-punts** són els ítems de la subsecció. Per això aquests —i només
    aquests— sí que porten text propi.
  - `_LlicResumSubpunt` (pura) es queda amb **el text fins a l'enllaç, i
    l'enllaç**: a REQ1 la primera línia és l'etiqueta i el que ve després ja és
    el requeriment sencer. Retorna **línies** (text + `[[URL]] …`), que és el que
    `_LlicEscriuPunt` ja sap emetre com a pic + hipervincle. Decisió de
    l'usuari, vista la comparació amb «només la primera frase» —que deixava
    tres ítems de gas idèntics—.
  - **Aquelles dues subseccions desapareixen del pas «Projecte»**
    (`_LlicSubseccionsFora` + `_LlicSeccionsSenseSubseccions`, pures): no es
    poden demanar dues vegades. *Inspeccions periòdiques* s'hi queda.
  - **`Instal·lacions` és l'última secció** del bloc, i per això el punt
    `Insp. periòdica - PCI` s'ha mogut al final de la llista de `LLIC.json`:
    l'ordre de la pantalla i el de l'informe són el mateix.
- **El pas «Projecte» de Llicència accepta no triar res** (`Select-Items
  -permetreBuit`): hi ha activitats sense cap deficiència de projecte, i
  l'avís «No s'ha seleccionat cap deficiència» hi bloquejava l'assistent. A
  «Requeriment - Nou» l'avís es queda: allà un informe buit no té sentit.
- **La CLASSIFICACIÓ no va en negreta.** A la capçalera, l'etiqueta és en
  negreta i el **valor** no —així són totes les línies (`ID GIA:`, `Titular:`…)—,
  però el bloc `[[CAP:LLIC]]` es va escriure amb `Classificació: <<CLASSIFICACIO>>`
  en **un sol run** amb `<w:b/>`. Ara són tres runs, com a `Titular:`: etiqueta
  en negreta, tabulador, i el valor amb el mateix `rPr` **sense** `<w:b/>`. Hi
  ha prova sobre el `.docx` real.
- **L'ANNEX 1 va en TEXT PLA**: `Format-Plain` (`Format.ps1`) — sense sagnia,
  pics ni numeració. **Els números i els guions van escrits com a TEXT**
  (`1. `, `2. `… als `item`; `- ` als `subitem`), perquè la plantilla els porta
  amb numeració automàtica del Word i l'usuari els vol plans i sense sagnia. El
  comptador només avança amb els `item`; els `text` del mig no es numeren, i al
  full de signatures no s'hi posa cap marca. **Negreta només** als dos títols, i des del «Document
  d'acceptació…» (`_LlicEsTitolAcceptacio`, pura) **pàgina nova i cos 9**
  (`$Script:LlicAnnexSignaturaCos`; a la plantilla, `sz=18` mig-punts).
- **La CLASSIFICACIÓ no és a la capçalera genèrica**: `_ReadHeaderControls` no
  la retorna (és només de Llicència). S'omple després del pas 2 amb
  `_LlicClassificacio`: es busca a l'Excel per ID GIA i **sempre es mostra** en
  un quadre per confirmar-la o corregir-la. Sense això sortia una línia
  «Classificació:» **buida** a l'informe.
- **El nom del fitxer no porta el titular** (decisió de l'usuari): ja surt a la
  capçalera del document.
- **Els comentaris «Es disposa» demanen DADES, i són per punt.** Al Word de
  l'usuari hi havia `XXX` on va l'Id Firmadoc i, segons el punt, l'Expedient /
  Referència / Registre. Al catàleg són ara **`[CAMP: nom]`**, i al bloc ABANS
  els camps es pinten **inline** amb `_RenderRichInto` (`Camps.ps1`), la mateixa
  funció que REQ1. (Abans hi havia una columna «Omplir…» que obria un diàleg,
  `Select-LlicDadesPunt`; es va esborrar en passar la pantalla a llista+detall i
  ara la funció tampoc no hi és.)
  - **No es fa servir el diccionari de camps compartit**: allà la clau és el
    NOM del camp, i «Id Firmadoc» val **una cosa diferent a cada document**. El
    valor es resol punt a punt (`_LlicAplicaCamps`, pura) i el punt se'n va amb
    el text **ja resolt**. Un camp sense valor deixa el forat buit, mai el
    marcador a la vista.
  - Al bloc **DESPRÉS** es marca si es disposa del document, però **no** es
    demanen les dades (`$ambDades` només a ABANS): l'usuari ho va demanar així.
  - Prova al catàleg real: **cap `XXX`** i **tot «Es disposa» demana com a
    mínim l'Id Firmadoc**.
- **La tria de cada pantalla es RECORDA** (`$st.MemAbans` / `$st.MemDespres`,
  indexada per `_LlicClauPunt` = clau de REQ1, o `#títol` per als punts propis):
  tornar Enrere ja no esborra el que s'havia marcat. Es recorda **tot** —marcat
  o no, l'estat i les dades—, no només el que estava marcat.
- **L'assistent NO tenia `catch`.** Qualsevol error a dins **matava el programa
  en silenci** («es tanca i no passa res, tampoc es genera cap informe»).
  `Invoke-NouWizard` sí que en té des de sempre; aquest se'l va deixar. Ara
  mostra el missatge **i el fitxer i la línia**, i torna al menú.
- **LA MATEIXA TRAMPA TÉ UNA SEGONA CARA: en ARGUMENTS d'una crida.**
  `_AddBrandHeader $form 'X' 'refer' + [char]0x00E8 + 'ncia' 56` no concatena
  res: PowerShell passa `'refer'`, `'+'`, `'è'`, `'+'`, `'ncia'` i `56` com a
  arguments **solts**, i el `56` (l'alçada) acaba a un altre paràmetre. El
  programa peta amb *"no se puede convertir el valor '+' al tipo System.Int32"*
  i, com que és una excepció de WinForms, surt el quadre gros de .NET. Va
  passar al botó «Omplir…» de Llicència.
  - **Detector definitiu, sense falsos positius**: es recorre l'AST de tot
    `suport/` i es busca un argument que sigui **literalment `'+'`** — això
    només pot venir d'una concatenació sense parèntesis. Ho diu el propi
    parser, no una expressió regular. Hi ha prova, i s'ha comprovat que
    **falla** quan s'hi injecta el cas (i diu fitxer i línia).
  - Regla pràctica: **qualsevol concatenació que vagi com a argument, entre
    parèntesis**. Les dues cares d'aquesta trampa (dins d'un `@()` i en
    arguments) ja han costat dues rondes amb l'usuari.
- **TERCERA CARA: `f -or $x` NO ÉS UN «O», és una CRIDA a `f`.** Una funció
  sense parèntesis obre una **crida a una ordre**, i tot el que ve darrere en
  són **arguments**: `_GoldenEsRefer -or -not (Test-Path $p)` li passa `-or`,
  `-not` i el resultat del `Test-Path` com a paràmetres, la funció els ignora i
  retorna el seu valor de sempre → **la condició sempre és certa**. Ho vaig
  patir escrivint els fitxers d'or: el codi anava directe a la branca de
  «llegeix el fitxer» amb un fitxer que encara no existia.
  - **Regla**: una crida a funció dins d'una condició, **entre parèntesis**:
    `(f) -or (-not (Test-Path $p))`.
  - És la mateixa família que el `'+'` solt: PowerShell té **dos modes de
    parseig** (expressió i ordre), i el que decideix quin s'aplica és com
    comença el token. Quan una condició es comporta al revés del que diu,
    sospita d'això abans que de la lògica.
- **Els camps es resolen per BLOC, no línia a línia** (`Apply-FieldsToLines`,
  `Camps.ps1`). Cada paràgraf del catàleg és una **BodyLine**, i l'editor deixa
  prémer Enter **a dins** d'un `[OPCIO: …]`. Llavors el marcador queda partit
  entre dues línies, cap de les dues en té un de sencer, i `Apply-Fields` línia
  a línia no hi trobava res: **el `[OPCIO: …]` anava al Word tal qual**. Va
  passar de debò al punt 1 del Req1 del GIA 1463.
  - **El que despistava**: la *detecció* sí que funcionava, perquè
    `Get-FieldsFromSelection` ajunta les BodyLines. El desplegable sortia bé a
    la pantalla i no es veia res estrany fins a obrir el document.
  - Ajuntar amb `\n`, resoldre i tornar a partir arregla **les dues cares**: el
    marcador partit es resol, i un salt de línia **dins del valor triat** torna
    a sortir com a **paràgraf propi** al Word.
  - Per això `Get-FieldsFromSelection` i `_RichTextOfBodyLines` ajunten amb
    **salt de línia i no amb espai**: si la pantalla i el generador no veuen el
    mateix text, el valor triat al desplegable no és el que s'escriu.
- **Una opció BUIDA és una opció.** `[OPCIO: Afegitó? | | text]` vol dir «res o
  aquest text». `_ParseOpcio` les llençava (`if ($o -ne '')`), o sigui que
  l'afegitó sortia **sempre** i no hi havia manera de dir que no. Ara hi entra,
  i com que sol anar primera també és el valor per defecte —que és el que toca—.
  - Al desplegable es pinta **`(res)`**: una fila en blanc no es distingeix d'un
    desplegable trencat.
  - I el `ComboBox` **es llegeix per ÍNDEX**, no pel text de la fila:
    l'etiqueta és només per veure-la (col·lapsa els salts perquè càpiga en una
    fila) i el valor que es desa ha de ser el del catàleg, salts inclosos.
- **`.GetNewClosure()` copia VALORS, no referències.** Un scriptblock que es
  crida a si mateix — `$pinta = { ... & $pinta $idx ... }.GetNewClosure()` — es
  queda amb `$pinta = $null`, perquè quan es va crear encara no s'havia
  assignat. Al clic peta amb *«l'expressió que segueix a `&` … no és un nom
  d'ordre ni un scriptblock»*, i com que ve d'un handler de WinForms surt el
  quadre gris de .NET. Va passar en triar «Es disposa del document».
  - **La variant silenciosa és pitjor**: `$cerca = $null` … `{ $cerca.Text
    }.GetNewClosure()` … `$cerca = _AddSearchBox …`. La closure es queda amb
    `$null`, `$null.Text` **no peta**, i el cercador simplement **no filtra
    mai**. Ningú se n'assabenta.
  - **Solució**: un hashtable de funcions. `$fn = @{}` declarat primer i
    després `$fn.Pinta = { … & $fn.Pinta … }.GetNewClosure()`: el hashtable
    **sí** es captura per referència, i `$fn.Pinta` es resol en cridar-lo. Si
    el handler ja rep el control (`add_TextChanged` passa el `$sender`), millor
    encara: no cal capturar res.
  - Un scriptblock **sense** `.GetNewClosure()` no té el problema (resol en
    temps d'execució): `$propagate` de `SeleccioItems.ps1` és recursiu i va bé.
  - **Detector, sense falsos positius**: es recorre l'AST de tot `suport/` i es
    busca un `.GetNewClosure()` que referenciï **la variable a què s'està
    assignant**. Hi ha prova, i s'ha comprovat que **falla** quan s'hi injecta
    el cas.
- **LA SEGONA CARA: una closure DINS d'una altra perd el que ve de fora.**
  `.GetNewClosure()` copia **només els locals del context que la crida**. Una
  closure creada a dins d'una altra es queda amb els locals d'aquella invocació
  —paràmetres i variables que s'hi assignen— i **perd tot el que venia del
  mòdul de la closure exterior**:

  ```powershell
  $fn = @{}; $fora = 'x'
  $fn.Pinta = { param($idx)
      $local = 'l'
      $inner = { "$local $idx $fora $fn" }.GetNewClosure()   # $fora i $fn: NULL
  }.GetNewClosure()
  ```

  Vaig arreglar la primera cara posant les funcions de la pantalla en un
  hashtable `$fn`… i el handler dels radios, que es crea **a dins** de
  `$fn.Pinta`, seguia rebent `$fn` buit: `& $null.Pinta` en triar «Es disposa
  del document». **El detector de la primera cara no ho veu**, perquè `$fn` no
  és la variable que s'està assignant.
  - **Solució**: una **còpia local** just abans de crear els handlers
    (`$fnAquest = $fn`). Els paràmetres i les variables assignades dins de la
    funció (`$idx`, `$e`, els controls) sí que arriben bé.
  - **Detector propi**: una `.GetNewClosure()` imbricada dins d'una altra només
    pot referenciar variables **locals de la de fora** (paràmetres,
    assignacions, variables de `foreach`) o seves. S'exclouen les automàtiques i
    les de `$Script:`/`$Global:`/`$env:`. Sobre el codi d'avui: 1 encert i **0
    falsos positius**; validat injectant el cas.
- **UN EMOJI ASTRAL AMB `[char]` I EL PROGRAMA NO S'OBRE.**
  `Icon = [string][char]0x1F5C2 + [char]0xFE0F` → *«no se puede convertir el
  valor 128450 al tipo System.Char»*. Un `[char]` és de **16 bits** (màxim
  `0xFFFF`) i `0x1F5C2` (🗂) no hi cap. Com que aquell codi és a `Select-Mode`,
  que construeix la **primera** pantalla, el programa **no arrencava gens**.
  - La manera bona ja era **dues línies més amunt** al mateix fitxer:
    `[System.Char]::ConvertFromUtf32(0x1F5C2)`. Els emoji del pla bàsic
    (`✏️` = `[char]0x270F + [char]0xFE0F`) sí que van amb `[char]`.
  - **Per què cap prova ho va veure**: `Select-Mode` és WinForms i les proves no
    el criden mai. Cap prova de comportament pot enxampar això.
  - **Detector**: es recorre l'AST de tot `suport/` buscant un `[char]` amb una
    constant més gran que `0xFFFF`. Zero falsos positius —un `[char]` així no
    pot ser correcte MAI— i validat injectant el cas.
  - **Lliçó general**: el codi que només corre a la interfície (menú, diàlegs)
    no el toca cap prova, o sigui que **hi ha d'arribar el parser**. Els tres
    detectors AST que hi ha (el `'+'` solt, les closures i aquest) són tots del
    mateix estil i per la mateixa raó.
- **CONTROLS QUE ES TREPITGEN: ho ha de veure el PROGRAMA, no jo.** És el
  defecte recurrent d'aquest projecte —coordenades a mà, i un text que creix o
  un control nou que hi passa per sobre—. L'últim: el xip «Dades» tapant el
  «LL Prov» del menú.
  - **Al menú, arreglat per construcció**: el `$paintHandler` calcula **primer**
    els xips i després dibuixa el títol dins d'un **rectangle acotat** pel xip
    més a l'esquerra, amb `EndEllipsis`. Digui el que digui el títol i hi hagi
    els xips que hi hagi, no s'hi pot posar a sobre. (Abans es dibuixava en un
    `Point`, sense límit d'amplada.)
  - **A tot arreu, comprovació en viu**: `_NewForm` (`UiComuns.ps1`) enganxa un
    `Shown` que crida **`_AvisaSolapaments`** → recorre l'arbre de controls i
    compara **només els germans** (dins d'un contenidor les coordenades són
    relatives a ell). Si en troba, ho diu **amb els noms i les coordenades**, un
    sol cop per pantalla i sessió.
  - **La geometria és pura i es prova a Linux** (`_TrobaSolapaments`). NO és
    solapament que un contingui l'altre del tot (és un fons) ni que es toquin
    per la vora. S'ignoren els `Dock` (els col·loca WinForms) i els invisibles.
  - **Per què no un analitzador estàtic**: es va provar. Resol els controls amb
    `Location`/`Size` constants i les variables enteres seqüencials (`$y = 76`,
    `$y += 32`), i sobre el codi d'avui troba **zero** — perquè el que falla de
    debò és text pintat a mà (no són controls) i codi dins de bucles. La
    comprovació en viu sí que ho veu tot.
- **Un clic en un RadioButton dispara DOS esdeveniments**: el que es marca i el
  germà que es desmarca. Amb un handler compartit la pantalla es repintava dues
  vegades, i la segona llegia uns controls que `Controls.Clear()` acabava de
  treure. Un handler per radio amb `if (-not $rb.Checked) { return }` i s'acaba.
- **Els errors dels handlers de WinForms no arriben al `try/catch` de qui va
  obrir la finestra**: els para el bucle de missatges, que ensenya el seu quadre
  gris en l'idioma del Windows i sense dir on ha passat. Per això `Motor.ps1`
  posa `SetUnhandledExceptionMode('CatchException')` + `add_ThreadException` i
  els mostra amb **fitxer i línia**, com la resta del programa.
- **La pantalla de documentació és un ARBRE per seccions** (com el Pas 3), no
  una llista plana: 43 punts tots a la mateixa alçada no es podien llegir.
  L'agrupació (`_LlicAgrupaPunts`, pura) surt de la **clau** del punt
  (`Secció::Ítem`, `_ItemKey`) — no cal cap camp nou a cap JSON —, i els punts
  sense clau (els PROPIS, i els llegits d'un informe ja emès) van al **primer
  nivell, sense capçalera**. És **només de pantalla**: l'informe es munta
  recorrent els punts en l'ordre del catàleg, i hi ha prova que agrupar no
  perd, no duplica i no reordena res.
- **La trampa de la coma, altre cop i al meu propi codi**: `@('Pl' + [char]0x00E0 + 'nols')`
  són **TRES** elements, i a la pantalla hi sortien cinc caselles —Projecte,
  `Pl`, `à`, `nols`, Annexos— en lloc de tres. Per això la llista és ara una
  funció pura (`_LlicDocsSignats`) **que es pot comptar en una prova**: la
  regla ja era a aquest document i no va evitar res; la prova sí.
- **El favorable POST LLEGEIX el pre-llicència**, no el torna a demanar
  (`_LlicPuntsDelDocxAnterior`, pura): el post diu «Després d'haver comprovat la
  següent documentació presentada:» i ha de llistar **exactament** el que deia
  l'informe anterior. Fer-ho triar altre cop era repetir una feina que ja consta
  escrita i obrir la porta que les dues llistes no quadressin.
  - **Es reconeix pel TEXT**, sense mirar la numeració del Word: aquests
    informes els genera aquest mateix programa i allà el número i el pic
    s'**escriuen com a text** (`Format-Item` teclegia `"N. "`, `Format-Bullet`
    `U+2022` + tabulador). Per tant `^\d+\.\s` = punt nou, `U+2022` = sub-punt,
    la resta = línia de cos.
  - El **«Quan:» es descarta**: al post la documentació ja s'ha presentat i el
    termini no hi pinta res.
  - El bloc es talla amb `_LlicFinalsDeBloc` (CONDICIONS, ANNEX 1, «Ho poso al
    seu coneixement», la conclusió...). Si no se'n treu res, s'avisa i es cau a
    la llista del catàleg: l'informe s'ha de poder fer igualment.
- **Les caselles surten MARCADES al bloc DESPRÉS** (`Select-LlicDocumentacio
  -marcatPerDefecte`), i al bloc ABANS no. El Word de l'usuari portava tots els
  punts del DESPRÉS i ell hi anava esborrant el que no tocava; picar quinze
  caselles cada vegada era el que feia que l'eina no compensés. A ABANS cada
  punt demana a més decidir «No es disposa / Es disposa», que sí que és una
  decisió per activitat. Hi ha **«Marcar-ho tot» / «Desmarcar-ho tot»**.

## ELS TRES INFORMES DE LLICÈNCIA SÓN EL MATEIX DOCUMENT
Això va costar de veure i és el canvi més gros d'aquesta família. El favorable
**POST** estava fet com un informe **curt** que **llegia** el pre-llicència i en
copiava el bloc DESPRÉS. L'usuari va enviar el que fa a mà i és **l'informe
sencer**: documentació del projecte, bloc ABANS i bloc DESPRÉS amb els seus
«Quan:». L'única cosa que canvia entre fases és **què es diu de cada punt del
bloc DESPRÉS**:

| Fase | Sota el «Quan:» |
|---|---|
| `requeriment` | res (encara no toca dir si es té o no) |
| `favorable-pre` | **No es disposa de la documentació.** (negreta: falta) |
| `favorable-post` | Es disposa del document (Id Firmadoc: …) |

- Ho decideix **`_LlicEstatDespres`** (pura) i ho aplica **`_LlicPuntsAmbEstatFase`**
  al **pas 3** de l'assistent, de manera que **la pantalla del pas 7 i el
  document fan servir exactament els mateixos punts**. Si es fes només a la
  composició, la pantalla no ensenyaria el que sortirà.
- Un punt que ja porti text propi al catàleg **se'l queda**: el genèric només hi
  entra quan no n'hi ha.
- **El «Quan:» va ABANS del comentari** (`_LlicEscriuPunt`): a l'informe fet a mà
  cada punt diu primer quan s'ha de tenir i després si es té. Al revés, el
  termini quedava penjat al final del punt.
- **Fora**: `_LlicEntradaPost`, `_LlicPuntsDelDocxAnterior` i `_LlicFinalsDeBloc`
  — només servien per al post que llegia el `.docx` anterior. `Select-PreviousReport`
  es queda: la fa servir *Seguiment*.

### Altres coses que va destapar comparar el generat amb el fet a mà
- **La documentació del projecte va DALT DE TOT**, sota el títol de secció
  `DOCUMENTACIÓ PROJECTE` i **fora de la numeració**. Sortia al final del bloc
  ABANS, numerada i sota un subtítol subratllat «Documentació».
- **La classificació no porta tabulador**: l'etiqueta és més llarga que la
  primera parada de tabulació i el valor saltava a la següent, molt a la dreta.
- **ANNEX 1**: una línia en blanc entre punts numerats; l'aclariment d'un punt va
  **dins del mateix paràgraf** (`Format-Append`, nou a `Format.ps1`, que és
  l'únic `Format-*` que **no** comença amb `TypeParagraph`); i el full de
  signatures porta dues línies en blanc davant de cada declaració.
- **L'ANNEX 1 no s'escriu si ja es disposa de l'autorització d'usos i obres
  provisionals** (`_LlicCalAnnex1`): l'annex diu **com demanar-la**, o sigui que
  si ja es té no pinta res. El punt es reconeix per la **condició `provisional`**,
  no pel títol (que es pot reescriure des de l'editor i trencaria el lligam en
  silenci).
- **La pantalla «Documentació» recorda què s'ha signat**: Projecte / Plànols /
  Annexos i el seu Id Firmadoc no es desaven enlloc i calia tornar-ho a marcar a
  cada informe. Ara van a `$st.TecnicDocs` i a la base de dades de llicències
  (`ConvertTo-LlicenciaDocs`), i la pantalla surt ja marcada.

## Modificació NO Substancial i Traspàs (`suport/MnsTraspas.ps1`)
Dos informes **curts** que van al **mateix menú** que els tres de sempre (pas 1
de Llicència) perquè comparteixen capçalera i tràmit, però el document no
s'assembla gens: tres o quatre paràgrafs fixos i cap bloc de documentació.
`_LlicTotesLesFases` ajunta `_LlicFases` + `_MnsFases`.

- **L'única cosa que es tria és si hi ha observacions.** Amb observacions →
  «…amb la següent observació:» i **un paràgraf de llista de Word buit**; sense →
  «…sense més observacions en relació a aquest tràmit.» i cap llista.
- **`Format-ListItem` (`Format.ps1`) és una llista de Word DE VERITAT**
  (`ListFormat.ApplyNumberDefault()`), no un número escrit com a text. A la resta
  de l'informe el número s'escriu perquè el document ja surt fet; aquí el
  paràgraf surt **buit** i l'usuari hi escriu al Word, i vol que en prémer Enter
  la llista continuï sola. Perquè el paràgraf següent **no** continuï la llista,
  `_Apply-Indent` fa `ListFormat.RemoveNumbers()` **abans** de posar la sagnia
  (el Word, en treure la numeració, també toca la sagnia).
- **El text viu a `ESTRUCTURALS/MNSTRAS.json`** (un sol catàleg per als dos, com
  va demanar l'usuari), família `mnstraspas`. Una secció per informe i un fill
  per paràgraf: `text` → paràgraf normal, `item` → paràgraf de llista buit. La
  **clau** diu quan hi entra: `amb-observacions`, `sense-observacions`,
  `llista-observacions`, o res = sempre (`_MnsNodeEntra`, pura).
- **El vermell del Word original era una marca de l'usuari**, no part del
  document: ni els títols «MODIFICACIÓ NO SUBSTANCIAL» / «TRASPÀS» ni les dues
  variants es pinten de cap color. Els títols **no s'escriuen**: només serveixen
  per saber quin informe és.
- `MNSTRAS.json` queda **fora de `Get-Catalegs`** (com `LLIC.json`) i té vista en
  Word (`_VistaMnsTraspas`), que ensenya els dos informes amb les dues variants.

