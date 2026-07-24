# Notes per a Claude (mantenir entre sessions)

## Desplegament de l'usuari
- L'usuari executa el programa des d'un **clone de git local** al seu PC.
- Per actualitzar fa doble clic a **`Actualitzar.bat`** (fa `git pull` de
  `main`); per executar el programa fa doble clic a **`GenerarInforme.bat`**.
- Per instal·lar en una **màquina nova** hi ha **`suport/Instalar.bat`**:
  instal·la Git (winget o descàrrega directa), clona el repo a `main` (o
  converteix un ZIP extret en clone), i deixa el programa operatiu. La URL
  pública del repo està fixada al `.bat`: `https://github.com/xexifm/informes-cornella`.
- La branca **estable i de desplegament és `main`**. El clone de l'usuari
  segueix `main`, i `main` és la branca per defecte del repositori.

## IMPORTANT: tota la feina ha de convergir a `main`
Cada sessió de Claude Code (web) treballa en una branca pròpia `claude/...`.
Si la feina es queda només en aquesta branca, **l'usuari no la rebrà mai**
amb el seu `git pull` de `main`, i semblarà que "no ha canviat res".

Per tant, **al final de cada sessió**:
1. Assegura't que tot està commitejat a la branca de la sessió.
2. Fusiona la feina a `main` i fes push de `main`:
   ```
   git fetch origin
   git checkout main
   git pull --ff-only origin main
   git merge --no-ff <branca-de-la-sessio>
   git push origin main
   ```
   (o, si la branca de sessió ja conté `main`, n'hi ha prou amb
   `git push origin <branca-de-la-sessio>:main`).
3. Confirma a l'usuari que ja pot actualitzar amb `Actualitzar.bat`.

Si tens dubtes sobre si pots fer push a `main`, pregunta-ho; però el model
de desplegament de l'usuari depèn que la feina arribi a `main`.

## Base d'informes (informes-db.json)
- El motor de la base d'informes és `suport/Informes.ps1`: escaneja `$InformesDir`
  (per defecte `...\5.- Sergi Fadurdo\Informes`) i, per cada informe (`.docx` o
  `.doc` antic amb data al principi del nom), en treu **data + ID GIA +
  conclusió**, agrupat per activitat (per GIA; si no en té, per **carpeta**), a
  `BASE DE DADES ACTIVITATS\informes-db.json` (gitignored). Botons al menú:
  **🗃 Actualitzar** i **📋 Editar** (marc "Base d'informes").
- Lectura de `.docx` **sense Word** (zip) reutilitzant les primitives de
  `Seguiment.ps1` (`_LoadDocxXml`, `_ParagraphTextXml`). Lectura de `.doc`
  antics (Word 97-2003) via **Word COM** (`_ReadDocParagraphsWord`): instància
  creada mandrosament a `Invoke-InformesDbScan` només si cal reprocessar algun
  `.doc`, i tancada (`Quit()`) en un `finally`. Funcions de text PURES (dates,
  GIA, expedient, conclusió) amb tests a `run-tests.ps1`.
- **ID GIA:** cadena document → carpeta ("GIA 361") → Excel per expedient.
  `_ExtractIdGia` ignora placeholders com `"-"`, `"XXX"`, `"N/A"` (activitats
  encara sense GIA assignat) perquè no s'ajuntin activitats diferents sota una
  mateixa "activitat" fantasma.
- **Conclusió:** `$Script:ConclusioStartPhrases` a `Informes.ps1` llista les
  frases d'inici reconegudes, cada una amb el seu `Font` (família de tràmit).
  `"Vist l'anterior"` i `"Tenint en consideració el risc"` es consideren
  fiables (Font `vist_anterior`/`risc`, decisió pròpia i diferenciada de cada
  informe). `"S'informa favorablement"` (MNS) i `"El titular/L'organitzador és
  responsable d'executar"` (actes extraordinàries) també es capturen i es
  desen al `informes-db.json`, però com que són clàusules gairebé idèntiques
  entre informes diferents, `_ConclusioIgnorarPerDefecte` fa que
  Get-InformeData marqui l'informe **"ignorat" PER DEFECTE** (només la
  primera vegada que es veu; si l'usuari el desmarca des de l'editor, el seu
  criteri es conserva als escanejos següents). Si cap frase coneguda hi
  apareix, la conclusió queda buida (motiu `"sense conclusio"`, va a
  "a_revisar").
- Validat contra la carpeta REAL d'informes (~43 GB, 720 informes): 0 grups
  GIA corromputs per placeholders, cobertura de conclusió 70% → 87%.
- **Conclusió breu / Estat actual:** cada informe té una `conclusio_breu`
  (`_ConclusioBreu`, funció pura) que classifica el TEXT de la conclusió (no
  el nom de l'arxiu, que l'usuari ha anat modificant amb el temps de manera
  inconsistent) en una de `$Script:ConclusioBreuOpcions`: Requeriment, FI
  Requeriment (inclou "denúncia tancada"), Precinte / Cessament, FI Precinte /
  Cessament, Favorable, Ampliació termini, Sense efecte, Altres, Revisar.
  `Revisar` és el resultat per defecte quan no es reconeix cap frase — inclou
  deliberadament "desfavorable" (per no confondre'l amb "Favorable"). `Altres`
  és NOMÉS una opció manual des de l'editor; el classificador automàtic mai
  la retorna. Cada ACTIVITAT té un `estat_actual` (`_EstatActualActivitat`,
  funció pura) = `conclusio_breu` del seu informe **no ignorat** més recent
  **per `data`** (no per data de modificació del fitxer). A **Editar base
  d'informes** la columna "Conclusio breu" és un desplegable editable
  (`DataGridViewComboBoxColumn`) i "Estat activitat" és només lectura,
  derivada; en editar "Ignorar" o "Conclusio breu" de qualsevol informe es
  recalcula i es propaga l'estat a totes les files de la mateixa activitat.
  Una conclusio que diu que **NO** es pot donar per tancat/finalitzat (qualsevol
  "no es pot donar...") es **Requeriment** (pendent), no "FI Requeriment": la
  comprovacio del "no" va abans que la del "si" a `_ConclusioBreu`.
- **Menú Pas 1 — 3 apartats de rajoles** (`Select-Mode`, `Seguiment.ps1`, helper
  `$addTileRow`; dispatch al `switch` de `Main`):
  - **EINES** (3): 📍 *Generar ruta* (`ruta`), 🔒 *Activitats precintades* (url),
    📅 *Controls periòdics* (`controlsperiodics`).
  - **INFORMES** (5): 🗃 *Actualitzar base* (`informesdb`), 📋 *Editar base*
    (`informesdbedit`), 📁 *Copiar informes* (`copiarinformes`), ✅ *Comprovar
    Excel* (`comprovarexcel`), 📄 *Word a PDF* (`convertirpdf`).
  - **MÒBIL** (2): 📧 *Textos del correu* (`emailtextos`), 📥 *Revisar mòbil*
    (`revisarmobil`).
  Sota
  *Actualitzar base*, *Copiar informes* i *Comprovar Excel* es mostra, en petit,
  l'**última execució** (`_LastRunText`), llegida de `actualitzat_el`
  (`informes-db.json`), `copiat_el` (`copia-informes-state.json`) i `comprovat_el`
  (`comprovar-excel-state.json`, que ara escriu `Invoke-ComprovarExcel` via
  `_SaveRunTimestamp`). *Editar base* no en té.
- **Textos del correu del mòbil** (`Invoke-EmailTextos`, `suport/EmailTextos.ps1`,
  rajola 📧 a MÒBIL, acció `emailtextos`): editor dels textos que l'app mòbil
  envia al titular per EmailJS. Viuen a **`docs/dades/email-textos.json`** (sense
  dades personals → committejable) que `docs/app.js` llegeix (`carregarJson` +
  `aplicarEmailTextos`, amb els defaults `EMAIL_TEXTOS_DEFAULT` de fallback).
  **Model simplificat: només 2 claus, `assumpte` i `cos`.** Al **cos** hi surt
  TOT (capçalera, text CA/ES, avís…) i els requeriments seleccionats s'insereixen
  allà on hi ha la variable **`{REQUERIMENTS}`** (`buildEmailBody`/`buildEmailHTML`
  fan `cos.split("{REQUERIMENTS}")` i hi encasten `buildRequirementsList/HTML`).
  Variables: `{REQUERIMENTS}{ID_GIA}{ADRECA}{ACTIVITAT}{TITULAR}{DATA}` (`fillPh`),
  **`**negreta**`** (`mdHtml`→`<b>`, `stripMarkers` al text pla) i **auto-enllaç**
  dels URLs http(s) (`autolinkHtml`). Cada línia del cos → un `<div>` a l'HTML.
  L'editor (`Invoke-EmailTextos`) té 2 camps (assumpte + cos gran) i avisa si el
  cos no conté `{REQUERIMENTS}`. En desar, l'`Actualitzar.bat` publica
  `email-textos.json` (pas **2b**, commit ABANS del stash). Funcions pures a
  `EmailTextos.ps1` (`_DefaultEmailTextos`, `_EmailTextosFields`,
  `_LoadEmailTextos`, `_SaveEmailTextos`) amb tests; la finestra només a Windows.
- **Feedback del xip ✏️ (editor de catàlegs):** a `Select-Mode`, els botons de
  tipus d'informe tenen un xip clicable que obre l'editor; en passar-hi el ratolí
  (`add_MouseMove`/`add_MouseLeave` → `$entry.ChipHover`) el cursor passa a **mà**
  i el xip es **ressalta** (fons més intens + vora granat), repintant només quan
  l'estat de hover canvia.
- **Exportar llistats (CSV):** botó a *Editar base d'informes* →
  `Export-EstatsActivitats` (`Informes.ps1`). Escriu un CSV (`;`, UTF-8 amb BOM,
  a `_ResolveOutputDir`) amb una fila per informe de les activitats en Estat
  `Requeriment` i `Precinte / Cessament`: Estat, GIA, Titular, Adreça (creuada
  amb l'Excel per GIA), Expedient, Data informe, Conclusió breu. Filtrable a
  Excel per la columna Estat.
- **Copiar informes** (`Invoke-CopiarInformes`, `Informes.ps1`): còpia **plana**
  (tots els Word a una sola carpeta) i **incremental** de `$InformesDir` a
  `$CopiaInformesDir` (nova carpeta configurable, vegeu Configuració). **Només
  copia INFORMES**: `.doc`/`.docx` (ignora `~$…`) **amb data al principi del
  nom** (`_ParseDataInformeFromName`, el mateix criteri que "Actualitzar base");
  qualsevol altre Word NO es copia. Guarda `copia-informes-state.json`
  (`copiat_el`, mateix patró que `actualitzat_el`) i només mira els fitxers
  modificats després de l'última còpia; si el nom ja és al destí NO el recopia;
  **mai** esborra res del destí; si el destí desat canvia, fa còpia completa.
  Mostra una **finestra de progrés amb botó Cancel·lar** i **confirma abans de
  copiar** (amb el nombre d'informes) — mai comença "a cegues". Si es cancel·la,
  NO desa `copiat_el` (la propera vegada torna a comprovar el que faltava).
- **Comprovar Excel** (`Invoke-ComprovarExcel`, `Informes.ps1`): per cada
  activitat en Estat `Precinte / Cessament` de la base d'informes, comprova que a
  l'Excel (fulla "Estès", indexat per GIA = col 1) tingui un **Camp Info** amb
  Nom ∈ `$Script:ExcelPrecinteCampNoms` (`requerit per decret?` / `precinte?`) i
  **Valor que comenci per "SI"** (`_ExcelActivitatActualitzada`, pura + tests).
  Llista en una finestra les desactualitzades, les no trobades a l'Excel i les
  sense GIA (no verificables). Lector de Camp Info autònom (`_ReadExcelCampInfoPerGia`
  + `_FindCampInfoPairs`, pura + tests) — NO dot-sourceja `rutes/Ruta.ps1`.
- **Word a PDF (i signar)** (`Invoke-ConvertirPdf`, `suport/PdfSignar.ps1`, rajola
  📄 a EINES, acció `convertirpdf`): converteix tots els Word (`.doc`/`.docx`) d'una
  carpeta (i subcarpetes) a **PDF al mateix lloc i mateix nom** (via Word COM
  `ExportAsFixedFormat`, `17`=wdExportFormatPDF); salta els que ja tenen un PDF al
  dia si no es marca "sobreescriure". Opcionalment **signa** cada PDF amb
  **AutoFirma** per línia de comandes (`sign -store windows -format pades`). El
  **certificat es tria d'un desplegable** poblat del magatzem de Windows
  (`Cert:\CurrentUser\My`, amb clau privada); del certificat triat se'n treu el
  CN (`_CertCommonName`) i es passa a AutoFirma com a `-filter subject.contains:<CN>`
  perquè el triï sol, sense diàleg. El PDF signat substitueix el sense signar. Es
  va triar **AutoFirma + magatzem de Windows** perquè reutilitza el certificat que
  l'usuari ja fa servir i és l'eina oficial (signatura vàlida per a
  l'administració). La **tria de carpeta** fa servir el helper comú `_AddConfigRow`
  (Configuracio.ps1) — quadre editable + "..." + indicador ✓/⚠ en viu — igual que
  a Configuració (tots els selectors de carpeta del programa han de fer servir
  aquest format). Funcions pures testejables (`_PdfPathForDoc`, `_PdfShouldConvert`,
  `_CertFilterValue`, `_CertCommonName`, `_BuildAutoFirmaSignArgs`,
  `_AutoFirmaCandidatePaths` — retorna un **array pla**, no `,$ArrayList`, perquè
  `@()` l'enumeri bé). Word (COM) i AutoFirma només a Windows. Opcions/estat a
  `pdf-signar-state.json`. **Pendent de provar la signatura a Windows** (AutoFirma).
- **Editar base d'informes — filtres i ordre:** a sobre de la graella hi ha la
  cerca global (conte, totes les columnes) i, a la 2a fila, **filtres per
  columna de SELECCIO MULTIPLE**: Conclusio breu, Estat activitat, Motiu i
  Ignorats (Actius / Ignorats). Cada filtre es un desplegable amb items
  marcables (helper comu `_MakeMultiFilter` a `GenerarInforme.ps1`, un boto +
  `ContextMenuStrip` que no es tanca en marcar): cap opcio marcada = passa tot;
  amb diverses marcades, la fila passa si el seu valor es entre les triades
  (OR dins del filtre; AND entre filtres diferents). Clicar una **capcalera** ordena per aquella
  columna (asc/desc, amb fletxa), pero l'**agrupament per activitat sempre es
  la clau primaria** i la data la darrera: la columna triada nomes desempata
  DINS de cada activitat (ordenacio programatica; `SortMode='Programmatic'`).
- **On és la carpeta d'informes:** a la feina, `$InformesDir` per defecte és
  `I:\Activitats_Ordenances\Activitats\5.- Sergi Fadurdo\Informes`. **A casa**,
  l'usuari en té una còpia en un **disc extern**:
  `F:\FEINA\2022 Ajuntament Cornellà\5.- Sergi Fadurdo` (per tant, els informes
  són a `F:\FEINA\2022 Ajuntament Cornellà\5.- Sergi Fadurdo\Informes`). Per
  treballar-hi en local **NO editis `suport/config.ps1`** (és compartit via
  git i trepitjaria l'altra màquina) — fes servir el botó **⚙ Configuració**
  del programa (vegeu secció següent) o, en una sessió de Claude Code sense
  GUI, escriu directament a `%LOCALAPPDATA%\InformesCornella\settings.json`
  (`{"InformesDir": "F:\\...\\Informes", "ActivitatsDir": "F:\\...\\2_Controls Excels"}`).

## Pas 2 — origen de l'informe (capçalera genèrica REQ1)
- Al **Pas 2** (capçalera genèrica; les actes extraordinàries es queden igual)
  hi ha una tria **Origen de l'informe**: *Documentació aportada* o *Visita
  d'inspecció* (radios; per defecte "doc"). Segons la tria es mostren uns camps
  o uns altres i canvia la línia **"Objecte:"** de la capçalera:
  - **doc** → camps `NUM_ANOTACIO` + `DATA_ANOTACIO`; Objecte: `Doc. aportada
    amb Núm. d'anotació <NUM_ANOTACIO> del <DATA_ANOTACIO>`.
  - **insp** → camp `DATA_INSPECCIO`; Objecte: `Visita inspecció <DATA_INSPECCIO>`.
- La línia "Objecte:" de `ESTRUCTURALS/0 CAPCALERA.docx` (bloc REQ1) és ara un
  únic placeholder **`<<ORIGEN>>`**, que `Apply-HeaderReplacements` substitueix
  pel text muntat per **`_BuildOrigenText`** (funció PURA, testejada) segons
  `ORIGEN_TIPUS` del `$header`. `<<DATES>>` (actes extraordinàries) NO es toca:
  és un flux diferent.
- **App mòbil (`docs/`)**: el Pas 2 també ofereix la tria d'origen (bloc "Origen
  de l'informe" a `index.html` + `muntarOrigen`/`renderCampsOrigen` a `app.js`),
  però al mòbil el valor **per defecte és `insp` (Visita d'inspecció)**. El
  paquet inclou `ORIGEN_TIPUS` i `DATA_INSPECCIO` (a `HEADER_KEYS`). Al PC, si un
  paquet ANTIC no porta `ORIGEN_TIPUS`, es manté el fallback 'doc'.

## ESTRUCTURALS en JSON — FORMAT ESTÀNDARD ÚNIC, editable des del programa
- **Objectiu (petició de l'usuari):** deixar de dependre del Word "rudimentari";
  format estructurat, editable des del programa, amb nivells, negreta/cursiva i
  afegir/treure requeriments. L'usuari va demanar **un únic format estàndard,
  replicable entre tots els ESTRUCTURALS** (mateixos nivells, formats i opcions),
  amb **fills imbricats**.
- **Format estàndard únic** (`suport/CatalegJson.ps1` ho documenta):
  ```
  { "tipus","familia","intro":[<paràgraf>],
    "nodes":[ {"nivell","marca","titol","cos":[<paràgraf>],"fills":[<node>]} ] }
  <paràgraf> = { "runs":[ {"t","b","i"} ], "url": bool }
  ```
  Un **run** és un fragment de text amb negreta (`b`) i/o cursiva (`i`); un
  paràgraf pot ser `"url": true` (enllaç). Els **fills són imbricats** dins cada
  node. La `familia` + la `marca` de cada node donen la semàntica (l'estructura
  és sempre la mateixa):
  - `cataleg` (REQ1, TERMINI): nivell1 `marca=seccio`; nivell2
    `marca=item|subseccio|intro`; nivell3 `marca=fill` (imbricat dins l'ítem).
    `intro` = cos fix (TERMINI no té nodes → cos fix).
  - `conclusions` (0 CONCLUSIONS): nivell1 `marca=grup` (titol=tipus d'informe)
    amb fills `marca=conclusio`; nivell1 `marca=sempre` (frases ::SEMPRE::).
    `intro` = `[capçalera]`.
  - `actextr` (ACT_EXTR_REQ/FAV): nivell1 `marca=seccio` (títol visual, pot dur
    `[[KEY]]`); nivell2 `marca=bloc` (titol = `"[[KEY]] … etiqueta"`, cos =
    contingut). El lector recorre l'arbre → llista ordenada de records
    `@{Text;Style}` per a `Build-ActExtrBlocks`.
- **`0 CAPCALERA` es queda en Word**: és una carta amb escut/taula/format real
  (la generació COPIA el .docx i hi substitueix `<<PLACEHOLDERS>>`), no un
  llistat reconstruïble des d'un model de runs.
- **Lectura** (`suport/CatalegJson.ps1`, headless/testejable): `Read-CatalegJson`,
  `Read-ConclusionsJson` i `Read-ActExtrRecordsJson` tornen EXACTAMENT el mateix
  model en memòria que `Parse-Cataleg` / `Read-Conclusions` / `Build-ActExtrBlocks`.
  El cos s'**aplana** a la mateixa cadena amb marques (`_RunsToMarkup`:
  `**negreta**`, `//cursiva//`, `[[URL]] …`) que ja entenen `Type-RichText`/
  `_SplitTextAndUrls` → **la generació des de JSON és idèntica a la del .docx**.
  `Get-ParsedCataleg`, `Read-Conclusions` i `Parse-ActExtrTemplate` fan servir el
  `.json` si existeix al costat del `.docx` (mateix nom), amb **fallback segur al
  .docx** si el JSON falla. Els `.docx` es conserven (còpia de seguretat).
- **Conversió inicial**: els 5 JSON es van generar a partir dels `.docx` amb un
  convertidor (Python, a scratchpad) que replica `Parse-Cataleg`/
  `Read-Conclusions`/`Build-ActExtrBlocks` i **comprova byte a byte** que (a)
  `flatten(runs)` reprodueix cada línia original i (b) el model del lector NOU és
  idèntic al del lector VELL (ja validat). Tests a `run-tests.ps1` (`_RunsToMarkup`,
  `_JsonParaToBodyLine`, lectura dels JSON reals amb fills imbricats i records
  ACT_EXTR). Nota: `Type-RichText` tracta `**`/`//` com a spans NO solapats (un
  run és negreta O cursiva, mai totes dues alhora).
- **Editor visual "Editar catàlegs"** (`suport/EditorCatalegs.ps1`, WinForms):
  una sola finestra edita QUALSEVOL ESTRUCTURAL en JSON (desplegable amb tots els
  `*.json`). Arbre de nodes (nivell1/2/3, fills imbricats) + entrada especial per a
  la introducció/capçalera; a la dreta s'edita títol, tipus (marca segons família)
  i **cos amb negreta/cursiva reals** (RichTextBox ↔ runs) amb botons per inserir
  `[CAMP:]`/`[OPCIO:]` i enllaços. Es poden afegir/eliminar/moure nodes. En desar
  s'escriu el JSON (sense BOM) amb **validació** (re-llegeix amb el lector) i
  **còpia `.bak`** de seguretat; l'editor només ESCRIU, la generació no en depèn.
  - **Punt d'entrada**: a la finestra principal (`Seguiment.ps1 Select-Mode`),
    el **xip del document** (REQ1/TERMINI/ACT_EXTR) dels botons de tipus d'informe
    porta un emoji d'editar ✏️; clicar-lo (hit-test del rectangle `DocChipRect`
    via `MouseClick`) obre l'editor centrat en aquell document. Dispatch:
    `'editcataleg' → Show-CatalegEditor -focusDoc <Doc>` (ACT_EXTR → ACT_EXTR_REQ).
  - **Funcions pures testejables** (headless): `_Ed_JsonToModel`/`_Ed_ModelToJson`
    (model editable ↔ JSON), `_Ed_SegmentsToRuns` (fragments RTB → runs, forçant la
    invariant de no-solapament), `_Ed_CosToRich`/`_Ed_RichToCos`, `_Ed_MarcaOptions`/
    `_Ed_CanAddChild`/`_Ed_ChildMarca`. Tests a `run-tests.ps1` comproven que
    model→JSON→lector és **idèntic** a llegir l'original (sense pèrdues) per als 5
    fitxers. La finestra (WinForms) només es pot provar a Windows.
  - **Notes d'implementació WinForms** (per no repetir errors): les funcions que
    retornen col·leccions fan servir `return ,$coll` (el `,` evita que el pipeline
    desenrotlli l'ArrayList i trenqui `.Add`); els handlers capturen només `$state`
    amb `.GetNewClosure()`; la reentrada arbre↔editor es controla amb `$state.Busy`;
    el `FormClosing` (no el botó Enrere) fa la pregunta de desar (evita doble prompt).

## Configuració per ordinador (portabilitat)
- El programa és **portable**: cap ordinador nou hauria de necessitar tocar
  codi ni `config.ps1`. Hi ha tres capes de prioritat creixent: valor
  hardcodejat al codi → `suport/config.ps1` (compartit via git, valor per
  defecte comú) → `%LOCALAPPDATA%\InformesCornella\settings.json` (NOMÉS
  aquest PC, mai versionat).
- `suport/Settings.ps1`: `Load-AppSettings`/`Save-AppSettings` (mateix idioma
  que `Save-LastReport`/`Load-LastReport`) + funcions pures testejables
  `_ResolveEffectiveValue` (override si no buit, sino per defecte) i
  `_BuildSettingsOverrides` (què es desa: només els camps que l'usuari ha
  deixat diferents del valor per defecte, mai buits).
- `suport/Configuracio.ps1`: pantalla **⚙ Configuració** del menú principal.
  6 camps (Informes, Excel d'activitats, sortida d'informes, sortida de
  rutes, Drive mòbil, i **carpeta de còpia dels informes** `CopiaInformesDir`,
  que fa servir l'eina *Copiar informes*) amb explorador de carpetes i indicador ✓/⚠ en viu
  (`Test-Path`, no bloqueja desar). Secció "Manteniment": branca + últim
  commit i el botó **🔄 Actualitzar el programa**, que llança el mateix
  `Actualitzar.bat` (`Start-Process`, finestra visible) i tanca l'app —
  `Actualitzar.bat` NO s'ha tocat, segueix funcionant igual en doble clic.
- **Important:** `GenerarInforme.ps1` i `suport/rutes/Ruta.ps1` són processos
  independents (cadascun amb el seu propi `config.ps1`); cada un dot-sourceja
  `Settings.ps1` i aplica l'override pel seu compte, DESPRÉS de carregar el
  seu `config.ps1` i ABANS de derivar-ne res més (p.ex. les subcarpetes de
  Drive a partir de `$DriveBaseDir`).

## Controls periòdics (eina EINES)
- `suport/ControlsPeriodics.ps1` (fitxer NOU, amb BOM): eina **📅 Controls
  periòdics** del menú (secció EINES; acció `controlsperiodics`). Llegeix
  l'Excel d'activitats (fulla "Estès") i llista les activitats amb
  **Classificació general annex** = II o III **o** **Classificació general
  Apartat** amb un número que **comença per 561** (561, 5610…). Reutilitza
  l'estructura d'**Editar base d'informes** (DataGridView + filtre + ordre
  programàtic; filtre de **selecció múltiple** `_MakeMultiFilter`). Columnes: ID Activitat, Raó social, Raó soc. E-mail, Rep. Leg.
  E-mail, Adreça (Emp. Tipus via+Carrer+Número+Lletra), Data llicència/
  comunicació, Data control inicial/verificació, Periodicitat CP, Data control
  periòdic, Proper CP previst, Classif. annex, Classif. Apartat, Activitat
  principal. Filtre de **selecció múltiple** (II / III / 561; cap = tots) + cerca de text;
  **ordre per defecte: Proper CP previst ascendent** (més antic primer; els buits al final); clic a
  la capçalera reordena (les columnes de data ordenen per data real, no pel
  text). Botó **Exportar (CSV)**. Funcions pures testejades:
  `_ControlPeriodicClassify` (II/III per límit de paraula; 561 = número que
  comença per 561, exclou 1561) i `_ParseCellDate`. Les columnes de l'Excel es
  localitzen pel text de capçalera (`_FindColIndex`), amb fallback als índexs
  fixos coneguts de l'adreça.
- **Generar informes (lot)**: la graella té una **columna de casella "Generar"**
  (col 0, editable; la resta només lectura; la selecció es desa a l'objecte fila
  `.Sel` i sobreviu a filtres/ordre) i un botó **Generar informes**. Per cada
  activitat marcada genera un **requeriment** (com "Requeriment - Nou") amb el
  motor NO interactiu. **NO fa servir cap catàleg nou**: el catàleg és **REQ1**;
  d'entre les seves deficiències (Títol 2) tria la de control periòdic segons la
  classificació — `Decret 112/2010 - control periòdic` (561), `Annex III Llei
  20/2009 - control periòdic` (III) o `Annex II Llei 20/2009 - control periòdic`
  (II) — via `_ControlSectionTitle` + `_FindItemKeysByTitle` (recorda: al parser
  Títol 1 = secció, Títol 2 = item), i la conclusió **"Requeriment"**
  (`Read-Conclusions -reportType 'REQ1'` + `Build-ConclusionsFromTitles`).
  Capçalera amb les dades de l'activitat i **sense Objecte** (`ORIGEN_TIPUS='cap'`
  → `_BuildOrigenText` torna ''). Una activitat mai compleix més d'un criteri;
  `_ControlCatalegKind` manté igualment una precedència (561 > III > II). Sortida
  a `_ResolveOutputDir` (Informes generats), amb finestra de progrés +
  Cancel·lar. Si l'item de control periòdic no és a REQ1, l'informe d'aquella
  activitat s'omet amb avís. Funcions pures testejades: `_ControlCatalegKind`,
  `_ControlSectionTitle`, `_FindItemKeysByTitle`.

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
