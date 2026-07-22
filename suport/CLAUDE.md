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
- **Menú Pas 1 — apartat INFORMES:** sota EINES hi ha un apartat propi
  **INFORMES** amb 4 rajoles: 🗃 *Actualitzar base* (`informesdb`), 📋 *Editar
  base* (`informesdbedit`), 📁 *Copiar informes* (`copiarinformes`) i ✅
  *Comprovar Excel* (`comprovarexcel`). Les rajoles es dibuixen amb el helper
  `$addTileRow` de `Select-Mode` (`Seguiment.ps1`), reutilitzat per EINES i
  INFORMES. Dispatch al `switch` de `Main` (`GenerarInforme.ps1`).
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
- **Editar base d'informes — filtres i ordre:** a sobre de la graella hi ha la
  cerca global (conte, totes les columnes) i, a la 2a fila, **filtres per
  columna** (desplegables): Conclusio breu, Estat activitat, Motiu i Ignorats
  (Tots / Actius / Ignorats). Clicar una **capcalera** ordena per aquella
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

## Plànol públic d'activitats precintades
- `suport/rutes/Precintades.ps1` genera `docs/dades/precintades.json` a partir
  de l'Excel d'activitats (fulla "Estès"): les activitats amb el camp lliure
  "PRECINTE ACTIVITAT?" i valor que comença per "SI". La pàgina pública
  `docs/precintades.html` (GitHub Pages) el llegeix i pinta el mapa (Leaflet).
- Ho refresca i puja a `main` **`Actualitzar.bat`** (pas 7). URL pública:
  `https://xexifm.github.io/informes-cornella/precintades.html`.
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
