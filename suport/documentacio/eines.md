# Les eines del menú EINES: com funcionen i per què

> Surt del `suport/CLAUDE.md` (setembre 2026), que havia tornat a passar de
> 2.000 línies. Són les històries de diagnòstic de cada eina: **llegeix la
> secció de l'eina abans de tocar-la**. Les regles que valen per a tot el
> projecte continuen al `CLAUDE.md`.

## Eina «Seguiment» (fila GIA): d'on surt cada cosa
`SeguimentGia.ps1` substitueix un Excel de fórmules de l'usuari
(`0_PLANTILLA.xlsx`). Val la pena tenir apuntat com estava fet, perquè el codi
n'és la traducció literal:

- La plantilla tenia les columnes **desplaçades 15 posicions** (`ID Activitat` a
  la P, no a l'A) perquè hi havien inserit **15 columnes ocultes** d'ajuda: **5
  blocs de 3** (A-B-C … M-N-O), un per pestanya. Dins de cada bloc, la 3a
  columna feia `MATCH(criteri, P:FZ, 0)` → la posició del `Camp Info N - Nom`
  que coincidia; la 2a era un comptador que només avançava quan la fila
  coincidia; la 1a, la clau que després buscava el `VLOOKUP` de la pestanya.
- **El criteri és NOMÉS tenir aquell `Camp Info`, digui el que digui el valor.**
  No es demana que comenci per SI. Comprovat contra les dades reals: a REQUERIT
  DECRET hi havia dues activitats amb valor `PROCEDIMENT ESMENA` i
  `CONTROL PERIODIC VTO. 27-04-2026…`. **`Invoke-ComprovarExcel` (Informes.ps1)
  fa servir un criteri DIFERENT**, allà sí que cal el SI: són dues coses
  distintes i no s'han d'unificar.
- **ANNEX II**: `Classificació general annex` = `II` **i** `Descripció lliure`
  amb contingut. Ull: la base de dades té aquesta capçalera **dues vegades**; la
  plantilla **mostra la primera** (CZ) i **filtra per la segona** (DG). A les
  dades reals les dues són idèntiques a totes 1.312 activitats, però
  `_SgFilesPerFulla` ho reprodueix igual (busca explícitament la 2a per al
  criteri) per no canviar el resultat si algun dia divergeixen.
- Les columnes es resolen **pel nom de la capçalera**, com el `MATCH` de la
  plantilla, i es reaprofita **`_FindCampInfoPairs`** (`Informes.ps1`), que ja
  localitzava les parelles `Camp Info N - Nom`/`- Valor`.
- **Validat cel·la a cel·la** contra la plantilla real: 26 / 24 / 48 / 8 / 51
  files, mateixos ID i mateixos valors i en el mateix ordre. L'única diferència
  volguda són les dues columnes de data d'ANNEX II, que ara surten com a
  `dd/MM/aaaa` (`_FormatDateOnly`) en lloc de `2020-12-22 00:00:00.0`.
- **Colors, calcats de la plantilla** (`$Script:SgColors`, trets del seu
  `styles.xml`): capçalera amb lletra **blanca sobre blau marí** — al fitxer són
  colors **indexats de la paleta antiga**, `indexed 9` = `FFFFFF` i `indexed 18`
  = `000080` — i files de dades amb **ratllat de zebra** `E8E8E8` (`theme 2`),
  la 1a ombrejada i després una sí una no. La columna `N` de la capçalera **no
  porta fons**. Abans hi havia un `#D9E1F2` que m'havia inventat. L'Excel vol el
  color com a **`R + G*256 + B*65536`** (BGR), no com un `#RRGGBB`.
- **L'alçada de la capçalera s'ajusta DESPRÉS de posar les amplades**: amb
  `WrapText`, `Rows(2).AutoFit()` calcula l'alçada a partir de l'amplada de la
  columna, o sigui que fer-ho abans deixa el text tallat igualment.
- **Impressió** (i per tant el PDF): horitzontal, **A3**, `Zoom=$false` +
  `FitToPagesWide=1` + `FitToPagesTall=$false` (hi caben totes les columnes),
  marges 0,5 cm, `PrintTitleRows='$1:$2'` i peu `&A` / `Pàgina &P de N`.
- **El peu compta les pàgines DE CADA PESTANYA**, no del PDF sencer:
  - `FirstPageNumber = 1` a cada fulla — per defecte l'Excel numera de correguda
    per tot el treball d'impressió i ANNEX II hauria començat per la 540.
  - El total **no pot ser `&N`**: en una exportació de diverses pestanyes, `&N`
    és el total del PDF. Es llegeix **`$sh.PageSetup.Pages.Count`** i s'escriu el
    número literal al peu. S'ha de llegir **al final**, quan la paginació ja està
    decidida (orientació, paper i ajust a l'ample); abans donaria un altre
    número. Si no es pot llegir, el peu es queda amb `Pàgina &P` (sense total),
    mai amb un total fals.
- **Es TRIA què s'exporta** (caselles a la finestra, totes marcades per defecte).
  `_SgOpcionsExport` és **l'única llista** — la fan servir tant la finestra com
  `_SgConstruirLlibre`, o sigui que no es poden desincronitzar — i
  `_SgSeleccioTeEstes` / `_SgFullesTriades` (pures) decideixen què es munta.
  El llibre porta **només** el que s'ha triat, i per això el PDF s'exporta
  sencer: `$wb.ExportAsFixedFormat`, que respecta el `PageSetup` de cada
  pestanya. **`$excel.ActiveWindow.SelectedSheets.ExportAsFixedFormat` NO
  EXISTEIX**: `SelectedSheets` és una col·lecció `Sheets`, i aquest mètode només
  el tenen `Workbook`, `Worksheet`, `Chart` i `Range`. L'Excel ho deia clar —
  *"[System.__ComObject] no contiene ningún método llamado 'ExportAsFixedFormat'"*.
- **El full buit del llibre nou fa de PLACEHOLDER**: un llibre no pot quedar-se
  sense cap fulla, així que només s'esborra **al final**, quan ja hi ha les
  pestanyes de debò. Abans es podia esborrar de seguida perquè `Estès` es copiava
  sempre; ara `Estès` pot no estar triada.
- **A l'Excel de l'usuari, `Range.Value2` NOMÉS ACCEPTA CADENES.** Dues rondes
  seguides amb el mateix patró:
  - `$rang.Value2 = $matriu` → *"Unable to cast object of type
    'System.Object[,]' to type 'System.String'"*
  - `$cel.Value2 = 1` → *"Unable to cast object of type 'System.Int32' to type
    'System.String'"*
  - …mentre que el títol i les capçaleres (cadenes) s'escrivien **sempre** bé.

  L'adaptador COM d'aquell PowerShell resol el `put` de `Value2` com si demanés
  una cadena. Per això **tot el que s'escriu passa per `_SgValorCella`** (pura),
  que retorna sempre un `String`. No es perd res: de tota la taula, l'únic valor
  que no era text ja era la columna **`N`** (el número de fila) — tota la resta
  ve de `$cel`, que ja retorna cadenes — i l'Excel interpreta el text en
  assignar-lo igual que si l'escrivissis a mà (`"1"` segueix sent el número 1).
  `_SgEscriuMatriu` prova tres camins (bloc → bloc per `InvokeMember` →
  **cel·la a cel·la amb text**) i, si fallen tots tres, peta dient què ha passat
  a cada un. Cel·la a cel·la aquí es pot permetre: aquests llistats són de
  desenes de files (26/24/48/8/51), no de milers. El «matriu vs cel·la a cel·la
  és de minuts a segons» valia per a la **lectura** de la base sencera
  (1.312 × 152), no per a aquesta escriptura.

  **Pendent de saber:** si el problema és només de `Value2` o de **qualsevol**
  `put` que no sigui cadena. Els booleans (`$excel.Visible = $false`) i les
  cadenes (`$sh.Name`) funcionen; encara no s'ha arribat a executar cap
  assignació NUMÈRICA (`RowHeight`, `Font.Size`, `ColumnWidth`, `Interior.Color`…
  són totes a `_SgFormatarFulla`, que va després). Si algun dia surt l'avís de
  «pestanya sense format», la resposta és que sí i caldrà passar aquelles
  assignacions per un helper amb respatller.
- **El format d'una pestanya no pot endur-se el fitxer**: `_SgFormatarFulla` va
  dins d'un `try/catch` **dins del bucle** (mateixa lliçó que els reintents de la
  signatura) i, si falla, s'afegeix un avís al resum i es continua. Les dades
  són el que importa; els colors, no.
- **Quan una eina de COM peta, ha de dir ON.** `_SgConstruirLlibre` porta un
  `$pas` («obrint l'Excel», «copiant la fulla Estès», «bolcant les dades a
  PRECINTES»…) i `_SgTextError` hi afegeix **la línia exacta del codi**. Sense
  això, un missatge com el de dalt obliga a endevinar quina de les vint crides a
  l'Excel ha estat, i cada intent costa una volta sencera amb l'usuari.
- **PER QUÈ VA PETAR EL PRIMER DIA** («No se puede convertir el valor
  "System.Object[]" … al tipo "System.Int32"»): `_FindCampInfoPairs`
  (`Informes.ps1`) acaba amb `return ,@($pairs)`. La coma hi és a posta —
  protegeix el cas d'UNA sola parella, perquè `$x = f` no rebi el hashtable pelat
  i `.Count` no li doni el nombre de CLAUS — i per tant **s'ha de consumir SENSE
  `@()`**. `SeguimentGia.ps1` hi posava un `@()`, que **hi torna a posar la
  capa**: `$pairs` quedava com un array d'UN element que contenia l'array de
  parelles, `$p.NomCol` feia **enumeració de membres** (retornava un `Object[]`
  amb tots els `NomCol`) i `[int]$p.NomCol` petava. Ara: la crida va sense `@()`
  **i** `_SgFilesPerFulla` passa el que rep per **`_SgAplanaPairs`** (pura), que
  accepta les dues formes. Hi ha prova de regressió amb la forma embolcallada.
- Els segells de «última vegada» del menú (`Select-Mode`) s'indexaven per
  POSICIÓ dins de la fila de rajoles; en moure *Comprovar Excel* a la fila GIA
  haurien anat a la rajola equivocada. Vegeu la secció del segell, més avall.
- **Les amplades de columna han de tenir un FORAT** on la plantilla no en
  defineix cap: `Rep. Leg. Mòbil` no porta amplada pròpia (es queda amb la de
  defecte, `baseColWidth=10`). Al principi no hi era i **totes les amplades de
  després ballaven una posició**: la columna ampla del Valor es quedava sense
  amplada i la del text ajustat sortia estreta. Als arrays d'amplades el **0**
  vol dir «no la toquis».

## «Enviar correu»: el GIA surt del DOCUMENT, no de l'últim informe

Bug real i vistós (setembre 2026): l'usuari genera un **Seguiment** del GIA 1466
i el diàleg d'enviar surt amb l'assumpte **«GIA 1000 Requeriments»** i els
destinataris d'una altra activitat.

- **La causa**: `Send-CorreuPerDocx` muntava tota la capçalera amb
  `Load-LastReport`, que és **l'últim informe fet amb l'ASSISTENT**. Un informe
  de *Seguiment* no hi passa (no desa capçalera), o sigui que allà encara hi
  havia el «Requeriment - Nou» d'abans. **El fitxer que s'enviava i les dades del
  correu venien de dues activitats diferents** — i res no ho deia. És la regla 1
  d'aquest document al revés: dues fonts per a la mateixa cosa, i ningú les
  comparava.
- **La regla ara**: el GIA surt del **document que s'envia**, per aquest ordre:
  1. el **nom del fitxer** (`_GiaDelNomFitxer`, pura) — els noms els fa
     `_GetOutputFileName` (`..._GIA <id>.docx`) i `_SeguimentOutputName` els
     **conserva**. Només s'agafen els **dígits** de darrere de «GIA»: el sufix
     d'unicitat `_2`/`_3` de `_GetUniqueOutputPath` NO és part de l'id;
  2. la **capçalera del document** (`_ReadDocxParagraphs` + `_ExtractIdGia`, les
     dues ja existien a `Informes.ps1`).
- **`_CorreuGiaDecideix` (pura) NO ENDEVINA MAI**: si no en troba cap (informes
  antics sense ID GIA) **o si els dos no coincideixen**, retorna
  `CalPreguntar = $true` i `_DemanaIdGia` ho demana, dient **els dos valors** que
  ha vist. Cancel·lar vol dir **no enviar**: val més no enviar que enviar a qui
  no és.
- **La capçalera es reconstrueix des de l'EXCEL per aquell GIA**
  (`_CorreuHeaderMerge`, pura). La de `Load-LastReport` només s'aprofita **si el
  seu `ID_GIA` és el mateix** —i només per omplir el que l'Excel no té (núm.
  d'anotació)—, mai per trepitjar l'Excel. Això arregla alhora l'assumpte
  (`{ID_GIA}`) i les variables `{TITULAR}`/`{ADRECA}`/`{ACTIVITAT}`, no només els
  destinataris.
- **El GIA es mira ABANS de llegir el cos**: si s'ha de preguntar o es
  cancel·la, no s'ha fet feina de franc. (Ara el cos tampoc no obre el Word:
  vegeu *El format del correu*, més avall.)
- Si el GIA no és a l'Excel, **s'avisa i es continua** amb el destinatari a mà:
  aturar-ho seria pitjor que deixar enviar-lo.
- `_CorreuEmailsActivitat` s'ha esborrat: obria l'Excel per treure NOMÉS els dos
  correus i ara `_CorreuActivitatPerGia` retorna la fitxa sencera (una sola
  obertura) que alimenta capçalera i destinataris alhora.
- **Colors dels botons**: blau marí «Enviar» / vermell «No enviar»
  (`$Script:CorreuBlauMari` / `$Script:CorreuVermell` + **`_StyleAccentButton`**,
  nou a `UiComuns.ps1` al costat de `_StylePrimaryButton` — no una tercera còpia
  de l'estil). Un correu **no es pot desenviar**: les dues accions oposades han
  de distingir-se d'un cop d'ull.
- **I AQUELLS COLORS VAN TRENCAR EL PROGRAMA EN HEADLESS.** Es van posar a
  **àmbit d'script**, i allà `[System.Drawing.Color]` s'avalua **en carregar el
  fitxer**: en headless (`Actualitzar.bat`, `RecordatorisAuto`, les proves)
  `Motor.ps1` no fa l'`Add-Type`, i el motor **sencer** peta amb *«No se
  encuentra el tipo [System.Drawing.Color]»*. L'usuari es va quedar sense vistes
  en Word, sense dades del mòbil i sense refresc del Drive a cada actualització.
  - **La convenció ja hi era i no es va seguir**: `$Script:BrandMaroon`
    (`UiComuns.ps1`) i `$Script:ConfigUiAccent` (`Configuracio.ps1`) declaren a
    `$null` i omplen **dins d'un `if (-not $Script:HeadlessTest)`**. Dins d'una
    **funció** sí que hi poden anar: només s'avalua en cridar-la, i aquestes les
    crida només la interfície.
  - **PER QUÈ LA SUITE NO HO VA VEURE, i és el que cal recordar**: al **pwsh 7
    de Linux** `System.Drawing.Color` viu a `System.Drawing.Primitives`, que és
    del framework i sempre hi és → resolia bé. Al **Windows PowerShell 5.1** viu
    a `System.Drawing.dll`, que en headless no s'ha carregat → peta. **Un tipus
    de .NET pot estar en assemblatges diferents entre 5.1 i 7**: que la suite
    passi a Linux no vol dir que el fitxer es pugui carregar al PC.
  - **Guard** (`06-guards.ps1`, validat injectant el defecte i comprovant que
    diu fitxer i línia): recorre l'AST de tot `suport/` i falla si un
    `[System.Drawing.*]` o `[System.Windows.Forms.*]` queda a **àmbit d'script**
    sense estar dins d'una funció ni d'un `if` que miri una bandera de headless.
    Sobre el codi d'avui: **zero falsos positius** (tots els altres ja hi eren).

### Destinatari buit = correu de prova per a tu

L'usuari esborrava el destinatari per rebre el correu només ell en CCO i el
diàleg ho aturava («Indica almenys un destinatari»). **No es pot enviar només en
CCO**: la plantilla d'EmailJS necessita el `to_email` i, buit, el servei el
rebutja. Per això `_CorreuDestinatariBuit` (pura): si el destinatari és buit, la
**CCO marcada per defecte** a `docs/dades/email-textos.json` passa a ser el
destinatari i **surt de la CCO** (no arriba dos cops). L'adreça no és al codi
(repositori públic). Si no hi ha cap CCO per defecte, es manté l'avís. El mòbil
(`docs/app.js`) encara exigeix destinatari.

## El format del correu: el de REQ1, i el mateix al PC i al mòbil

Petició de l'usuari (setembre 2026), amb dos correus de prova del GIA 1398 al
davant: el del PC i el del mòbil no s'assemblaven entre ells ni a l'informe.

- **Què passava.** Cada un pintava pel seu compte. El PC obria el Word i
  **començava al primer paràgraf en majúscules**: «ID GIA: 1398» ho és, i el
  correu repetia la capçalera sencera, «INFORME» i la nota de l'Ordenança, i
  sagnava totes les línies de cos. El mòbil posava les seccions en negreta, els
  enllaços sagnats, cap línia en blanc entre punts i **cap conclusió**. I tots
  dos deien «s'han detectat a la visita…» encara que l'informe fos de
  documentació aportada.
- **Ara**:
  - la **capçalera del correu** són les línies amb etiqueta de `0 CAPCALERA`
    (ID GIA, Exp. Núm., Adreça, Activitat, Titular, Objecte), en taula perquè
    l'Outlook respecti la columna (2,75 cm, la sagnia de la plantilla; hi ha
    guard). Una línia sense valor no surt;
  - `{INTRO}` tria la frase segons l'**Objecte**: `introDoc` (amb el núm. i la
    data d'anotació), `introDocSenseAnotacio` o `introInsp` (sense data, la
    d'avui, també a l'Objecte). Les frases són a `email-textos.json`; l'editor
    de textos no les edita però **les conserva en desar**, com el `bcc`;
  - el **cos** té el format de l'informe: secció en majúscules sense negreta,
    número en negreta, enllaç sense sagnar a 10 pt, sub-punts amb pic i sagnia
    francesa (12 pt el primer, 6 els altres), una línia en blanc després de cada
    punt. **Sense** conclusions (vegeu la tercera ronda), ni «Ho poso al seu
    coneixement» ni la signatura.
- **Com es garanteix que coincideixin** (`suport/CorreuFormat.ps1` i
  `docs/correu.js`, que n'és la còpia en JavaScript):
  - les mides surten de `$ReportFormatConfig` (`_CorreuFormat`) i el mòbil llegeix
    les mateixes de `docs/dades/correu-format.json`, que genera `ExportaDades`.
    Guard: el fitxer publicat ha de ser igual a `_CorreuFormat`;
  - les línies de la capçalera i les plantilles de l'Objecte
    (`$Script:OrigenPlantilles`, ara en un sol lloc a `MotorInforme.ps1`) van a
    `docs/dades/capcalera.json` (`Correu`, `Origen`);
  - **prova creuada amb Node** (`tests/proves/08-correu-format.ps1`): la selecció
    de REQ1 sencera (els punts sense camps) passa per `Build-CatalegBlocs` +
    `_CorreuBlocsAHtml` i per `correu.js`, i l'HTML ha de ser **idèntic** (uns
    180.000 caràcters). Validada injectant un defecte a `correu.js`. Sense Node
    diu `OMES`.
- **El PC llegeix el `.docx` sense Word** (`_CorreuDocxLlegeix`, XML): respecta la
  sagnia, els espais i la negreta **de cada paràgraf**, que és el que cal per a un
  informe de seguiment (anotacions datades, punts pendents en negreta), i
  comença **després** de la frase del catàleg («…deficiències… esmenar»). S'obre
  amb `FileShare.ReadWrite`: l'informe pot ser obert al Word.
- **El peu** («feu-hi constar: ID GIA …, Adreça …, Titular …») agafa de la
  capçalera de l'informe el que l'Excel no té.
- **Segona ronda (setembre 2026)**, demanada per l'usuari:
  - **INFORME** sota la capçalera, centrat i en negreta, com a l'informe. El
    text surt de `0 CAPCALERA` (`_CorreuCapcaleraTitol`: el primer text després
    de l'última línia amb etiqueta) i el mòbil el llegeix de `capcalera.json`
    (`Titol`).
  - **Seguiment: negreta NOMÉS al comentari** («No s'aporta.»). El seguiment
    actual ja no posa el punt en negreta, però els d'abans sí (la marca que
    encara llegeix `_InferResolvedFromBold`): `_CorreuSenseNegretaSeguiment`
    treu la negreta d'un punt que la té **sencera** i la torna a posar només al
    número. Un punt amb una part en negreta (del catàleg) no es toca.
  - **El peu, per idiomes**: «Com presentar la documentació / Cómo presentar
    la documentación» en **una línia, en vermell**; a sota, el bloc CATALÀ i el
    bloc ESPAÑOL (títols sense negreta), cadascun amb tota la seva informació
    (la instància, a qui va dirigida i l'avís IMPORTANT), en lloc de les frases
    barrejades. Res del peu en negreta, i sense línia separadora. El vermell és una
    marca nova dels textos del correu, **`!!text!!`** (`_TextToHtml` i
    `correu.js`), que també entenen els recordatoris i els controls periòdics.
- **Tercera ronda (setembre 2026): fora les conclusions**, demanat per l'usuari,
  al PC i al mòbil. El PC deixa de llegir el `.docx` al títol **CONCLUSIONS**
  (`_CorreuDocxLlegeix`: «conclusions» o «conclusió», normalitzat; si l'informe
  no el porta, s'atura a la frase de tancament com abans). El mòbil ja no afegeix
  les conclusions triades al Pas 4 ni a l'HTML ni al text pla (han desaparegut
  `blocsConclusions` de `correu.js` i `conclusionsTriades` d'`app.js`); les
  conclusions continuen anant a l'**informe**. Guard a `06-guards.ps1`.

## Base d'informes (informes-db.json)
- El motor de la base d'informes és `suport/Informes.ps1`: escaneja `$InformesDir`
  (per defecte `...\5.- Sergi Fadurdo\Informes`) i, per cada informe (`.docx` o
  `.doc` antic amb data al principi del nom), en treu **data + ID GIA +
  conclusió**, agrupat per activitat (per GIA; si no en té, per **carpeta**), a
  `local\base-dades-activitats\informes-db.json` (dins de `local/`: mai es puja).
  Botons al menú: **🗃 Actualitzar** i **📋 Editar** (marc "Base d'informes").
- Lectura de `.docx` **sense Word** (zip) reutilitzant les primitives de
  **`Docx.ps1`** (`_LoadDocxXml`, `_ParagraphTextXml`). Lectura de `.doc`
  antics (Word 97-2003) via **Word COM** (`_ReadDocParagraphsWord`): instància
  creada mandrosament a `Invoke-InformesDbScan` només si cal reprocessar algun
  `.doc`, i tancada (`Quit()`) en un `finally`. Funcions de text PURES (dates,
  GIA, expedient, conclusió) amb tests a `run-tests.ps1`.
- **Informe de seguiment — sub-punts (fills):** `_BuildSeguimentModel`
  (`Seguiment.ps1`) aplana a **UNITATS accionables**: un requeriment sense fills →
  1 unitat; un requeriment **amb fills** (sub-punts amb pic; `IsBulletChild` a
  `_CollectParaRecordsXml`: `numId≠0` i (pics o `ilvl>0`)) → **1 unitat per fill**
  (el requeriment fa de capçalera i cada fill es resol per separat). Cada unitat
  té `ParaIndex` propi, i el motor (`_SeguimentBlocksXml`/`_ApplySeguimentTransform`)
  hi ancora la seva anotació datada sense canvis (les subseccions subratllades i
  els espaiadors buits tallen el bloc). La UI (`Prompt-SeguimentComments`) mostra
  un checkbox+comentari per unitat, amb `Label` "Req. N (tema): <fill>" i sagnat
  per als fills. Funcions pures amb tests (`_ShortenText`, `_SeguimentParentTopic`,
  `_BuildSeguimentModel` amb fills).
- **L'anotació NO hereta l'espaiat del requeriment.** El clon del `pPr` es fa per
  quedar-se la **sagnia i l'estil**, no els espais: `_MakeAnnotationParagraphXml`
  esborra el `w:spacing` clonat i després hi posa el que decideix ell. Sense
  això, un requeriment que porti un `after` (el que el separa del punt següent)
  l'encomanava a **totes** les seves anotacions i, a la segona entrega,
  apareixia un **forat entre les dues línies datades**. Es notava en un sol punt
  de l'informe — el que casualment duia aquell `after` — i per això semblava
  aleatori.
- **L'espai de sota del bloc es MOU, no es copia**: ha d'anar sempre a l'**últim**
  paràgraf del bloc, i cada anotació nova passa a ser-ho.
  `_TakeSpacingAfterXml` el pren del que ho era fins ara (el cos del requeriment
  o l'anotació anterior) i el passa a la nova. Si no es mogués, se n'acumularia
  un a cada ronda. **Compte que això també passava entre dues anotacions d'un
  SUB-PUNT**, on l'`after` el posa la regla del sub-punt.
- **Format de l'anotació d'un sub-punt:** `_MakeAnnotationParagraphXml` clona el
  `pPr` del paràgraf que anota i hi força `numId=0` perquè l'anotació **no
  s'enumeri**. Això té un efecte col·lateral: si el sub-punt treia la sagnia de
  la **numeració** (llista real del Word, sense `w:ind` propi), l'anotació la
  perdia i quedava desalineada. Per això, quan `$req.IsChild`, l'anotació rep
  (només si el `pPr` clonat no en portava cap) una **sagnia explícita**
  `AnnotationIndentCm` i un **espai a sota** `AnnotationSpaceAfterPt`, perquè el
  sub-punt següent no li quedi enganxat. Els valors viuen a
  `$ReportFormatConfig` de **`Format.ps1`** (que és qui mana en el format del
  document) i `Seguiment.ps1` només els llegeix — `_AnnotationFormatTwips` els
  passa a **twips** (1 cm = 1440/2,54; 1 pt = 20), amb els mateixos valors per
  defecte si `Format.ps1` no s'ha carregat. L'ordre dels elements dins de
  `<w:pPr>` (`pStyle, numPr, spacing, ind`) es respecta: fora d'ordre el Word
  es queixa del document.
- **Separació ítem → primer sub-punt:** `Format-Bullet -First` (`Format.ps1`)
  aplica `PrimerSubpuntSpaceBeforePt` (12 pt) en lloc de `BulletSpaceBeforePt` (6 pt) al
  **primer** punt que penja d'un ítem numerat, perquè no quedi enganxat al text
  de l'ítem; els punts següents entre ells mantenen els 6 pt. `Motor.ps1` marca
  el primer fill EMÈS (no el primer del catàleg: els fills sense línies es
  salten). Es posa l'espai a **`SpaceBefore` del fill** i no a `SpaceAfter` de
  l'ítem perquè entre l'ítem i els fills hi pot haver línies extra o un URL, i
  llavors l'espai separaria l'ítem del seu propi cos.
- **Sangria dels fills:** la vinyeta d'un fill s'alinea el text a
  `BulletChildIndentCm` = **1 cm** amb francesa `BulletChildHangCm` = **0,5 cm**
  (al XML: `w:ind left="567" hanging="283"`). És el **mateix** 1 cm que
  `ChildIndentCm`, que fan servir les sub-línies i els enllaços del fill
  (`Format-Body`/`Format-Url -IsChild`), de manera que tot el bloc del fill
  queda alineat. Els punts de **primer nivell** (només l'informe favorable
  d'activitat extraordinària) mantenen `BulletIndentCm` 1,25 / `BulletHangCm`
  0,62: són un altre document i no s'han tocat.
- **La negreta del número d'un ítem s'aplica pel RANG**, no amb
  `$sel.Font.Bold = 1` … `= 0`. Motiu real: el `Bold = 0` d'després d'escriure
  el número actua sobre el **format d'escriptura del punt d'inserció**, i el
  Word no sempre l'hi aplica; quan no ho feia, **tot** el text de l'ítem sortia
  en negreta i el número i el text quedaven fusionats en un sol `<w:r>` (es veia
  a tots els ítems de CONTROLS INICIALS i CONTROLS PERIÒDICS de la vista de
  REQ1, i només allà). `Format-Item` escriu ara el número sense negreta, es
  guarda `Range.Start`/`Range.End` i al final fa
  `$sel.Document.Range($numStart,$numEnd).Font.Bold = $true`: així la negreta
  només pot tocar el número i el cos no se la pot encomanar mai. La negreta
  **inline** del cos (`**...**` de `Type-RichText`) no es toca.
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
- **Menú Pas 1 — 4 apartats de rajoles** (`Select-Mode`, `Menu.ps1`, helper
  `$addTileRow`; dispatch al `switch` de `Main`):
  - **EINES** (4): 📍 *Generar ruta* (`ruta`), 🗺 *Coordenades*
    (`coordenades`), 🔒 *Activitats precintades* (`url`, acció `precintades`
    només per al segell), 📅 *Controls periòdics* (`controlsperiodics`).
  - **INFORMES** (4): 🗃 *Actualitzar base* (`informesdb`), 📋 *Editar base*
    (`informesdbedit`), 📁 *Copiar informes* (`copiarinformes`), 📄 *Word a PDF*
    (`convertirpdf`).
  - **GIA** (2): ✅ *Comprovar Excel* (`comprovarexcel`), 📊 *Seguiment*
    (`seguimentgia`).
  - **MÒBIL** (2): 📧 *Textos del correu* (`emailtextos`), 📥 *Revisar mòbil*
    (`revisarmobil`).
- **Segell d'«última execució»: UN sol registre per a totes les rajoles.**
  `local\base-dades-activitats\eines-state.json` → `{ "<accio>": "<ISO>" }`.
  - S'escriu en **un sol lloc**: al final del bucle de `Main` (`Wizard.ps1`),
    quan l'eina torna. Per tant la data vol dir **«l'última vegada que has obert
    i tancat aquesta eina»**, no «l'última vegada que va acabar bé» — és l'única
    cosa que el despatxador pot saber sense tocar les onze eines, i està dit al
    comentari perquè ningú no ho llegeixi com una altra cosa.
  - La llista que es manté és la dels que **NO** en porten
    (`$Script:AccionsSenseSegell` = `nou`, `seguiment`, `actextr`, `config`,
    `editcataleg`), no la dels que sí: així **una rajola nova hi entra sola**.
  - El segell es llegeix per **`$it.Action`** (clau del registre), no per posició
    ni per etiqueta. Abans anava per posició dins d'una fila concreta i, en moure
    *Comprovar Excel* a la fila GIA, hauria anat a la rajola equivocada.
  - Dues excepcions llegeixen **la seva pròpia marca** si la tenen, perquè
    l'escriu el procés mateix quan ha treballat de debò i és més precisa:
    `informesdb` → `actualitzat_el`, `copiarinformes` → `copiat_el`
    (`$Script:SegellPropi`). Si no hi és, es cau al registre.
  - `_SaveRunTimestamp` i `comprovat_el` (`comprovar-excel-state.json`) **es van
    esborrar**: el seu únic ús era pintar aquest segell. `comprovar-excel-state.json`
    ja no s'escriu (si en queda un de vell al disc, és inofensiu).
  - Com que **totes** les files porten segell, els dos helpers de fila
    (`$addTileRow` + `$addTileRowAmbSegells`) es van tornar a fondre en **un
    sol**. Cada rajola es guarda la seva etiqueta a `$tool.StampLabel`, que
    serveix perquè la rajola d'**enllaç** (que no tanca el menú, i per tant no
    passa pel despatxador) s'apunti i es refresqui el segell allà mateix.
  - `_FormatRunStamp` (pura, amb proves) fa el format `dd/MM/aa HH:mm`;
    `_LastRunText` l'ha de fer servir i no duplicar-lo. **Compte a les proves**:
    la marca es desa en hora LOCAL amb desplaçament, o sigui que una asserció amb
    una cadena fixa falla si la màquina va en una altra zona horària — s'ha de
    comprovar l'anada i tornada.
- **L'interruptor A/M de «Copiar informes»** (menú, setembre 2026). L'única
  rajola amb commutador (`Interruptor = $true` a la seva entrada). Va **a
  l'espai del segell**: la data on hi havia la data i la pastilla A/M **on hi
  havia l'hora** (l'hora de l'última còpia no interessava). Ni un píxel més que
  les altres rajoles.
  - **A verd** = automàtic · **M gris** = manual. I **la data també parla**:
    verda si l'última còpia la va fer l'automàtic, grisa si la vas fer tu — així
    es veu si l'automàtic treballa de debò o només està encès.
  - És un `Panel` **dibuixat a mà** (un `Label` no pot portar la pastilla), amb
    el mateix patró de *hit-test* que el xip ✏️ de les rajoles: el rectangle del
    commutador el guarda el **Paint** (`$auto.Rect`), que és l'únic que sap on ha
    quedat després de centrar data + pastilla.
  - La pastilla són **dos semicercles i un rectangle**: el GDI+ no té rectangle
    arrodonit i un `GraphicsPath` serien vint línies per a 24×12 px.
  - El rellotge és un **`Timer` de WinForms d'un minut** (mai un bucle: el menú
    ha de respondre) i **mor amb la finestra** (`FormClosed` → `Stop`+`Dispose`):
    un timer viu disparant sobre controls destruïts peta dins del bucle de
    missatges, on no ho veu ningú. La primera comprovació és al `Shown` (és la
    de «en obrir el programa»).
  - `_FormatRunStamp` va guanyar un segon paràmetre (`$ambHora`) i `_LastRunText`
    / `_LastRunEina` s'han partit en **ISO + format** (`_LastRunIso`,
    `_LastRunIsoEina`): l'interruptor vol **la mateixa marca** amb un altre
    format, i duplicar la cerca era la manera que un dia els dos segells
    diguessin coses diferents.
- **El «?» d'ajuda de cada rajola** (menú, setembre 2026, demanat per l'usuari).
  Cada rajola d'eina porta a la cantonada de dalt a la dreta una rodona amb un
  «?» (dibuixada al `Paint` de la rajola, sense emoji). El clic es mira contra
  `$t.AjudaRect`, que guarda el `Paint`, i mostra el text en un `MessageBox`
  **sense obrir l'eina**. Els textos són a `$Script:AjudaEines` (`Menu.ps1`,
  indexats per acció, com el segell) i es llegeixen amb `_AjudaEina`. Prova a
  `04-correu.ps1`: **cada** acció de les rajoles del menu té text, no n'hi ha cap
  de sobrant i cap passa de 220 caràcters (validada traient-ne un).
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
  `EmailTextos.ps1` (`_LoadEmailTextos`, `_SaveEmailTextos`) amb tests; la finestra només a Windows.
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
- **Copiar informes** (`Invoke-CopiarInformes`, `CopiaInformes.ps1`; fins a la
  revisió d'arquitectura de setembre 2026 vivia a `Informes.ps1`): còpia **plana**
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
- **Copiar informes, mode AUTOMÀTIC** (interruptor **A/M** del menú, setembre
  2026). La mateixa còpia es fa de dues maneres i per això la feina viu en
  **quatre funcions sense cap finestra** (`_CopiaInformesPrepara`,
  `_CopiaInformesCerca`, `_CopiaInformesTria`, `_CopiaInformesCopia`); el que
  difereix —avisar, preguntar, pintar la barra— es queda a la crida, en
  **scriptblocks**. Hi ha **un sol `Copy-Item` a tot `Informes.ps1`**, i un guard
  ho vigila: si el manual i l'automàtic es munten cada un el seu bucle, un dia
  copiaran coses diferents.
  - **Quan toca: UNA sola pregunta**, `_CopiaAutoToca` (pura, amb proves) —
    «des de l'últim **venciment** (les 14:30 que tocaven), s'ha fet cap passada
    automàtica?». Serveix per als dos casos que va demanar l'usuari (el rellotge
    de les 14:30 amb el programa obert, i la passada perduda que es recupera en
    obrir-lo) **i** per al que s'escapava de tots dos: obrir el programa a la
    tarda el mateix dia que no s'ha fet. Amb dues regles separades, aquell cas
    es perdia fins l'endemà.
  - **En un PROCÉS A PART** (`suport/CopiaInformesAuto.ps1`, llançat per
    `Start-ScriptSegonPla`): recórrer la carpeta d'informes pot trigar, i fet
    dins del menú la finestra es quedaria **congelada** —el contrari de «no es
    veurà res»— i, amb els `DoEvents`, l'usuari podria obrir una eina a mig
    copiar. El fill agafa un mutex `Global\InformesCornella.CopiaInformesAuto`
    i **no espera**: si ja n'hi ha un fent la feina, plega.
  - **L'estat** (`copia-informes-state.json`) hi afegeix tres claus: `auto`
    (l'interruptor), `auto_el` (**l'última passada, encara que no copiés res** —
    és el que evita que es repeteixi cada minut) i `mode` (`auto`/`manual`: qui
    va fer l'última còpia de debò). Una passada que no copia res **no** toca
    `mode`, si no la data del menú deixaria de dir res.
  - **`_CopiaInformesDesaEstat` desa NOMÉS les claus que li dones**, damunt del
    que ja hi ha: engegar l'interruptor no pot esborrar la data de l'última
    còpia, ni al revés.
  - **La rajola copia SEMPRE**, digui el que digui l'interruptor (ho va demanar
    l'usuari amb totes les lletres). El commutador és un control **a part** —el
    segell de sota— i el seu clic es mira contra el **seu rectangle**; un guard
    comprova que el clic de la rajola no consulta l'interruptor.
  - Diagnòstic a `%LOCALAPPDATA%\InformesCornella\copia-informes-log.txt`
    (`_CopiaAutoLog`): d'un mode que no ensenya res, si no és per aquest fitxer
    no se'n sap res.
- **Comprovar Excel** (`Invoke-ComprovarExcel`, `ComprovarExcel.ps1`; abans a `Informes.ps1`): per cada
  activitat en Estat `Precinte / Cessament` de la base d'informes, comprova que a
  l'Excel (fulla "Estès", indexat per GIA = col 1) tingui un **Camp Info** amb
  Nom ∈ `$Script:ExcelPrecinteCampNoms` (`requerit per decret?` / `precinte?`) i
  **Valor que comenci per "SI"** (`_ExcelActivitatActualitzada`, pura + tests).
  Llista en una finestra les desactualitzades, les no trobades a l'Excel i les
  sense GIA (no verificables). Lector de Camp Info autònom (`_ReadExcelCampInfoPerGia`
  + `_FindCampInfoPairs`, pura + tests) — NO dot-sourceja `rutes/Ruta.ps1`.
- **Word a PDF (i signar)**: la secció sencera —AutoFirma, el caixetí, el
  reempaquetat del CMS i les sis rondes sobre la validesa— és a
  **`suport/documentacio/signatura-pdf.md`**.
- **Selector de carpetes MODERN (a tot el programa)** (`_PickFolderModern`,
  `UiComuns.ps1`): el botó "..." de `_AddConfigRow` obre el diàleg **IFileOpenDialog**
  amb `FOS_PICKFOLDERS` (estil Explorer: barra d'adreça on es pot **enganxar la ruta**,
  panell lateral d'unitats/xarxa, cerca), en lloc del `FolderBrowserDialog` clàssic
  (arbre bàsic). Les interfícies COM (`IFileOpenDialog`/`IShellItem`) es defineixen per
  `Add-Type` (C#) i es compilen EN VIU el primer cop (mai en headless); si res falla,
  **fallback** al `FolderBrowserDialog` de sempre. Com que `_AddConfigRow` el fan servir
  TOTS els selectors (Configuració, Word a PDF…), el canvi és automàtic arreu.
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
  a `_ResolveOutputDir` (`local\informes-generats\`), amb finestra de progrés +
  Cancel·lar. Si l'item de control periòdic no és a REQ1, l'informe d'aquella
  activitat s'omet amb avís. Funcions pures testejades: `_ControlCatalegKind`,
  `_ControlSectionTitle`, `_FindItemKeysByTitle`.
- **Avisar titulars per correu (esborranys a Outlook)** (`suport/ControlsCpEmail.ps1`,
  botó **"Enviar correu (esborranys)"** a la finestra de Controls periòdics): per a
  les activitats **marcades** (mateixa columna "Generar"/`.Sel`), crea un correu per
  titular avisant que constava un **control periòdic** a passar (data prevista) per
  la seva activitat/adreça. Els correus i les dades surten de l'Excel (`RaoEmail`,
  `RepEmail`, `Adreca`, `ActPrincipal`, `ProperCP`, `DataControlPer`…). Fa servir
  **Outlook per COM** (`CreateItem(0)` → `.Save()`): deixa **esborranys** a la
  carpeta *Esborranys*; **MAI** `.Send()` (l'usuari revisa i envia). Titular a
  **Per a**, representant a **CC** (`_ControlsCpRecipients`; si en falta un, l'altre
  passa a To; les activitats sense correu vàlid es llisten com a omeses). El text és
  **editable** (assumpte + cos amb variables `{ACTIVITAT}{ADRECA}{ID_GIA}{TITULAR}`
  `{PROPER_CP}{DATA_CONTROL}{DATA}` + `**negreta**` + `//cursiva//` + enllaços) via
  el botó **"Editar text"** (`Invoke-ControlsCpEmailTextos`, mateix patró que *Textos del correu*),
  però es desa **LOCALMENT** a `%LOCALAPPDATA%\InformesCornella\controls-cp-email.json`
  (fora del repo: cap dada personal, sobreviu a `Actualitzar.bat`). Funcions pures
  testejades: `_DefaultControlsCpEmail`, `_ControlsCpRecipients`, `_FillControlsCpPh`.
  L'HTML del cos el fa **`_CosAHtml`** (`EnviarCorreu.ps1`), la mateixa que els
  recordatoris — i per tant aquest correu **també accepta `//cursiva//`**, cosa
  que la còpia pròpia que hi havia no feia. Outlook (COM) i finestres només a
  Windows. Es dot-sourceja a `GenerarInforme.ps1` després de `ControlsPeriodics.ps1`.

## L'editor de catàlegs: desar era lent i el Tipus estava bloquejat

Dues queixes de l'usuari, i totes dues tenien la mateixa arrel —fer feina cara al
lloc equivocat.

- **Desar trigava 10-15 segons** perquè `_Ed_SaveDoc` obria el **Word** i
  redibuixava el catàleg sencer per COM **a cada desat**. Mesurat: el JSON
  (model→objecte→text→validació) són ~350 ms en pwsh 7; la resta és Word.
  - Ara la vista es marca **pendent** (`$state.VistaPendent`) i es refà **en
    segon pla en tancar l'editor** (`_Ed_RefrescaVistes` llança
    `GeneraVistes.ps1` amb `Start-Process -WindowStyle Hidden`). Com que
    `Invoke-ExportarVistesWord` ja mira quins JSON són més nous que la seva vista
    (`_VistaCalRegenerar`), refà **només** el que s'acaba de desar; i si no arriba
    a passar, l'`Actualitzar.bat` les torna a mirar al pas 4b.
  - **`GeneraVistes.ps1` té ara un mutex** (`Global\InformesCornella.GeneraVistes`):
    ara el poden llançar l'editor i l'`Actualitzar.bat`, i **dos processos
    conduint el Word alhora** (`Documents.Add` + `SaveAs`) és la manera de treure
    una vista a mitges. Espera fins a 2 minuts; corre en segon pla i no bloqueja
    ningú.
  - **Les cometes les posem nosaltres** al `-File`: `Start-Process
    -ArgumentList` no enquota (la trampa de sempre) i el clone té espais.
  - La validació ja no escriu cap fitxer temporal: `_LoadEstructuralJson` accepta
    **una ruta o un objecte ja parsejat**, i l'editor li passa el que acaba de
    serialitzar. S'estalvia un `ConvertFrom-Json` sencer del fitxer.
- **El desplegable «Tipus» sortia bloquejat** sempre que el pare només admetia un
  tipus (un ítem dins d'una subsecció, un subítem dins d'un ítem). Dues coses:
  1. `_Ed_TipusOptions 'cataleg' 'subseccio'` només oferia `item`, però el lector
     **sí** que llegeix un `text` dins d'una subsecció (n'hi ha un a REQ1). Ara
     ofereix `item, text` i el combo es desbloqueja allà.
  2. El que faltava de debò era **canviar de nivell**: `← Treure` i `→ Ficar`
     (sota l'arbre) treuen el node del seu pare o el fiquen dins del germà de
     sobre, ajustant-ne el tipus (`_Ed_TipusEnMoure`). És el moviment que calia
     per fer el canvi de la secció anterior i que obligava a editar el JSON a mà.
  - La feina va a `_Ed_MouNivell`, **pura** (només toca el model) i provada sense
    Windows; `_Ed_CanviaNivell` només hi posa el missatge. `_Ed_TrobaPare`
    retorna **un hashtable**, no una col·lecció, per no caure al desenrotllat del
    pipeline.
  - El combo porta un **tooltip** que diu per què està bloquejat i on és la
    sortida: bloquejar un control sense explicar-ho és el que feia que semblés
    que el programa no deixava fer-hi res.

## Recordatoris periòdics als titulars (eina EINES)

`suport/Recordatoris.ps1` + `suport/EmailQuota.ps1` + `suport/RecordatorisAuto.ps1`.
Rajola 🔔 *Recordatoris* a EINES (acció `recordatoris`). Avisa periòdicament els
titulars amb tràmits pendents, a partir de l'`estat_actual` de la base d'informes.

- **DUES CAMPANYES INDEPENDENTS dins d'UNA sola eina** (decisió de l'usuari):
  `requeriments` (estat `Requeriment`) i `precintes` (estat `Precinte / Cessament`),
  cada una amb encesa/apagada, periodicitat, espera inicial, topall per tanda,
  mode (manual/automàtic) i **text propi**. Es defineixen en UN SOL LLOC
  (`_RecCampanyes`) — la finestra i l'execució automàtica hi beuen, així no es
  poden desincronitzar — i hi ha prova que els seus estats són **disjunts**.
- **La decisió de "a qui li toca" és PURA** (`_RecToca` / `_RecDueActivitats`) i
  per tant es prova a Linux, que és tot el sentit d'haver-la separada de la
  finestra. Ordre de les regles: sense GIA → fora (sense GIA no hi ha correu a
  l'Excel; es compta a part, **mai en silenci**); exclosa a mà → fora; **espera
  inicial** (el termini del requeriment encara corre); **periodicitat**.
  La data surt de **`_InformeQueDeterminaEstat`** (`Informes.ps1`): l'estat i la
  data han de venir del MATEIX informe, si no el correu diria una data que no
  lliga amb el que s'hi explica.
- **LA QUOTA D'EMAILJS ÉS EL CONDICIONANT DE TOT.** El pla gratuït són 200
  correus/mes. `EmailQuota.ps1` en compta **150** (reserva de 50) i qui hi suma
  és **`Send-EmailJs`**, no cada eina: així hi entren TOTS els enviaments del PC
  (l'eina *Enviar correu* i els recordatoris), que és l'única manera que el
  topall protegeixi de debò. Dues limitacions dites a la interfície: el mes
  d'EmailJS es reinicia el **dia de facturació**, no l'1 (la reserva de 50 és el
  coixí), i els correus enviats **des del mòbil** no es poden comptar des del PC.
- **LA BASE D'INFORMES DESFASADA ÉS EL RISC REAL**, no la quota: amb un
  `informes-db.json` vell s'escriuria a titulars que **ja han complert**, i això
  no es pot desfer. Per això la finestra ensenya l'antiguitat i **avisa en
  vermell** a partir de 30 dies, i el mode automàtic **es nega a enviar res** a
  partir de 45 (`$Script:RecMaxAntiguitatDbDies`), ho apunta al registre i surt.
- **Es desa DESPRÉS DE CADA enviament** (no al final de la tanda): si peta o es
  cancel·la, el que ja ha sortit consta i no es torna a enviar. El `try/catch` va
  **dins** del bucle (lliçó de la signatura), però un **401/403 atura la tanda
  sencera**: si les claus no valen, els 14 correus següents fallaran igual i no
  té sentit cremar-los.
- **L'Excel es carrega UNA vegada per tanda** (`Initialize-ActivitatsCache` +
  `Get-ActivitatFromCache`). `_CorreuEmailsActivitat` obre l'Excel a cada crida i
  serveix per a UN correu; en una tanda de 15 seria inviable.
- **Els destinataris els munta `_CorreuDestinatarisPerDefecte`** (`EnviarCorreu.ps1`),
  que ja combina *Raó soc. E-mail* + *Rep. Leg. E-mail* i dedupe. La plantilla
  d'EmailJS **no té camp CC**: les dues adreces van juntes a `to_email` separades
  per coma. El BCC **no consumeix quota** (una crida = un correu).
- **L'HTML del cos viu en UN sol lloc**: `_CosAHtml` (`EnviarCorreu.ps1`) fa els
  `<div>` per línia i cada línia passa per **`_TextToHtml`**, que escapa i aplica
  `**negreta**`, `//cursiva//` i l'autoenllaç. Abans n'hi havia **dues còpies
  idèntiques línia a línia** —aquesta i `_ControlsCpEmailHtml` dels controls
  periòdics, amb el mateix estil inline— i una tercera funció de línia
  (`_ControlsCpLineHtml`) que feia el mateix **sense cursiva**. Ara els dos
  correus passen per la mateixa. Hi ha guard que l'estil inline no es torni a
  copiar.
- **ELS URLs S'APARTEN ABANS DE MIRAR LA CURSIVA, i és un defecte real que hi
  havia.** La cursiva és `//...//` i un `https://` en porta un `//` a dins: amb
  **dues adreces a la mateixa línia**, l'expressió es menjava tot el tros d'una a
  l'altra i les destrossava totes dues —
  `Mira https://a.cat i tambe https://b.cat` → `Mira https:<i>a.cat i tambe
  https:</i>b.cat`—. No era hipotètic: `_TextToHtml` ja la feien servir els
  recordatoris i el text el pot editar l'usuari. Ara cada URL es substitueix per
  una marca amb caràcters de control (que cap de les dues expressions toca) i es
  torna a posar, ja com a enllaç, al final. Hi ha prova.
- **Dues coses del text estan blindades amb proves** perquè no es puguin perdre
  editant-lo: l'**avís de «si ja ho heu presentat, no en feu cas»**
  (`_RecAvisJaPresentat`, bilingüe) i l'**article 5 de l'Ordenança**
  (`_RecArticle5`, literal, versió vigent des del 19/06/2025). Viuen en funcions
  pròpies justament perquè una prova els pugui vigilar.
- **La campanya neix APAGADA i en manual** (`_RecDefaultConfig`): una eina que
  envia correus a ciutadans no es pot activar sola en actualitzar el programa.
  Hi ha prova que ho vigila.
- **Mode automàtic**: `RecordatorisAuto.ps1` és headless (patró de
  `mobil/Vigilant.ps1`: `$MotorSenseGui = $true`, una passada i surt) i el llança
  una **tasca del Windows** (`schtasks`) que es crea des del botó *Automàtic...*.
  `_RecSchtasksTr`/`_RecSchtasksArgv` són pures i **enquoten les rutes**: el clone
  té espais i `Start-Process -ArgumentList` no enquota (trampa de sempre).
- **On es desa**: `%LOCALAPPDATA%\InformesCornella\recordatoris.json` (config +
  historial) i `emailjs-quota.json`. **MAI al repositori**: porten ID GIA i dates
  d'enviament, i el repositori és PÚBLIC. A `%LOCALAPPDATA%` i no a `local/`
  perquè han de sobreviure a tornar a clonar.
- `_RecHistorialAMapa` / `ConvertTo-Mapa` (Json.ps1) desfan el que fa `ConvertFrom-Json`
  (PSCustomObjects): l'historial s'indexa per GIA i sense això `.ContainsKey` no
  existiria i **es perdria tot en silenci**. Hi ha prova d'anada i tornada **amb
  el JSON pel mig**, que és on aquest projecte s'ha trencat sempre.

## Normativa (eina EINES, setembre 2026)
Petició de l'usuari: tota la normativa de REQ1 **i la dels seus marcadors de
Chrome** (carpeta *Aj. Cornellà*) en una sola carpeta, classificada pel nom
(«Vector ambiental_Residus_Decret XXX»). Preguntat i confirmat: rajola al menú,
**text consolidat**, índex en Excel, que s'actualitzi sola, l'any al nom i
l'enllaç des de la fitxa d'ajuda.
- **`suport/normativa.json`** (al repositori, sense res personal: hi ha guard
  contra enllaços de OneDrive i adreces): ~165 normes amb `Ambit`, `Tema`, `Any`,
  `Tipus`, `Num`, `Titol`, `Url`, `Derogada` i, per a les que no tenen número
  (ordenances, DB del CTE…), `Claus`. Es va generar amb un script de la sessió a
  partir de les fitxes de REQ1 (que ja porten URL ELI) i dels marcadors.
- **`NormativaDades.ps1`** (pur): el nom del fitxer, la font (`boe`/`pdf`/`web`/
  `manual`), la informació de la pàgina del BOE (id, data d'última actualització,
  PDF original), quan cal tornar a baixar, l'índex `.xlsx` **fet a mà en
  OpenXML** (sense Excel; enllaços amb la fórmula `HYPERLINK` relativa) i com es
  reconeix una norma dins d'un text (`_NormativaNormText` porta «Real Decreto»,
  «Reial Decret» i «RD» a la mateixa forma; el número ha d'anar sencer).
  **Està partit de `Normativa.ps1`** perquè `Show-Ajuda` (`UiComuns.ps1`) en
  necessita la cerca i, en un sol fitxer, `UiComuns` i `Normativa` dependrien
  l'un de l'altre (el guard de cicles ho va enxampar).
- **`Normativa.ps1`** (només Windows): les baixades i la finestra. BOE → pàgina
  → `/buscar/pdf/<any>/<id>-consolidado.pdf` (o l'original si no n'hi ha).
  Portal Jurídic, CIDO i altres pàgines → **Edge headless `--print-to-pdf`** amb
  un **perfil propi** (`--user-data-dir`: si l'Edge de l'usuari és obert, sense
  això l'ordre se li'n va a ell i torna sense fer res). Un PDF de menys de 20 KB
  es dona per fallat (pàgina no carregada).
- **No s'ha pogut provar cap baixada des d'aquí**: el proxy d'aquest entorn
  bloqueja el BOE, el Portal Jurídic, EUR-Lex i cornella.cat. L'índex diu el
  motiu de cada fallada perquè l'usuari ho pugui passar.
- **Actualitzar sola**: BOE per la data d'última actualització; la resta, cada
  `$Script:NormativaDiesRefresc` (180) dies. L'anterior va a `anteriors\` amb la
  data si de debò canvia (BOE: nova data; web: la mida canvia més d'un 2 %,
  perquè imprimir la mateixa pàgina dues vegades no dona els mateixos bytes).
- Proves a `tests/proves/09-normativa.ps1`, entre elles que **la norma de cada
  fitxa d'ajuda de REQ1 és al catàleg** (validada traient el Decret 197/2016).
  El punt nou de tatuatge **no té fitxa d'ajuda**, o sigui que aquesta prova no
  el mira; el Decret 90/2008 hi és igualment (és als marcadors).
- **Segona ronda (setembre 2026)**, de l'usuari:
  - **El Portal Jurídic té un botó «PDF»** («Descarrega PDF RDF TTL XML»). Ara,
    per a tot el que no és del BOE, `_NormativaBaixaWeb`: (1) si l'URL ja és un
    PDF, aquell; (2) el **botó PDF de la pàgina** (`_NormativaPdfsDeHtml`, pura:
    primer l'enllaç que es diu «PDF», després els `.pdf`/`format=pdf`; mai
    RDF/TTL/XML ni el **resum fet amb IA** del costat, «Descarrega (CA)»), buscat
    a l'HTML del servidor i, si no hi és —la pàgina es munta amb JavaScript—, al
    DOM que torna l'Edge amb `--dump-dom`; (3) si res, la pàgina impresa. L'estat
    apunta la **`Via`** i l'índex diu «pàgina desada com a PDF» quan s'ha hagut
    d'imprimir: és el senyal que el PDF de la norma no s'ha trobat.
  - Les dues normes sense enllaç ja en tenen: l'Ordenança d'activitats (anunci
    del BOPB) i l'ordenança tipus d'olors de la Diputació (document directe).
  - **Guies i manuals** dels marcadors (`Guia: true`, tema **Guies** de cada
    àmbit; ~27). Fora les pàgines que són índexs de documents (industria.gob.es,
    Interior…) i la **UNE 123001** (còpia d'una norma UNE amb drets en un blog).
  - La fitxa de Farmàcies de REQ1 deia «Decret 40/2006»: és el **Decret
    40/1992** (confirmat per l'usuari). Fitxer d'or de la vista refet.
  - `_NormativaFilesIndex` tornava la fila desfeta quan només hi havia una norma
    (`return` sense coma): ho va enxampar la prova de les guies.
