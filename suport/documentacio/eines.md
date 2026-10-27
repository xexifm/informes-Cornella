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

### El Seguiment AUTOMÀTIC i el correu amb el PDF (octubre 2026)

L'usuari: «automatitzar l'eina Seguiment cada dues setmanes + enviament del
document adjunt per correu», «només els PDF, tot menys la fulla Estès».

- **«Cada dues setmanes»** és una freqüència nova de tots els automatismes
  (`quinzena` a `ModeAutomatic.ps1`, i al desplegable de Configuració →
  Automatismes, `$Script:FreqsAuto`): un dia de la setmana, les setmanes
  **parelles des d'una data fixa** (`$Script:AutoAncoraQuinzena`, el dilluns
  12/10/2026). És una data fixa i no «dues setmanes des de l'última»: si un dia
  es fa tard, la següent no es mou. La tasca del Windows dels Recordatoris
  també ho sap fer (`WeeksInterval`). El Seguiment ve per defecte cada dues
  setmanes, el dilluns a les 13:00.
- **La rajola porta l'interruptor A/M** (`$Script:ModesAuto['seguimentgia']`,
  `SeguimentGiaAutomatic.ps1`). La passada corre en segon pla
  (`SeguimentGiaAuto.ps1`): fa el PDF amb **la mateixa funció que el botó**
  (`_SgGeneraFitxer`, que ara no obre cap finestra; `_SgExportar` només hi
  posa els missatges), amb tots els llistats **menys la fulla Estès**
  (`_SgSeleccioAuto`), i l'envia.
- **El correu** (`Send-SeguimentCorreu`) va a les **adreces fixes** de l'eina
  `seguiment` (Configuració → Configuració correu electrònic), **només per l'Outlook**,
  perquè porta l'adjunt. Sense destinataris no s'envia res, i l'interruptor no
  s'engega (`Requisit`). Si va a Esborranys, el menú avisa com amb els
  recordatoris. El text es tria a *Textos del correu* i es desa a
  `seguiment-correu.json` (`%LOCALAPPDATA%`). Les variables són `{DATA}`,
  `{RESUM}` (les activitats de cada llistat) i `{ORIGEN}`.
- **El botó «PDF i enviar»** de la finestra fa el mateix ara, amb les pestanyes
  marcades.
- **Per què el correu és a `SeguimentGia.ps1`** i no a l'automàtic: el fan
  servir tots dos, i si fos a l'automàtic, aquest fitxer i aquell dependrien
  l'un de l'altre (ho va enxampar el guard de cicles).

## Normativa: només la que es cita; la resta, a `derogades` (octubre 2026)
La primera revisió de debò va donar 62 files i quasi cap era feina de debò:
- **Falses «derogades» del BOE.** `_RevEstatBoe` buscava «norma anulada» o
  «fue derogada por» a **tot** el text, i el preàmbul en parla d'altres normes
  (REBT, gas, alta tensió, ascensors, Ley 34/1998). Ara només compten les marques
  d'estat: la capçalera «Norma derogada, con efectos de…» i l'etiqueta
  «[Disposición derogada]». L'única de debò era el RD 1836/1999 (radioactives),
  derogat pel RD 1217/2024: canviada la fitxa i afegit a `normativa.json`.
- **Quasi tot el Portal Jurídic «no s'ha pogut saber»**, sense motiu. Si l'Edge
  es penja amb la primera norma, `$Script:NormativaEdgeKO` el deixa de fer servir
  per a tota la web. Ara l'Excel diu el motiu (`_RevMotiuPjur`,
  `$Script:NormativaDomEdgeError`). La segona revisió (13:29) ho va confirmar:
  les 7 primeres, l'Edge torna la pàgina però **sense** l'etiqueta on la busca
  `_RevEstatPjur`; a la 8a es penja (45 s) i la resta queden «L'Edge es va penjar
  amb portaljuridic.gencat.cat». Des d'aquí no s'hi arriba (proxy 403), o sigui
  que `suport\Provar-Vigencia.bat` (`ProvarVigencia.ps1`) desa al PC de
  l'usuari la pàgina del servidor, les metadades, els scripts de la pàgina
  (l'aplicació ha de cridar alguna API amb la vigència), el DOM de l'Edge amb
  més temps, i del BOE la pàgina i les dades obertes (`metadatos`) de quatre
  normes, en un `.zip` a `local\revisions`.
- **Normativa que no cita ningú** (marcadors de Chrome). Decisió de l'usuari:
  *«Quan una norma (o guia) ja no es cita a REQ1 no cal revisar i la mous a
  derogades»*, i les derogades també (opció B: s'acaba el tema «Antic» com a
  lloc on guardar-les). `_NormativaSepara` (pura, `NormativaDades.ps1`): fora les
  `Derogada` i les que no cita **cap** catàleg (`Get-NormativaTextCatalegs`, tots
  els JSON d'ESTRUCTURALS: Llicència en cita alguna que REQ1 no); es queden
  sempre les **guies** i **col·leccions** (l'usuari: «s'ha de tractar com a
  normativa») i les que no tenen cap clau per reconèixer-les (no es pot saber).
  `normativa.json` **no es toca**: és tot el que es coneix, i una norma que es
  torni a citar torna sola. `Invoke-NormativaBaixada` rep la llista sencera,
  mou el PDF de les retirades a `local\normativa\derogades`
  (`_NormativaMouRetirades`, mai esborra) i baixa i indexa només les actives. La
  revisió de vigència només mira les actives.
- `_EnllacosDeCataleg`: un sub-punt sense títol porta el nom del punt de sobre
  (els «No es disposa…» de Llicència sortien amb el punt en blanc).

## REQUERIMENT ANTERIOR / REQUERIMENT ACTUAL i els punts A1, A2… (octubre 2026)
L'usuari enganxava l'anterior requeriment al lloc de *COPIAR REQUERIMENT* i
l'informe quedava mal numerat (1, 1, 2, 2, 3…: el text enganxat duia la
numeració automàtica del Word i el punt «S'ha de donar resposta…» consumia
l'1) i amb seccions repetides. Primer es va proposar fusionar l'anterior amb
l'actual (reconeixent els punts vells al catàleg) i l'usuari ho va descartar:
*«És complicar-ho massa… Ben diferenciat el requeriment anterior i l'actual,
copiaré el text literal i en comptes d'1 i 2 posaré A1 i A2»*.
- **`Build-CatalegBlocs`** (`MotorInforme.ps1`): la secció que porta
  `COPIAR REQUERIMENT` en algun punt (`_EsSeccioReqAnterior`; per la marca, no
  pel títol, que l'usuari pot canviar a l'editor) surt com a `titolbloc`
  **Requeriment anterior** amb les línies del punt **sense número**
  (`_BlocsReqAnterior`), i abans de la secció següent hi va el `titolbloc`
  **Requeriment actual** + la frase d'intro (`_BlocsReqActual`). La numeració de
  l'actual comença per l'1. Només a l'informe: la vista del catàleg
  (`-SenseCamps`) no hi entra. Afecta REQ1 (l'únic catàleg amb la marca) i MNS/Transmissió (sense frase
  d'intro); **Llicència no** (té el seu muntatge amb lletres A., B.…).
- El **correu** (PC i mòbil) pinta el `titolbloc` subratllat, i `docs/correu.js`
  té la mateixa regla (`esSeccioReqAnterior`…): la prova creuada hi passa una
  secció d'anterior.
- El **catàleg** REQ1: el *COPIAR REQUERIMENT* porta el recordatori de canviar
  1, 2… per A1, A2…, i el «I a més a més, esmenar els següents punts:» ja no hi
  és (el substitueix el bloc REQUERIMENT ACTUAL).
- **Quants punts tenia l'anterior** (octubre 2026, substitueix el recordatori
  de l'A1, A2…). L'usuari: «si marco que tenia 4 requeriments, els següents
  començaran per 5, 6, 7… en comptes de per 1, 2, 3». És un `[CAMP:]` normal
  en una línia pròpia de l'*Anterior requeriment* («Punts del requeriment
  anterior: [CAMP: …]»), o sigui que surt sol a la pantalla de *Text i
  opcions* (PC i mòbil). El motor reconeix la línia **ja resolta** pel prefix
  (`_PuntsAnteriorDeLinia`, `$Script:PrefixPuntsAnterior`), no la posa al
  document i en fa començar el `$num` de `Build-CatalegBlocs`. Pel prefix i no
  pel `[CAMP:]` cru a posta: `docs/correu.js` només veu les línies resoltes, i
  la prova creuada hi passa una anterior amb 4 punts (l'actual ha de començar
  pel 5 als dos costats). Llicència (punts amb lletres) també salta la línia.
- **Seguiment** (`Seguiment.ps1`): `$SeguimentReqRegex` accepta `A1.`, `A2.`…
  (`Number` = `'A1'`, text; els numèrics continuen sent enters), i
  `_SeguimentParentTopic` i `_CorreuSenseNegretaSeguiment` també.

## «Requeriment - Seguiment»: una paraula en negreta no fa pendent un punt resolt

L'estat de cada punt viu **al mateix `.docx`**: el seguiment escriu la data
sense negreta i el comentari **tot en negreta si queda pendent** i sense
negreta si es resol. Abans es mirava la negreta del **paràgraf sencer** i un
paràgraf **mixt** comptava com a pendent. Cas real (octubre 2026):
«16/09/2026: S'aporta. L'horari serà **DIÜRN** (fins les 23h).» —l'usuari hi
havia destacat una paraula— es llegia com a pendent, el punt sortia sense
«Resolt» i el seguiment següent hi va escriure «No s'aporta.» en negreta.

- Ara ho decideix **`_AnotacioPendent`** (pura, `Seguiment.ps1`) amb els trossos
  del paràgraf (`_ParagraphTrossosXml`, `Docx.ps1`): mira **només el
  comentari** (treu la data) i és pendent si **la majoria de les lletres** van
  en negreta. Per majoria i no «tot» perquè també aguanti el cas contrari (un
  pendent on algú ha tret la negreta d'una paraula). Sense comentari, mana el
  paràgraf.
- **I ja no s'esborra la negreta de l'usuari**: en un seguiment nou només es
  des-negreten les anotacions anteriors que eren **pendents** (la marca del
  programa). Abans es des-negretaven totes i la paraula destacada es perdia.
- Proves a `01-motor.ps1`, validades tornant a posar les dues regles antigues
  (3 FAIL).

## «Enviar correu»: QUALSEVOL informe, triat (octubre 2026)

L'usuari: «no ha de ser de l'últim informe fet, sinó que ha de poder ser de
qualsevol informe fet». L'eina enviava sempre el `.docx` més recent de la
carpeta sense preguntar. Ara obre un selector (`Select-InformeDocx`,
`EnviarCorreu.ps1`) amb el més recent ja triat (`_InformeDialegInici`, pura): un
Intro i es fa com abans. El *Seguiment* hi tria el seu informe anterior
(`-SenseUltim`). En acabar de generar un informe (`Offer-EnviarCorreu`) no es
pregunta: és el que s'acaba de fer. Guard a `04-correu.ps1`.

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
  - **La convenció ja hi era i no es va seguir**: els granats de la marca
    s'omplien **dins d'un `if (-not $Script:HeadlessTest)`**. Des de l'octubre
    de 2026 la paleta sencera és `Initialize-BrandColors` (`UiFinestra.ps1`), i
    **el guard de headless és al CRIDADOR**: dins d'una funció el tipus es
    resol en *compilar-ne el cos*, abans de la primera línia, o sigui que un
    `if` a dins no aturaria res (la trampa de `_BuildCaixetiImageBase64`).
    (`$Script:ConfigUiAccent` ja no existeix: era un cinquè literal de la
    paleta que no llegia ningú.)
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

## Per on surten els correus: EmailJS o l'Outlook (octubre 2026)

L'usuari: *«Es poden tenir les dues opcions? Poder triar entre EmailJS per no
desfer-ho i Outlook per fer proves?»*. Conviuen, i tot el que fa sortir un correu
viu a **`suport/CorreuVia.ps1`**. Tres vies (`$Script:CorreuVies`):

| Via | Què fa |
|---|---|
| `emailjs` (per defecte) | com sempre: la plantilla d'EmailJS i la quota de 200/mes |
| `outlook-esborrany` | l'Outlook de l'ordinador, i el correu es queda a **Esborranys**: no s'envia res |
| `outlook` | l'Outlook de l'ordinador, i **s'envia** (surt de la bústia de l'usuari) |

- **Una sola preferència per PC** (`CorreuVia` a `settings.json`), que es tria a
  **⚙ Configuració** (requadre «Correus que s'envien des d'aquest PC») o al
  desplegable **«Enviar amb:»** del diàleg d'*Enviar correu* (canviar-lo allà
  també la desa). No és un interruptor A/M de rajola: la rajola de Recordatoris
  ja en porta un i voldria dir una altra cosa.
- **`Set-CorreuVia` no trepitja la resta de `settings.json`**
  (`_SettingsAmbCorreuVia`, pura i provada): el diàleg d'enviar no pot esborrar
  les carpetes ni els automatismes. La via per defecte no s'hi escriu, i una via
  desconeguda (un `settings.json` escrit a mà) torna a EmailJS
  (`_CorreuViaValida`): mai un PC sense correu.
- **Una «sessió» d'enviament** (`Open-` / `Send-` / `Close-CorreuSessio`): l'Outlook
  s'obre UN cop per tanda, no un cop per correu. **No es fa mai `Quit()`**
  (tancaria l'Outlook de l'usuari, com ja diu Controls periòdics); si l'Outlook
  no era obert i s'ha enviat alguna cosa, en tancar se li demana
  `SendAndReceive` perquè no es quedi a la Safata de sortida.
- **L'Outlook s'obre NOMÉS a `New-OutlookApp`**, com el Word a `Motor.ps1` i
  l'Excel a `Excel.ps1`: Controls periòdics també hi passa. Guard a
  `06-guards.ps1` (validat injectant un `New-Object -ComObject` a
  `ControlsCpEmail.ps1`).
- **Per què `CorreuVia.ps1` i no `EnviarCorreu.ps1`**: les claus d'EmailJS i
  `Send-EmailJs` vivien a `EnviarCorreu.ps1` i, amb la tria, els dos fitxers
  haurien depès l'un de l'altre (el guard de cicles ho va aturar). Ara
  `CorreuVia.ps1` és **la porta del servei** (també a la llista de fitxers que
  poden fer peticions HTTP, `09-normativa.ps1`).
- **Només Outlook clàssic**: el «nou Outlook» de Windows no es pot controlar per
  COM. Si no hi és, el missatge diu de tornar a EmailJS a Configuració.
- **Per què EmailJS segueix sent el per defecte**: l'enviament directe per
  l'Outlook pot fer saltar l'avís de seguretat («un programa intenta enviar
  correu») o el pot bloquejar la política d'informàtica. Desar esborranys no el
  fa saltar (per això Controls periòdics només en desa: hi ha guard que no
  obri la sessió en cap altre mode ni faci cap `.Send()`).
- **Els recordatoris, també els AUTOMÀTICS, segueixen la via triada** (primer
  anaven sempre per EmailJS; l'usuari: *«vull que Recordatoris també funcioni
  amb Outlook, que salti un avís al programa que els correus estan a
  Esborranys»*). La tasca del Windows corre amb la sessió iniciada, que és el
  que l'Outlook necessita. Si l'Outlook no era obert, en acabar se li demana
  `SendAndReceive` i s'espera (fins a 60 s) que la Safata de sortida es buidi.
- **Un recordatori desat a Esborranys CONSTA a l'historial**, com un d'enviat.
  La primera versió no l'hi apuntava (raonament: no ha sortit), però amb
  l'automàtic això vol dir que **l'endemà la tasca en tornaria a desar un altre**
  per al mateix titular, i així cada dia. Que s'acabin enviant ho vigila
  l'avís dels esborranys pendents.
- **L'avís dels esborranys pendents** (`Add-` / `Get-` / `Clear-CorreuEsborranysPendents`
  i `Show-AvisEsborranysSiCal`, `CorreuVia.ps1`): la tanda **automàtica** apunta
  els que ha desat a `%LOCALAPPDATA%\InformesCornella\correus-esborranys.json`
  (porten noms i adreces: mai al repositori), i el **menu** ho avisa en obrir-se
  i al seu rellotge de cada minut (la tasca en pot deixar amb el programa
  obert). *Sí* = ja els he enviat (es buida); *No* = calla fins que n'arribin de
  nous. Els manuals no s'hi apunten: el resum de la tanda ja ho diu en acabar.
  La marca «ja avisat» es posa **abans** del `MessageBox`: el rellotge segueix
  tocant mentre el quadre és obert i, si no, el repetiria. Guard: la tanda els
  apunta i el menú crida l'avís als dos llocs (validat traient-ne cada un).
- **EL REMITENT** (l'usuari: *«a la feina puc enviar correus des d'adreces
  diferents; com ho puc seleccionar per no haver de canviar-la a tots els
  correus?»*): una adreça per defecte a **⚙ Configuració → Configuració correu
  electrònic** («Des d'on surten (Outlook)», `Set-CorreuRemitent`;
  `CorreuRemitent` als settings; buit = el compte per defecte de l'Outlook). El
  botó *Comptes de l'Outlook* omple el desplegable amb els comptes del perfil
  (`Get-OutlookComptes`); una **bústia compartida** no hi surt i s'escriu a mà.
  La sessió busca el compte **un cop** (`_OutlookTriaCompte`): si és un compte,
  `SendUsingAccount` (amb el respatller d'`InvokeMember`, perquè assignar un
  objecte COM a una propietat COM falla en alguns PowerShell); si no,
  `SentOnBehalfOfName` (cal permís d'enviar en nom de la bústia a l'Exchange).
  Val per a Enviar correu, Recordatoris **i Controls periòdics**, que ara també
  passa per la sessió (amb el CC del representant: `Send-CorreuSessio` en té).
  **Amb EmailJS no compta**: el remitent és el del servei connectat a EmailJS.
  El diàleg d'*Enviar correu* diu «des de …» al costat de la via.
- **Dos defectes que hi havia a Recordatoris i que la tria va destapar**: la
  tanda cridava `Send-EmailJs $ecfg` amb **`$ecfg` sense definir** (des del
  commit `353c1a0`: tots els recordatoris, manuals i automàtics, haurien
  fallat), i la quota s'apuntava **dues vegades** per correu (`Send-EmailJs` ja
  ho fa). Totes dues arreglades.
- **No s'ha provat amb un Outlook de debò** (la suite corre a Linux): les proves
  fan servir un doble de COM que comprova que el mode esborrany **desa i no
  envia**, que les adreces van separades per `;`, que el remitent va a
  `SendUsingAccount` o a `SentOnBehalfOfName` segons si és un compte, i que en
  tancar s'envia i es mira la Safata de sortida.

## Els correus de CADA EINA (octubre 2026) — `CorreuEines.ps1`

L'usuari: «a Configuració s'ha de poder triar entre les opcions (Outlook
esborrany, Outlook enviament, EmailJS) segons quina eina s'estigui usant […]
també s'ha de veure clar a qui s'envien els correus […] i CCO […] un botó a
cadascuna de les opcions per enviar un correu de prova».

- **`$Script:CorreuEines`, el registre** (un sol lloc): `mobil` (Enviar correu,
  l'informe de visita), `rec-requeriments`, `rec-precintes`, `controls`,
  `seguiment`. Cada eina hi diu **quines vies admet**:
  - **Controls periòdics, només esborranys**: l'usuari els revisa; la regla i
    el guard de sempre es mantenen.
  - **Seguiment, només Outlook**: porta PDF adjunts i EmailJS no en sap.
    `Send-CorreuSessio` **peta** si se li passen adjunts per EmailJS, en lloc
    d'enviar el correu sense ells.

  Cada eina diu també a qui pot anar: titular, representant legal, persones
  autoritzades (de la base de contactes, `Get-CorreuAutoritzats`).
- **`correus.json` a `%LOCALAPPDATA%\InformesCornella`**, no `settings.json`:
  - el «Desar» de Configuració reescriu `settings.json` sencer;
  - **les adreces no poden anar mai al repositori**, que és públic. L'usuari
    les omple un sol cop a Configuració → *Configuració correu electrònic*.

  Per eina hi ha `via`, `dest`, `fixes` i `cco`. La `cco` **només si s'ha
  configurat**: sense, l'eina fa servir la de sempre (email-textos.json per a
  Enviar correu, la de la campanya per als Recordatoris). La pantalla no la desa
  si és la mateixa (com a conjunt). També hi ha la `prova`: el GIA de les dades
  i qui la rep.
- **La via de tot el PC** (`CorreuVia`, CorreuVia.ps1) queda com a valor per
  defecte de les eines encara no configurades: qui l'havia triada no veu cap
  canvi. El desplegable d'*Enviar correu* desa la via de l'eina `mobil`.
- **El correu de PROVA** (`Send-CorreuProva`) va **només** a l'adreça de prova,
  sense CCO. Al davant porta un requadre que diu a qui hauria anat de debò i la
  CCO. Cada eina registra el seu a `$Script:CorreuProves` (i la CCO d'abans a
  `$Script:CorreuCcoAbans`) **des del seu fitxer**: `CorreuEines.ps1` no pot
  cridar cap funció de les eines sense fer un cicle, perquè totes criden
  `Get-CorreuEina`. Per això es carrega **abans** que elles. La prova d'Enviar
  correu fa servir el darrer informe generat amb les dades de l'activitat de
  prova.
- **A qui va**: `_CorreuDestinataris`, que dóna l'ordre titular, representant,
  autoritzats i fixes, sense repetides. Controls periòdics hi posa el primer a
  «Per a» i la resta a «CC» (`_CorreuParteixToCc`, que substitueix
  `_ControlsCpRecipients`).
- **La CCO: «sempre» o «a triar cada vegada»** (octubre 2026). L'usuari: «afegir
  una opció de les que poso en CCO […] i triar si sempre (com el meu correu) o es
  pot triar cada vegada». Abans només *Enviar correu* oferia adreces per marcar
  (les no `def` de la clau `bcc` d'`email-textos.json`), i a *Correus de cada
  eina* només s'hi veien les marcades: **les altres no sortien enlloc**. Ara cada
  eina té `cco` (sempre) i `ccoTria` (a triar), totes dues a `correus.json` i
  totes dues amb la d'abans si no s'han configurat (`$Script:CorreuCcoAbans` /
  `$Script:CorreuCcoTriaAbans`; només `mobil` en tenia d'a triar).
  - **Sempre**: a tots els correus de l'eina, també als automàtics.
  - **A triar**: surten **desmarcades** quan s'envia **a mà**. A *Enviar correu*,
    a les caselles de la seva finestra (`_CorreuCcoOpcions`: primer les de
    sempre, marcades). A Recordatoris, Controls periòdics i Seguiment, en una
    finestreta abans de la tanda (`Select-CorreuCcoTria`, **`CorreuCcoTria.ps1`**:
    a part perquè `CorreuEinesPantalla.ps1` crida les eines i les eines la
    criden, i juntes farien un cicle). Sense cap adreça a triar no s'obre res.
  - **En automàtic, mai** les a triar: no hi ha ningú per marcar-les, i una
    finestra en una tasca del Windows es quedaria esperant sola. Guard a
    `06-guards.ps1` (validat injectant-ho).
  - El **mòbil** segueix amb la clau `bcc` d'`email-textos.json`: `correus.json`
    és d'aquest PC i no arriba a la web.
- **Tot el correu des d'aquesta pantalla** (octubre 2026, l'usuari: «s'ha de
  poder configurar tot el correu des d'aquesta pantalla»). Cada fila té
  **Editar el text**, que obre el mateix editor que *Textos del correu*
  (`$Script:CorreuTextosEditors`; hi ha prova que totes les eines del registre en
  tenen). Per fer lloc als dos botons, la finestra fa 1.100 d'ample
  (`_CorreuFilaGeometria`).
- **UN SOL BOTÓ A CONFIGURACIÓ** (octubre 2026, l'usuari: «tot el bloc que hi
  ha sobre Automatismes, que es converteixi en un botó "Configuració correu
  electrònic". I allà dins poses l'adreça des d'on s'envien els correus»). La
  finestra (abans *Correus de cada eina*) ara porta a dalt el **remitent**
  (`_CorreuCapGeometria`: remitent i correu de prova l'un al costat de
  l'altre) i el desa a `settings.json` amb `Set-CorreuRemitent`, sense tocar
  la resta. Com que el «Desar» de Configuració reescriu `settings.json`
  sencer, hi torna a posar la via i el remitent **tal com són en aquell
  moment** (`_SettingsAmbCorreuDeAra`, llegint el fitxer), no com eren en obrir
  la pantalla: s'han pogut canviar a la finestra del correu amb Configuració
  oberta.
- **L'APP DEL MÒBIL, UNA FILA MÉS** (l'usuari: «App del mòbil (web) no sé per
  què té un botó, quan podria ser un apartat més com la resta. Tampoc sé la
  diferència entre App del mòbil i Enviar correu»). És el **mateix correu** que
  la fila `mobil` (l'informe de visita, el mateix text), però l'envia l'app web
  per EmailJS; per això va just a sota (`_CorreuFilaAppMobil`) i les dues es
  diuen *Informe de visita · des del PC (Enviar correu)* i *· des del mòbil
  (app web)*. No té via ni adreces fixes ni prova; *Editar el text* obre el
  mateix editor, i *Prompt...* la finestreta del prompt
  (`Show-CorreuMobilPrompt`). En desar, `Save-CorreuMobilConfig` només toca el
  que ha canviat (`_CorreuMobilCanvis`, pura: la CCO com a conjunt): la CCO, al
  PC i al Drive; a qui va, a `email-textos.json` (relegit en aquell moment, per
  no trepitjar el que s'hagi editat amb *Editar el text*), i llavors ofereix
  Actualitzar.
- **L'ORDRE DELS AUTOMATISMES** a Configuració és el de les eines d'aquesta
  finestra (`_ConfigOrdreAutomatismes`, pura: clau de correu → automatisme) i
  després els que no envien correus, en l'ordre del registre.
- **El correu del MÒBIL des del programa** (`CorreuMobil.ps1`, la fila de
  l'app a la finestra del correu). El mòbil llegeix `docs/dades/email-textos.json`; s'hi tria
  **a qui va** (clau nova `dest`: `titular` i/o `representant`; sense la clau, el
  titular, que és el que feia; `docs/app.js` → `destinatariPerDefecte`, amb
  l'`EMAIL_REP` de l'activitat) i la **CCO sempre / a triar** (la clau `bcc` de
  sempre, amb el `def`). En desar, si ha canviat a qui va, ofereix Actualitzar,
  que committeja aquell fitxer; li arriba com a scriptblock des de Configuració
  (cridar-lo des d'aquí faria un cicle). Per al que viu a `docs/config.js`
  (EmailJS, el remitent) —que Actualitzar **no** puja— o al codi, **Prompt...**
  (`_CorreuMobilPrompt`, pura) fa un text amb la configuració d'ara i
  el canvi que s'hi escriu, per enganxar a Claude. *Textos del correu* conserva
  `dest` en desar (guard).
- **LES ADRECES DE CCO, FORA DEL GITHUB** (octubre 2026, l'usuari: «podem posar
  les adreces en local i fora de GitHub perquè no siguin públiques?»). Eren a la
  clau `bcc` d'`email-textos.json`, a la web pública. Ara:
  - al PC, a `%LOCALAPPDATA%\InformesCornella\correu-mobil.json`
    (`Read-CorreuMobilCco` / `Save-CorreuMobilCco`, `EmailTextos.ps1`); la clau
    `bcc` de `_LoadEmailTextos` en surt, o sigui que *Enviar correu* i el
    destinatari buit (`_CorreuDestinatariBuit`) funcionen igual;
  - al mòbil, del **Drive privat** (`Dades/correu-mobil.json`, al costat
    d'`activitats.json`): el PC l'hi puja en desar la fila de l'app (si la CCO ha canviat) i a cada
    Actualitzar (`Publish-CorreuMobilCco`, pas 6); `docs/drive.js` →
    `llegirCorreuMobil`, **`docs/cco.js`** (`Cco.muntar` / `Cco.seleccionat`).
  - **La CCO al mòbil, a `docs/cco.js`** (octubre 2026, l'usuari: «no em deixa
    escollir el CCO, m'hauria de sortir el CCO sempre i els CCO a triar»).
    Abans era a `app.js` i, si el Drive no responia o encara no hi havia el
    fitxer, el quadre sortia **buit sense dir res**. Ara: dos grups («CCO
    sempre», marcades i desmarcables; «CCO a triar», desmarcades); si el Drive
    no està connectat prova `reconnectarSilenci`; i si no hi ha adreces o no es
    poden llegir, ho diu (on es configuren al PC) amb un enllaç per tornar-ho a
    provar. Prova a `navegador/prova-mobil.mjs`. Fora d'`app.js` també perquè
    aquell fitxer és a tocar del seu sostre de 1.250 línies.
  - **«No hi ha cap adreça» amb la CCO posada al PC** (octubre 2026, amb
    captures). No es va poder veure el fitxer del Drive de l'usuari, o sigui
    que es va tancar per tots els costats: (1) el PC la **torna a pujar a cada
    Desar** de *Configuració correu electrònic* (abans només si canviava o a
    Actualitzar, i si aquell pas fallava no ho deia ningú) i, amb l'API del
    Drive, **la torna a llegir** i diu quantes adreces hi ha quedat
    (`_CorreuMobilCcoCompta`, pura); (2) el mòbil llegeix la llista en
    qualsevol forma que hi pugui haver (`Cco.llistaDe`: `{value, Count}` d'un
    PowerShell 5.1, un sol objecte, `Addr`/`Default`); (3) si la llista és
    buida, diu **de quan és el fitxer** (`modifiedTime`, `drive.js`) i que cal
    fer Desar al PC.
- **EL REPLY-TO** (l'usuari: «Reply to ha de ser suport.activitats@…»): va al
  camp **Reply To de la plantilla d'EmailJS** (al panell d'EmailJS), no al
  codi. Una adreça @aj-cornella.cat no pot anar a `docs/` ni al codi (guard a
  `04-correu.ps1`: el repositori és públic), i un paràmetre `reply_to` igualment
  no faria res si la plantilla no el fes servir. Com que el PC i el mòbil fan
  servir la mateixa plantilla, val per als dos.
- **MENYS INICIS DE SESSIÓ DE GOOGLE al mòbil** (`docs/drive.js`, octubre
  2026, l'usuari: «em fa iniciar sessió a Google moltes vegades»). Google no
  dóna *refresh tokens* a una app de navegador (el token dura ~1 h). El que ho
  empitjorava, i ara no:
  1. el token vivia a la memòria de la **pàgina**: del plànol a l'informe
     («Fer informe»), enrere o recarregar el perdia, i la renovació de
     l'arrencada no surt d'un toc (el mòbil bloqueja la finestra). Ara va a la
     **sessionStorage** fins que caduca (menys 2 min de marge); no a la
     localStorage, perquè no quedi al telèfon en tancar l'app;
  2. si la renovació silenciosa fallava, es tornava a demanar amb
     `prompt: "consent"` (la pantalla de permisos sencera). Ara una sola
     petició amb `prompt: ""` des del toc, i `"none"` a l'arrencada;
  3. amb més d'un compte al telèfon calia triar-lo cada vegada: es recorda
     (`drive/v3/about`, localStorage `driveCompte`) i es passa com a `hint`.
  Un 401 oblida el token. Prova amb el `drive.js` de debò i un Google simulat a
  `navegador/prova-mobil.mjs`. Tancar l'app del tot o passar d'una hora encara
  demana un toc.
- **Següent sense confirmar el titular** (mòbil): el motiu surt en vermell
  **a sobre del botó Següent** (`#msg-nav`, `avisNav` a `app.js`), no sota l'ID
  GIA, que quedava fora de la pantalla; el de «cap deficiència» també (abans era
  un `alert`). Se'n va en marcar la casella i en canviar de pas.
  - **La migració és sola** (`_CorreuMobilCcoMigra`): si el fitxer del
    repositori encara porta `bcc`, la copia al PC (si no n'hi ha) i la hi treu
    —també si el *Restore* d'Actualitzar l'hi torna a posar—; si ja ha baixat
    sense, la treu de l'**historial del git** (`_CorreuMobilCcoHistoriaGit`; a
    les proves, mai). Sense historial llegible no es dona per feta.
  - Guard: cap `@aj-cornella.cat` ni al codi ni a cap `.json`/`.html` de `docs/`
    (validat injectant-ne una). **L'historial del git encara les té**: treure-les
    d'allà vol dir reescriure `main` amb un force-push, i no s'ha fet.
- **La prova del Seguiment FA EL PDF EN AQUELL MOMENT** (octubre 2026,
  l'usuari: «s'envia l'últim arxiu que jo he creat a mà, però la gràcia és que
  aquell document es faci abans de ser enviat i que estigui actualitzat del
  mateix dia»). Abans adjuntava el darrer PDF de l'estat (`ultim_pdf`) o de la
  carpeta. Ara crida `_SgGeneraFitxer 'pdf' @(_SgSeleccioAuto)`, com
  l'automàtic, i n'usa el resum i l'Excel d'origen. Si no es pot fer, surt
  **sense adjunt** i diu per què, amb el cos sense les línies de dades
  (`_SgCosSenseDades`): mai un PDF vell. El mateix per a **Excel Errors**
  (`Invoke-ExcelErrorsLlistats`; abans, els tres `.xlsx` més nous de la
  carpeta). Guard a `06-guards.ps1`. Les dades són les de l'Excel d'activitats
  més nou de la carpeta (`Find-LatestActivitatsExcel`): el programa no el
  baixa del GIA.

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

### Els `.doc` antics i el Word en segon pla (octubre 2026)

**Què passava:** després de pujar la versió del classificador (que obliga a
tornar a classificar tots els informes), la passada **automàtica** obria amb
el Word, un per un, tots els `.doc` antics **directament de la unitat de
xarxa**. El Word hi escriu els seus temporals (`~WRS{…}.tmp`) i, a la primera
interrupció de xarxa, ensenyava «Word no puede abrir el archivo ~WRS…», «Word
ha realizado una lectura doble…» o «Word perdió datos debido a una conexión de
red…»: finestres que **no treu el `DisplayAlerts`**. Sense ningú davant, el
procés s'hi quedava esperant amb el mutex de la base agafat, i el botó deia
cada vegada «s'està actualitzant sola ara mateix».

**Com està ara:**
- `_ReadDocParagraphsWord` **copia el `.doc` al `%TEMP%`** i obre la còpia
  (amb contrasenyes de mentida perquè un document protegit doni error en lloc
  de preguntar). Mai l'original de la xarxa.
- El text de cada `.doc` es desa la primera vegada a
  `local\base-dades-activitats\informes-doc-text.json` (`_LiniesDeDoc`) i es
  torna a fer servir mentre el fitxer no canviï: un canvi de classificador ja
  no necessita el Word.
- **La passada automàtica no obre mai el Word** (`Invoke-InformesDbEscaneig
  … $ambWord = $false`, també per als PDF del repàs de contactes). Un `.doc`
  que no es pot llegir **no es classifica en blanc**: es queda el que ja se'n
  sabia, amb `no_llegit`, i el torna a provar la passada següent (el botó, que
  sí que fa servir el Word) encara que el fitxer no hagi canviat. Abans es
  desava sense conclusió i, com que no canviava, ja no es tornava a llegir mai.
- `ValidarClassificacio` fa servir el text desat de debò
  (`$Script:DocTextCacheDir`), no el de la seva carpeta temporal.
- El motor de la base d'informes és `suport/Informes.ps1`: escaneja `$InformesDir`
  (per defecte `...\5.- Sergi Fadurdo\Informes`) i, per cada informe (`.docx` o
  `.doc` antic amb data al principi del nom), en treu **data + ID GIA +
  conclusió**, agrupat per activitat (per GIA; si no en té, per **carpeta**), a
  `local\base-dades-activitats\informes-db.json` (dins de `local/`: mai es puja).
  Botons al menú: **🗃 Actualitzar** i **📋 Editar** (marc "Base d'informes").
  L'escaneig viu a `InformesEscaneig.ps1`: el **nucli sense finestres**
  (`Invoke-InformesDbEscaneig`) el comparteixen el botó (`Invoke-InformesDbScan`,
  amb la finestra de progrés) i el mode automàtic (`Invoke-InformesDbAuto`).
- **Actualitzar base, mode AUTOMÀTIC** (interruptor **A/M** de sota la rajola,
  octubre 2026: «ja que serà tan important per fer el Plànol activitats»). La
  mateixa regla i la mateixa hora que *Copiar informes* (`ModeAutomatic.ps1`):
  cada dia a les **13:00** amb el programa obert i, si l'última vegada que tocava
  no es va fer, en obrir el programa; en un **procés a part** (`BaseInformesAuto.ps1`), sense res a la
  pantalla. Estat a `informes-db-auto.json` (`auto`, `auto_el`, `mode`) —**no**
  dins de la base, que l'editor reescriu sencera. Registre a
  `%LOCALAPPDATA%\InformesCornella\informes-db-log.txt`.
  - **Un sol escaneig a la vegada**: el botó i l'automàtic agafen el mateix mutex
    (`$Script:BaseMutexNom`) i **no esperen**. Si l'automàtic el troba ocupat,
    apunta el venciment com a servit (la base ja s'està posant al dia); si és el
    botó, ho diu.
  - **L'editor obert mentre l'automàtic escriu.** `Save-BaseEditada` desa dins del
    mateix mutex i compara el segell (`actualitzat_el`) amb el que va carregar: si
    la base ha canviat, **hi fusiona** les correccions de l'editor informe a
    informe per la ruta (`_FusionaEdicionsBase`, pura) en lloc de reescriure-la
    —si no, desar tornaria enrere els informes nous—, i ho diu. Si s'està
    escrivint ara mateix, no desa i demana tornar-hi (els canvis no es perden).
- **Les correccions a mà («Editar base») PREVALEN** sobre el que surti
  d'«Actualitzar base» (octubre 2026). Es marca a l'**informe**
  (`editat_a_ma`, amb el valor automàtic guardat a `auto_conclusio_breu` /
  `auto_ignorat` per poder-ho desfer amb el botó *Desfer canvi a mà*). L'escaneig
  (`_AplicaEdicioPrevia`) manté la correcció i hi guarda el nou automàtic; un
  informe **no** corregit agafa el que diu ara el text (abans es congelava la
  conclusió breu de tots). A l'editor, l'*Estat activitat* d'una activitat amb
  alguna correcció surt **en vermell**. Ordre per **ID GIA numèric**
  (`_GiaNumeric`; abans 10, 1000, 103…) i el clic a la capçalera **mana**
  (`_OrdenaFilesBase`).
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
- **Conclusió, conclusió breu, tipus i estat → `InformesClassificacio.ps1`**
  (octubre 2026; vivia a `Informes.ps1`, que amb això passava de les 1.200
  línies). Tot pur. **Per què es va refer**: el 7/10/2026 es van llegir un per
  un els 802 informes de la carpeta real (427 activitats) i es va comparar amb
  el que en treia *Actualitzar base*: 110 informes en `Revisar`, 82 activitats
  sense estat útil i, pitjor, activitats amb l'estat **equivocat** (una
  llicència informada favorablement s'ignorava per defecte i l'activitat es
  quedava en `Requeriment` per un informe de dos anys abans). L'estat alimenta
  els **Recordatoris**: un `Requeriment` fals és un correu que no s'havia
  d'enviar.
  - **Un sol punt d'entrada, `_ClassificaInforme($lines, $fitxer, $expedient)`**
    → `{ Conclusio; Font; Breu; Tipus; Motius }`. El criden `Get-InformeData` i
    `ValidarClassificacio.ps1`: el que es valida és el que es fa servir.
  - **Frases d'inici** (`$Script:ConclusioStartPhrases`): les de sempre, més
    «Cal requerir l'esmena de les deficiències indicades» (la conclusió de
    `LLIC`, sense «Vist l'anterior»). **«Informo favorablement» i «s'informa
    amb caràcter favorable» competeixen amb les de primera fila** i mana la que
    surt primer al document (8/10/2026: un concert deia «S'informa amb caràcter
    favorable» i, més avall, «El titular és responsable d'executar»; guanyava la
    segona i sortia `Revisar`). Les **de segona** (`Segona = $true`: «és
    pertinent suspendre/precintar», «el Control Periòdic és FAVORABLE»,
    «estimar la sol·licitud d'ampliació», i els favorables/FI sense la frase de
    sempre: «s'equipara a resultat favorable», «COMPLEIX el que estableix el
    Decret 112/2010», «no s'aprecia cap irregularitat», «no se li poden requerir
    més mesures», «no provenen de l'empresa») **només es miren si no n'hi ha cap
    de les altres**: també poden sortir al cos (un seguiment que copia una conclusió
    anterior) i, com que la conclusió comença a la PRIMERA línia que en conté
    una, s'endurien mig informe. En segona passada no poden canviar res del que
    ja es trobava.
  - **La línia d'inici no pot ser el final.** La transmissió diu «…presentada a
    l'Ajuntament de Cornellà de Llobregat, s'informa FAVORABLEMENT…» i el
    «Cornellà de Llobregat,» (una de les frases de signatura) la tallava abans
    de començar: **totes** les transmissions sortien «sense conclusio». Es va
    veure passant `_ClassificaInforme` pels fitxers d'or (els textos de les
    plantilles), que és una manera barata de mesurar el classificador sense la
    carpeta real.
  - **`_ConclusioBreu`: l'ORDRE mana**, i el comentari diu quin cas real va
    posar cada bloc davant del següent. Davant de tot, **«es deixa sense efecte
    (la comunicació)»** → `Sense efecte` i el **precinte o la suspensió
    efectius** → `Precinte / Cessament`, encara que el text continuï amb les
    deficiències i un «Vist l'anterior, cal requerir…» (8/10/2026: sis
    activitats suspeses i una comunicació anul·lada sortien en `Requeriment`).
    `_PrecinteEfectiu` mira **frase a frase**: «és pertinent
    suspendre/precintar» només és l'advertiment si dins de la mateixa frase va
    precedit de «En cas contrari», «Si es disposen de més elements…» o «Si es
    detecta», o seguit de «en el cas de no presentar» — **excepte** «Si es
    detecta un ús de la cuina **estant precintada**…»: allà el precinte ja hi
    és, és vigent — **llevat que el text tanqui l'expedient** (`_TextTancament`:
    «es pot donar per tancada/finalitzat», «es pot aixecar/desprecintar», sense
    «no» al davant): el tancament de la denúncia copia l'advertiment i després
    diu «No s'ha detectat ús de la cuina… es pot donar per tancada la denúncia»; una afirmació directa
    («…fins a esmenar les deficiències», «…fins que hagi obtingut la
    llicència», «és pertinent precintar la cuina fins a…») és un precinte. Un
    precinte que només és advertiment dóna `Requeriment`, però **després** dels
    FI («es pot aixecar el precinte. Si es detecta…, és pertinent precintar» és
    un FI). Després, el que deixa l'expedient
    **pendent** encara que la frase digui «es pot donar per tancada la
    denúncia», «desprecintar» o «favorablement» («…però NO donar per finalitzat»
    i el «ni donar per finalitzat» amb el NO oblidat; «D'altra banda, es
    requereix…»; «es valora favorablement la solució… s'hauran de…»; el control
    periòdic «FAVORABLE… incorrecte havent de ser DESFAVORABLE»; «cal requerir
    l'esmena», que abans queia a `Revisar`). Després, **inici del procediment
    d'esmena**: la retirada que l'acompanya és l'advertiment («és pertinent que
    es retiri»), com el «determini el cessament» de sempre (un precinte efectiu
    ja ha sortit abans). Va davant del «es pot donar per tancada
    la denúncia» perquè «tanca la denúncia i inicia el procediment d'esmena» és
    un requeriment. Després els FI, i finalment precinte/suspensió («és
    pertinent suspendre» sense la frase literal del risc), ampliació («estimar»,
    no «desestimar»), favorable i les clàusules d'un requeriment nou.
    «Desfavorable» (i «no es pot informar favorablement») continua a `Revisar`
    a posta. Dues de noves (8/10/2026): «De totes maneres… s'ha de presentar»
    després d'un favorable és `Requeriment` (pas 1), i **una obligació**
    (`$Script:RxObligacio`: «S'ha de presentar/justificar/realitzar/aportar/
    retirar…», «S'haurà de…», «Cal justificar/presentar») sense cap altra
    decisió és `Requeriment` — però **darrere del favorable**, a posta: el
    favorable d'una activitat extraordinària llista «S'haurà de presentar…» i
    no és cap requeriment (hi ha una prova que passa els fitxers d'or pel
    classificador i ho vigila).
  - **Sense frase de conclusió** (97 dels 802: no són rars, són els formats
    d'abans) → `_EstatSenseConclusio`. Seguiment punt per punt (respostes a
    l'inici de línia, després de la data «dd/MM/aaaa: » que hi posa l'eina
    Seguiment): alguna negativa → `Requeriment`; **algun requeriment (línia amb
    una obligació) sense cap resposta abans del següent** → `Requeriment` (un
    concert amb «OK» sota uns punts i res sota el Pla d'Autoprotecció sortia FI;
    per això **una resposta positiva que no s'hi reconeix fa un Requeriment
    fals**: si en surt una de nova, va a `$Script:RespostesPositives` —«Es
    presenta», «S'aplica», «Es tramita», «S'ha realitzat», «S'han retirat», «No
    es requereix», «No cal»… les negatives es miren primer);
    totes contestades → `FI Requeriment` **amb motiu** «estat deduit, sense
    conclusio». Requeriment antic («S'han observat les següents deficiències que
    cal esmenar…»), **un cos que llista obligacions** o «…per tant no es pot
    informar» → `Requeriment`; van **després** del seguiment perquè el seguiment
    d'un requeriment els copia sencers. Denúncia d'accessibilitat sense res a
    requerir → `FI Requeriment`. La resta, `Revisar` + «sense conclusio».
  - **Plantilla sense omplir** (`_EsPlantillaSenseOmplir`: el SÍ i el NO de
    «es pot donar per tancada la denúncia» alhora, o «Copiar requeriment» al
    cos) → `Revisar` + «plantilla sense omplir». Abans era `Requeriment` (el
    «no es pot donar» guanyava).
  - **Cap informe s'ignora per defecte** (`_ConclusioIgnorarPerDefecte` es va
    esborrar). Abans s'ignorava tot el que tenia Font `mns` o `act_extr`
    perquè el **text** d'aquelles conclusions és gairebé igual entre informes;
    però per a l'**estat** barrejava tres coses diferents, que ara distingeix el
    **tipus** (`_TipusInforme`, desat a l'informe com a `tipus` perquè l'editor
    pugui recalcular sense obrir el `.docx`): `llicfav` (favorables de
    llicència, decideixen sempre), `mns` (MNS, canvi de nom, canvi de
    titularitat/transmissió) i `actextr` (nom del fitxer amb «ActExtr»/«Act
    Extr», o la conclusió del Decret 112/2010; i, **només si ni el text ni el
    nom diuen res**, l'expedient de la **sèrie** 2569/2565 —l'expedient és
    any/número/sèrie i només compta la sèrie: una MNS de la sèrie 2562 amb el
    número 2565 sortia com a activitat extraordinària—. Que el text mani sobre
    la sèrie ve d'una MNS amb la capçalera mal escrita, «2026/1/2565», a la
    carpeta de la 2562).
  - **`_InformeQueDeterminaEstat($act)` — rep l'ACTIVITAT, no la llista**, perquè
    la regla d'`actextr` depèn de si té GIA. Recorre els informes per ordre i es
    queda amb l'últim que decideix: ignorat → mai; `Altres` → només si no hi ha
    res més; `actextr` sota un GIA → no (abans l'estat de l'estadi el decidia el
    requeriment d'un concert); `mns` favorable → **neutre** (no tapa
    `$Script:EstatsPendents`, però si no hi ha res pendent dóna `Favorable`; abans
    una activitat amb només MNS quedava amb l'estat buit). Les dues darreres no
    s'apliquen a un informe `editat_a_ma`: el que l'usuari hi ha posat mana. Si
    li arriba una llista, **peta**; i un guard d'AST (`06-guards.ps1`, validat
    tornant a posar `$act.informes` a `ComprovarExcel.ps1`) mira totes les
    crides. `_EstatActualActivitat`, Recordatoris i Comprovar Excel en són
    clients.
  - **Mateixa data**: `_OrdenaInformesActivitat` desempata pel nom del fitxer (i
    la ruta). El `Sort-Object` del PowerShell 5.1 no és estable i l'estat podia
    canviar d'una passada a l'altra.
  - **`$Script:ClassificadorVersio`**: es desa a la base (`versio_classificador`)
    i, si no coincideix, l'escaneig torna a llegir **tots** els informes. Sense
    això una millora del classificador només arribava als informes que algú
    tornés a desar. **Puja-la cada cop que canviï el que en surt.**
- **Les opcions** de conclusió breu són `$Script:ConclusioBreuOpcions`
  (Requeriment, FI Requeriment —inclou «denúncia tancada»—, Precinte /
  Cessament, FI Precinte / Cessament, Favorable, **Favorable pre-llicència
  (Requeriment)**, **Favorable post-llicència (Favorable)**, Ampliació termini,
  Sense efecte, Altres, Revisar).
- **Els dos favorables de llicència** (octubre 2026, l'usuari; abans era la
  idea pendent «Favorable pendent doc.»): un informe `llicfav` amb «a l'espera
  de rebre la citada documentació» és **pre** (queda obert), la resta de
  `llicfav` **post**. El parèntesi diu com es **tracta**: `_EstatEquivalent`
  dóna `Requeriment` per al pre i `Favorable` per al post, i hi passen **tots**
  els que comparen un estat amb un altre de concret: les campanyes de
  Recordatoris (el pre entra a Requeriments), el color del Plànol (groc / verd)
  i els llistats d'*Editar base*. El pre és a `$Script:EstatsPendents` (una MNS
  favorable no el tapa). La referència de la validació deia `Favorable`:
  **`suport\ActualitzarReferenciaLlicencies.ps1`** en fa una còpia nova
  (`classificacio-informes_<data>_llic.json`, al costat de la de partida, que no
  es toca) amb aquests informes posats on toca, llegint-ne el text. `Altres` només es tria a mà. A **Editar base
  d'informes** la columna «Conclusio breu» és un desplegable editable i «Estat
  activitat» és només lectura: en canviar «Ignorar» o «Conclusio breu» d'un
  informe es recalcula (`_EstatActualActivitat`) i es propaga a totes les files
  de l'activitat.
- **Les correccions a mà no depenen de la unitat** (`_ClauInforme`): l'escaneig,
  `_FlattenInformesDb` i `_FusionaEdicionsBase` casen els informes per la ruta
  **relativa a `carpeta_arrel`**. Abans era la ruta absoluta i, amb la mateixa
  base i la carpeta a `F:` en lloc d'`I:`, no casava res: es reprocessava tot i
  es perdien totes les correccions sense avís. Si l'arrel és una altra de debò i
  no hi casa ni la meitat (`_AvisCanviArrel`), el botó **pregunta** abans
  d'escriure («la base és de X, ara s'escanejaria Y: s'hi perdrien N
  correccions»); l'automàtic, que no pot preguntar, no escriu i ho apunta al
  registre.
- **Correccions que ja no ho són**: si en reprocessar un informe corregit a mà el
  valor automàtic nou ja és el corregit (`conclusio_breu` i `ignorat`),
  `_AplicaEdicioPrevia` **treu la marca** i els `auto_*`. La base d'octubre de
  2026 en portava 179 que eren defectes del classificador, no gustos de
  l'usuari, i pintaven mitja base en vermell.
- **Agrupament**: un informe sense ID GIA en una carpeta on tots els altres que
  en tenen són del mateix GIA va amb aquell GIA (`_GiaDelsGermans`; cas real:
  un informe de llicència quedava com una activitat a part de la del seu
  expedient, perquè `_GiaFromFolderName` només mira el nom de la carpeta). Si
  l'ID GIA del document no és el de la carpeta, es queda amb el del document
  però surt a `a_revisar` («GIA del document diferent del de la carpeta»):
  abans se n'anava en silenci a l'activitat d'un altre titular.
- **Només el PRIMER NIVELL** (octubre 2026, petició de l'usuari):
  `Get-FitxersPrimerNivell` torna els fitxers que són **directament** dins de
  cada carpeta d'activitat de la carpeta d'informes (`<arrel>\<carpeta>\fitxer`).
  Ni els de l'arrel ni els de les subcarpetes: a les subcarpetes hi ha
  l'expedient del GIA descarregat (~25.000 fitxers) i informes vells que no
  han de decidir l'estat. Abans era un `Get-ChildItem -Recurse`. El fan servir
  la base d'informes i el repàs de contactes. Una entrada de la classificació
  de referència o de `gia-assignats` d'una subcarpeta surt ara com a «no
  trobada».
- **L'ID GIA assignat a mà** (octubre 2026): els informes antics no porten
  l'ID GIA a la capçalera, o el porten malament. L'assignació, llegida un per
  un, és a `local/base-dades-activitats/gia-assignats_*.json` (el **més
  recent** pel nom; porta dades personals i, com la classificació, no es puja
  mai): `{ "informes": [ { ruta_relativa, id_gia, confianca, motiu, forcar } ],
  "sense_gia": [ { ruta_relativa, motiu } ] }`, amb `ruta_relativa` = la clau
  de `_ClauInforme`. A `Get-InformeData` l'ordre és **`forcar=true` > document
  > assignat > carpeta «GIA n» > Excel per expedient** (i després els germans);
  `forcar` és per a una capçalera amb l'ID **equivocat** (els dos d'un titular
  deien 239 i eren del 293). Un GIA assignat surt amb `GiaFont = 'assignat'`
  (desat a l'informe com a `gia_font`), sense «sense ID GIA» i, si és forçat,
  sense «GIA del document diferent del de la carpeta».
  - `Read-GiaAssignats` **no peta mai**: un fitxer trencat (i `Read-JsonFile`
    torna `$null` en silenci, no llança) dóna les llistes buides i un `Error`
    que surt al resum del botó, al registre de l'automàtic i a la validació.
  - **La signatura del fitxer** (nom|data|mida) es desa a la base
    (`gia_assignats`): si canvia, l'escaneig torna a llegir **tots** els
    informes, com amb la versió del classificador. Sense això una assignació
    nova només arribaria als informes que algú tornés a desar.
  - Les entrades que ja no troben l'informe (mogut o reanomenat) tornen a
    `GiaAssignatsNoTrobats` i surten com a **AVÍS** a `ValidarClassificacio`.
    `sense_gia` de moment només es fa servir per a aquest avís.
- **Validar contra la carpeta REAL: `suport/ValidarClassificacio.ps1`** (no és
  de la suite). La classificació feta a mà
  (`local/base-dades-activitats/classificacio-informes_2026-10-08.json`, que
  substitueix la del 10-07) porta
  dades personals i el repositori és públic: a la suite hi ha **textos
  inventats**, i aquest script fa l'escaneig de debò en una carpeta temporal
  (sense la base ni les correccions de l'usuari) i escriu la **llista** de
  discrepàncies d'informe i d'estat d'activitat a
  `local/base-dades-activitats/validacio-classificacio_<data>.txt`. Deixa fora
  les entrades amb `nota` «DUBTE…» i llista a part les de judici de l'usuari
  («JUDICI: …», decisions de lectura). Als informes `mns` i `actextr` **no
  compara l'ignorat** (l'estat el decideix el tipus i el programa no en desa
  cap): només la conclusió breu, i tampoc no l'aplica a l'estat. Una
  diferència d'estat d'activitat que **només ve d'un informe JUDICI** (amb el
  del JUDICI com el diu el programa l'estat surt igual) va a un bloc a part i
  no compta.
  L'agrupament per activitat també fa servir el `gia-assignats_*.json` més
  recent: el busca **abans** de passar l'escaneig a la carpeta temporal (on no
  n'hi ha cap) i el passa a `Invoke-InformesDbEscaneig` (`-GiaAssignats` per
  triar-ne un altre). La classificació de referència no canvia: és per informe.
  `powershell -NoProfile -ExecutionPolicy Bypass -File suport\ValidarClassificacio.ps1`
  **Els dos favorables de llicència contra la referència** (`_BreuEquivalentRef`,
  octubre 2026): la classificació del 10-08 és d'abans del «Favorable
  pre-llicència (Requeriment)» i del «post-llicència (Favorable)» i hi diu
  «Favorable». Ara són **equivalents**: no compten com a discrepància, ni a
  l'informe ni a l'estat de l'activitat (es posa el d'ara en lloc del de la
  referència), i es llisten a part. El classificador 2026-10-09.1 només va
  partir el «Favorable» en pre i post: no pot fer d'un favorable un
  «Requeriment». Un «Llicència favorable» que surt «Requeriment» o bé és un pre
  (el seu estat es tracta com un requeriment, `_EstatEquivalent`) o bé era una
  correcció a mà que es va perdre amb la base refeta de zero.
- **Estats nous pendents de decidir amb l'usuari** (no fets): «Favorable pendent
  doc.» (el `favorable-pre` de llicència, que avui cau a `Favorable` i surt en
  verd al plànol amb l'expedient obert) i «Desfavorable». Si es fan, toquen
  `$Script:ConclusioBreuOpcions`, el desplegable de l'editor, els colors de
  `rutes/PlanolDades.ps1`, les campanyes de `Recordatoris.ps1` i el
  `LLEGEIX-ME.md`.
- **Menú Pas 1 — una graella de 4 × 4 rajoles a la dreta dels informes,
  agrupades per MOMENT DE LA FEINA** (octubre 2026, acordat amb l'usuari;
  `Select-Mode`, `Menu.ps1`, helper `$addTileRow`; dispatch al `switch` de
  `Main`):
  - **CARRER**: 📍 *Generar ruta* (`ruta`), 🗺 *Coordenades* (`coordenades`),
    🔒 *Activitats precintades* (`url`, acció `precintades` només per al
    segell), 📥 *Revisar mòbil* (`revisarmobil`).
  - **TITULARS**: 📧 *Enviar correu* (`enviarcorreu`), ✉ *Textos del correu*
    (`emailtextos`), 🔔 *Recordatoris* (`recordatoris`), 📅 *Controls
    periòdics* (`controlsperiodics`).
  - **BASE D'INFORMES**: 🗃 *Actualitzar base* (`informesdb`), 📋 *Editar base*
    (`informesdbedit`), 📁 *Copiar informes* (`copiarinformes`), 📄 *Word a PDF*
    (`convertirpdf`).
  - **GIA | NORMATIVA** (comparteixen fila): 📊 *Seguiment* (`seguimentgia`),
    ✅ *Excel Comprovar* (`comprovarexcel`; abans *Comprovar Excel*), ⚠ *Excel
    Errors* (`excelerrors`) | 📚 *Normativa* (`normativa`), 🔍
    *Revisar requeriments* (`revisio`).
  - **Per què així**: primer les eines anaven a sota dels informes (~980 px
    d'alt, calia fer scroll); després en una columna a la dreta amb files de
    5, 4, 2+2 i 3, i l'usuari va dir que «no hi ha harmonia amb les
    alineacions». Ara: **`_MenuDisposaGrups`** (pura) posa els grups en una
    graella de quatre columnes —entre grups hi ha el mateix espai que entre
    rajoles, i per això NORMATIVA cau just a la tercera columna— i
    **`_MenuFilesY`** (pura) estira les files perquè la primera comenci a
    l'altura del primer botó d'informe i la darrera acabi on acaba l'últim.
    Els títols (INFORMES a l'esquerra, els grups a la dreta) van tots a la
    mateixa línia i amb el mateix estil; Capçalera · Conclusions, alineats a
    la dreta de la columna; les icones de la banda, al mateix marge dret que
    les rajoles.
  - **Una eina nova** va a l'array del seu grup; **cap grup passa de quatre**
    (si no hi cap, un grup nou). Guard a `06-guards.ps1`: calcula l'alçada a
    partir del codi i exigeix que càpiga en una pantalla de 768 px, i cap grup
    de més de quatre (validat injectant entrades d'informe i una rajola de més).
- **El botó «↻ Actualitzar» és a la banda del menú** (a més de Configuració):
  l'usuari el fa servir molt. Tots dos criden `Invoke-ActualitzarPrograma`
  (`Configuracio.ps1`): el missatge de confirmació, el `Start-Process` del
  `.bat` i el `[Environment]::Exit` són en un sol lloc (guard).
- **Els emojis, EN COLOR** (`suport/emojis/*.png` + `_DibuixaEmoji`,
  `UiFinestra.ps1`). El GDI de WinForms no sap pintar lletres de colors: amb
  `Segoe UI Emoji` sortien d'un sol color, i l'usuari els va veure en color en
  un esbós fet amb el navegador. Cada emoji de la interfície és una imatge de
  64 × 64 generada amb Noto Color Emoji (vegeu el `LLEGEIX-ME.txt` d'allà); el
  nom són els punts de codi sense el FE0F (`_EmojiFitxer`, pura). Sense
  imatge, es dibuixa amb la lletra com abans. `_PosaIcona` també la fa servir
  (l'enllaç 🔗 de l'editor), però **els símbols** (✓, ⚠, ⓘ, fletxes) no en
  tenen a posta: han d'anar del color del text. Les imatges es llegeixen **a
  memòria**, no amb `Image.FromFile`, que deixaria el fitxer agafat i
  `Actualitzar.bat` no el podria canviar a la unitat de xarxa. Guards: cada
  emoji del menú té imatge (validat esborrant-ne una) i el menú no pinta cap
  emoji amb `TextRenderer` (validat tornant-hi).
  - **Per afegir un emoji**: renderitza'l a 64 × 64 amb fons transparent i
    desa'l amb el nom de `_EmojiFitxer` (com es van fer: Chromium + la font
    Noto Color Emoji, `font-size:52px` en una caixa de 64).
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
- **L'interruptor A/M de «Copiar informes»** (menú, setembre 2026; des de
  l'octubre també a «Actualitzar base»). Les rajoles amb commutador porten
  `Interruptor = $true` a la seva entrada i la seva eina s'apunta a
  **`$Script:ModesAuto`** (`ModeAutomatic.ps1`: `Actiu`, `DesaActiu`,
  `UltimMode`, `SiToca`, `Requisit`, `TipA`/`TipM`); el menú només recorre el
  registre, i un guard comprova que les dues llistes coincideixen. L'estat de
  cada commutador va al `Tag` del seu `Panel`, i els mateixos scriptblocks
  serveixen per a tots. Va **a
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
- **PROGRAMACIONS CONFIGURABLES** (octubre 2026: «com són ja uns quants, haurien
  de ser configurables des de la configuració»). Cada automatisme s'apunta amb
  `Register-ProgramacioAuto <clau> <títol> [dia|setmana] [dia 1-7] [hora]`
  (`ModeAutomatic.ps1`); per defecte tots són a les 13:00 (`$Script:AutoHora`) i
  cada dia, tret del *Plànol activitats* (cada **dilluns**). ⚙ *Configuració* té
  un grup **Automatismes** que es munta sol amb el que hi ha registrat:
  interruptor (si és a `$Script:ModesAuto`), «cada dia / cada setmana», el dia i
  l'hora. Es desa a `Automatismes` dels settings **només si difereix del valor
  per defecte** (`ConvertTo-AutomatismesSettings`), i `Get-ProgramacioAuto`
  rellegeix els settings a cada consulta: **no cal reiniciar**. Un valor
  invàlid (hora mal escrita, dia fora de 1-7) torna al de per defecte. El
  venciment setmanal fa servir el dia ISO (diumenge = 7). Els Recordatoris, que
  són una tasca del Windows, es reescriuen en desar (`Update-RecordatorisTascaSiCal
  -Forca`, `ScheduleByWeek` si és setmanal); l'interruptor dels recordatoris
  continua sent el botó *Automàtic...* de la seva eina.
- **L'HORA DE TOTS ELS MODES AUTOMÀTICS: les 13:00** per defecte (octubre 2026; abans la
  còpia era a les 14:30, la base a les 14:00 i els recordatoris a les 09:00). Viu
  en un sol lloc, `$Script:AutoHora`/`$Script:AutoMinut` (`ModeAutomatic.ps1`),
  i la regla és la mateixa per a tots: **si l'última vegada que tocava no es va
  fer, es fa tan aviat com es pot**. Copiar informes i Actualitzar base ho miren
  en obrir el programa (`_AutoToca`). Els Recordatoris, que són una tasca del
  Windows, es creen amb **XML** (`_RecTascaXml`) per poder-hi posar
  `<StartWhenAvailable>`: amb `schtasks /SC DAILY /ST` no es pot, i amb el PC
  apagat a l'hora aquell dia no s'enviava res. Una tasca d'abans es reescriu sola
  en obrir el programa (`Update-RecordatorisTascaSiCal`, una vegada per
  execució; si no existeix, no es crea). I com que la base s'actualitza a la
  mateixa hora, `RecordatorisAuto.ps1` espera el seu mutex (`Wait-MutexLliure`,
  fins a 30 min) per fer servir la base d'avui i no la d'ahir.
- **Copiar informes, mode AUTOMÀTIC** (interruptor **A/M** del menú, setembre
  2026). La mateixa còpia es fa de dues maneres i per això la feina viu en
  **quatre funcions sense cap finestra** (`_CopiaInformesPrepara`,
  `_CopiaInformesCerca`, `_CopiaInformesTria`, `_CopiaInformesCopia`); el que
  difereix —avisar, preguntar, pintar la barra— es queda a la crida, en
  **scriptblocks**. Hi ha **un sol `Copy-Item` a tot `Informes.ps1`**, i un guard
  ho vigila: si el manual i l'automàtic es munten cada un el seu bucle, un dia
  copiaran coses diferents.
  - **Quan toca: UNA sola pregunta**, `_CopiaAutoToca` (pura, amb proves) —
    «des de l'últim **venciment** (les 13:00 que tocaven), s'ha fet cap passada
    automàtica?». Serveix per als dos casos que va demanar l'usuari (el rellotge
    de les 13:00 amb el programa obert, i la passada perduda que es recupera en
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

## Contactes de les activitats (octubre 2026) — `Contactes*.ps1`

**Què demanava l'usuari:** que «Actualitzar base» (el botó i l'automàtic de
les 13:00), a més de l'estat de cada activitat, **repassi les dades de
contacte** de cada ID GIA: que **completi** les que falten a l'Excel
d'activitats (titular, representant legal, persones autoritzades) i que
**avisi** de les que hi són malament. El cas més freqüent: **el tècnic posat
com a representant legal**, o el seu correu o el seu mòbil com si fossin del
titular. **L'Excel és una exportació del GIA i el programa no el toca mai**:
el que surt és una base pròpia, una finestra per revisar-la i una llista de
correccions per entrar-les al GIA a mà.

**Els fitxers**, un per cosa:

| Fitxer | Què hi ha |
|---|---|
| `PdfText.ps1` | el text d'un PDF generat, sense biblioteques (C# compilat amb el de `PdfUnio.ps1`) |
| `ContactesExtraccio.ps1` | què diu cada document (pur): e-TRAM, instància, autorització, XML antics |
| `ContactesRegles.ps1` | de qui és cada dada i els avisos contra l'Excel (pur) |
| `ContactesDb.ps1` | `contactes-db.json`, el repàs, les correccions a mà, l'ús als correus, l'exportació, la validació |
| `ContactesPantalla.ps1` | la finestra «Contactes» (des d'«Editar base») |
| `ValidarContactes.ps1` · `DiagnosticPdf.ps1` | per al PC de l'usuari (`local\ValidarContactes.bat`, `local\DiagnosticPdf.bat`) |

### Quins documents: tres tipus, triats pel NOM

A la carpeta d'informes hi ha ~25.000 fitxers (cada carpeta porta l'expedient
del GIA descarregat). L'usuari ho va mesurar: gairebé totes les dades surten de
tres tipus de document, petits. `_CtTipusDocument` els tria **pel nom** abans
d'obrir res:

| Tipus | Nom | Què se'n treu |
|---|---|---|
| `etram` | `*etram-tramit*.xml` | `<solicitant>`, `<representant>`, `<canalSms>`, `<canalCorreuElectronic>`, `<dataCreacio>`, `<nomTramit>` i les `<dada>` de l'establiment |
| `instancia` | `*instància genèrica*`, `*Instancia_generica*`, `*Esmena sol·licitud*`, `*SEU  OAC*` (PDF) | les seccions de l'interessat i del representant, la data i el títol, i les signatures `(R:CIF)` |
| `autoritzacio` | `*autoritz*`, `*autoriz*`, `*representaci*` (no `Índex electrònic…`) | qui signa en nom de l'empresa (el **representant legal**) i a qui autoritza |
| `xmlantic` | `*XML_TRAMIT*.xml`, `*sicres3*.xml` | `Nombre_Interesado`… i `_Representante` |

**Tot l'arbre** de cada carpeta d'activitat, subcarpetes incloses
(`Get-FitxersContactes`; abans només el primer nivell, i l'usuari va trobar
~195 documents bons sense llegir: `Firmadoc\`, `OME\`, `Final\` i sobretot la
carpeta d'expedient del titular **dins** de la d'una queixa). Els fitxers de
l'arrel no. Els certificats, les memòries i els projectes no es llegeixen:
pesen molt i gairebé només donen el tècnic.

**L'ID GIA de cada document** (`_CtGiaDocument`) és el de **la carpeta més
propera que digui de quina activitat és**, de dins enfora: el «GIA n» del nom,
els informes que hi ha (si tots són del mateix GIA, com `_GiaDelsGermans` de la
base d'informes) o el seu expedient a l'Excel (`AAAA-N-NNNN`, provat en els dos
ordres: el document l'escriu `2025/3165/16` i la carpeta `2025-16-3165`).
**El GIA que diu el document, només si cap carpeta no en diu cap**: l'escriu el
ciutadà i sovint malament (carpeta `GIA 595`, esmena «GIA 5»: l'activitat 5
rebia una esmena d'una altra i sortia un fals canvi de titular). Després,
l'expedient del document → Excel. Una carpeta de diverses activitats amb un
document que no diu de quina: «sense GIA» (surt a la validació).

**Incremental**, com els informes: per cada document es desa la ruta relativa,
la data de modificació i el que se n'ha tret, i només es tornen a llegir els
nous o modificats (o tots si canvia `$Script:ContactesVersio`). **Compte amb el
pwsh 7**: torna les dates ISO del JSON com a `[datetime]`; la comparació va amb
`Read-JsonIso` (ho va enxampar la prova de l'incremental).

### Llegir els PDF sense biblioteques (`PdfText.ps1`)

Per què no PDFsharp ho explica `PdfUnio.ps1`. Les instàncies de la seu i els
formularis de l'e-TRAM són **PDF generats, amb text**. El lector fa servir el
mateix `PdfDoc` de la unió (xref, fluxos d'objectes, FlateDecode amb predictor;
s'hi han afegit ASCII85 i ASCIIHex) i hi posa:

- els operadors de text `Tj`, `TJ`, `'` i `"`, i la posició (`Tm`, `Td`, `TD`,
  `T*`, `cm`, `q`/`Q`) per tornar a fer les **línies**: cada tros es desa amb la
  seva posició i, en acabar la pàgina, s'ordena de dalt a baix i d'esquerra a
  dreta. **Els formularis posen l'etiqueta i el valor a la mateixa alçada però en
  qualsevol ordre al contingut**: llegint en ordre del flux sortien barrejats;
- la `ToUnicode` (`bfchar`/`bfrange`) de cada font, i sense, WinAnsi +
  `/Differences`; l'amplada (`/Widths`, `/W`) per saber si dos trossos de la
  mateixa línia porten un espai al mig;
- les formes (`Do` d'un XObject `/Form`) i salta les imatges en línia (`BI … EI`).

El C# de `PdfText` va **dins del namespace amb els `using` a dins** i es compila
**enganxat** al de `PdfUnio` (`_PdfUnioCarrega`): necessita `PdfDoc`, i dos
`Add-Type` separats no es poden referenciar en memòria al PowerShell 5.1.

Si no en surt text (una font `Identity-H` sense `ToUnicode`), **el Word**
(`_PdfTextWord`, obert un sol cop per passada i només si cal). Si tampoc,
**escanejat**: a la llista «no llegibles». L'OCR del Windows queda per més
endavant (l'usuari: «ara no»).

**EL WORD NO POT QUEDAR-SE ESPERANT UNA PREGUNTA** (octubre 2026, va passar el
primer dia): en obrir un PDF, el Word ensenya «Word convertirá el PDF a un
documento de Word editable…» amb Aceptar/Cancelar, i **ni `ConfirmConversions`
ni `DisplayAlerts` el treuen**. Amb el Word amagat, el procés s'hi queda
esperant (i, mentrestant, l'explorador de Windows deixava de respondre). Per
això:

- `_PdfTextWord` no obre cap PDF sense haver posat abans
  `DisableConvertPdfWarning = 1` a `HKCU\Software\Microsoft\Office\<versió>\Word\Options`
  (`_WordSenseAvisPdf`: és la casella «No volver a mostrar este mensaje»), i
  només si després s'hi llegeix;
- els PDF **xifrats no van mai al Word** (demanaria la contrasenya);
- **el Word no obre mai el PDF de la xarxa**, sinó una còpia al `%TEMP%`
  (`Open-WordCopiaLocal`, `Motor.ps1`, el mateix que ja feien els `.doc` de la
  base d'informes). Llegint de la unitat de xarxa, el Word hi escriu els seus
  temporals `~WRS{…}.tmp` i a la primera interrupció ensenya «Word ha realizado
  una lectura doble de su archivo», «Word perdió datos debido a una conexión de
  red incorrecta» o «Word no puede abrir el archivo ~WRS…», que el
  `DisplayAlerts` no treu (l'usuari, octubre 2026, en fer «Actualitzar base»:
  amb les subcarpetes, molts més escanejats van al Word);
- la passada **automàtica** (sense ningú davant) no fa servir el Word
  (`Invoke-ContactesEscaneig -SenseWord`): els que l'haurien necessitat queden
  `pendent_word` i els torna a llegir el pròxim «Actualitzar base» a mà, encara
  que no hagin canviat. Guards a `11-contactes.ps1`, validats injectant-hi el
  defecte.

**Les etiquetes, en català i en castellà** (versió 2026-10-15.1). La instància
de la seu porta cada etiqueta en les dues llengües i el text del PDF les talla
en línies diferents: totes les etiquetes primer i els valors a sota («Nom i
cognoms o raó social Nombre y apellidos o razón social» / «DNI / NIF / NIE /
Pas.» / «EXEMPLE SL» / «B00000000»). Sortien 132 titulars amb el nom «Nombre y
apellidos o razón social DNI / NIF / NIE / Pas». Ara:

- una línia que, sense la puntuació, **només té paraules d'etiqueta** és una
  etiqueta (`_CtEsLiniaEtiqueta`; cal com a mínim una paraula «forta», perquè
  «DE LA FONT» és un nom);
- el nom és **la primera línia que no és etiqueta** i el NIF, la següent
  (`_CtPrimerValorNom`); a la instància telemàtica antiga, amb les etiquetes i
  els valors a la mateixa línia, el que va darrere de les etiquetes
  (`_CtSenseEtiquetes`);
- `_CtPersona` no accepta mai com a nom una etiqueta (`_CtNomEsEtiqueta`:
  «razón social», «representado», «DNI / NIF», «nom i cognoms»…), ni
  `_CtNetejaNom` un DNI dins d'un nom;
- l'expedient és només `AAAA/NNNN/N`: s'hi colaven les dates (262 documents);
- a l'e-TRAM antic, `<raoSocial>rao</raoSocial>` (i `-`, `.`, menys de tres
  lletres) no és cap nom: es fa servir nom + cognoms (`_CtRaoBuida`).

**L'autorització** (`Read-AutoritzacioText`): a més d'«autoritzo», els verbs
`designa` (amb «com a tècnica responsable … a Z»), `nomena`, `atorga la seva
representació a`, `otorga su representación a`, `apodera` i `faculta`; sense «Jo,»
ni tractament, el signant és el nom de davant d'«amb DNI», i el nom s'atura a
`amb|con DNI|NIF|NIE|N.I.`, la coma, `actuant`, `actuando`, `en representació`,
`en nombre`, `com a`, `como`, `en qualitat`, `en calidad` (`$Script:CtReFiNom`).

**La signatura electrònica de representant, la font més segura**
(`_CtSignaturesRep`, a tots els PDF que es llegeixen): «Firmado digitalmente por
00000000T MARIA EXEMPLE (R:B00000000)», també partida en línies i amb el CIF
tallat (s'ajunten abans de buscar). Qui signa amb `(R:<CIF>)` és el representant
legal de l'entitat d'aquell CIF.

**No s'ha pogut veure cap instància de debò** (porten dades personals): el
lector i els lectors de cada document es van fer amb documents inventats.
`DiagnosticPdf.ps1` (`local\DiagnosticPdf.bat`: s'hi arrossega un PDF) escriu
el text extret i el que en treuen els lectors a
`local\base-dades-activitats\diagnostic-pdf_<nom>.txt`. **És el primer que s'ha
de mirar quan la validació doni discrepàncies.**

### Les regles (`ContactesRegles.ps1`)

1. **El tècnic no és mai el representant legal.** Tècnics, enginyers,
   arquitectes, gestories i tramitadors són **persones autoritzades**.
   `_CtMotiuTecnic` els reconeix: a `tecnics-coneguts_*.json` (el més recent),
   el correu o el telèfon surten a **activitats de titulars diferents**
   (comptat sobre tota la base, Excel inclòs, a `_CtContext`: el senyal més
   fort), el correu és professional (`engin`, `arquitec`, `gestor`…), una
   autorització els autoritza, o presenten amb un NIF d'entitat (gestoria).
   **`-Fort`**: per treure un correu o un telèfon AL TITULAR només valen la
   llista i els titulars diferents; un «estudi de dansa» o una «oficina» poden
   ser el titular, i el patró professional se'l menjaria.
   **També el NIF i el nom** de qui presenta pel titular (el representant del
   tràmit i l'autoritzat, `NifTitulars`/`NomTitulars` de `_CtContext`): el
   mateix DNI o el mateix nom per a ≥2 titulars diferents és tècnic, i el
   correu que hi ha posat, si no té res de tècnic, **és del client** (va al
   titular, «probable»). Patrons professionals: també `advocat`, `abogad`,
   `icab.cat`, `legal`, `lawyer`, `assessors`, `asesores`, `gestio`, `ebcn`.
   Sense `tecnics-coneguts_*.json`, el resum d'«Actualitzar base» ho diu
   («AVÍS: no hi ha llista de tècnics coneguts»).
2. **Representant legal**: persona física → en blanc (és ella mateixa);
   persona jurídica → **qui signa amb el certificat de representant del CIF del
   titular** (`(R:<CIF>)`: «segur», `via: signatura`, i mana sobre les regles:
   només el treu una correcció a mà «És el tècnic»), qui signa l'autorització en
   nom de l'empresa («segur»), el representant d'un tràmit que no sigui tècnic o
   la persona física que presenta en nom propi amb el correu de l'empresa
   («probable»). Mai el tècnic. Qui signa amb `(R:)` d'**una altra** entitat
   (una enginyeria, una gestoria) és persona autoritzada; les de les
   administracions (CIF `P`, `Q`, `S`) no són de ningú.
3. **El titular** = l'interessat, camp a camp i del més recent, **tret del que
   sigui del tècnic**: el correu o el mòbil d'una persona autoritzada no pot ser
   del titular. **Si el NIF de l'interessat no és el de l'Excel** (i el nom
   tampoc), el seu correu i el seu telèfon **sí** que són del titular, com a
   «probable» (abans es perdien); el nom i el NIF van a l'històric. Si el nom
   extret és buit o una etiqueta, la Raó social de l'Excel.
4. **L'establiment**: el telèfon i el correu de l'e-TRAM, si no són del tècnic.
5. **Fora**: les **queixes** (sèrie `2579`: els interessats són els veïns),
   tret que el NIF sigui el del titular. Ho diu l'expedient del document, el
   nom del fitxer o **la carpeta d'expedient més propera** (`AAAA-N-2579`, o les
   paraules QUEIXA, QUEXIA, DENUNCIA); la carpeta del titular **dins** de la
   d'una queixa no és queixa (`_CtEsQueixa`). Les administracions
   (`cornella.cat`, `gencat.cat`, `aoc.cat`, Mossos, Bombers…) i les entitats
   de control (`tuv.com`, `aucatel`, `eicon`, `addient`, `eca-`, `dekra`,
   `applus`…, sobre el domini) **no són mai ni del titular ni persones
   autoritzades**.
6. **Recintes** amb organitzadors d'actes (GIA 28 i 1324; es canvia a
   `local\base-dades-activitats\contactes-config.json`, `{ "recintes": [...] }`):
   no es proposa res del titular ni del representant.
7. **El més recent mana** (per la data del document); la resta, a `historic`.
8. Cada dada porta **font** (ruta relativa), **data** i **confiança**.

Els telèfons es reparteixen **pel prefix** (6/7 mòbil, 8/9 fix), no pel camp on
els va escriure el formulari: és el mateix criteri amb què es comparen amb
l'Excel.

### Els avisos contra l'Excel

L'Excel es llegeix **una sola vegada** (la de la base d'informes):
`Initialize-ActivitatsCache` hi afegeix les columnes de contacte (NIF, telèfons,
representant legal) **a `$cache.Contactes`, no a `ById`**, que és el que es puja
al Drive del mòbil (hi ha prova amb un doble de l'Excel, validada injectant-hi
el NIF).

| Tipus | Quan |
|---|---|
| `es_el_tecnic` | el representant legal, un correu o un telèfon de l'Excel són d'una persona autoritzada (diu de qui) |
| `falta` | l'Excel és buit i els documents donen la dada |
| `diferent` | els documents més recents diuen una altra cosa |
| `error` | `gmailo.com`, `hotmal`, `holmail`, un domini a 1–2 lletres del dels documents, 10 xifres, un fix al mòbil, un correu al telèfon, la lletra del DNI, el CIF de l'empresa al NIF del representant |
| `canvi_titular` | el document no-queixa més recent és d'un altre NIF, d'una entitat o persona que no és el tècnic (només avís: l'Excel pot ser més nou). No ho és el mateix nom de l'Excel amb un altre NIF (l'Excel el té malament), qui signa amb el certificat del titular, ni la persona física que presenta per l'empresa |

### On es guarda i les correccions a mà

`local\base-dades-activitats\contactes-db.json` (no es puja mai): per GIA,
`titular_tipus`, `titular`, `representant_legal`, `persones_autoritzades[]`,
`establiment`, `avisos[]`, `documents[]` — **el mateix format que la
referència**, perquè la validació hi pugui comparar camp a camp — i, al costat,
la fila de l'Excel (`excel`). A dalt, la cache dels documents.

Les **correccions a mà** (`correccions`, per GIA) **manen**, com `editat_a_ma` a
la base d'informes: «És el tècnic», «És el representant legal», «Descarta
l'avís» i «Edita». **No es desen a sobre del resultat**: entren a les regles
(`_CtManDeCorreccions` → `Get-ContactesActivitat`) i es tornen a aplicar a cada
passada, o sigui que el pròxim «Actualitzar base» no les desfà. Cada una guarda
el valor automàtic per a «Desfer canvi a mà», i la finestra torna a calcular
l'activitat sense tornar a llegir cap document (`Update-ContactesActivitats`).
Un sol escaneig a la vegada: el mutex de la base d'informes.

### La finestra «Contactes» i el resum

- Al final d'«Actualitzar base»: «Contactes: N activitats amb dades noves · M
  avisos (X del tècnic, Y errors…) · K documents no llegibles» (al missatge del
  botó i al registre de l'automàtic).
- **Editar base → Contactes...** (s'hi apunta pel registre
  `$Script:EditarBaseBotonsExtra`: l'editor no depèn dels seus clients). La
  llista per GIA amb filtre per tipus d'avís, el detall amb l'Excel i els
  documents de costat, «Obre el document», els quatre botons de correcció i
  **«Exportar correccions (Excel)»** (GIA, titular, camp, valor a l'Excel,
  problema, proposta, font), amb el mateix escriptor d'`.xlsx` sense Excel que
  l'índex de la Normativa.

### Fer-la servir als correus

On el programa agafa el correu del titular o del representant de l'Excel
(Enviar correu, Recordatoris, Controls periòdics), `Get-ContactesEmailsCompletats`:

- si l'Excel és buit, proposa el de la base de contactes i **diu que surt dels
  documents**;
- si el correu de l'Excel té un avís `es_el_tecnic`, **avisa abans d'enviar**
  (Enviar correu: abans del diàleg; Recordatoris a mà: pregunta per a cada un;
  en automàtic, al registre; Controls periòdics, que són esborranys, a la llista
  del final).

La destinació «Persones autoritzades / tècnic» de Configuració → Correus de cada
eina surt d'aquí (`Get-ContactesAutoritzatsEmails`).

### La validació (`ValidarContactes.ps1`)

Germana de `ValidarClassificacio.ps1`: fa el repàs de debò **sense correccions
a mà i sense desar** (`-SenseCorreccions -NoDesis`) i el compara amb
`contactes-referencia_*.json` (336 activitats revisades a mà): e-mail i
telèfons del titular (els telèfons com a conjunt), nom del representant legal,
els e-mails de les persones autoritzades i els avisos `es_el_tecnic`. Les dades
«probable» de la referència van a part. Escriu
`local\base-dades-activitats\validacio-contactes_<data>.txt`. **L'objectiu és 0
discrepàncies en les «segur».** Al principi del `.txt` (`_CtEstadistiques`): quants
documents s'han llegit per tipus i quants són de subcarpetes, quants noms són
etiquetes (ha de ser 0), quantes persones autoritzades són d'una entitat de
control (0) i quants representants legals surten de la signatura `(R:CIF)`. Cal haver fet «Actualitzar base» abans
(l'atribució al GIA és la de la base d'informes).

També: els representants legals per on surten (`via`: `signatura`, `interessat`,
`autoritzacio`, `tramit`), els documents atribuïts pel NIF, els noms que
«presenten» per a més de 15 titulars (si n'hi ha cap, o és un tècnic de debò o
una etiqueta que s'ha escapat) i, per cada document sense GIA, per què
(`_CtMotiuSenseGia`: els GIA dels informes de les carpetes de sobre).

### La 2a ronda (versió 2026-10-15.2)

La primera validació a la feina va donar 0 noms-etiqueta, 0 ECA autoritzades i
881 documents de subcarpetes llegits; el que quedava era el **representant
legal** (la referència en té 187 i el repàs en trobava 71) i uns quants falsos
«és el tècnic». El que es va canviar, i per què:

- **La signatura `(R: CIF)`** (`_CtSignaturesRep`): la forma més corrent porta un
  **espai** després dels dos punts i, al caixetí, el nom **sense el DNI**
  («MARIA EXEMPLE (R: B00000000) Firmado digitalmente por…»). Abans es demanava
  el DNI just davant del nom: se'n trobaven 81 de 183. Ara, per cada `(R:CIF)`,
  el tros de davant fins a «Firmado digitalmente por», un altre `)` o 150
  caràcters: el DNI si n'hi ha i les paraules en majúscules del final. La
  mateixa persona amb DNI i sense, una sola vegada.
- **L'interessat persona física d'una entitat** és candidat a representant legal
  («probable», `via: interessat`), tingui el correu que tingui (GIA 17, 293).
  Només no és canvi de titular si, a més, presenta en nom propi amb el correu
  de l'empresa (`_CtInteressatDeLEmpresa`).
- **L'autorització sense «Jo,»** («X amb DN(I): …, en nom i representació de Y…»):
  el signant comença **a la seva línia** (`_CtAjuntaLinies` recorda on hi havia
  salts) i, si al davant hi ha text en minúscules, són les paraules en
  majúscules. A «designa com a Tècnica Responsable … a Z», Z és el que va
  darrere de **l'última « a »** abans de la coma (`_CtZDespresDeCarrec`): abans
  havia de ser en 80 caràcters i l'autoritzat sortia «com a Tècnica…».
- **Noms brossa** (`_CtNomBrossa`): fora els que comencen per «N.I», els que
  porten «interviene», «calidad», «qualitat», «número» i, en una persona, xifres
  («BAR 2000 SL» sí que és un nom). Els tractaments del davant, amb la
  puntuació («/D»», «D./Dña.»).
- **Etiquetes**: més paraules (avinguda, escala, mitjà de notificació…), i una
  línia amb tres paraules fortes o més i el 80% d'etiqueta també ho és (les de
  l'adreça, plenes de parèntesis).
- **`_CtMateixNom` demana un cognom**: les paraules del nom curt han de ser al
  llarg i com a mínim una no ha de ser un nom de pila corrent
  (`$Script:CtNomsPila`). «Maria del Mar» ja no és «Maria del Mar Garcia».
- **El correu de la llista que porta una altra persona** (`_CtCorreuTecnicAltri`,
  GIA 87): la persona conserva el seu paper i el correu és de la gestoria
  (persona autoritzada amb el nom de la llista).
- **El domini del titular no és del tècnic** (`_CtDominisTitular`, GIA 92 i 154):
  una dada que el tècnic porta amb el domini de l'empresa (o que l'interessat
  dona en un document sense tècnic) no és seva: no queda a la persona
  autoritzada (ni els telèfons que porta amb aquell correu), va al titular, i
  no fa cap avís «és el tècnic». El domini del correu de l'Excel només compta
  si aquell correu no surt a d'altres titulars (és l'error més corrent de
  l'Excel).
- **Titulars diferents: 3**, o 2 amb un altre senyal (correu professional, o
  presenta com a representant als dos) (`_CtProuTitulars`). I dos titulars que
  comparteixen el representant legal de l'Excel (NIF o nom) o el domini del
  correu són **un sol grup** (`_CtGrupsTitulars`), només si el comparteixen com
  a molt tres titulars i no té res de tècnic: si no, el tècnic posat de
  representant legal a moltes activitats (l'error més corrent) faria de pont.
- **Correus de gestoria pel nom** (`_CtCorreuGestoriaPelNom`: «asesoria…@»
  no és del titular, tret que el titular ho sigui) i **el domini d'un tècnic de
  la llista**, igual o a 2 lletres (`_CtDominiTecnic`; mai gmail i companyia).
- **Documents sense GIA**: el GIA dels informes de la carpeta és el de **més de
  la meitat** (`_CtGiaMajoritari`; abans havien de ser tots del mateix), i si
  cap carpeta no diu res, el **NIF de l'interessat** si és el d'una sola
  activitat de l'Excel (`gia_via: nif`, dades «probable»).
- **L'OCR dels escanejats** (punt D, opcional): no s'ha fet. A Linux no es pot
  provar ni el `Windows.Media.Ocr` ni el `Windows.Data.Pdf` (WinRT), i un codi
  sense provar dins d'«Actualitzar base» no hi pot anar.

**Les bases no es tornen a fer de zero en silenci** (`Read-JsonFileStrict` i
`Backup-JsonFile`, `Json.ps1`). A la feina, `informes-db.json` es va tornar a
generar de zero (`generat_el` nou) i es van perdre les 170 correccions a mà: si
la base hi era però no s'havia pogut llegir (el clone viu en una unitat de
xarxa), `Read-JsonFile` tornava `$null` i l'escaneig ho prenia per «no hi ha
base». Ara, si el fitxer hi és i no es pot llegir (tres intents), «Actualitzar
base» s'atura amb un avís i no escriu res, i el repàs de contactes igual. I
abans de sobreescriure `informes-db.json` o `contactes-db.json` (l'escaneig,
l'editor de la base i les correccions de la finestra Contactes) se'n fa una
còpia `<nom>_AAAAMMDD-HHMM.bak` al costat; se'n guarden les 5 últimes.

Els dos `.bat` de `local\` els escriu la migració (`_BatsLocal`,
`Invoke-MigracioLocal`): `local\` no es puja, o sigui que no poden anar al
repositori.

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
- **Avisar titulars per correu (esborranys a Outlook)** (`suport/ControlsCpEmail.ps1`; passa
  per la sessió de `CorreuVia.ps1` en mode esborrany, que hi posa el remitent de Configuració;
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

### Controls periòdics AUTOMÀTIC (octubre 2026)

L'usuari: «aplica l'automatització de les eines a l'eina Controls periòdics».
`ControlsPeriodicsAutomatic.ps1` (interruptor A/M, registre, programació
`setmana` / dilluns) + `ControlsPeriodicsAuto.ps1` (la passada en segon pla).

- **Què és «automatitzar» aquí**: la feina que l'usuari feia a mà —obrir la
  llista, marcar els que ja han passat el *Proper CP previst* i preparar-ne
  els correus—. El text del correu ja diu «havia de passar … i no ens consta»:
  per això només els **vençuts** (`KProperCP` < avui), no els que vencen aviat.
  Generar els informes en lot **no** s'automatitza: són Word, i no se'n
  demanava.
- **NOMÉS ESBORRANYS**, també en automàtic: la regla de l'eina i el seu guard
  es mantenen. La passada **no obre cap sessió de correu pel seu compte**:
  passa per `_CpPreparaEsborranys` (`ControlsCpEmail.ps1`), la mateixa funció
  que el botó, que és qui obre `Open-CorreuSessio 'outlook-esborrany'`. Abans tot
  vivia dins del botó, amb finestres; copiar-lo per a l'automàtic era la manera
  segura que divergissin (guard: un sol `Send-CorreuSessio` al fitxer). Els
  esborranys de l'automàtic s'apunten a `Add-CorreuEsborranysPendents` i el menú
  avisa, com amb els recordatoris.
- **L'historial** (`controls-periodics-avisats.json` a
  `local\base-dades-activitats`, porta GIA i dates: mai al GitHub): per GIA, el
  *Proper CP previst* del qual ja s'ha preparat el correu. **També hi apunta el
  botó**: el que fas a mà, l'automàtic no ho repeteix. La clau és la data: si
  l'Excel en porta una de nova (s'ha passat el control i n'hi ha un altre de
  previst), torna a tocar (`_CpJaAvisat`). S'hi apunta **encara que s'aturi a
  mig camí** (el `finally`): el que ja és a Esborranys no es torna a fer. La
  finestra en treu la columna «Correu preparat» (`_CpMarcaAvisats`).
- **Màxim 20 per passada, primer els que han vençut fa menys**
  (`_CpAutoTriades`, pura). La primera vegada hi pot haver centenars de
  controls vençuts de fa anys; el que val la pena que surti sol és el que acaba
  de vèncer. La resta va sortint (o es fa a mà).
- **La CCO «a triar», mai** en automàtic (guard de sempre, amb els dos fitxers
  nous a la llista).

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

- **Octubre 2026, el que l'usuari no entenia de la finestra** (es va partir:
  la finestra és ara a `RecordatorisPantalla.ps1`; `Recordatoris.ps1` havia
  passat de les 1.200 línies):
  - **«Només els que toquen avui» en marcava molts més que el màxim** (155 amb
    un màxim de 15): `_RecDueActivitats` marcava `Sel` a TOTES les que tocaven
    i el topall només s'aplicava en enviar. Ara **`_RecPreselecciona`** (pura)
    en marca com a molt `maxPerTanda`, en l'ordre de prioritat, i l'etiqueta
    diu quantes n'hi ha de marcades.
  - **La casella «Campanya activa» ja no hi és**: deia el mateix que el
    Manual/Automàtic de cada pestanya. Ara una campanya en **Automàtic** és la
    que entra a la passada de la tasca del Windows (que s'engega amb
    l'interruptor A/M de la rajola). Una configuració d'abans **apagada i en
    Automàtic** passa a **Manual** (`_RecNormalitzaConfig`): cap campanya no
    pot començar a enviar per una actualització.
  - **Una ⓘ al costat de «Cada (dies)», «Espera (dies)», «Màx. per tanda» i
    del Manual/Automàtic** (`_AddInfoIcona`, `UiFinestra.ps1`: globus i també
    un clic, perquè el globus es pot no veure). Els textos són a
    `_RecAjudaCamp`.

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
- **Tercera ronda: les COL·LECCIONS** (ITC de Bombers, TINSCI, ITC antigues).
  No tinc la llista (el proxy d'aquí bloqueja interior.gencat.cat) i Interior en
  publica de noves: l'entrada porta `Colleccio: true` i l'URL de la pàgina, i
  `_NormativaBaixaColleccio` en treu els PDF cada vegada
  (`_NormativaDocsDeColleccio`; si no n'hi ha, el DOM de l'Edge; si tampoc, un
  nivell de pàgines filles, `_NormativaSubpagines`). Cada document va a l'estat
  amb `Pare` = la col·lecció i surt a l'índex sota seu. El nom és
  `Ambit_Tema_<text de l'enllaç>.pdf` (`_NormativaNomDocColleccio`, sense el
  «(PDF, 1,2 MB)» del final).
  - Les llistes tornen **sense coma** (`return $out.ToArray()`): el cridador fa
    `@()`, i amb la coma en sortia una llista d'una llista. Ho van enxampar les
    proves.
  - El bucle de baixada és ara **`Invoke-NormativaBaixada`** (sense finestra):
    el fan servir l'eina Normativa i la revisió.
- **Quarta ronda: les TINSCI «sense cap document»**. La pàgina de les TINSCI
  **no enllaça els PDF**: enllaça la FITXA de cada document al repositori
  d'Interior (DSpace, `dsp.interior.gencat.cat/handle/20.500.14007/<n>`), que és
  un altre servidor, i `_NormativaSubpagines` només acceptava pàgines del mateix
  lloc. Ara hi entren les fitxes del DSpace de qualsevol servidor
  (`$Script:NormativaDspaceFitxa`, també `hdl.handle.net` i les `/items/<uuid>`
  del DSpace nou), i a la fitxa es reconeixen els fitxers
  (`/bitstream/handle/…/X.pdf?sequence=…`, que ja passava, i
  `/bitstreams/<uuid>/download`). Els enllaços relatius de la fitxa pengen de
  l'adreça **final** després de les redireccions (`_NormativaUrlFinal`).
  L'adreça que va passar l'usuari és a la prova. **No provat contra el
  servidor real** (proxy): si el DSpace torna a no donar res, mirar l'HTML de
  la fitxa amb l'Edge.
  - Els noms de les ITC portaven **«(Obre en una nova finestra)»** (el text
    per a lectors de pantalla de la web d'Interior). Es treu, i els fitxers ja
    baixats amb el nom vell **es reanomenen** (per l'adreça, a l'estat) en lloc
    de baixar-los de nou i deixar el vell a la carpeta.
  - **La ⓘ d'un punt d'ITC obre el PDF de la ITC** (`_NormativaSpDeText` +
    `_NormativaFitxerSp`, abans del catàleg): la fitxa cita també la Llei
    3/2010 i, si no, obria la llei.

- **Cinquena ronda (índex de l'1 d'octubre de 2026: 254 baixades, 5 errors)**.
  - **Les «ITC Bombers antigues» eren les TINSCI**: el marcador
    «Incendis_SP-XXX (antigues)» apuntava a la pàgina «Documentació normativa:
    TINSCI», i els 36 documents baixats són DT-x (vigents i anteriors). Ara són
    una font més de la col·lecció TINSCI (`AltresUrls`), i la pàgina nova
    (documents-tinsci) es continua provant; els documents de totes dues es
    reuneixen sense repetits. Els fitxers ja baixats **es reanomenen** perquè
    la col·lecció porta `Abans` = «Col·lecció ITC Bombers antigues». Les
    versions «Document anterior/antic» van a Antic
    (`Incendis_Antic_TINSCI Document anterior DT-5.pdf`).
  - Pàgines filles: un nivell més **només si són del repositori (DSpace)** (la
    pàgina pot enllaçar la col·lecció i no cada document), i l'Edge només a la
    pàgina de la col·lecció i a les del repositori: abans es feia a les 150
    filles, fins a 45 s cadascuna.
  - **RD 1002/2002**: no té text consolidat i el BOE respon **404** a `/con`.
    L'`Invoke-WebRequest` llança, i el respatller sense `/con` (que ja hi era) no
    s'arribava a provar mai. **Arreglat això, l'ELI sense `/con` també va tornar
    404** (segon índex de l'1 d'octubre): ara el catàleg apunta a
    `buscar/doc.php?id=BOE-A-2002-19574`, la publicació original, que porta
    l'enllaç `/boe/dias/…pdf` que llegeix `_NormativaBoeInfo`.
  - **2016/679 i 2017/745**, de la pàgina del BOE (DOUE-L-2016-80807 i
    DOUE-L-2017-80916), com els altres reglaments europeus.
  - **Circular 093 d'APABCN**: és a l'àrea privada (cal iniciar sessió). Camp
    `Manual: true` (`_NormativaFontDe`): l'índex en conserva l'enllaç i diu
    «Web amb accés restringit: obre l'enllaç i desa-la a mà amb aquest nom».

## REQ1: la subsecció «ITC de Bombers» (Incendis, setembre 2026)
24 punts (un per ITC vigent del web de Bombers; SP 144 i SP 147 amb
sub-punts), escrits **llegint els PDF oficials** que l'usuari va deixar al Drive
(carpeta «Normativa»). Tres coses per si s'han de tornar a tocar:
- **Els PDF de 2021 ençà tenen les xifres mal codificades** a la capa de text
  (`1, m` per `1,60 m`, `5 m²` per `500 m²`): ni `pypdf` ni el lector del Drive
  les treuen. Es van llegir **renderitzant les pàgines a imatge** (PyMuPDF). No
  es va copiar cap xifra de la capa de text sense veure-la a la imatge.
- La **Nota aclaridora de la DGPEIS** (ITC anteriors al RD 164/2025) deixa sense
  aplicació SP 103, 107, 108, 116, 117, 119, 122, 123 i 140 (no hi són) i adapta
  SP 113, 121, 128, 131 i 145: els punts ja porten el text adaptat i la fitxa
  ho diu. La **Nota 1 de la SP 144** (Llei 11/2026): des del 14.7.2026 les
  activitats esporàdiques en espais oberts ja no tenen informe de la DGPEIS.
- Les **TINSCI** van a la seva subsecció (més avall).
- **Reorganitzades a l'octubre 2026** (petició de l'usuari). L'ordre de les
  subseccions d'Incendis és ara **CTE DB SI · RSCIEI · RIPCI · ITC de Bombers ·
  TINSCI** (els quatre punts solts del principi no es mouen). Les ITC es
  parteixen en els **grups de la web de Bombers**, cada un amb el seu text fix:
  *ITC de Bombers - RSCIEI*, *- CTE DB SI*, *- Genèriques*, *- Certificats SP
  136* i *- Altres*. **D'on surt el grup de cada ITC** (la web d'Interior no
  s'hi arriba des d'aquest entorn): l'ordre que l'aprova, que ja era a la
  fitxa — ISP/19/2025 = RSCIEI, ISP/20/2025 = DB SI, ISP/28/2025 = genèriques
  —; les sis del 2012, per la capçalera del DOGC dels PDF (SP 109, 110 i 114,
  una disposició; SP 113, 120 i 121, l'altra) i el contingut (INT/323/2012 =
  DB SI, INT/324/2012 = genèriques). **SP 126, 138 i 147 no les aprova cap
  ordre** i van a *Altres*, amb un text fix que no diu «caràcter
  reglamentari»: és una deducció, si la web les posa en un altre grup cal
  moure-les. Títols amb «ITC SP nnn» i textos escurçats als conceptes clau
  (tots del text anterior). L'script és `scratchpad/py/itc_reorg.py` (de la
  sessió).
- **La subsecció «Documentació (ITC SP)» ja no existeix**: l'intro de
  l'Ordenança i els models A, B i C de l'SP 136 són *ITC de Bombers -
  Certificats SP 136*, i les **lluernes en coberta** són un punt del RSCIEI
  genèric. `LLIC.json` expandeix ara aquella subsecció al bloc DESPRÉS (la
  clau) — o sigui que **a Llicència les lluernes han passat de DESPRÉS a
  PROJECTE**, que és el que toca per a un requisit del RSCIEI.

## REQ1: la subsecció «TINSCI» (Incendis, octubre 2026)
16 punts, un per document TINSCI vigent (DT-4 a DT-19; DT-9 amb tres
sub-punts), escrits **llegint els PDF** que l'usuari va deixar al Drive. Les
DT-1, 2 i 3 només hi són com a versions «anteriors» i no tenen punt.
- **No són normativa**: són criteris de la Taula d'Interpretació. Per això el
  text del punt diu «segons el document TINSCI DT-x» i la competència de la
  fitxa ho diu (la DGPEIS l'aplica en l'annex 1 de la Llei 3/2010; l'Ajuntament
  el pot prendre com a referència). Als documents, el **negre és normatiu i el
  gris és criteri TINSCI**: les fitxes ho recorden on importa.
- Referències velles dins dels documents: DT-7 i DT-12 citen el RD 2267/2004
  (ara RD 164/2025). La fitxa ho diu; el text del punt no ho reprodueix.
- De la DT-18 hi havia **dues versions** al Drive (febrer i desembre 2023): el
  punt segueix la de desembre. `_NormativaFitxerDt` tria igualment
  l'«actualitzat» quan n'hi ha dos.
- Els annexos de la DT-13 (taules de forjats) tenen el text il·legible a la capa
  de text; el punt només hi remet.
- **La ⓘ d'un punt TINSCI obre el PDF** (`_NormativaDtDeText` +
  `_NormativaFitxerDt`): només si la norma diu «TINSCI», perquè «DT» sol és
  massa curt per reconèixer-ho a qualsevol text.
- `REQ1.json` es va escriure imitant el `ConvertTo-Json` del PowerShell 5.1
  (el diff només té línies afegides), i `docs/dades/cataleg-REQ1.json` es va
  refer amb `ExportaDades -Plantilles` i el mateix format.

## Revisar requeriments (eina NORMATIVA, setembre 2026)
Petició de l'usuari: una eina per «actualitzar el programa» que miri si la
normativa dels requeriments segueix vigent, si els enllaços funcionen, si cal
baixar normativa nova que substitueixi l'anterior i si tots els punts tenen la
fitxa ⓘ. **Informa, no canvia res**: l'informe (Excel, `local\revisions`) diu
què passa i on; canviar un requeriment es fa llegint la norma.
- **`RevisioDades.ps1`** (pur): `_RevPuntsSenseFitxa` (punts i sub-punts sense
  `norma` ni `criteri`), `_RevEstatBoe` (derogada només si la pàgina ho diu de
  la norma sencera —«Norma derogada», «Estado: derogada», «queda derogada por»—;
  una derogació parcial no; captura la substituta amb la `Ref. BOE-A-…`),
  `_RevEstatPjur` (l'etiqueta **VIGENT/DEROGAT en majúscules just després de
  «Copia la URI ELI»**: un «(Derogat)» dins del text d'un article no compta, i
  «DEROGATÒRIA» no és «DEROGAT»). Si la pàgina no ho diu clar → `?` («mira-ho a
  mà»): val més dir-ho que endevinar-ho. **No s'ha pogut provar contra el BOE ni
  el Portal Jurídic reals** (proxy): els patrons surten del que ensenyen les
  pàgines (la captura de l'usuari per al Portal Jurídic).
- **`Enllacos.ps1`** (només defineix): `_EnllacosDeCataleg` (text **i** fitxa,
  amb el punt) i `Test-EnllacViu`. Abans eren dins de `Comprova-Enllacos.ps1`,
  que a més **cridava `Read-JsonFile` sense carregar `Json.ps1`** i petava a cada
  catàleg; ara hi carrega els dos fitxers.
- **`Revisio.ps1`**: la finestra (quatre caselles) i la vigència de cada norma
  (BOE: la pàgina; Portal Jurídic: el DOM de l'Edge, perquè l'etiqueta es posa
  amb JavaScript).
- Menú: fila nova **NORMATIVA** amb 📚 Normativa i 🔍 Revisar requeriments.
- Proves a `tests/proves/10-revisio.ps1`.
- **Quarta ronda: el Portal Jurídic sense l'Edge** (l'usuari, amb l'índex d'una
  primera passada: l'Edge es penjava 2 minuts per norma i la passada es va
  quedar a la 5a).
  - **El botó PDF del Portal Jurídic apunta al DOGC**, per número de versió:
    `portaldogc.gencat.cat/utilsEADOP/AppJava/PdfProviderServlet?versionId=N&type=01`
    (adreça copiada per l'usuari de la Llei 3/2010). `_NormativaFontsPjur` busca
    el número **sense navegador**: a la pàgina tal com la dona el servidor i a les
    metadades ELI (les descàrregues RDF/TTL/XML que enllaça la pàgina, i
    `<eli>/rdf|ttl|xml`), amb `_NormativaPdfPjurDeText` (l'enllaç sencer o un
    `versionId` en dades JSON). **No s'ha pogut veure si el número hi és**: si
    no hi és, la baixada cau a l'Edge com abans.
  - **L'Edge, en un sol lloc** (`_NormativaEdge`): perfil **nou a cada crida**
    (un de compartit quedava bloquejat pel que s'havia penjat i feia penjar els
    següents), `taskkill /T` de tot l'arbre si no acaba, i si es penja **un cop**
    ja no es torna a fer servir en aquella passada (`$Script:NormativaEdgeKO`).
    Temps: 45 s el DOM, 60 s imprimir.
  - La revisió mira la vigència del Portal Jurídic primer a les metadades ELI
    (`_RevEstatEli`: `InForce-inForce` / `notInForce` / `partiallyInForce`) i a
    la pàgina del servidor; l'Edge només si no ho diuen.
- **Cinquena ronda** (índex de la segona passada de l'usuari: 207 baixades, 54
  errors). Les del Portal Jurídic que es van baixar ho van fer **amb l'Edge**
  (el DOM dibuixat): el número de versió no surt sense navegador. El que va
  fallar: la pàgina del **CIDO** va penjar l'Edge i, com que la retirada era
  global, **34 normes del Portal Jurídic** que venien darrere no el van poder
  fer servir. Ara la retirada és **per servidor** (`$Script:NormativaEdgeKO`
  és un hashtable; del tot només si es penja amb 3 webs diferents).
  - Els reglaments europeus: **EUR-Lex no deixa baixar fora d'un navegador**
    (torna una pàgina de comprovació). Els cinc que tenen la fitxa `DOUE-L` al
    BOE (marcadors de l'usuari) es baixen d'allà; el 2016/679 i el 2017/745
    segueixen a EUR-Lex (no tinc el seu `DOUE-L` i un de fals baixaria un altre
    reglament amb el seu nom).
  - BOE sense text consolidat (RD 1002/2002): si la pàgina `/con` no dona
    l'identificador, es prova l'ELI sense `/con`.
  - L'índex diu **per quin camí** s'ha baixat cada una (`Baixada (PDF de la
    pàgina, amb l'Edge)`…), per poder-ho diagnosticar sense l'estat.

## «Excel Comprovar» (abans «Comprovar Excel»): visites d'inspecció i avisos descartats (octubre 2026)

> **El nom.** L'usuari (octubre 2026) la va reanomenar «Excel Comprovar», al costat
> de la nova «Excel Errors». **L'acció segueix sent `comprovarexcel`**, i és a
> posta: el segell d'«última vegada» del menú s'indexa per l'acció, i canviar-la
> l'hauria deixat a «(mai)».

L'usuari: «ha de poder descartar avisos […] si jo considero que està arreglat
per una altra banda» i «els informes que faig amb Objecte Visita inspecció +
data […] aquesta data ha de constar a la fulla "Estès" columna "Data control
inicial/verificació"».

- **La data de la visita entra a la base d'informes** (`visita`, `dd/MM/yyyy`;
  `''` = mirat i no és una visita). La treu `_ExtractVisitaInspeccio` de la línia
  «Objecte: Visita inspecció …» (`Get-InformeData`). Els informes que ja hi eren
  (sense el camp) la llegeixen a l'**Actualitzar base** següent **sense tornar-se
  a classificar** (només la capçalera; un `.doc`, de la memòria del text).
- **Què es comprova**: la **darrera** visita de cada activitat (sense els
  informes ignorats) ha de ser entre les dates de la cel·la
  (`_DatesDeCella`: número d'Excel, DateTime o text amb una o més dates). Només
  la darrera, a posta: la columna és de *data de control* i en porta la més nova.
  Si la columna no hi és, es diu al peu i no s'avisa de res.
- **Cada avís té una clau estable** (`precinte|GIA`, `visita|GIA|data`,
  `visitaqueixa|GIA|data`, `notrobat|GIA`, `sensegia|carpeta`; `Get-ComprovacioExcelAvisos`, pura).
  **Descartar** desa les claus a `local/base-dades-activitats/comprovar-excel-descartats.json`
  (porta noms de titulars: mai al GitHub). Un descartat que ja no avisa es treu
  sol de la llista (`_CxFiltraDescartats`), i una visita **nova** té una altra
  clau: torna a avisar. **Recuperar descartats** els torna a ensenyar tots.

### El MOTIU de la visita: Queixa / Pla de verificació (octubre 2026)

L'usuari: «he de poder diferenciar […] si aquella visita és per una Queixa o pel
Pla de verificació 2026 (l'any vinent serà 2027)», també totes dues
(«aprofito la queixa per fer el control inicial»), i que Excel Comprovar miri
**Queixa → columna `Inspecció`** i **Pla → `Data control inicial/verificació`**.

- **Ha de constar a l'informe**, i és d'allà que es llegeix: l'Objecte porta
  `Visita inspecció 22/10/2026 - Queixa`, `… - Pla de verificació 2026` o
  `… - Queixa + Pla de verificació 2026`. **L'any és el de la data de la
  visita**, no el d'avui (una visita del desembre redactada al gener).
- **Un sol vocabulari**: `MOTIU_VISITA` de la capçalera = `queixa`, `pla`,
  `queixa+pla` o `''`. `_MotiuVisitaNet` (MotorInforme.ps1) passa qualsevol
  text a aquest valor (tolera majúscules i accents; `\bpla\b`, perquè «plaça» no
  és el pla) i `_MotiuVisitaText` en fa el text; `_BuildOrigenText` el posa
  darrere de la data. Al mòbil, `Correu.motiuVisitaText` / `origenText`
  (`docs/correu.js`), amb **prova creuada** a `08-correu-format`.
- **On es tria**: Pas 2 del PC (caselles *Queixa* / *Pla de verificació* al lloc
  de la data d'anotació; `_MotiuVisitaDeCaselles`) i Pas 2 del mòbil
  (`#motiu-queixa`, `#motiu-pla`). **Totes dues obliguen** a marcar-ne almenys
  un: sense motiu, Excel Comprovar no sabria quina columna mirar. El paquet del
  mòbil el porta (`HEADER_KEYS`, `Paquet.ps1`), i *Recuperar dades* el torna a
  posar. Del correu, `_OrigenDesDeText` separa la data del motiu.
- **La base d'informes** desa `visita_motiu` (`_ExtractVisitaMotiu` de la línia
  Objecte, a `Get-InformeData`). La `visita` segueix sent **només la data**: el
  «2026» del pla no és cap data.
- **Excel Comprovar** (`_CxVisitesPerMotiu` + `Get-ComprovacioExcelAvisos`):
  la darrera visita de **cada** motiu, a la seva columna. Grups i claus:
  `visita|GIA|data` (Pla), `visitaqueixa|GIA|data` (Queixa) i `visitasense`
  (informes d'abans, sense motiu: només si és la visita més nova, i val
  qualsevol de les dues columnes). **La clau del Pla i de la de sense motiu és
  la d'abans**, a posta: els avisos que l'usuari ja havia descartat segueixen
  descartats. Si falta una columna, el seu motiu no avisa i es diu al peu.

## «Pla verificació» (fila BASE D'INFORMES, octubre 2026) — `PlaVerificacio.ps1`

L'usuari, amb el Programa de verificació 2026 i la Memòria del 2025 al davant:
«vull que l'eina mostri les següents taules (en relació a la base d'informes)»,
amb la definició de cada indicador (vegeu la capçalera de `PlaVerificacio.ps1`
i el `LLEGEIX-ME`). Abans del 2025 no hi ha programa.

- **Tot el càlcul és pur** (`Get-PlaVerificacioIndicadors`): rep les files de la
  fulla Estès (`_ReadPlaVerificacio`, amb `Read-FullaEstesa`), l'any, les
  previstes i els informes per GIA (`_PvInformesPerGia`, sense els ignorats).
  `$null` d'informes = sense base: les greus surten `-`, no un 0 que enganyi.
- **Les columnes** (`_PvColumnes`): `Inspecció` ha de ser **exacta** (n'hi ha
  d'altres que porten la paraula) i n'hi ha dues de *Classificació general
  annex*: es pren la primera, com Excel Errors.
- **Les dates de l'Excel nou són TEXT** («2025-07-25 00:00:00.0»): les llegeix
  `_DataDeCella` (`Excel.ps1`, la poden fer servir els dos processos). Això va
  destapar que **Excel Comprovar no en veia cap**: `_DatesDeCella` només
  entenia dd/mm/aaaa i totes les visites sortien com a no apuntades. Ara també
  passa per `_DataDeCella`.
- **Greus**: de les del programa, el primer informe (no ignorat) **de visita
  d'inspecció** (el camp `visita` de la base, de l'«Objecte: Visita inspecció»)
  del mateix any que el control i amb conclusió breu `Precinte / Cessament`
  (`_PvInformePrecinte`). L'usuari: «ha de tenir Precinte / Cessament però amb
  l'objecte de visita d'inspecció». Un precinte d'una documentació aportada no
  és una discrepància trobada a la verificació. La llista ensenya quin.
- **`suport/pla-verificacio.json`** (`Read-PvDades` / `ConvertTo-PvDades`, que
  el passa a hashtables): els **programes** signats (`previstes`, `annex1` = les
  de criticitat més alta en l'ordre del PDF, `annex2` = les aleatòries) i les
  **fotos** dels anys tancats (les dades de la Memòria). Són documents públics
  signats i només porten IDs d'activitat: poden anar al GitHub. Un any nou
  s'hi afegeix quan hi hagi el programa signat; mentrestant, les previstes es
  posen a la finestra i es desen a `local/base-dades-activitats/pla-verificacio.json`
  (`Read/Save-PvPrevistes`), que mana sobre el del repositori.
  - Les llistes es van treure dels PDF amb `pypdf` (una fila comença després
    de la criticitat de l'anterior; el peu de pàgina de Firmadoc s'ha de
    saltar, si no hi surten IDs falsos com «1» o «2025») i es van comparar amb
    les pàgines. Del 2025 només hi ha les 30 previstes («seran les 30 primeres
    de la llista»); del 2026, les 30 de l'annex I i les 15 aleatòries.
- **LA FOTO** (l'usuari: «de moment posa de foto del 2025 les dades de la
  memòria 2025»). `Get-PvFotoIndicadors` dona la mateixa forma que
  `Get-PlaVerificacioIndicadors` a partir de les llistes de la Memòria (els
  comptes surten de les llistes: programa = annex no II/III, EC-PCA = III);
  sense dates, i dins de cada grup en l'ordre de la Memòria (un `Sort-Object`
  no és estable a PowerShell 5.1). La finestra ensenya la foto si l'any en té,
  amb una casella per calcular-ho amb l'Excel d'ara.
- **SEGUIMENT DEL PROGRAMA** (botó secundari *Activitats del programa...*;
  l'usuari: «però no com a cosa primordial»): `Get-PvSeguimentPrograma`, per
  cada activitat prevista: verificada (control de l'any), verificada un altre
  any, pendent o *No és a l'Excel*; amb foto, el que diu la Memòria. El botó
  no es diu «Seguiment…» a posta: el guard dels botons de peu fets a mà hi veu
  un «Següent».
- **Copiar les taules** posa al porta-retalls l'HTML de les dues taules amb el
  format de la Memòria (`_PvTaulesHtml`) en format CF_HTML (`_PvCfHtml`: les
  posicions són en **bytes**, per això l'HTML és ASCII i els accents van com a
  entitats, `_PvHtmlEsc`) i el text tabulat de respatller.
- Comprovat amb l'Excel real del 14/10/2026: el 2025 surten 8 del programa
  (la Memòria en diu 10: la 760 i la 425 tenen la data de control buida, la
  311 ja no és a l'Excel, i hi surt la 546), 9 EC-PCA (la Memòria, 7: la 366
  i la 1291 de més) i 20 denúncies (21: la 600 ja no és a l'Excel). Les
  diferències són de l'Excel, no del càlcul; per això el 2025 surt amb la foto.
- Proves: `tests/proves/16-pla-verificacio.ps1`.

## «Excel Errors» (fila GIA, octubre 2026) — `ExcelErrors.ps1`

L'usuari: «una nova eina que comprovi errors de la base de dades d'activitats
[…] que falta alguna dada de les importants (raó social, representant legal,
adreça, cadastre, UTM, classificació, activitat), dades del tècnic a
representant legal, etc. […] extreure en un Excel indicant el ID Activitat i el
tipus d'error (deixa'm marcar el tipus d'error que vull que s'exporti)».

- **`$Script:ExTipus` és l'ÚNICA llista de tipus**: la finestra, la tria desada i
  l'exportació hi beuen. Un tipus nou s'hi afegeix i surt a tot arreu.
- **La fila es comprova en una funció PURA** (`Get-ExcelErrorsFila`): rep els
  valors per clau de columna i les columnes (`_ExColumnes`, per la capçalera amb
  `_ColumnaAmbTermes`). **Una columna que no hi és (0) no fa saltar cap error**:
  aquella comprovació no es fa i el peu diu quina falta. Si el GIA reanomena una
  capçalera, es veu al peu en lloc de sortir 1.300 errors falsos. Pel mateix
  motiu, el respatller 23/25 de `Get-ColumnesContacte` (mòbil i correu) només
  val si la capçalera parla de mòbil o de correu.
- **Res d'escriure les regles dues vegades**: les errades de correu, telèfon i
  NIF són `_CtErrada` / `_CtTipusNif` (el repàs de contactes); el control
  periòdic, `_ControlPeriodicClassify` (l'eina Controls periòdics). «Dades del
  tècnic» i «no quadra amb els documents» són els avisos de la **base de
  contactes** (`es_el_tecnic`; `falta`, `diferent`, `canvi_titular`). Els
  `error` de la base **no** hi entren: són els mateixos que es miren aquí a
  l'Excel i sortirien dues vegades. Sense base de contactes, el peu ho diu.
- **Les UTM** s'entenen amb coma decimal i punts de milers (`_ExNumero`; la
  `ConvertTo-UtmNumber` de rutes és de l'altre procés). La zona
  (`$Script:ExUtmZona`) és generosa a posta: vol enxampar X/Y girades, graus,
  un zero de més; no una parcel·la del límit del terme.
- **El cadastre**: 20 caràcters, o els 14 de la parcel·la.
- **La tria es desa** a `local\base-dades-activitats\excel-errors.json` com a
  llista de **desmarcats**: així un tipus nou surt marcat.
- **L'exportació** és `_NormativaXlsxBytes` (sense Excel): *ID Activitat*,
  *Classificació general annex*, *Classificació general Apartat* de la base i
  *Tipus d'error* a les primeres columnes (l'usuari va demanar l'annex al costat de l'ID i treure la raó
  social de l'Excel; a la graella de la finestra s'hi queda). La graella i
  l'exportació ordenen pel mateix lloc (`_ExTriats`); **l'Excel del GIA no s'obre mai
  per escriure** (guard).
- **El CheckedListBox avisa ABANS** del canvi (`ItemCheck`): la finestra porta
  el canvi pendent a `$st.Canvi` en lloc de refer la graella més tard.
- **Els tres llistats per correu i l'automàtic** (octubre 2026, l'usuari: «vull
  automatitzar també els errors: un correu amb 3 Excels», Innòcues, Bars i
  Restaurants, Annexos II i III). `ExcelErrorsLlistats.ps1`: el repartiment
  (`_ExLlistatDe`, pura, compara l'annex sense accents ni signes: «L18-Proj. i
  Cert.» → `l18projicert`; el 561 amb `_ControlPeriodicClassify`), les tres
  fitxes, les activitats **tretes dels llistats** (`excel-errors-exclosos.json`,
  per GIA) i el correu (`Send-ExcelErrorsCorreu`, com el del Seguiment: adreces
  fixes, només Outlook perquè porta adjunts). L'automàtic
  (`ExcelErrorsAutomatic.ps1` + `ExcelErrorsAuto.ps1`, cada dues setmanes) fa
  **la mateixa** `Invoke-ExcelErrorsLlistats` + `Send-ExcelErrorsCorreu` que el
  botó (guard), amb la tria desada i mai la CCO «a triar».
  - **Qui no va a cap llistat**, a posta i dit a l'usuari: annex buit, annex I,
    L18-Cert. o Sectorial amb 561 i qualsevol altre valor. Amb els Excels que va
    passar l'usuari (76 activitats amb errors): 15 sense annex i 1 de l'annex
    I.1. La finestra en mostra la columna «Llistat» i el recompte.
  - **La finestra va a `ExcelErrorsPantalla.ps1`**: la fan servir la lectura
    (ExcelErrors.ps1) i els llistats; si fos dins d'ExcelErrors.ps1, aquell i el
    dels llistats dependrien l'un de l'altre (guard de cicles).
  - El guard de les closures va enxampar `$Script:ExLlistats` dins del `$fill`
    de la graella (hauria llegit buit): es captura en una variable local.
  - **Qui va a «(cap)», segona ronda** (l'usuari, amb les 16 al davant i
    l'Excel del 14/10/2026): «Deixa a (cap) les "Sense llicència o no comunicada"
    i les autoritzacions ambientals. Les altres que ara mateix van a cap haurien
    d'anar a innòcues», i «segueix vigent el criteri d'Annexos i Bars i
    Restaurants». O sigui: primer els criteris dels Annexos i de Bars i
    Restaurants (no canvien), després Innòcues; el que no hi entrava va a
    Innòcues **llevat de** l'annex I i l'estat «sense llicència». Per això la
    lectura porta ara la columna «Estat de legalització de l'activitat» (`ESTAT`
    a `_ExColumnes`; no és cap comprovació). A tota la base, 9 a (cap).

## Tres canvis petits (octubre 2026)

- **Informes programa** (Configuració, sobre els Automatismes): la clau
  `OferirCorreuInforme` (`'si'`) dels settings decideix si `Offer-EnviarCorreu`
  obre la finestra d'enviar en acabar un informe (`_OferirCorreuInforme`,
  pura). **Sense la clau, no surt**: l'usuari no envia correus mentre no
  funcioni bé Actualitzar base («sempre cancel·lo»). L'informe ja s'ha obert al
  Word abans; no cal cap altre avís.
- **Generar ruta**: l'activitat (*Activitat principal*, per nom de capçalera
  amb `Find-HeaderColumn`) surt sota l'adreça, més petita i en gris (`.act`), a
  la llista i a la fitxa del punt.
- **REQUERIMENT ANTERIOR**: si l'informe comença pel requeriment anterior, la
  frase d'intro («S'han observat les següents deficiències…») ja no va a dalt:
  surt només sota REQUERIMENT ACTUAL (`Build-CatalegBlocs`). El fitxer d'or de
  REQ1 va perdre aquelles dues línies i res més. El correu del mòbil ja ho feia
  així (el seu intro és el del correu).

## Excel Errors: tres tipus nous (octubre 2026)

L'usuari, amb un Excel d'errors fet a mà («falten columnes»): la resta ja els
trobava l'eina; faltaven aquests.
- **`superficie`**: ni *Superfície edificada* ni *Superfície útil* (la primera,
  la del costat de l'edificada: n'hi ha dues) no són un número més gran que 0.
  Un 0 també és no tenir-la. A l'Excel del 14/10/2026: 91 activitats (les 83 de
  la llista de l'usuari menys les que han canviat des d'aleshores).
- **`data_registre`**: *Data registre entrada* buida (14, totes «sense
  llicència o no comunicada»).
- **`cp_annex3`**: si és **annex III**, la *Periodicitat CP* ha de ser **8** i
  no buida (49 buides). Aquell cas ja no surt també com a `control_periodic`
  (no es diu dues vegades). El *Proper CP previst* **no es mira**: el posa sol
  el GIA a partir de la periodicitat (l'usuari). Una primera versió també
  comprovava que fos el darrer control + 8 anys; es va treure.
- Les diferències que queden amb la llista de l'usuari són activitats que han
  canviat entre els dos Excels o que ja no hi són.
