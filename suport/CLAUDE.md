# Notes per a Claude (mantenir entre sessions)

## On és cada cosa (mapa de mòduls)
`Motor.ps1` havia arribat a **3.316 línies i 70 funcions**, amb la configuració,
l'Excel, el Drive, la selecció, els camps, les conclusions, el document, el mode
mòbil i el menú tot barrejat. Es va partir. **El mapa complet i actualitzat és a
la capçalera de `suport/Motor.ps1`** — mira'l allà abans de tocar res; aquí només
el resum:

| Fitxer | Què hi ha |
|---|---|
| `Motor.ps1` (635 l.) | rutes, config, càrrega de mòduls, `%LOCALAPPDATA%`, instància única, Word, accés al catàleg |
| `Wizard.ps1` | `Main` + assistent de "Requeriment - Nou" |
| `Capcalera.ps1` · `SeleccioItems.ps1` · `Camps.ps1` · `Document.ps1` | els passos 2, 3+5, camps i composició |
| `Activitats.ps1` | Excel d'activitats: caché per ID GIA + pujada a Drive |
| `Paquet.ps1` | generar sense assistent (mòbil) |
| `Migracio.ps1` | rutes de `local/` (`Get-LocalSubdir`) + endreç de les carpetes velles |

**Partir un fitxer és barat i segur**: tot va amb dot-source al **mateix àmbit**,
o sigui que moure una funció d'un fitxer a un altre no en canvia el comportament.
L'únic que compta és l'**ordre de càrrega**, i només per als mòduls que
**calculen alguna cosa en carregar-se** (`Activitats.ps1` → `$LocalActivitatsDir`,
`ActExtr.ps1` → rutes del registre): han d'anar **després** del bloc de rutes.
En partir, verifica que la llista de noms de funció de tot `suport/` és idèntica
abans i després.

## EXECUTA LES PROVES. De debò, executa-les
`suport/tests/run-tests.ps1` passa sencer **en un Linux sense Word ni Excel**:

```
apt-get install -y powershell   # o el tar.gz de github.com/PowerShell/PowerShell
GENINFORME_TEST=1 pwsh -NoProfile -File suport/tests/run-tests.ps1
```

Val la pena insistir-hi perquè durant molt de temps **no es van executar mai**
(en aquell contenidor no hi havia `pwsh` i es validava tot amb rèpliques en
Python). El dia que es van poder executar van sortir **quatre errors reals de
seguida**, tres d'ells amb les proves ja escrites i afirmant el que tocava:

1. L'eina *Seguiment* petava a la primera pestanya (el `@()` de
   `_FindCampInfoPairs`, vegeu la seva secció).
2. La crida de prova a `Parse-ActExtrTemplate` passava **dos** arguments: la
   signatura havia canviat en treure el lector de `.docx` i ningú no ho havia
   vist. El `throw` **matava la resta de la suite**, o sigui que tot el que hi
   havia després no s'executava.
3. Dins d'un literal `@(...)` la **coma lliga més fort que el `+`**: a les dades
   de prova, `'DEN' + [char]0x00DA + 'NCIA?'` sense parèntesis es convertia en
   TRES elements i desalineava la fila sencera. **Parentetitza sempre** les
   concatenacions dins d'un array.
4. Les amplades de columna anaven desplaçades una posició (vegeu *Seguiment*).

Coses a tenir en compte perquè la suite pugui córrer fora de Windows:
- **`Join-Path` resol la UNITAT**: fora de Windows, `Join-Path 'C:\x' 'y'` peta
  amb *"A drive with the name 'C' does not exist"*. El codi de producció el pot
  fer servir (només corre a Windows), però **les proves de rutes han de fer
  servir una arrel vàlida a la plataforma on corren** (`$tstClone`, `$tstSep`…).
  `[System.IO.Path]` tampoc no serveix per a rutes de Windows en un Linux: allà
  la `\` no és separador.
- **System.Drawing (GDI+) no hi és fora de Windows**, i l'excepció del
  carregador de tipus **s'escapa del `try/catch`** perquè salta en *compilar* el
  cos de la funció, abans de la primera línia. Per això el guard de plataforma va
  al **cridador** (`_AutoFirmaVisibleExtraParams`) i no dins de
  `_BuildCaixetiImageBase64`.

## La carpeta `local/`: què és del repositori i què no
- **`ESTRUCTURALS/` = FONTS**: els 5 `.json` + `0 CAPCALERA.docx` (l'única
  plantilla de Word de veritat, que no es pot regenerar). Res més.
- **`local/` = tot el que és d'aquest ordinador**, i el `.gitignore` l'exclou
  **sencera** (`local/*` + `!local/README.txt`). Hi van: informes generats,
  mapes de ruta, l'Excel local, `informes-db.json`, el registre de la signatura
  i les **vistes en Word** dels catàlegs (derivades, es regeneren soles).
- Això és **seguretat, no estètica**: abans un fitxer local nou a l'arrel es
  pujava per defecte i calia recordar-se d'afegir una regla; el
  `pdf-signar-log.txt`, que porta noms i adreces de titulars, ja s'hi va escapar
  un cop **en un repositori públic**. Ara no pot passar.
- Els noms de les subcarpetes viuen **NOMÉS** a `Migracio.ps1`
  (`$Script:LocalSubdirs` + `Get-LocalSubdir`). `rutes/Ruta.ps1` s'executa en un
  procés propi i abans repetia `'BASE DE DADES ACTIVITATS'` pel seu compte: era
  **l'única duplicació real del projecte**; ara demana la ruta al mateix lloc.
- **NO** es mou `%LOCALAPPDATA%\InformesCornella\` (settings, caché, còpies dels
  catàlegs, credencials del Drive, `running.pid`): és estat d'usuari de Windows,
  ha de sobreviure a tornar a clonar, i les credencials del Drive no han de
  viure dins d'una carpeta que es pugui comprimir i enviar.
- `Invoke-MigracioLocal` (idempotent, no llança mai) la criden `Motor.ps1` en
  arrencar i `Actualitzar.bat` al pas 4a, després del `pull`.

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
- **Impressió** (i per tant el PDF): horitzontal, **A3**, `Zoom=$false` +
  `FitToPagesWide=1` + `FitToPagesTall=$false` (hi caben totes les columnes),
  marges 0,5 cm, `PrintTitleRows='$1:$2'` i peu `&A` / `Página &P`. Al PDF hi
  van **només els 5 llistats**, seleccionats com a grup i exportats amb
  **`$excel.ActiveWindow.SelectedSheets.ExportAsFixedFormat`** — amb
  `ActiveSheet` només sortiria una pestanya.
- Escriure amb COM va **per matriu** (`$range.Value2 = $matriu`), no cel·la a
  cel·la: la diferència és de minuts a segons.
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

## Trampa: el que surt d'un CMDLET ve embolcallat en un `PSObject`
`Join-Path` és un **cmdlet**, i el seu resultat arriba dins d'un `PSObject`.
Normalment no es nota (PowerShell el desembolcalla sol), **però NO quan el valor
s'ha de passar a una crida COM per referència**:

```
$doc.SaveAs([ref]$out, [ref]16)
→ no se puede convertir el valor "...\REQ1.docx" de tipo "psobject" al tipo "Object"
```

Va passar exactament això en canviar `_VistaWordPathFor` de
`[System.IO.Path]::ChangeExtension(...)` (mètode .NET → `String` net) a
`Join-Path` (cmdlet → `PSObject`): **cap vista es va poder desar**. Solució:
`return [string](Join-Path …)` a `Get-LocalDir`/`Get-LocalSubdir`/
`_VistaWordPathFor` **i** `[string]$out = …` al punt d'ús.

No es pot detectar amb una prova pura (`-is [string]`, `.GetType()` i
`.psobject.BaseObject` diuen `String` en tots dos casos), o sigui que la regla
és: **si un valor ha d'anar a un `[ref]` d'una crida COM, força'n el tipus**.

## Res de llegir `.docx` per treure'n contingut
El lector de `.docx` (`Parse-Cataleg`, les branques `.docx` de `Read-Conclusions`
i `Parse-ActExtrTemplate`, `Read-ConclusionsXml`, `Test-StyleMatch`) es va
esborrar. **No el tornis a posar "per si de cas"**: els `.docx` d'ESTRUCTURALS ja
no són catàlegs, són **vistes generades en format d'informe**. Un respatller que
els llegís no fallaria — generaria un informe **silenciosament equivocat**. Si el
`.json` no hi és, val més petar amb un missatge clar.

## Arquitectura: motor, punt d'entrada i UiComuns
- **`suport/Motor.ps1`** = la base. NOMÉS defineix (funcions, rutes,
  configuració); carregar-lo no obre res ni genera res. Abans es deia
  `GenerarInforme.ps1` i era alhora motor i programa.
- **`suport/GenerarInforme.ps1`** = el punt d'entrada (~45 línies): té el
  `param(-DesDePaquet)`, carrega `Motor.ps1` i crida `Main` (o
  `Invoke-GenerateFromPaquet`). És el que llancen el `.bat`/`.vbs`.
- **Reutilitzar el motor com a biblioteca**: `$MotorSenseGui = $true` i
  dot-source de `Motor.ps1`. Ho fan `mobil/Vigilant.ps1` i
  `mobil/ExportaDades.ps1`. **Ja no cal `$env:GENINFORME_TEST = '1'`** per
  evitar que s'executi el programa: aquesta bandera ara només vol dir "no
  carreguis WinForms" i la fan servir les proves.
- **`suport/UiComuns.ps1`** = helpers de WinForms compartits; **es carrega el
  primer** de tots els mòduls i no coneix res del motor. Hi viuen `_NewForm`,
  `_AddBrandHeader`, `_AddStepBar`, `_StylePrimaryButton`/`_StyleSecondaryButton`,
  **`_MakeMultiFilter`** (abans al punt d'entrada) i **`_AddConfigRow`** (abans
  a `Configuracio.ps1`). Regla: si un helper d'interfície el fan servir dues
  pantalles, va aquí — mai a la pantalla que el va estrenar.
- `_MakeMultiFilter().GetSelected` retorna **sempre** un `[string[]]` (el `,`
  del `return ,([string[]]$s)` és deliberat). Consumeix-lo com
  `$sel = & $mf.GetSelected`, **sense `@()`** al voltant: `@()` l'embolcallaria
  en un array d'un element i trencaria el "cap opció marcada = passa tot".
- **Helpers de graella** (`UiComuns.ps1`): `_StyleListGrid` (carcassa),
  `_AddSearchBox`, `_EnableHeaderSort` + `_SetSortGlyph` (ordre programàtic amb
  fletxa). Els fan servir *Editar base d'informes* i *Controls periòdics*.
  L'estat d'ordre s'ha de dir **`$state.SortColIdx` / `$state.SortAsc`** (és el
  que espera `_EnableHeaderSort`). Deliberadament **NO** hi ha un "constructor
  de graelles" únic: les columnes, el filtratge i l'ordenació de les dues
  pantalles són massa diferents (una agrupa sempre per activitat, l'altra
  ordena per data real i té casella de selecció) i el genèric sortiria més
  complicat que les dues pantalles juntes.
- **Compte amb les col·lisions de noms**: tot el codi cau al mateix àmbit
  global (dot-source), i **el darrer fitxer carregat guanya, en silenci**.
  Abans d'afegir una funció nova, `grep` del nom. Ja va passar: es va crear un
  `_TextMatches` a `UiComuns.ps1` sense veure que ja n'hi havia un a
  `Motor.ps1` (el del cercador del TreeView de catàlegs) — el de `Motor.ps1`
  el tapava i les proves ho van destapar. El filtre de text de les graelles
  reutilitza el de `Motor.ps1`; se li passa el text de cerca **ja net**
  (`.Trim().ToLower()`), que és el seu contracte.

## El clone viu en una unitat de XARXA: manteniment del git desactivat
- El clone de l'usuari està a `\\fitxers\arrel\Activitats_Ordenances\...`. Allà, el
  **`geometric-repack`** que el git llança sol després d'un `fetch`/`push` **falla**:
  no pot reanomenar el `.idx` (SMB el té bloquejat) i escup
  `Permission denied` + `error: task 'geometric-repack' failed`.
- El pitjor no és el soroll: git **PREGUNTA** `Should I try again? (y/n)` i això pot
  deixar `Actualitzar.bat` **ATURAT esperant una tecla** (al registre de l'usuari
  s'hi veu fins i tot un `Sorry, I did not understand your answer`).
- Per això `Actualitzar.bat` (a dalt de tot) i `Instalar.bat` (subrutina
  `:NO_MAINTENANCE`) desactiven el manteniment **en aquest clone**:
  `git config maintenance.auto false` + `git config gc.auto 0`, i s'hi posa
  `GIT_ASK_YESNO=false` per si de cas. El repositori és petit: no perdem res.

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

## CRÍTIC: els catàlegs de l'usuari no es poden perdre mai
- **Què va passar (bug real, no ho repeteixis):** `Actualitzar.bat` detectava
  canvis a **qualsevol** fitxer d'`ESTRUCTURALS` però només feia
  `git add "ESTRUCTURALS/*.docx"`. Els **`.json` dels catàlegs** (els que escriu
  l'editor "Editar catàlegs") no es committejaven mai, quedaven bruts, i el
  `git stash push -u` del pas següent se'ls enduia. El `pull` restaurava la versió
  del repositori i **la feina de l'usuari desapareixia del programa** (quedava al
  stash, però ningú no la treia mai d'allà). A sobre, el mòbil SÍ que rebia els
  canvis (les dades es regeneren abans del stash), o sigui que PC i mòbil quedaven
  incoherents.
- **Regla:** els catàlegs editats a l'ordinador de l'usuari **SÓN L'AUTORITAT**.
  Es committegen, es pugen a `main` i **prevalen** sobre el que baixi del repositori.
- **Com està resolt** (`suport/SincronitzaCatalegs.ps1` + `Actualitzar.bat`):
  1. **Còpia de seguretat SEMPRE i abans de tocar res de git** (`-Fase Backup`) de
     tot el que l'usuari pot editar (`ESTRUCTURALS` + `docs/dades`) a
     `%LOCALAPPDATA%\InformesCornella\backups\<data-hora>\` amb `manifest.txt`.
  2. `git add "ESTRUCTURALS/*.docx" "ESTRUCTURALS/*.json" "docs/dades/*.json"` +
     commit (així **no** queda res brut i el stash no se'n pot endur res).
  3. Després del `pull`, **`-Fase Restore`** torna a aplicar els fitxers de la
     còpia → la versió de l'usuari **preval** — i es committeja i puja.
  4. Si el `rebase` troba un **conflicte**: `rebase --abort` + `reset --hard
     origin/main` i es tornen a aplicar els catàlegs de la còpia (mai es deixa
     l'usuari amb un rebase a mitges).
  5. Si tot i així quedés res brut a `ESTRUCTURALS`, s'avisa **ben visible** amb la
     ruta de la còpia (mai un stash silenciós).
- **`Recuperar-catalegs.bat`** (arrel) → `-Fase Recuperar`: busca als stashes antics
  els catàlegs que es van perdre amb la versió anterior i els **extreu** a
  `%LOCALAPPDATA%\InformesCornella\recuperats\` (no fa `pop`: no toca res del clone).
- `.gitignore` té `ESTRUCTURALS/*.bak` (còpies que fa l'editor en desar).
- Funcions pures amb tests: `_CatalegEsProtegible`, `_ParseGitStatusPaths` (compte:
  git **enquota** els noms amb espais, i retorna **array pla**), `_CatalegsBackupName`.

## IMPORTANT: tota la feina ha de convergir a `main`
Cada sessió de Claude Code (web) treballa en una branca pròpia `claude/...`.
Si la feina es queda només en aquesta branca, **l'usuari no la rebrà mai**
amb el seu `git pull` de `main`, i semblarà que "no ha canviat res".

Per tant, **al final de cada sessió**:
1. Assegura't que tot està commitejat a la branca de la sessió.
2. Fusiona la feina a `main` i fes push de `main`:
   ```
   git fetch origin
   git rebase origin/main
   git push origin HEAD:main
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
  `local\base-dades-activitats\informes-db.json` (dins de `local/`: mai es puja).
  Botons al menú: **🗃 Actualitzar** i **📋 Editar** (marc "Base d'informes").
- Lectura de `.docx` **sense Word** (zip) reutilitzant les primitives de
  `Seguiment.ps1` (`_LoadDocxXml`, `_ParagraphTextXml`). Lectura de `.doc`
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
  aplica `ItemSpaceAfterPt` (12 pt) en lloc de `BulletSpaceBeforePt` (6 pt) al
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
- **Menú Pas 1 — 4 apartats de rajoles** (`Select-Mode`, `Seguiment.ps1`, helper
  `$addTileRow`; dispatch al `switch` de `Main`):
  - **EINES** (3): 📍 *Generar ruta* (`ruta`), 🔒 *Activitats precintades*
    (`url`, acció `precintades` només per al segell), 📅 *Controls periòdics*
    (`controlsperiodics`).
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
  (`UiComuns.ps1`) — quadre editable + "..." + indicador ✓/⚠ en viu — igual que a
  Configuració (tots els selectors de carpeta del programa fan servir aquest format).
  - **Per defecte hi surt l'ÚLTIM INFORME GENERAT** (`_UltimInformeGenerat`): el
    `.docx`/`.doc` més nou de la carpeta de sortida, sense `-Recurse` (els informes
    es desen plans), saltant els temporals `~$`. El cas d'ús real és sempre el
    mateix — acabes de generar un informe i el vols passar a PDF i signar-lo — i
    abans tocava anar-hi a buscar cada vegada. Si no se'n troba cap (o la unitat de
    xarxa no hi és) es recupera la cascada d'abans: última ruta desada →
    `$InformesDir` → buit. `_CarpetaInformesGenerats` repeteix la cascada de
    `_ResolveOutputDir` (`Motor.ps1`) però **sense crear la carpeta**: obrir un
    diàleg no ha de crear res al disc. Es pot canviar sempre: el quadre és editable
    i hi ha els botons **Carpeta** i **Document**.
  - **Signatura VISIBLE (caixetí)**: en signar, per defecte s'afegeix un **caixetí a
    dalt a la dreta de la pàgina 1** que reprodueix l'aspecte "CERTIFICAT SENSE DNI"
    (nom / càrrec / organisme / data, **sense DNI**). AutoFirma per línia de comandes
    **NO pot triar un "Aspecto" desat de la GUI**, així que es reprodueix amb
    `layer2Text` via `-config` (`_AutoFirmaVisibleExtraParams` → `signaturePage=1`,
    `signaturePositionOnPage...` dalt-dreta, `layer2Text` amb les línies del caixetí
    unides per `\n` LITERAL i el marcador `$$SIGNDATE=yyyy.MM.dd HH:mm:ss$$`). El text
    del caixetí és **editable** al diàleg d'opcions (casella "Signatura visible" +
    quadre de text) i es desa a `pdf-signar-state.json` (`caixeti`, `visibleSign`);
    per defecte `_DefaultCaixeti`. Sense caixetí, `_BuildAutoFirmaSignArgv` es comporta
    com abans (signatura invisible, cap `-config`).
  - **Format de `-config`: propietats separades pel `\n` LITERAL (2 caràcters),
    MAI per salts de línia reals.** Això no és una suposició, és el codi
    d'AutoFirma. `CommandLineParameters.java` guarda el valor **tal qual** (NO el
    descodifica de Base64) i l'ordre `sign` el converteix a `Properties` amb
    `CommandLineLauncher.buildProperties()`, que fa
    `while ((endIndex = params.indexOf("\\n", beginIndex)) != -1)`: en Java el
    literal `"\\n"` són els caràcters `\` + `n` i `indexOf` busca una **cadena
    literal**, no una expressió regular. El mateix codi ho diu al comentari:
    *"La division no funciona correctamente con split porque el caracter salto de
    linea se protege al insertarse por consola, asi que lo hacemos manualmente."*
    (Compte: `loadSignConfig`, unes línies més amunt, sí que fa `split("\n")` amb
    salts reals — però **aquest camí no és el de `sign`**. És exactament la trampa
    en què vam caure.) Amb salts REALS el bucle no troba res i AutoFirma es queda
    amb **una sola propietat**: clau `signaturePage` i com a valor tota la resta →
    **signa bé (codi 0) però sense caixetí**. Al PDF es reconeix perquè el widget
    surt amb `/Rect[0 0 0 0]` i l'aparença amb `/BBox[0 0 0 0]`. Un intent encara
    anterior el passava en **Base64**, amb el mateix resultat. `$Script:AutoFirmaConfigSep`
    conté el separador; posició tunejable a `$Script:AutoFirmaCaixetiPos`.
  - **Codi de sortida 0 NO vol dir caixetí visible**: `_PdfCaixetiEsInvisible`
    mira **només el tros que el signador ha afegit al final** del PDF (la revisió
    incremental, des de la mida del fitxer original) i hi busca `/BBox[0 0 0 0]`
    (`_PdfTextCaixetiInvisible`, pura). Si el caixetí ha quedat invisible, l'intent
    es dona per fallat, es registra i es passa al següent. Sense això el programa
    donava el fitxer per bo i l'usuari no veia res, sense cap avís.
  - **LES COMETES LES POSEM NOSALTRES (2n error, també real)**: a PowerShell 5.1
    `Start-Process -ArgumentList @(...)` **NO enquota** els elements: els ajunta
    amb espais. Amb rutes com `5.- Sergi Fadurdo`, AutoFirma rebia la ruta
    **tallada** al primer espai i responia *"El fichero de entrada no existe:
    I:\…\5.-"*. Ara `_ArgvToCommandLine` (pura) construeix la línia i enquota, i
    s'executa amb **`ProcessStartInfo`** (`_RunAutoFirma`), que dona control
    exacte de la línia d'ordres i recull sortida i codi de sortida. Les barres
    invertides **finals** es dupliquen quan l'argument va entre cometes: si no,
    la barra escaparia la cometa de tancament (`CommandLineToArgvW`).
  - **La data la resolem NOSALTRES (3r error real)**: amb
    `$$SIGNDATE=yyyy.MM.dd HH:mm:ss$$` dins de `layer2Text`, AutoFirma petava amb
    *"Error no reconocido: begin 0, end -1, length 21"* (un `substring` amb un
    índex no trobat) i no signava. `_ResolveCaixetiDate` (pura) substitueix el
    marcador per la data abans de cridar AutoFirma i **treu qualsevol altre
    `$$...$$`**, de manera que AutoFirma no veu mai cap marcador. A la interfície
    el marcador es manté (l'usuari pot triar el format de data).
  - **EL `layer2Text` NOMÉS ADMET UNA LÍNIA (4t error real, ja confirmat)**, i ara
    se n'entén el perquè: el `\n` literal és el **separador de propietats**, de
    manera que un salt dins del `layer2Text` fa que AutoFirma hi talli i que el
    tros següent (sense cap `=`) faci petar
    `keyValue.substring(0, keyValue.indexOf('='))` amb `indexOf` = −1. És
    exactament l'error del registre: *"Error no reconocido: begin 0, end -1,
    length 21"*, i **21 són els caràcters de `Enginyer d'Activitats`**, la 2a línia
    del caixetí. Per això `_AutoFirmaVisibleExtraParams` en mode `text` **sempre**
    posa el caixetí en una sola línia (`_CaixetiUnaLinia`, unit amb ` · `, que
    també talla pel `\n` literal i treu qualsevol barra invertida que quedi).
  - **Per tenir-lo de DIVERSES LÍNIES: com a IMATGE**. `_BuildCaixetiImageBase64`
    (System.Drawing) dibuixa el caixetí en un JPEG amb la mateixa proporció que el
    requadre de la signatura i el passa a `signatureRubricImage` (base64), que és
    el mecanisme documentat d'AutoFirma per a la rúbrica. En aquest mode NO s'hi
    posa `layer2Text` (la imatge ja ho porta tot).
  - **LÍMIT DE LA LÍNIA D'ORDRES (5è error real)**: la imatge viatja en **base64
    dins de l'ordre**, i Windows no admet més de **32767** caràcters. Amb la
    imatge a escala x4, `Process.Start` petava amb *"El nombre del archivo o la
    extensión es demasiado largo"*. Ara la imatge va a **escala x2 amb JPEG de
    qualitat 70** i hi ha **dos topalls**: `$Script:MaxCaixetiBase64` (la imatge
    no es genera si passa de 20000 car.) i `$Script:MaxCommandLine` (un intent que
    no hi cabria **ni es prova**, se salta i es registra).
  - **El `try/catch` va DINS del bucle d'intents**: quan estava a fora, una
    excepció d'un intent s'enduia **tots** els altres i el fitxer es quedava
    **sense signar**. Va passar exactament això amb la imatge massa gran.
  - **Reintents escalats**: caixetí **(imatge)** → caixetí **(text d'una línia)**
    → **sense** caixetí. Així mai es queda un PDF sense signar i del registre se'n
    dedueix quin ha funcionat. Al registre la imatge surt **resumida** (mida en
    caràcters), mai el base64 sencer, que faria el fitxer inservible.
  - **El registre distingeix els salts**: `_AutoFirmaArgvToText` mostra els salts
    de línia REALS com a `<LF>` i deixa els `\n` LITERALS tal qual. Abans tots
    dos sortien com a `\n` i el log no permetia saber quin era quin — que és
    justament el que calia per depurar.
  - **Registre de diagnòstic**: cada execució desa a `pdf-signar-log.txt` (al costat
    de `pdf-signar-state.json`) **l'ordre exacta** passada a AutoFirma, el codi de
    sortida i la seva sortida (`_PdfSignarLog`, `_AutoFirmaArgvToText`). El resum
    ofereix obrir-lo. Serveix per no haver d'endevinar si el caixetí no surt.
    **Va al `.gitignore`**: conté les RUTES COMPLETES dels informes i les carpetes
    porten el nom i l'adreça del titular — i aquest repositori és PÚBLIC. (S'hi
    ignora pel nom exacte, no `*.txt`, perquè el `README.txt` d'aquella carpeta sí
    que va al repositori.)
  - **Carpeta O document**: el quadre accepta una **carpeta** (tots els Word de dins
    i subcarpetes) o **un sol document** Word; hi ha dos botons (Carpeta / Document).
    `_RunConvertPdf` ho distingeix amb `Test-Path -PathType Leaf`.
  Funcions pures testejables (`_PdfPathForDoc`, `_PdfShouldConvert`, `_CertFilterValue`,
  `_CertCommonName`, `_BuildAutoFirmaSignArgv`, `_AutoFirmaVisibleExtraParams`,
  `_AutoFirmaArgvToText`, `_DefaultCaixeti`, `_AutoFirmaCandidatePaths` — retorna un
  **array pla**, no `,$ArrayList`, perquè `@()` l'enumeri bé). Word (COM) i AutoFirma
  només a Windows. Opcions/estat a `pdf-signar-state.json`.
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
    "nodes":[ {"tipus","titol","clau"?,"cos":[<paràgraf>],"fills":[<node>]} ] }
  <paràgraf> = { "runs":[ {"t","b","i"} ], "url": bool }
  ```
  Un **run** és un fragment de text amb negreta (`b`) i/o cursiva (`i`); un
  paràgraf pot ser `"url": true` (enllaç). Els **fills són imbricats de veritat**
  dins cada node (una subsecció conté els seus ítems, un ítem els seus subítems…);
  el **nivell = fondària d'imbricació**, no un camp. El `tipus` de cada node fa
  servir el **MATEIX vocabulari a tots els catàlegs** (petició de l'usuari: "mateixos
  noms a tot") i, amb la `familia`, dona la semàntica. El lector torna a mapar el
  `tipus` al `Kind`/`Style` intern d'abans → **generació byte-idèntica**. Vocabulari:
  - `seccio` — contenidor de 1r nivell (secció de catàleg; grup de conclusions amb
    `titol`=tipus d'informe; secció visual d'ACT_EXTR).
  - `subseccio` — sub-contenidor de catàleg que agrupa els seus ítems (fills).
  - `item` — unitat de contingut (deficiència / conclusió / bloc d'ACT_EXTR).
  - `subitem` — sub-ítem amb pic (fill d'un ítem; a ACT_EXTR = bloc `::CHILD::`).
  - `text` — text/introducció (a ACT_EXTR = bloc `::TEXT::`).
  - `sempre` — (conclusions) frase que s'inclou sempre (`::SEMPRE::`).
  - `nota`/`etiqueta`/`capcalera`/`paragraf` — (ACT_EXTR) estils
    `::NOTE::`/`::LABEL::`/`::HEADER::`/`::CONC::` de l'informe favorable.
  - `intro` (camp de dalt): cos fix (TERMINI no té nodes) o capçalera (conclusions).
  Per família: `cataleg` (REQ1, TERMINI) fa servir seccio→(item|subseccio|text),
  subseccio→item, item→subitem; `conclusions` (0 CONCLUSIONS) fa servir
  seccio→item i seccio `sempre` a l'arrel; `actextr` (ACT_EXTR_REQ/FAV) fa servir
  seccio→(item|subitem|text|nota|etiqueta|capcalera|paragraf).
  - **`clau`**: només a ACT_EXTR. És la `[[KEY]]` funcional (Decret 112) del bloc,
    ara un **atribut a part** (abans vivia DINS el títol junt amb `::CHILD::`). El
    lector reconstrueix la capçalera `"[[clau]] ::TOKEN:: titol"` que espera
    `Build-ActExtrBlocks` (que només fa servir clau + token i **ignora l'etiqueta**),
    de manera que els blocs surten idèntics. Les claus funcionals NO es toquen mai
    (`INCENDIS`, `PAU_CAT/LOCAL`, `VIGILANTS`, `RC`, `MEMORIA_A..G`, `REQ_INTRO`,
    `FAV_*`…). L'editor la mostra en un camp **bloquejat** (mai s'edita des d'aquí).
- **ELS `.docx` JA NO SERVEIXEN PER GENERAR: la font de veritat és el `.json`.**
  `Get-Catalegs` llista **`*.json`** d'ESTRUCTURALS (abans `*.docx`), i
  `$ConclusionsPath`, les plantilles d'ACT_EXTR (`ActExtr.ps1`), el catàleg de
  *Controls periòdics* i `Invoke-GenerateFromPaquet` apunten tots al `.json`.
  **Per què:** l'usuari va apartar els `.docx` (ja no calien) i van
  **desaparèixer del menú** "Requeriment - Nou" i "Ampliació de termini", perquè
  el menú es construïa llistant `*.docx`. `Read-ConclusionsXml` (Seguiment, que
  llegia el `.docx` com a ZIP) també delega ara al JSON.
- **`0 CAPCALERA` es queda en Word**: és una carta amb escut/taula/format real
  (la generació COPIA el .docx i hi substitueix `<<PLACEHOLDERS>>`), no un
  llistat reconstruïble des d'un model de runs. És **l'únic `.docx` que és una
  plantilla de veritat** i **no es pot regenerar**: `Actualitzar.bat` el recupera
  del repositori (`git checkout --`) si falta, en lloc de commitejar-ne l'esborrat.
- **VISTES en Word dels catàlegs** (`suport/VistaWord.ps1` + `suport/GeneraVistes.ps1`):
  la resta de `.docx` d'ESTRUCTURALS són ara **vistes generades des dels JSON**
  per poder consultar tot el contingut (tots els requeriments, totes les
  conclusions…) sense obrir el programa. **El format és EXACTAMENT el de l'informe**:
  criden les mateixes `Format-*` de `Format.ps1` que fa servir `Build-Document`
  (secció en MAJÚSCULES, subsecció subratllada, ítems numerats amb el número en
  negreta, fills amb pic, URLs com a hipervincle, mateixos espaiats). A sobre,
  cada títol rep un **nivell d'esquema** (`OutlineLevel` 1/2/3) perquè surti al
  **panell de navegació** de Word: l'OutlineLevel NO canvia com es veu el
  paràgraf, només el fa navegable. Compte: Word **hereta** el nivell al paràgraf
  següent, per això el cos el torna sempre a 10 (`wdOutlineLevelBodyText`). Es regeneren **en desar des de
  l'editor de catàlegs** (`_Ed_SaveDoc`) i des de **`Actualitzar.bat`** (al pas
  **4b, DESPRÉS del `pull`**: si es fessin abans, es generarien amb la versió
  ANTIGA del programa i un canvi de format no hi arribaria mai — caldria executar
  `Actualitzar.bat` dues vegades), i sobreescriuen el `.docx` del mateix nom (mai
  `0 CAPCALERA.docx`). Quan canvia `$Script:VistaWordVersio` es regeneren **totes**
  una vegada (la versió es desa a `%LOCALAPPDATA%`, no al repositori).
  **Només si el JSON és més nou que la vista** (`_VistaCalRegenerar`): si es
  regeneressin sempre, cada `Actualitzar.bat` faria un commit d'un `.docx` "nou"
  (Word hi posa dates internes) i el repositori s'ompliria de canvis inútils.
  Funcions pures amb tests: `_VistaWordPathFor`, `_VistaEsProtegit`,
  `_VistaActExtrTitol`, `_VistaCalRegenerar`.
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
  `*.json`). Arbre de nodes amb **fills imbricats de veritat** (les subseccions i
  els subítems pengen del seu pare) + entrada especial per a la introducció/
  capçalera; a la dreta s'edita títol, **Tipus** (vocabulari unificat; el combo
  s'omple amb `_Ed_TipusOptions $familia $parentTipus` i és **SEMPRE canviable**
  quan hi ha >1 opció) i **cos amb negreta/cursiva reals** (RichTextBox ↔ runs) amb
  botons per inserir `[CAMP:]`/`[OPCIO:]` i enllaços. A ACT_EXTR es mostra a més un
  camp **Clau** (`ReadOnly`) amb la `[[KEY]]` funcional, que no s'edita mai. El tag
  de cada TreeNode duu `ParentNode` (per calcular els tipus vàlids). Es poden
  afegir/eliminar/moure nodes. En desar s'escriu el JSON (sense BOM; `clau` només
  quan té valor) amb **validació** (re-llegeix amb el lector) i **còpia `.bak`** de
  seguretat; l'editor només ESCRIU, la generació no en depèn.
  - **Punt d'entrada**: a la finestra principal (`Seguiment.ps1 Select-Mode`),
    el **xip del document** (REQ1/TERMINI/ACT_EXTR) dels botons de tipus d'informe
    porta un emoji d'editar ✏️; clicar-lo (hit-test del rectangle `DocChipRect`
    via `MouseClick`) obre l'editor centrat en aquell document. Dispatch:
    `'editcataleg' → Show-CatalegEditor -focusDoc <Doc>` (ACT_EXTR → ACT_EXTR_REQ).
  - **Funcions pures testejables** (headless): `_Ed_JsonToModel`/`_Ed_ModelToJson`
    (model editable ↔ JSON; node `@{tipus;titol;clau;cos;fills}`), `_Ed_SegmentsToRuns`
    (fragments RTB → runs, forçant la invariant de no-solapament), `_Ed_CosToRich`/
    `_Ed_RichToCos`, `_Ed_TipusOptions`/`_Ed_DefaultTipus`/`_Ed_CanAddChild`/
    `_Ed_ChildTipus` (vocabulari unificat segons família i tipus del pare). Tests a
    `run-tests.ps1` comproven que model→JSON→lector és **idèntic** a llegir l'original
    (sense pèrdues) per als 5 fitxers. La finestra (WinForms) només es prova a Windows.
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
  `{PROPER_CP}{DATA_CONTROL}{DATA}` + `**negreta**` + enllaços) via el botó **"Editar
  text"** (`Invoke-ControlsCpEmailTextos`, mateix patró que *Textos del correu*),
  però es desa **LOCALMENT** a `%LOCALAPPDATA%\InformesCornella\controls-cp-email.json`
  (fora del repo: cap dada personal, sobreviu a `Actualitzar.bat`). Funcions pures
  testejades: `_DefaultControlsCpEmail`, `_ControlsCpRecipients`, `_FillControlsCpPh`,
  `_ControlsCpLineHtml`/`_ControlsCpEmailHtml`. Outlook (COM) i finestres només a
  Windows. Es dot-sourceja a `GenerarInforme.ps1` després de `ControlsPeriodics.ps1`.

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
