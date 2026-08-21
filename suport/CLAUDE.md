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
- **`_StyleListGrid` fa `Dock='Fill'` — i això és un CONTRACTE, no un detall**:
  està pensada per a graelles que viuen **dins d'un panell** (Editar base,
  Controls periòdics). Si la graella conviu amb botons posats a mà **al mateix
  formulari**, el `Dock` la fa ocupar TOTA la finestra i **tapa els botons**:
  la pantalla sembla morta (cap botó visible ni clicable). Va passar de debò a
  la tria de documentació de Llicència — «no em deixa posar següent ni enrere
  ni res». Solució: desfer el `Dock` i fixar `Location/Size/Anchor` **després**
  de l'estil (abans, l'estil ho trepitja). Hi ha prova de font que ho vigila.
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
- **Segona part del mateix bug, mesos després:** `0 CAPCALERA.docx` **està al
  repositori i és una plantilla de veritat**, però quan es va passar tot a JSON
  el `git add` es va deixar en `"ESTRUCTURALS/*.json"`. Resultat: l'usuari
  canvia la mida de l'escut de la capçalera i **el canvi no es committeja mai**
  → el `git pull --rebase` es nega a començar (*"You have unstaged changes"*) →
  es va al camí d'error → el `git reset --hard origin/main` **se l'endú**. Només
  se salvava per la còpia de seguretat del pas 2, i a l'execució següent tornava
  a passar exactament igual. Ara hi ha **`git add -u -- ESTRUCTURALS`** als dos
  commits (estadia el que ja té seguiment) i **no** un `add ESTRUCTURALS/*.docx`
  a pèl, que en un clone antic pujaria les vistes en Word encara sense migrar.
  Lliçó: quan es treu un tipus de fitxer d'un flux, s'ha de mirar si **algun
  d'aquells fitxers seguia sent una font**.
- **UN BINARI NO ES POT FUSIONAR, i «l'usuari mana» hi vol dir una altra cosa**
  (incident real, agost 2026). L'usuari tenia `0 CAPCALERA.docx` retocat i el
  repositori hi acabava d'afegir el bloc `[[CAP:LLIC]]`. El `rebase` va petar
  (*"Cannot merge binary files"*), es va anar al camí d'error i, tot seguit, el
  **Restore va tornar a posar-hi la còpia de l'usuari a sobre**: el bloc va
  desaparèixer i la versió sense el bloc **es va pujar a `main`**. La regla és
  bona per als `.json` (text, i com a molt es torna a escriure un text), però en
  un binari vol dir **llençar el fitxer sencer de l'altra banda**, i és allà on
  hi ha les peces que el programa necessita.
  - `_CatalegEsBinari` + `_CatalegHiHaColisio` (pures): si un fitxer binari ha
    canviat a les **dues** bandes, el Restore **no el toca**. Es queda la del
    repositori (el programa la necessita sencera), la de l'usuari es queda a la
    còpia de seguretat i s'avisa amb la ruta. Com que el fitxer no es modifica,
    el clone queda **net** i per tant no es puja res: ningú decideix en silenci.
  - La comparació es fa **pel sha del blob** (`git rev-parse <base>:<ruta>`
    contra `git hash-object`), no llegint binaris amb PowerShell. El **commit de
    base** l'apunta el Backup a `base.txt` (corre abans del commit i del pull,
    o sigui que `HEAD` encara és d'on venia el fitxer de l'usuari).
  - **Sense sha de base no es decideix en contra de l'usuari**: es torna a
    aplicar la seva còpia, que és el comportament de sempre.
  - `Actualitzar.bat` recull el **codi de sortida 2** i torna a dir-ho **al
    final**: enmig de l'actualització l'avís queda amunt i no es veu.
- **La prova que havia de protegir la capçalera anava amb `-like`** i per tant no
  provava res: en un patró de `-like`, `[[CAP:LLIC]` és una **classe de
  caràcters**, o sigui que `*[[CAP:LLIC]]*` dona per bo qualsevol text amb un
  dels caràcters `[ C A P : L I` seguit d'un `]`. Amb aquest `.docx` encertava de
  casualitat (no hi ha cap `]`). Ara va amb `.Contains()` i **sobre el text
  descodificat dels paràgrafs** — al XML cru el marcador surt escapat
  (`&lt;&lt;CLASSIFICACIO&gt;&gt;`) i buscar-hi `<<CLASSIFICACIO>>` no trobaria
  mai res. Hi ha també una prova que no hi hagi cap **etiqueta sense marcador**
  (una línia acabada en `:` sense cap `<<...>>` al darrere només pot sortir
  BUIDA a l'informe: és el que passava amb `Classificació:` al bloc genèric,
  que sortia en blanc a tots els REQ1 i TERMINI).
- **Res destructiu sense desar-ho abans**: el camí d'error del pull feia
  `rebase --abort` (que escopia *"fatal: no rebase in progress"* quan el pull ni
  havia arrencat) i tot seguit `reset --hard`. Ara l'abort només es fa si hi ha
  un rebase de debò a mitges (`.git\rebase-merge` / `rebase-apply`) i **abans del
  reset sempre es fa un `git stash push -u`**.
- El «hi ha alguna cosa bruta?» del pas 3 es mira amb **`git status --porcelain`**
  i no amb `git diff`: el `diff` **no veu els fitxers sense seguiment**, i que la
  detecció d'allà i la del `pull` no miressin el mateix és com s'arribava a un
  pull que peta havent dit que tot estava net.
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
    (System.Drawing) dibuixa el caixetí amb la mateixa proporció que el requadre
    de la signatura i el passa a `signatureRubricImage` (base64), que és el
    mecanisme documentat d'AutoFirma per a la rúbrica. En aquest mode NO s'hi
    posa `layer2Text` (la imatge ja ho porta tot).
  - **LA RÚBRICA HA D'ANAR EN JPEG. NO EN PNG.** Es va provar de passar-la en PNG
    (amb text pla ocupa molt menys: 16.244 caràcters contra els 35.000 llargs del
    JPEG equivalent) i AutoFirma la va rebutjar amb
    *"Se ha proporcionado una imagen de rubrica que no esta codificada en JPEG"*.
    L'intent amb imatge falla i el caixetí cau al respatller de **text d'una
    línia** — que és com es nota, perquè no surt cap error a la pantalla. Hi ha
    prova que ho blinda (cap entrada de `$Script:CaixetiIntents` pot ser d'un
    altre format).
  - **La lletra es veia BORROSA i era la QUALITAT del JPEG**, no el format:
    s'hi anava a **qualitat 70** i **escala x2** (144 ppp), i el JPEG a qualitat
    baixa deixa halos al voltant del text negre sobre blanc. Ara
    `$Script:CaixetiIntents` és una escala d'intents **de més a menys qualitat**
    (x3 q92 → q88 → q84 → q78 → x2 q95 → q90) i es fa servir **el primer que hi
    càpiga**: així s'aprofita tot el pressupost en lloc d'anar a la fixa amb el
    pitjor. El codificador JPEG del Windows no dona la mateixa mida que cap
    altre, per això hi ha tants graons.
  - **`MaxCaixetiBase64` era massa just** (20000, posat "per si de cas"): amb
    l'escut de fons només hi cabia l'escala x2, justament la borrosa. **Mesurat
    amb una ordre real del registre** (rutes a `I:\…\5.- Sergi Fadurdo\…`, filtre
    amb el CN sencer i les 6 propietats de posició), tot el que **no** és la
    imatge ocupa **628 caràcters** — o sigui que del límit dur de Windows (32767)
    en sobren més de 30.000. Ara `MaxCommandLine` = 32000 i el topall de la
    imatge, 30500. La comprovació de `MaxCommandLine` segueix sent la xarxa de
    seguretat si una ruta fos molt més llarga.
  - **El registre diu quina variant d'imatge ha entrat** (`$Script:CaixetiUltimIntent`
    → línia `imatge: jpeg x3 q88 (28… car.)`), i si no n'hi cap cap, les mides de
    totes les provades. Sense això, quan el caixetí sortia en text calia endevinar
    si havia estat la mida o el format.
  - **Aspecte** (`$Script:CaixetiEstil`, tot tunejable en un sol lloc): contorn
    gris fosc, **escut de l'Ajuntament de fons a la dreta** (pintat amb
    `ColorMatrix.Matrix33` = opacitat; va **abans** del text perquè aquest hi
    passi per sobre) i `FactorLletra` = 0,72 (era 0,58): com més alt, més omplen
    les lletres la línia i **menys espai buit** queda entremig. El requadre ha
    passat de **75 a 48 pt** d'alçada per al mateix nombre de línies.
  - **`Icon.ToBitmap()` NO serveix amb `suport/cornella.ico`**: un `.ico` és un
    contenidor amb diverses mides a dins, i aquest les porta **totes set
    comprimides en PNG** (16…256 px). El .NET, amb icones així, no les
    descomprimeix bé i el resultat surt **buit** — l'escut no apareixia al
    caixetí i **no ho deia ningú**, perquè el dibuix va dins d'un `try/catch`.
    `_IcoTriaFrame` (pura, amb proves contra el fitxer real) llegeix la taula del
    `.ico` — capçalera de 6 bytes i una entrada de 16 per imatge, amb l'amplada
    en **un sol byte** on el 0 vol dir 256 — i en tria la més petita que ja sigui
    prou gran; després el PNG va a `Image.FromStream`. Queda el respatller a
    `Icon(path, Size)` per si algun dia el `.ico` porta imatges DIB.
    `$Script:CaixetiAvisEscut` / `CaixetiEscutDibuixat` ho deixen dit al
    registre: **un `try/catch` que s'empassa un error de dibuix ha de deixar
    rastre**, si no la cosa surt malament en silenci.
  - **Resolució: mesura-la, no la suposis.** Del registre de l'usuari surt que
    l'escala x3 amb qualitat 92 ocupa **26.416 caràcters**, mentre que les meves
    proves n'estimaven 33.192: el codificador JPEG del Windows fa la imatge un
    **20% més petita**. Amb això calibrat, l'escala **x4 (800×192, 288 ppp) hi
    cap fins a qualitat 85**, i amb TEXT més resolució val més que més qualitat.
    La imatge que hi ha dins d'un PDF ja signat es pot treure i mirar
    (`/DCTDecode`) — és la manera de saber què hi ha arribat de debò en lloc de
    discutir sobre una captura de pantalla.
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
  - **LA FIRMA NO ES VALIDAVA ALS ALTRES ORDINADORS** (mesurat comparant els dos
    PDF, no deduït). Al PC de l'usuari sortia vàlida; als dels companys **i al
    visor corporatiu de l'Ajuntament** (`aytos-fdocweb`), *"La validez de la
    firma es DESCONOCIDA / la identidad del firmante es desconocida"*. Un informe
    signat a mà **amb el mateix certificat** sí que s'hi validava.
    - Es va treure el CMS de dins del `/Contents` dels dos PDF i comparar-los:
      **el certificat signant és EL MATEIX** (mateix número de sèrie
      `49160A73…`, TCAT d'empleat públic, emès per `SubCA SECTOR PUBLIC Q (G3)
      A.1`). O sigui que **no era el certificat**, ni caducat ni res.
    - Diferències reals: el nostre anava amb `SubFilter` **`ETSI.CAdES.detached`**
      i **3 certificats a dins** (arrel + subCA + signant), amb
      `signingCertificateV2` (els OID de política que s'hi veuen són **del
      certificat**, dins d'aquell atribut: **no** hi ha cap política de firma —
      les preferències d'AutoFirma diuen «Ninguna política»); el que funciona,
      **`adbe.pkcs7.detached`** i **només 1** (el signant). **Cap dels dos** porta
      segell de temps ni dades de revocació (el `revocationInfoArchival` del que
      funciona és un SET **buit**).
    - **Per què es creu que és això**: l'avís d'Adobe diu que no són de confiança
      *"sus certificados PRINCIPALES"* — i els pares només els té perquè **els hi
      encastem nosaltres**. Si l'ancoratge de confiança d'aquell ordinador no és
      l'arrel de l'AOC sinó la **subCA** (que és el que publica la Llista de
      Confiança europea), donar-li la cadena sencera el fa aturar-se en una arrel
      que no coneix. Que falli **també al visor corporatiu** —que és de servidor i
      té la llista de confiança ben posada— apunta al **fitxer**, no a cada PC.
    - `_AutoFirmaCompatLines` (pura) hi posa `signatureSubFilter=adbe.pkcs7.detached`
      i `includeOnlySignningCertificate=true`. **El nom porta DUES ENES**
      (`Signning`): és així al Client @firma i, si algú l'hi "corregeix",
      AutoFirma l'ignora **en silenci** i tornem a encastar la cadena. Hi ha prova.
    - Aquestes línies van al `-config` **sempre**, amb caixetí o sense: abans el
      `-config` només existia si hi havia caixetí i una signatura invisible no
      se n'hauria beneficiat.
    - Per tornar enrere: `$Script:SignaturaCompat`, i prou.
    - **3a ronda, amb els DOS PDF del MATEIX ordinador** (el signat a mà surt
      vàlid, el nostre desconegut): això **retira definitivament** la teoria del
      magatzem de confiança de l'altre PC — és el fitxer. I els dos canvis
      anteriors **sí que van entrar** (el PDF ja porta `adbe.pkcs7.detached` i
      un sol certificat), o sigui que **tampoc no eren la causa**.
    - **L'OID de l'algorisme tampoc.** Al `SignerInfo` nosaltres emetíem
      `rsaEncryption` i Adobe `sha256WithRSAEncryption`. Es va poder provar
      **sense tornar a signar**, perquè aquell camp **no està cobert per la
      signatura**: es va canviar **1 sol byte** del PDF ja signat (l'OID passa de
      `…01 01 01` a `…01 01 0B`, mateixa llargada, `/ByteRange` intacte),
      l'OpenSSL seguia dient *CMS Verification successful*… i l'Adobe seguia
      dient desconegut. **Aquest truc del byte és reutilitzable** per provar
      qualsevol cosa que estigui fora dels `signedAttrs`.
    - **L'única diferència que queda: `signingCertificateV2`** (atribut ESS que
      AutoFirma posa i Adobe no). El seu `certHash` és **correcte** (SHA-256 del
      certificat, comprovat), però l'atribut hi porta a dins un camp
      **`policies`** amb les polítiques del certificat — i el **RFC 5035** diu
      que això **no és informatiu**: obliga a validar la cadena **restringida a
      aquelles polítiques**. Al PC de l'usuari, amb tota la cadena de l'AOC
      instal·lada, passa; en un altre, no. Això explica per què encastar o no la
      cadena no va canviar res: el que falla no és **trobar** els pares, és
      **validar-los sota aquelles polítiques**.
    - `signingCertificateV2=false` demana la versió antiga de l'atribut
      (SigningCertificate, RFC 2634), que **no porta `policies`**. És l'últim
      recurs abans de muntar la signatura nosaltres.
    - **`signingCertificateV2=false` tampoc.** Provat: segueix desconeguda.
      S'han acabat els paràmetres d'AutoFirma.
    - **LA SOLUCIÓ: `suport/PdfCms.ps1`.** En lloc de seguir perseguint
      diferències d'una en una, es deixa que **AutoFirma munti el PDF** (que això
      ho fa bé: el document mai surt alterat, el camp de signatura i el caixetí
      són correctes) i **se li reemplaça NOMÉS el CMS de dins del `/Contents`**
      per un fet aquí amb `SignedCms` de .NET, amb l'estructura **exacta**
      d'Adobe: `contentType` + `messageDigest` + `adbe-revocationInfoArchival`
      buit, un sol certificat i **cap atribut ESS**.
      - **Per què és segur**: no es toca ni un byte del document. El
        `/ByteRange` no canvia (i és el que diu on és el forat, així que no cal
        endevinar on és el `<`), el forat que deixa AutoFirma és de **27.000
        bytes** i el nostre CMS n'ocupa ~2.600, i el que es firma són
        **exactament** els mateixos bytes → el *"no ha habido modificaciones"*
        continua sortint igual. La mida del fitxer **no pot canviar**: si
        canviés, el `/ByteRange` (que ja està escrit i firmat) deixaria de
        quadrar.
      - Cal l'**empremta** del certificat (`CertThumb` a les opcions), no només
        el filtre de CN que fa servir AutoFirma: sense triar-ne un al
        desplegable no es pot refer i es queda la signatura d'AutoFirma (i el
        registre ho diu).
      - Si el reempaquetat falla, **es deixa la d'AutoFirma**: val més una
        signatura que només es validi al PC de l'usuari que cap signatura.
      - **`CmsSigner` de .NET, dos paranys**: (1) per defecte fa **SHA-1**, s'ha
        de posar SHA-256 explícitament; (2) si no se li posa **cap** atribut
        signat, firma el contingut directament i el PDF queda **sense
        `messageDigest`**, que és el que Adobe espera. Per això s'hi posa
        l'`adbe-revocationInfoArchival` buit — el mateix que hi posa Adobe — i
        llavors .NET hi afegeix sol el `contentType` i el `messageDigest`.
      - `CryptographicAttributeObject` **no** és a `...Cryptography.Pkcs` sinó a
        `System.Security.Cryptography`. L'assemblatge és `System.Security` al
        PowerShell 5.1 i `System.Security.Cryptography.Pkcs` al 7 (les proves):
        `_CmsCarregaTipus` prova els dos.
      - **Es va escriure una funció que "corregia" l'OID de l'algorisme del
        SignerInfo i la prova la va enxampar corrompent el CERTIFICAT**: el patró
        que buscava (`SEQUENCE{rsaEncryption, NULL}`) també surt a la clau
        pública del certificat, i el del SignerInfo no porta el `NULL`, així que
        la cerca "l'última aparició" queia sobre el certificat. Com que ja
        s'havia demostrat que aquell OID **no canvia la validesa**, la funció es
        va esborrar. Codi que no cal, fora.
    - **4a ronda: l'usuari va reportar «tampoc» amb el reempaquetat ja publicat,
      SENSE registre** — o sigui que no se sap si la signatura refeta es va
      arribar a aplicar (sense certificat triat al desplegable no es pot refer,
      i una excepció al `SignedCms` queia en silenci al respatller). Resposta:
      **tancar l'ambigüitat**, no provar més coses a cegues:
      - **Res no s'escriu sense comprovar-ho** (`_CmsComprova`): el CMS nou es
        descodifica, `CheckSignature($true)` contra el contingut del PDF, 1
        certificat, `messageDigest` present, **cap atribut ESS**. Si falla, el
        PDF es queda com estava i el motiu surt al registre **i al resum**.
      - **L'OID del SignerInfo ara també s'iguala** (`_CmsOidComAdobe`), amb
        guarda POSICIONAL: dins d'un SignerInfo l'algorisme va DESPRÉS dels
        atributs signats, o sigui que l'última aparició de `rsaEncryption` només
        és la bona si queda **després de l'últim `messageDigest`** (la del
        certificat queda sempre abans). El .NET l'escriu `30 0B` sense
        paràmetres i l'Adobe `30 0D` amb NULL: tots dos legals, i el NULL no es
        pot inserir sense re-encodar longituds — es deixa sense.
      - **El resum ho diu ben gros**: «N amb la signatura REFETA I COMPROVADA» o
        «ATENCIÓ: N s'han quedat amb la d'AutoFirma (només vàlida aquí)». I si
        es vol signar **sense** certificat triat, s'avisa **abans** de començar.
      - La suite genera un **certificat efímer en memòria** (`CertificateRequest`)
        i prova el cicle sencer: crear → igualar OID → comprovar → PDF sintètic
        → rellegir i verificar. El camí que corre a Windows és el provat.
    - **TRAMPA de PowerShell (test enxampat)**: un literal hex com a argument
      (`AssertEq $x 0x30`) arriba com a 48 però el PSObject **conserva el text
      del token**, i el `[string]` de dins d'`AssertEq` en fa `"0x30"` → la
      comparació falla mentre el missatge mostra «48» contra «48». La
      interpolació `"$x"` dóna "48"; el cast `[string]$x` dóna "0x30". Als
      arguments de test, els valors esperats **en decimal**. (Mateixa família
      que el PSObject del `Join-Path`.)
    - **5a ronda — LA FI DE LES DIFERÈNCIES DE FITXER.** L'usuari va enviar el
      PDF nou (resum: «REFETA I COMPROVADA») que als PCs dels companys seguia
      sortint desconegut. Es va analitzar contra el vàlid de l'Adobe:
      - certificat del signant **byte a byte idèntic** (mateix SHA-256);
      - CMS del mateix **exacte** nombre de bytes (2.629) i **esquelet ASN.1
        idèntic node a node (156/156)**: les úniques diferències són el hash del
        document i la firma, que canvien per força;
      - **pyHanko** (validador PAdES independent, amb l'arrel de l'AOC com a
        àncora): els DOS fitxers `INTACT:TRUSTED,UNTOUCHED`, cobertura
        `ENTIRE_FILE`, modificacions `NONE`.
      **Conclusió: ja no queda cap diferència de fitxer.** El que difereix és la
      CONFIANÇA dels ordinadors que miren: l'Adobe, per defecte, **només es refia
      de les seves llistes AATL/EUTL** (documentat per Adobe), no del magatzem de
      Windows. El següent element de diagnòstic és la pestanya **Confianza** del
      certificat a l'Adobe d'un company (diu la FONT de la confiança) amb els dos
      fitxers costat a costat.
    - **`suport/Confiar-certificats-AOC.bat`**: la solució de desplegament — la
      mateixa que documenten el BOE i els ministeris per als seus PDF. Porta
      **incrustats** els dos certificats PÚBLICS de l'AOC (arrel `ROOT-A` +
      `SubCA SECTOR PUBLIC Q (G3) A.1`, verificats per hash contra els del CMS
      real) i els instal·la amb `certutil -user -addstore` (sense administrador);
      després cal marcar la Integració amb Windows a l'Adobe (o «Agregar a
      certificados de confianza» des del propi PDF). ASCII pur (els `.bat` amb
      accents es trenquen amb les codepages).
    - **6a ronda — el requisit real de l'usuari**: la firma ha de sortir vàlida
      **també fora de la feina**, sense que cap receptor instal·li res. L'únic
      artefacte que ho compleix EMPÍRICAMENT és el PDF signat pel **propi
      Adobe**. Per això el diàleg té ara **dos modes de signatura**
      (`SignMode` a l'estat/opcions; per defecte `'adobe'`):
      - **`adobe`**: l'eina obre cada PDF a l'Adobe (`_TrobaAdobeExe` /
        `_AdobeExeCandidats`, rutes en text pla — `Join-Path` peta fora de
        Windows), l'usuari signa a mà, i en dir «Sí» es **comprova** que hi hagi
        firma de debò (`_PdfTrobaFirma`); si no n'hi ha (desada amb un altre
        nom!), s'avisa i es compta a part (`$senseFirmaAdobe`). `Cancel·la`
        atura la resta. El certificat/caixetí/AutoFirma del diàleg es
        desactiven en aquest mode.
      - **`autofirma`**: tot el pipeline d'abans (AutoFirma + repack + caixetí),
        intacte i encara provat per la suite. Vàlid on es confiï en l'AOC.
      La qüestió de fons (per què el mateix CMS byte-idèntic es fia en un fitxer
      i en l'altre no en aquells ordinadors) segueix SENSE explicació mesurada:
      la pestanya **Confianza** de l'Adobe d'un company amb els dos fitxers és
      la dada que falta. No barrejar les dues coses: el mode `adobe` és el camí
      pràctic, no la resposta al misteri.
    - **Geometria del diàleg**: `$Script:PdfDlgAmple` (550) i els botons
      col·locats a partir de `$y` (`$yBotons`), amb la `ClientSize` calculada al
      final. Abans anaven clavats a `y=438` i, en afegir-hi els dos radios del
      mode, **van quedar fora de la finestra**; i el botó «Document» de
      `_AddConfigRow` (que acaba a **x=536**) ja sortia tallat amb els 510 px
      d'amplada. Hi ha prova de FONT que ho vigila — acotada al tros del diàleg
      d'opcions, perquè la finestra de **progrés** sí que és de mida fixa.
    - **Pendent**: el que faria la firma validable a qualsevol banda i d'aquí a
      anys és un **segell de temps (TSA)** + dades de revocació (PAdES-LTV).
      AutoFirma ho admet (`tsaURL`, `tsaPolicy`, `tsaHashAlgorithm`…), però cal
      una TSA a què l'Ajuntament doni accés.
  - **L'ASPECTE del caixetí està a `$Script:CaixetiAspecte`** i ara val
    `'defecte'`: només se li diu **on** va (les coordenades de sempre) i el
    dibuixa AutoFirma, que és el mateix que surt amb l'eina *"Utilizar un
    certificado"* de l'Adobe. Amb `'propi'` torna el caixetí nostre (lletra,
    interlineat, contorn i escut de fons): **no s'ha esborrat res**, tot el codi
    hi és i hi ha proves que ho comproven. Es va canviar a petició de l'usuari
    mentre es perseguia el problema de validesa; **que quedi dit: l'aspecte no
    pot canviar la validesa** — el que es valida és el certificat i la seva
    cadena, no el dibuix.
  - **Reintents escalats**: caixetí **(imatge)** → caixetí **(text d'una línia)**
    → **sense** caixetí. Així mai es queda un PDF sense signar i del registre se'n
    dedueix quin ha funcionat. Al registre la imatge surt **resumida** (mida en
    caràcters), mai el base64 sencer, que faria el fitxer inservible.
  - **El registre distingeix els salts**: `_AutoFirmaArgvToText` mostra els salts
    de línia REALS com a `<LF>` i deixa els `\n` LITERALS tal qual. Abans tots
    dos sortien com a `\n` i el log no permetia saber quin era quin — que és
    justament el que calia per depurar.
  - **Posició: ALINEADA amb la capçalera de l'informe, i els números són
    MESURATS**, no posats a ull. Es descomprimeix el flux de contingut de la
    pàgina 1 d'un informe ja generat i se'n treu que: la imatge del logo es
    col·loca amb el seu dalt a `y=806,52` **però porta 18 px de blanc a dalt**
    (de 199) = 6,51 pt, o sigui que la punta de l'escut és a **`y=800`**; i el
    requadre de la «Nota:» va de `x=85,58` a **`x=552,45`**, que és el marge dret
    del text. D'aquí `AutoFirmaCaixetiPos` = 352/752/552/800. Hi ha proves que
    lliguen el caixetí a aquestes dues referències.
  - **Registre de diagnòstic**: cada execució desa a `pdf-signar-log.txt` (al costat
    de `pdf-signar-state.json`) **l'ordre exacta** passada a AutoFirma, el codi de
    sortida i la seva sortida (`_PdfSignarLog`, `_AutoFirmaArgvToText`). S'obre
    (o no) segons la casella **«Obrir el registre de la signatura en acabar»** del
    diàleg d'opcions, que es recorda a `pdf-signar-state.json` (`obrirRegistre`,
    per defecte **no**). Abans es preguntava en acabar cada vegada, i preguntar
    sempre una cosa que és de diagnòstic fa nosa. Serveix per no haver d'endevinar si el caixetí no surt.
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
  Referència / Registre. Al catàleg són ara **`[CAMP: nom]`**, i la pantalla del
  bloc ABANS té una columna **«Omplir…»** que obre `Select-LlicDadesPunt` amb
  els camps que demani aquell text (`_LlicCampsDelText`, pura).
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
- **`Coordenades.ps1` PORTA BOM I L'HA DE PORTAR.** Tot el text que l'usuari veu
  al mapa (llegenda, popups, capçaleres de l'Excel que es baixa) viu dins del
  here-string de `Build-CoordenadesHtml`, en català i amb accents. Sense BOM, el
  Windows PowerShell 5.1 llegeix el fitxer com a ANSI i el mapa surt ple de
  `Ã§`. `Geocodificador.ps1`, en canvi, és ASCII pur i no en porta (com
  `Precintades.ps1`).
- **Dins del here-string `@"…"@` no hi pot haver cap `$` ni cap `` ` `` que no
  sigui una interpolació volguda**: el JavaScript del mapa està escrit
  expressament sense `$` ni template literals. Si hi afegeixes codi, comprova-ho
  (`$` dins del here-string = variable de PowerShell).
- **L'`.xlsx` el genera el NAVEGADOR, sense cap biblioteca.** Un `.xlsx` és un ZIP
  amb cinc XML a dins; amb el mètode «sense compressió» només cal el CRC-32 i les
  capçaleres del ZIP (`crc32`/`zipStore`/`buildXlsx` al mateix HTML). Els textos
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
- **`estat[]` va per POSICIÓ dins d'`ITEMS`, no per ID**: si algun dia la base
  portés dos cops el mateix ID Activitat, dues fitxes es trepitjarien. Al
  `localStorage`, en canvi, es desa **per ID**, que és el que ha de sobreviure
  quan es torni a generar el mapa.
- **Es va provar el mapa SENCER en un navegador de debò** (Chromium + Playwright,
  amb un doble de Leaflet perquè el CDN estava bloquejat): estat inicial,
  arrossegament, desat i recuperació al navegador, esborrat de correccions,
  filtre, i el `.xlsx` baixat rellegit amb `openpyxl`. El **SRI** dels `<script>`
  de Leaflet bloqueja qualsevol doble: a la còpia de prova s'ha de treure
  l'`integrity` (mai al fitxer de veritat).
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
