# Notes per a Claude (mantenir entre sessions)

## LLEGEIX AIXO PRIMER: hi ha tres documents mes

Aquest fitxer s'havia fet de **2.750 linies** i ja no es podia llegir sencer, que
es justament el que el feia inutil. Les tres histories mes llargues -i son
llargues perque cada una recull diverses rondes de diagnostic amb l'usuari, amb
hipotesis ja descartades- viuen ara a part:

| Abans de tocar... | Llegeix |
|---|---|
| `suport/PdfSignar.ps1`, `suport/PdfCms.ps1` (Word a PDF, AutoFirma, la validesa de la signatura) | **`suport/documentacio/signatura-pdf.md`** |
| `suport/Llicencia.ps1`, `LlicenciaDb.ps1`, `MnsTraspas.ps1`, `ESTRUCTURALS/LLIC.json` | **`suport/documentacio/llicencia.md`** |
| `suport/rutes/` (rutes, coordenades, el planol public de precintades) | **`suport/documentacio/rutes-i-mapes.md`** |
| Posar el mobil en marxa (Drive, EmailJS, GitHub Pages) | **`suport/documentacio/DESPLEGAMENT-MOBIL.md`** |
| Provar el programa al PC despres d'una tanda de canvis | **`suport/documentacio/provar-al-pc.md`** (porta un prompt per enganxar) |

**No hi son per estalviar espai, hi son perque es llegeixin.** Si toques un
d'aquells fitxers sense llegir el seu document, et trobaras reproduint una cosa
que ja es va provar i descartar: la signatura en porta SIS rondes, i Llicencia,
tres informes que resulta que son el mateix document.

El que queda aqui es el que val per a TOT el projecte: com executar les proves,
les trampes del llenguatge, `local/` i la privadesa, l'arquitectura, i les
lliçons que no son de cap modul concret.

## LES REGLES D'ARQUITECTURA (i com es fan complir)

Surten d'un repàs sencer del projecte i de les coses que s'hi van trobar. **No
són preferències d'estil: cada una ve d'un defecte real que hi havia**, i la
majoria tenen un *guard* a la suite que les vigila. Si escrius codi nou aquí,
segueix-les; si en trobes una que no es compleix, és un defecte, no una
excepció.

### 1. Una cosa, un sol lloc — i si n'hi ha dues, compara-les

La lliçó de tot el repàs: **quan una cosa està escrita dues vegades, les dues
còpies ja han divergit**, i les diferències no són variants volgudes sinó
defectes. Va passar amb tot el que es va mirar:

- Els textos del correu (tres còpies) → **el JSON ja tenia un enllaç que les
  altres dues havien perdut**.
- Set lectors de l'Excel → **tres deixaven un `EXCEL.EXE` orfe** i no
  comprovaven el `$null`.
- Cinc obertures del Word → **tres es deixaven l'`AutomationSecurity`**, que és
  el que evita la Vista protegida a la unitat de xarxa.
- Dues funcions d'HTML del correu → en unificar-les va sortir que **dos URLs a
  la mateixa línia es destrossaven**.

**Per tant: abans d'escriure una funció, `grep` del que vulguis fer.** I si
n'has d'unificar dues, **compara-les línia a línia primer**: la diferència sol
ser el defecte que busques.

### 2. Unifica el que és igual; deixa a la crida el que difereix

L'error contrari també es paga. Una funció compartida amb vuit paràmetres per
cobrir tres casos és pitjor que les tres còpies. La regla que ha funcionat:

- **El que és literalment el mateix** (obrir l'Excel, muntar l'HTML d'un cos,
  pintar una pantalla d'assumpte + cos) va a un sol lloc.
- **El que difereix de debò** es queda a la crida, sovint com a **scriptblock**:
  què vol dir «desar», quines columnes es llegeixen, quin mapa de variables.

I hi ha coses que **NO s'han d'unificar, i queda escrit per què**: les dues
graelles (el genèric sortiria més complicat que les dues pantalles juntes) i les
finestres de progrés (només dues són iguals, i amb geometries diferents). **Una
decisió de no fer-ho també s'escriu**, si no algú la refarà a cegues.

### 3. Un mòdul no pot dependre del seu client

`Write-InformeDocx` (el motor) cridava cinc funcions de `Document.ps1` (un dels
seus clients). Si el genèric necessita alguna cosa del particular, **la cosa és
del genèric i s'ha de moure**, amb els seus helpers privats.

**Moure funcions entre fitxers és barat i segur** —tot va amb dot-source al
mateix àmbit—, i hi ha dues comprovacions que ho demostren i s'han de fer:

1. La **llista de noms de funció** de tot `suport/` idèntica abans i després.
2. Els **19 fitxers d'or** idèntics.

### 4. Cada fitxer, una cosa

`Seguiment.ps1` tenia l'informe de seguiment, les primitives d'OOXML **i el menú
principal del programa**. El senyal d'alarma no va ser la mida: va ser que
**quatre dels sis guards que llegien aquell fitxer eren guards del menú**.

Si per parlar del que fa un fitxer necessites la paraula «i», mira-t'ho.

### 5. El que és de dos processos, en un fitxer que NOMÉS defineix

`rutes/Ruta.ps1` corre en un procés propi i **no pot carregar `UiComuns.ps1`**,
que executa coses en carregar-se (AppUserModelID, icona). Per això
`UiFinestra.ps1`, `Json.ps1`, `Excel.ps1` i `Docx.ps1` **només defineixen
funcions**: és l'única manera que els puguin compartir els dos processos. Si
escrius un fitxer compartit, que no executi res en carregar-se.

### 6. Un guard per cada cosa que es pot desfer sense adonar-se'n

El projecte té l'historial de defectes que **no fallen, empitjoren en silenci**.
Contra això la suite té *guards* de font: cap Excel fora d'`Excel.ps1`, cap Word
fora de `Motor.ps1`, cap format fora de `Format.ps1`, l'estil del correu en un
sol lloc, tots els `.ps1` parsegen i tenen BOM…

**Un guard nou s'ha de VALIDAR INJECTANT EL DEFECTE** i comprovant que passa a
vermell. Un guard que no s'ha vist fallar no se sap si vigila res.

### 7. Mesura-ho, no ho dedueixis

Sobretot amb les trampes del PowerShell (el desenrotllat dels arrays, les
closures, la coma que lliga més que el `+`). En aquest repàs, **raonar va donar
la resposta equivocada tres vegades**: una coma que semblava imprescindible i no
feia res, un `.GetNewClosure()` que semblava necessari i no ho era, i una funció
que semblava morta i la cridava un altre fitxer. Les tres es van resoldre
executant-ho.

**Corol·lari**: una excepció que s'escapa d'un bloc de proves **el mata sencer**
i la suite segueix dient «0 FAIL», perquè només compta el que s'arriba a
executar. Ha passat dues vegades. Els blocs llargs van dins d'un `try` que ho
reporta com a fallada.

### 8. Digues per què, no què

Els comentaris d'aquest projecte expliquen **quin defecte hi havia** i **què es
va provar i descartar**. És el que fa que no es repeteixi. Un comentari que
només tradueix el codi a paraules no serveix de res.

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
La suite passa sencera **en un Linux sense Word ni Excel**:

```
apt-get install -y powershell   # o el tar.gz de github.com/PowerShell/PowerShell
GENINFORME_TEST=1 pwsh -NoProfile -File suport/tests/run-tests-all.ps1
```

**`run-tests-all.ps1`, no `run-tests.ps1`**: hi ha SIS suites (`run-tests`,
`-actextr`, `-golden`, `-ruta`, `-precintades`, `-coordenades`) i `run-tests.ps1`
n'és només una. Executar-la sola deixa fora els fitxers d'or —que són la xarxa
de seguretat del motor de composició— i el planificador de rutes.

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

## El codi que estava CLONAT (i què va destapar desclonar-lo)

Ve de l'auditoria d'arquitectura. La lliçó general: **cada còpia s'havia anat
separant de les altres, i les diferències no eren variants volgudes sinó
defectes**. Desclonar no era estètica — era la manera de trobar-los.

### `suport/Excel.ps1` — un sol lector de la fulla «Estès»

**SET** funcions feien la mateixa seqüència (obrir l'Excel per COM, buscar
«Estès», llegir `UsedRange.Value2`, treure'n files/columnes/capçalera, tancar) i
només es diferenciaven en **què** en treien: `Activitats.ps1`, `Informes.ps1`,
`ControlsPeriodics.ps1`, `SeguimentGia.ps1`, `rutes/Ruta.ps1`,
`rutes/Precintades.ps1`, `rutes/Coordenades.ps1`.

Tres defectes que hi havia amagats, tots trobats **comparant les còpies**:

1. **UN EXCEL.EXE ORFE.** Les tres còpies de `rutes/` feien `$wb.Close($false)`
   i `$excel.Quit()` al `finally` **sense `try`**. Si el `Close` peta —i un
   llibre obert des d'una unitat de xarxa ho pot fer— el `Quit` no s'executa
   mai: queda un Excel corrent, invisible, amb el fitxer agafat. Les de
   `suport/` sí que els embolicaven. Ara **cada pas del tancament va dins del
   seu `try`**.
2. **CAP GUARDA DEL `$null`** a les mateixes tres. `New-Object -ComObject` pot
   tornar `$null` sense llançar, i llavors peta 40 línies més avall amb un
   «mètode sobre NULL».
3. El missatge de «no trobo la fulla» **només deia els noms de les pestanyes a
   UNA de les set**. Ara sempre.

- **`Read-FullaEstesa $fitxer { param($x) … }`**: el cos arriba com a
  **scriptblock** i **sense `.GetNewClosure()`** — mateix patró i mateix motiu
  que `Write-InformeDocx`. El context porta `Data`, `Rows`, `Cols`, `Headers`,
  `Sheet` i `Noms`.
- **Què NO decideix**: un Excel buit no és un error (el cos es crida igual amb 0
  files i cada eina en torna el que li toca), i si l'Excel no arrenca **llança**
  — els dos cridadors que volen un `@{ Ok=$false; Error }` s'ho emboliquen amb
  un `catch`, que és una línia.
- **NOMÉS DEFINEIX FUNCIONS**, i és el que permet carregar-lo des dels **dos
  processos** (`rutes/Ruta.ps1` corre a part i no pot carregar `UiComuns.ps1`).
  Mateix patró que `Json.ps1` i `UiFinestra.ps1`.
- **`SeguimentGia` es queda amb la seva instància**, i està raonat al codi: no
  només llegeix, també **copia la fulla a un llibre NOU** amb la mateixa
  instància i l'exporta a PDF, o sigui que necessita l'Excel obert més enllà del
  cos. Comparteix `_TrobaFullaEstesa`.

**LA REGLA DE CONSUM, i està MESURADA.** El cridador **assigna primer** i després
torna amb coma:

```powershell
function Read-XFromExcel($f) {
    $out = Read-FullaEstesa $f { param($x) …; return ,@($registres) }
    return ,@($out)
}
```

**Mai `return (Read-FullaEstesa …)` directe**: allà hi ha DOS `return` seguits,
el pipeline desenrotlla una capa a cada un, i una llista d'**UN SOL** registre
arriba al cridador com l'objecte **pelat** (sortia un `PSCustomObject` on toca un
`Object[]`). Amb l'assignació pel mig no passa, perquè assignar no desenrotlla.
Dins de `Read-FullaEstesa` **no hi ha cap coma** al `return $resultat`, i també
està comprovat: amb aquesta forma d'ús, posar-n'hi una no canvia res, i una coma
que sembla que fa falta i no en fa només despista.

**Un sol normalitzador**: `_NormalizeText` (45 usos) i `_RutaNormalize` (6) feien
el mateix; l'única diferència real era `ToLower()` contra `ToLowerInvariant()`.
Ara és **`_NormalitzaText` amb `ToLowerInvariant`**: un normalitzador que serveix
per **comparar** no pot dependre de l'idioma del Windows.

### El Word s'obre en UN sol lloc

`New-WordApp` (`Motor.ps1`) ja feia tres coses que calen sempre: la guarda del
`$null` amb un missatge útil, `Visible`/`DisplayAlerts`, i sobretot
**`AutomationSecurity = 1`**, que és el que impedeix que el Word obri en **Vista
protegida** els fitxers d'una unitat de xarxa — i els informes són a
`I:\Activitats_Ordenances\…`, que és exactament el cas per al qual aquella línia
hi és. **Quatre** llocs se la saltaven amb un `New-Object` a pèl i cada un hi
perdia coses; el pitjor, `EnviarCorreu.ps1`, **no tenia cap guarda del `$null`**
i el «mètode sobre NULL» arribava tal qual al quadre «Error llegint l'informe»
que veu l'usuari. Ara tots quatre van amb `New-WordApp -Opcional` i **cada un es
queda el seu missatge**, que és l'únic que difereix de debò.

### El correu: `_CosAHtml` i `_OmpleVariables`

`_RecCosHtml` (Recordatoris) i `_ControlsCpEmailHtml` (Controls periòdics) eren
**idèntiques línia a línia**, el mateix estil inline inclòs; i hi havia **tres**
bucles iguals per substituir les variables `{X}`. Ara l'HTML el fa `_CosAHtml` i
el bucle `_OmpleVariables`; **el mapa de variables el segueix posant cada eina**,
que és l'única cosa que difereix. Hi ha guard perquè l'estil inline no es torni a
copiar.

El defecte dels **dos URLs a la mateixa línia** que això va destapar està explicat
a la secció de *Recordatoris periòdics*.

### `Show-EditorAssumpteCos` — les tres pantalles d'«assumpte + cos»

`Invoke-EmailTextos`, `Invoke-ControlsCpEmailTextos` i `Invoke-RecordatorisTextos`
eren tres pantalles gairebé iguals **amb les mateixes coordenades**, ~100 línies
cadascuna. Ara la pantalla és una, a `UiComuns.ps1`, i **el que difereix de debò
es queda a la crida**: què vol dir **desar** i què vol dir **restaurar**, totes
dues com a scriptblock.

- **Els blocs `-Desa`/`-Restaurar` NO porten `.GetNewClosure()`, i és a posta.**
  Es defineixen a cada eina i s'executen des d'un handler de la pantalla; està
  **mesurat** que així segueixen veient els locals de qui els va escriure. Importa
  perquè el `-Desa` dels textos del mòbil llegeix `$textos['bcc']` i, si li
  arribés buit, **desar s'enduria la llista de CCO en silenci**. Hi ha prova,
  validada substituint el bloc per un de `[scriptblock]::Create` (que no té scope
  de definició) i comprovant que el valor hi arriba **buit**.
- **Els salts de línia, en un sol lloc**: el `TextBox` multilínia de WinForms
  només ensenya els salts com a CRLF i els cossos es desen amb LF. Abans cada
  pantalla ho havia de recordar i **la de Recordatoris no ho feia**.
- Desviacions unificades (totes de la de Recordatoris): l'ajuda anava **sota** el
  quadre gran, la lletra era **Consolas**, i no normalitzava els salts.
- En desar, la de Recordatoris **torna a llegir l'estat** abans d'escriure-hi: la
  finestra ha estat oberta i l'enviament automàtic hi escriu l'historial.

### Les finestres de progrés: per què NO s'han fos

Semblava que n'hi havia **quatre** iguals (barra indeterminada). De prop, només
**dues** ho són:

| On | |
|---|---|
| `Informes.ps1` (Actualitzar base) | hi encaixa |
| `ControlsPeriodics.ps1` | hi encaixa, però amb una altra mida (420×130 contra 560×170) |
| `Informes.ps1` (Copiar informes) | porta **Cancel·lar** i la seva lògica |
| `rutes/Coordenades.ps1` | porta **Cancel·lar** *i* corre al **procés a part**, que no pot carregar `UiComuns.ps1` |

Fondre dues còpies de ~15 línies que ni tan sols tenen la mateixa geometria és el
mateix cas que el constructor de graelles (vegeu més amunt): el genèric sortiria
més complicat que les dues juntes. **Decidit que no**, i queda escrit perquè
ningú no ho refaci a cegues.

### Els guards que ho mantenen

Tots **validats injectant el defecte** i comprovant que passen a vermell:

1. Cap fitxer fora d'`Excel.ps1` (i `SeguimentGia`, raonat) obre l'Excel per COM.
2. Cap fitxer fora de `Motor.ps1` obre el Word per COM, i `New-WordApp` conserva
   l'`AutomationSecurity` i el `-Opcional`.
3. L'estil inline del cos del correu viu en un sol lloc.
4. Les tres eines de text passen per `Show-EditorAssumpteCos`, que està definida
   una sola vegada.
5. `Read-FullaEstesa` **tanca el llibre i surt de l'Excel encara que el cos
   llanci** — la prova va amb un **doble de COM** (`pscustomobject` +
   `Add-Member`, el patró de la prova de `Write-InformeDocx`), que és l'única
   manera de provar-ho en un Linux sense Excel.

**Una cosa que va sortir escrivint aquelles proves i val per a tota la suite:**
en injectar el defecte de l'orfe, l'excepció **s'escapava i matava el bloc
sencer**… i la suite seguia dient «0 FAIL», perquè només compta el que s'ha
arribat a comprovar. Per això aquell bloc va dins d'un `try` que ho reporta com a
fallada. És la mateixa lliçó del `throw` que va matar mitja suite.

## On viu cada cosa: dos fitxers que estaven al lloc equivocat

Del mateix repàs d'arquitectura. Aquí no hi havia còpies —hi havia codi
**compartit vivint dins d'un dels seus clients**, que és el que fa que un
projecte sembli més embolicat del que és.

### El motor de composició depenia del seu client

`Write-InformeDocx` (`MotorInforme.ps1`) és el motor genèric, i `Document.ps1`
és el composador de REQ1: un dels seus clients. Però el motor cridava **cinc
funcions que vivien al client** — `_ResolveOutputDir`, `_GetUniqueOutputPath`,
`_OpenOutputDocument`, `Select-CapcaleraBlock`, `Apply-HeaderReplacements` — i
**cap no és de REQ1**: `_ResolveOutputDir` la fan servir sis fitxers més. Eren
allà només perquè és on es van estrenar. Ara són a `MotorInforme.ps1` amb els
seus helpers privats (`_CapMarcador`, `_BuildOrigenText`…); si aquells s'hi
haguessin quedat, el motor hi seguiria depenent per un altre camí.
`_GetOutputFileName` sí que s'ha quedat a `Document.ps1`: només el fa servir
`Build-Document`.

### `Seguiment.ps1` tenia tres coses que no hi pintaven

Eren 2.002 línies amb l'informe de seguiment, les primitives d'OOXML i **el menú
principal del programa sencer**. Ara:

| Fitxer | Què hi ha |
|---|---|
| **`Docx.ps1`** | Llegir i editar un `.docx` **sense Word** (ZIP + WordprocessingML) |
| **`Menu.ps1`** | `Select-Mode` (Pas 1) i el segell d'última execució de les eines |
| `Seguiment.ps1` | Només l'informe de seguiment |

- **Que el menú estava mal posat ho deia la pròpia suite**: **quatre dels sis**
  guards que llegien `Seguiment.ps1` eren guards del **menú** (el títol acotat
  pel xip, el `$result.Choice` amb closure, el botó de la carpeta i l'ordre dels
  informes). Ara apunten a `Menu.ps1`, i s'ha comprovat que hi segueixen
  vermellejant quan s'hi injecta el defecte.
- **`Docx.ps1` hi porta només el que sap d'OOXML i no sap res del seguiment.**
  S'hi han quedat les que porten «Seguiment» a dins encara que toquin XML, i en
  particular **`_ApplyBodyFontXml` i `_MakeBodyRunXml`, que criden
  `_SeguimentFontName`** — escriuen amb la lletra d'aquell informe, no són
  genèriques—, i `_CollectParaRecordsXml`, que decideix `IsBulletChild`,
  vocabulari del model de seguiment i de ningú més.
- `Informes.ps1` ja feia servir tres d'aquelles primitives: una eina depenia
  d'una altra eina per una cosa que no és de cap de les dues.

**Moure funcions entre fitxers és neutre** —tot va amb dot-source al mateix
àmbit— i les dues comprovacions que ho demostren, totes dues fetes: la **llista
de noms de funció de tot `suport/` idèntica** abans i després (783), i els **19
fitxers d'or idèntics**.

**Segona vegada que passa el mateix, i val la pena tenir-ho present**: en
repuntar els guards, un d'ells va petar (`Substring` amb `-1`) i **va matar la
resta de la suite**, que va acabar dient «0 proves, 0 fallades» en lloc de
fallar. El resum només compta el que s'ha arribat a executar.

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

## Llicència, MNS i Traspàs → un document a part

Tot això (l'assistent, `LLIC.json`, la base de dades de llicències, els tres
informes que resulta que són el mateix, i MNS/Traspàs) és a
**`suport/documentacio/llicencia.md`**.

## La capçalera i les conclusions, editables des del programa
Els fan servir **tots** els informes, no pengen de cap rajola del menú i no hi
havia manera d'arribar-hi. Ara hi ha **dos enllaços al costat de «Què vols
fer?»** de la pantalla principal (acció `editcataleg` amb `Doc` = `0 CAPCALERA`
/ `0 CONCLUSIONS`) i l'editor **diu a quin tipus d'informe s'aplica** cada
secció (`_Ed_AplicaText`, reaprofitant el camp de només lectura que a ACT_EXTR
mostra la `[[KEY]]`).

**`0 CAPCALERA.docx` NO es converteix a JSON**: és l'única plantilla de Word de
veritat (escut, la taula del requadre de la «Nota:», les parades de tabulació,
els marges) i no es pot regenerar. El que es fa és **separar les dues coses**:

```
el .docx mana en el FORMAT   (escut, taula, lletra, tabulacions)
el .json mana en el TEXT     (etiquetes, valors fixos, la nota)
```

- **`suport/CapcaleraJson.ps1`**: `_CapBlocsDelXml` (docx → blocs i línies) i
  `_CapAplicaAlXml` (JSON → docx), totes dues **pures i provades sobre el fitxer
  real**. `Sync-CapcaleraJson` regenera el JSON en obrir l'editor —així no pot
  quedar desincronitzat— i `Apply-CapcaleraJson` el torna a escriure en desar,
  amb **còpia de seguretat** al costat.
- **Tot són edicions de TEXT sobre `word/document.xml`**, mai un serialitzador
  d'XML (això ja va corrompre el fitxer un cop i va deixar el programa
  inservible; vegeu la secció de Llicència). El ZIP es reescriu amb
  `ZipArchiveMode::Update`, que només toca aquella entrada.
- **Com es llegeix una línia** (`_CapLiniaDeRuns`): l'etiqueta són els runs en
  **negreta abans del tabulador**, i el valor, els de després. Sense tabulador,
  l'etiqueta és la negreta del principi. En escriure, **tot el text va al PRIMER
  run del grup i els altres es queden buits**: així el `<w:rPr>` (lletra,
  negreta, mida) no es toca mai.
- **Les línies es localitzen per POSICIÓ** (la clau `pN` del JSON), no pel
  contingut: la capçalera té una estructura fixa i el que s'edita és el que hi
  diu, no on va. Per això l'editor **no deixa afegir ni treure línies**
  (`_Ed_CanAddChild` → `$false` per a `capcalera`).
- **Comprovat sobre el `.docx` real**: aplicar el JSON sense canvis no en mou ni
  un byte; canviar una etiqueta només toca aquella línia; i el document segueix
  amb els seus **19 espais de noms**, el `mc:Ignorable` i **cap prefix `ns0:`**
  inventat.
- **`_VistaEsProtegit` ja protegia `0 CAPCALERA*`**, i és el que impedeix que el
  generador de vistes en Word sobreescrigui la plantilla a partir del JSON nou.

### Els enllaços del menú: dues trampes en una
Els enllaços **Capçalera** / **Conclusions** del costat de «Què vols fer?» van
sortir malament de primera:

1. **Un `LinkLabel` té UNA sola lletra per a tot el text**, i amb la Segoe UI del
   programa el llapis `✏️` sortia com un **quadrat**. Als xips de les rajoles sí
   que n'hi ha perquè allà es dibuixa a part, amb `Segoe UI Emoji`
   (`TextRenderer.DrawText`). Als enllaços: **sense emoji**.
2. **Un scriptblock SENSE `.GetNewClosure()` no veu els LOCALS de la funció que
   el crea**, només l'àmbit de l'script. El bloc es va posar **abans** de
   declarar `$result` i sense closure: `$result.Choice = …` queia sobre `$null`,
   el menú es tancava i **el programa sortia sense fer res** («es tanca el
   programa i no passa res més»). Ara va **després** de `$result` i **amb**
   closure, com tots els altres handlers de `Select-Mode`.
   - Prova que ho vigila: **cada `$result.Choice =` de `Menu.ps1` ha
     d'anar després de la declaració i dins d'un bloc amb `.GetNewClosure()`**.
     Validada injectant el cas.
   - Ull amb la variant contrària: un scriptblock **sense** closure sí que
     resol bé quan la variable és de l'àmbit de l'**script** (`$propagate` de
     `SeleccioItems.ps1`). El que no veu són els **locals d'una funció**.

### Ancorar el programa a la barra de tasques (`suport/AccesDirecte.ps1`)
**Windows no deixa ancorar un `.bat`** (ni un `.vbs`): només accessos directes
que apuntin a un **executable**. Per això el `.lnk` apunta a
`wscript.exe "<clone>\suport\GenerarInforme.vbs"` — que és exactament el que ja
fa `GenerarInforme.bat` — amb `IconLocation` = `suport\cornella.ico`. Es deixa a
l'**escriptori** i al **menú Inici** (des d'allà se cerca i s'ancora).

- **No s'ancora sol i no s'ha d'intentar**: des del Windows 10, el verb *Pin to
  taskbar* ja no és invocable per codi. Els trucs que corren escriuen al
  registre (`Taskband`) i reinicien l'explorer: són fràgils i poden carregar-se
  la barra de tasques de l'usuari. Es deixa el `.lnk` fet i un clic dret.
- **LA DRECERA HA DE PORTAR L'`AppUserModelID`, i el mateix que el procés.** Sense
  això, per a Windows la icona ancorada i la finestra del programa són **dues
  aplicacions diferents**: l'ancorada surt **sense icona** i en obrir-la apareix
  un **segon botó** a la barra de tasques en lloc d'il·luminar-se el que ja hi
  havia. Va passar exactament així.
  - El procés ja se'l posava (`SetCurrentProcessExplicitAppUserModelID`, a
    `UiComuns.ps1`); el que faltava era **posar-l'hi a la drecera**.
  - `WScript.Shell` **no sap** escriure aquesta propietat: cal
    `IShellLink` → `IPersistFile.Load` → QI `IPropertyStore` →
    `SetValue(PKEY_AppUserModel_ID, …)` → `Commit` → `Save`. Les interfícies es
    declaren amb `Add-Type` i es compilen en viu (mateix patró que
    `_PickFolderModern`), en **C# 5** (PowerShell 5.1 no en compila de més nou).
  - `PROPVARIANT` per a una cadena: `vt = 31` (VT_LPWSTR) i el punter **al byte
    8** (2 del tipus + 6 de reservats), tant a 32 com a 64 bits; s'allibera amb
    `PropVariantClear`. Comprovat amb `Marshal.SizeOf`: `PROPERTYKEY` = 20.
  - **L'identificador està escrit en UN SOL LLOC** (`AccesDirecte.ps1`) i
    `UiComuns.ps1` el llegeix d'allà — per això `UiComuns.ps1` carrega
    `AccesDirecte.ps1` si no hi és. Hi ha prova que compta el literal a tot
    `suport/` i falla si n'apareix un segon (validada injectant-lo).
  - **`IconLocation` amb l'índex** (`…\cornella.ico,0`): sense ell hi ha Windows
    que es queden amb la icona genèrica.
  - **Si ja estava ancorat, s'ha de desancorar i tornar a ancorar**: Windows es
    queda la còpia del dia que es va ancorar. Ho diu el `.bat` i el `LLEGEIX-ME`.
- **…I DESPRÉS ES VA PERDRE LA ICONA.** L'AppUserModelID va arreglar l'agrupació
  i va destapar la segona meitat: **`New-Object System.Drawing.Icon($path)` dóna
  una icona BUIDA amb aquest `.ico`**, perquè el de l'Ajuntament porta **les set
  mides comprimides en PNG** i el GDI+ no les sap descomprimir — és exactament la
  mateixa trampa que ja feia sortir l'escut buit al caixetí de la signatura.
  Abans no es notava perquè la barra de tasques agrupava el programa sota el
  PowerShell i hi sortia **la icona d'ell**; en donar-li identificador propi va
  passar a fer servir la de la finestra, que era buida.
  - **`_IcoTriaFrame` ha passat de `PdfSignar.ps1` a `UiComuns.ps1`**: ara la fan
    servir dos mòduls, i allà és on van els helpers compartits. `PdfSignar` la
    segueix cridant igual.
  - **`_IconaDeIco`** (nova, a `UiComuns.ps1`): llegeix la taula del `.ico`,
    agafa el PNG de la mida demanada i en fa una icona de veritat amb
    `Bitmap.GetHicon()` + `Icon.FromHandle`. Com que `FromHandle` **no** es fa
    seva la nansa, se'n fa un `Clone()` gestionat i es crida `DestroyIcon`.
    Respatller a `Icon(path, Size)` si algun dia el `.ico` porta imatges DIB.
  - **L'escut de la drecera es copia al disc LOCAL**
    (`%LOCALAPPDATA%\InformesCornella\cornella.ico`) i l'`IconLocation` hi apunta:
    **el clone de l'usuari viu en una unitat de xarxa** i l'explorador no és de
    fiar carregant icones d'allà per a un element ancorat. Si la còpia falla, es
    fa servir la del clone.
- `Crear-acces-directe.bat` (arrel) i el pas 4 d'`Instalar.bat` el criden.
- **Un `.bat` amb una ordre de PowerShell llarga és un niu d'errors**, i per això
  la feina és al `.ps1` (`Invoke-CrearAccesDirecte`). Concretament: **dins de
  cometes dobles el `cmd` NO interpreta el `|`, però sí que deixa passar el `^`
  literal** — o sigui que un `^|` escrit «per si de cas» arriba tal qual al
  PowerShell i peta. Hi ha prova de font que no hi hagi cap `^` dins de cometes
  al `.bat`, i que sigui ASCII pur.
- `Get-AccesDirecteObjectiu` / `Get-AccesDirecteDestins` són **pures** i es
  proven a Linux (destí, arguments entre cometes, icona, barra final del clone).

### El botó 📁 del menú: la carpeta dels informes
A l'esquerra de ⚙, obre la carpeta on es desen els informes. La ruta surt de
**`_ResolveOutputDir`** —la mateixa que fa servir la generació, o sigui la de
**Configuració** amb el seu respatller local—: **cap ruta escrita al codi**, i
hi ha prova que ho vigila (validada injectant una ruta `I:\…`). **No tanca el
menú**: obrir una carpeta no és triar cap opció. L'emoji de carpeta és
**astral** (U+1F4C1) i va amb `ConvertFromUtf32`, mai amb `[char]`.

### `continue` dins d'un `switch` NO continua el `foreach` de fora
Només surt del `switch`; l'execució segueix a la línia de sota, dins de la
mateixa volta del bucle. Va passar al desat de la base de llicències: les
entrades de la documentació del projecte (que no pertanyen a cap bloc de punts)
queien al codi dels blocs i petaven. Solució: `if` + `continue`, no `switch`.
Hi ha una prova que ho deixa escrit **i comprovat** contra el propi PowerShell.

## Que una finestra hi CAPIGA sempre (`suport/UiFinestra.ps1`)

En una pantalla més baixa —el PC de casa, o el Windows amb l'escalat al 125%—
diverses pantalles d'aquest programa surten **més altes que l'àrea de treball**,
i els botons de baix queden fora i no s'hi pot arribar. Els casos: l'editor de
catàlegs (`MinimumSize` 836×**700**), la base de llicències i el pas de
documentació (client 1080×660), les actes extraordinàries (1160×**680**).

Es resol en dos temps, i **calen tots dos**:

1. **`AutoScroll = $true` a totes les finestres.** Si s'encongeixen i algun
   control queda per sota, surt la barra vertical. Es deixa que WinForms calculi
   sol la zona a recórrer (`AutoScrollMinSize` a zero): així una graella
   `Dock='Fill'` **segueix encongint-se** com sempre en lloc d'estrenar una barra
   que no calia.
2. **Només quan la finestra no hi cap**: es retalla a l'àrea de treball i,
   **abans**, se li baixa el `MinimumSize` —sense això el Windows es nega a
   encongir-la i el pas 1 no serveix de res—. En aquest cas sí que s'hi fixa
   `AutoScrollMinSize` = **l'alçada de disseny**, que és l'única manera de
   garantir que s'arriba a tot, també al que està ancorat a baix (que si no puja
   i es comprimeix). L'amplada es deixa a 0: els controls ancorats a la dreta ja
   s'estrenyen sols i posar-hi l'amplada trauria una barra horitzontal inútil.
   La finestra també es **corre** perquè quedi sencera dins de l'àrea: una
   centrada que sobresurt per baix també sobresurt per dalt, i llavors ni la
   barra de títol es pot agafar.

- La decisió és **pura** (`_MidaFinestraDinsPantalla`) i es prova a Linux;
  `_AjustaFinestraAPantalla` només l'aplica. Va al **`Shown`**, no abans: fins
  llavors el `ClientSize` encara pot canviar.
- **Per què un fitxer a part i no `UiComuns.ps1`**: el planificador de rutes
  (`rutes/Ruta.ps1`, i `Coordenades.ps1` que el carrega) corre en un **procés
  propi** i no pot carregar `UiComuns.ps1`, que **sí que executa coses en
  carregar-se** (AppUserModelID, icona). `UiFinestra.ps1` només defineix
  funcions, i per això el poden compartir els dos processos sense arrossegar-ne
  els efectes.
- Les pantalles del programa hi entren totes per `_NewForm`; les cinc finestres
  que es fan a mà (dues a `Ruta.ps1`, dues a `Coordenades.ps1`, una a
  `EnviarCorreu.ps1`) criden `_AjustaFinestraAPantalla` des del seu `Shown`.
  **Hi ha una prova de FONT que compta els `New-Object …Forms.Form` de cada
  fitxer i exigeix el mateix nombre de crides**, validada injectant una finestra
  òrfena i comprovant que passa a vermell.
- L'ordre al `Shown` importa: **primer `_AvisaSolapaments`, després l'ajust**.
  Al revés, encongir la finestra faria que controls ancorats a baix es trepitgin
  i sortirien avisos de solapament que no són cap defecte.

## Controls que es trepitgen: la tolerància
`_TrobaSolapaments` avisava de pantalles que es veuen perfectament (una etiqueta
que passa **un píxel** per sota d'un radio, el títol i el subtítol de la banda
granat) i l'avís es va tornar soroll que ningú llegia — l'usuari: *«Calen tots
aquests avisos de Pantalla mal col·locada??»*. Ara ha de trepitjar com a mínim
**8 px en les DUES direccions** (`$Script:SolapMinPx`) **i** cobrir el **15%**
del control més petit (`$Script:SolapMinPct`). Amb això calla en els casos reals
que no molesten i segueix cridant quan un botó tapa una etiqueta de debò —que és
com es va trobar el defecte de `Select-LlicFase`, on els botons anaven a una `Y`
clavada al codi i ara surten del **peu real** de l'última etiqueta.

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

## El MOTOR DE COMPOSICIÓ (`suport/MotorInforme.ps1`)
Ve de la petició de l'usuari: **«el programa el més senzill, estandarditzat i
modulable possible; no vull duplicitats»** — canviar la lletra en un lloc i que
afecti a tot, canviar un requeriment de REQ1 i que canviï a tot arreu, els salts
de pàgina, les conclusions.

**El repartiment és aquest, i val per a tot el programa:**

```
Format.ps1        COM es veu un paràgraf   (lletra, sagnies, espaiats)
MotorInforme.ps1  COM es munta un document (l'ordre, l'aire, els nivells)
els Build-*Blocs  QUÈ s'escriu             (purs: es proven sense Word)
```

### El que hi havia duplicat (i on ha anat a parar)
| Duplicitat | Còpies | On viu ara |
|---|---|---|
| Obrir/desar el `.docx` (~20 l.) | 4 (`Build-Document`, `Build-ActExtrDocument`, `Build-LlicenciaDocument`, `Build-MnsDocument`) | `Write-InformeDocx` |
| Escriure una línia (text + enllaços) | 3 (`$emitLine`, `_LlicEmetLinia`, `_VLine`) | `Write-Linia` |
| Escriure un punt (número + cos + fills + URLs) | 2 (`_WriteCatalegBody`, `_VistaCataleg`) | `Build-CatalegBlocs` + `Write-Informe` |
| `if ($cfg.SpacerAfterX) { Format-Spacer }` | **34** | `Format-Aire $sel '<clau>'` |
| Salt de pàgina i `OutlineLevel` a pèl | 2 fitxers | `Format-SaltPagina` / `Format-Nivell` |
| Valors i conversions de format | `Seguiment.ps1` en tenia còpia | llegeix `$ReportFormatConfig` + `_CmToTwips`/`_PtToTwips` |

### `Write-InformeDocx`
Nom únic → còpia a `%TEMP%` (si no, el Word obre el fitxer en *Vista protegida*
quan el destí és una unitat de xarxa) → bloc de capçalera → `<<PLACEHOLDERS>>` →
**el cos, que arriba com a scriptblock** → desar, tancar i moure al destí.
- **El scriptblock del cos NO porta `.GetNewClosure()`**, i és a posta: ha de
  veure els locals del `Build-*` **en temps d'execució**, i la closure en
  copiaria els valors del moment de crear-lo (vegeu la secció de les closures).
- **Si el cos peta, el document es tanca** abans de rellançar l'error. Abans
  només ho feia ACT_EXTR; les altres tres deixaven el Word amb un document
  obert i el `%TEMP%` brut.

### `Format-Aire` i les banderes
La clau és el **nom del bloc que s'acaba d'escriure** (`seccio`, `subseccio`,
`item`, `intro`, `introparagraf`, `conclusions`) i `$Script:AireFlagPerClau` la
tradueix a la bandera. **Una clau desconeguda no posa aire i no peta**: un nom
mal escrit no pot afegir una línia en blanc a un informe en silenci.
- **`Test-FormatAire` és pura** i es prova sense Word.
- **`aire` ≠ `espai`**: `aire` depèn d'una bandera; `espai` és una línia en
  blanc **sempre** (el cos fix de TERMINI, l'ANNEX 1).
- Les **vistes** tenen `_VAire`, que passa per `_VSpacer` per tornar el nivell
  d'esquema a cos.

### El vocabulari de blocs (`Write-Informe`)
`seccio`, `subseccio`, `etiqueta`, `item`, `cos`, `pic`, `nota`, `enllac`,
`pla`, `continua`, `llista`, `conclusio`, `conclusiocap`, `aire`, `espai`,
`saltpagina` i **`unitat`**.

**`unitat` és la peça que ho fa funcionar.** És el contenidor d'un punt sencer:
- en obrir-se marca que **el pròxim `pic` és el primer sub-punt** → `-First`
  (12 pt en lloc de 6);
- en tancar-se hi posa l'aire d'ítem **només si ha escrit alguna cosa**, de
  manera que l'espai va després de l'ítem **complet** (amb els seus fills i
  enllaços). Això és el que abans feia el `$itemWritten` a mà a cada família.

**`-AmbNivells`** és l'**única** cosa que diferencia una vista d'un informe:
posa l'`OutlineLevel` a cada paràgraf. Els paràgrafs **buits també el tornen a
cos**: el Word l'hereta, i un espaiador després d'un títol es quedaria a nivell
1 i sortiria com una entrada buida al panell de navegació.

**`-SenseCamps`**: a la vista els `[CAMP:]`/`[OPCIO:]` es veuen **tal qual**. És
una vista del **catàleg**, no l'informe d'una activitat. (Primer intent:
passar-hi un diccionari buit. Els deixava **en blanc** i la vista perdia
justament el que hi vas a mirar; ho va enxampar un fitxer d'or.)

**Un tipus de bloc desconegut PETA.** Val més això que generar un document al
qual li falta un tros sense que ho digui ningú.

### Les proves de font que ho mantenen
Quatre guards, tots **validats injectant el cas** i comprovant que passen a
vermell:
1. Cap fitxer fora de `MotorInforme.ps1` crida `_OpenOutputDocument`.
2. Cap fitxer fora de `Format.ps1` llegeix les banderes `Spacer*`.
3. **Cap fitxer fora de `Format.ps1` toca el Word per format**
   (`ParagraphFormat.`, `.OutlineLevel`, `InsertBreak(`). És la invariant que
   fa que «canviar-ho en un lloc» sigui veritat.
4. Cap alineat escrit a pèl a `Format.ps1` (l'únic literal que hi pot quedar és
   el centrat del títol CONCLUSIONS).

### ELS FITXERS D'OR (`suport/tests/dades/emit-*.txt`)
**La xarxa de seguretat per tocar el motor.** Cada fitxer és la seqüència
sencera de crides `Format-*` d'una família, una per línia, tal com les
enregistra `FormatDoubles.ps1`. Es llegeixen com el document:

```
BODY|En relació a la sol·licitud de Modificació NO Substancial …
AIRE|item
LLISTA|
CONCLCAP|CONCLUSIONS
```

- **Per què**: aquest projecte té l'historial de defectes que **no fallen sinó
  que empitjoren en silenci** — un espai que desapareix, un sub-punt que passa
  de 12 a 6 pt, un enllaç que canvia de lloc. Cap prova puntual els veu tots;
  una comparació línia a línia de tot el document, sí.
- **19 escenaris FIXOS** (els N primers ítems de cada secció; cap dada que
  depengui de la data ni de la màquina): REQ1, TERMINI, ACT_EXTR req i fav, les
  7 vistes, els 3 informes de Llicència, la provisional amb ANNEX 1, i
  MNS/Traspàs amb punts de REQ1 i sense.
- Quan una comparació falla diu **la primera línia que difereix**, amb
  l'esperat i l'obtingut: la resta acostuma a ser el mateix desplaçat una
  posició, i abocar-ho tot amagaria la línia que importa.
- **Per refer-los després d'un canvi VOLGUT:**
  ```
  GENINFORME_GOLDEN=1 pwsh -NoProfile -File suport/tests/run-tests-golden.ps1
  ```
  **…i després mira't el `git diff`: és tota la gràcia.** Si el diff no és
  exactament el que esperaves, el canvi no era el que et pensaves. Va passar
  dues vegades el mateix dia (els `[CAMP:]` que es buidaven, i un paràgraf
  duplicat a la vista de TERMINI que hi era des de sempre).

### Què queda per migrar (i per què no corre pressa)
`MnsTraspas.ps1`, `ActExtr.ps1` i `_LlicEscriuPunt` encara criden les
`Format-*` directament. **No és duplicació**: totes tres ja passen per
`Format-Aire`, `Write-Linia` i, quan escriuen punts de catàleg, per
`_WriteCatalegBody`. El que els queda és lògica **pròpia** (la frase
d'observacions de MNS, el repartiment per token d'ACT_EXTR, i l'ordre dels
enllaços respecte del comentari a Llicència). Si es migren, **una família per
commit i comparant el fitxer d'or a cada pas**.

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
  `Get-ParsedCataleg`, `Read-Conclusions` i `Parse-ActExtrTemplate` llegeixen
  **NOMÉS el `.json`**. El respatller al `.docx` que hi havia aquí **es va
  esborrar** —vegeu «Res de llegir `.docx` per treure'n contingut», més amunt—:
  els `.docx` d'ESTRUCTURALS ja no són fonts sinó **vistes generades**, o sigui
  que el respatller no hauria fallat, hauria generat un informe silenciosament
  equivocat. Si el `.json` no hi és, peta amb un missatge clar.
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
  - **Punt d'entrada**: a la finestra principal (`Menu.ps1`, `Select-Mode`),
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
  `{PROPER_CP}{DATA_CONTROL}{DATA}` + `**negreta**` + `//cursiva//` + enllaços) via
  el botó **"Editar text"** (`Invoke-ControlsCpEmailTextos`, mateix patró que *Textos del correu*),
  però es desa **LOCALMENT** a `%LOCALAPPDATA%\InformesCornella\controls-cp-email.json`
  (fora del repo: cap dada personal, sobreviu a `Actualitzar.bat`). Funcions pures
  testejades: `_DefaultControlsCpEmail`, `_ControlsCpRecipients`, `_FillControlsCpPh`.
  L'HTML del cos el fa **`_CosAHtml`** (`EnviarCorreu.ps1`), la mateixa que els
  recordatoris — i per tant aquest correu **també accepta `//cursiva//`**, cosa
  que la còpia pròpia que hi havia no feia. Outlook (COM) i finestres només a
  Windows. Es dot-sourceja a `GenerarInforme.ps1` després de `ControlsPeriodics.ps1`.

## Rutes, coordenades i el plànol públic → un document a part

L'eina «Coordenades» (Excel contra façana, el Cadastre) i el plànol públic
d'activitats precintades són a **`suport/documentacio/rutes-i-mapes.md`**.

## Els requeriments d'INCENDIS de REQ1 (agost 2026)

**El RSCIEI que citàvem estava DEROGAT.** El RD 2267/2004 el va substituir el
**Real Decreto 164/2025, de 4 de marzo** (BOE 10/04/2025, en vigor el 10/05/2025),
que a més **modifica el RIPCI (RD 513/2017) i el CTE DB SI**. Si algun dia
tornen a sortir dubtes d'incendis, el punt de partida és aquest, no el del 2004.

- **L'error que es requeria de debò**: les pressions de les BIE deien «45 mm
  entre 5,4 i 8,4 bar; 25 mm entre 3,5 i 6,5 bar» — valors del **RD 1942/1993**,
  derogat. El vigent lliga el mínim al cabal i a la K (25 mm K=42 → 4 bar;
  45 mm K=85 → 3,5 bar) amb un **únic màxim de 9 bar**.
- **Novetats estructurals del RD 164/2025** que toquen el que informem: els
  nivells de risc intrínsec se subdivideixen en **8 subnivells** (abans 3); la
  **configuració tipus E desapareix**; l'article 4 obliga les zones subsidiàries
  de més de **250 m²** (100 m² d'aparcament) a ser sector independent sota el
  CTE; i el CTE guanya l'**ús Almacén**, que compleix amb la taula 1.1 del DB SI 1
  **més** els annexos I-IV del RSCIEI (els trasters de lloguer hi entren sempre).
- **Quan aplica el RSCIEI, la inspecció del RIPCI ja va inclosa dins la del
  RSCIEI** (ho diu la Guia Tècnica a l'article 13): no s'han de requerir totes
  dues.

### El lector de PDF del Drive TALLA el text
`mcp__Google_Drive__read_file_content` retorna com a molt ~293.000 caràcters:
del RSCIEI en donava fins a la pàgina **80 de 107**, i les **seccions 4 i 5 de
l'Annex II** (intervenció dels bombers i resistència estructural) i **tot
l'Annex III** quedaven fora. Amb el text truncat vaig arribar a escriure al pla
que un requisit «no es podia confirmar».

**La via bona**: `mcp__Google_Drive__download_file_content` retorna el fitxer en
base64 i, com que passa del límit, **el harness el desa a un fitxer** en lloc de
posar-lo al context. D'allà es descodifica i s'extreu amb `pypdf`:

```
json.load(<fitxer del tool-result>)['content'] -> base64.b64decode -> .pdf -> pypdf
```

`drive.google.com`, `boe.es` i `codigotecnico.org` els **bloqueja el proxy**
d'aquest entorn (403 al CONNECT): no es poden baixar amb `curl`.

### EDITAR ELS `.json` D'ESTRUCTURALS SENSE EMBRUTAR EL `git diff`
Els catàlegs els desa l'editor amb el **`ConvertTo-Json -Depth 40` del Windows
PowerShell 5.1**, que té un format propi: sagnat **alineat a la columna del
valor**, **dos espais** després dels dos punts, i `'`, `>` i `&` escapats com a
`\u0027`, `\u003e` i `\u0026`. Reescriure el fitxer amb el `json` de Python el
canvia **sencer** i el diff deixa de servir per revisar res.

`scratchpad/psjson.py` (de la sessió, no del repositori) imita aquell format, i
**es valida amb l'anada i tornada sobre el fitxer real**: llegir `REQ1.json` i
tornar-lo a serialitzar ha de donar **els mateixos bytes**. Si algun dia s'ha de
tornar a fer una edició massiva d'un catàleg, aquest és el camí — i la
comprovació d'anada i tornada, la condició per fiar-se'n.

**Convenció de format del catàleg, respectada:** el text **citat en castellà**
va en **cursiva**; el text de connexió en català, pla.

### La prova dels textos fixos no pot exigir-los a TOTES les famílies
La prova «els textos fixos de REQ1 a totes les famílies» descobreix **sola** cada
subsecció amb text fix. En afegir `Incendis :: RSCIEI` va passar a vermell… i no
hi havia cap defecte: **Llicència només expandeix les subseccions que li diu
`LLIC.json`**, i aquella no n'és una. Ara a cada família s'hi comproven **només
els grups que aquella família porta de debò** (`$grupsLlicTF`). Validat injectant
el defecte original —que l'intro no s'emeti— i comprovant que torna a donar
6 FAIL a les tres fases de Llicència.

## Un TEXT FIX pot ser de la SECCIÓ, no només de la subsecció

Ve d'una petició senzilla —moure l'intro de l'article 4 de dins de
`Legalitzacions` a la secció `Instal·lacions`, perquè ara parla també de les
inspeccions— i va destapar que **el text hauria desaparegut de tots els
informes**.

- La regla d'abans era: «una subsecció nova invalida l'intro pendent»
  (`Build-CatalegBlocs`). Serveix perquè l'intro de la subsecció A no surti sobre
  els ítems de la B. Però un intro posat a la **secció** queda abans del primer
  marcador de subsecció, i els seus ítems pengen de les subseccions: el primer
  `SUB` se l'enduia i **no sortia mai**.
- Ara es distingeixen els dos casos: **abans de la primera subsecció = de la
  secció** (sobreviu als canvis de subsecció, surt UN sol cop i **abans** del
  títol de la subsecció); **dins d'una subsecció = d'aquella subsecció** (mor amb
  ella, com sempre).
- **La mateixa regla és a TRES llocs** i s'han de tocar els tres:
  `Build-CatalegBlocs` (`MotorInforme.ps1`), `_LlicItemsAmbUbicacio` +
  `_LlicEscriuPunts` (`Llicencia.ps1`) i `_VistaLlicencia` (`VistaWord.ps1`).
  A Llicència l'intro viatja dins del punt (`Intro`), o sigui que cal el camp
  **`IntroDeSeccio`** —i copiar-lo als **tres** llocs que munten el registre del
  punt, que és on em vaig deixar dos i el text sortia **tres vegades**, una per
  subsecció.
- **Ho va enxampar el fitxer d'or i res més.** La suite no baixava; simplement el
  paràgraf no hi era. És el defecte típic d'aquest projecte: no falla, empitjora
  en silenci.

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
- `_RecHistorialAMapa` / `_RecObjAMapa` desfan el que fa `ConvertFrom-Json`
  (PSCustomObjects): l'historial s'indexa per GIA i sense això `.ContainsKey` no
  existiria i **es perdria tot en silenci**. Hi ha prova d'anada i tornada **amb
  el JSON pel mig**, que és on aquest projecte s'ha trencat sempre.
