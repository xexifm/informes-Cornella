# Informes Cornellà — Generador

Script PowerShell que munta informes de deficiències per a llicències
d'activitat de l'Ajuntament de Cornellà.

---

## 1. Què tens a la carpeta

```
informes-Cornella/
├── GenerarInforme.bat            ← doble clic per generar un informe
├── Ruta.bat                      ← doble clic per planificar una ruta d'inspecció
├── Vigilant.bat                  ← genera sol els informes que arriben del mòbil
├── Actualitzar.bat               ← doble clic per actualitzar el programa
├── ESTRUCTURA.md                  ← MAPA: quins fitxers fa servir cada executable
├── ESTRUCTURALS/                  ← plantilles del programa (les pots editar al Word)
│   ├── 0 CAPCALERA.docx           ← capçalera fixa de l'informe
│   ├── 0 CONCLUSIONS.docx         ← conclusions triables (per tipus) + fixes
│   ├── REQ1.docx                  ← catàleg de deficiències
│   └── TERMINI.docx               ← informe de cos fix (sense deficiències a triar)
├── BASE DE DADES ACTIVITATS/      ← (fallback local) copia local de l'Excel quan no hi ha xarxa
├── Informes generats/             ← (local, gitignored) on cauen els .docx generats
├── Rutes generades/               ← (local, gitignored) on cauen els mapes de ruta HTML
└── suport/                        ← codi del programa, no cal tocar-lo
    ├── GenerarInforme.ps1         ← MOTOR + entrada de GenerarInforme.bat
    ├── Format.ps1 · Seguiment.ps1 · DriveApi.ps1   ← mòduls del motor
    ├── config.ps1                 ← (opcional) sobreescriu rutes locals
    ├── Instalar.bat               ← instal·lar en una màquina nova (vegeu cap. 2)
    ├── rutes/Ruta.ps1             ← planificador de rutes (entrada de Ruta.bat)
    ├── mobil/                     ← integració mòbil/Drive
    │   ├── Vigilant.ps1           ← entrada de Vigilant.bat
    │   ├── ExportaDades.ps1       ← el crida Actualitzar.bat (exporta dades)
    │   └── Authorize-Drive.ps1    ← autoritza el PC a Google Drive
    ├── tests/                     ← proves (run-tests.ps1, run-tests-ruta.ps1)
    ├── documentacio/              ← guies (PLA-MOBIL.md, DESPLEGAMENT-MOBIL.md)
    ├── README.md                  ← aquest document
    └── CLAUDE.md                  ← notes per a sessions amb Claude
```

> Per saber **quins fitxers pertanyen a cada executable**, mira
> **`ESTRUCTURA.md`** a l'arrel: hi ha una taula amb el mapa complet.

**Pots moure la carpeta `informes-Cornella` on vulguis dins del PC**:
tot és relatiu. L'única ruta absoluta del programa és la **base de dades
d'activitats** a la xarxa de la feina
(`I:\Activitats_Ordenances\Activitats\5.- Sergi Fadurdo\2_Controls Excels`).
Si executes el programa fora de la feina i la xarxa no és accessible,
mira el [capítol 7: ús fora de la feina](#7-ús-fora-de-la-feina-fallback-local).

---

## 2. Instal·lació en una màquina nova

Per posar el programa en marxa en un ordinador **nou i net** (per exemple, el
de casa) hi ha el fitxer **`suport/Instalar.bat`**. Fa tot el necessari de
manera automàtica.

### Què fa
1. **Instal·la Git** si no hi és (primer prova `winget`; si no, descarrega
   l'instal·lador oficial de Git per a Windows i l'instal·la en silenci).
2. **Baixa el programa** de GitHub a la branca `main`.
3. Et deixa a punt per fer servir `GenerarInforme.bat` i `Actualitzar.bat`.

### Com fer-ho servir — dues maneres
**(a) Des del ZIP de GitHub**
1. A GitHub, botó verd **Code → Download ZIP**.
2. Extreu el ZIP on vulguis.
3. Entra a la carpeta `suport/` i fes doble clic a **`Instalar.bat`**.
   Convertirà la carpeta extreta en un clone operatiu seguint `main`.

**(b) Enviant només el `.bat`**
1. Envia a algú el fitxer **`Instalar.bat`** (és autònom: porta dins la
   direcció de GitHub).
2. Que el posi en una carpeta buida i hi faci doble clic.
   Es descarregarà tot el programa a una carpeta `informes-Cornella`.

### Què NO instal·la (ho has de tenir tu)
- **Microsoft Word i Excel**: el programa els necessita i **no** es poden
  instal·lar automàticament (llicència). El `.bat` avisa si no detecta Word.
- **PowerShell**: no cal; Windows ja porta Windows PowerShell 5.1, i el
  programa s'executa amb `-ExecutionPolicy Bypass` (no cal tocar polítiques).

### Notes
- A casa, **sense haver iniciat sessió a GitHub**, pots **baixar i
  actualitzar** el programa (el repositori és públic), però **no pujar**
  plantilles. Si edites plantilles a `ESTRUCTURALS/` i executes
  `Actualitzar.bat`, l'intent de pujar fallarà sense trencar res (els canvis
  es queden al teu PC). Per poder pujar des de casa caldria configurar el
  login de GitHub (fora de l'abast de l'instal·lador).
- Si l'empresa imposa una **política d'execució de PowerShell per GPO**, el
  `-ExecutionPolicy Bypass` s'ignora. És poc habitual en equips de casa i no
  es pot sobreescriure des d'un `.bat`.

---

## 3. Ús diari

### Generar un informe
**Doble clic a `GenerarInforme.bat`**. Et guia per 5 passos:

| Pas | Què fa |
|-----|--------|
| 1 | Triar catàleg (si n'hi ha més d'un a `ESTRUCTURALS/`) |
| 2 | Dades de la capçalera (auto-fill per ID GIA des de l'Excel) |
| 3 | Marcar les deficiències (TreeView amb filtre). **En marcar-ne una, el seu text surt al panell de la dreta i, si té opcions o camps lliures, els omples allà mateix.** |
| 4 | Triar les conclusions: es mostra el **cos sencer** de cada una i, si té opcions/camps, els omples **dins del propi text**. |
| 5 | Es genera el .docx i s'obre amb Word |

> Les opcions (desplegables) i els camps lliures s'omplen **inline**, allà on
> apareixen al text (Pas 3 i Pas 4); ja no hi ha un pas separat de "camps".

Els botons **Enrere** dels passos 2-4 permeten tornar enrere conservant les
dades. Hi ha també un botó **Recuperar dades últim informe** al pas 2 per
clonar l'informe anterior i tirar pas a pas.

> En obrir `GenerarInforme.bat` surt primer una pantalla per triar entre
> **Generar informe nou** (els passos d'aquí) i **Fer seguiment d'un
> informe existent** (vegeu a sota).

### Fer el seguiment d'un informe

Serveix per indicar, sobre un informe ja emès, si l'activitat ha **resolt o
no** cada requeriment. És **iteratiu**: si una activitat entrega diverses
vegades, sota cada punt s'hi acumulen línies datades:

```
1. Baixa tensió. S'ha d'entregar la legalització.
01/06/2026: No s'entrega.
03/06/2026: S'entrega però falten dades.
05/06/2026: S'entrega correctament.
```

A la pantalla inicial tria **Fer seguiment d'un informe existent**. Després:

| Pas | Què fa |
|-----|--------|
| 1 | Tries l'informe anterior (`.docx`). Pot ser fet amb el programa o a mà, sempre que tingui requeriments enumerats (`1.`, `2.`…). |
| 2 | Tries el **primer paràgraf de conclusions a esborrar** (et llista els paràgrafs a partir de l'últim punt; el detectat ve preseleccionat). |
| 3 | Escrius la **data** d'aquesta entrega (per defecte, avui). |
| 4 | Per cada requeriment: **comentari** (per defecte *"No s'aporta."*, o *"S'aporta."* si marques **Resolt**; editable) i casella **Resolt**. Es veu l'historial. |
| 5 | Tries les **conclusions** noves (com a l'informe normal) i omples els camps que calguin. |
| 6 | Es genera un `.docx` nou i s'obre amb Word. |

**Què fa exactament:**
- **No toca** la capçalera ni el text dels requeriments de l'informe anterior.
- **Esborra** les conclusions antigues (el que va després de l'últim punt
  enumerat, normalment des de *"Vist l'anterior…"* fins a *"Cornellà de
  Llobregat,"*) i hi posa les noves que triïs.
- Sota cada requeriment, **a sota de tot** (després del cos i de l'enllaç),
  afegeix la línia `data: comentari`. Si no es marca *Resolt*, el comentari per
  defecte és *"No s'aporta."*; si es marca, *"S'aporta."* (sempre editable).
- **Negreta dinàmica**: mentre un requeriment NO estigui resolt, **totes** les
  seves línies (punt + sub-línies + anotacions) van en **negreta** —menys
  l'enllaç, que mai—; quan el marques com a resolt, deixen d'anar-hi.
- El text afegit (anotacions i conclusions) surt en **Bookman Old Style 11**
  justificat.

El resultat torna a ser un "informe anterior" vàlid: pots tornar-hi a passar el
seguiment a la ronda següent i s'hi afegirà una línia nova sense duplicar res.

**Nom del fitxer:** el seguiment **incrementa el número** del catàleg amb la data
d'avui: d'un `..._Req1_GIA <id>` en surt `<avui>_Req2_GIA <id>.docx`; el següent,
`Req3`, etc. Si en generes dos el mateix dia, el segon porta `_2`.

**Tècnic:** el seguiment edita el `.docx` directament (XML intern, sense Word);
Word només s'obre al final per ensenyar-te el resultat. Les frases que marquen
on comencen les conclusions a esborrar es poden personalitzar a
`suport/config.ps1` (`$SeguimentConclusionPhrases`).

### Planificar una ruta d'inspecció
**Doble clic a `Ruta.bat`**. Serveix per visitar diverses activitats en un
sol viatge amb el camí més curt. Fa el següent:

1. Surt una finestra: **escriu o enganxa els ID Activitat** a visitar
   (separats per espais, comes o salts de línia). Exemple: `1429 1428 1427`.
2. El programa busca cada ID a la base de dades (fulla **"Estès"**) i agafa:
   - les coordenades **"UTM X"** i **"UTM Y"** per situar-les al mapa, i
   - l'adreça de l'emplaçament (**"Emp. Tipus via" + "Emp. Carrer" +
     "Emp. Número" + "Emp. Lletra"**) per identificar-les.
3. Calcula la **ruta circular més ràpida** que les visita totes i torna al
   punt de partida. La ruta **comença sempre per l'activitat més propera a la
   base** (per defecte **Carrer de l'Energia, 97**) i hi torna al final.
   Per defecte fa servir el servei de rutes per carretera **OSRM**; si no hi
   ha internet, fa una **ruta aproximada en línia recta**.
4. Obre un **mapa** (al navegador) amb cada parada **numerada per ordre de
   visita** (la 1 en verd és l'inici), un panell lateral amb la llista
   ordenada (núm. · ID · adreça) i la distància/temps totals.
5. Botó **"Imprimir / Desar com a PDF"**: imprimeix amb *Microsoft Print to
   PDF* (o *Desar com a PDF*) per tenir la ruta en PDF.

Els mapes es desen a `Rutes generades/` (a l'arrel del clone, ignorada per
git). Detalls:
- **Privacitat**: al servei de rutes només s'hi envien **coordenades**, mai
  noms ni adreces. Si vols anar sense internet o amb un servidor propi,
  configura `$OsrmBaseUrl` a `suport/config.ps1` (vegeu cap. 10).
- Si algun ID no existeix o no té coordenades UTM, el programa t'avisa i
  continua amb la resta (sempre que en quedin per situar).
- És un programa **independent**: només necessita la base de dades d'Excel
  (xarxa de la feina o còpia local), com el generador.

### Actualitzar el programa
**Doble clic a `Actualitzar.bat`**. Fa el següent:

1. Detecta si tens Word obert en alguna plantilla (`~$*.docx`): si sí, t'avisa i s'atura.
2. Si has editat plantilles a `ESTRUCTURALS/*.docx`, les **commiteja i puja a GitHub**.
3. Si has tocat codi (`.ps1`, `.bat`), el guarda al **stash** com a còpia de seguretat.
4. Baixa l'última versió de la branca `main`.
5. Et mostra el commit final per saber on ets.

---

## 4. Sortida

Els informes es desen per defecte a `Informes generats/` (a l'arrel del
clone). Aquesta carpeta està **ignorada per git**: viu només al teu PC.

Nom dels fitxers:
```
YYYY-MM-DD_<Cataleg>_GIA <id>.docx
```
Exemple: `2026-05-29_Req1_GIA 1379.docx`. Si en generes diversos del
mateix dia/GIA, el segon és `..._2.docx`, el tercer `..._3.docx`, etc.

Per canviar la ruta de sortida a una unitat de xarxa, edita `suport/config.ps1`:
```powershell
$OutputDir = 'I:\Activitats_Ordenances\Activitats\5.- Sergi Fadurdo\0_Plantilles\Informes generats'
```

---

## 5. Editar el catàleg de deficiències (`REQ1.docx`)

El programa llegeix `REQ1.docx` **paràgraf a paràgraf** segons l'**estil
del paràgraf** (no segons aparença visual). Selecciona el paràgraf i a
la galeria d'estils del Word tria un dels següents:

| Estil de Word | Què representa | Apareix com |
|---|---|---|
| **Título 1** | Secció | Arrel del TreeView (Pas 3) |
| **Título 2** | Ítem (nom curt) | Checkbox del TreeView |
| **Normal** | Cos de l'ítem (el text que surt a l'informe) | Sota l'ítem |
| **Cita** | Enllaç URL | Hipervincle de 10 pt en paràgraf propi |

> El programa reconeix les variants `Heading 1/2`, `Titulo 1/2`,
> `Título 1/2`, `Ttulo1/2`, `Titol 1/2`... independentment d'idioma o
> de si tenen accent/espai.

### Prefixos especials al **Título 2**

Escriu-los al **principi** del text:

| Prefix | Què fa |
|---|---|
| `::SUB:: Nom` | **Subsecció** dins de la secció (visualment subratllada). |
| `::CHILD:: Nom` | **Sub-ítem** (fill) de l'ítem anterior. Al Pas 3 es pot marcar a part; surt indentat sense numeració. |
| `::INTRO::` | **Bloc d'introducció**. No surt al TreeView; el seu cos s'emet abans del primer ítem de la secció. |

### Enllaços
Posa el paràgraf de l'URL en **estil "Cita"**. Surt sempre en paràgraf
propi, com a hipervincle, en cos 10 pt. Si oblides l'estil i poses
`https://...` en Normal, també funciona (retrocompatible).

### Camps a omplir: `[CAMP: ...]` i `[OPCIO: ...]`

> Els `[OPCIO:]`/`[CAMP:]` s'omplen **inline**, dins del propi text, a la
> mateixa pantalla on marques la deficiència (Pas 3) o la conclusió (Pas 4).

**Camp de text lliure**:
```
Cal aportar el certificat de la [CAMP: entitat acreditada].
```
On aparegui el text, hi surt una caixa de text per escriure "entitat acreditada".

**Camp amb ajuda**:
```
[CAMP: Òrgan homologació PAU (Protecció civil de Catalunya / Protecció civil local)]
```
El text dins els parèntesis surt com a hint a sota del camp.

**Desplegable**:
```
S'ha de presentar un projecte tècnic [OPCIO: Destinatari | a l'ajuntament | a l'ACA] amb el contingut...
```
Allà mateix surt un desplegable (inline) amb les opcions. La primera està
preseleccionada per defecte. L'opció triada substitueix el `[OPCIO: ...]`
al document final.

> El separador d'opcions és `|` (no `/`), perquè el text legal en fa
> servir molt de barres.

**Detalls**:
- Si el mateix nom de camp apareix a diversos ítems, només es demana
  un cop i el valor es replica.
- Pots barrejar `[CAMP:]` i `[OPCIO:]` al mateix paràgraf.

### Negreta i cursiva inline

A qualsevol paràgraf del REQ1 o CONCLUSIONS:

| Sintaxi | Resultat |
|---|---|
| `**text en negreta**` | **text en negreta** |
| `//text en cursiva//` | *text en cursiva* |

### Exemple complet (REQ1)

```
Instal·lacions                       ← Título 1   (secció)
::SUB:: Legalitzacions                ← Título 2   (subsecció)
Instal·lació de baixa tensió          ← Título 2   (ítem)
Cal legalitzar la [CAMP: instal·lació]. Termini: [OPCIO: Termini | 1 mes | 3 mesos].
https://canalempresa.gencat.cat/...   ← Cita       (URL)
::CHILD:: Documentació tècnica         ← Título 2   (fill)
Memòria tècnica signada per **tècnic competent**.
```

---

## 6. Editar les conclusions (`0 CONCLUSIONS.docx`)

Les conclusions **depenen del tipus d'informe**. El fitxer s'organitza en
**grups**: un `Título 1` (Heading 1) per cada tipus d'informe (el nom ha de
coincidir amb el del catàleg: `REQ1`, `TERMINI`...), i a sota, cada conclusió
triable d'aquell tipus com a `Título 2` (Heading 2) + un `Normal` amb el cos.

Estructura del fitxer:

```
CONCLUSIONS                                 ← Normal CENTRAT + NEGREITA  (títol del bloc)

REQ1                                         ← Título 1  (TIPUS D'INFORME = grup)
Terrassa projecte                            ← Título 2  (títol curt; surt al Pas 4)
La terrassa que apareix... no forma part... ← Normal     (cos que s'imprimeix si la tries)

Requeriment                                  ← Título 2
Vist l'anterior, cal requerir l'esmena...   ← Normal

TERMINI                                      ← Título 1  (un altre TIPUS D'INFORME)
Ampliar                                      ← Título 2
Vist l'anterior... es valora ampliar...     ← Normal

No ampliar                                   ← Título 2
Vist l'anterior... es valora NO ampliar...  ← Normal

::SEMPRE:: Ho poso al seu coneixement...    ← Normal amb prefix ::SEMPRE::
::SEMPRE:: Cornellà de Llobregat,           ← Normal amb prefix ::SEMPRE::
```

### Regles
- **Títol del bloc** (paràgraf 1, centrat-negreta): es copia tal qual.
  Si vols canviar el text "CONCLUSIONS" per un altre, edita aquest paràgraf.
- **Grup = tipus d'informe**: cada **Título 1** obre el grup de conclusions
  d'un tipus d'informe. El text ha de coincidir amb el nom del catàleg
  (`REQ1.docx` → `REQ1`, `TERMINI.docx` → `TERMINI`). Al Pas 4 només surten
  les conclusions del tipus que estiguis fent.
- **Conclusions triables**: dins d'un grup, cada una és un **Título 2** (títol
  curt, llegible al Pas 4) seguit d'un **Normal** amb el cos.
- **Parts fixes**: paràgrafs Normal que comencin amb `::SEMPRE:: `.
  No surten al Pas 4 i són **globals** (s'imprimeixen sempre, per a qualsevol
  tipus d'informe), en l'ordre del fitxer, al final de l'informe.
- **Camps i opcions**: pots fer servir `[CAMP: ...]` i `[OPCIO: ...]`
  als cossos. S'omplen inline, dins del propi text, al Pas 4.
- **Negreta i cursiva** inline (`**...**`, `//...//`) també funcionen.

### Separació visual
Cada conclusió té una separació de 12 pt sota. Si vols més o menys,
edita `suport/Format.ps1`:
```powershell
ConclusionSpaceAfterPt = 12     # prova 18 si vols més espai
```

---

## 7. Ús fora de la feina (fallback local)

Quan executes el programa **fora de la xarxa de la feina** (la unitat
`I:` no és accessible), el programa fa servir la carpeta
**`BASE DE DADES ACTIVITATS/`** dins de `informes-Cornella` com a font
alternativa de la base de dades.

### Com fer-ho servir
1. A l'oficina, copia el fitxer `YYYY-MM-DD ACTIVITATS.xls` o `.xlsx`
   més recent dins de `informes-Cornella/BASE DE DADES ACTIVITATS/`.
2. Al PC de fora, executa `GenerarInforme.bat` normalment.

### Ordre de cerca
El programa busca la base de dades en aquest ordre:
1. **Xarxa de la feina** (`$ActivitatsDir`, ruta `I:\...`).
2. **Fallback local** (`BASE DE DADES ACTIVITATS/` del clone).

Si el primer és accessible, sempre s'usa la xarxa. Si no, cau al local.

### Com sé quina s'està fent servir
Al **Pas 2** ho indica clarament a l'etiqueta superior:
- **Normal (blau)**: `Base de dades d'activitats: 2026-05-29 ACTIVITATS.xlsx ...`
- **Fallback (taronja, negreta)**: `[FALLBACK LOCAL]  Base de dades d'activitats: ...`

A més, en obrir el Pas 2 amb fallback surt un avís emergent recordant
de comprovar que la còpia local sigui prou recent.

### Notes
- Els fitxers `.xls` / `.xlsx` dins de `BASE DE DADES ACTIVITATS/` **no**
  es pugen a GitHub (estan al `.gitignore`). Es queden només al teu PC.
- Si vols actualitzar la còpia local, simplement copia el fitxer més
  recent a la carpeta i esborra els antics (o deixa'ls: el programa
  agafa el de data més recent per nom).
- Si executes el programa **sense xarxa i sense còpia local**, el
  programa s'atura amb un missatge explicant les dues ubicacions.

---

## 8. Capçalera (`0 CAPCALERA.docx`)

Placeholders amb format `<<NOM>>` que el programa substitueix amb els
valors del Pas 2:

| Placeholder | Valor |
|---|---|
| `<<ID_GIA>>` | ID de l'activitat |
| `<<EXP_NUM>>` | Núm. d'expedient (auto des de l'Excel) |
| `<<TITULAR>>` | Titular (auto) |
| `<<ADRECA>>` | Adreça (auto) |
| `<<ACTIVITAT>>` | Activitat principal (auto) |
| `<<NUM_ANOTACIO>>` | Núm. registre entrada (auto) |
| `<<DATA_ANOTACIO>>` | Data registre entrada (auto, només data) |

Els camps "auto" es llegeixen de la base de dades d'activitats
(`YYYY-MM-DD ACTIVITATS.xlsx`, fulla "Estès") per ID GIA. Pots editar-los
manualment després de la cerca.

---

## 9. Afegir un catàleg nou

Posa un altre `.docx` a `ESTRUCTURALS/` (per ex. `REQ2.docx`) amb la
mateixa estructura del REQ1. **No pot començar per `0 `** (aquests són
plantilles fixes: capçalera i conclusions). Si hi ha més d'un catàleg,
el Pas 1 et deixarà triar.

### Informes de **cos fix** (p. ex. `TERMINI.docx`)

Si vols un tipus d'informe en què **el cos sempre és el mateix** (no es
trien deficiències), crea el `.docx` a `ESTRUCTURALS/` **sense cap
`Título 1/2`**: només paràgrafs `Normal` (i, si cal, `Cita` per a enllaços).
El programa ho detecta automàticament (un catàleg sense seccions) i:

- **salta el Pas 3** (no hi ha deficiencies a triar);
- imprimeix tal qual els paràgrafs del document com a cos de l'informe
  (amb `**negreta**`, `//cursiva//`, `[CAMP:]` i `[OPCIO:]` igual que sempre).

Les conclusions d'aquest informe surten del grup de `0 CONCLUSIONS.docx` amb
el mateix nom que el fitxer (`TERMINI.docx` → grup `Título 1` **TERMINI**),
tal com s'explica a la secció 6.

---

## 10. Configuració local (`suport/config.ps1`)

Fitxer **opcional** per personalitzar rutes/constants només al teu PC.
Si no existeix s'usen els valors per defecte. Variables disponibles:

```powershell
$OutputDir              = '...'    # ruta on desar els informes (.docx)
$ActivitatsDir          = '...'    # carpeta on viu l'Excel d'activitats
$AlwaysConclusionsCount = 2        # (obsolet, ara s'usa ::SEMPRE::)

# Planificador de rutes (Ruta.bat):
$OsrmBaseUrl            = '...'    # servidor de rutes OSRM (buit = ruta recta)
$RutesOutputDir         = '...'    # ruta on desar els mapes de ruta (HTML)
$RutaOrigenUtmX         = 424456   # base de sortida (UTM X). Per defecte
$RutaOrigenUtmY         = 4578205  #   Carrer Energia 97. La ruta comenca per
                                   #   l'activitat mes propera a aquest punt.
```

---

## 11. Resolució de problemes

### "Word obert: tanca'l"
Tens una plantilla oberta (`~$ESTRUCTURALS\...`). Tanca el Word i torna
a executar.

### "S'ha produït un error al script"
Llegeix el missatge a la finestra negra. Si menciona un fitxer i una
línia, copia-ho a una sessió de Claude per al fix.

### Tot ha canviat a GitHub però el meu programa no es comporta diferent
Comprova en quin commit ets:
```
cd informes-Cornella
git log -1 --oneline
```
Si no és el més recent, executa `Actualitzar.bat`.

### Tinc canvis locals que m'arrosseguen problemes
`Actualitzar.bat` ho gestiona automàticament: els canvis a plantilles
s'envien a GitHub, i els canvis a codi es guarden al stash:
```
git stash list      # veure el què hi ha
git stash pop       # recuperar el més recent
git stash drop      # esborrar-lo
```

---

## 12. Per a desenvolupadors

### Estructura tècnica
- **Mode headless** (`$env:GENINFORME_TEST = '1'`): el script només
  defineix funcions, no obre WinForms ni executa Main. Permet provar
  la lògica a Linux sense Word/Office.
- **Tests**: `pwsh -File suport/tests/run-tests.ps1` (95 OK).
- **Persistència**: l'últim informe queda a
  `%LOCALAPPDATA%\InformesCornella\lastreport.json` per a la funció
  "Recuperar dades últim informe".

### Branca i flux git
- Branca de desplegament: **`main`**.
- Cada sessió de Claude treballa en una branca pròpia `claude/...` i
  acaba fusionant a `main` (mira `CLAUDE.md`).
- El `.bat` d'actualització fa `git pull --ff-only origin main` per
  defecte, amb fallback a rebase automàtic si el clone i el remot
  divergeixen.

### Commit per `Actualitzar.bat`
Si l'usuari ha editat plantilles localment, el `.bat` les commiteja
amb autoria genèrica i les puja a `main`. **No** puja codi (`.ps1` /
`.bat`): aquest el toca només la sessió de Claude.

---

## 13. Preparar informes des del mòbil

Pots **preparar** un informe des del mòbil (triar deficiències, conclusions i
camps), **enviar els requeriments per correu** a un destinatari, i fer que el
**`.docx` complet es generi sol al PC** perquè estigui a punt quan hi arribis.

- El mòbil obre un **formulari web** (GitHub Pages, carpeta `docs/`).
- El mòbil prepara un **paquet JSON**, el deixa a una carpeta privada de
  **Google Drive**, i el PC el converteix en `.docx` amb **`Vigilant.bat`**
  (mode `GenerarInforme.ps1 -DesDePaquet`).
- Les **dades d'activitats** (noms/adreces) **no surten** mai al GitHub públic:
  van només a Drive privat. Les plantilles (sense dades personals) sí que es
  publiquen, per servir el formulari.
- Tot s'actualitza sol: `Actualitzar.bat` refresca les dades del web quan
  canvies plantilles; generar al PC refresca les activitats a Drive.

Posada en marxa pas a pas: **`suport/documentacio/DESPLEGAMENT-MOBIL.md`**.
