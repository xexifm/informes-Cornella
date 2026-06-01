# Informes Cornellà — Generador

Script PowerShell que munta informes de deficiències per a llicències
d'activitat de l'Ajuntament de Cornellà.

---

## 1. Què tens a la carpeta

```
informes-Cornella/
├── Actualitzar.bat               ← doble clic per actualitzar el programa
├── GenerarInforme.bat            ← doble clic per generar un informe
├── ESTRUCTURALS/                  ← plantilles del programa (les pots editar al Word)
│   ├── 0 CAPCALERA.docx           ← capçalera fixa de l'informe
│   ├── 0 CONCLUSIONS.docx         ← conclusions triables + fixes
│   └── REQ1.docx                  ← catàleg de deficiències
├── Informes generats/             ← (local, gitignored) on cauen els .docx generats
└── suport/                        ← codi del programa, no cal tocar-lo
    ├── GenerarInforme.ps1
    ├── Format.ps1
    ├── config.ps1                 ← (opcional) sobreescriu rutes locals
    ├── tests/run-tests.ps1        ← proves automàtiques
    ├── README.md                  ← aquest document
    └── CLAUDE.md                  ← notes per a sessions amb Claude
```

**Pots moure la carpeta `informes-Cornella` on vulguis dins del PC**:
tot és relatiu. L'única ruta absoluta del programa és la **base de dades
d'activitats** (`I:\Activitats_Ordenances\Activitats\5.- Sergi Fadurdo\2_Controls Excels`).

---

## 2. Ús diari

### Generar un informe
**Doble clic a `GenerarInforme.bat`**. Et guia per 6 passos:

| Pas | Què fa |
|-----|--------|
| 1 | Triar catàleg (si n'hi ha més d'un a `ESTRUCTURALS/`) |
| 2 | Dades de la capçalera (auto-fill per ID GIA des de l'Excel) |
| 3 | Marcar les deficiències a incloure (TreeView amb filtre) |
| 4 | Triar les conclusions (checkboxes amb el títol curt) |
| 5 | Omplir camps i triar opcions (només els que apareixen als ítems/conclusions seleccionats) |
| 6 | Es genera el .docx i s'obre amb Word |

Els botons **Enrere** dels passos 2-5 permeten tornar enrere conservant les
dades. Hi ha també un botó **Recuperar dades últim informe** al pas 2 per
clonar l'informe anterior i tirar pas a pas.

### Actualitzar el programa
**Doble clic a `Actualitzar.bat`**. Fa el següent:

1. Detecta si tens Word obert en alguna plantilla (`~$*.docx`): si sí, t'avisa i s'atura.
2. Si has editat plantilles a `ESTRUCTURALS/*.docx`, les **commiteja i puja a GitHub**.
3. Si has tocat codi (`.ps1`, `.bat`), el guarda al **stash** com a còpia de seguretat.
4. Baixa l'última versió de la branca `main`.
5. Et mostra el commit final per saber on ets.

---

## 3. Sortida

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

## 4. Editar el catàleg de deficiències (`REQ1.docx`)

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

**Camp de text lliure**:
```
Cal aportar el certificat de la [CAMP: entitat acreditada].
```
Al Pas 5 surt una caixa de text amb l'etiqueta "entitat acreditada".

**Camp amb ajuda**:
```
[CAMP: Òrgan homologació PAU (Protecció civil de Catalunya / Protecció civil local)]
```
El text dins els parèntesis surt com a hint a sota del camp.

**Desplegable**:
```
S'ha de presentar un projecte tècnic [OPCIO: Destinatari | a l'ajuntament | a l'ACA] amb el contingut...
```
Al Pas 5 surt un desplegable etiquetat "Destinatari" amb les opcions.
La primera està preseleccionada per defecte. L'opció triada substitueix
el `[OPCIO: ...]` al document final.

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

## 5. Editar les conclusions (`0 CONCLUSIONS.docx`)

Estructura del fitxer:

```
CONCLUSIONS                                 ← Normal CENTRAT + NEGREITA  (títol del bloc)

Terrassa projecte                            ← Título 1  (títol curt; surt al Pas 4)
La terrassa que apareix... no forma part... ← Normal     (cos que s'imprimeix si la tries)

Requeriment                                  ← Título 1
Vist l'anterior, cal requerir l'esmena...   ← Normal

...

::SEMPRE:: Ho poso al seu coneixement...    ← Normal amb prefix ::SEMPRE::
::SEMPRE:: Cornellà de Llobregat,           ← Normal amb prefix ::SEMPRE::
```

### Regles
- **Títol del bloc** (paràgraf 1, centrat-negreta): es copia tal qual.
  Si vols canviar el text "CONCLUSIONS" per un altre, edita aquest paràgraf.
- **Conclusions triables**: cada una és un **Título 1** (títol curt,
  llegible al Pas 4) seguit d'un **Normal** amb el cos.
- **Parts fixes**: paràgrafs Normal que comencin amb `::SEMPRE:: `.
  No surten al Pas 4: s'imprimeixen sempre, en l'ordre del fitxer,
  al final de l'informe.
- **Camps i opcions**: pots fer servir `[CAMP: ...]` i `[OPCIO: ...]`
  als cossos. El Pas 5 els demana junt amb els del REQ1.
- **Negreta i cursiva** inline (`**...**`, `//...//`) també funcionen.

### Separació visual
Cada conclusió té una separació de 12 pt sota. Si vols més o menys,
edita `suport/Format.ps1`:
```powershell
ConclusionSpaceAfterPt = 12     # prova 18 si vols més espai
```

---

## 6. Capçalera (`0 CAPCALERA.docx`)

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

## 7. Afegir un catàleg nou

Posa un altre `.docx` a `ESTRUCTURALS/` (per ex. `REQ2.docx`) amb la
mateixa estructura del REQ1. **No pot començar per `0 `** (aquests són
plantilles fixes: capçalera i conclusions). Si hi ha més d'un catàleg,
el Pas 1 et deixarà triar.

---

## 8. Configuració local (`suport/config.ps1`)

Fitxer **opcional** per personalitzar rutes/constants només al teu PC.
Si no existeix s'usen els valors per defecte. Variables disponibles:

```powershell
$OutputDir              = '...'    # ruta on desar els informes
$ActivitatsDir          = '...'    # carpeta on viu l'Excel d'activitats
$AlwaysConclusionsCount = 2        # (obsolet, ara s'usa ::SEMPRE::)
```

---

## 9. Resolució de problemes

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

## 10. Per a desenvolupadors

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
