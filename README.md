# Informes Cornellà — Generador

Script PowerShell que munta informes de deficiències per a llicències
d'activitat de l'Ajuntament de Cornellà.

## Estructura

```
informes-Cornella/
├── GenerarInforme.bat            ← doble clic per executar
├── GenerarInforme.ps1            ← script principal (cridat pel .bat)
├── Format.ps1                    ← funcions de format del document Word
├── config.ps1                    ← (opcional) sobreescriu rutes/constants
└── ESTRUCTURALS/
    ├── 0 CAPCALERA.docx          ← capçalera fixa (placeholders <<NOM>>)
    ├── 0 CONCLUSIONS.docx        ← un paràgraf per conclusió alternativa
    └── REQ1.docx                 ← catàleg de deficiències
```

Estat persistent (sessió i cache de catàleg) es guarda a
`%LOCALAPPDATA%\InformesCornella\` perquè no embruti el repositori.

Per executar: **doble clic a `GenerarInforme.bat`**. El .bat obre PowerShell
amb la política d'execució correcta i llança el script. Si vols posar una
drecera a l'escriptori, fes-la sobre el .bat (no sobre el .ps1).

Els .docx d'`ESTRUCTURALS/` que comencen per `0 ` són fixos. Tots els
altres es consideren **catàlegs** i apareixen al Pas 1 si n'hi ha més
d'un (si només hi ha REQ1, el script salta el Pas 1 i el fa servir
directament).

## Convencions del catàleg (REQ1.docx)

Estils Word que el parser reconeix:

| Estil      | Significat                                             |
|------------|--------------------------------------------------------|
| Heading 1  | Títol de secció (Autoritzacions, Controls inicials...) |
| Heading 2  | Nom curt de l'ítem (apareix al TreeView)               |
| Normal     | Cos: 1a línia = text principal, següents = URLs/extra  |

**Sub-bullets:** un Heading 2 que comenci per `::CHILD:: ` es tracta
com a fill de l'ítem Heading 2 anterior. Al TreeView apareix indentat
i seleccionable per separat.

**Placeholders al cos:**

- `[CAMP: nom_camp]` — el script demana el valor un sol cop per nom.
- `[CAMP: nom_camp (text d'ajuda)]` — el text entre parèntesis surt
  com a hint a sota del camp del formulari.

Exemples reals al fitxer:
- `Grup CAPCA: [CAMP: Grup CAPCA].` (apareix en 3 ítems → es demana 1 cop)
- `[CAMP: Òrgan homologació PAU (Protecció civil de Catalunya / Protecció civil local)]`

## Capçalera (CAPCALERA.docx)

Placeholders amb format `<<NOM>>` substituïts pels valors del Pas 2.
Els camps actualment demanats:

`ID_GIA`, `EXP_NUM`, `ADRECA`, `ACTIVITAT`, `PETICIONARI`, `DATA`,
`DECISIO`, `TECNIC`.

## Flux d'execució

1. **Pas 1** — Selecció de catàleg (si n'hi ha més d'un).
2. **Pas 2** — Formulari de dades de la capçalera.
3. **Pas 3** — TreeView amb checkboxes per cada secció / ítem / fill.
4. **Pas 4** — Formulari amb un camp per cada `[CAMP: …]` únic detectat.
5. **Pas 5** — Checkboxes per a les conclusions a incloure.
6. Es genera `Informes generats/Informe - <Activitat> - <data>.docx` i
   s'obre amb Word.

## Numeració

Tots els ítems seleccionats (pares i fills) es numeren **globalment**
1, 2, 3, … al document final, agrupats per secció. Els títols de
secció apareixen abans del seu grup d'ítems com a Heading 1.

## Sortida

Per defecte els informes es desen a:

```
I:\Activitats_Ordenances\Activitats\5.- Sergi Fadurdo\0_Plantilles\Powershell\Informes generats
```

amb format de nom `YYYY-MM-DD_<Cataleg>_GIA <id>.docx`, on `<Cataleg>` és
el nom base del fitxer triat al Pas 1 amb la primera lletra en majúscula
(p. ex. `REQ1.docx` → `Req1`). Exemple:

```
2026-05-26_Req1_GIA 1000.docx
```

Si la unitat `I:` no és accessible (script executat fora de la xarxa
municipal), el fitxer cau automàticament a una carpeta `Informes generats`
al costat del `.ps1`.

## Configuració local (`config.ps1`)

Si vols sobreescriure les rutes per defecte (per exemple en un altre PC
o sense unitat de xarxa), crea o edita `config.ps1` al costat del `.ps1`
i posa-hi els valors que vulguis sobreescriure:

```powershell
$OutputDir              = 'D:\Informes\Sortida'
$ActivitatsDir          = 'D:\Informes\Excels'
$AlwaysConclusionsCount = 2
```

El fitxer és opcional. Si no existeix, s'usen els valors per defecte del
`GenerarInforme.ps1`.

## Recuperació de sessió

Després de cada pas, l'estat s'auto-desa a
`%LOCALAPPDATA%\InformesCornella\session.json`. Si la propera execució
detecta una sessió que no es va completar, pregunta si vols
**precarregar les dades** als formularis. Quan la generació acaba amb
èxit, la sessió es neteja.

Útil quan caduca la connexió, et confons al pas 5 i vols repetir, o
generes diversos informes amb capçalera similar.

## Cache del catàleg

El parseig del `.docx` del catàleg es cacheja a
`%LOCALAPPDATA%\InformesCornella\cache\<NomCataleg>.json` amb un hash
SHA-256 del fitxer com a clau. Si el `.docx` no ha canviat, el catàleg
no es torna a parsejar (estalvia entre 1 i 2 segons per execució).

## Filtre al Pas 3

Hi ha una caixa de **filtre** sobre el TreeView del Pas 3: a mesura que
escrius, només es mostren les deficiències que contenen el text (a títol
de secció, ítem o sub-ítem). Els checks marcats es preserven encara que
canviïs el filtre.

## Estat actual

REQ1.docx conté **21 seccions** amb 131 ítems pare i 92 sub-ítems. Inclou
totes les seccions del catàleg original:

| Bloc                                          | Ítems | Fills |
|-----------------------------------------------|------:|------:|
| Autoritzacions / Informes preceptius          | 19    | 9     |
| Pla d'Autoprotecció                           | 1     | 0     |
| Controls inicials                             | 3     | 2     |
| Controls periòdics                            | 7     | 2     |
| Instal·lacions — Legalitzacions               | 2     | 18    |
| Instal·lacions — Inspeccions inicials         | 5     | 0     |
| Instal·lacions — Inspeccions periòdiques      | 14    | 0     |
| Registres                                     | 19    | 0     |
| Incendis — Evacuació                          | 1     | 0     |
| Incendis — Documentació (ITC SP)              | 5     | 8     |
| Incendis — RIPCI                              | 14    | 0     |
| Incendis — CTE DB SI                          | 6     | 21    |
| Projecte                                      | 3     | 0     |
| Activitat                                     | 8     | 0     |
| Restauració — Cuina i extracció de fums       | 9     | 0     |
| Denúncia soroll                               | 2     | 0     |
| Denúncia calor                                | 1     | 0     |
| Denúncia accessibilitat                       | 1     | 0     |
| Denúncia emmagatzematge material exterior     | 2     | 8     |
| Documentació / Rètols                         | 3     | 24    |
| Accessibilitat                                | 6     | 0     |

**Placeholders únics:** 18. `Grup CAPCA` i `Òrgan homologació PAU` són
els únics que es reutilitzen en més d'un ítem (es demanen una sola
vegada).

**Agrupacions naturals creades:** Les seccions originals
"Instal·lacions" i "Incendis" del REQ1 contenien blocs molt grans amb
subepígrafs implícits; s'han partit en 3 i 4 seccions respectivament
per facilitar la selecció al TreeView. La secció "RESTAURACIÓ" del
document original (només títol) s'ha fusionat amb "Cuina — Extracció
de fums".
