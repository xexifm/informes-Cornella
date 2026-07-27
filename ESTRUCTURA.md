# Estructura del projecte — qui fa servir cada fitxer

Aquest document és el **mapa** del repositori: et diu quins fitxers fa servir
cada programa (cada `.bat` que pots clicar) i quins són compartits.

## Idea general

- **A l'arrel** hi ha els `.bat` que cliques i les carpetes de dades.
- **`suport/`** conté el **motor compartit** (el codi que reutilitzen diversos
  programes) i, en **subcarpetes**, els scripts propis de cada programa.

> **Per què el motor és a `suport/` i no dins d'una carpeta per executable?**
> Perquè `Motor.ps1` (+ els seus mòduls) és el motor que reutilitzen el
> **generador**, el **vigilant** i l'**exportador**. No pertany a un sol
> executable: és compartit. Per això viu a l'arrel de `suport/` i els programes
> que el fan servir són a subcarpetes.

> **Motor i punt d'entrada són fitxers DIFERENTS.** `Motor.ps1` només
> *defineix* (funcions, rutes, configuració): carregar-lo no obre cap finestra
> ni genera res. `GenerarInforme.ps1` només *arrenca* (carrega el motor i crida
> `Main`, o `Invoke-GenerateFromPaquet` amb `-DesDePaquet`). Qui vol el motor
> com a **biblioteca** (vigilant, exportador, proves) carrega `Motor.ps1`; qui
> vol el **programa** executa `GenerarInforme.ps1`.

---

## Arrel del clone (el que veus i cliques)

| Fitxer / carpeta            | Què és                                                        |
|-----------------------------|---------------------------------------------------------------|
| `GenerarInforme.bat`        | ▶ Programa principal. Llança el programa **sense cap finestra de consola** (via `suport\GenerarInforme.vbs`). Si ja està obert, **no n'obre un segon**: porta al davant la finestra existent. Al menú (Pas 1) hi ha, a més dels tipus d'informe, els botons **📍 Generar ruta**, **🔒 Activitats precintades** (obre el plànol públic), **🗃 Actualitzar base d'informes**, **📋 Editar base d'informes**, **📥 Revisar entrades del mòbil** (comprovació d'un sol cop), **⚙ Configuració** (rutes d'aquest PC + actualitzar el programa) i **❓ Ajuda** (obre el manual). |
| `Actualitzar.bat`           | ▶ Actualitzar el programa des de GitHub i refrescar dades. Si el programa està obert, **el tanca abans** d'actualitzar. |
| `ESTRUCTURALS/`             | Plantilles `.docx` (capçalera, conclusions, catàleg REQ, ACT_EXTR). |
| `BASE DE DADES ACTIVITATS/` | Còpia local de l'Excel d'activitats (fallback sense xarxa) i `informes-db.json` (base d'informes generada des del menú, gitignored). |
| `BASE DE DADES ACT_EXTR/`   | Registre local d'activitats extraordinàries (mode ACT_EXTR, gitignored). |
| `Informes generats/`        | Sortida `.docx` (local, ignorada per git).                    |
| `Rutes generades/`          | Sortida dels mapes de ruta HTML (local, ignorada per git).    |
| `docs/`                     | Web pública (GitHub Pages): formulari del mòbil (`index.html`) i **plànol públic d'activitats precintades** (`precintades.html`). |
| `suport/`                   | Codi: motor compartit + scripts de cada programa + proves.    |

---

## Dins de `suport/`

```
suport/
├── GenerarInforme.vbs     ← llançador SENSE consola (el crida GenerarInforme.bat)
├── GenerarInforme.ps1     ← PUNT D'ENTRADA (només arrenca: carrega Motor.ps1 i crida Main)
├── Motor.ps1              ← MOTOR compartit (només definicions; no executa res)
├── UiComuns.ps1           ← mòdul del motor (helpers WinForms compartits: _NewForm,
│                            banda granat, estils de botó, _MakeMultiFilter, _AddConfigRow)
├── Format.ps1             ← mòdul del motor (format del .docx)
├── Seguiment.ps1          ← mòdul del motor (informes de seguiment + tria de mode)
├── ActExtr.ps1            ← mòdul del motor (mode ACT_EXTR: activitats extraordinàries)
├── Informes.ps1           ← mòdul del motor (escàner d'informes → informes-db.json)
├── DriveApi.ps1           ← mòdul compartit (accés a Google Drive)
├── Settings.ps1           ← mòdul compartit (rutes d'aquest PC, %LOCALAPPDATA%\...\settings.json)
├── Configuracio.ps1       ← mòdul del motor (pantalla "⚙ Configuració" + actualitzar el programa)
├── Comprova-Enllacos.ps1  ← utilitat: comprova els enllaços dels catàlegs
├── ComprovarEnllacos.bat  ← ▶ entrada (doble clic) del comprovador d'enllaços
├── config.ps1             ← configuració COMPARTIDA (git) de valors per defecte (rutes, OSRM…)
├── Instalar.bat           ← instal·lador per a una màquina nova
│
├── rutes/                 ← PROGRAMA(es): rutes i mapes a partir de l'Excel
│   ├── Ruta.ps1               (el llança el botó "📍 Generar ruta" del menú)
│   └── Precintades.ps1        (genera docs/dades/precintades.json per al plànol
│                               públic; el crida Actualitzar.bat. Reutilitza
│                               les funcions de Ruta.ps1 en mode headless)
│
├── mobil/                 ← PROGRAMA(es): integració amb el mòbil/Drive
│   ├── Vigilant.ps1           (el llança el botó "📥 Revisar entrades del mòbil":
│   │                           mira un sol cop si han arribat informes del mòbil)
│   ├── ExportaDades.ps1       (el crida Actualitzar.bat per exportar dades)
│   └── Authorize-Drive.ps1    (autoritza el PC a Google Drive, un sol cop)
│
├── tests/                 ← proves automàtiques
│   ├── run-tests.ps1              (motor / generador d'informes)
│   ├── run-tests-ruta.ps1         (planificador de rutes)
│   └── run-tests-precintades.ps1  (mapa d'activitats precintades)
│
└── documentacio/          ← guies tècniques
    ├── PLA-MOBIL.md
    └── DESPLEGAMENT-MOBIL.md
```

> `README.md` i `CLAUDE.md` es queden a `suport/` (manual i notes de
> manteniment).

---

## Quins fitxers fa servir cada executable

Llegenda: **●** = punt d'entrada · **○** = el carrega (dot-source) · **·** = el pot fer servir.

| Fitxer                            | GenerarInforme | **Ruta** | Vigilant | Actualitzar |
|-----------------------------------|:--------------:|:--------:|:--------:|:-----------:|
| `GenerarInforme.ps1`              | ●              |          |          |             |
| `Motor.ps1`                       | ○              |          | ○        | ○           |
| `UiComuns.ps1`                    | ○              |          | ○        | ○           |
| `Format.ps1`                      | ○              |          | ○        | ○           |
| `Seguiment.ps1`                   | ○              |          | ○        | ○           |
| `ActExtr.ps1`                     | ○              |          | ○        | ○           |
| `Informes.ps1`                    | ○              |          | ○        | ○           |
| `DriveApi.ps1`                    | ○              |          | ○        | ○           |
| `Settings.ps1`                    | ○              | ○        | ○        | ○           |
| `Configuracio.ps1`                | ○              |          | ○        | ○           |
| `config.ps1`                      | ·              | ·        | ·        | ·           |
| `rutes/Ruta.ps1`                  |                | **●**    |          | ○ (Precint.)|
| `rutes/Precintades.ps1`           |                |          |          | ● (○ Ruta)  |
| `mobil/Vigilant.ps1`              |                |          | ●        |             |
| `mobil/ExportaDades.ps1`          |                |          |          | ● (○ motor) |
| `mobil/Authorize-Drive.ps1`       |                |          | ·        |             |

**Com es troben els fitxers entre ells**
- Cada script calcula la seva ubicació (`$ScriptRoot`) i, si cal, puja al
  directori `suport/` (`$SuportDir`) per carregar el motor o `config.ps1`.
  Així el motor pot viure a `suport/` i els programes a subcarpetes sense
  rutes fràgils.
- **`rutes/Ruta.ps1` és independent**: només llegeix `config.ps1` (no carrega
  el motor; no necessita Word). També llegeix `Settings.ps1` (dot-source
  propi) per aplicar l'override d'aquest PC a `$ActivitatsDir`/`$RutesOutputDir`.
- **`Settings.ps1`** és l'ÚNIC lloc on es llegeix/escriu
  `%LOCALAPPDATA%\InformesCornella\settings.json` (rutes d'aquest PC, mai a
  git). El carreguen per separat `Motor.ps1` i `rutes/Ruta.ps1`
  (processos/scopes independents), cadascun DESPRÉS del seu propi `config.ps1`.
- **`rutes/Precintades.ps1`** genera les dades del **plànol públic** d'activitats
  precintades (`docs/dades/precintades.json`). Carrega `Ruta.ps1` en mode
  headless per reutilitzar-ne les funcions (conversió UTM, format d'adreça,
  cerca de l'Excel). El crida `Actualitzar.bat`, que puja el JSON a `main`.
- **El motor compartit** és `Motor.ps1` + `UiComuns.ps1` + `Format.ps1` +
  `Seguiment.ps1` + `DriveApi.ps1`. Un canvi aquí afecta el generador, el
  vigilant i l'exportador alhora.
- **Reutilitzar el motor des d'un script de consola**: posa
  `$MotorSenseGui = $true` i fes dot-source de `Motor.ps1` (ho fan
  `mobil/Vigilant.ps1` i `mobil/ExportaDades.ps1`). La bandera només evita
  carregar WinForms; carregar el motor mai arrenca el programa.
- **`UiComuns.ps1` es carrega el primer** i no coneix res del motor. Hi van els
  helpers de WinForms que fan servir diverses pantalles, perquè cap mòdul hagi
  de dependre del punt d'entrada (ni d'una altra pantalla) per dibuixar-se.

---

## Proves

```
pwsh -File suport/tests/run-tests-all.ps1          # TOTES (recomanat)
```
Cada suite corre en un procés a part (els dobles i les variables d'entorn
d'una no poden contaminar la següent) i el codi de sortida és 0 només si
passen totes. També es poden triar: `run-tests-all.ps1 -Suite ruta,actextr`.

Per executar-ne una de sola:
```
pwsh -File suport/tests/run-tests.ps1              # motor / generador d'informes
pwsh -File suport/tests/run-tests-ruta.ps1         # planificador de rutes
pwsh -File suport/tests/run-tests-precintades.ps1  # mapa d'activitats precintades
pwsh -File suport/tests/run-tests-actextr.ps1      # mode ACT_EXTR (Decret 112/2010)
```
(en un Windows sense PowerShell 7, fes servir `powershell` en lloc de `pwsh`.)

Fitxers compartits de les proves: `tests/TestLib.ps1` (`Assert`, `AssertEq`,
`AssertNear`, `Write-TestSummary`) i `tests/FormatDoubles.ps1` (els dobles de
les funcions `Format-*`, que capturen a `$global:emitCalls` què rebria Word).
