# Estructura del projecte — qui fa servir cada fitxer

Aquest document és el **mapa** del repositori: et diu quins fitxers fa servir
cada programa (cada `.bat` que pots clicar) i quins són compartits.

## Idea general

- **A l'arrel** hi ha els `.bat` que cliques i les carpetes de dades.
- **`suport/`** conté el **motor compartit** (el codi que reutilitzen diversos
  programes) i, en **subcarpetes**, els scripts propis de cada programa.

> **Per què el motor és a `suport/` i no dins d'una carpeta per executable?**
> Perquè `GenerarInforme.ps1` (+ els seus mòduls) és alhora el **generador
> d'informes** *i* el motor que reutilitzen el **vigilant** i l'**exportador**.
> No pertany a un sol executable: és compartit. Per això viu a l'arrel de
> `suport/` i els programes que el fan servir són a subcarpetes.

---

## Arrel del clone (el que veus i cliques)

| Fitxer / carpeta            | Què és                                                        |
|-----------------------------|---------------------------------------------------------------|
| `GenerarInforme.bat`        | ▶ Programa principal. Llança el programa **sense cap finestra de consola** (via `suport\GenerarInforme.vbs`). Si ja està obert, **no n'obre un segon**: porta al davant la finestra existent. Al menú (Pas 1) hi ha, a més dels tipus d'informe, el botó **📍 Generar ruta** i l'interruptor **Vigilant del mòbil** (activar/aturar). |
| `Actualitzar.bat`           | ▶ Actualitzar el programa des de GitHub i refrescar dades. Si el programa està obert, **el tanca abans** d'actualitzar. |
| `ESTRUCTURALS/`             | Plantilles `.docx` (capçalera, conclusions, catàleg REQ, ACT_EXTR). |
| `BASE DE DADES ACTIVITATS/` | Còpia local de l'Excel d'activitats (fallback sense xarxa).   |
| `BASE DE DADES ACT_EXTR/`   | Registre local d'activitats extraordinàries (mode ACT_EXTR, gitignored). |
| `Informes generats/`        | Sortida `.docx` (local, ignorada per git).                    |
| `Rutes generades/`          | Sortida dels mapes de ruta HTML (local, ignorada per git).    |
| `docs/`                     | Formulari web del mòbil (GitHub Pages).                        |
| `suport/`                   | Codi: motor compartit + scripts de cada programa + proves.    |

---

## Dins de `suport/`

```
suport/
├── GenerarInforme.vbs     ← llançador SENSE consola (el crida GenerarInforme.bat)
├── GenerarInforme.ps1     ← MOTOR + punt d'entrada de GenerarInforme.bat
├── Format.ps1             ← mòdul del motor (format del .docx)
├── Seguiment.ps1          ← mòdul del motor (informes de seguiment + tria de mode)
├── ActExtr.ps1            ← mòdul del motor (mode ACT_EXTR: activitats extraordinàries)
├── DriveApi.ps1           ← mòdul compartit (accés a Google Drive)
├── Comprova-Enllacos.ps1  ← utilitat: comprova els enllaços dels catàlegs
├── ComprovarEnllacos.bat  ← ▶ entrada (doble clic) del comprovador d'enllaços
├── config.ps1             ← la TEVA configuració local (rutes, OSRM…)
├── Instalar.bat           ← instal·lador per a una màquina nova
│
├── rutes/                 ← PROGRAMA: planificador de rutes
│   └── Ruta.ps1               (el llança el botó "📍 Generar ruta" del menú)
│
├── mobil/                 ← PROGRAMA(es): integració amb el mòbil/Drive
│   ├── Vigilant.ps1           (el llança l'interruptor "Vigilant del mòbil" del menú)
│   ├── ExportaDades.ps1       (el crida Actualitzar.bat per exportar dades)
│   └── Authorize-Drive.ps1    (autoritza el PC a Google Drive, un sol cop)
│
├── tests/                 ← proves automàtiques
│   ├── run-tests.ps1          (motor / generador d'informes)
│   └── run-tests-ruta.ps1     (planificador de rutes)
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
| `GenerarInforme.ps1`              | ●              |          | ○        | ○           |
| `Format.ps1`                      | ○              |          | ○        | ○           |
| `Seguiment.ps1`                   | ○              |          | ○        | ○           |
| `ActExtr.ps1`                     | ○              |          | ○        | ○           |
| `DriveApi.ps1`                    | ○              |          | ○        | ○           |
| `config.ps1`                      | ·              | ·        | ·        | ·           |
| `rutes/Ruta.ps1`                  |                | **●**    |          |             |
| `mobil/Vigilant.ps1`              |                |          | ●        |             |
| `mobil/ExportaDades.ps1`          |                |          |          | ● (○ motor) |
| `mobil/Authorize-Drive.ps1`       |                |          | ·        |             |

**Com es troben els fitxers entre ells**
- Cada script calcula la seva ubicació (`$ScriptRoot`) i, si cal, puja al
  directori `suport/` (`$SuportDir`) per carregar el motor o `config.ps1`.
  Així el motor pot viure a `suport/` i els programes a subcarpetes sense
  rutes fràgils.
- **`rutes/Ruta.ps1` és independent**: només llegeix `config.ps1` (no carrega
  el motor; no necessita Word).
- **El motor compartit** és `GenerarInforme.ps1` + `Format.ps1` +
  `Seguiment.ps1` + `DriveApi.ps1`. Un canvi aquí afecta el generador, el
  vigilant i l'exportador alhora.

---

## Proves

```
pwsh -File suport/tests/run-tests.ps1          # motor / generador d'informes
pwsh -File suport/tests/run-tests-ruta.ps1     # planificador de rutes
pwsh -File suport/tests/run-tests-actextr.ps1  # mode ACT_EXTR (Decret 112/2010)
```
(en un Windows sense PowerShell 7, fes servir `powershell` en lloc de `pwsh`.)
