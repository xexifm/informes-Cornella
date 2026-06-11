# Estructura del projecte — qui fa servir cada fitxer

Aquest document és el **mapa** del repositori: et diu quins fitxers fa servir
cada programa (cada `.bat` que pots clicar) i quins són compartits.

> **Per què no hi ha una carpeta per a cada executable?**
> Perquè els programes **comparteixen un nucli**: `suport/GenerarInforme.ps1`
> és alhora el generador d'informes *i* el motor que reutilitzen el vigilant i
> l'exportador de dades. A més, els scripts deriven l'arrel del clone
> (`$RepoRoot`) suposant que viuen exactament a `suport/`. Moure'ls trencaria
> les rutes de `ESTRUCTURALS/`, `BASE DE DADES ACTIVITATS/` i `docs/`. Per
> això es manté **una sola carpeta de codi (`suport/`)** i s'organitza la
> claredat amb aquest mapa.

---

## Arrel del clone (el que veus i cliques)

| Fitxer / carpeta            | Què és                                                        |
|-----------------------------|---------------------------------------------------------------|
| `GenerarInforme.bat`        | ▶ Generar un informe de deficiències.                         |
| `Ruta.bat`                  | ▶ Planificar la ruta d'inspecció d'un llistat d'activitats.   |
| `Vigilant.bat`              | ▶ Generar sol els informes que arriben del mòbil.             |
| `Actualitzar.bat`           | ▶ Actualitzar el programa des de GitHub i refrescar dades.    |
| `ESTRUCTURALS/`             | Plantilles `.docx` (capçalera, conclusions, catàleg REQ).     |
| `BASE DE DADES ACTIVITATS/` | Còpia local de l'Excel d'activitats (fallback sense xarxa).   |
| `Informes generats/`        | Sortida `.docx` (local, ignorada per git).                    |
| `Rutes generades/`          | Sortida dels mapes de ruta HTML (local, ignorada per git).    |
| `docs/`                     | Formulari web del mòbil (GitHub Pages).                        |
| `suport/`                   | Tot el codi (PowerShell) + documentació + proves.             |

---

## Quins fitxers de `suport/` fa servir cada executable

Llegenda: **●** = punt d'entrada (el que llança el `.bat`) · **○** = el carrega (dot-source) · **·** = el pot fer servir.

| Fitxer de `suport/`     | GenerarInforme | **Ruta** | Vigilant | Actualitzar | Instalar |
|-------------------------|:--------------:|:--------:|:--------:|:-----------:|:--------:|
| `GenerarInforme.ps1`    | ●              |          | ○        | ○           |          |
| `Format.ps1`            | ○              |          | ○        | ○           |          |
| `Seguiment.ps1`         | ○              |          | ○        | ○           |          |
| `DriveApi.ps1`          | ○              |          | ○        | ○           |          |
| `Authorize-Drive.ps1`   |                |          | ·        |             |          |
| `Vigilant.ps1`          |                |          | ●        |             |          |
| `ExportaDades.ps1`      |                |          |          | ● (○ del nucli) |      |
| **`Ruta.ps1`**          |                | **●**    |          |             |          |
| `Instalar.bat`          |                |          |          |             | ●        |
| `config.ps1`            | ·              | ·        | ·        | ·           |          |
| `tests/`                | proves         | proves   |          |             |          |

**Notes**
- **`Ruta.ps1` és independent**: l'única cosa que comparteix amb la resta és
  `config.ps1` (rutes i servidor de rutes opcional). No depèn de Word ni del
  nucli del generador. Per això té la seva pròpia lògica de lectura d'Excel.
- **El nucli compartit** és `GenerarInforme.ps1` + `Format.ps1` +
  `Seguiment.ps1` + `DriveApi.ps1`. Qualsevol canvi aquí afecta el generador,
  el vigilant i l'exportador alhora.
- **`config.ps1`** (opcional, no es versiona si conté rutes locals) el
  llegeixen tots els programes que el tenen al costat.

---

## Documentació (a `suport/`)

| Fitxer                   | Tema                                                        |
|--------------------------|-------------------------------------------------------------|
| `README.md`              | Manual d'usuari complet.                                     |
| `ESTRUCTURA.md` (arrel)  | Aquest mapa.                                                 |
| `PLA-MOBIL.md`           | Disseny del flux des del mòbil.                              |
| `DESPLEGAMENT-MOBIL.md`  | Posada en marxa del mòbil pas a pas.                         |
| `CLAUDE.md`              | Notes per a sessions amb Claude (manteniment).              |

---

## Proves

```
pwsh -File suport/tests/run-tests.ps1        # nucli del generador d'informes
pwsh -File suport/tests/run-tests-ruta.ps1   # planificador de rutes (Ruta.ps1)
```
