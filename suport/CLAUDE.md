# Notes per a Claude (mantenir entre sessions)

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

## IMPORTANT: tota la feina ha de convergir a `main`
Cada sessió de Claude Code (web) treballa en una branca pròpia `claude/...`.
Si la feina es queda només en aquesta branca, **l'usuari no la rebrà mai**
amb el seu `git pull` de `main`, i semblarà que "no ha canviat res".

Per tant, **al final de cada sessió**:
1. Assegura't que tot està commitejat a la branca de la sessió.
2. Fusiona la feina a `main` i fes push de `main`:
   ```
   git fetch origin
   git checkout main
   git pull --ff-only origin main
   git merge --no-ff <branca-de-la-sessio>
   git push origin main
   ```
   (o, si la branca de sessió ja conté `main`, n'hi ha prou amb
   `git push origin <branca-de-la-sessio>:main`).
3. Confirma a l'usuari que ja pot actualitzar amb `Actualitzar.bat`.

Si tens dubtes sobre si pots fer push a `main`, pregunta-ho; però el model
de desplegament de l'usuari depèn que la feina arribi a `main`.

## Plànol públic d'activitats precintades
- `suport/rutes/Precintades.ps1` genera `docs/dades/precintades.json` a partir
  de l'Excel d'activitats (fulla "Estès"): les activitats amb el camp lliure
  "PRECINTE ACTIVITAT?" i valor que comença per "SI". La pàgina pública
  `docs/precintades.html` (GitHub Pages) el llegeix i pinta el mapa (Leaflet).
- Ho refresca i puja a `main` **`Actualitzar.bat`** (pas 7). URL pública:
  `https://xexifm.github.io/informes-cornella/precintades.html`.
- **Privadesa**: el JSON només conté activitat genèrica (p.ex. "BAR"), adreça de
  l'establiment, ID intern i coordenades — **mai** la raó social ni el text
  lliure del Valor (que conté noms i tràmits interns). No hi afegeixis dades
  personals: aquesta pàgina és pública.
- Reutilitza les funcions de `Ruta.ps1` carregant-lo en mode headless
  (`RUTA_TEST`); si canvies `Ruta.ps1`, executa també
  `run-tests-precintades.ps1`.
