# Notes per a Claude (mantenir entre sessions)

## Desplegament de l'usuari
- L'usuari executa el programa des d'un **clone de git local** al seu PC.
- Per actualitzar fa doble clic a **`Actualitzar.bat`** (fa `git pull` de
  `main`); per executar el programa fa doble clic a **`GenerarInforme.bat`**.
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
