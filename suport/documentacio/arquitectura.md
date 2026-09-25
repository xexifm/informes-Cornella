# Arquitectura: el mapa, el que s'ha simplificat i el que queda proposat

> Resultat d'executar `revisio-arquitectura.md` (setembre 2026). Les xifres són
> mesurades amb eines (AST de PowerShell sobre tot `suport/`, sense `tests/`), no
> deduïdes. Quan el programa torni a créixer, torna a executar aquell prompt i
> actualitza aquest document.

## 1. El mapa

**El detall fitxer per fitxer és a la capçalera de `suport/Motor.ps1`**, que és
qui els carrega. Aquí, les capes i qui depèn de qui:

```
                 Menu.ps1 · Wizard.ps1 · Seguiment.ps1          (entrades)
                                   │
   ┌──────────────┬────────────────┼──────────────────┬───────────────────┐
   │ REQ1 i cia.  │ Llicencia*.ps1 │ ActExtr*.ps1     │ eines del menú     │
   │ Capcalera    │ Dades · Blocs  │ Dades · Blocs    │ PdfSignar · PdfUnio│
   │ SeleccioItems│ Pantalles ·    │ Pantalles ·      │ Informes · Copia-  │
   │ Document     │ l'assistent    │ l'orquestrador   │ Informes · Recorda-│
   │ MnsTraspas   │ LlicenciaDb    │                  │ toris · Controls...│
   └──────┬───────┴───────┬────────┴────────┬─────────┴─────────┬─────────┘
          │  Build-*Blocs (PURES) ──► Write-Informe (MotorInforme.ps1)      │
          │                              │                                  │
          ▼                              ▼                                  ▼
   Camps.ps1 · CatalegJson.ps1      Format.ps1 (COM es veu)       UiComuns.ps1 (finestres)
          │                                                         UiFinestra.ps1
          ▼
   Json.ps1 · Excel.ps1 · Docx.ps1 · Settings.ps1 · Migracio.ps1   (base: no depenen de ningú)
```

**Els que més fan servir els altres** (fitxers que en criden alguna funció):
`Json.ps1` 28 · `Motor.ps1` 20 · `UiComuns.ps1` 19 · `MotorInforme.ps1` 13 ·
`Camps.ps1` 12 · `Excel.ps1` 12 · `Migracio.ps1` 11 · `Activitats.ps1` 10 ·
`CatalegJson.ps1` 9.

**Dependències**: 233 arestes entre fitxers i **cap cicle** (hi ha guard).
`Format.ps1`, `Json.ps1`, `Excel.ps1`, `UiFinestra.ps1` i `Docx.ps1` no depenen
de cap altre fitxer; `MotorInforme.ps1` només de `Camps`, `CatalegJson`, `Format`
i `Migracio` (tots de més avall). Cap mòdul genèric no crida cap client seu.

## 2. Què s'ha fet (un commit cadascun, tot amb la suite en verd i els 19 fitxers d'or)

| Canvi | Per què | Com es va demostrar |
|---|---|---|
| `Llicencia.ps1` (2.558 l.) → `LlicenciaDades` / `Blocs` / `Pantalles` / `Llicencia` | feia quatre coses | llista de funcions idèntica, or idèntic |
| `Informes.ps1` (1.820 l.) → + `CopiaInformes.ps1` i `ComprovarExcel.ps1` | tres eines en un fitxer | ídem |
| `ActExtr.ps1` (1.390 l.) → `ActExtrDades` / `Blocs` / `Pantalles` / `ActExtr` | ídem que Llicència | ídem |
| Codi mort fora: `Write-Linia`, `Write-Tancament`, `_WriteActExtrBodyFav`, `_LastRunText` | només els cridaven les proves | AST de tot `suport/` + `.bat` + `.vbs`; llista = foto − 4 |
| Un sol convertidor JSON → hashtable: `ConvertTo-Mapa` (`Json.ps1`) | n'hi havia quatre còpies | proves noves; or idèntic |
| Vistes d'ACT_EXTR, MNS/Traspàs i conclusions en blocs + `Write-Informe -AmbNivells` | quinze embolcalls `_V*` repetien el motor | or idèntic tret de 10 línies `AIRE|` a la vista de MNS (comprovat per programa) |
| Defecte del motor: l'`aire` apagat treia el títol del panell de navegació de la vista | trobat en migrar les vistes | prova que el reprodueix sense l'arranjament |
| `_LlicLletra` a `LlicenciaDades` | era l'únic cicle entre fitxers | guard de cicles, validat tornant-la al lloc vell |
| `CLAUDE.md` 2.058 → ~1.420 línies; les històries de les eines a `documentacio/eines.md` | ningú no el llegia sencer | — |

**Guards nous** (a `tests/proves/06-guards.ps1`, cadascun validat injectant el
defecte): cap fitxer de `suport/` passa de **1.200 línies** (excepcions amb
sostre propi: `PdfSignar` 1.600, `rutes/Coordenades` 1.550, `EditorCatalegs`
1.450); cap nom de funció definit a **dos fitxers**; cap **cicle** entre
fitxers; `VistaWord.ps1` no crida cap `Format-*` directe.

## 3. Propostes pendents, per benefici / risc

Totes toquen codi que **només corre a Windows** (WinForms, Word, AutoFirma): cap
prova de Linux no les pot validar i canvien, encara que sigui poc, el que es veu.
Per això no s'han fet sense la teva confirmació.

### 3.1 Peu de botons comú — **fet** (a petició de l'usuari)
- `_AddPeuBotons` (`UiFinestra.ps1`, perquè també el fa servir el procés de
  rutes; els `_Style*Button` hi van anar amb ell). **Convenció**: sortir o tornar
  a l'esquerra, avançar a la dreta, l'acció principal al capdavall de la dreta i
  en granat; 32 d'alt, 15 de marge, 10 entre botons, amplada segons el text.
- 36 peus (totes les finestres i diàlegs amb botons de sortir/avançar): 626
  línies fora, 305 dins. Els de la dreta es
  recol·loquen a cada canvi de mida (un panell amb `Dock` o una pestanya encara
  no tenen l'amplada bona quan s'hi posen).
- Canvis que es veuen: els botons que no tenien estil (Seguiment, ACT_EXTR,
  rutes) ara el tenen (al procés de rutes no hi ha el granat i es queden amb
  l'aspecte del sistema); «Enrere» porta sempre la fletxa; «Seguent» ja porta
  la ü; a Enviar correu, «No enviar» i «Enviar» van cadascun a una punta.
- Proves: la col·locació (`_PeuPosicions`, `_PeuAmple`) és pura i es prova a
  Linux; guard: cap botó de peu fet a mà fora d'`UiFinestra.ps1`. La resta, al
  PC (i `_AvisaSolapaments` avisa sol si un botó en trepitja un altre).

### 3.2 Partir `PdfSignar.ps1` (1.521 l.) — **proposat, toca la signatura**
Word → PDF, AutoFirma (amb els reintents) i el registre són tres coses. Seria un
tall mecànic com els altres, però la regla és preguntar abans de tocar la
signatura, i les sis rondes de diagnòstic de `signatura-pdf.md` aconsellen fer-ho
en un moment tranquil i provar-ho al PC.

### 3.3 Joc de proves compartit del format de catàleg PC ↔ mòbil (benefici mitjà, risc baix) — **proposat**
El format (`[CAMP:]`, `[OPCIO:]`, `**`, `//`, `[[URL]]`) està escrit dues vegades:
`Camps.ps1`/`CatalegJson.ps1` (PowerShell) i `docs/app.js` (1.201 l.). No es pot
fer una sola implementació, però sí un `docs/dades/proves-format.json` amb
entrades i sortides esperades que llegeixin la suite de PowerShell i una prova de
`node` per a `app.js`. Cal decidir si la suite pot dependre de `node`; toca
`docs/`, que és el que publica el mòbil.

### 3.4 Màquina de passos dels assistents (benefici baix ara) — **no, de moment**
Tres assistents (`Wizard.ps1`, `Llicencia.ps1`, `Seguiment.ps1`) fan
`while` + `switch ($step)` amb Enrere/Endavant. Els passos s'assemblen, però
cadascun fa coses entre pas i pas (carregar la base, saltar-ne si és MNS…) que un
genèric hauria d'aprendre amb paràmetres. **Si surt un quart assistent**, llavors
sí: `Invoke-Passos` amb una llista de scriptblocks que tornen `fwd`/`back`/`exit`.

### 3.5 Partir `EditorCatalegs.ps1` (1.371 l.) — **no**
És una sola finestra: els controls es fan referència entre ells per closures
(el patró `$fn` hashtable). Partir-lo escamparia l'estat d'una finestra per tres
fitxers. Té sostre propi al guard perquè no creixi.

## 4. Què NO s'ha de fer, i per què

- **No unificar les pantalles de llista + detall** (`Select-LlicDocumentacio`,
  `Show-LlicenciaDb`, `Select-Items`). S'assemblen de lluny: una és un arbre amb
  camps inline, una altra una graella amb filtres i l'altra una llista amb Id
  Firmadoc. El genèric necessitaria vuit paràmetres per cobrir tres casos, i
  això és pitjor que les tres còpies.
- **No apujar el límit de 1.200 línies** quan un fitxer hi arribi: parteix-lo.
- **No fer una sola implementació del format de catàleg** entre PC i mòbil: són
  dos llenguatges i dues plataformes. El que es comparteix són les proves (3.3).
- **No tornar a escriure `Format-*` des d'una família o una vista**: tot passa
  per `Build-*Blocs` + `Write-Informe` (guards).

## 5. Xifres

| | Abans (b07ff1c) | Després |
|---|---|---|
| Fitxers `.ps1` (sense proves) | 56 | 64 |
| Línies `.ps1` (sense proves) | 30.120 | ~30.180 (les capçaleres dels fitxers nous) |
| Fitxer més gran | `Llicencia.ps1` 2.558 | `PdfSignar.ps1` 1.521 |
| Fitxers de més de 1.200 línies | 6 | 3 (excepcions amb sostre) |
| Funcions | 786 | 769 |
| Cicles entre fitxers | 1 | 0 |
| Proves | 2.544 | 2.558 |
| `CLAUDE.md` | 2.058 línies | ~1.420 |
