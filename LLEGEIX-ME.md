# Generador d'informes — Ajuntament de Cornellà de Llobregat

Programa per muntar informes de deficiències de llicències d'activitat, fer-ne
el seguiment, planificar rutes d'inspecció i passar-los a PDF signat.

> **Per començar:** doble clic a **`GenerarInforme.bat`**. Tot surt d'aquí.
> Al menú hi ha el botó **❓ Ajuda** que obre aquest document.

---

## 1. Què hi ha a la carpeta

```
informes-Cornella/
├── GenerarInforme.bat     ← DOBLE CLIC: el programa
├── Actualitzar.bat        ← DOBLE CLIC: baixar l'última versió i pujar els teus catàlegs
├── LLEGEIX-ME.md          ← això que estàs llegint
│
├── ESTRUCTURALS/          ← les FONTS dels catàlegs (el que dona contingut als informes)
│   ├── 0 CAPCALERA.docx     capçalera dels informes (l'única plantilla de Word)
│   ├── 0 CONCLUSIONS.json   conclusions, per tipus d'informe
│   ├── REQ1.json            catàleg de deficiències
│   ├── TERMINI.json         informe de cos fix (ampliació de termini)
│   └── ACT_EXTR_*.json      activitats extraordinàries (requeriment i favorable)
│
├── local/                 ← TOT el que és d'aquest ordinador. No es puja mai.
│   └── (vegeu local/README.txt: informes generats, rutes, Excel, vistes…)
│
├── docs/                  ← el web del mòbil (GitHub Pages). No el toquis a mà.
└── suport/                ← el codi. No cal tocar-lo.
```

Pots **moure la carpeta `informes-Cornella` on vulguis**: tot és relatiu. Les
rutes de la feina (`I:\…`) es poden canviar des del botó **⚙ Configuració**.

### Les dues regles que expliquen tota l'organització

1. **`ESTRUCTURALS` = font. `local` = derivat i teu.** Els catàlegs són els
   `.json` d'`ESTRUCTURALS` i s'editen amb l'**editor de catàlegs** del
   programa. Les còpies en Word (`local/vistes-catalegs/`) es regeneren soles i
   **no s'editen**: si les toques, els canvis es perdran.
2. **Tot el que és d'aquest ordinador va a `local/`**, que està exclosa del
   GitHub. Com que el repositori és **públic** i els informes porten noms i
   adreces de titulars, res del que hi hagi a dins es pot pujar per accident.

La configuració (rutes, credencials del Drive, l'últim informe…) no és a
`local/` sinó a `%LOCALAPPDATA%\InformesCornella\`, perquè sobrevisqui encara
que tornis a baixar el programa de zero.

---

## 2. Instal·lar-lo en un ordinador nou

Doble clic a **`suport/Instalar.bat`** (o només aquest fitxer, si algú te
l'ha passat solt). Instal·la el Git si cal, baixa el programa del GitHub i el
deixa a punt.

**No instal·la** el Microsoft Word ni l'Excel (fan falta, per llicència no es
poden instal·lar sols) ni l'AutoFirma (només per signar).

Després: obre `GenerarInforme.bat` i, si no ets a la feina amb la unitat `I:`,
prem **⚙ Configuració** i posa-hi les teves carpetes.

---

## 3. El dia a dia

### Generar un informe

| Pas | Què fa |
|-----|--------|
| 1 | **Menú**: tries alhora QUÈ vols fer i, si és un informe nou, el catàleg. |
| 2 | **Dades de la capçalera**. Si escrius un ID GIA que és a l'Excel, s'omple sol. |
| 3 | **Marcar les deficiències** (arbre amb filtre). En marcar-ne una, el text surt a la dreta i, si té opcions o camps, els omples **allà mateix**. |
| 4 | **Triar les conclusions**, també amb els camps dins del propi text. |
| 5 | Es genera el `.docx` i s'obre amb el Word. |

Els botons **Enrere** conserven el que has posat; enrere al Pas 2 torna al menú.
Al Pas 2 hi ha **Recuperar dades últim informe** per clonar l'anterior.

### Seguiment d'un informe

Sobre un informe ja emès, marques quins punts s'han resolt i quins no. És
**iteratiu**: cada entrega afegeix una línia datada sota el punt corresponent, i
només l'última entrega pendent queda en negreta. Les conclusions es
regeneren amb les del grup **SEGUIMENT**.

### Ruta d'inspecció

Botó **📍 Generar ruta**. Escrius els ID d'activitat a visitar i et calcula la
ruta circular més curta des de la base (per defecte Carrer de l'Energia, 97),
amb un mapa numerat que pots imprimir a PDF. Els mapes van a
`local/rutes-generades/`.

> **Privacitat:** al servei de rutes només s'hi envien **coordenades**, mai noms
> ni adreces.

### Word a PDF (i signar)

Botó **📄 Word a PDF**. Per defecte hi surt **l'últim informe que has generat**;
pots triar-ne un altre o una carpeta sencera. Converteix a PDF al mateix lloc i,
si ho marques, els signa amb l'**AutoFirma** i el teu certificat de Windows,
amb el caixetí a dalt a la dreta de la primera pàgina.

### Altres botons del menú

- **🗃 Actualitzar base d'informes**: recorre els informes ja fets i n'extreu
  data, ID GIA i conclusió a `local/base-dades-activitats/informes-db.json`.
- **📋 Editar base d'informes**, **📥 Revisar entrades del mòbil**,
  **⏱ Controls periòdics**, **Activitats extraordinàries**.

---

## 4. Editar els catàlegs

**Des del programa**, amb el botó **✏ Editar catàlegs**. Hi pots afegir,
esborrar i moure seccions, subseccions, ítems i sub-punts, i escriure'n el text.

> Abans els catàlegs eren documents de Word i s'editaven amb els estils
> "Títol 1"/"Títol 2". **Ja no**: la font són els `.json` i l'editor. Els `.docx`
> que veus a `local/vistes-catalegs/` són **còpies per llegir**, es regeneren
> soles i editar-les no serveix de res.

### El que pots escriure dins del text

| Marcador | Què fa |
|---|---|
| `[CAMP: nom]` | Un camp de text que hauràs d'omplir. Mateix nom = mateix valor. |
| `[CAMP: nom (ajuda)]` | Igual, amb un text d'ajuda a sota. |
| `[OPCIO: nom \| A \| B]` | Un desplegable amb les opcions A i B. |
| `**negreta**` | Text en negreta. |
| `//cursiva//` | Text en cursiva (per exemple, els títols de normativa en castellà). |

Un enllaç (URL) es marca com a tal a l'editor i surt a l'informe com a
hipervincle, en cos més petit.

Quan deses, el programa **regenera les vistes en Word** i, en fer
`Actualitzar.bat`, **puja els teus catàlegs al GitHub**. Els teus canvis
sempre manen: si el repositori ha tocat el mateix fitxer, guanya la teva versió.

### Afegir un catàleg nou

Posa un `.json` nou a `ESTRUCTURALS` (el més fàcil: copia'n un i edita'l amb
l'editor). Apareixerà tot sol al menú. Els que comencen per `0 ` i els
`ACT_EXTR_*` no surten al menú d'informes nous: són fitxers de sistema.

### Comprovar que els enllaços funcionen

Doble clic a `suport/ComprovarEnllacos.bat`: prova tots els enllaços dels
catàlegs i et diu quins han caigut.

---

## 5. Treballar fora de la feina

Si no tens la unitat de xarxa `I:`, copia l'Excel
`AAAA-MM-DD ACTIVITATS.xls` més recent a **`local/base-dades-activitats/`**. El
programa prova primer la ruta de la feina i, si no hi arriba, fa servir el més
recent d'aquesta carpeta i t'ho indica al Pas 2 amb l'etiqueta
**[FALLBACK LOCAL]** en taronja.

Amb el botó **⚙ Configuració** pots canviar totes les rutes d'aquest
ordinador (es desen a `%LOCALAPPDATA%`, no al repositori, o sigui que cada PC
té les seves).

---

## 6. Actualitzar el programa

Doble clic a **`Actualitzar.bat`**. Fa, per aquest ordre:

1. Còpia de seguretat dels teus catàlegs (a `%LOCALAPPDATA%`, per si de cas).
2. Els commiteja i els puja al GitHub.
3. Es posa al dia amb el GitHub (`git pull --rebase`).
4. Endreça la carpeta `local/` si véns d'una versió antiga.
5. **Torna a aplicar els teus catàlegs**: la teva versió preval sempre.
6. Regenera les vistes en Word i les dades del mòbil.
7. Torna a obrir el programa.

---

## 7. Si alguna cosa no va

| Símptoma | Què fer |
|---|---|
| *"Word obert: tanca'l"* | Tens un fitxer d'`ESTRUCTURALS` obert al Word. Tanca'l. |
| El programa no s'obre | Comprova que ets a l'última versió: `Actualitzar.bat`. |
| Un error a la finestra negra | Copia el missatge sencer (fitxer i línia) i passa'l a una sessió de Claude. |
| El caixetí de la signatura no surt | Mira `local/base-dades-activitats/pdf-signar-log.txt`: hi diu exactament què ha passat a cada intent. |
| Vull recuperar canvis de codi que havia fet a mà | `git stash list` per veure'ls i `git stash pop` per recuperar-los. |

---

## 8. Preparar informes des del mòbil

Es pot omplir un informe des del telèfon i que el `.docx` es generi sol al PC.
La posada en marxa (Google Drive, credencials, web) és a
**`suport/documentacio/DESPLEGAMENT-MOBIL.md`**.

---

## 9. Per si ho toca algú altre (o Claude)

- El **codi** és a `suport/`. El mapa dels mòduls és a la capçalera de
  `suport/Motor.ps1`; les decisions tècniques i el perquè de cada cosa, a
  `suport/CLAUDE.md`.
- **Proves**: `pwsh -File suport/tests/run-tests.ps1`. Amb
  `$env:GENINFORME_TEST = '1'` el motor només defineix funcions (ni finestres ni
  Word), o sigui que la lògica pura es pot provar fins i tot en un Linux.
- **Branca de desplegament**: `main`. `Actualitzar.bat` només hi puja
  **catàlegs i dades**, mai codi.
