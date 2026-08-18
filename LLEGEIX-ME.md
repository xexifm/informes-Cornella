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

### Llicència (Annex II / LL Prov)

Botó **📜 Llicència (Annex II / LL Prov)**. És el tràmit de llicència d'activitat
de l'Annex II de la Llei 20/2009 i el de llicència provisional. **No és un
informe, són tres**, i tries quin fas al primer pas:

| Fase | Com acaba |
|------|-----------|
| **Requeriment** | «Cal requerir l'esmena de les deficiències indicades…» |
| **Favorable pre-llicència** | «S'informa favorablement a l'espera de rebre la citada documentació…» |
| **Favorable post-llicència** | «S'informa favorablement l'activitat i es dóna per tancat l'expedient.» |

Al mateix pas hi ha la casella **«Llicència provisional»**, que canvia el punt
condicional del principi (compatibilitat urbanística) i afegeix l'**ANNEX 1** al
final — però **només** a la fase de Requeriment.

Passos: fase → capçalera (amb la **Classificació** ja omplerta des de l'Excel:
`Llei 20/2009; Annex II; Epígraf …`) → documentació **abans** de la resolució →
**Projecte** (la mateixa pantalla de deficiències de sempre) → dades del tècnic
redactor i els Id Firmadoc → documentació **després** de la resolució (amb el
«Quan:») → i, si és el favorable pre, les **condicions** de la llicència, en un
quadre de text lliure.

La diferència amb un requeriment normal és que aquí els punts **no són
deficiències sinó documentació**, i surten tant si es té com si no: de cada punt
tries **«No es disposa…»** (surt en negreta) o **«Es disposa… (Id Firmadoc: …)»**.

Al bloc de **després** de la resolució les caselles surten **totes marcades** (és
com tenies el Word: hi eren totes i n'anaves esborrant); hi ha **«Marcar-ho
tot»** i **«Desmarcar-ho tot»**. Al bloc d'**abans** surten desmarcades, perquè
allà cada punt demana a més dir si ja es disposa de la documentació.

El **favorable post-llicència** no et fa tornar a triar res: et demana el `.docx`
del **pre-llicència** i en treu la documentació que hi constava, ja sense el
«Quan:». Només has de desmarcar el que no s'hagi arribat a comprovar.

> **D'on surt el text:** de **REQ1**, en viu. `LLIC.json` només hi afegeix el que
> és propi de Llicència (els dos comentaris i el «Quan:»); si canvies un
> requeriment a REQ1, aquí canvia sol. Si algú reanomena un requeriment de REQ1,
> el programa **avisa** que aquell punt s'ha quedat sense text en lloc de
> callar-s'ho.

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
amb el caixetí a dalt a la dreta de la primera pàgina, alineat amb l'escut i el
marge dret de la capçalera de l'informe.

El text del caixetí és editable a les opcions. Allà mateix hi ha la casella
**«Obrir el registre de la signatura en acabar»**: normalment la vols desmarcada;
marca-la si el caixetí no surt i vols veure què s'ha enviat a l'AutoFirma.

> **La signatura es valida a QUALSEVOL ordinador.** L'AutoFirma munta el PDF i
> el programa li refà la signatura de dins amb la mateixa forma que la fa
> l'Adobe, que és l'única que es validava a tot arreu. Per això **has de triar
> el teu certificat al desplegable**: si hi deixes «(triar-lo a AutoFirma en
> signar)», el programa no sap quin és i no la pot refer. Si algun dia et torna
> a sortir «desconeguda», obre el registre de la signatura i passa'm-lo.
>
> El que ho deixaria resolt per sempre és afegir-hi un **segell de temps**, i
> per això cal demanar a Informàtica de l'Ajuntament la **URL del servei de
> segellat de temps (TSA)** que fan servir. Amb aquella adreça és afegir-la a
> les opcions i llestos.

### Seguiment (fila GIA)

Botó **📊 Seguiment**. Genera els cinc llistats de seguiment a partir de la base
de dades d'activitats, en un sol fitxer amb una pestanya per cada un:

| Pestanya | Què hi surt |
|---|---|
| `Estès` | còpia de la base de dades d'activitats |
| `PRECINTES` | activitats amb el Camp Info **PRECINTE ACTIVITAT?** |
| `DENÚNCIES` | … amb **DENÚNCIA?** |
| `REQUERIT DECRET` | … amb **REQUERIT PER DECRET?** |
| `SONOMETRIA` | … amb **SONOMETRIA?** |
| `ANNEX II` | annex **II** amb **Descripció lliure** escrita |

Hi surt **tota activitat que TINGUI aquell camp**, digui el que digui el valor
(no cal que comenci per SI). Substitueix l'Excel de fórmules que hi havia abans,
i el resultat és el mateix però sense columnes ocultes ni res a recalcular.

**Tries què vols exportar** amb les caselles: hi són totes marcades i en pots
desmarcar les que no necessitis (hi ha un enllaç per marcar-les o desmarcar-les
totes de cop). Al fitxer només hi haurà les pestanyes marcades.

Dos botons: **Exportar a Excel** i **Exportar a PDF**. El PDF surt en horitzontal
i A3, ajustat perquè hi càpiguen totes les columnes, amb les dues primeres files
repetides a cada pàgina. Al peu hi ha el nom de la pestanya i **la pàgina dins
d'aquella pestanya**: encara que el PDF sencer en tingui 600, ANNEX II comença
per *Pàgina 1 de …* amb el total d'ANNEX II. Els fitxers es desen a
`local/seguiment-gia/`.

> Per al PDF val més **desmarcar `Estès`**: són 152 columnes i no està pensada
> per imprimir. Si la deixes marcada, hi sortirà igualment — mana el que triïs.

> Necessita tenir l'Excel una estona treballant: amb tota la base de dades pot
> trigar. Mentre ho fa, els botons queden desactivats.

### Comprovar Excel (fila GIA)

Botó **✅ Comprovar Excel**. Agafa les activitats que la base d'informes té en
estat *Precinte / Cessament* i comprova que a l'Excel hi tinguin el Camp Info
**REQUERIT PER DECRET?** o **PRECINTE ACTIVITAT?** amb un valor que comenci per
**SI**. Les que no, te les llista amb la data de l'informe que les va deixar en
aquell estat, perquè puguis actualitzar l'Excel.

> Compte, que no és el mateix criteri que el de **Seguiment**: aquí sí que es
> demana el «SI», i allà no.

### Altres botons del menú

- **🗃 Actualitzar base d'informes**: recorre els informes ja fets i n'extreu
  data, ID GIA i conclusió a `local/base-dades-activitats/informes-db.json`.
- **📋 Editar base d'informes**, **📥 Revisar entrades del mòbil**,
  **⏱ Controls periòdics**, **Activitats extraordinàries**.

### Sota cada eina, quan la vas fer servir

A totes les eines del menú hi surt, en gris i lletra petita, **l'última vegada
que la vas obrir** (o `(mai)`). Serveix per no haver de recordar si ja havies
passat, per exemple, el *Comprovar Excel* aquesta setmana. Es desa a
`local/base-dades-activitats/eines-state.json` i no es puja mai.

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

### El catàleg de Llicència (`LLIC.json`)

És diferent de la resta i val la pena saber-ho abans de tocar-lo: **no hi ha el
text dels requeriments**. Cada punt hi porta només una **clau** que apunta a un
requeriment de REQ1 (bloquejada a l'editor, com la d'ACT_EXTR) i el que és propi
de Llicència:

| Tipus | Què és |
|---|---|
| `nodisposa` | El comentari de quan **falta** la documentació. |
| `sidisposa` | El de quan **ja es té** (hi va l'Id Firmadoc). |
| `quan` | El «Quan:» del bloc de després de la resolució. |

Els punts que **no** existeixen a REQ1 (els dos condicionals de compatibilitat i
l'ANNEX 1) van a la secció `PROPIS` i sí que hi porten el text sencer.

### Afegir un catàleg nou

Posa un `.json` nou a `ESTRUCTURALS` (el més fàcil: copia'n un i edita'l amb
l'editor). Apareixerà tot sol al menú. Els que comencen per `0 `, els
`ACT_EXTR_*` i `LLIC` no surten al menú d'informes nous: són fitxers de sistema
(el de Llicència té botó propi).

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
