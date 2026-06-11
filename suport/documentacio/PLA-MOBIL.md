# Pla: generar informes des del mòbil

> Document de planificació (no és codi en producció). Descriu com afegir la
> possibilitat de **preparar un informe des del mòbil** sense Word, enviar els
> **requeriments per correu** a un destinatari, i fer que el **.docx complet**
> es generi sol al PC perquè estigui a punt quan hi arribis.
>
> Estat: **IMPLEMENTAT** (pendent de prova real al PC). Vegeu la guia de posada
> en marxa a **`suport/DESPLEGAMENT-MOBIL.md`**.

## Decisió presa

- **Transport i dades personals → Google Drive privat.** Les plantilles
  (REQ/conclusions/capçalera, sense dades personals) van al **GitHub públic**
  (GitHub Pages, carpeta `/docs`); les **activitats** (noms i adreces per ID
  GIA) van **només** a una carpeta privada de Google Drive. El repositori es
  manté **públic**, així `Instalar.bat`/`Actualitzar.bat` i Pages segueixen
  funcionant sense friccions.
- **Capçalera:** el PC exporta a Drive només els camps de capçalera per ID GIA
  cada cop que genera un informe; el mòbil els llegeix per auto-emplenar. Si el
  mòbil només envia l'ID GIA, el **PC completa la capçalera** en generar.
- **Auto-actualització:** `Actualitzar.bat` regenera i puja les dades del web
  quan canvies plantilles; el mòbil sempre està al dia sense gestió manual.

## On ha quedat cada peça

| Peça | Fitxer(s) |
|---|---|
| Mode "des de paquet" | `suport/GenerarInforme.ps1` (`-DesDePaquet`, `Invoke-GenerateFromPaquet`) |
| Reconstrucció (proves) | `suport/tests/run-tests.ps1` (`Build-*FromKeys/Titles/Paquet`) |
| Export plantilles + activitats | `suport/mobil/ExportaDades.ps1` |
| Vigilant | `suport/mobil/Vigilant.ps1`, `Vigilant.bat` |
| Web del mòbil | `docs/` (`index.html`, `app.js`, `drive.js`, `config.js`, `estil.css`) |
| Auto-update | `Actualitzar.bat`, hook a `GenerarInforme.ps1` (Pas 2) |

---

## 1. Idea en una frase

El mòbil fa de **selector**: tries catàleg, dades, deficiències, conclusions i
camps, igual que ara al PC (mode "nou", **no** seguiment). El resultat és un
**paquet JSON** —el mateix model que el programa ja desa a `lastreport.json`—.
Amb aquest paquet passen dues coses:

1. **S'envia un correu** amb el **text dels requeriments** (no el .docx) a un destinatari.
2. El paquet **viatja a Google Drive**; el PC el recull i genera el **.docx complet** sol.

El .docx mai surt del PC. El mòbil mai necessita Word ni Excel.

```
   MÒBIL (web, GitHub Pages)                 GOOGLE DRIVE            PC (Windows + Word)
┌───────────────────────────┐            ┌───────────────┐     ┌──────────────────────────┐
│ Passos 1-5 (selecció)     │            │ /Entrada/     │     │ Watcher (Vigilant.ps1)   │
│  → genera paquet.json     │── puja ───▶│  paquet.json  │────▶│  detecta fitxer nou      │
│                           │            │               │     │  → GenerarInforme.ps1    │
│ Botó "Enviar requeriments"│            │ /Processats/  │◀────│     -DesDePaquet          │
│  → email (mailto/EmailJS) │            └───────────────┘     │  → .docx a "Informes     │
└───────────────────────────┘                                   │     generats/"           │
        │                                                        └──────────────────────────┘
        └──▶ destinatari (només text dels requeriments)
```

---

## 2. La peça que ho fa viable: el JSON ja existeix

El programa ja desa cada informe com un paquet JSON (vegeu `GenerarInforme.ps1`,
funcions `Save-LastReport` / `Load-LastReport`, format "versió 1"):

```json
{
  "Version": 1,
  "Timestamp": "<ISO 8601>",
  "CatalegBaseName": "REQ1",
  "Header":          { "ID_GIA": "...", "TITULAR": "...", "ADRECA": "...", ... },
  "SelectedKeys":    [ "Secció::Ítem", "Secció::Ítem::Fill", ... ],
  "FieldValues":     { "nom_camp": "valor", ... },
  "ConclusionTexts": [ "text conclusió 1", ... ]
}
```

Tota la feina es recolza en aquest model: el mòbil **el produeix**, el PC **el
consumeix**. Proposo afegir-hi camps nous sense trencar el format:

```jsonc
{
  "Version": 2,                    // el PC accepta v1 i v2
  "Origen": "mobil",               // d'on ve el paquet (traçabilitat)
  "Email": {                       // opcional: dades del correu de requeriments
    "Destinatari": "...",
    "Assumpte": "..."
  }
  // ...la resta igual que v1
}
```

---

## 3. Les cinc peces

### Peça 1 — Mode "des de paquet" al PC *(base de tot)*

**Fitxer:** `suport/GenerarInforme.ps1` (modificació).

Avui `Main` (línia ~1908) obre l'assistent WinForms de 6 passos, que recull les
dades i al final crida `Build-Document(...)` (línia 1868). Cal afegir un camí
alternatiu que **salti l'assistent** i alimenti `Build-Document` directament
des d'un paquet JSON:

```
GenerarInforme.ps1 -DesDePaquet "C:\ruta\paquet.json"
```

Lògica nova (reaprofitant el que ja hi ha):
1. Llegir el paquet JSON (`Load-LastReport` ja fa el `ConvertFrom-Json`).
2. `Parse-Cataleg` del `REQ<n>.docx` indicat a `CatalegBaseName` → estructura del catàleg.
3. Reconstruir `$selectedSections` a partir de `SelectedKeys` (la funció de
   clau "Secció::Ítem" ja existeix, línia ~180).
4. `Read-Conclusions` (línia 1434) + casar `ConclusionTexts`.
5. `Add-FieldsFromConclusions` + `FieldValues` del paquet.
6. Cridar `Build-Document(...)` exactament com ara → .docx a `Informes generats/`.

**Important:** aquest mode **sí** que necessita Word (és al PC, està bé). No
toca res del flux actual; és una porta d'entrada nova. Es pot provar de seguida
generant un paquet a mà i executant la comanda.

**Esforç:** mitjà. Risc baix (codi additiu).

---

### Peça 2 — Exportar el catàleg a JSON *(perquè el mòbil tingui dades)*

**Fitxer nou:** `suport/ExportaCataleg.ps1`.

El mòbil ha de mostrar les seccions, ítems, camps `[CAMP:]`, opcions `[OPCIO:]`
i les conclusions. Avui això viu dins de `REQ1.docx` i `0 CONCLUSIONS.docx` i
només es pot llegir amb Word (`Parse-Cataleg`, `Read-Conclusions`). Solució:
un script al PC que **exporta el catàleg a un JSON estàtic** que el mòbil
descarrega:

```
ExportaCataleg.ps1  →  suport/web/cataleg.json
```

Conté: seccions/ítems/fills, cossos, enllaços, i la llista de camps i opcions
detectats (parsejant `[CAMP:]` / `[OPCIO:]`), més les conclusions triables.

**Quan s'executa:** cada cop que canviïs les plantilles. Es pot enganxar a
`Actualitzar.bat` perquè es regeneri i es pugi automàticament a GitHub, de
manera que GitHub Pages sempre serveixi un catàleg al dia.

**Esforç:** baix-mitjà. Reaprofita `Parse-Cataleg` / `Read-Conclusions`.

> Limitació honesta: l'**auto-emplenat de la capçalera per ID GIA** (Pas 2 al
> PC) llegeix l'Excel de la unitat de xarxa `I:\...`, que el mòbil no veu. Al
> mòbil, la capçalera s'omple **a mà** (o es deixa buida i el PC l'auto-emplena
> en generar, ja que el PC sí que té l'Excel). Recomanat: que el PC ompli la
> capçalera des de l'Excel en el moment de generar, i al mòbil només es demani
> l'**ID GIA** + el que vulguis sobreescriure.

---

### Peça 3 — Formulari web al mòbil

**Fitxers nous:** `suport/web/` (`index.html`, `app.js`, `estil.css`).
**Allotjament:** GitHub Pages (gratuït; el repo ja és públic).

Una pàgina responsive que:
1. Carrega `cataleg.json`.
2. Reprodueix els passos: ID GIA + capçalera mínima → marcar deficiències
   (llista amb cerca, com el TreeView) → triar conclusions → omplir camps/opcions.
3. Construeix el paquet JSON (format v2).
4. Dos botons finals:
   - **Enviar requeriments per correu** (peça 4).
   - **Preparar informe al PC** → puja el paquet a Google Drive (peça 5).

**Esforç:** mitjà-alt (és el gruix visual). Tecnologia: HTML/JS senzill, sense
framework pesat, perquè sigui fàcil de mantenir.

---

### Peça 4 — Enviar els requeriments per correu

Del paquet es genera el **text dels requeriments** (cos de cada deficiència
seleccionada, en text pla llegible — no el .docx). Dues variants:

- **`mailto:` prefarcit** *(zero infraestructura)*: el botó obre l'app de correu
  del mòbil amb destinatari, assumpte i cos ja escrits; tu prems "Enviar".
  Limitació: el cos llarg pot quedar truncat segons l'app.
- **EmailJS o servei similar** *(enviament automàtic)*: el correu surt sol des
  de la pàgina, sense obrir l'app. Requereix un compte gratuït i una clau.

**Recomanació:** començar amb `mailto:` (simple i sense comptes); passar a
EmailJS si vols enviament automàtic o cossos llargs.

**Esforç:** baix.

---

### Peça 5 — Transport per Google Drive + watcher al PC

**Pujada des del mòbil:** la pàgina puja `paquet.json` a una carpeta de Drive
(p. ex. `Informes-Cornella/Entrada/`) via l'API de Google Drive (login Google
des del mòbil; encaixa amb el teu Gmail).

**Fitxer nou al PC:** `suport/mobil/Vigilant.ps1` (+ `Vigilant.bat` per arrencar-lo).
Un procés que:
1. Vigila la carpeta de Drive sincronitzada al PC (Drive d'escriptori).
2. En detectar un `paquet.json` nou, crida `GenerarInforme.ps1 -DesDePaquet ...`.
3. Mou el paquet a `Processats/` i deixa el .docx a `Informes generats/`.
4. (Opcional) avisa per correu que ja està fet.

**Esforç:** variable. La part de Drive al mòbil és la més delicada (OAuth);
la del watcher al PC és senzilla (`FileSystemWatcher` de PowerShell).

> Avís: el PC ha d'estar **engegat i amb Drive sincronitzant** perquè
> l'informe es generi sol. És inherent a "que ja estigui fet quan arribi".

---

## 4. Ordre d'implementació (fases)

Cada fase deixa alguna cosa **provable**, per no comprometre's amb tot de cop:

| Fase | Què | Resultat provable | Depèn de |
|------|-----|-------------------|----------|
| **F1** | Peça 1 (mode `-DesDePaquet`) | Genero un paquet a mà i el PC en treu el .docx | — |
| **F2** | Peça 2 (`ExportaCataleg.ps1`) | `cataleg.json` correcte a partir de les plantilles | — |
| **F3** | Peça 3 (web) + Peça 4 (`mailto:`) | Formulari al mòbil que envia requeriments i **descarrega** el paquet | F2 |
| **F4** | Peça 5 (Drive + `Vigilant.ps1`) | Cicle complet automàtic mòbil → PC | F1, F3 |
| **F5** | Polit | Auto-export al `Actualitzar.bat`, EmailJS si cal, avisos | totes |

Després de F1+F2+F3 ja tindries valor real (preparar al mòbil, enviar
requeriments, i passar el paquet al PC encara que sigui manualment). F4 és el
que ho fa **automàtic**.

---

## 5. Riscos i decisions obertes

1. **Capçalera al mòbil:** sense Excel, l'auto-emplenat per ID GIA no és
   possible al mòbil. Decisió recomanada: el PC l'omple en generar (sí té
   Excel); al mòbil només ID GIA + ajustos manuals. *(Cal confirmar.)*
2. **OAuth de Google Drive** des d'una pàgina estàtica: factible però és la
   part més fina. Alternativa més simple si es complica: que el mòbil
   **descarregui** el paquet i el pugis tu a Drive (semiautomàtic).
3. **Política de PowerShell per GPO** al PC: si l'empresa la imposa, el
   watcher pot no arrencar. Poc probable en un PC de casa.
4. **Seguretat:** el paquet conté dades d'activitats. La carpeta de Drive ha de
   ser privada del teu compte. El repo és públic: **mai** s'hi puja cap paquet
   ni cap base de dades (ja estan al `.gitignore`).
5. **Manteniment del catàleg:** si edites `REQ1.docx`, cal regenerar
   `cataleg.json` (automatitzable al `Actualitzar.bat`).

---

## 6. Estimació global

- **Feina nova:** 4-5 peces. Cap depèn de tecnologia exòtica.
- **Es reaprofita molt:** el model JSON, `Parse-Cataleg`, `Read-Conclusions`,
  `Build-Document` i tota la lògica de format ja existeixen.
- **El flux actual del PC no es toca:** tot són portes d'entrada noves.
- **Camí recomanat:** F1 → F2 → F3 (valor ràpid), i després F4 (automatisme).

---

## 7. Què necessito de tu per arrencar

- Confirmar la decisió sobre la **capçalera** (punt 5.1).
- Dir-me el **destinatari habitual** dels requeriments (o si és variable).
- Per F4: tenir **Google Drive d'escriptori** instal·lat al PC i decidir la
  carpeta de treball.
- Per quina **fase** vols que comenci.
