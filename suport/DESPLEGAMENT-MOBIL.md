# Posada en marxa del mòbil

Guia per deixar operatiu el flux **preparar informe des del mòbil**. Segueix
els blocs en ordre. Un cop fet, no hauràs de tornar-hi: tot s'actualitza sol.

> **Privadesa.** Les dades d'activitats (noms i adreces) **no** es pugen mai al
> GitHub públic: viuen només a la teva carpeta privada de Google Drive. El
> GitHub públic només conté plantilles (textos de deficiències/conclusions),
> que no tenen dades personals.

---

## Què s'ha afegit (resum tècnic)

| Peça | On | Què fa |
|---|---|---|
| Mode `-DesDePaquet` | `suport/GenerarInforme.ps1` | Genera el `.docx` des d'un paquet JSON, sense l'assistent. |
| `ExportaDades.ps1` | `suport/` | Exporta plantilles → `docs/dades/*.json` (web) i activitats → Drive privat. |
| `Vigilant.ps1` + `Vigilant.bat` | `suport/` i arrel | Vigila la carpeta Entrada de Drive i genera els informes que arriben del mòbil. |
| Web del mòbil | `docs/` | Formulari responsive (GitHub Pages): selecció + correu de requeriments + paquet a Drive. |
| `Actualitzar.bat` | arrel | Ara també regenera i puja les dades del mòbil quan canvies plantilles. |
| Auto-export d'activitats | `GenerarInforme.ps1` (Pas 2) | Cada cop que generes al PC, refresca `activitats.json` a Drive. |

---

## A · Activar GitHub Pages (allotjament gratuït del web)

1. Al GitHub, repositori `informes-cornella` → **Settings → Pages**.
2. **Source: Deploy from a branch**. **Branch: `main`**, carpeta **`/docs`**. Desa.
3. Al cap d'un parell de minuts tindràs una URL tipus
   `https://xexifm.github.io/informes-cornella/`. Aquesta és l'adreça que
   obriràs al mòbil (afegeix-la a la pantalla d'inici).

> De moment mostrarà un error de "dades": és normal fins que facis el bloc D
> (la primera exportació puja els `docs/dades/*.json`).

---

## B · Google Drive

### B.1 Carpetes a Drive
Crea al teu Drive una carpeta `Informes-Cornella` amb tres subcarpetes:
```
Informes-Cornella/
├── Entrada/      (els paquets que arriben del mòbil)
├── Processats/   (paquets ja generats)
└── Dades/        (activitats.json — privat)
```
De cada carpeta, obre-la al navegador i copia l'**ID** de la URL
(`.../folders/`**`AQUEST_TROS`**). Necessitaràs el d'**Entrada** i el de **Dades**.

### B.2 Com accedeix el PC a Drive — dues variants

**Variant 1 — amb Google Drive d'escriptori** (la més senzilla):
Instal·la **Google Drive per a escriptori** al PC i inicia sessió, de manera
que la carpeta `Informes-Cornella` aparegui com una carpeta local. Si la ruta
**no** és `%USERPROFILE%\Google Drive\Informes-Cornella`, posa la teva a
`suport/config.ps1`:
```powershell
$DriveBaseDir = 'G:\El meu Drive\Informes-Cornella'
```

**Variant 2 — SENSE Drive d'escriptori (per API)**, si no el pots instal·lar:
el PC parla amb Drive per HTTPS. Necessites també l'**ID de la carpeta
Processats** (a més d'Entrada i Dades) i fer la configuració del **bloc G**.
Comprova abans que la xarxa deixa sortir cap a Google amb el test del bloc G.

---

## C · Client OAuth de Google (perquè el mòbil llegeixi/escrigui a Drive)

1. Ves a **Google Cloud Console** → crea un projecte (o reutilitza'n un).
2. **API i serveis → Biblioteca**: activa **Google Drive API**.
3. **Pantalla de consentiment OAuth**: tipus **Extern**, afegeix el teu correu
   com a **usuari de prova** (n'hi ha prou per a ús personal).
4. **Credencials → Crea credencials → ID de client OAuth → Aplicació web**:
   - **Orígens de JavaScript autoritzats**: la URL de Pages **sense barra final**,
     p. ex. `https://xexifm.github.io`.
   - Crea i copia el **Client ID** (acaba en `.apps.googleusercontent.com`).
5. Edita **`docs/config.js`** i omple:
   ```js
   GOOGLE_CLIENT_ID:        "....apps.googleusercontent.com",
   DRIVE_ENTRADA_FOLDER_ID: "<ID carpeta Entrada>",
   DRIVE_DADES_FOLDER_ID:   "<ID carpeta Dades>",
   EMAIL_DESTINATARI:       "elteu@destinatari.cat"   // opcional
   ```
6. Puja el canvi de `config.js` (és codi, no plantilla): fes-ho via una sessió
   de Claude o un PR, o amb git si en saps. (`Actualitzar.bat` no puja codi.)

> `config.js` **no conté dades personals**, només identificadors; es pot pujar
> al repo públic sense problema.

---

## D · Primera exportació i engegar el vigilant (al PC)

1. **Exporta-ho tot un cop** (genera les dades del web i puja les d'activitats a Drive):
   ```
   powershell -ExecutionPolicy Bypass -File suport\ExportaDades.ps1
   ```
   - Crea `docs/dades/*.json`. **Puja'ls** (via `Actualitzar.bat` editant
     qualsevol plantilla, o per PR) perquè GitHub Pages els serveixi.
   - Crea `Dades/activitats.json` a Drive.
2. **Deixa el vigilant obert** mentre treballis: doble clic a **`Vigilant.bat`**.
   Vigila `Entrada/` i, quan arribi un paquet del mòbil, genera el `.docx` a
   `Informes generats/` i mou el paquet a `Processats/`.

A partir d'aquí, **cada cop que generes un informe al PC** s'actualitza sol
`activitats.json` a Drive, i **cada cop que canvies una plantilla i fas
`Actualitzar.bat`** s'actualitza sol el web del mòbil. No has de fer res més.

---

## Ús diari

**Al mòbil** (obre la URL de Pages):
1. Tria catàleg (si n'hi ha més d'un) → ID GIA → **Cercar** (auto-emplena des de Drive).
2. Marca deficiències → tria conclusions → omple camps.
3. Pas final:
   - **Enviar requeriments per correu** → s'obre el correu amb el text ja escrit.
   - **Preparar informe al PC** → puja el paquet a Drive.

**Al PC**: el vigilant genera el `.docx` complet en segons. Quan hi arribis, ja
el tens a `Informes generats/`.

---

## E · Enviar els requeriments amb un sol clic (EmailJS)

Perquè el botó **Enviar requeriments per correu** enviï sol (sense obrir cap
app), s'usa **EmailJS** (gratuït per a poc volum):

1. Crea un compte a **https://www.emailjs.com**.
2. **Email Services → Add Service**: connecta el teu correu (p. ex. Gmail).
   Aquesta serà **l'adreça des de la qual s'envia**.
3. **Email Templates → Create Template** amb aquestes variables al cos:
   - Camp **To**: `{{to_email}}`
   - Camp **Subject**: `{{subject}}`
   - Cos del missatge: `{{{message}}}` ← **amb TRES claus** (el correu s'envia
     en **HTML**, amb subseccions subratllades, negretes, etc.). Si poses només
     `{{message}}` (dues claus), es veuran les etiquetes HTML com a text.
4. A **Account → API Keys** copia la **Public Key**, i anota el **Service ID**
   i el **Template ID**.
5. Omple'ls a **`docs/config.js`**:
   ```js
   EMAILJS_PUBLIC_KEY:  "...",
   EMAILJS_SERVICE_ID:  "...",
   EMAILJS_TEMPLATE_ID: "...",
   EMAIL_DESTINATARI:   "destinatari@exemple.cat"
   ```
6. (Recomanat) Al panell d'EmailJS, restringeix l'enviament al teu domini de
   Pages per evitar que algú altre faci servir la teva quota.

L'assumpte del correu surt sempre com **`GIA <id> Requeriments`**. Si no
configures EmailJS, el botó torna al comportament d'obrir l'app de correu.

## F · Logo i colors de l'Ajuntament

- **Logo:** posa el fitxer del logo a **`docs/img/logo.png`** (o `.svg` canviant
  el `src` a `index.html`). Si no n'hi ha, simplement no es mostra.
- **Colors:** són variables CSS a dalt de **`docs/estil.css`** (`--brand`,
  `--brand-dark`...). Ara hi ha un verd corporatiu **aproximat**; passa'm els
  codis hex exactes (o el logo) i els deixo idèntics als de cornella.cat.

## Qui pot accedir al web?

El web és a **GitHub Pages**, o sigui **públic**: qualsevol que tingui l'enllaç
el pot obrir. Però:
- **No conté dades personals** (només plantilles).
- **Cercar** (dades d'activitats) necessita el **teu** compte de Google: un
  estrany no hi té accés.
- **Preparar informe al PC** puja a **el teu** Drive: un estrany no hi pot
  escriure.

És a dir, el formulari és visible, però les dades i el teu PC queden protegits.
Si vols, et puc afegir un **PIN senzill** d'entrada (dissuasiu, no una seguretat
forta) o moure-ho a un allotjament amb contrasenya.

## G · PC sense Google Drive d'escriptori (accés per API)

Si no pots instal·lar Drive d'escriptori, el PC accedeix a Drive per API.

**G.1 — Comprova que la xarxa deixa sortir cap a Google.** A PowerShell:
```powershell
$urls="https://www.googleapis.com/discovery/v1/apis","https://oauth2.googleapis.com/tokeninfo","https://www.googleapis.com/drive/v3/about"
foreach($u in $urls){try{$r=Invoke-WebRequest $u -UseBasicParsing -TimeoutSec 15;"ACCESSIBLE $($r.StatusCode)  $u"}catch{$x=$_.Exception.Response;if($x){"ACCESSIBLE $([int]$x.StatusCode)  $u"}else{"BLOQUEJAT  $u"}}}
```
Si surten **ACCESSIBLE**, endavant. Si surt **BLOQUEJAT**, usa un PC amb Drive d'escriptori.

**G.2 — Crea un segon client OAuth, de tipus *Aplicació d'escriptori*** (mateix
projecte de Google Cloud que el del mòbil): **Credencials → Crea credencials →
ID de client OAuth → Aplicació d'escriptori**. Apunta el **Client ID** i el
**Client Secret**.

**G.3 — Posa els IDs de les tres carpetes** a `suport/config.ps1`:
```powershell
$DriveEntradaId    = '<ID carpeta Entrada>'
$DriveProcessatsId = '<ID carpeta Processats>'
$DriveDadesId      = '<ID carpeta Dades>'
```

**G.4 — Autoritza el PC una sola vegada:**
```
powershell -ExecutionPolicy Bypass -File suport\Authorize-Drive.ps1
```
Enganxa el Client ID i el Secret, autoritza al navegador amb el teu compte i
ja està. Les credencials queden a `%LOCALAPPDATA%\InformesCornella\` (mai al repo).

A partir d'aquí, `ExportaDades.ps1` puja `activitats.json` a Drive per API i
`Vigilant.bat` recull els paquets de Drive per API automàticament. No cal
carpeta sincronitzada.

## Mode sense Drive (fallback)

Si encara no has fet el bloc C, el web funciona igual però:
- La capçalera s'omple **a mà** (el PC l'acabarà d'omplir des de l'Excel en generar).
- En comptes de pujar el paquet, el botó **Baixar paquet** te'l descarrega;
  deixa'l tu a la carpeta `Entrada/` de Drive i el vigilant farà la resta.

---

## Resolució de problemes

- **El web diu "Error carregant les dades"**: encara no has pujat
  `docs/dades/*.json` (bloc D, pas 1) o Pages no està actiu (bloc A).
- **"No s'ha trobat activitats.json a Drive"**: genera un informe al PC un cop
  (o executa `ExportaDades.ps1 -Activitats`) perquè es creï a `Dades/`.
- **Google demana permisos cada vegada**: normal en apps en "mode de prova";
  per a ús personal és suficient.
- **El vigilant no genera res**: comprova que `Vigilant.bat` està obert, que la
  ruta `$DriveBaseDir` és correcta i que Drive d'escriptori està sincronitzant.
- **Un paquet falla**: el vigilant el mou a `Processats/` amb sufix `.error`;
  obre'l i revisa'l, o torna a preparar-lo des del mòbil.

---

## Nota de verificació

Les parts que depenen de **Word/Excel** (mode paquet, exportació) i de **Google
Drive** (web) no s'han pogut executar en l'entorn on s'ha programat (Linux sense
Office). La **lògica pura** (reconstrucció de selecció, camps, claus) sí que té
proves automàtiques a `suport/tests/run-tests.ps1`. Convé fer una primera prova
real al PC: prepara un informe senzill al mòbil i comprova que el `.docx`
generat pel vigilant és correcte.
