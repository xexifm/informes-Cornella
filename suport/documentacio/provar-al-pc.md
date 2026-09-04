# Provar el programa al PC — que ho faci Claude, no tu

Aquest document és per a **Claude Desktop (o Claude Code) corrent al PC de la
feina**, amb Windows, Office, la unitat de xarxa `I:` i el clone de git.

**L'usuari NO tocarà l'ordinador.** Tot ho ha de fer Claude des del `cmd`: obrir
el programa, clicar, escriure, mirar què surt i decidir si està bé. Aquest
document diu **com** fer-ho quan una comprovació no es pot fer amb codi.

Serveix perquè hi ha una part del programa que **no es pot provar al contenidor
de Linux on es fan els canvis**: tot el que és WinForms (les finestres) i tot el
que va per COM (Word, Excel, Outlook). La suite de 2.269 asserts cobreix la
lògica pura; això cobreix la resta.

---

## Com conduir el programa des del `cmd`

Tres eines, de la més fiable a la menys:

**1. Comprovar-ho amb PowerShell, sense obrir res** — sempre que es pugui.
La majoria de coses d'aquesta llista es poden verificar carregant el motor com a
biblioteca i cridant les funcions, sense cap finestra:

```powershell
$env:GENINFORME_TEST = '1'      # nomes definicions: no obre res
. .\suport\Motor.ps1
# ...i ja pots cridar Get-HeaderData, Read-FullaEstesa, _LlicNomFitxer...
```

Per a les que necessiten Word/Excel/Outlook de debò, treu el `GENINFORME_TEST` i
crida la funció directament (per exemple `Invoke-InformesDbScan`), que és molt
més fiable que clicar.

**2. Obrir el programa i clicar-hi** — quan la cosa a provar ÉS la finestra.
Llança'l amb `Start-Process` i condueix-lo amb UI Automation, que és el que
Windows ofereix per a WinForms:

```powershell
Start-Process -FilePath 'GenerarInforme.bat'
Start-Sleep -Seconds 6
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$root = [System.Windows.Automation.AutomationElement]::RootElement
# ...FindFirst per ControlType.Window / Button / Edit i Invoke()/SetValue()
```

**3. Captura de pantalla** — per mirar-ho amb els teus ulls quan la
UI Automation no arriba (text pintat a mà, colors, solapaments):

```powershell
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
$b = [System.Drawing.Bitmap]::new(
        [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width,
        [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height)
$g = [System.Drawing.Graphics]::FromImage($b)
$g.CopyFromScreen(0, 0, 0, 0, $b.Size)
$b.Save("$env:TEMP\prova.png", [System.Drawing.Imaging.ImageFormat]::Png)
```

…i després mira't el `.png`. **Fes captura sempre que una comprovació falli**,
encara que l'hagis feta amb codi: és el que permet explicar què passava.

### Regles

- **No enviïs res a ningú.** Els correus es queden en esborrany o en vista
  prèvia. Si un pas sembla que enviarà un correu de debò a un titular, **atura't
  i pregunta**.
- **No facis `git push` ni executis `Actualitzar.bat`** si no t'ho diuen.
- Si una cosa peta, **apunta el text sencer de l'error i segueix** amb la resta.
  No t'aturis a la primera.
- Fes servir una **activitat de prova** sempre que puguis. Si has de fer servir
  una de real, no desis res que la modifiqui.
- Al final, **deixa el PC com estava**: tanca el programa, el Word, l'Excel i
  l'Outlook que hagis obert, i comprova-ho al Gestor de tasques.

---

## La llista

### 0. La suite (sense obrir res)

```
GENINFORME_TEST=1 pwsh -NoProfile -File suport/tests/run-tests-all.ps1
```
(o `powershell` si no hi ha `pwsh`). **Han de sortir 2.269 asserts i 0
fallades.** Si no, para i digues què falla: la resta no té sentit.

### 1. El menú principal  · *el codi s'ha mogut a `Menu.ps1`*

Obre el programa i **fes captura**. Comprova-hi:
1. Els quatre apartats de rajoles (EINES, INFORMES, GIA, MÒBIL) i els botons de
   tipus d'informe.
2. Sota cada rajola, la data d'última execució (o «(mai)»).
3. Que **cap títol quedi tapat** pels xips ✏️ / Dades.
4. El botó 📁 obre la carpeta dels informes i **no tanca el menú**.
5. Els enllaços «Capçalera» i «Conclusions» obren l'editor de catàlegs.

### 2. Els textos del correu  · *les tres pantalles ara són una de sola*

6. MÒBIL → «Textos del correu». Captura: l'assumpte, el cos, i **la línia
   d'ajuda A SOBRE** del quadre gran.
7. Prem «Desar» i després, **des del `cmd`**:
   ```powershell
   (Get-Content .\docs\dades\email-textos.json -Raw | ConvertFrom-Json).bcc.Count
   ```
   **Ha de donar 4.** *Aquest és el punt més important de tota la llista*: el
   bloc `bcc` no s'edita des d'aquella pantalla però és del mateix fitxer, i si
   el desat no el tornés a escriure **es perdrien les quatre adreces de còpia
   oculta**.
8. EINES → «Controls periòdics» → «Editar text»: prova «Restaurar original» i
   després «Desar».
9. EINES → «Recordatoris» → «Editar text…»: desa, tanca, torna-hi i comprova
   que el text hi és.

### 3. L'Excel d'activitats  · *un sol lector, i el lector de cel·la compartit*

10. EINES → «Generar ruta»: genera una ruta i obre el mapa. Captura.
11. EINES → «Coordenades»: genera el mapa, arrossega un punt i baixa l'Excel.
    Comprova que s'obre.
12. EINES → «Controls periòdics»: surt la llista d'activitats **amb les
    columnes plenes** (adreça, dates, correus). Si alguna columna surt buida,
    és el lector de cel·la: apunta-ho.
13. INFORMES → «Actualitzar base»: acaba i diu quants informes ha trobat.
14. GIA → «Comprovar Excel» i GIA → «Seguiment».
15. **Al Gestor de tasques: cap `EXCEL.EXE` corrent** després de tot això.
    ```powershell
    Get-Process EXCEL -ErrorAction SilentlyContinue
    ```
    *És el defecte que s'ha arreglat: tres lectors deixaven l'Excel obert.*

### 4. El Word  · *s'obre sempre per `New-WordApp`*

16. «Requeriment - Nou»: fes un informe sencer d'una activitat de prova i
    obre'l. Mira capçalera, numeració, enllaços i conclusions.
17. **El nom del fitxer** ha de ser `aaaa-mm-dd_Req1_GIA <n>.docx`.
    *(El sanejat de caràcters invàlids ha canviat de `_` a `-`, però amb un GIA
    numèric no es nota.)*
18. INFORMES → «Word a PDF»: converteix l'últim informe generat.
19. Edita un catàleg des del xip ✏️ i desa: en tancar l'editor s'han de
    regenerar les vistes en Word d'`ESTRUCTURALS`.
20. «Informe de seguiment» sobre un informe anterior: **compara'l amb un de fet
    abans d'aquests canvis** — les anotacions datades han de sortir amb la
    mateixa lletra i el mateix espaiat, i **en negreta només les pendents**.

### 5. El correu

21. Genera un requeriment i prova «Enviar correu»: la **vista prèvia** es veu bé
    i els enllaços són clicables. Captura. **No l'enviïs.**
22. Prem «Fet» i comença un informe nou: la vista prèvia **s'ha d'haver
    buidat**. *(Abans es quedava la de l'anterior.)*
23. Posa **dues adreces http a la mateixa línia** del text i mira la vista
    prèvia: **les dues han de ser enllaços**. *(Abans la cursiva se les
    menjava.)*
24. Controls periòdics → «Enviar correu (esborranys)» amb **una** activitat de
    prova: mira l'esborrany a l'Outlook (**no l'enviïs**). Els enllaços
    clicables, i si el text porta `//cursiva//` ha de sortir en cursiva.

### 6. El mòbil

25. Obre la pàgina del mòbil al navegador. Al Pas 2 **no hi ha d'haver camps de
    text que es diguin ORIGEN, DATES ni CLASSIFICACIO**.
26. El comptador ha de dir **«Pas X / 4»**, no «/ 5».
27. Fes un informe sencer i mira la vista prèvia del correu.

---

## L'informe final

Quan acabis, digues:

1. **Què ha anat bé**, per bloc.
2. **Què ha fallat**: el número del punt, el text exacte de l'error i la captura.
3. **Què no has pogut provar** i per què (no hi havia dades, no s'hi arribava…).
   Això és tan important com la resta: val més dir-ho que donar-ho per bo.

## Per què cada bloc

| Bloc | Què s'ha canviat i què podria fallar |
|---|---|
| **1** | `Select-Mode` ha passat de `Seguiment.ps1` a `Menu.ps1`. Si falla, el menú no s'obre o una rajola no despatxa. |
| **2** | Les tres pantalles d'«assumpte + cos» són una de sola. El punt **7** és el crític (el `bcc`). |
| **3** | Els set lectors de l'Excel són un, i el lector de cel·la ara ve amb el context. El punt **12** el prova de debò; el **15**, l'`EXCEL.EXE` orfe. |
| **4** | Les cinc obertures del Word són una (`New-WordApp`), que hi afegeix l'`AutomationSecurity` — el que evita la **Vista protegida** amb fitxers de la unitat de xarxa. El **20** prova que `_ShouldBeBold`, que ara sí que fa servir el motor, decideix la negreta igual que abans. |
| **5** | L'HTML del correu és una sola funció i els URLs s'aparten abans d'aplicar la cursiva (**23**). El **22** era un defecte: el reinici feia `.value` sobre un `<div>`. |
| **6** | El mòbil pintava tres quadres de text que no anaven enlloc, i el comptador deia «/ 5» quan només s'hi arriba a 4. |

## Si alguna cosa falla

Tot el que s'ha tocat són **moviments de codi entre fitxers** i **unificacions de
funcions duplicades**: no hi hauria d'haver cap canvi de comportament, i els 19
fitxers d'or ho confirmen per als documents generats. Si en surt un, serà una
crida que ha quedat apuntant on no toca — i es veurà de seguida.
