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
    └── Crear-acces-directe.bat ← DOBLE CLIC: l'accés directe per ancorar-lo a la barra de tasques
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

### Tenir-lo a la barra de tasques

Windows **no deixa ancorar un `.bat`** a la barra de tasques: només hi admet
accessos directes que apuntin a un programa. Doble clic a
**`suport\Crear-acces-directe.bat`** i te'n deixa un a l'**escriptori** i al **menú
Inici**, amb l'escut de l'Ajuntament.

Després, **clic dret** damunt de l'accés directe → **«Ancorar a la barra de
tasques»**. Al Windows 11 potser primer has de triar «Mostra més opcions».

**Si ja el tenies ancorat**, treu-lo de la barra i torna-hi a posar: Windows es
queda la còpia del dia que el vas ancorar, i fins que no el tornis a ancorar
seguirà sense icona i obrirà un segon botó.

> No s'ancora sol perquè des del Windows 10 aquesta ordre ja no es pot fer per
> codi: el que es troba per internet són trucs que toquen el registre i
> reinicien l'explorador, i es poden carregar la barra de tasques.
>
> L'accés directe l'obre `Instalar.bat` per tu quan instal·les de zero. Si mous
> la carpeta del programa, torna a fer doble clic al `.bat` per refer-lo.

---

## 3. El dia a dia

### Generar un informe

| Pas | Què fa |
|-----|--------|
| 1 | **Menú**: tries alhora QUÈ vols fer i, si és un informe nou, el catàleg. |
| 2 | **Dades de la capçalera**. Si escrius un ID GIA que és a l'Excel, s'omple sol. |
| 3 | **Marcar les deficiències** (arbre amb filtre). En marcar-ne una, el text surt a la dreta i, si té opcions o camps, els omples **allà mateix**. Els punts amb una **ⓘ** al davant porten fitxa d'ajuda: el punter es torna una mà quan hi passes per sobre; clica-la (o prem **F1**) i et diu si allò s'ha de requerir o no. |
| 4 | **Triar les conclusions**, també amb els camps dins del propi text. |
| 5 | Es genera el `.docx` i s'obre amb el Word. |

Els botons **Enrere** conserven el que has posat; enrere al Pas 2 torna al menú.
Al Pas 2 hi ha **Recuperar dades últim informe** per clonar l'anterior.

**Si hi ha un requeriment anterior sense resposta**, marca *Requeriment pendent →
Anterior requeriment*. L'informe surt en dos blocs ben separats:

- **REQUERIMENT ANTERIOR** (subratllat): l'avís «S'ha de donar resposta a
  l'anterior requeriment…», **sense número**, i a sota el text *COPIAR
  REQUERIMENT*. Allà hi enganxes el text **literal** de l'anterior i en canvies
  els números **1, 2, 3… per A1, A2, A3…** (el mateix *COPIAR REQUERIMENT* t'ho
  recorda).
- **REQUERIMENT ACTUAL** (subratllat): la frase «S'han observat les següents
  deficiències…» i els punts d'ara, numerats **des de l'1**.

El **Seguiment** llegeix tant els A1, A2… com els 1, 2…, o sigui que podràs
marcar també els punts de l'anterior.

### Seguiment d'un informe

Sobre un informe ja emès, marques quins punts s'han resolt i quins no. És
**iteratiu**: cada entrega afegeix una línia datada sota el punt corresponent, i
només l'última entrega pendent queda en negreta. Les conclusions es
regeneren amb les del grup **SEGUIMENT**.

### Llicència (Annex II / LL Prov)

Botó **📜 Llicència (Annex II / LL Prov)**. És el tràmit de llicència d'activitat
de l'Annex II de la Llei 20/2009 i el de llicència provisional. **No és un
informe, són cinc**, i tries quin fas al primer pas:

| Informe | Com acaba |
|------|-----------|
| **Requeriment** | «Cal requerir l'esmena de les deficiències indicades…» |
| **Favorable pre-llicència** | «S'informa favorablement a l'espera de rebre la citada documentació…» |
| **Favorable post-llicència** | «S'informa favorablement l'activitat i es dóna per tancat l'expedient.» |
| **Modificació NO Substancial** | informe curt, a part (vegeu més avall) |
| **Transmissió** | informe curt, a part (vegeu més avall) |

Al mateix pas hi ha la casella **«Llicència provisional»**, que canvia el punt
condicional del principi (compatibilitat urbanística) i afegeix l'**ANNEX 1** al
final — però **només** a la fase de Requeriment, i **només si encara no es
disposa** d'aquella autorització: l'annex diu com demanar-la, i si ja la tens no
hi pinta res. (La casella no s'aplica als dos informes curts.)


Passos: fase → capçalera (amb la **Classificació** ja omplerta des de l'Excel:
`Llei 20/2009; Annex II; Epígraf …`) → documentació **abans** de la resolució →
**Projecte** (la mateixa pantalla de deficiències de sempre) → dades del tècnic
redactor i els Id Firmadoc → documentació **després** de la resolució (amb el
«Quan:») → i, als dos favorables, **qui posa condicions**.

**Les condicions.** Als favorables, l'últim pas és una llista dels organismes
que informen els punts d'*Autoritzacions / Informes preceptius* (OGAU, Agència
de Residus de Catalunya, Direcció General de Canvi Climàtic i Qualitat
Ambiental…). Surten ja marcats els que tenen l'informe preceptiu com a «Es
disposa». Si en marques algun, la conclusió diu «…sota les condicions que es
determinen en els següents informes (adjunts a continuació):» i a sota hi
surten els que has marcat (a., b., c.…); els seus informes els adjuntes tu
darrere. Si no en marques cap, la conclusió no parla de condicions. La llista
és a `LLIC.json` (secció `CONDICIONS`) i la pots ampliar des de l'editor de
catàlegs.

Al costat de cada organisme hi ha el botó **«PDF…»** per triar el seu informe.
El programa en guarda una còpia a
`local\base-dades-llicencies\GIA 924\a.OGAU.pdf` (la lletra és la de l'informe)
i se'n recorda per al post. Quan passis l'informe a PDF amb **📄 Word a PDF**,
aquests informes s'hi afegeixen **darrere** abans de signar. Les seves
signatures s'hi continuen **veient** però ja no són camps de signatura, o sigui
que l'Adobe no dirà «signatura no vàlida»; els originals signats queden a la
carpeta de la llicència. Si algun adjunt no hi és, aquell PDF no es signa i el
resum t'ho diu.

La diferència amb un requeriment normal és que aquí els punts **no són
deficiències sinó documentació**, i surten tant si es té com si no: de cada punt
tries **«No es disposa…»** (surt en negreta) o **«Es disposa… (Id Firmadoc: …)»**.

La pantalla és com la de marcar deficiències: la **llista a l'esquerra** (amb
cercador) i el **detall a la dreta**. Al detall tries «No es disposa / Es
disposa» i, al bloc d'**abans**, hi omples els camps **allà mateix** — l'Id
Firmadoc i, segons el punt, l'expedient, la referència o el registre. Cada
document té les seves dades.

**Tots** els punts d'abans et deixen dir que es disposa del document amb el seu
Id Firmadoc, encara que `LLIC.json` no en digui res: els que hi tenen una
redacció pròpia la mantenen, i la resta agafen
«Es disposa del document (Id Firmadoc: …)».

Al bloc d'**abans** hi surten els punts de quatre seccions de REQ1
(*Autoritzacions*, *Pla d'Autoprotecció*, *Controls inicials* i *Registres*), i
per això **ja no surten al pas «Projecte»**: no els has de demanar dues vegades.

Al bloc de **després**, els punts que tenen sub-punts —els certificats
d'inscripció i les inspeccions inicials— et deixen **triar quines
instal·lacions** apliquen.

Si tornes **Enrere**, el que havies marcat (i les dades que havies posat) **es
conserva**.

#### La base de dades de llicències

Els informes d'una llicència van en cadena: primer el requeriment, després el
favorable pre i després el post. El programa **se'n recorda**: quan poses un
**ID GIA** que ja té un informe de llicència fet, t'avisa i **surt tot omplert**
—el que aplicava, si es disposava o no de cada document, els **Id Firmadoc**,
els expedients, les referències, els punts del Projecte, el tècnic redactor i
qui posava condicions—. Ho pots canviar tot; només hi és perquè no ho hagis de tornar
a escriure.

Es desa sol quan generes l'informe, a
`local\base-dades-llicencies\llicencies-db.json`. Com tot el que hi ha a
`local\`, **no puja mai al GitHub**.

Al menú principal, la fila de **Llicència** té un botó **🗂 Dades** al costat
del de `✏️ LLIC`: obre la base de dades, on pots veure el que s'ha desat de
cada activitat, **corregir una dada** que s'hagués entrat malament i esborrar la
fitxa d'una llicència. Esborrar la fitxa no toca cap informe ja generat: només
es perd la memòria per al següent.

Cada fila de la llista té un botó **Generar**: fa l'informe següent d'aquella
activitat sense haver de buscar-la. S'obre l'assistent de sempre (el de
Llicència o el de Modificació No Substancial / Transmissió, segons la fitxa) amb la
fase de l'últim informe ja triada (canvia-la al Pas 1 si ara toca una altra) i,
al Pas 2, l'ID GIA posat i la capçalera omplerta de l'Excel. La resta surt de
la fitxa, com sempre. Si havies canviat alguna dada de la fitxa sense desar-la,
es desa abans d'obrir l'assistent.

Al bloc de **després** de la resolució les caselles surten **totes marcades** (és
com tenies el Word: hi eren totes i n'anaves esborrant); hi ha **«Marcar-ho
tot»** i **«Desmarcar-ho tot»**. Al bloc d'**abans** surten desmarcades, perquè
allà cada punt demana a més dir si ja es disposa de la documentació.

#### Els tres informes són el mateix document

El **favorable pre** i el **post** no són informes a part: són **el mateix
informe sencer** (documentació del projecte, bloc d'abans i bloc de després amb
els seus «Quan:»). L'única cosa que canvia és què diu de cada punt del bloc de
**després**:

| Informe | Sota el «Quan:» |
|---|---|
| Requeriment | res |
| Favorable pre-llicència | res |
| Favorable post-llicència | Es disposa del document (Id Firmadoc: …) — o, si encara falta, **No es disposa de la documentació.** |

Al post, doncs, el bloc de després també et deixa dir de cada punt si es disposa
del document i posar-hi l'**Id Firmadoc**. Al pre no: aquella documentació
encara no toca tenir-la, i cada punt ja diu **quan** s'ha de presentar.

La **documentació del projecte** (el tècnic redactor i els documents signats) va
**dalt de tot**, sota el títol `DOCUMENTACIÓ PROJECTE` i **fora de la
numeració**; la numeració dels punts va seguida de cap a peus.

#### Modificació NO Substancial i Transmissió

Són dos informes **curts**, amb el seu botó al menú. Tenen la capçalera de
Llicència i **només et pregunten una cosa: quins punts de REQ1 hi vols
adjuntar** (la pantalla de sempre). Pots no marcar-ne cap.

- **Amb algun punt** → l'informe diu «…amb les següents observacions:», hi posa
  els punts (amb el mateix format que un requeriment) i, a CONCLUSIONS, el text
  de requeriment.
- **Sense cap** → diu «…sense més observacions en relació a aquest tràmit.», sense
  cap número i sense el «Vist l'anterior, cal requerir l'esmena…».

El fitxer es diu `AAAA-MM-DD_MNS_GIA n.docx` o `AAAA-MM-DD_TRANS_GIA n.docx`.

A la *Modificació NO Substancial* hi ha, a més, la llista de les modificacions
justificades, que hi surt **sempre**.

El text el pots canviar des de l'editor de catàlegs: hi ha un sol document per
als dos, **`MNSTRANS.json`**.

> **D'on surt el text:** de **REQ1**, en viu. `LLIC.json` només hi afegeix el que
> és propi de Llicència (els dos comentaris i el «Quan:»); si canvies un
> requeriment a REQ1, aquí canvia sol. Si algú reanomena un requeriment de REQ1,
> el programa **avisa** que aquell punt s'ha quedat sense text en lloc de
> callar-s'ho.

### Plànol activitats (🏘)

La primera eina de **CARRER**. Fa un plànol de Cornellà amb les **parcel·les on hi
ha activitats** pintades segons el seu estat, per veure d'un cop d'ull on hi ha
activitats legalitzades, on n'hi ha amb coses pendents i on no n'hi ha cap:

- **🔴 vermell** — precintada (a l'Excel d'activitats, o el darrer informe és de
  precinte o cessament).
- **🟡 groc** — el darrer informe és un **requeriment** o una ampliació de termini.
- **🔵 blau** — no té cap informe a la base (o està per revisar): no sabem si està
  legalitzada.
- **🟢 verd** — sense res pendent (favorable, o el requeriment ja s'ha tancat).

**En prémer la rajola**, si ja n'hi ha un de fet, et pregunta si vols
**consultar l'últim** (et diu de quin dia és; s'obre al moment) o **fer-ne un de
nou** (torna a llegir els Excel i pot trigar una estona). Amb l'interruptor en
**A** se'n fa un de nou sol cada setmana, o sigui que normalment n'hi ha prou amb
consultar.

Si en una parcel·la hi ha activitats en estats diferents, la parcel·la té el color
**del més greu**. Clica-hi i surten totes, cada una amb el seu estat.

**D'on surt:** l'Excel d'activitats i l'Excel d'**ESTABLIMENTS** més nous (els dos
a la «Carpeta de l'Excel d'activitats», amb el nom `AAAA-MM-DD ACTIVITATS.xls` i
`AAAA-MM-DD ESTABLIMENTS.xls`), i la **base d'informes** (fes abans
*Actualitzar base*). L'Excel d'establiments és el que diu que una activitat té
**més d'un establiment** (naus, locals contigus): l'activitat surt a totes les
seves parcel·les.

**Els ID GIA**: amplia el mapa i surten damunt de cada parcel·la, a la seva
**coordenada UTM de l'Excel d'activitats** (columnes UTM X i UTM Y). Una
**línia de punts** l'uneix amb el punt de la parcel·la segons el Cadastre: com
més llarga, més lluny ha quedat de la coordenada del Cadastre. Si un ID no
és on toca, corregeix-lo amb l'eina **Coordenades**: quan la correcció s'hagi
importat al GIA i baixis l'Excel nou, el plànol ja el posarà al seu lloc. En
**vermell**, al centre, els que tenen la coordenada fora de la parcel·la o no en
tenen. Al plànol hi ha només els números; clica la parcel·la i a la fitxa hi ha,
per a cada activitat, el seu **local, planta o porta**, l'**adreça de la base
d'activitats** i l'**adreça del Cadastre** (poden no ser iguals). Si l'Excel no
diu el local, el programa ho pregunta al Cadastre (surt marcat «(Cadastre)»), i
si el Cadastre tampoc ho sap, surt el número d'unitat de la referència
cadastral («unitat 0011»).

**Clica un ID** i surt la fitxa d'aquella activitat (amb tots els seus
establiments); **clica la parcel·la** i surten totes les activitats que hi ha. Una
activitat amb diversos establiments surt un sol cop, amb la llista dels seus
locals. Si el GIA té el mateix local dues vegades (un amb l'activitat i un de
buit), no es compta com a local buit: surt a *Per revisar → Local buit duplicat
al GIA* perquè es corregeixi allà.

**Filtres** (a la dreta): per estat, els **locals buits** (en gris; no surten si
no ho marques) i **«Per revisar»**: activitats en un local marcat com a buit (que
potser ja han plegat), activitats que no són a la base d'activitats i activitats
sense cap establiment. Hi ha un cercador per ID GIA, adreça o activitat, i pots
**baixar en CSV** la llista del que es veu.

**La primera vegada triga uns minuts**: ha de demanar al Cadastre el dibuix de
cada parcel·la (unes 950) i la planta/porta dels locals que no la porten. Ho
desa, i les vegades següents ja és qüestió de segons. Si cancel·les, el plànol
es fa igualment amb el que ja tingui (les parcel·les que falten surten com un
punt).

> El plànol és **només del teu ordinador** (`local\planol-activitats\`): porta
> requeriments pendents i noms d'activitats. Al Cadastre només s'hi envia la
> **referència cadastral**.

Si el Cadastre no respon o el plànol surt sense parcel·les dibuixades, fes doble
clic a `suport\rutes\Provar-Planol.bat`: fa una consulta de prova i desa la
resposta.

### Ruta d'inspecció

Botó **📍 Generar ruta**. Escrius els ID d'activitat a visitar i et calcula la
ruta circular més curta des de la base (per defecte Carrer de l'Energia, 97),
amb un mapa numerat que pots imprimir a PDF. Els mapes van a
`local/rutes-generades/`.

> **Privacitat:** al servei de rutes només s'hi envien **coordenades**, mai noms
> ni adreces.

### Coordenades dels establiments

Botó **🗺 Coordenades**, just a la dreta de *Generar ruta*.

El GIA agafa les coordenades del Cadastre, i el Cadastre situa cada activitat al
centre de la seva **parcel·la**, no al local. Per això totes les activitats d'un
mateix edifici cauen exactament al mateix punt: a la base del 18/08/2026, 1.380
activitats només tenien 899 punts diferents, i a Ctra. de l'Hospitalet 147 n'hi
havia **19 apilades**.

**Es treballa per zones, no tot de cop.** En obrir l'eina tries les zones que
vols repassar avui: el municipi va per quadres de 400 m, i de cada zona hi surt
el nom, els seus dos carrers principals i **quantes activitats hi ha**. Amb la
base del 18/08 són 40 zones, la més gran de 40 activitats. Només es consulta el
Cadastre de les zones que hagis marcat.

Al mapa hi veus, de cada activitat:

- **🔴 vermell** — la coordenada que hi ha ara a l'Excel. No es mou.
- **🟢 verd** — la del portal segons l'adreça, unida al vermell per una línia
  grisa. El color et diu d'on surt:
  - **verd** — el portal amb el número exacte de l'activitat.
  - **groc** — hi havia més d'un portal amb aquell número: mira-te'l.
  - **verd clar** — aquell número no hi era: el portal més proper de la parcel·la.
  - **blanc** — no s'ha trobat cap portal. Comença a sobre del vermell i l'has de
    moure tu.

A la llista i a la fitxa de cada punt hi surt el **titular** (la raó social), i
el cercador també el troba. L'adreça de l'activitat surt **sencera** (amb el
bloc, l'escala, el pis i la porta) i a la fitxa hi ha també l'**adreça del
titular**.

**Per mirar un punt sense tocar-lo, passa-hi el ratolí per sobre**: surt una
targeta amb l'ID, el titular, l'adreça, d'on surt el punt i si està validat. Fes
servir això si tens dubtes amb un punt que ja has mogut: un **clic** el
desvalida i el torna al portal (si et passa, **Ctrl+Z** o «Desfer» ho
recupera).

**Quines activitats**: a la finestra de triar zones pots triar **només les
apilades**, **només les NO apilades** o **totes**; **amagar les ja corregides**; o
obrir **només les marcades per revisar** (de totes les zones). Les que ja estaven
corregides a l'Excel (ja no són al punt del Cadastre) surten amb el punt
**lila** i una **línia de punts** fins al punt de la parcel·la al Cadastre, i
tenen un filtre propi a «Mostra:».

**Les ja corregides (lila) es tornen a tocar des d'on són**, no des del portal:
el punt que mous comença on el té ara l'Excel. **Arrossega'l** per corregir-lo, o
**clica'l** i tria: *deixar-la on la té l'Excel*, *tornar-la al punt de la
parcel·la (Cadastre)* o *portar-la al portal*. Un clic a una lila **no la
valida** sola (no se sabria què volies), i tot es pot desfer amb **Ctrl+Z**.

**Triar les zones al plànol**: a la finestra de triar zones, el botó **«Triar-les
al plànol…»** obre el mapa amb la **graella de zones** a sobre (cada una amb les
seves activitats i quantes en queden de pendents). Clica les que vulguis
repassar i **«Fet»**; les pots tornar a triar quan vulguis amb **«Triar zones»**
a la barra de baix, i el mapa se'n recorda la propera vegada. Les zones que ja
hagis marcat a la llista hi surten triades. La primera vegada triga més: ha de
demanar al Cadastre totes les parcel·les (després ja queda desat).

**Punts estranys**: a la fitxa del punt, **«⚠ Marca per revisar»**, amb una nota
si vols. No el valida ni el mou: queda marcat amb un **!** vermell, i al
desplegable «Mostra:» hi ha **«Per revisar»** per tornar-hi després. Les marques
van també a l'Excel del repàs (columna «Per revisar»), però no a l'Excel per
importar.

Si vols veure les parcel·les i els edificis com al Cadastre, marca **«plànol del
Cadastre»** a la barra de dalt: es posa a sobre del fons i es recorda per a la
vegada següent (també al Plànol activitats).

**Amplia el mapa fins que surtin els números dels portals** (a partir del zoom de
carrer). Són tots els portals de la illa, tal com al plànol del Cadastre, també
els que no tenen cap activitat: són el que et deixa dir si un punt és al lloc.

Aleshores, de cada punt verd:

- si ja és on toca, **fes-hi clic** per donar-lo per bo (li surt un anell blau);
- si no, **arrossega'l** fins on toqui (queda taronja);
- si has de validar-ne molts de cop, el botó **«Validar tot el que es veu»** fa
  els que hi hagi a la llista de la dreta (i el cercador la filtra).

Per anar més de pressa:

- **«Següent pendent»** (o la tecla **N**) et porta a la pendent **més propera**
  i n'obre la fitxa: així acabes un edifici abans de passar al següent.
- El desplegable **«Mostra:»** deixa veure només un tipus de punt (els blancs,
  els grocs, els moguts, les pendents...), a la llista **i** al mapa. Al costat
  de cada color de la llegenda hi diu quants n'hi ha.
- L'activitat que tens seleccionada es **ressalta**: la fila en groc, el punt
  verd més gros i amb un anell taronja, i la línia fins al seu vermell en blau.
  Quan n'hi ha moltes juntes, així saps quin vermell va amb quin verd.
- **«Desfer»** (o **Ctrl+Z**) torna enrere l'últim canvi d'un punt: un
  arrossegament, o un clic que l'ha validat o desvalidat. Pots desfer-ne tants
  com vulguis mentre no tanquis la pàgina.

El botó **Baixar Excel (.xlsx)** et dona **tot el que hagis validat d'aquesta
base de dades**, també el d'altres zones i altres dies: ID GIA, referència
cadastral, adreça, zona, coordenada vella, coordenada nova, d'on surt, quants
metres s'ha mogut i de quina base de dades és. La idea és que acabis amb **un
sol fitxer**: el teu **repàs**.

El repàs es recorda en aquest navegador, així que pots tancar i tornar-hi un
altre dia. **Baixa't l'Excel de tant en tant**: és la teva còpia de seguretat. Si
el Chrome perd les dades (les esborres, o canvies d'ordinador), el botó
**«Carregar repàs…»** del mapa el torna a carregar des d'aquell Excel (també si
l'has obert i desat amb l'Excel). El que ja és al navegador no es trepitja: del
fitxer només s'afegeix el que hi falta.

**Per passar les correccions a qui les ha d'importar:** a la finestra de
Coordenades (la de triar zones), botó **«Excel per importar…»**. Tries l'Excel
del repàs i el programa fa una **còpia de la base de dades d'activitats** (el
mateix fitxer, amb el mateix format i totes les columnes) amb les coordenades
corregides escrites **en vermell**. Si alguna activitat ja té a la base una
coordenada diferent de la que hi havia quan la vas repassar (potser algú ja
l'ha corregida), **no es toca** i t'ho diu. La base original no es modifica mai:
la còpia es desa a `local\geocodificacio\` i s'obre sola. Quan hagis abocat les coordenades noves a l'Excel d'activitats i en
generis un de nou, el comptador es buida sol i comences net — i com que aquelles
activitats ja no estaran apilades, **cada tanda serà més curta que l'anterior**.

> L'eina **no toca res**: ni l'Excel, ni els mapes de ruta, ni el plànol
> d'activitats precintades. El que et baixes és un full a part.

> **Privacitat:** al Cadastre només s'hi envia la **referència cadastral**, mai
> el nom del titular ni la raó social.

Si algun dia els punts verds no surten (tots es queden a sobre del vermell), fes
doble clic a **`suport\rutes\Provar-Cadastre.bat`**: fa una sola consulta al
Cadastre i et diu si el servei respon i quants portals n'ha entès. No obre res ni
toca cap fitxer.

### Word a PDF (i signar)

Botó **📄 Word a PDF**. Per defecte hi surt **l'últim informe que has generat**;
pots triar-ne un altre o una carpeta sencera. Converteix a PDF al mateix lloc i,
si marques **«Signar els PDF amb AutoFirma»**, els signa sols, sense clics.

- **Tria el teu certificat al desplegable.** Així la signatura es valida a
  qualsevol ordinador. Si hi deixes «(triar-lo a AutoFirma en signar)», el
  programa t'avisa abans de començar, perquè aleshores només seria vàlida en
  aquest ordinador.
- **Signatura visible**: posa el caixetí a dalt a la dreta de la primera pàgina.
- En acabar et diu quants PDF ha generat, quants ha signat i els errors, si n'hi
  ha. Si alguna vegada cal saber què ha passat amb una signatura, hi ha un
  registre a `local\base-dades-activitats\pdf-signar-log.txt`.

### Normativa (📚)

Botó **📚 Normativa** (secció NORMATIVA). Baixa a la carpeta
**`local\normativa`** el text **vigent** (consolidat) de totes les normes que
cita REQ1 i de les dels teus marcadors de Chrome (unes 165), i les guies i
manuals dels marcadors. Tot va a **una sola
carpeta** i el **nom** del fitxer diu de què és, amb la mateixa classificació
que els marcadors:

```
Vector ambiental_Residus_2016_Decret 197-2016 Comunicació prèvia i registres productors i gestors.pdf
Instal·lacions_Ascensors_2024_RD 355-2024 ITC AEM 1 Ascensores.pdf
Incendis_Antic_2004_RD 2267-2004 RSCIEI anterior (derogat pel RD 164-2025).pdf
```

Ordenat per nom, queden agrupades per àmbit, dins de cada àmbit per tema i dins
de cada tema per any.

**Només hi ha la normativa que es fa servir.** Una norma que **no cita cap punt**
de cap catàleg (ni al text ni a la fitxa ⓘ), o que està **derogada**, surt de la
llista: el seu PDF es mou a la subcarpeta **`derogades`** (no s'esborra mai res)
i ja no surt a l'índex ni a la revisió. Si treus un punt de REQ1 i la seva norma
ja no la cita ningú més, la propera vegada que facis servir l'eina se'n va sola a
`derogades`; si la tornes a citar, torna. Les **guies i manuals** es queden
sempre, encara que no les citi cap punt.

- **La primera vegada** triga una estona (són moltes normes). Les del BOE es
  baixen en PDF. Les del Portal Jurídic, el BOPB i el CIDO es baixen amb el
  **botó «PDF» de la mateixa pàgina**; només si no se'n troba cap, la pàgina es
  desa com a PDF amb l'Edge (sense finestra), i l'índex ho diu.
- **Guies i manuals**: també hi són les guies, manuals, quadres i taules dels
  teus marcadors, al tema **Guies** de cada àmbit
  (`Vector ambiental_Guies_Manual Gestió de residus industrials a Catalunya (ARC).pdf`).
  A l'índex surten com a «Guia» i no com a norma.
- **Es manté al dia sola**: cada vegada que la fas servir, les del BOE es tornen a
  baixar si el BOE n'ha publicat una actualització, i la resta si fa més de mig
  any que es van baixar. **La versió anterior no es perd**: va a la subcarpeta
  `anteriors` amb la data fins a la qual va ser la bona.
- **L'índex** (`0 Index normativa.xlsx`, a la mateixa carpeta) té cada norma amb
  el seu fitxer (clic i s'obre), l'enllaç web, la versió, quan es va baixar, si ha
  anat bé i **a quins punts de REQ1 surt**. Es pot filtrar.
- **A la fitxa d'ajuda (ⓘ) de cada requeriment** hi ha el botó **«Obre el PDF
  desat»** quan la norma ja és a la carpeta.
- Si una norma no té enllaç, a l'índex surt com a **«Sense enllaç: desa-la a mà
  amb aquest nom»**: si la guardes a la carpeta amb aquell nom, el programa ja la
  troba.
- Si alguna falla, l'índex en diu el motiu i es torna a provar la vegada següent.
  Passa-me'l si en veus alguna que no se solucioni.

- **Les ITC de Bombers i les TINSCI** (i les ITC antigues) es baixen **totes**
  de les pàgines d'Interior: el programa en treu la llista de la pàgina cada
  vegada, o sigui que si Interior en publica una de nova, es baixa sola. Van al
  tema `ITC Bombers` / `TINSCI` d'Incendis, amb el nom de l'enllaç
  (`Incendis_ITC Bombers_SP 120 ...pdf`), i l'índex diu quants documents hi ha
  de cada col·lecció.

La llista de normes i la classificació són a `suport/normativa.json`.

### Revisar requeriments (🔍)

Botó **🔍 Revisar requeriments** (secció NORMATIVA). Serveix per tenir el
programa al dia. Tria què vols revisar i prem **Revisar**:

- **Punts de REQ1 sense la fitxa d'informació** (la i): els que has afegit i
  encara no en tenen.
- **Enllaços que no funcionen**, de tots els catàlegs: els del text dels punts i
  els del botó «Obre la norma» de les fitxes.
- **Normativa que ja no és vigent**: mira la pàgina de cada norma **que cita
  algun punt** al BOE i al Portal Jurídic (la que no cita ningú no es revisa: va
  a `derogades`). Si una està derogada, diu **quina la substitueix** i **quins
  punts de REQ1 la citen**.
- **Baixar la normativa nova i les versions noves** (desmarcat per defecte): fa
  el mateix que l'eina Normativa.

En acabar deixa un **Excel a `local\revisions`** amb una fila per cosa a fer:
què passa, on (catàleg i punt), l'enllaç i què cal fer. **No canvia res dels
catàlegs**: canviar un requeriment perquè una norma ha canviat s'ha de fer
llegint-la. Passa'm l'Excel i ho arreglem.

Les normes que no se sap si són vigents surten com a «mira-ho a mà», amb
l'enllaç i **el motiu** (per exemple, que l'Edge no ha pogut obrir la pàgina del
Portal Jurídic).

Si moltes normes del Portal Jurídic surten «no s'ha pogut saber», fes doble clic
a **`suport\Provar-Vigencia.bat`**: desa el que responen el Portal Jurídic i el
BOE per a unes quantes normes en un `.zip` a `local\revisions` (no toca res).
Passa'm el `.zip` i ho arreglem.

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

### Recordatoris als titulars (🔔)

Botó **🔔 Recordatoris** (secció EINES). Envia recordatoris **periòdics** als
titulars que tenen un tràmit pendent, segons l'estat de la base d'informes.

Hi ha **dues campanyes independents**, una a cada pestanya:

| Campanya | A qui escriu |
|---|---|
| **Requeriments** | activitats en estat *Requeriment* |
| **Precintes** | activitats en estat *Precinte / Cessament* |

Cada una té la seva **casella d'activa**, la seva **periodicitat** (cada quants
dies es repeteix a la mateixa activitat), la seva **espera inicial** (dies que es
deixen passar des de l'informe abans del primer avís, perquè el termini encara
corre), el seu **topall per tanda** i el seu **text**, que pots editar amb
*Editar text...*.

El correu ja porta, de sèrie, l'avís que **si ja s'ha presentat la documentació
no cal que en facin cas**, i l'**article 5 de l'Ordenança** (el que diu que no es
pot transmetre l'activitat mentre hi hagi un expedient obert).

**Com fer-ho servir la primera vegada:**

1. Executa **🗃 Actualitzar base** (si no, treballaries amb dades velles).
2. Obre **🔔 Recordatoris** i mira la llista.
3. **Exporta el CSV i revisa'l** — és la manera de veure a qui escriuries
   **sense escriure a ningú**.
4. Posa el topall per tanda a 1 i envia'n un a una adreça teva per veure com
   queda.
5. Quan et convenci, puja el topall i activa la campanya.

> **El límit de correus.** EmailJS només en deixa enviar **200 al mes**. El
> programa en compta **150** i es planta: els altres 50 queden de reserva perquè
> mai et quedis sense. El comptador de dalt de la finestra suma **tots** els
> correus que surten del PC, també els de l'eina *Enviar correu*.

> **EmailJS o l'Outlook.** A **⚙ Configuració** (requadre *Correus que s'envien
> des d'aquest PC*) o al desplegable **Enviar amb:** de la finestra d'*Enviar
> correu* pots triar per on surten els correus d'aquest PC: **EmailJS** (com
> sempre), **Outlook: el deixa a Esborranys** (no s'envia res, és per veure com
> queda) o **Outlook: envia el correu** (surt de la teva bústia, sense límit de
> 200). Cal l'Outlook clàssic; el «nou Outlook» no serveix. Els recordatoris
> **automàtics** van sempre per EmailJS. Si l'Outlook dona problemes (un avís de
> seguretat, o informàtica no ho deixa), torna a EmailJS i ja està.

> **Compte amb la base desfasada.** Si la base d'informes és vella, escriuries a
> gent que ja ha complert. La finestra t'ho avisa en vermell a partir de 30 dies,
> i el mode automàtic directament **no envia res** si passa dels 45.

**Mode automàtic.** Cada campanya pot anar en *Manual* (tu obres l'eina i
cliques *Enviar tanda*) o en *Automàtic*. Per a l'automàtic, el botó
**Automàtic...** crea una tasca del Windows que cada dia a les **13:00** envia el
que toqui; si a aquella hora el PC estava apagat, ho fa en engegar-lo. Només
corre amb la sessió iniciada, i **els correus surten sense que ningú els
revisi**. Si ja la tenies creada d'abans (a les 09:00), el programa la posa al
dia sol en obrir-se. El mateix botó també serveix per esborrar la tasca.

Amb **Excloure / incloure** treus una activitat concreta dels recordatoris (per
exemple, si en portes el seguiment per una altra via).

### Altres botons del menú

- **🗃 Actualitzar base d'informes**: recorre els informes ja fets i n'extreu
  data, ID GIA i conclusió a `local/base-dades-activitats/informes-db.json`.
  També té interruptor **A / M** (vegeu més avall): en **A** es fa sola cada
  dia a les **13:00**.
- **📋 Editar base d'informes**: el que hi corregeixes a mà (la conclusió breu o
  l'«ignorar») **mana** sobre el que surti d'*Actualitzar base*, i l'*Estat
  activitat* d'aquelles activitats surt **en vermell** perquè ho tinguis present.
  Si t'has equivocat, selecciona la fila i clica **Desfer canvi a mà**.
- **📋 Editar base d'informes**, **📥 Revisar entrades del mòbil**,
  **⏱ Controls periòdics**, **Activitats extraordinàries**.

### Com decideix l'estat de cada activitat

L'*Estat activitat* és el que fan servir el **Plànol activitats**, **Comprovar
Excel** i, sobretot, els **Recordatoris**: un *Requeriment* que no ho és és un
correu al titular que no s'havia d'enviar. Per això, l'octubre del 2026 es van
llegir un per un els 802 informes de la carpeta i es va corregir el que el
programa entenia malament. Ara:

- **Cap informe s'ignora sol.** Abans s'ignoraven tots els que començaven per
  «S'informa favorablement», i una llicència informada favorablement es quedava
  en *Requeriment* per un informe de dos anys abans. Ara cada informe té un
  **tipus**, i és el tipus el que diu si decideix l'estat:
  - **Favorable de llicència**: decideix sempre (*Favorable*).
  - **MNS, canvi de nom o de titularitat** informats favorablement: **no tapen**
    un requeriment, un precinte o una ampliació de termini pendents (la MNS pot
    arribar enmig d'un requeriment obert); si no hi ha res pendent, l'estat és
    *Favorable*.
  - **Activitats extraordinàries** (concerts, fires…): si l'acte és dins d'un
    establiment amb GIA (l'estadi, per exemple), **no decideixen l'estat de
    l'establiment**, ni el favorable ni el requeriment. Si l'acte no té GIA,
    l'«activitat» és l'acte, i sí que decideixen.
  - Un informe que hagis marcat *Altres* (notes informatives) només dóna l'estat
    si l'activitat no en té cap altre.
- **El que corregeixes a mà sempre mana**: un informe que ignores no compta mai,
  i el que hi has canviat passa per davant de tot això.
- **Formats d'abans, sense frase de conclusió**: el requeriment antic («S'han
  observat les següents deficiències que cal esmenar…») és *Requeriment*; un
  seguiment punt per punt és *Requeriment* si alguna resposta diu «No s'aporta»,
  «No es justifica»… i *FI Requeriment* si totes són positives (aquest surt a
  revisar amb «estat deduït, sense conclusió», perquè hi donis un cop d'ull).
- **Plantilles sense omplir** (la conclusió diu alhora que es pot i que no es
  pot tancar la denúncia, o hi ha quedat «Copiar requeriment.») surten a revisar
  amb «plantilla sense omplir», en lloc de passar per un requeriment.
- Un informe **sense ID GIA** en una carpeta on tots els altres són del mateix
  GIA va amb aquell GIA. Si el GIA del document no és el de la carpeta, surt a
  revisar («GIA del document diferent del de la carpeta»).
- **La base funciona igual a la feina (I:) que a casa (F:)**: els informes es
  reconeixen pel camí dins de la carpeta d'informes, no per la unitat, i les
  correccions a mà no es perden. Si mai la carpeta d'informes és una altra de
  debò i s'hi haguessin de perdre correccions, *Actualitzar base* t'ho pregunta
  abans (en automàtic no la toca i ho apunta al registre).
- La primera vegada després d'actualitzar el programa, *Actualitzar base* torna
  a llegir **tots** els informes (triga més): les correccions a mà que ara ja
  coincideixen amb el que diu el programa deixen de sortir en vermell.

### El «?» de cada eina

Cada rajola d'eina té a la cantonada de dalt a la dreta un **?** petit. Si hi
cliques, et diu en una o dues frases què fa aquella eina, **sense obrir-la**.

### Sota cada eina, quan la vas fer servir

A totes les eines del menú hi surt, en gris i lletra petita, **l'última vegada
que la vas obrir** (o `(mai)`). Serveix per no haver de recordar si ja havies
passat, per exemple, el *Comprovar Excel* aquesta setmana. Es desa a
`local/base-dades-activitats/eines-state.json` i no es puja mai.

### Copiar informes sol cada dia (l'interruptor **A** / **M**)

> **Per defecte, tot el que es fa sol es fa a les 13:00**: copiar informes,
> actualitzar la base i els recordatoris cada dia, i el plànol d'activitats cada
> dilluns. Ho pots canviar a ⚙ *Configuració → Automatismes* (vegeu més avall). I
> sempre amb la mateixa regla: si l'última vegada que tocava no es va poder fer,
> es fa tan aviat com es pot.

Sota la rajola **📁 Copiar informes**, allà on les altres eines tenen l'hora, hi
ha un **interruptor petit**:

| | Què vol dir |
|---|---|
| **A** (verd) | **Automàtic.** Amb el programa obert, cada dia **a les 13:00** es copien sols els informes nous. Si l'última vegada que tocava no es va poder fer (el programa estava tancat), es fa **en obrir el programa**. |
| **M** (gris) | **Manual.** Només es copia quan cliques la rajola. |

Es canvia clicant-hi al damunt. **La rajola segueix funcionant igual en tots dos
casos**: si la cliques, la còpia es fa al moment, estigui en A o en M.

La còpia automàtica **es fa en segon pla i no es veu res**: cap finestra, cap
pregunta i cap barra de progrés. No pregunta perquè no fa res que es pugui
desfer: copia els informes nous a la carpeta de còpia i **mai esborra res**.

**La data que hi ha al costat de l'interruptor et diu qui va fer l'última
còpia**: si surt **en verd**, la va fer sola; si surt en gris, la vas fer tu. Així
d'un cop d'ull saps si l'automàtic està treballant de debò o només està encès.

> Per posar-lo en **A** cal tenir configurada la **carpeta on copiar els
> informes** (⚙ *Configuració*). Si no hi és, el programa t'ho diu i es queda en
> manual.

> Si mai vols saber què ha fet, hi ha un registre a
> `%LOCALAPPDATA%\InformesCornella\copia-informes-log.txt` amb una línia per
> passada.

### Actualitzar la base sola cada dia (també **A** / **M**)

Sota **🗃 Actualitzar base** hi ha el mateix interruptor. En **A**, la base
d'informes s'actualitza sola cada dia **a les 13:00** (o en obrir el programa, si
l'última vegada que tocava no es va poder fer), en segon pla i sense que es vegi
res. La
data surt en verd si l'última actualització la va fer sola.

- El que hagis corregit a **📋 Editar base** **no es perd**: mana sobre el que
  surti de l'actualització.
- Si tens l'editor obert mentre s'actualitza, en desar s'hi afegeixen els teus
  canvis i el programa t'avisa perquè el tornis a obrir i hi vegis els informes
  nous. Si li toca desar just mentre s'està actualitzant, et dirà que hi tornis
  d'aquí una estona.
- Registre: `%LOCALAPPDATA%\InformesCornella\informes-db-log.txt`.

### El plànol d'activitats sol cada setmana (també **A** / **M**)

Sota **Plànol activitats** hi ha el mateix interruptor. En **A**, el plànol es
torna a fer sol **cada dilluns a les 13:00** (o en obrir el programa, si aquella
setmana encara no s'ha fet), en segon pla i sense que es vegi res, i **es puja
sol al Drive** perquè el mòbil tingui el plànol al dia. Per posar-lo en A cal
tenir configurada la carpeta de la base d'activitats.
Registre: `%LOCALAPPDATA%\InformesCornella\planol-log.txt`.

### Quan es fa cada cosa: ⚙ Configuració → Automatismes

A **⚙ Configuració** hi ha el grup **Automatismes**, amb una fila per cada cosa
que el programa fa sol (Copiar informes, Actualitzar base, Plànol activitats i
Recordatoris). A cada fila pots:

- **encendre-la o apagar-la** (és el mateix interruptor A/M del menú);
- triar **cada dia** o **cada setmana**, i en aquest cas **quin dia**;
- triar **l'hora**.

En desar, el canvi val **de seguida**, sense reiniciar el programa. *Restaura*
torna a posar les hores de fàbrica. Els **Recordatoris** són una tasca del
Windows: si canvies quan es fan, el programa la reescriu sol, però encendre'ls
o apagar-los es continua fent amb el botó *Automàtic...* de la seva eina.

---

## 4. Editar els catàlegs

**Des del programa**, amb el botó **✏ Editar catàlegs**. Hi pots afegir,
esborrar i moure seccions, subseccions, ítems i sub-punts, i escriure'n el text.

**La capçalera i les conclusions** les fan servir *tots* els informes i per això
no pengen de cap botó: hi entres pels dos enllaços que hi ha **al costat de «Què
vols fer?»** de la pantalla principal — **✏ Capçalera** i **✏ Conclusions**. A
dins, cada bloc et diu **a quin tipus d'informe s'aplica**.

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

### La fitxa d'ajuda de cada requeriment (el botó **ⓘ**)

Cada ítem i sub-punt d'un catàleg pot portar una **fitxa d'ajuda**: el criteri
per decidir si allò s'ha de requerir o no, sense haver d'anar a buscar el Word
d'inspeccions. Té **cinc camps**, sempre els mateixos:

| Camp | Què hi va |
|---|---|
| **Norma** | La norma i l'article exactes. |
| **Enllaç** | El **text consolidat** de la norma. A la vista en Word surt com a hipervincle sota la línia de la norma; al Pas 3, com a botó **🔗 Obre la norma**. 220 de les 254 fitxes en tenen; les que no, són ordenances municipals, les OME i normes UNE, que no tenen text consolidat en línia. |
| **Criteri** | El llindar, l'excepció, **quan sí i quan no**. |
| **Vigència** | Si és d'una sola vegada o s'ha d'anar renovant, i cada quant. |
| **A qui s'aplica** | Activitat nova, existent, modificació… És el camp que evita exigir normativa posterior a una activitat legalitzada abans. |
| **Competència** | Qui ho pot exigir. Mitja secció d'Autoritzacions són obligacions d'altres administracions: allà l'Ajuntament **constata i dona trasllat**. |
| **Revisat** | `AAAA-MM`. Serveix per saber quines fitxes s'han quedat endarrerides quan canvia una norma. |

> **Els enllaços es construeixen sols.** Tant el BOE com el Portal Jurídic
> publiquen permalinks **ELI**: `boe.es/eli/es/rd/2017/05/22/513/con` i
> `portaljuridic.gencat.cat/eli/es-ct/d/2023/11/28/209`. Si has d'afegir la
> d'una norma nova, segueix el mateix patró: tipus, any, mes, dia i número.

**On es veu:** al **Pas 3**, amb la icona **ⓘ** que surt al davant dels punts que
en tenen (clic o **F1**); i a la **vista en Word** del catàleg, en gris sota cada
punt. **A l'informe del titular no hi surt mai** — és material de consulta teu.

**Com s'omple:** a **✏ Editar catàlegs**, amb el botó **ⓘ** que hi ha al final
de la barra del cos (passa-hi el ratolí per sobre i t'ho diu). Surt com a
**ⓘ ✓** quan el punt ja en té, i queda apagat als títols de secció i de
subsecció, que no es requereixen. Si esborres tots cinc camps, la fitxa
desapareix i la **ⓘ** deixa de sortir al Pas 3.

> Quan canviï una norma, el que has de tocar és **la fitxa i el text del punt**,
> i posar la data nova a *Revisat*. Les proves comproven que cap fitxa no es
> queda sense norma ni sense criteri, i que la data té el format `AAAA-MM`.

Quan deses, el programa **regenera les vistes en Word** i, en fer
`Actualitzar.bat`, **puja els teus catàlegs al GitHub**. Els teus canvis
sempre manen: si el repositori ha tocat el mateix fitxer, guanya la teva versió.

### La capçalera (`0 CAPCALERA`)

És l'únic document que **segueix sent un Word de veritat**: hi ha l'escut, el
requadre de la «Nota:», les tabulacions i els marges, i això no es pot refer des
de cap fitxer de text. Per això va **a mitges**:

- el **`.docx`** mana en el **format**;
- el **`.json`** mana en el **text** (les etiquetes, els valors fixos i la nota).

Quan obres l'editor, el `.json` es torna a llegir del document —o sigui que mai
et pot dir una cosa diferent del que hi ha—, i quan deses, el text que hi hagis
posat **es torna a escriure dins del `.docx`** (i se'n fa una còpia de seguretat
al costat). No s'hi poden afegir ni treure línies: la capçalera té una
estructura fixa i el que s'edita és **el que hi diu**.

Hi ha **tres blocs**, i l'editor et diu a quin informe va cadascun:

| Bloc | S'aplica a |
|---|---|
| Capçalera general | Requeriment - Nou, Ampliació de termini, Controls periòdics |
| Capçalera d'activitats extraordinàries | Activitats extraordinàries |
| Capçalera de llicència | Llicència, Modificació NO Substancial, Transmissió |

> Els `<<ID_GIA>>`, `<<TITULAR>>`, `<<ADRECA>>`… són els forats que omple el
> programa: **no els esborris**. Una línia «Camp:» sense cap forat surt buida a
> l'informe.

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

Al costat, el botó **📁** obre la **carpeta on es desen els informes generats** —
la que hi ha posada a Configuració, sigui quina sigui. No tanca el menú.

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

En obrir la web del mòbil hi ha dues opcions:

1. **Consultar plànol**: el mateix *Plànol activitats* del PC. Cada vegada que
   el generes al PC, se'n puja una còpia al teu Drive privat (carpeta `Dades`,
   `planol.html`), i el mòbil la llegeix amb el teu compte de Google: **genera'l
   al PC de tant en tant** perquè el del mòbil estigui al dia. A la pantalla del
   telèfon, els filtres s'obren amb el botó **Filtres**. A la fitxa de cada
   activitat hi ha **Fer informe**: obre el formulari amb l'ID GIA ja cercat al
   Pas 2; revisa les dades i prem **Següent**.
2. **Generar informe**: el formulari de sempre.

---

## 9. Per si ho toca algú altre (o Claude)

- El **codi** és a `suport/`. El mapa dels mòduls és a la capçalera de
  `suport/Motor.ps1`; les decisions tècniques i el perquè de cada cosa, a
  `suport/CLAUDE.md` — que remet a tres documents més, a
  `suport/documentacio/`, per a la signatura de PDF, la Llicència i les rutes.
- **Proves**: `pwsh -File suport/tests/run-tests-all.ps1` (les passa totes sis;
  `run-tests.ps1` n'és només una). Amb
  `$env:GENINFORME_TEST = '1'` el motor només defineix funcions (ni finestres ni
  Word), o sigui que la lògica pura es pot provar fins i tot en un Linux.
- **Branca de desplegament**: `main`. `Actualitzar.bat` només hi puja
  **catàlegs i dades**, mai codi.
