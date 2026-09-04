# Rutes, coordenades i el plànol públic

> Ve de `suport/CLAUDE.md`, que s'havia fet massa gros per llegir-lo
> sencer. **Llegeix-lo ABANS de tocar `suport/rutes/`**

## Eina «Coordenades» — Excel vs façana (`rutes/Coordenades.ps1` + `rutes/Geocodificador.ps1`)
- **El problema, mesurat** (base del 18/08/2026): el GIA porta les coordenades
  del Cadastre, i el Cadastre georeferencia la **parcel·la**, no el local. 1.380
  activitats → **899 punts diferents**; 711 files (52%) en comparteixen un amb
  alguna altra. De les 898 parcel·les, **227 tenen més d'una activitat** i
  concentren aquelles 711 files. Pitjor cas: `4091106DF2749A` (Ctra. de
  l'Hospitalet 147) amb **19 apilades**.
- **La coordenada verda NO surt de geocodificar el text de l'adreça.** Surt de
  la `Ref. cadastral` que ja hi ha a l'Excel: els **14 primers caràcters** són la
  parcel·la, i al Cadastre se li demanen els **portals** d'aquella parcel·la
  (servei INSPIRE d'Adreces, `wfsAD.aspx`, consulta desada `GetadByRefcat`). Cada
  portal és un punt d'**entrada** amb el seu número de carrer. Avantatges: **una
  consulta per parcel·la** (no per activitat) i la resposta ja ve en **EPSG:25831**,
  el mateix sistema que l'Excel — cap reprojecció, cap error de conversió.
- **NO toca Ruta.ps1 ni Precintades.ps1.** Va ser petició explícita de l'usuari:
  aquells segueixen amb la coordenada original. L'eina només MIRA i genera un
  fitxer.
- **Tres xarxes de seguretat, i totes tres hi són a posta:**
  1. `Resolve-CoordEstabliment` descarta qualsevol portal a més de
     `$GeoDistanciaMaximaM` (250 m) de la parcel·la. Val més un marcador imprecís
     que un marcador mentider.
  2. `ConvertFrom-CatastroAdXml` parseja **per `local-name()`**, sense lligar-se a
     cap espai de noms ni nivell de l'arbre, i gira els eixos si venen a l'inrevés
     (en UTM 31N l'est ~420.000 va molt per sota del nord ~4.578.000).
  3. Res del mòdul llança mai: si el servei no respon, la parcel·la queda sense
     portals i cada activitat es queda amb la seva coordenada de sempre.
- **`$GeoCatastroUrlTemplate` és una VARIABLE, no una cadena enterrada al codi**:
  si el Cadastre canvia el nom del paràmetre de la consulta desada, s'arregla des
  de `config.ps1` sense tocar el programa. Per això `Coordenades.ps1` carrega
  `Geocodificador.ps1` **ABANS** de `Ruta.ps1` — que és qui carrega `config.ps1`:
  si es carregués després, els valors per defecte del mòdul trepitjarien el que
  l'usuari hagués posat a `config.ps1`.
- **EL CLIENT DE XARXA NO S'HA POGUT PROVAR CONTRA EL SERVEI REAL.** L'entorn on
  es va escriure tenia `ovc.catastro.meh.es` bloquejat per política de sortida
  (403 al CONNECT; també ICGC, Cartociudad i Nominatim). La fixture
  `tests/dades/wfsAD-exemple.xml` està **muntada a mà** seguint l'esquema INSPIRE,
  no gravada. Abans de fiar-se'n, a la feina:
  ```
  suport\rutes\Provar-Cadastre.bat        (doble clic; accepta una refcat com a argument)
  ```
  Ha de llistar els portals de Cadis i Huelva amb els seus números.
  **NO cridis `. Ruta.ps1` a pèl** per fer-ho: `Ruta.ps1` executa la seva `Main`
  i t'obre el planificador de rutes (va passar). `Coordenades.ps1` en mode
  headless ja ho carrega tot sense obrir res.
  `Test-Geocodificador` compta les adreces de la resposta **crua** (per
  `regex`, sense parsejar) i les compara amb les que ha entès el parseig: així
  es distingeix «el servei no ha tornat res» de «no n'he sabut treure res». I
  desa **sempre** la resposta sencera a `local/geocodificacio/resposta-<rc>.xml`,
  que és l'única cosa que permet arreglar el parseig sense anar a les palpentes.
- **`Coordenades.ps1` PORTA BOM I L'HA DE PORTAR.** Tot el text que l'usuari veu
  al mapa (llegenda, popups, capçaleres de l'Excel que es baixa) viu dins del
  here-string de `Build-CoordenadesHtml`, en català i amb accents. Sense BOM, el
  Windows PowerShell 5.1 llegeix el fitxer com a ANSI i el mapa surt ple de
  `Ã§`. `Geocodificador.ps1`, en canvi, és ASCII pur i no en porta (com
  `Precintades.ps1`).
- **Dins del here-string `@"…"@` no hi pot haver cap `$` ni cap `` ` `` que no
  sigui una interpolació volguda**: el JavaScript del mapa està escrit
  expressament sense `$` ni template literals. Si hi afegeixes codi, comprova-ho
  (`$` dins del here-string = variable de PowerShell).
- **L'`.xlsx` el genera el NAVEGADOR, sense cap biblioteca.** Un `.xlsx` és un ZIP
  amb cinc XML a dins; amb el mètode «sense compressió» només cal el CRC-32 i les
  capçaleres del ZIP (`crc32`/`zipStore`/`buildXlsx` al mateix HTML). Els textos
  van **inline** (`t="inlineStr"`), així no cal `sharedStrings.xml`. Verificat:
  el fitxer generat el valida `zipfile` i l'obre `openpyxl` **sense avisos**, amb
  números com a números i accents intactes. Sense el `<cellStyles>` a
  `styles.xml`, `openpyxl` es queixa («no default style»).
- **`latLonToUtm31` (al JS) és la INVERSA de `Convert-UtmToLatLon`** i cal perquè
  Leaflet dona graus quan s'arrossega un punt i nosaltres hem d'exportar metres.
  Comprovada d'anada i tornada sobre 525 punts de tot el terme municipal: error
  màxim **0,07 mm**. Una coordenada que **no** s'ha mogut a mà s'exporta amb els
  metres **tal com van arribar** (`utmActual`), sense reprojectar: així no s'hi
  acumula l'error d'anar i tornar.
- **`estat[]` va per POSICIÓ dins d'`ITEMS`, no per ID**: si algun dia la base
  portés dos cops el mateix ID Activitat, dues fitxes es trepitjarien. Al
  `localStorage`, en canvi, es desa **per ID**, que és el que ha de sobreviure
  quan es torni a generar el mapa.
- **Es va provar el mapa SENCER en un navegador de debò** (Chromium + Playwright,
  amb un doble de Leaflet perquè el CDN estava bloquejat): estat inicial,
  arrossegament, desat i recuperació al navegador, esborrat de correccions,
  filtre, i el `.xlsx` baixat rellegit amb `openpyxl`. El **SRI** dels `<script>`
  de Leaflet bloqueja qualsevol doble: a la còpia de prova s'ha de treure
  l'`integrity` (mai al fitxer de veritat).
- **LA TRAMPA DEL `return ,@(...)`, TERCERA APARICIÓ — i la primera crida real
  al Cadastre la va destapar.** `ConvertFrom-CatastroAdXml`, `Get-RegistresApilats`
  i `Get-RefcatsAConsultar` acabaven amb `return ,@(...)` **i** totes les crides
  les embolcallaven amb `@()`. Les dues coses juntes hi posen la capa **dues
  vegades**: `$portals.Count` valia **1**, `$p.Numero` feia enumeració de
  membres i el diagnòstic escrivia `numero='System.Object[]'`. El servei
  funcionava perfectament; el que fallava era el consum.
  - **Tria una convenció i escriu-la al costat de la funció.** Aquí és **array
    pla + `@()` al lloc de la crida** (com `_AutoFirmaCandidatePaths`), no la de
    `_FindCampInfoPairs` (`,@(...)` consumit **sense** `@()`). Les dues són
    correctes; barrejar-les no.
  - **Les proves ho haurien enxampat** (`AssertEq $portals.Count 5`,
    `AssertEq @(ConvertFrom-CatastroAdXml '').Count 0`), però no s'havien pogut
    executar: a l'entorn no hi ha `pwsh`, i la rèplica en Python **no modela
    aquesta semàntica** — retorna llistes planes i el problema no hi existeix.
    Lliçó: una rèplica en Python valida la LÒGICA, mai les trampes del llenguatge.
- **EL DIAGNÒSTIC ÉS UN `.bat`, i ho és per dues rascades seguides.**
  `suport/rutes/Provar-Cadastre.bat`, doble clic. Els dos intents d'escriure la
  comanda a mà van fallar tots dos:
  1. `. Ruta.ps1` a pèl → `Ruta.ps1` executa la seva `Main` al final i **obre el
     planificador de rutes**. L'usuari es va trobar la finestra oberta sense
     saber què fer.
  2. `powershell -NoProfile -Command "$env:COORDENADES_TEST=1; …"` llançat
     **des d'un PowerShell** → el shell de FORA expandeix `$env:` (buit) abans
     de passar-ho i al de dins li arriba `=1; …` → *«El término '=1' no se
     reconoce»*. Dins de cometes dobles, el `$` és del shell exterior.
  Al `.bat` la variable la posa el `cmd` i a la línia de PowerShell **no hi ha
  cap `$`**. Si algú ja té un PowerShell obert a l'arrel, el que sí que va és
  `$env:COORDENADES_TEST=1; . .\suport\rutes\Coordenades.ps1; Test-Geocodificador '…'`
  (sense embolcallar-ho en un altre `powershell`).
  Regla general: **una comanda de diagnòstic que s'ha d'escriure a mà amb
  cometes niuades no és una comanda de diagnòstic, és un `.bat` que falta.**
- **`Find-HeaderColumn` ha passat de `Precintades.ps1` a `Ruta.ps1`**: és
  utillatge comú de `rutes/` i ara la fan servir dos fitxers. Tot va a dot-source
  al mateix àmbit, o sigui que Precintades la segueix veient.
- La memòria cau dels portals viu a `local/geocodificacio/portals.json` (clau
  `Geocodificacio` a `$Script:LocalSubdirs`) i **es desa cada 25 parcel·les
  noves**: si es cancel·la a mitja tanda, no es perd el que ja s'ha demanat. Les
  entrades amb portals valen 365 dies; les buides, 30 (per si el servei era
  caigut). Si la crida **falla**, no s'hi escriu res.
- **LES SIGLES DE VIA DEL CADASTRE HI HAN DE SER TOTES.** El Cadastre escriu
  `CL CADIS`, i a `Get-ViaNormalitzada` hi faltava **`CL`** (hi havia `CALLE`,
  `CARRER` i `C`, però la `C` demana un espai al darrere i a `CL` la segueix una
  `L`). Resultat: `'CL CADIS' ≠ 'CADIS'`, **cap adreça no ha coincidit mai per
  carrer** i la tria del portal es feia només pel número. En una illa amb
  entrades per dos carrers això vol dir agafar el número del carrer del costat:
  `C HUELVA 1` va acabar a 129 m d'on tocava. La llista viu a
  `$Script:GeoSiglesVia`; si un dia surten desplaçaments estranys, mira primer
  si hi ha una sigla nova. Va passar desapercebut perquè **no falla, empitjora
  en silenci**: seguia trobant un portal, només que el que no era.
- **Dos portals amb el mateix número existeixen.** A la illa de Cadis n'hi ha
  dos amb l'1, i un d'ells cau exactament al centre de la parcel·la. Quan passa,
  `Select-PortalFacana` retorna `facana-dubtosa`: al mapa surt en ambre perquè
  l'usuari el miri, en lloc de triar-ne un a l'atzar i callar.
- **El guard del JSON ha de mirar la SORTIDA, no el `Count`.** El
  `if ($arr.Count -eq 1) { "[$json]" }` dona per fet que `ConvertTo-Json`
  desembolcalla quan hi ha un sol element — i **en el PowerShell de l'usuari no
  ho fa**: sortia `[[{…}]]` i el mapa d'una sola activitat no arrencava. Ara es
  mira `StartsWith('[')`. **`Ruta.ps1` tenia el mateix patró** i s'ha arreglat
  igual (amb una sola parada hauria petat exactament igual).
- **La graella de zones s'ancora a un origen CONSTANT** (`$CoordZonaX0` /
  `$CoordZonaY0`), no al mínim de les dades. Si sortís de les dades, n'hi hauria
  prou que una activitat nova caigués més a l'oest perquè **totes** les zones es
  desplacessin i «la zona C6» volgués dir una altra cosa que la setmana passada.
  400 m: 40 zones, la més gran de 40 activitats i la mediana de 17.
- **Els noms de les zones surten de les dades, no del codi.** `Get-CarrersDominants`
  els calcula dels carrers de les activitats que hi cauen: al codi no hi ha
  escrit cap nom de cap carrer de Cornellà, i per tant no es pot desfasar.
- **`Test-CoordPlausible` i per què cal.** La base porta coordenades impossibles
  (el GIA 1009 té `X=423,37`: les xifres bones dividides per mil). Si es colen, el
  mapa s'estira fins a l'Atlàntic i la resta de punts queden tots en un píxel. Es
  descarten i es llisten al resum del final, mai en silenci.
- **El progrés del repàs viu al NAVEGADOR, i el PowerShell no hi té accés.** Per
  això la finestra de tria diu quantes activitats té cada zona però **no** quantes
  en portes de repassades: això ho diu el mapa, que sí que pot llegir el
  `localStorage`. No intentis posar-ho a la finestra sense una via real de
  retorn del navegador al disc.
- **El `localStorage` desa la FILA SENCERA**, no només la posició: l'Excel ha de
  poder portar tot el que s'hagi validat d'aquesta base, també el de les zones
  que avui no estan obertes. La clau és el **nom del fitxer d'origen**, de manera
  que en canviar de base d'activitats el repàs es buida sol — que és el que toca,
  perquè les correccions ja seran a dins de la base nova.
- **Proves**: `tests/run-tests-coordenades.ps1` (registrada a `run-tests-all.ps1`).
  A l'entorn de desenvolupament no hi havia `pwsh` (ni paquet ni GitHub), o sigui
  que la lògica es va validar amb una **rèplica en Python** — el mateix recurs
  que ja s'havia fet servir en aquest projecte — i les suites de PowerShell les
  ha d'executar l'usuari a Windows.


## Plànol públic d'activitats precintades
- `suport/rutes/Precintades.ps1` genera `docs/dades/precintades.json` a partir
  de l'Excel d'activitats (fulla "Estès"): les activitats amb el camp lliure
  "PRECINTE ACTIVITAT?" i valor que comença per "SI". La pàgina pública
  `docs/precintades.html` (GitHub Pages) el llegeix i pinta el mapa (Leaflet).
- Ho refresca i puja a `main` **`Actualitzar.bat`** (pas 7). URL pública:
  `https://xexifm.github.io/informes-Cornella/precintades.html`.
- **Privadesa**: el JSON només conté activitat genèrica (p.ex. "BAR"), adreça de
  l'establiment, ID intern i coordenades — **mai** la raó social ni el text
  lliure del Valor (que conté noms i tràmits interns). No hi afegeixis dades
  personals: aquesta pàgina és pública.
- **Etiquetes del plànol**: tant a `precintades.html` com al mapa de ruta
  (`Ruta.ps1`, `Build-RouteHtml`) els marcadors mostren l'**ID Activitat (GIA)**
  en una "pastilla" (no un número correlatiu; l'amplada s'ajusta als dígits). A
  la ruta, la BASE conserva el "0" i s'afegeixen unes quantes **fletxes de
  sentit** (~9, `addRouteArrows`) orientades al traçat. El GIA ja era públic al
  JSON i al popup, així que no hi ha cap dada nova exposada.
- Reutilitza les funcions de `Ruta.ps1` carregant-lo en mode headless
  (`RUTA_TEST`); si canvies `Ruta.ps1`, executa també
  `run-tests-precintades.ps1`.

