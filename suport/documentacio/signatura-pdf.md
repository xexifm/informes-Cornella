# Word a PDF i signatura electrònica

> Ve de `suport/CLAUDE.md`, que s'havia fet massa gros per llegir-lo
> sencer. **Llegeix-lo ABANS de tocar `suport/PdfSignar.ps1` o `suport/PdfCms.ps1`**

Aquesta secció és llarga perquè hi ha SIS rondes de diagnòstic amb l'usuari, i
cada una descarta una hipòtesi. Val la pena llegir-la sencera abans de proposar
res sobre la validesa de la signatura: gairebé tot el que sembla evident ja
està provat i descartat.

- **Word a PDF (i signar)** (`Invoke-ConvertirPdf`, `suport/PdfSignar.ps1`, rajola
  📄 a EINES, acció `convertirpdf`): converteix tots els Word (`.doc`/`.docx`) d'una
  carpeta (i subcarpetes) a **PDF al mateix lloc i mateix nom** (via Word COM
  `ExportAsFixedFormat`, `17`=wdExportFormatPDF); salta els que ja tenen un PDF al
  dia si no es marca "sobreescriure". Opcionalment **signa** cada PDF amb
  **AutoFirma** per línia de comandes (`sign -store windows -format pades`). El
  **certificat es tria d'un desplegable** poblat del magatzem de Windows
  (`Cert:\CurrentUser\My`, amb clau privada); del certificat triat se'n treu el
  CN (`_CertCommonName`) i es passa a AutoFirma com a `-filter subject.contains:<CN>`
  perquè el triï sol, sense diàleg. El PDF signat substitueix el sense signar. Es
  va triar **AutoFirma + magatzem de Windows** perquè reutilitza el certificat que
  l'usuari ja fa servir i és l'eina oficial (signatura vàlida per a
  l'administració). La **tria de carpeta** fa servir el helper comú `_AddConfigRow`
  (`UiComuns.ps1`) — quadre editable + "..." + indicador ✓/⚠ en viu — igual que a
  Configuració (tots els selectors de carpeta del programa fan servir aquest format).
  - **Per defecte hi surt l'ÚLTIM INFORME GENERAT** (`_UltimInformeGenerat`): el
    `.docx`/`.doc` més nou de la carpeta de sortida, sense `-Recurse` (els informes
    es desen plans), saltant els temporals `~$`. El cas d'ús real és sempre el
    mateix — acabes de generar un informe i el vols passar a PDF i signar-lo — i
    abans tocava anar-hi a buscar cada vegada. Si no se'n troba cap (o la unitat de
    xarxa no hi és) es recupera la cascada d'abans: última ruta desada →
    `$InformesDir` → buit. `_CarpetaInformesGenerats` repeteix la cascada de
    `_ResolveOutputDir` (`Motor.ps1`) però **sense crear la carpeta**: obrir un
    diàleg no ha de crear res al disc. Es pot canviar sempre: el quadre és editable
    i hi ha els botons **Carpeta** i **Document**.
  - **Signatura VISIBLE (caixetí)**: en signar, per defecte s'afegeix un **caixetí a
    dalt a la dreta de la pàgina 1** que reprodueix l'aspecte "CERTIFICAT SENSE DNI"
    (nom / càrrec / organisme / data, **sense DNI**). AutoFirma per línia de comandes
    **NO pot triar un "Aspecto" desat de la GUI**, així que es reprodueix amb
    `layer2Text` via `-config` (`_AutoFirmaVisibleExtraParams` → `signaturePage=1`,
    `signaturePositionOnPage...` dalt-dreta, `layer2Text` amb les línies del caixetí
    unides per `\n` LITERAL i el marcador `$$SIGNDATE=yyyy.MM.dd HH:mm:ss$$`). El text
    del caixetí és **editable** al diàleg d'opcions (casella "Signatura visible" +
    quadre de text) i es desa a `pdf-signar-state.json` (`caixeti`, `visibleSign`);
    per defecte `_DefaultCaixeti`. Sense caixetí, `_BuildAutoFirmaSignArgv` es comporta
    com abans (signatura invisible, cap `-config`).
  - **Format de `-config`: propietats separades pel `\n` LITERAL (2 caràcters),
    MAI per salts de línia reals.** Això no és una suposició, és el codi
    d'AutoFirma. `CommandLineParameters.java` guarda el valor **tal qual** (NO el
    descodifica de Base64) i l'ordre `sign` el converteix a `Properties` amb
    `CommandLineLauncher.buildProperties()`, que fa
    `while ((endIndex = params.indexOf("\\n", beginIndex)) != -1)`: en Java el
    literal `"\\n"` són els caràcters `\` + `n` i `indexOf` busca una **cadena
    literal**, no una expressió regular. El mateix codi ho diu al comentari:
    *"La division no funciona correctamente con split porque el caracter salto de
    linea se protege al insertarse por consola, asi que lo hacemos manualmente."*
    (Compte: `loadSignConfig`, unes línies més amunt, sí que fa `split("\n")` amb
    salts reals — però **aquest camí no és el de `sign`**. És exactament la trampa
    en què vam caure.) Amb salts REALS el bucle no troba res i AutoFirma es queda
    amb **una sola propietat**: clau `signaturePage` i com a valor tota la resta →
    **signa bé (codi 0) però sense caixetí**. Al PDF es reconeix perquè el widget
    surt amb `/Rect[0 0 0 0]` i l'aparença amb `/BBox[0 0 0 0]`. Un intent encara
    anterior el passava en **Base64**, amb el mateix resultat. `$Script:AutoFirmaConfigSep`
    conté el separador; posició tunejable a `$Script:AutoFirmaCaixetiPos`.
  - **Codi de sortida 0 NO vol dir caixetí visible**: `_PdfCaixetiEsInvisible`
    mira **només el tros que el signador ha afegit al final** del PDF (la revisió
    incremental, des de la mida del fitxer original) i hi busca `/BBox[0 0 0 0]`
    (`_PdfTextCaixetiInvisible`, pura). Si el caixetí ha quedat invisible, l'intent
    es dona per fallat, es registra i es passa al següent. Sense això el programa
    donava el fitxer per bo i l'usuari no veia res, sense cap avís.
  - **LES COMETES LES POSEM NOSALTRES (2n error, també real)**: a PowerShell 5.1
    `Start-Process -ArgumentList @(...)` **NO enquota** els elements: els ajunta
    amb espais. Amb rutes com `5.- Sergi Fadurdo`, AutoFirma rebia la ruta
    **tallada** al primer espai i responia *"El fichero de entrada no existe:
    I:\…\5.-"*. Ara `_ArgvToCommandLine` (pura) construeix la línia i enquota, i
    s'executa amb **`ProcessStartInfo`** (`_RunAutoFirma`), que dona control
    exacte de la línia d'ordres i recull sortida i codi de sortida. Les barres
    invertides **finals** es dupliquen quan l'argument va entre cometes: si no,
    la barra escaparia la cometa de tancament (`CommandLineToArgvW`).
  - **La data la resolem NOSALTRES (3r error real)**: amb
    `$$SIGNDATE=yyyy.MM.dd HH:mm:ss$$` dins de `layer2Text`, AutoFirma petava amb
    *"Error no reconocido: begin 0, end -1, length 21"* (un `substring` amb un
    índex no trobat) i no signava. `_ResolveCaixetiDate` (pura) substitueix el
    marcador per la data abans de cridar AutoFirma i **treu qualsevol altre
    `$$...$$`**, de manera que AutoFirma no veu mai cap marcador. A la interfície
    el marcador es manté (l'usuari pot triar el format de data).
  - **EL `layer2Text` NOMÉS ADMET UNA LÍNIA (4t error real, ja confirmat)**, i ara
    se n'entén el perquè: el `\n` literal és el **separador de propietats**, de
    manera que un salt dins del `layer2Text` fa que AutoFirma hi talli i que el
    tros següent (sense cap `=`) faci petar
    `keyValue.substring(0, keyValue.indexOf('='))` amb `indexOf` = −1. És
    exactament l'error del registre: *"Error no reconocido: begin 0, end -1,
    length 21"*, i **21 són els caràcters de `Enginyer d'Activitats`**, la 2a línia
    del caixetí. Per això `_AutoFirmaVisibleExtraParams` en mode `text` **sempre**
    posa el caixetí en una sola línia (`_CaixetiUnaLinia`, unit amb ` · `, que
    també talla pel `\n` literal i treu qualsevol barra invertida que quedi).
  - **Per tenir-lo de DIVERSES LÍNIES: com a IMATGE**. `_BuildCaixetiImageBase64`
    (System.Drawing) dibuixa el caixetí amb la mateixa proporció que el requadre
    de la signatura i el passa a `signatureRubricImage` (base64), que és el
    mecanisme documentat d'AutoFirma per a la rúbrica. En aquest mode NO s'hi
    posa `layer2Text` (la imatge ja ho porta tot).
  - **LA RÚBRICA HA D'ANAR EN JPEG. NO EN PNG.** Es va provar de passar-la en PNG
    (amb text pla ocupa molt menys: 16.244 caràcters contra els 35.000 llargs del
    JPEG equivalent) i AutoFirma la va rebutjar amb
    *"Se ha proporcionado una imagen de rubrica que no esta codificada en JPEG"*.
    L'intent amb imatge falla i el caixetí cau al respatller de **text d'una
    línia** — que és com es nota, perquè no surt cap error a la pantalla. Hi ha
    prova que ho blinda (cap entrada de `$Script:CaixetiIntents` pot ser d'un
    altre format).
  - **La lletra es veia BORROSA i era la QUALITAT del JPEG**, no el format:
    s'hi anava a **qualitat 70** i **escala x2** (144 ppp), i el JPEG a qualitat
    baixa deixa halos al voltant del text negre sobre blanc. Ara
    `$Script:CaixetiIntents` és una escala d'intents **de més a menys qualitat**
    (x3 q92 → q88 → q84 → q78 → x2 q95 → q90) i es fa servir **el primer que hi
    càpiga**: així s'aprofita tot el pressupost en lloc d'anar a la fixa amb el
    pitjor. El codificador JPEG del Windows no dona la mateixa mida que cap
    altre, per això hi ha tants graons.
  - **`MaxCaixetiBase64` era massa just** (20000, posat "per si de cas"): amb
    l'escut de fons només hi cabia l'escala x2, justament la borrosa. **Mesurat
    amb una ordre real del registre** (rutes a `I:\…\5.- Sergi Fadurdo\…`, filtre
    amb el CN sencer i les 6 propietats de posició), tot el que **no** és la
    imatge ocupa **628 caràcters** — o sigui que del límit dur de Windows (32767)
    en sobren més de 30.000. Ara `MaxCommandLine` = 32000 i el topall de la
    imatge, 30500. La comprovació de `MaxCommandLine` segueix sent la xarxa de
    seguretat si una ruta fos molt més llarga.
  - **El registre diu quina variant d'imatge ha entrat** (`$Script:CaixetiUltimIntent`
    → línia `imatge: jpeg x3 q88 (28… car.)`), i si no n'hi cap cap, les mides de
    totes les provades. Sense això, quan el caixetí sortia en text calia endevinar
    si havia estat la mida o el format.
  - **Aspecte** (`$Script:CaixetiEstil`, tot tunejable en un sol lloc): contorn
    gris fosc, **escut de l'Ajuntament de fons a la dreta** (pintat amb
    `ColorMatrix.Matrix33` = opacitat; va **abans** del text perquè aquest hi
    passi per sobre) i `FactorLletra` = 0,72 (era 0,58): com més alt, més omplen
    les lletres la línia i **menys espai buit** queda entremig. El requadre ha
    passat de **75 a 48 pt** d'alçada per al mateix nombre de línies.
  - **`Icon.ToBitmap()` NO serveix amb `suport/cornella.ico`**: un `.ico` és un
    contenidor amb diverses mides a dins, i aquest les porta **totes set
    comprimides en PNG** (16…256 px). El .NET, amb icones així, no les
    descomprimeix bé i el resultat surt **buit** — l'escut no apareixia al
    caixetí i **no ho deia ningú**, perquè el dibuix va dins d'un `try/catch`.
    `_IcoTriaFrame` (pura, amb proves contra el fitxer real) llegeix la taula del
    `.ico` — capçalera de 6 bytes i una entrada de 16 per imatge, amb l'amplada
    en **un sol byte** on el 0 vol dir 256 — i en tria la més petita que ja sigui
    prou gran; després el PNG va a `Image.FromStream`. Queda el respatller a
    `Icon(path, Size)` per si algun dia el `.ico` porta imatges DIB.
    `$Script:CaixetiAvisEscut` / `CaixetiEscutDibuixat` ho deixen dit al
    registre: **un `try/catch` que s'empassa un error de dibuix ha de deixar
    rastre**, si no la cosa surt malament en silenci.
  - **Resolució: mesura-la, no la suposis.** Del registre de l'usuari surt que
    l'escala x3 amb qualitat 92 ocupa **26.416 caràcters**, mentre que les meves
    proves n'estimaven 33.192: el codificador JPEG del Windows fa la imatge un
    **20% més petita**. Amb això calibrat, l'escala **x4 (800×192, 288 ppp) hi
    cap fins a qualitat 85**, i amb TEXT més resolució val més que més qualitat.
    La imatge que hi ha dins d'un PDF ja signat es pot treure i mirar
    (`/DCTDecode`) — és la manera de saber què hi ha arribat de debò en lloc de
    discutir sobre una captura de pantalla.
  - **LÍMIT DE LA LÍNIA D'ORDRES (5è error real)**: la imatge viatja en **base64
    dins de l'ordre**, i Windows no admet més de **32767** caràcters. Amb la
    imatge a escala x4, `Process.Start` petava amb *"El nombre del archivo o la
    extensión es demasiado largo"*. Ara la imatge va a **escala x2 amb JPEG de
    qualitat 70** i hi ha **dos topalls**: `$Script:MaxCaixetiBase64` (la imatge
    no es genera si passa de 20000 car.) i `$Script:MaxCommandLine` (un intent que
    no hi cabria **ni es prova**, se salta i es registra).
  - **El `try/catch` va DINS del bucle d'intents**: quan estava a fora, una
    excepció d'un intent s'enduia **tots** els altres i el fitxer es quedava
    **sense signar**. Va passar exactament això amb la imatge massa gran.
  - **LA FIRMA NO ES VALIDAVA ALS ALTRES ORDINADORS** (mesurat comparant els dos
    PDF, no deduït). Al PC de l'usuari sortia vàlida; als dels companys **i al
    visor corporatiu de l'Ajuntament** (`aytos-fdocweb`), *"La validez de la
    firma es DESCONOCIDA / la identidad del firmante es desconocida"*. Un informe
    signat a mà **amb el mateix certificat** sí que s'hi validava.
    - Es va treure el CMS de dins del `/Contents` dels dos PDF i comparar-los:
      **el certificat signant és EL MATEIX** (mateix número de sèrie
      `49160A73…`, TCAT d'empleat públic, emès per `SubCA SECTOR PUBLIC Q (G3)
      A.1`). O sigui que **no era el certificat**, ni caducat ni res.
    - Diferències reals: el nostre anava amb `SubFilter` **`ETSI.CAdES.detached`**
      i **3 certificats a dins** (arrel + subCA + signant), amb
      `signingCertificateV2` (els OID de política que s'hi veuen són **del
      certificat**, dins d'aquell atribut: **no** hi ha cap política de firma —
      les preferències d'AutoFirma diuen «Ninguna política»); el que funciona,
      **`adbe.pkcs7.detached`** i **només 1** (el signant). **Cap dels dos** porta
      segell de temps ni dades de revocació (el `revocationInfoArchival` del que
      funciona és un SET **buit**).
    - **Per què es creu que és això**: l'avís d'Adobe diu que no són de confiança
      *"sus certificados PRINCIPALES"* — i els pares només els té perquè **els hi
      encastem nosaltres**. Si l'ancoratge de confiança d'aquell ordinador no és
      l'arrel de l'AOC sinó la **subCA** (que és el que publica la Llista de
      Confiança europea), donar-li la cadena sencera el fa aturar-se en una arrel
      que no coneix. Que falli **també al visor corporatiu** —que és de servidor i
      té la llista de confiança ben posada— apunta al **fitxer**, no a cada PC.
    - `_AutoFirmaCompatLines` (pura) hi posa `signatureSubFilter=adbe.pkcs7.detached`
      i `includeOnlySignningCertificate=true`. **El nom porta DUES ENES**
      (`Signning`): és així al Client @firma i, si algú l'hi "corregeix",
      AutoFirma l'ignora **en silenci** i tornem a encastar la cadena. Hi ha prova.
    - Aquestes línies van al `-config` **sempre**, amb caixetí o sense: abans el
      `-config` només existia si hi havia caixetí i una signatura invisible no
      se n'hauria beneficiat.
    - Per tornar enrere: `$Script:SignaturaCompat`, i prou.
    - **3a ronda, amb els DOS PDF del MATEIX ordinador** (el signat a mà surt
      vàlid, el nostre desconegut): això **retira definitivament** la teoria del
      magatzem de confiança de l'altre PC — és el fitxer. I els dos canvis
      anteriors **sí que van entrar** (el PDF ja porta `adbe.pkcs7.detached` i
      un sol certificat), o sigui que **tampoc no eren la causa**.
    - **L'OID de l'algorisme tampoc.** Al `SignerInfo` nosaltres emetíem
      `rsaEncryption` i Adobe `sha256WithRSAEncryption`. Es va poder provar
      **sense tornar a signar**, perquè aquell camp **no està cobert per la
      signatura**: es va canviar **1 sol byte** del PDF ja signat (l'OID passa de
      `…01 01 01` a `…01 01 0B`, mateixa llargada, `/ByteRange` intacte),
      l'OpenSSL seguia dient *CMS Verification successful*… i l'Adobe seguia
      dient desconegut. **Aquest truc del byte és reutilitzable** per provar
      qualsevol cosa que estigui fora dels `signedAttrs`.
    - **L'única diferència que queda: `signingCertificateV2`** (atribut ESS que
      AutoFirma posa i Adobe no). El seu `certHash` és **correcte** (SHA-256 del
      certificat, comprovat), però l'atribut hi porta a dins un camp
      **`policies`** amb les polítiques del certificat — i el **RFC 5035** diu
      que això **no és informatiu**: obliga a validar la cadena **restringida a
      aquelles polítiques**. Al PC de l'usuari, amb tota la cadena de l'AOC
      instal·lada, passa; en un altre, no. Això explica per què encastar o no la
      cadena no va canviar res: el que falla no és **trobar** els pares, és
      **validar-los sota aquelles polítiques**.
    - `signingCertificateV2=false` demana la versió antiga de l'atribut
      (SigningCertificate, RFC 2634), que **no porta `policies`**. És l'últim
      recurs abans de muntar la signatura nosaltres.
    - **`signingCertificateV2=false` tampoc.** Provat: segueix desconeguda.
      S'han acabat els paràmetres d'AutoFirma.
    - **LA SOLUCIÓ: `suport/PdfCms.ps1`.** En lloc de seguir perseguint
      diferències d'una en una, es deixa que **AutoFirma munti el PDF** (que això
      ho fa bé: el document mai surt alterat, el camp de signatura i el caixetí
      són correctes) i **se li reemplaça NOMÉS el CMS de dins del `/Contents`**
      per un fet aquí amb `SignedCms` de .NET, amb l'estructura **exacta**
      d'Adobe: `contentType` + `messageDigest` + `adbe-revocationInfoArchival`
      buit, un sol certificat i **cap atribut ESS**.
      - **Per què és segur**: no es toca ni un byte del document. El
        `/ByteRange` no canvia (i és el que diu on és el forat, així que no cal
        endevinar on és el `<`), el forat que deixa AutoFirma és de **27.000
        bytes** i el nostre CMS n'ocupa ~2.600, i el que es firma són
        **exactament** els mateixos bytes → el *"no ha habido modificaciones"*
        continua sortint igual. La mida del fitxer **no pot canviar**: si
        canviés, el `/ByteRange` (que ja està escrit i firmat) deixaria de
        quadrar.
      - Cal l'**empremta** del certificat (`CertThumb` a les opcions), no només
        el filtre de CN que fa servir AutoFirma: sense triar-ne un al
        desplegable no es pot refer i es queda la signatura d'AutoFirma (i el
        registre ho diu).
      - Si el reempaquetat falla, **es deixa la d'AutoFirma**: val més una
        signatura que només es validi al PC de l'usuari que cap signatura.
      - **`CmsSigner` de .NET, dos paranys**: (1) per defecte fa **SHA-1**, s'ha
        de posar SHA-256 explícitament; (2) si no se li posa **cap** atribut
        signat, firma el contingut directament i el PDF queda **sense
        `messageDigest`**, que és el que Adobe espera. Per això s'hi posa
        l'`adbe-revocationInfoArchival` buit — el mateix que hi posa Adobe — i
        llavors .NET hi afegeix sol el `contentType` i el `messageDigest`.
      - `CryptographicAttributeObject` **no** és a `...Cryptography.Pkcs` sinó a
        `System.Security.Cryptography`. L'assemblatge és `System.Security` al
        PowerShell 5.1 i `System.Security.Cryptography.Pkcs` al 7 (les proves):
        `_CmsCarregaTipus` prova els dos.
      - **Es va escriure una funció que "corregia" l'OID de l'algorisme del
        SignerInfo i la prova la va enxampar corrompent el CERTIFICAT**: el patró
        que buscava (`SEQUENCE{rsaEncryption, NULL}`) també surt a la clau
        pública del certificat, i el del SignerInfo no porta el `NULL`, així que
        la cerca "l'última aparició" queia sobre el certificat. Com que ja
        s'havia demostrat que aquell OID **no canvia la validesa**, la funció es
        va esborrar. Codi que no cal, fora.
    - **4a ronda: l'usuari va reportar «tampoc» amb el reempaquetat ja publicat,
      SENSE registre** — o sigui que no se sap si la signatura refeta es va
      arribar a aplicar (sense certificat triat al desplegable no es pot refer,
      i una excepció al `SignedCms` queia en silenci al respatller). Resposta:
      **tancar l'ambigüitat**, no provar més coses a cegues:
      - **Res no s'escriu sense comprovar-ho** (`_CmsComprova`): el CMS nou es
        descodifica, `CheckSignature($true)` contra el contingut del PDF, 1
        certificat, `messageDigest` present, **cap atribut ESS**. Si falla, el
        PDF es queda com estava i el motiu surt al registre **i al resum**.
      - **L'OID del SignerInfo ara també s'iguala** (`_CmsOidComAdobe`), amb
        guarda POSICIONAL: dins d'un SignerInfo l'algorisme va DESPRÉS dels
        atributs signats, o sigui que l'última aparició de `rsaEncryption` només
        és la bona si queda **després de l'últim `messageDigest`** (la del
        certificat queda sempre abans). El .NET l'escriu `30 0B` sense
        paràmetres i l'Adobe `30 0D` amb NULL: tots dos legals, i el NULL no es
        pot inserir sense re-encodar longituds — es deixa sense.
      - **El resum ho diu ben gros**: «N amb la signatura REFETA I COMPROVADA» o
        «ATENCIÓ: N s'han quedat amb la d'AutoFirma (només vàlida aquí)». I si
        es vol signar **sense** certificat triat, s'avisa **abans** de començar.
      - La suite genera un **certificat efímer en memòria** (`CertificateRequest`)
        i prova el cicle sencer: crear → igualar OID → comprovar → PDF sintètic
        → rellegir i verificar. El camí que corre a Windows és el provat.
    - **TRAMPA de PowerShell (test enxampat)**: un literal hex com a argument
      (`AssertEq $x 0x30`) arriba com a 48 però el PSObject **conserva el text
      del token**, i el `[string]` de dins d'`AssertEq` en fa `"0x30"` → la
      comparació falla mentre el missatge mostra «48» contra «48». La
      interpolació `"$x"` dóna "48"; el cast `[string]$x` dóna "0x30". Als
      arguments de test, els valors esperats **en decimal**. (Mateixa família
      que el PSObject del `Join-Path`.)
    - **5a ronda — LA FI DE LES DIFERÈNCIES DE FITXER.** L'usuari va enviar el
      PDF nou (resum: «REFETA I COMPROVADA») que als PCs dels companys seguia
      sortint desconegut. Es va analitzar contra el vàlid de l'Adobe:
      - certificat del signant **byte a byte idèntic** (mateix SHA-256);
      - CMS del mateix **exacte** nombre de bytes (2.629) i **esquelet ASN.1
        idèntic node a node (156/156)**: les úniques diferències són el hash del
        document i la firma, que canvien per força;
      - **pyHanko** (validador PAdES independent, amb l'arrel de l'AOC com a
        àncora): els DOS fitxers `INTACT:TRUSTED,UNTOUCHED`, cobertura
        `ENTIRE_FILE`, modificacions `NONE`.
      **Conclusió: ja no queda cap diferència de fitxer.** El que difereix és la
      CONFIANÇA dels ordinadors que miren: l'Adobe, per defecte, **només es refia
      de les seves llistes AATL/EUTL** (documentat per Adobe), no del magatzem de
      Windows. El següent element de diagnòstic és la pestanya **Confianza** del
      certificat a l'Adobe d'un company (diu la FONT de la confiança) amb els dos
      fitxers costat a costat.
    - **`suport/Confiar-certificats-AOC.bat`**: la solució de desplegament — la
      mateixa que documenten el BOE i els ministeris per als seus PDF. Porta
      **incrustats** els dos certificats PÚBLICS de l'AOC (arrel `ROOT-A` +
      `SubCA SECTOR PUBLIC Q (G3) A.1`, verificats per hash contra els del CMS
      real) i els instal·la amb `certutil -user -addstore` (sense administrador);
      després cal marcar la Integració amb Windows a l'Adobe (o «Agregar a
      certificados de confianza» des del propi PDF). ASCII pur (els `.bat` amb
      accents es trenquen amb les codepages).
    - **6a ronda — el requisit real de l'usuari**: la firma ha de sortir vàlida
      **també fora de la feina**, sense que cap receptor instal·li res. L'únic
      artefacte que ho compleix EMPÍRICAMENT és el PDF signat pel **propi
      Adobe**. Per això el diàleg té ara **dos modes de signatura**
      (`SignMode` a l'estat/opcions; per defecte `'adobe'`):
      - **`adobe`**: l'eina obre cada PDF a l'Adobe (`_TrobaAdobeExe` /
        `_AdobeExeCandidats`, rutes en text pla — `Join-Path` peta fora de
        Windows), l'usuari signa a mà, i en dir «Sí» es **comprova** que hi hagi
        firma de debò (`_PdfTrobaFirma`); si no n'hi ha (desada amb un altre
        nom!), s'avisa i es compta a part (`$senseFirmaAdobe`). `Cancel·la`
        atura la resta. El certificat/caixetí/AutoFirma del diàleg es
        desactiven en aquest mode.
      - **`autofirma`**: tot el pipeline d'abans (AutoFirma + repack + caixetí),
        intacte i encara provat per la suite. Vàlid on es confiï en l'AOC.
      La qüestió de fons (per què el mateix CMS byte-idèntic es fia en un fitxer
      i en l'altre no en aquells ordinadors) segueix SENSE explicació mesurada:
      la pestanya **Confianza** de l'Adobe d'un company amb els dos fitxers és
      la dada que falta. No barrejar les dues coses: el mode `adobe` és el camí
      pràctic, no la resposta al misteri.
    - **Geometria del diàleg**: `$Script:PdfDlgAmple` (550) i els botons
      col·locats a partir de `$y` (`$yBotons`), amb la `ClientSize` calculada al
      final. Abans anaven clavats a `y=438` i, en afegir-hi els dos radios del
      mode, **van quedar fora de la finestra**; i el botó «Document» de
      `_AddConfigRow` (que acaba a **x=536**) ja sortia tallat amb els 510 px
      d'amplada. Hi ha prova de FONT que ho vigila — acotada al tros del diàleg
      d'opcions, perquè la finestra de **progrés** sí que és de mida fixa.
    - **Pendent**: el que faria la firma validable a qualsevol banda i d'aquí a
      anys és un **segell de temps (TSA)** + dades de revocació (PAdES-LTV).
      AutoFirma ho admet (`tsaURL`, `tsaPolicy`, `tsaHashAlgorithm`…), però cal
      una TSA a què l'Ajuntament doni accés.
  - **L'ASPECTE del caixetí està a `$Script:CaixetiAspecte`** i ara val
    `'defecte'`: només se li diu **on** va (les coordenades de sempre) i el
    dibuixa AutoFirma, que és el mateix que surt amb l'eina *"Utilizar un
    certificado"* de l'Adobe. Amb `'propi'` torna el caixetí nostre (lletra,
    interlineat, contorn i escut de fons): **no s'ha esborrat res**, tot el codi
    hi és i hi ha proves que ho comproven. Es va canviar a petició de l'usuari
    mentre es perseguia el problema de validesa; **que quedi dit: l'aspecte no
    pot canviar la validesa** — el que es valida és el certificat i la seva
    cadena, no el dibuix.
  - **Reintents escalats**: caixetí **(imatge)** → caixetí **(text d'una línia)**
    → **sense** caixetí. Així mai es queda un PDF sense signar i del registre se'n
    dedueix quin ha funcionat. Al registre la imatge surt **resumida** (mida en
    caràcters), mai el base64 sencer, que faria el fitxer inservible.
  - **El registre distingeix els salts**: `_AutoFirmaArgvToText` mostra els salts
    de línia REALS com a `<LF>` i deixa els `\n` LITERALS tal qual. Abans tots
    dos sortien com a `\n` i el log no permetia saber quin era quin — que és
    justament el que calia per depurar.
  - **ELS QUATRE VALORS DE POSICIÓ HAN DE SER ENTERS.** AutoFirma llegeix
    `signaturePositionOnPage*` com a **nombres enters**: amb un `324.48` es
    queda sense la configuració del caixetí i **la signatura surt INVISIBLE**,
    sense cap error ni cap codi de sortida diferent. Va passar de debò en fer
    que la dreta sortís del marge del text (`595,276 − 70,8 = 524,476`): el
    caixetí va desaparèixer i el programa no ho va dir. Mig punt no es veu, o
    sigui que arrodonir no costa res; el que costava era el decimal. Hi ha
    prova que ho vigila —sobre el mapa i sobre les línies que es passen de
    debò—, validada tornant a posar-hi els decimals.
  - **Posició: ALINEADA amb la capçalera de l'informe, i els números són
    MESURATS**, no posats a ull. Es descomprimeix el flux de contingut de la
    pàgina 1 d'un informe ja generat i se'n treu que: la imatge del logo es
    col·loca amb el seu dalt a `y=806,52` **però porta 18 px de blanc a dalt**
    (de 199) = 6,51 pt, o sigui que la punta de l'escut és a **`y=800`**; i el
    requadre de la «Nota:» va de `x=85,58` a **`x=552,45`**, que és el marge dret
    del text. D'aquí `AutoFirmaCaixetiPos` = 352/752/552/800. Hi ha proves que
    lliguen el caixetí a aquestes dues referències.
  - **Registre de diagnòstic**: cada execució desa a `pdf-signar-log.txt` (al costat
    de `pdf-signar-state.json`) **l'ordre exacta** passada a AutoFirma, el codi de
    sortida i la seva sortida (`_PdfSignarLog`, `_AutoFirmaArgvToText`). S'obre
    (o no) segons la casella **«Obrir el registre de la signatura en acabar»** del
    diàleg d'opcions, que es recorda a `pdf-signar-state.json` (`obrirRegistre`,
    per defecte **no**). Abans es preguntava en acabar cada vegada, i preguntar
    sempre una cosa que és de diagnòstic fa nosa. Serveix per no haver d'endevinar si el caixetí no surt.
    **Va al `.gitignore`**: conté les RUTES COMPLETES dels informes i les carpetes
    porten el nom i l'adreça del titular — i aquest repositori és PÚBLIC. (S'hi
    ignora pel nom exacte, no `*.txt`, perquè el `README.txt` d'aquella carpeta sí
    que va al repositori.)
  - **Carpeta O document**: el quadre accepta una **carpeta** (tots els Word de dins
    i subcarpetes) o **un sol document** Word; hi ha dos botons (Carpeta / Document).
    `_RunConvertPdf` ho distingeix amb `Test-Path -PathType Leaf`.
  Funcions pures testejables (`_PdfPathForDoc`, `_PdfShouldConvert`, `_CertFilterValue`,
  `_CertCommonName`, `_BuildAutoFirmaSignArgv`, `_AutoFirmaVisibleExtraParams`,
  `_AutoFirmaArgvToText`, `_DefaultCaixeti`, `_AutoFirmaCandidatePaths` — retorna un
  **array pla**, no `,$ArrayList`, perquè `@()` l'enumeri bé). Word (COM) i AutoFirma
  només a Windows. Opcions/estat a `pdf-signar-state.json`.
