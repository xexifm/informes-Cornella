#requires -Version 5.1
<#
.SYNOPSIS
  Que diu cada informe i quin estat en surt per a l'activitat ("Actualitzar
  base", "Editar base", Recordatoris, Comprovar Excel, Planol activitats).

.DESCRIPTION
  Vivia a Informes.ps1. El 7 d'octubre de 2026 es van llegir un per un els 802
  informes de la carpeta real (427 activitats) i es van comparar amb el que en
  treia "Actualitzar base": 110 informes en 'Revisar', 82 activitats sense estat
  util i, pitjor, activitats amb l'estat EQUIVOCAT (una llicencia informada
  favorablement s'ignorava per defecte i l'activitat es quedava en
  'Requeriment' per un informe de dos anys abans). L'estat alimenta els
  Recordatoris (correus als titulars): un 'Requeriment' fals es un correu que no
  s'havia d'enviar. Corregir-ho va fer creixer prou aquesta part perque
  Informes.ps1 passes de les 1.200 linies.

  Tot el que hi ha aqui es PUR (nomes text i objectes): es prova en headless
  amb textos inventats (la classificacio real te dades personals i el
  repositori es public; es valida en local amb ValidarClassificacio.ps1).

  NOMES DEFINEIX FUNCIONS.
#>

# La versio del classificador. Es desa a informes-db.json (versio_classificador)
# i, si la de la base no es aquesta, "Actualitzar base" torna a llegir TOTS els
# informes, no nomes els modificats: sense aixo, una millora del classificador
# nomes arribaria als informes que algu tornes a desar, i els 802 de la carpeta
# es quedarien amb la classificacio vella per sempre. CANVIA-LA cada vegada que
# canviis el que en surt (conclusio, conclusio breu, tipus).
$Script:ClassificadorVersio = '2026-10-08.3'

# Una propietat d'un objecte de la base (o $null si no la te), sense petar amb
# les bases d'abans que no porten els camps nous. La fan servir aquest fitxer,
# Informes.ps1 i InformesEscaneig.ps1: viu aqui, a baix de tot, perque cap dels
# tres no depengui d'un altre en cercle.
function _PropInf($o, [string]$nom) {
    if ($null -eq $o -or $null -eq $o.PSObject.Properties[$nom]) { return $null }
    return $o.$nom
}

# Normalitza un numero d'expedient per comparar-lo (l'informe fa servir "/", la
# carpeta "-", i l'Excel pot portar zeros al davant): parteix en grups i treu els
# zeros inicials de cada grup numeric. "2025/1/2563" i "2025/01/2563" -> "2025-1-2563".
function _NormalitzaExpedient($s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return '' }
    $groups = ([string]$s).Trim() -split '[^\dA-Za-z]+' | Where-Object { $_ -ne '' }
    $norm = $groups | ForEach-Object {
        if ($_ -match '^\d+$') { [string][int]$_ } else { $_.ToUpper() }
    }
    return ($norm -join '-')
}

# Normalitzacio per comparar frases: sense accents, minuscules i SENSE
# apostrofs. Els informes fan servir l'apostrof TIPOGRAFIC (U+2019), pero les
# nostres frases de referencia el recte (U+0027); traient-los tots dos (i altres
# variants) la comparacio casa igual. Fem servir codepoints [char]0x.... per no
# dependre de l'encoding amb que PowerShell 5.1 llegeix aquest fitxer.
function _ConclNorm($s) {
    $t = _NormalitzaText $s
    $apos = @([char]0x0027, [char]0x2018, [char]0x2019, [char]0x02BC, [char]0x00B4, [char]0x0060)
    foreach ($a in $apos) { $t = $t.Replace([string]$a, '') }
    return $t
}

# Frases que poden marcar l'INICI de la conclusio d'un informe: cada familia de
# tramits tanca la decisio d'una manera diferent (vist a la carpeta real).
#
# 'Segona' = nomes si cap de les altres hi es. Son frases que tambe poden sortir
# al COS (un seguiment que copia la conclusio d'un informe anterior, "es
# pertinent precintar" dins d'un requeriment...): la conclusio comenca a la
# PRIMERA linia que en conte una, i si una d'aquestes passes davant de la
# "Vist l'anterior" de debo, la conclusio s'enduria mig informe. En segona
# passada no poden canviar res del que ja es trobava abans.
#
# 'Font' ja no decideix res de l'estat (abans 'mns' i 'act_extr' feien
# l'informe "ignorat" per defecte): ho fa el TIPUS (_TipusInforme), que mira el
# text sencer de la conclusio. Es desa per saber d'on ha sortit.
$Script:ConclusioStartPhrases = @(
    [pscustomobject]@{ Font = 'vist_anterior'; Segona = $false; Phrase = "Vist l'anterior" },
    [pscustomobject]@{ Font = 'risc';          Segona = $false; Phrase = 'Tenint en consideració el risc' },
    [pscustomobject]@{ Font = 'mns';           Segona = $false; Phrase = "S'informa favorablement" },
    [pscustomobject]@{ Font = 'act_extr';      Segona = $false; Phrase = "El titular és responsable d'executar" },
    [pscustomobject]@{ Font = 'act_extr';      Segona = $false; Phrase = "L'organitzador és responsable d'executar" },
    # La conclusio de LLIC (requeriment de llicencia): sense "Vist l'anterior".
    [pscustomobject]@{ Font = 'requeriment';   Segona = $false; Phrase = "Cal requerir l'esmena de les deficiències indicades" },
    # Suspensions i precintes sense la frase literal del risc ("Tenint en
    # consideracio l'incompliment greu...", "Es ratifica que...", "A l'haver-se
    # exhaurit el termini...").
    [pscustomobject]@{ Font = 'risc';          Segona = $true;  Phrase = 'és pertinent suspendre' },
    [pscustomobject]@{ Font = 'risc';          Segona = $true;  Phrase = 'és pertinent precintar' },
    # Formats d'abans: control periodic conforme i favorables "a l'antiga".
    [pscustomobject]@{ Font = 'favorable';     Segona = $true;  Phrase = 'el Control Periòdic és FAVORABLE' },
    # Les frases de DECISIO favorable competeixen amb les de primera fila i mana
    # la que surt primer al document. Eren de segona, i el concert que deia
    # "S'informa amb caracter favorable" i, mes avall, "El titular es responsable
    # d'executar..." comencava la conclusio a la segona: el text que en quedava
    # no deia "favorable" i sortia 'Revisar' (8/10/2026).
    [pscustomobject]@{ Font = 'favorable';     Segona = $false; Phrase = 'informo favorablement' },
    [pscustomobject]@{ Font = 'favorable';     Segona = $false; Phrase = "s'informa amb caràcter favorable" },
    [pscustomobject]@{ Font = 'termini';       Segona = $true;  Phrase = "estimar la sol·licitud d'ampliació" },
    # Favorables i FI sense la frase habitual (formats d'abans: denuncies i
    # controls del Decret 112/2010). De segona: "no s'aprecia cap irregularitat"
    # tambe pot sortir al cos d'un requeriment.
    [pscustomobject]@{ Font = 'favorable';     Segona = $true;  Phrase = "s'equipara a resultat favorable" },
    [pscustomobject]@{ Font = 'favorable';     Segona = $true;  Phrase = 'COMPLEIX el que estableix el Decret 112/2010' },
    [pscustomobject]@{ Font = 'fi';            Segona = $true;  Phrase = "no s'aprecia cap irregularitat" },
    [pscustomobject]@{ Font = 'fi';            Segona = $true;  Phrase = 'no se li poden requerir més mesures' },
    [pscustomobject]@{ Font = 'fi';            Segona = $true;  Phrase = "no provenen de l'empresa" }
)

# Les OBLIGACIONS que llista el cos d'un requeriment ("S'ha de presentar...",
# "S'haura d'entregar...", "Cal justificar..."), ja normalitzades (_ConclNorm:
# sense accents ni apostrofs, o sigui "sha de presentar", "sha daportar"). Un
# informe que en porta i no diu cap decisio es un requeriment: 97 dels 802
# informes no tenen frase de conclusio, i els requeriments d'abans (i el de les
# activitats extraordinaries d'avui, que no en te) sortien a 'Revisar'.
$Script:RxObligacio = '\b(s?ha|s?haura|s?hauran)\s+d(e\s+)?(presentar|justificar|realitzar|aportar|entregar|acreditar|obtenir|retirar)\b|\bcal\s+(justificar|presentar|aportar|acreditar)\b'

# Conclusio: des del primer paragraf que conte una de $Script:ConclusioStartPhrases
# fins (exclos) el que marca el tancament de l'informe (signatura). Uneix els
# paragrafs amb un espai. Retorna un objecte { Text; Font }: Text es el text
# ORIGINAL (no normalitzat), '' si no es troba cap frase d'inici coneguda;
# Font indica quina frase ha disparat la deteccio. Les 'Segona' nomes es miren
# si no hi ha cap de les altres (vegeu $Script:ConclusioStartPhrases).
function _ExtractConclusio($lines) {
    $lines = @($lines)
    $endPhrases  = @(
        (_ConclNorm 'Ho poso al seu coneixement'),
        (_ConclNorm 'Cornella de Llobregat,'),
        (_ConclNorm "S'informa als efectes oportuns,"),
        (_ConclNorm 'A Cornella de Llobregat, en la data')
    )
    $norm = @($lines | ForEach-Object { _ConclNorm $_ })
    $start = -1
    $font = ''
    foreach ($segona in @($false, $true)) {
        $starts = @($Script:ConclusioStartPhrases | Where-Object { [bool]$_.Segona -eq $segona } | ForEach-Object {
            [pscustomobject]@{ Font = $_.Font; Norm = (_ConclNorm $_.Phrase) }
        })
        for ($i = 0; $i -lt $norm.Count; $i++) {
            foreach ($sp in $starts) {
                if ($norm[$i].Contains($sp.Norm)) { $start = $i; $font = $sp.Font; break }
            }
            if ($start -ge 0) { break }
        }
        if ($start -ge 0) { break }
    }
    if ($start -lt 0) { return [pscustomobject]@{ Text = ''; Font = '' } }
    # La linia d'INICI no pot ser el final: la transmissio diu "...presentada a
    # l'Ajuntament de Cornella de Llobregat, s'informa FAVORABLEMENT..." i el
    # "Cornella de Llobregat," de la signatura la tallava abans de comencar (totes
    # les transmissions sortien "sense conclusio"; vist passant el classificador
    # pels fitxers d'or, octubre 2026).
    $parts = New-Object System.Collections.ArrayList
    for ($i = $start; $i -lt $norm.Count; $i++) {
        $ln = [string]$lines[$i]
        $isEnd = $false
        if ($i -gt $start) { foreach ($ep in $endPhrases) { if ($norm[$i].Contains($ep)) { $isEnd = $true; break } } }
        if ($isEnd) { break }
        if (-not [string]::IsNullOrWhiteSpace($ln)) { [void]$parts.Add($ln.Trim()) }
    }
    return [pscustomobject]@{ Text = ($parts -join ' '); Font = $font }
}

# Opcions valides de "conclusio breu" (l'estat en que queda l'activitat
# despres d'aquell informe). 'Altres' es nomes manual: el classificador
# automatic no la torna mai. 'Revisar' es el que torna quan no hi ha prou
# senyal.
# CADA ESTAT TE NOM, i la llista es fa amb els noms. Abans la llista era l'unic
# lloc on sortien, pero els que la resta del programa necessita anomenar
# ('Requeriment' i 'Precinte / Cessament') es tornaven a escriure A MA a
# Informes.ps1, Recordatoris.ps1, ComprovarExcel.ps1 i rutes/PlanolDades.ps1.
# Aquells quatre comparen contra estat_actual: si algun dia es reanomena un
# estat, el Where-Object no troba res i l'eina diu "no hi ha cap activitat" en
# lloc de petar. Es el defecte tipic d'aqui: no falla, calla.
$Script:EstatRequeriment   = 'Requeriment'
$Script:EstatFiRequeriment = 'FI Requeriment'
$Script:EstatPrecinte      = 'Precinte / Cessament'
$Script:EstatFiPrecinte    = 'FI Precinte / Cessament'
$Script:EstatFavorable     = 'Favorable'
$Script:EstatAmpliacio     = 'Ampliaci' + [char]0x00F3 + ' termini'
$Script:EstatSenseEfecte   = 'Sense efecte'
$Script:EstatAltres        = 'Altres'
$Script:EstatRevisar       = 'Revisar'

$Script:ConclusioBreuOpcions = @(
    $Script:EstatRequeriment,
    $Script:EstatFiRequeriment,
    $Script:EstatPrecinte,
    $Script:EstatFiPrecinte,
    $Script:EstatFavorable,
    $Script:EstatAmpliacio,
    $Script:EstatSenseEfecte,
    $Script:EstatAltres,
    $Script:EstatRevisar
)

# Els estats que deixen alguna cosa PENDENT a l'activitat. Una MNS favorable no
# els tapa (vegeu _InformeQueDeterminaEstat).
$Script:EstatsPendents = @($Script:EstatRequeriment, $Script:EstatPrecinte, $Script:EstatAmpliacio)

# "es pertinent suspendre/precintar" no negat ("no es pertinent..." no compta).
$Script:RxPrecinte = '(?<!\bno (es )?)pertinent (suspendre|precintar)'

# Hi ha un precinte o una suspensio DE DEBO, i no nomes l'ADVERTIMENT? $n ja
# normalitzat (_ConclNorm). Es mira FRASE A FRASE: "es pertinent
# suspendre/precintar" nomes es condicional si, dins de la MATEIXA frase, va
# precedit de "En cas contrari", "Si es disposen de mes elements..." o "Si es
# detecta", o seguit de "en el cas de no presentar". Una afirmacio directa ("es
# pertinent suspendre l'activitat fins a esmenar les deficiencies / fins que
# hagi obtingut la llicencia / fins a aportar la documentacio", "es pertinent
# precintar la cuina fins a...") es un precinte encara que despres hi hagi
# "Vist l'anterior, cal requerir...". Abans n'hi havia prou amb que el text NO
# digues "cas contrari" en algun lloc, i el "cal requerir" de despres guanyava:
# sis activitats de la classificacio real del 7/10/2026 sortien en
# 'Requeriment' amb l'activitat suspesa.
# El text TANCA o FINALITZA l'expedient, o desprecinta ("es pot donar per
# tancada la denuncia / per finalitzat", "es pot aixecar / desprecintar", "es
# pertinent desprecintar"), sense un "no" al davant. $n ja normalitzat.
function _TextTancament([string]$n) {
    return ($n -match '(?<!\bno )es pot donar.{0,12}(finalitzat|tancad)' -or
            $n -match '(?<!\bno )es (pot|valora) (aixecar|desprecintar)' -or
            $n -match '(?<!\bno (es )?)pertinent desprecintar')
}

function _PrecinteEfectiu([string]$n) {
    $tancament = _TextTancament $n
    foreach ($frase in ($n -split '(?<=[.;])\s+')) {
        foreach ($m in [regex]::Matches($frase, $Script:RxPrecinte)) {
            $abans = $frase.Substring(0, $m.Index)
            $despres = $frase.Substring($m.Index + $m.Length)
            # "Si es detecta un us de la cuina ESTANT PRECINTADA, ...es pertinent
            # suspendre": la condicio parla d'un precinte que JA hi es, o sigui
            # que es un precinte vigent i no l'advertiment (8/10/2026). PERO no si
            # el text tanca l'expedient: el tancament de la denuncia copia
            # l'advertiment i despres diu "No s'ha detectat us de la cuina
            # durant el precintament... es pot donar per tancada la denuncia";
            # aquella regla, aqui al pas 0, li passava per davant i el feia
            # Precinte (validacio del 8/10/2026, 19:48).
            if (-not $tancament -and $abans -match 'si es detecta' -and $abans -match '(?<!\bno )(estant|esta|estiguin?|ja) precintad') { return $true }
            if ($abans -match 'en cas contrari|si es disposen? de mes elements|si es detecta') { continue }
            if ($despres -match 'en (el )?cas de no presentar') { continue }
            return $true
        }
    }
    return $false
}

# Classifica el text de la CONCLUSIO (ja extreta per _ExtractConclusio) en un
# dels $Script:ConclusioBreuOpcions, mirant les frases reals amb que Sergi tanca
# cada tipus de tramit. 'Revisar' quan no hi ha conclusio o no es reconeix cap
# frase (inclou "desfavorable", deliberadament: no es vol confondre amb
# "Favorable"). Funcio PURA (nomes text).
#
# L'ORDRE MANA, i cada bloc va davant del seguent per un cas real que fallava:
#   0. "Es deixa sense efecte (la comunicacio)" i el precinte o la suspensio
#      EFECTIUS (_PrecinteEfectiu) van davant de TOT: el text sovint continua
#      amb les deficiencies i un "Vist l'anterior, cal requerir...", i allo no
#      treu que la comunicacio quedi anul·lada o l'activitat suspesa.
#   1. El que deixa l'expedient PENDENT encara que la frase digui "es pot donar
#      per tancada la denuncia", "desprecintar" o "favorablement": "...pero NO
#      donar per finalitzat el procediment d'esmena" (i el "ni donar per
#      finalitzat", que es el mateix amb el NO oblidat), "D'altra banda, es
#      requereix...", "es valora favorablement la solucio... s'hauran de...",
#      el control periodic "FAVORABLE... incorrecte havent de ser DESFAVORABLE",
#      i la conclusio de requeriment del cataleg d'avui ("cal requerir
#      l'esmena", sense termini), que queia a 'Revisar'.
#   2. S'inicia el procediment d'esmena: el precinte o la retirada que
#      l'acompanyen son l'ADVERTIMENT ("En cas contrari es pertinent
#      precintar", "es pertinent que es retiri"), com el "determini el
#      cessament" de sempre (un precinte efectiu ja ha sortit al pas 0). Va DAVANT del "es pot donar per tancada la denuncia":
#      "tanca la denuncia i inicia el procediment d'esmena" es un requeriment.
#   3. Els FI (finalitzat, denuncia tancada, desprecintar/aixecar).
#   4. Un precinte que nomes es l'advertiment -> Requeriment (DESPRES dels FI:
#      "es pot aixecar el precinte. Si es detecta..., es pertinent precintar"
#      es un FI). Despres, el risc sense "pertinent precintar", ampliacio,
#      favorable, i les clausules d'un requeriment nou; l'ultima, una obligacio
#      ($Script:RxObligacio) sense cap altra decisio.
function _ConclusioBreu($text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return 'Revisar' }
    $n = _ConclNorm $text

    # 0. Comunicacio anul·lada, i precinte o suspensio de debo.
    if ($n -match '(?<!\bno (es )?)deixa sense efecte') { return $Script:EstatSenseEfecte }
    if (_PrecinteEfectiu $n) { return $Script:EstatPrecinte }
    # 1. Pendent, digui el que digui la resta de la frase.
    if ($n -match 'no s.?han esmenat' -or $n.Contains('no es pot donar') -or
        $n -match '\b(no|ni) donar per (finalitzat|tancat)') { return 'Requeriment' }
    if ($n.Contains('cal requerir lesmena')) { return 'Requeriment' }
    if ($n -match 'daltra banda,?\s*es requereix') { return 'Requeriment' }
    if ($n.Contains('resultat favorable incorrecte')) { return 'Requeriment' }
    if ($n.Contains('valora favorablement la solucio') -and $n.Contains('shauran de')) { return 'Requeriment' }
    # "el Control Periodic es FAVORABLE. De totes maneres, s'ha de presentar la
    # seguent documentacio": el favorable no tanca res (8/10/2026).
    if ($n -match 'de totes maneres' -and $n -match $Script:RxObligacio) { return 'Requeriment' }
    # 2. Inici del procediment d'esmena.
    if ($n -match 'inicia (dofici )?el procediment desmena') { return 'Requeriment' }
    # 3. Seguiment resolt (inclou denuncies tancades: mateix "final positiu").
    if ($n -match 'es pot donar.{0,12}finalitzat') { return 'FI Requeriment' }
    if ($n.Contains('es pot donar per tancada la denuncia')) { return 'FI Requeriment' }
    # Seguiments i denuncies resolts sense la frase de sempre (8/10/2026).
    if ($n.Contains('no saprecia cap irregularitat') -or $n.Contains('no se li poden requerir mes mesures') -or
        $n -match 'molesties.{0,60}no provenen de') { return 'FI Requeriment' }
    # Aixecament d'un precinte/suspensio.
    if ($n -match 'es (pot|valora) (aixecar|desprecintar)' -or $n.Contains('pertinent desprecintar')) { return 'FI Precinte / Cessament' }
    # 4. El precinte o la suspensio que queden son NOMES l'advertiment (els de
    # debo ja han sortit al pas 0): es un requeriment.
    if ($n -match $Script:RxPrecinte) { return $Script:EstatRequeriment }
    # Risc greu/imminent sense "es pertinent precintar", o el cessament ordenat.
    if ($n.Contains('tenint en consideracio el risc') -or $n -match 'ordeni el cessament') { return $Script:EstatPrecinte }
    # "estimar" i no "desestimar".
    if ($n -match '(^|[^a-z])estimar la sol.?licitud d.?ampliacio') { return $Script:EstatAmpliacio }
    # Desfavorable: deliberadament NO es classifica com a Favorable; cau a Revisar.
    # "no es pot informar favorablement" tambe es un desfavorable (sense aixo, el
    # "favorablement" el feia Favorable).
    if ($n.Contains('desfavorablement') -or $n.Contains('desfavorable') -or $n.Contains('no es pot informar favorablement')) { return 'Revisar' }
    if ($n.Contains('favorablement') -or $n.Contains('favorable')) { return 'Favorable' }
    # Control del Decret 112/2010 que "COMPLEIX el que estableix" sense dir
    # "favorable" a la mateixa conclusio.
    if ($n -match '(?<!\bno )compleix el que estableix el decret') { return 'Favorable' }
    if ($n.Contains('ampliar el termini')) { return ('Ampliaci' + [char]0x00F3 + ' termini') }
    # "Vist l'anterior, s'ha de retirar / s'ha de presentar...": una obligacio
    # sense cap altra decisio. Va DARRERE del favorable a posta: el favorable
    # d'una activitat extraordinaria ("s'informa favorablement tenint en compte
    # les seguents consideracions") llista "S'haura de presentar..." i no es cap
    # requeriment.
    if ($n -match $Script:RxObligacio -or $n -match '\bno es pot informar\b') { return 'Requeriment' }
    # Clausules estandard d'un requeriment NOU (encara sense "Vist l'anterior").
    if ($n.Contains('recepcio del requeriment') -or $n.Contains('esmenar les deficiencies') -or
        $n.Contains('mancances formals') -or $n.Contains('termini maxim de') -or
        $n.Contains('podran adoptar les mesures') -or
        $n.Contains('procediment desmena') -or $n.Contains('esmenar els defectes') -or
        $n -match 'cas contrari.{0,60}(cessament|precinte)' -or $n -match 'determini el (cessament|precinte)') {
        return 'Requeriment'
    }
    return 'Revisar'
}

# Les respostes d'un SEGUIMENT PUNT PER PUNT (formats d'abans, sense frase de
# conclusio): sota cada requeriment, una linia curta. Es miren a l'INICI de la
# linia, despres de la data que hi posa l'eina Seguiment ("dd/MM/aaaa: ").
$Script:RespostesNegatives = @('no es presenta', 'no saporta', 'manca aportar la documentacio',
    'no es justifica', 'no shan retirat', 'no estan esmenad', 'no es disposa')
# Les negatives es miren PRIMER: "No es presenta" no arriba mai a casar amb la
# positiva "es presenta" (i aquesta, amb el \b del final, no casa amb "es
# presentara"). "No es requereix" i "No cal" sota una obligacio la deixen
# resolta. Abans hi faltaven "Es presenta", "S'aplica", "Es tramita"... i un
# seguiment amb tots els punts resolts sortia Requeriment, perque aquelles
# respostes no es veien i els punts quedaven "sense resposta" (8/10/2026).
$Script:RespostesPositives = @('saporta', 'es justifica', 'sentrega', 'saclareix',
    'sha portat a terme amb resultat favorable', 'ok', 'es presenta', 'saplica',
    'es tramita', 'sha realitzat', 'shan realitzat', 'shan retirat', 'sha retirat',
    'no es requereix', 'no cal')

# L'estat d'un informe SENSE cap frase de conclusio (97 dels 802: no son rars,
# son els formats d'abans). Torna @{ Breu; Motiu } o $null si no se'n pot dir
# res (llavors queda 'Revisar', "sense conclusio"). $lines: els paragrafs.
#   - Seguiment punt per punt: alguna resposta negativa -> Requeriment. Si hi ha
#     respostes positives pero algun REQUERIMENT (una linia amb una obligacio,
#     $Script:RxObligacio) es queda sense cap resposta abans del seguent, tampoc
#     no es pot deduir FI: es Requeriment (el concert amb "OK" sota uns punts i
#     res sota el Pla d'Autoprotecció i l'assistencia sanitaria sortia FI,
#     8/10/2026). TOTES respostes -> FI Requeriment, amb motiu (deduit).
#   - Requeriment antic: "S'han observat les seguents deficiencies que cal
#     esmenar per poder (seguir) exercint l'activitat", i acaba a la signatura.
#     Va DESPRES del seguiment: el seguiment d'un requeriment antic el copia
#     sencer, amb aquesta frase inclosa.
#   - Un cos que llista obligacions ("S'ha de presentar...", "Cal justificar...")
#     o diu que el document "no es pot informar": Requeriment. Tambe despres del
#     seguiment, que les copia.
#   - Denuncia d'accessibilitat sense res a requerir ("Sense requeriments
#     especifics", "Actuacio: Cap").
function _EstatSenseConclusio($lines) {
    $neg = 0; $pos = 0; $reqAntic = $false; $accCap = $false; $accAltra = $false
    $oblig = 0; $senseResposta = 0; $pendent = $false; $noInforma = $false
    foreach ($ln in @($lines)) {
        $n = _ConclNorm $ln
        if ($n -eq '') { continue }
        $r = $n -replace '^[\s\-\*–•·]*(\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4}\s*:?\s*)?', ''
        $esResposta = $false
        if ($r.Length -le 200) {
            $esNeg = $false
            foreach ($p in $Script:RespostesNegatives) { if ($r.StartsWith($p)) { $esNeg = $true; break } }
            if ($esNeg) { $neg++; $esResposta = $true }
            else {
                foreach ($p in $Script:RespostesPositives) {
                    if ($r -match ('^' + [regex]::Escape($p) + '(\b|$)')) { $pos++; $esResposta = $true; break }
                }
            }
        }
        if ($esResposta) { $pendent = $false }
        elseif ($n -match $Script:RxObligacio) {
            $oblig++
            if ($pendent) { $senseResposta++ }
            $pendent = $true
        }
        if ($n -match '\bno es pot informar\b') { $noInforma = $true }
        if ($n -match 's.?han observat les seguents deficiencies que cal esmenar') { $reqAntic = $true }
        if ($n.Contains('sense requeriments especifics') -or $n -match 'actuacio\s*:\s*cap\b') { $accCap = $true }
        elseif ($n -match '^actuacio\s*:') { $accAltra = $true }
    }
    if ($pendent) { $senseResposta++ }
    if ($neg -gt 0) { return @{ Breu = 'Requeriment'; Motiu = '' } }
    if ($pos -gt 0 -and $senseResposta -gt 0) { return @{ Breu = 'Requeriment'; Motiu = '' } }
    if ($pos -gt 0) { return @{ Breu = 'FI Requeriment'; Motiu = 'estat deduit, sense conclusio' } }
    if ($reqAntic -or $oblig -gt 0 -or $noInforma) { return @{ Breu = 'Requeriment'; Motiu = '' } }
    if ($accCap -and -not $accAltra) { return @{ Breu = 'FI Requeriment'; Motiu = '' } }
    return $null
}

# El cataleg sencer sense triar res: la conclusio diu alhora que es pot i que
# NO es pot donar per tancada la denuncia, o al cos hi ha quedat el "Copiar
# requeriment" de la plantilla. Abans sortia com a 'Requeriment' (el "no es pot
# donar" guanya), i no ho es: no se sap que diu.
function _EsPlantillaSenseOmplir([string]$conclusio, $lines) {
    $n = _ConclNorm $conclusio
    $nNo = ([regex]::Matches($n, 'no es pot donar per tancada la denuncia')).Count
    $nTots = ([regex]::Matches($n, 'es pot donar per tancada la denuncia')).Count
    if ($nNo -gt 0 -and $nTots -gt $nNo) { return $true }
    foreach ($ln in @($lines)) { if ((_ConclNorm $ln) -match '^copiar requeriment') { return $true } }
    return $false
}

# Expedient de la serie de les activitats extraordinaries (2569/2565): els
# informes antics no porten "ActExtr" al nom. L'expedient es any/numero/SERIE
# ("2025/1/2563"): nomes compta la serie. Abans es mirava qualsevol grup despres
# de l'any, i una MNS de la serie 2562 amb el NUMERO 2565 o 2569 sortia com a
# activitat extraordinaria (i deixava de decidir l'estat del seu establiment).
function _EsExpedientActExtr([string]$expedient) {
    if ($expedient -notmatch '(?:^|\D)(?:19|20)\d{2}\s*[/\-.]\s*\d+\s*[/\-.]\s*0*(\d+)(?:\D|$)') { return $false }
    return ($Matches[1] -eq '2569' -or $Matches[1] -eq '2565')
}

# El TIPUS d'informe, per decidir si fixa l'estat de l'activitat
# (_InformeQueDeterminaEstat). Abans tot el que comencava per "S'informa
# favorablement" o "El titular es responsable d'executar" s'ignorava per
# defecte, perque el TEXT d'aquestes conclusions es gairebe igual d'un informe a
# l'altre; pero per a l'ESTAT barrejava tres coses diferents:
#   'actextr'  activitat extraordinaria (Decret 112/2010): sota el GIA d'un
#              establiment no en decideix l'estat (ni el favorable ni el
#              requeriment: abans l'estat de l'estadi el decidia el requeriment
#              d'un concert).
#   'llicfav'  favorable de llicencia: fixa 'Favorable' i NO s'ha d'ignorar mai
#              (era el cas de l'activitat que es quedava en 'Requeriment').
#   'mns'      MNS, canvi de nom o de titularitat informats: neutre.
#   ''         la resta.
# PURA. Es desa a l'informe ('tipus') perque l'editor pugui recalcular l'estat
# sense tornar a obrir el .docx.
#
# L'ORDRE: el nom del fitxer i el TEXT primer; la serie de l'expedient nomes
# quan cap dels dos no diu res. Una MNS que deia "favorablement de la
# Modificacio..." amb la capcalera mal escrita ("Exp. Num: 2026/1/2565", la
# carpeta era la 2562) sortia 'actextr' (8/10/2026).
function _TipusInforme([string]$conclusio, [string]$fitxer, [string]$expedient) {
    $n = _ConclNorm $conclusio
    if ($fitxer -match '(?i)(^|[^a-z])act[\s_-]?extr' -or
        $n.Contains('responsable dexecutar') -or
        $n.Contains('favorablement tenint en compte les seguents consideracions')) { return 'actextr' }
    if ($n -match 'sinforma favorablement (a lespera de rebre|lactivitat)' -or
        $n.Contains('posterior visita dinspeccio')) { return 'llicfav' }
    if ($n -match 'favorablement (de la|del|al|a la) (modificacio|canvi de nom|canvi de titularitat|transmissio)' -or
        $n.Contains('favorablement una modificacio')) { return 'mns' }
    if (_EsExpedientActExtr $expedient) { return 'actextr' }
    return ''
}

# Tot el que se'n treu del TEXT d'un informe, en un sol lloc (el fan servir
# l'escaneig i ValidarClassificacio.ps1, que han de dir el mateix). PURA.
# Torna @{ Conclusio; Font; Breu; Tipus; Motius }.
function _ClassificaInforme($lines, [string]$fitxer, [string]$expedient) {
    $ci = _ExtractConclusio $lines
    $motius = New-Object System.Collections.ArrayList
    if (_EsPlantillaSenseOmplir $ci.Text $lines) {
        $breu = 'Revisar'
        [void]$motius.Add('plantilla sense omplir')
    } elseif (-not [string]::IsNullOrWhiteSpace($ci.Text)) {
        $breu = _ConclusioBreu $ci.Text
    } else {
        $ded = _EstatSenseConclusio $lines
        if ($null -ne $ded) {
            $breu = [string]$ded.Breu
            if ($ded.Motiu) { [void]$motius.Add([string]$ded.Motiu) }
        } else {
            $breu = 'Revisar'
            [void]$motius.Add('sense conclusio')
        }
    }
    return @{
        Conclusio = $ci.Text
        Font      = $ci.Font
        Breu      = $breu
        Tipus     = (_TipusInforme $ci.Text $fitxer $expedient)
        Motius    = $motius.ToArray()
    }
}

# Els informes d'una activitat per ORDRE: data, i si dos tenen la mateixa data,
# el nom del fitxer (i la ruta). El Sort-Object del PowerShell 5.1 no es estable:
# amb dos informes del mateix dia ordenats nomes per data, l'estat podia canviar
# d'una passada a l'altra.
function _OrdenaInformesActivitat($informes) {
    return @(@($informes) | Where-Object { $null -ne $_ } | Sort-Object -Property `
        @{ Expression = { [string](_PropInf $_ 'data') } },
        @{ Expression = { [string](_PropInf $_ 'fitxer') } },
        @{ Expression = { [string](_PropInf $_ 'ruta') } })
}

# L'informe que DECIDEIX l'estat d'una ACTIVITAT ($act: id_gia + informes). Es
# la font unica: _EstatActualActivitat, Recordatoris (la data del recordatori) i
# Comprovar Excel (la data de l'INFORME ENGINYER) el demanen aqui, aixi no poden
# discrepar mai sobre quin informe mana.
#
# Recorre els informes per ordre (_OrdenaInformesActivitat) i es queda amb
# l'ultim que decideix:
#   - Un informe IGNORAT (a ma, des de l'editor) no compta mai.
#   - 'Altres' (informatius) nomes decideix si no hi ha cap altre informe.
#   - 'actextr' sota un GIA (l'establiment on es fa l'acte) no compta. Sense GIA
#     (agrupat per carpeta) l'"activitat" es l'acte, i si.
#   - 'mns' favorable es NEUTRE: no tapa un estat pendent ($Script:EstatsPendents:
#     l'activitat pot tenir un requeriment obert i presentar una MNS pel mig),
#     pero si no hi ha res pendent l'estat es 'Favorable' (abans una activitat
#     amb nomes MNS favorables quedava amb l'estat buit).
#   Les dues darreres no s'apliquen a un informe corregit a ma (editat_a_ma):
#   l'usuari l'ha mirat i el que hi ha posat MANA.
#
# Demana l'ACTIVITAT i no la llista d'informes perque la regla d'actextr depen
# de si te GIA. Si se li passa una llista, PETA: abans la signatura era la
# llista, i una crida que no s'hagues canviat no fallaria, decidiria l'estat
# sense saber el GIA (hi ha un guard a 06-guards.ps1).
function _InformeQueDeterminaEstat($act) {
    if ($null -eq $act) { return $null }
    if ($act -is [System.Collections.IList] -or $null -eq $act.PSObject.Properties['informes']) {
        throw "_InformeQueDeterminaEstat espera l'ACTIVITAT (id_gia + informes), no la llista d'informes."
    }
    $teGia = -not [string]::IsNullOrWhiteSpace([string](_PropInf $act 'id_gia'))
    $decideix = $null
    $altres = $null
    foreach ($inf in (_OrdenaInformesActivitat $act.informes)) {
        if ([bool](_PropInf $inf 'ignorat')) { continue }
        $breu = [string](_PropInf $inf 'conclusio_breu')
        if ($breu -eq 'Altres') { $altres = $inf; continue }
        if (-not [bool](_PropInf $inf 'editat_a_ma')) {
            $tipus = [string](_PropInf $inf 'tipus')
            if ($tipus -eq 'actextr' -and $teGia) { continue }
            if ($tipus -eq 'mns' -and $breu -eq 'Favorable') {
                if ($null -eq $decideix -or $Script:EstatsPendents -notcontains [string]$decideix.conclusio_breu) { $decideix = $inf }
                continue
            }
        }
        $decideix = $inf
    }
    if ($null -ne $decideix) { return $decideix }
    return $altres
}

# Estat actual d'una ACTIVITAT: la conclusio breu de l'informe que el decideix
# (_InformeQueDeterminaEstat). '' si no n'hi ha cap.
function _EstatActualActivitat($act) {
    $inf = _InformeQueDeterminaEstat $act
    if ($null -eq $inf) { return '' }
    return [string]$inf.conclusio_breu
}
