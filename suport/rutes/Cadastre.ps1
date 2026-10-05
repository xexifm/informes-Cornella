<#
  Cadastre.ps1 - El que tenen en comu TOTES les consultes al Cadastre:
  demanar una URL amb reintents, la memoria cau en disc i el bucle que recorre
  una llista de claus amb barra de progres i Cancel.lar.

  Per que es un fitxer a part: fins a l'octubre de 2026 nomes hi havia una
  consulta (els portals de Coordenades, Geocodificador.ps1) i tot aixo vivia
  alla. El "Planol activitats" en necessita dues mes (la geometria de cada
  parcel.la i la planta/porta de cada unitat), i copiar el bucle dues vegades
  mes era la manera segura que les tres copies divergissin (regla 1 del
  CLAUDE.md). Cada consulta posa nomes el que difereix: la URL i com s'enten
  la resposta.

  FORMAT DE LA MEMORIA CAU (un JSON per consulta, a local\geocodificacio\):
    { "Versio": 1, "<Arrel>": { "<clau>": { "Data": "...", "<Camp>": ... } } }
  <Arrel> i <Camp> els diu cada consulta. Els portals fan servir 'Parcelles' i
  'Portals', que es el format que ja tenia portals.json: res no s'ha de tornar
  a demanar.

  NO LLANCA MAI per culpa del servei: si una consulta falla, aquella clau queda
  sense resultat (i no s'escriu a la memoria cau, perque una caiguda de xarxa
  d'un moment no ens deixi trenta dies amb un buit).

  NOMES DEFINEIX FUNCIONS (i les variables de sota, que config.ps1 pot
  sobreescriure: per aixo es carrega ABANS de Ruta.ps1, que es qui carrega
  config.ps1). ASCII pur.
#>

# Segons d'espera per consulta i nombre d'intents, per a TOTES les consultes.
$CadastreTimeoutSec = 20
$CadastreIntents    = 2

# El motiu de l'ultima fallada, per al diagnostic.
$Script:CadastreUltimError = ''

# Demana una URL i en torna el text, o $null si no hi ha manera.
function Invoke-CadastreGet([string]$url) {
    $Script:CadastreUltimError = ''
    for ($attempt = 1; $attempt -le [int]$CadastreIntents; $attempt++) {
        try {
            $resp = Invoke-WebRequest -Uri $url -Method Get -TimeoutSec ([int]$CadastreTimeoutSec) -UseBasicParsing
            return [string]$resp.Content
        } catch {
            $Script:CadastreUltimError = "$($_.Exception.Message)  [$url]"
            if ($attempt -lt [int]$CadastreIntents) { Start-Sleep -Milliseconds 700 }
        }
    }
    return $null
}

# Diu si una entrada de la memoria cau encara val. PURA (la data d'ara se li
# passa) per poder-la provar sense esperar un any. Una entrada SENSE resultat
# (el camp buit o $null) val menys dies: potser el servei estava caigut.
function Test-CacheCadastreValida($entry, [datetime]$ara, [string]$camp, [double]$dies, [double]$diesBuit) {
    if ($null -eq $entry) { return $false }
    $data = [datetime]::MinValue
    try {
        $data = [datetime]::Parse([string]$entry.Data, [System.Globalization.CultureInfo]::InvariantCulture)
    } catch { return $false }
    $edat = ($ara - $data).TotalDays
    $valor = $null
    if ($null -ne $entry.PSObject.Properties[$camp]) { $valor = $entry.$camp }
    $buit = ($null -eq $valor) -or (@($valor).Count -eq 0)
    $limit = if ($buit) { $diesBuit } else { $dies }
    return ($edat -ge 0 -and $edat -le $limit)
}

function Get-CacheCadastrePath([string]$fitxer) {
    $dir = Get-LocalSubdir $RepoRoot 'Geocodificacio'
    return (Join-Path $dir $fitxer)
}

# Hashtable clau -> entrada. Una memoria cau corrupta no ha de tombar l'eina:
# es descarta i es torna a preguntar.
function Import-CacheCadastre([string]$fitxer, [string]$arrel) {
    try {
        $obj = Read-JsonFile (Get-CacheCadastrePath $fitxer)
        $out = @{}
        if ($null -ne $obj -and $null -ne $obj.PSObject.Properties[$arrel] -and $null -ne $obj.$arrel) {
            foreach ($p in $obj.$arrel.PSObject.Properties) { $out[$p.Name] = $p.Value }
        }
        return $out
    } catch {
        return @{}
    }
}

function Export-CacheCadastre([string]$fitxer, [string]$arrel, $cache) {
    $entrades = [ordered]@{}
    foreach ($k in @($cache.Keys | Sort-Object)) { $entrades[$k] = $cache[$k] }
    Write-JsonFile (Get-CacheCadastrePath $fitxer) ([ordered]@{ Versio = 1; $arrel = $entrades }) 10
}

# EL BUCLE: per a cada clau, la memoria cau si encara val; si no, la consulta.
# Torna una hashtable clau -> resultat (el que torni $parseja; $null si la
# consulta ha fallat).
#
#   $consulta : @{ Fitxer; Arrel; Camp; Dies; DiesBuit;
#                  Url     = { param($clau) <URL> };
#                  Parseja = { param($text) <resultat> } }
#   $onProgress (opcional): & $onProgress $fetes $total $clau ; si torna $false
#                s'atura i es torna el que s'hagi aconseguit (Cancel.lar).
#
# Els blocs Url i Parseja els escriu cada consulta i s'executen aqui: NO porten
# .GetNewClosure() i no ho necessiten, nomes fan servir el que reben.
function Get-AmbCacheCadastre($claus, $consulta, [scriptblock]$onProgress = $null) {
    $result = @{}
    $llista = @($claus | Where-Object { $null -ne $_ -and [string]$_ -ne '' } | ForEach-Object { [string]$_ } | Sort-Object -Unique)
    if ($llista.Count -eq 0) { return $result }

    $cache = Import-CacheCadastre $consulta.Fitxer $consulta.Arrel
    $ara = Get-Date
    $nous = 0
    $fetes = 0
    $total = $llista.Count
    foreach ($clau in $llista) {
        $fetes++
        if ($null -ne $onProgress) {
            $seguim = & $onProgress $fetes $total $clau
            if ($seguim -eq $false) { break }
        }

        $entry = $null
        if ($cache.ContainsKey($clau)) { $entry = $cache[$clau] }
        if (Test-CacheCadastreValida $entry $ara $consulta.Camp ([double]$consulta.Dies) ([double]$consulta.DiesBuit)) {
            $result[$clau] = $entry.($consulta.Camp)
            continue
        }

        $text = Invoke-CadastreGet (& $consulta.Url $clau)
        $valor = $null
        if ($null -ne $text) { $valor = & $consulta.Parseja $text }
        $result[$clau] = $valor
        # Nomes desem el que hem pogut PREGUNTAR (vegeu la capcalera).
        if ($null -ne $text) {
            $e = [ordered]@{ Data = $ara.ToString('yyyy-MM-ddTHH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture) }
            $e[$consulta.Camp] = $valor
            $cache[$clau] = [pscustomobject]$e
            $nous++
            # Cada 25 de noves: si es cancel.la o peta a mitja tanda, no es
            # perd tot el que s'ha demanat.
            if (($nous % 25) -eq 0) { try { Export-CacheCadastre $consulta.Fitxer $consulta.Arrel $cache } catch { } }
        }
    }
    if ($nous -gt 0) {
        try { Export-CacheCadastre $consulta.Fitxer $consulta.Arrel $cache } catch { }   # no poder desar-la no es motiu per fallar
    }
    return $result
}
