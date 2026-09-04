#requires -Version 5.1
<#
.SYNOPSIS
  Llegir docs\config.js des de PowerShell. Un sol lloc, i amb prova.

.DESCRIPTION
  docs\config.js es el fitxer que l'usuari edita per posar el mobil en marxa
  (DESPLEGAMENT-MOBIL.md). El navegador el carrega com un <script> SINCRON abans
  d'app.js, o sigui que ha de continuar sent JavaScript: no es pot convertir en
  un .json sense reestructurar el web.

  El que si que es podia arreglar es com el llegeix el PC. EnviarCorreu.ps1
  portava QUATRE expressions regulars escampades dins de _CorreuConfig, una per
  clau, i cadascuna exigia cometes DOBLES i el nom literal. Amb cometes simples,
  amb la clau comentada o amb una plantilla, la clau es quedava BUIDA en silenci
  i l'enviament de correu deixava de funcionar sense dir per que. No hi havia cap
  prova que ho cobris.

  Ara hi ha un sol lector, i una prova que el fa correr contra el docs\config.js
  de debo i comprova que en treu les claus que el programa necessita. Si algu
  reformata el fitxer, ho diu la suite ABANS que es trenqui res.

  NO es un interpret de JavaScript ni ho ha de ser: window.CONFIG es una llista
  plana de CLAU: "valor" i prou. El que no encaixi amb aixo no es llegeix, i per
  aixo la prova comprova les claus que fem servir de debo.
#>

# Ruta de docs\config.js dins del clone.
function Get-ConfigJsPath {
    $root = if ($RepoRoot) { $RepoRoot } else { (Get-Location).Path }
    return (Join-Path $root (Join-Path 'docs' 'config.js'))
}

# Llegeix els parells CLAU: "valor" de docs\config.js i els torna en un
# hashtable. Accepta cometes dobles o simples i salta les linies comentades amb
# //. Si el fitxer no hi es, torna un hashtable BUIT (mai peta): qui decideix si
# una clau que falta es un problema es el crider, que es qui sap per a que la vol.
#
# $Text permet passar-li un contingut ja llegit (les proves).
function Read-ConfigJs([string]$Text = '') {
    $out = @{}
    $t = $Text

    # SENSE TEXT, ES LLEGEIX EL FITXER -i es comprova el BUIT, no el $null-.
    #
    # Aixo va nomes escrit aixi perque la primera versio deia
    # '[string]$Text = $null' i preguntava 'if ($null -eq $t)'. Un parametre
    # TIPAT [string] amb valor per defecte $null NO arriba com a $null: el
    # PowerShell el converteix a CADENA BUIDA. La condicio no es complia mai, la
    # funcio no llegia el fitxer i tornava sempre un diccionari buit, o sigui que
    # els tres IDs de carpeta de Drive quedaven a ''. Ho va enxampar la prova
    # d'aqui sota, que compara el que llegeix amb el docs/config.js de debo.
    if ([string]::IsNullOrEmpty($t)) {
        $p = Get-ConfigJsPath
        if (-not (Test-Path -LiteralPath $p)) { return $out }
        try { $t = Get-Content -LiteralPath $p -Raw -Encoding UTF8 } catch { return $out }
    }
    if ([string]::IsNullOrWhiteSpace($t)) { return $out }

    foreach ($linia in ($t -split "`r?`n")) {
        $l = $linia.Trim()
        if ($l.StartsWith('//')) { continue }
        # CLAU: "valor"  |  CLAU: 'valor'
        $m = [regex]::Match($l, '^([A-Z][A-Z0-9_]*)\s*:\s*(?:"([^"]*)"|''([^'']*)'')')
        if (-not $m.Success) { continue }
        $clau = $m.Groups[1].Value
        $val  = if ($m.Groups[2].Success) { $m.Groups[2].Value } else { $m.Groups[3].Value }
        if (-not $out.ContainsKey($clau)) { $out[$clau] = $val }
    }
    return $out
}

# Una clau concreta, amb valor per defecte si no hi es o es buida.
function Get-ConfigJsValue($cfg, [string]$Clau, [string]$PerDefecte = '') {
    if ($null -eq $cfg -or -not $cfg.ContainsKey($Clau)) { return $PerDefecte }
    $v = [string]$cfg[$Clau]
    if ([string]::IsNullOrWhiteSpace($v)) { return $PerDefecte }
    return $v
}
