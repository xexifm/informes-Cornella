#requires -Version 5.1
<#
.SYNOPSIS
  Exporta les dades del programa a JSON perque el formulari web del mobil les
  llegeixi sense cap gestio manual.

.DESCRIPTION
  Genera dos tipus de sortida, separades per sensibilitat:

  1. PLANTILLES (NO contenen dades personals) -> suport/web/dades/*.json
       - cataleg-<BaseName>.json  (un per cada REQ*.docx d'ESTRUCTURALS)
       - conclusions.json
       - capcalera.json           (placeholders <<...>> detectats)
       - manifest.json            (llista de catalegs + data de generacio)
     Aquests fitxers SI es pugen al GitHub public (els serveix GitHub Pages al
     mobil). Els refresca i puja Actualitzar.bat.

  2. ACTIVITATS (dades personals: noms i adreces) -> $DriveDadesDir\activitats.json
     NOMES els camps de capcalera per ID GIA. Van a la carpeta PRIVADA de Google
     Drive. NO es pugen MAI al GitHub public.

.PARAMETER Plantilles
  Exporta nomes les plantilles (REQ/conclusions/capcalera) a web/dades.

.PARAMETER Activitats
  Exporta nomes la base de dades d'activitats a Drive.

  Si no s'indica cap dels dos, s'exporten totes dues coses.

.NOTES
  Reutilitza les funcions de GenerarInforme.ps1 (Parse-Cataleg, Read-Conclusions,
  Initialize-ActivitatsCache...) carregant-lo en mode headless. Necessita Word
  instal·lat (i Excel per a -Activitats).
#>
param(
    [switch]$Plantilles,
    [switch]$Activitats
)

$ErrorActionPreference = 'Stop'

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

# Carreguem GenerarInforme.ps1 en mode headless: defineix les funcions i les
# rutes ($EstructuralsDir, $ConclusionsPath, $HeaderPath, $DriveDadesDir...)
# sense obrir cap finestra ni executar Main.
$env:GENINFORME_TEST = '1'
. (Join-Path $ScriptRoot 'GenerarInforme.ps1')

# Si no s'especifica res, fem totes dues exportacions.
if (-not $Plantilles -and -not $Activitats) {
    $Plantilles = $true
    $Activitats = $true
}

$WebDadesDir = Join-Path $ScriptRoot (Join-Path 'web' 'dades')

# Converteix un element del cataleg (de Parse-Cataleg) en un objecte
# JSON-friendly (arrays plans en lloc d'ArrayList), preservant
# Kind/Short/BodyLines/Children. Les BodyLines es deixen TAL QUAL (amb els
# marcadors [[URL]] i els placeholders [CAMP:]/[OPCIO:]): el formulari web les
# interpreta igual que el motor del PC.
function _ItemToJson($el) {
    $children = @()
    if ($el.Children) {
        foreach ($ch in $el.Children) {
            $children += [pscustomobject]@{
                Kind      = $ch.Kind
                Short     = $ch.Short
                BodyLines = @($ch.BodyLines)
            }
        }
    }
    return [pscustomobject]@{
        Kind      = $el.Kind
        Short     = $el.Short
        BodyLines = @($el.BodyLines)
        Children  = $children
    }
}

function Export-Plantilles {
    if (-not (Test-Path -LiteralPath $WebDadesDir)) {
        New-Item -ItemType Directory -Path $WebDadesDir -Force | Out-Null
    }
    $word = New-WordApp
    try {
        $catalegs = @(Get-Catalegs)
        if ($catalegs.Count -eq 0) {
            Write-Host "  (cap cataleg REQ*.docx trobat a $EstructuralsDir)"
        }
        $catNames = @()
        foreach ($cat in $catalegs) {
            $parsed = Get-ParsedCataleg -word $word -path $cat.FullName
            $sectionsJson = @()
            foreach ($sec in $parsed.Sections) {
                $itemsJson = @()
                foreach ($el in $sec.Items) { $itemsJson += (_ItemToJson $el) }
                $sectionsJson += [pscustomobject]@{ Title = $sec.Title; Items = $itemsJson }
            }
            $obj = [pscustomobject]@{
                BaseName  = $cat.BaseName
                IntroText = $parsed.IntroText
                Sections  = $sectionsJson
            }
            $outFile = Join-Path $WebDadesDir ("cataleg-$($cat.BaseName).json")
            ($obj | ConvertTo-Json -Depth 12) | Set-Content -LiteralPath $outFile -Encoding UTF8
            $catNames += $cat.BaseName
            Write-Host "  cataleg-$($cat.BaseName).json"
        }

        # Conclusions
        $concl = Read-Conclusions -word $word -path $ConclusionsPath
        $conclObj = [pscustomobject]@{
            HeaderText = $concl.HeaderText
            Selectable = @($concl.Selectable | ForEach-Object { [pscustomobject]@{ Title=$_.Title; Body=$_.Body } })
            Always     = @($concl.Always)
        }
        ($conclObj | ConvertTo-Json -Depth 8) |
            Set-Content -LiteralPath (Join-Path $WebDadesDir 'conclusions.json') -Encoding UTF8
        Write-Host "  conclusions.json"

        # Capcalera: placeholders <<...>> presents al 0 CAPCALERA.docx.
        $headerFields = @()
        if (Test-Path -LiteralPath $HeaderPath) {
            $doc = $word.Documents.Open($HeaderPath, $false, $true)
            try {
                $txt = [string]$doc.Content.Text
                $seen = @{}
                foreach ($m in ([regex]'<<\s*([A-Za-z0-9_]+)\s*>>').Matches($txt)) {
                    $name = $m.Groups[1].Value
                    if (-not $seen.ContainsKey($name)) { $seen[$name] = $true; $headerFields += $name }
                }
            } finally { $doc.Close($false) }
        }
        ([pscustomobject]@{ Placeholders = @($headerFields) } | ConvertTo-Json -Depth 4) |
            Set-Content -LiteralPath (Join-Path $WebDadesDir 'capcalera.json') -Encoding UTF8
        Write-Host "  capcalera.json"

        # Manifest
        ([pscustomobject]@{
            GeneratedAt  = (Get-Date).ToString('o')
            Catalegs     = @($catNames)
            HeaderFields = @($headerFields)
        } | ConvertTo-Json -Depth 4) |
            Set-Content -LiteralPath (Join-Path $WebDadesDir 'manifest.json') -Encoding UTF8
        Write-Host "  manifest.json"
    } finally {
        Close-WordApp $word
    }
}

function Export-ActivitatsCmd {
    $latest = Find-LatestActivitatsExcel
    if ($null -eq $latest) {
        Write-Host "  (cap Excel d'activitats trobat; ometo l'exportacio d'activitats)"
        return
    }
    $cache = Initialize-ActivitatsCache -excelFile $latest.File
    $ok = Export-ActivitatsToDrive $cache $latest
    if ($ok) {
        Write-Host "  activitats.json -> $DriveDadesDir ($($cache.ById.Count) activitats, font $($latest.File.Name))"
    }
}

if ($Plantilles) {
    Write-Host "Exportant plantilles a $WebDadesDir ..."
    Export-Plantilles
}
if ($Activitats) {
    Write-Host "Exportant base de dades d'activitats a Drive ..."
    Export-ActivitatsCmd
}
Write-Host "Fet."
