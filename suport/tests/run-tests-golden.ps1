#requires -Version 5.1
<#
.SYNOPSIS
  Fitxers d'or: la seqüencia sencera de composicio de cada familia d'informes.

.DESCRIPTION
  Per cada familia es munta un escenari FIX (sempre el mateix cataleg, els
  mateixos punts triats, la mateixa fase) i es compara la llista sencera de
  crides Format-* amb el fitxer desat a dades/emit-<familia>.txt.

  Es la xarxa de seguretat per refer el motor de composicio: cap prova puntual
  no veu alhora un espai que desapareix, un sub-punt que passa de 12 a 6 pt i
  un enllac que canvia de lloc, i aquest projecte ja te l'historial de defectes
  que no fallen sino que empitjoren en silenci.

  ESCENARIS FIXOS, NO ALEATORIS: es trien SEMPRE els mateixos items (els N
  primers de cada seccio), i cap dada depen de la data ni de la maquina. Si un
  fitxer d'or canvia, es perque ha canviat el programa.

  Per refer-los despres d'un canvi VOLGUT:
      GENINFORME_GOLDEN=1 pwsh -NoProfile -File suport/tests/run-tests-golden.ps1
  ...i despres mira't el `git diff`.
#>

$env:GENINFORME_TEST = '1'
$ErrorActionPreference = 'Stop'
# A Linux no existeix LOCALAPPDATA; el donem perque el dot-source no falli
# (mateix guard que run-tests.ps1 i run-tests-actextr.ps1).
if ([string]::IsNullOrEmpty($env:LOCALAPPDATA)) { $env:LOCALAPPDATA = [System.IO.Path]::GetTempPath() }

. (Join-Path $PSScriptRoot 'TestLib.ps1')
. (Join-Path $PSScriptRoot 'Golden.ps1')
. (Join-Path (Split-Path -Parent $PSScriptRoot) 'Motor.ps1')
. (Join-Path $PSScriptRoot 'FormatDoubles.ps1')

$sel = [pscustomobject]@{}
$sel | Add-Member ScriptMethod InsertBreak { param($b) } -Force

function _GdNet { $global:emitCalls.Clear() }

# Els N primers items de cada seccio d'un cataleg. Tria DETERMINISTA: no depen
# de cap dada de cap activitat, nomes del cataleg.
function _GdTriaPrimers($parsed, [int]$perSeccio) {
    $claus = New-Object System.Collections.ArrayList
    foreach ($sec in @($parsed.Sections)) {
        $n = 0
        foreach ($el in @($sec.Items)) {
            if ([string]$el.Kind -eq 'subsection' -or [string]$el.Kind -eq 'intro') { continue }
            [void]$claus.Add((_ItemKey ([string]$sec.Title) ([string]$el.Short)))
            $n++
            if ($n -ge $perSeccio) { break }
        }
    }
    return $claus.ToArray()
}

# ---------------------------------------------------------------------------
Write-Host "`n--- Fitxers d'or: REQ1 i TERMINI ---"
foreach ($nomCat in @('REQ1', 'TERMINI')) {
    $jp = Join-Path $EstructuralsDir ($nomCat + '.json')
    if (-not (Test-Path -LiteralPath $jp)) { Assert $false "or: falta $nomCat.json"; continue }
    $parsed = Get-ParsedCataleg $jp
    $cfg = $Script:ReportFormatConfig
    $fields = [ordered]@{}
    _GdNet
    if ($parsed.IsFixedBody) {
        _WriteCatalegBody $sel $cfg @() $fields ([string]$parsed.IntroText) $true @($parsed.FixedBodyLines)
    } else {
        $seccions = Build-SelectionFromKeys $parsed.Sections (_GdTriaPrimers $parsed 2)
        _WriteCatalegBody $sel $cfg $seccions $fields ([string]$parsed.IntroText)
    }
    $concl = @()
    try { $concl = @(Build-ConclusionsFromTitles (Read-Conclusions $ConclusionsPath $nomCat).Selectable @('Requeriment')) } catch { }
    $sempre = @()
    try { $sempre = @((Read-Conclusions $ConclusionsPath $nomCat).Always) } catch { }
    _WriteConclusionsBlock $sel $cfg 'CONCLUSIONS' $concl $sempre $fields
    Assert-Golden ('cataleg-' + $nomCat.ToLower()) $global:emitCalls
}

# ---------------------------------------------------------------------------
Write-Host "`n--- Fitxers d'or: ACT_EXTR ---"
foreach ($mode in @('req', 'fav')) {
    $tpl = if ($mode -eq 'fav') { $script:ActExtrFavTemplate } else { $script:ActExtrReqTemplate }
    if (-not (Test-Path -LiteralPath $tpl)) { Assert $false "or: falta la plantilla ACT_EXTR $mode"; continue }
    $blocs = Parse-ActExtrTemplate $tpl
    # TOTS els blocs aplicables: aixi el fitxer d'or cobreix la plantilla sencera
    # i no nomes el tros que toqui a una activitat concreta.
    $ctx = @{ Decret = @{}; Computed = @{ PAU_ORGAN_KEY = 'CAT'; HAS_LASERS = $true }; Delivered = @{}; DefKeys = @(); StatusByKey = @{} }
    foreach ($b in $blocs) {
        $k = _ActExtrBlockPoint $b.Key
        if (-not $ctx.StatusByKey.ContainsKey($k)) { $ctx.StatusByKey[$k] = @{ Applies = $true } }
    }
    _GdNet
    _WriteActExtrBody $sel $blocs $mode $ctx $ctx.Computed
    Assert-Golden ('actextr-' + $mode) $global:emitCalls
}

# ---------------------------------------------------------------------------
Write-Host "`n--- Fitxers d'or: les VISTES dels catalegs ---"
foreach ($nomV in @('REQ1', 'TERMINI', 'ACT_EXTR_REQ', 'ACT_EXTR_FAV', 'LLIC', 'MNSTRAS', '0 CONCLUSIONS')) {
    $jp = Join-Path $EstructuralsDir ($nomV + '.json')
    if (-not (Test-Path -LiteralPath $jp)) { continue }
    $o = _LoadEstructuralJson $jp
    _GdNet
    switch ([string]$o.familia) {
        'cataleg'     { _VistaCataleg $sel $jp $nomV }
        'conclusions' { _VistaConclusions $sel $jp }
        'actextr'     { _VistaActExtr $sel $jp $nomV }
        'llicencia'   { _VistaLlicencia $sel $jp }
        'mnstraspas'  { _VistaMnsTraspas $sel $jp }
        default       { continue }
    }
    Assert-Golden ('vista-' + ($nomV -replace '[^A-Za-z0-9]+', '-').ToLower()) $global:emitCalls
}

# ---------------------------------------------------------------------------
# LLICENCIA i MNS/TRASPAS: el document SENCER
# ---------------------------------------------------------------------------
# Aquestes dues families no s'aturen a _WriteCatalegBody: munten el document
# senceres (Build-LlicenciaDocument / Build-MnsDocument), i alli hi ha la
# logica mes subtil del repositori -l'ordre dels enllacos respecte del
# comentari, l'estat de cada punt segons la fase, l'ANNEX 1-. Per aixo el
# fitxer d'or es de la generacio completa, amb un Word de mentida.
Write-Host "`n--- Fitxers d'or: LLICENCIA i MNS/TRASPAS (document sencer) ---"
$llicJson = Join-Path $EstructuralsDir 'LLIC.json'
$req1Json = Join-Path $EstructuralsDir 'REQ1.json'
if ((Test-Path -LiteralPath $llicJson) -and (Test-Path -LiteralPath $req1Json)) {
    $selD = [pscustomobject]@{ Range = [pscustomobject]@{ Start = 0; End = 0 } }
    $selD | Add-Member ScriptMethod EndKey { param($u) } -Force
    $selD | Add-Member ScriptMethod InsertBreak { param($b) } -Force
    $docD = [pscustomobject]@{}
    $docD | Add-Member ScriptMethod Activate {} -Force
    $docD | Add-Member ScriptMethod Save {} -Force
    $docD | Add-Member ScriptMethod Close { param($x) } -Force
    $wordD = [pscustomobject]@{ Selection = $selD }
    $script:_gdDoc = $docD
    function _ResolveOutputDir { return ([System.IO.Path]::GetTempPath()) }
    function _GetUniqueOutputPath($d, $b) { return (Join-Path $d $b) }
    function _OpenOutputDocument($w, $p) { return $script:_gdDoc }
    function Select-CapcaleraBlock($d, $w) { }
    function Apply-HeaderReplacements { param($doc, $header) }
    if ([string]::IsNullOrWhiteSpace($env:TEMP)) { $env:TEMP = [System.IO.Path]::GetTempPath() }

    $llicCat = Read-LlicCataleg $llicJson
    $req1    = Read-CatalegJson $req1Json
    $idx     = _LlicIndexReq1 $req1
    # Escenari FIX: els 3 primers punts de cada bloc, sempre els mateixos.
    # $req1 ES OBLIGATORI: sense ell el bloc ABANS no surt de REQ1 i els punts es
    # queden sense SECCIO ni INTRO -o sigui que el fitxer d'or no representaria
    # el que genera l'assistent, que si que l'hi passa.
    $bAbans  = @(@((_LlicPuntsPerBloc $llicCat $idx 'ABANS'   $req1).Punts) | Select-Object -First 4)
    # PROU PUNTS per arribar a les seccions EXPANDIDES de REQ1: els sis primers
    # son text propi de LLIC (sense seccio, com va decidir l'usuari) i les
    # subseccions i els textos fixos comencen despres. Amb menys, el fitxer d'or
    # no cobriria justament el que ha de vigilar.
    $bDespres= @(@((_LlicPuntsPerBloc $llicCat $idx 'DESPRES' $req1).Punts) | Select-Object -First 16)
    # I uns quants punts de PROJECTE, que venen de Select-Items (seccions
    # senceres): es el bloc on es veu que la seccio de REQ1 hi ha de sortir.
    $bProj = @(Build-SelectionFromKeys $req1.Sections (@(_GdTriaPrimers $req1 1) | Select-Object -First 3))
    $capcal  = @{ ID_GIA = '1457'; TITULAR = 'TITULAR DE PROVA SL'; CLASSIFICACIO = 'Llei 20/2009; Annex II; Epigraf 12.25' }

    foreach ($fase in @('requeriment', 'favorable-pre', 'favorable-post')) {
        $abans = @(@($bAbans) | ForEach-Object { $_ | Select-Object * | Add-Member NoteProperty Estat 'no' -PassThru -Force })
        $desp  = @(@(_LlicPuntsAmbEstatFase $bDespres $fase) | ForEach-Object { $_ | Add-Member NoteProperty Estat 'no' -PassThru -Force })
        $model = @{
            Fase = $fase; EsProvisional = $false
            Header = $capcal; Fields = [ordered]@{}
            Abans = $abans; Projecte = @(_LlicPuntsDeSeleccio $bProj); Despres = $desp
            Doc = @{ Text = 'La documentacio tecnica ve signada per:'; Items = @('Projecte', 'Planols') }
            Condicions = ''; Cataleg = $llicCat
        }
        _GdNet
        [void](Build-LlicenciaDocument $wordD $model)
        Assert-Golden ('llicencia-' + $fase) $global:emitCalls
    }

    # I la llicencia PROVISIONAL, que es l'unica que porta l'ANNEX 1.
    $abansP = @(@($bAbans) | ForEach-Object { $_ | Select-Object * | Add-Member NoteProperty Estat 'no' -PassThru -Force })
    $modelP = @{
        Fase = 'requeriment'; EsProvisional = $true
        Header = $capcal; Fields = [ordered]@{}
        Abans = $abansP; Projecte = @(); Despres = @()
        Doc = @{ Text = ''; Items = @() }; Condicions = ''; Cataleg = $llicCat
    }
    _GdNet
    [void](Build-LlicenciaDocument $wordD $modelP)
    Assert-Golden 'llicencia-provisional-annex1' $global:emitCalls

    # MNS i TRASPAS, amb punts de REQ1 i sense (les dues branques de la
    # frase d'observacions i del bloc de CONCLUSIONS).
    $mnsJson = Join-Path $EstructuralsDir 'MNSTRAS.json'
    if (Test-Path -LiteralPath $mnsJson) {
        $mnsCat = _LoadEstructuralJson $mnsJson
        $puntsReq1 = Build-SelectionFromKeys $req1.Sections (@(_GdTriaPrimers $req1 1) | Select-Object -First 2)
        foreach ($fase in @('mns', 'traspas')) {
            foreach ($amb in @($true, $false)) {
                $model = @{
                    Fase = $fase; Header = $capcal; Fields = [ordered]@{}
                    Cataleg = $mnsCat
                    Punts = $(if ($amb) { $puntsReq1 } else { @() })
                }
                _GdNet
                [void](Build-MnsDocument $wordD $model)
                Assert-Golden ($fase + $(if ($amb) { '-amb-punts' } else { '-sense-punts' })) $global:emitCalls
            }
        }
    }
}

exit (Write-TestSummary 'RESULTAT OR')
