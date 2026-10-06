<#
  Planol.ps1 - "Planol activitats": un planol de Cornella amb les PARCEL.LES
  pintades segons les activitats que hi ha i en quin estat estan.

  Que fa:
    1. Agafa l'Excel d'ACTIVITATS i el d'ESTABLIMENTS mes nous (xarxa de la
       feina o carpeta local), i la base d'informes.
    2. Pregunta al Cadastre la planta/porta de les unitats que l'Excel no
       distingeix, i la GEOMETRIA de cada parcel.la. Tot amb memoria cau: la
       primera vegada triga uns minuts; les seguents, segons.
    3. Genera un mapa HTML (PlanolMapa.html) a local\planol-activitats\ i l'obre.
       Es NOMES d'aquest ordinador: porta requeriments pendents i noms
       d'activitats, que no han de ser publics.

  Les dades i els colors: PlanolDades.ps1. El Cadastre: Cadastre.ps1.
  NO toca res: ni els Excel ni la base d'informes.

  Mode "headless" per a proves: amb $env:PLANOL_TEST o $env:GENINFORME_TEST
  NOMES es defineixen les funcions.
#>

$ErrorActionPreference = 'Stop'

$Script:PlanolHeadless = [bool]$env:PLANOL_TEST -or [bool]$env:GENINFORME_TEST

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$SuportDir  = Split-Path -Parent $ScriptRoot          # suport/
$RepoRoot   = Split-Path -Parent $SuportDir           # informes-Cornella/

if (-not $Script:PlanolHeadless) {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()
}

# Els moduls amb variables que config.ps1 pot sobreescriure van ABANS de
# Ruta.ps1 (que es qui carrega config.ps1), com a Coordenades.
. (Join-Path $ScriptRoot 'Cadastre.ps1')
. (Join-Path $ScriptRoot 'Geocodificador.ps1')    # Test-CoordPlausible
. (Join-Path $ScriptRoot 'PlanolDades.ps1')

# Ruta.ps1 en mode headless: nomes en volem les funcions (cerca de l'Excel,
# Find-HeaderColumn, conversio UTM, format d'adreca) i que carregui config.ps1,
# Excel.ps1, Json.ps1, UiFinestra.ps1...
$Script:_prevRutaTestPlanol = $env:RUTA_TEST
$env:RUTA_TEST = '1'
try {
    . (Join-Path $ScriptRoot 'Ruta.ps1')
} finally {
    if ($null -eq $Script:_prevRutaTestPlanol) {
        Remove-Item Env:\RUTA_TEST -ErrorAction SilentlyContinue
    } else {
        $env:RUTA_TEST = $Script:_prevRutaTestPlanol
    }
}

. (Join-Path $ScriptRoot 'MapaHtml.ps1')
. (Join-Path $ScriptRoot 'EinesUi.ps1')
$Script:EinaTitol = ('Pl' + [char]0x00E0 + 'nol activitats')
if (-not $Script:PlanolHeadless) { $Script:EinaIcon = Get-EinaIcon $SuportDir }

$Script:PlanolPlantilla = Join-Path $ScriptRoot 'PlanolMapa.html'
$PlanolOutputDir = Get-LocalSubdir $RepoRoot 'Planol'

# ============================================================================
# EL MAPA (HTML)
# ============================================================================
# $meta: { BaseActivitats; BaseEstabliments; BaseInformes; Avisos (array) }
function Build-PlanolHtml($dadesMapa, $meta) {
    $avisos = @($meta.Avisos | Where-Object { [string]$_ -ne '' })
    $avisosHtml = ''
    if ($avisos.Count -gt 0) {
        $avisosHtml = '<div id="avisos">' + ((@($avisos | ForEach-Object { _HtmlEncode ([string]$_) })) -join '<br>') + '</div>'
    }
    $valors = @{
        dadesJson    = (ConvertTo-JsonScript @($dadesMapa) -Llista -Fondaria 10)
        baseAct      = (_HtmlEncode ([string]$meta.BaseActivitats))
        baseEst      = (_HtmlEncode ([string]$meta.BaseEstabliments))
        baseInf      = (_HtmlEncode ([string]$meta.BaseInformes))
        generat      = (Get-Date).ToString('dd/MM/yyyy HH:mm')
        avisos       = $avisosHtml
    }
    return (Get-PlantillaHtml $Script:PlanolPlantilla $valors)
}

# ============================================================================
# LECTURA (COM) - nomes a Windows amb Excel
# ============================================================================
function Read-EstablimentsExcel($excelFile) {
    $out = Read-FullaEstesa $excelFile {
        param($x)
        if ($null -eq $x.Data) { return ,@() }
        return ,@(ConvertFrom-FullaEstabliments $x.Data $x.Rows $x.Headers)
    } -Fulla 'Establiments'
    return ,@($out)
}

function Read-ActivitatsPlanolExcel($excelFile) {
    $out = Read-FullaEstesa $excelFile {
        param($x)
        if ($null -eq $x.Data) { return @{} }
        return (ConvertFrom-FullaActivitatsPlanol $x.Data $x.Rows $x.Headers)
    }
    return $out
}

# Una feina del Cadastre amb barra de progres. Cancel.lar no avorta l'eina:
# deixa de preguntar i el mapa es fa amb el que ja hi ha (memoria cau inclosa).
# Torna { Resultat; Cancelat }.
function _PlanolAmbProgres($claus, [string]$text, [scriptblock]$feina) {
    $llista = @($claus)
    if ($llista.Count -eq 0) { return [pscustomobject]@{ Resultat = @{}; Cancelat = $false } }
    $prog = New-EinaProgres $llista.Count 'Consultant el Cadastre' $text
    $onProgress = {
        param($fetes, $total, $clau)
        $prog.Bar.Value = [math]::Min($fetes, $prog.Bar.Maximum)
        $prog.Label.Text = "$text`n$fetes de $total  ($clau)"
        [System.Windows.Forms.Application]::DoEvents()
        return (-not $prog.Estat.Cancelat)
    }.GetNewClosure()
    $r = @{}
    try {
        $r = & $feina $llista $onProgress
    } catch {
        $r = @{}
    } finally {
        if (-not $prog.Form.IsDisposed) { $prog.Form.Close() }
    }
    return [pscustomobject]@{ Resultat = $r; Cancelat = [bool]$prog.Estat.Cancelat }
}

# ============================================================================
# MAIN
# ============================================================================
function Invoke-PlanolMain {
    $xlsA = Find-LatestRutaExcel
    if ($null -eq $xlsA) {
        Show-EinaInfo ("No s'ha trobat cap base de dades d'activitats.`n`n" +
            "Busco un fitxer 'AAAA-MM-DD ACTIVITATS.xlsx' a:`n  1. $ActivitatsDir`n  2. $LocalActivitatsDir") '' 'Warning'
        return
    }
    $xlsE = Find-LatestRutaExcel 'ESTABLIMENTS'
    $avisos = @()
    if ($null -eq $xlsE) {
        $r = [System.Windows.Forms.MessageBox]::Show(
            ("No s'ha trobat l'Excel d'ESTABLIMENTS ('AAAA-MM-DD ESTABLIMENTS.xls') a:`n  1. $ActivitatsDir`n  2. $LocalActivitatsDir`n`n" +
             "Sense ell, cada activitat surt nomes a la parcel.la de l'Excel d'activitats (una de sola) i no es veuen els locals buits.`n`n" +
             "Vols fer el planol igualment?"), $Script:EinaTitol, 'YesNo', 'Warning')
        if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $avisos += "Sense l'Excel d'establiments: cada activitat surt nomes a la parcel" + [char]0x00B7 + "la de l'Excel d'activitats."
    }

    # 1. Llegir els Excel.
    $espera = New-EinaProgres 1 'Llegint les bases de dades' "Llegint $($xlsA.File.Name)..."
    $espera.Bar.Style = 'Marquee'
    [System.Windows.Forms.Application]::DoEvents()
    try {
        $acts = Read-ActivitatsPlanolExcel $xlsA.File
        $ests = @()
        if ($null -ne $xlsE) {
            $espera.Label.Text = "Llegint $($xlsE.File.Name)..."
            [System.Windows.Forms.Application]::DoEvents()
            # SENSE @(): Read-EstablimentsExcel ja torna la llista sencera (amb
            # coma), i un @() al voltant la tornava a embolcallar: el model rebia
            # UN establiment que les contenia tots, no en lligava cap amb la
            # seva activitat i TOTES sortien "sense establiment" (octubre 2026).
            $ests = Read-EstablimentsExcel $xlsE.File
        }
    } catch {
        if (-not $espera.Form.IsDisposed) { $espera.Form.Close() }
        Show-EinaInfo "Error llegint l'Excel:`n$($_.Exception.Message)" '' 'Error'
        return
    } finally {
        if (-not $espera.Form.IsDisposed) { $espera.Form.Close() }
    }

    # 2. La base d'informes (opcional: sense, tot surt en blau).
    $estats = @{}
    $baseInf = '(no hi ha base d''informes)'
    $dbPath = Join-Path $LocalActivitatsDir 'informes-db.json'
    if (Test-Path -LiteralPath $dbPath) {
        try {
            $db = Read-JsonFile $dbPath
            $estats = ConvertFrom-InformesDbPlanol $db
            $baseInf = 'Base d''informes'
            if ($null -ne $db.PSObject.Properties['actualitzat_el']) {
                try { $baseInf += ' del ' + ([datetime]::Parse([string]$db.actualitzat_el)).ToString('dd/MM/yyyy HH:mm') } catch { }
            }
        } catch { $estats = @{} }
    }
    if ($estats.Count -eq 0) {
        $avisos += "No s'ha pogut llegir la base d'informes: cap activitat no t" + [char]0x00E9 + " estat (surten totes en blau). Fes abans 'Actualitzar base'."
    }

    # 3. La planta/porta de les unitats que l'Excel no distingeix.
    $rcU = @(Get-UnitatsAConsultar $ests)
    $fetU = _PlanolAmbProgres $rcU ("Planta i porta de " + $rcU.Count + " locals sense local/planta/porta a l'Excel...") {
        param($l, $p) Get-UnitatsCadastre $l $p
    }
    $unitats = $fetU.Resultat
    $cancelU = $fetU.Cancelat

    # 4. El model i la geometria de les parcel.les.
    $model = Build-PlanolModel $ests $acts $estats $unitats
    $rcP = @($model.Parceles | Where-Object { $_.Rc -ne '' } | ForEach-Object { $_.Rc })
    $fetP = _PlanolAmbProgres $rcP ("Dibuix de " + $rcP.Count + " parcel" + [char]0x00B7 + "les...") {
        param($l, $p) Get-GeometriesParceles $l $p
    }
    $geos = $fetP.Resultat
    $cancelP = $fetP.Cancelat
    if ($cancelU -or $cancelP) { $avisos += "Has cancel" + [char]0x00B7 + "lat les consultes al Cadastre: algunes parcel" + [char]0x00B7 + "les surten com un punt. Torna-ho a generar per completar-les (el que ja s'ha demanat queda desat)." }
    $senseGeo = @($rcP | Where-Object { -not $geos.ContainsKey($_) -or @($geos[$_]).Count -eq 0 }).Count
    if ($senseGeo -gt 0 -and -not ($cancelU -or $cancelP)) {
        $avisos += "$senseGeo parcel" + [char]0x00B7 + "les sense dibuix del Cadastre: surten com un punt."
    }

    # 5. El mapa.
    $dades = ConvertTo-PlanolDadesMapa $model $geos
    $meta = [pscustomobject]@{
        BaseActivitats   = $xlsA.File.Name
        BaseEstabliments = if ($null -ne $xlsE) { $xlsE.File.Name } else { '(sense Excel d''establiments)' }
        BaseInformes     = $baseInf
        Avisos           = $avisos
    }
    $html = Build-PlanolHtml $dades $meta
    if (-not (Test-Path -LiteralPath $PlanolOutputDir)) { New-Item -ItemType Directory -Path $PlanolOutputDir -Force | Out-Null }
    $outPath = Join-Path $PlanolOutputDir ("Planol_" + (Get-Date).ToString('yyyy-MM-dd_HHmmss') + '.html')
    [System.IO.File]::WriteAllText($outPath, $html, (New-Object System.Text.UTF8Encoding($false)))
    Start-Process $outPath

    $res = $model.Resum
    $msg  = "Planol generat: $(@($dades).Count) parcel" + [char]0x00B7 + "les amb $($res.Activitats) activitats.`n`n"
    if ($res.SenseEstabliment -gt 0) { $msg += "Activitats sense cap establiment (situades amb la refer" + [char]0x00E8 + "ncia de l'Excel d'activitats): $($res.SenseEstabliment)`n" }
    if ($res.NoBase -gt 0) { $msg += "Establiments amb una activitat que no " + [char]0x00E9 + "s a la base d'activitats: $($res.NoBase)`n" }
    if ($res.MarcatsBuit -gt 0) { $msg += "Activitats en un local marcat com a buit (per revisar): $($res.MarcatsBuit)`n" }
    if ($res.SensePosicio -gt 0) { $msg += "Sense refer" + [char]0x00E8 + "ncia cadastral ni coordenades (no surten): $($res.SensePosicio)`n" }
    foreach ($a in $avisos) { $msg += "`n$a" }
    $msg += "`n`nFitxer: $outPath"
    Show-EinaInfo $msg
}

if (-not $Script:PlanolHeadless) {
    Invoke-PlanolMain
}
