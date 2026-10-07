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

# $PlanolNomesFuncions: el posa qui el carrega com a biblioteca (PlanolAuto.ps1,
# el mode automatic) abans de fer-ne el dot-source.
$Script:PlanolHeadless = [bool]$env:PLANOL_TEST -or [bool]$env:GENINFORME_TEST -or [bool]$PlanolNomesFuncions

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
. (Join-Path $ScriptRoot 'CadastreParceles.ps1')   # el dibuix i el punt de cada parcel.la (compartit amb Coordenades)
. (Join-Path $ScriptRoot 'Geocodificador.ps1')    # Test-CoordPlausible, els portals de cada parcel.la
. (Join-Path $ScriptRoot 'PlanolGeometria.ps1')  # l'entrada dins la parcel.la i les parcel.les juntades
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
. (Join-Path (Split-Path -Parent $ScriptRoot) 'SegonPla.ps1')   # la copia per al mobil (mobil/PujaPlanol.ps1)
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

$Script:PlanolTeCcae = $true
function Read-ActivitatsPlanolExcel($excelFile) {
    $out = Read-FullaEstesa $excelFile {
        param($x)
        if ($null -eq $x.Data) { return @{} }
        # Si no hi ha CCAE, el filtre d'allotjaments turistics no pot fer res: es diu.
        $Script:PlanolTeCcae = (@(Get-ColumnesCcae $x.Headers).Count -gt 0)
        return (ConvertFrom-FullaActivitatsPlanol $x.Data $x.Rows $x.Headers)
    }
    return $out
}

# Una feina del Cadastre amb barra de progres. Cancel.lar no avorta l'eina:
# deixa de preguntar i el mapa es fa amb el que ja hi ha (memoria cau inclosa).
# Torna { Resultat; Cancelat }.
# En SILENCI (el mode automatic, sense pantalla): la feina sense barra.
$Script:PlanolSilenci = $false
function _PlanolAmbProgres($claus, [string]$text, [scriptblock]$feina) {
    $llista = @($claus)
    if ($llista.Count -eq 0) { return [pscustomobject]@{ Resultat = @{}; Cancelat = $false } }
    if ($Script:PlanolSilenci) {
        $r = @{}
        try { $r = & $feina $llista $null } catch { $r = @{} }
        return [pscustomobject]@{ Resultat = $r; Cancelat = $false }
    }
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
# FER EL PLANOL (amb finestres o en silenci) I MAIN
# ============================================================================
# Fa el planol i el desa. Amb $silenci (el mode automatic setmanal, octubre
# 2026: PlanolAuto.ps1) no obre cap finestra ni pregunta res: sense l'Excel
# d'establiments el fa igualment (amb l'avis). Torna { Ok; Error; OutPath;
# Missatge } (Error buit + Ok fals = l'usuari ha dit que no).
function Invoke-PlanolGenera([bool]$silenci) {
    $Script:PlanolSilenci = $silenci
    $xlsA = Find-LatestRutaExcel
    if ($null -eq $xlsA) {
        return [pscustomobject]@{ Ok = $false; OutPath = ''; Missatge = ''
            Error = ("No s'ha trobat cap base de dades d'activitats.`n`n" +
                     "Busco un fitxer 'AAAA-MM-DD ACTIVITATS.xlsx' a:`n  1. $ActivitatsDir`n  2. $LocalActivitatsDir") }
    }
    $xlsE = Find-LatestRutaExcel 'ESTABLIMENTS'
    $avisos = @()
    if ($null -eq $xlsE) {
        if (-not $silenci) {
            $r = [System.Windows.Forms.MessageBox]::Show(
                ("No s'ha trobat l'Excel d'ESTABLIMENTS ('AAAA-MM-DD ESTABLIMENTS.xls') a:`n  1. $ActivitatsDir`n  2. $LocalActivitatsDir`n`n" +
                 "Sense ell, cada activitat surt nomes a la parcel.la de l'Excel d'activitats (una de sola) i no es veuen els locals buits.`n`n" +
                 "Vols fer el planol igualment?"), $Script:EinaTitol, 'YesNo', 'Warning')
            if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return [pscustomobject]@{ Ok = $false; Error = ''; OutPath = ''; Missatge = '' } }
        }
        $avisos += "Sense l'Excel d'establiments: cada activitat surt nomes a la parcel" + [char]0x00B7 + "la de l'Excel d'activitats."
    }

    # 1. Llegir els Excel.
    $espera = $null
    if (-not $silenci) {
        $espera = New-EinaProgres 1 'Llegint les bases de dades' "Llegint $($xlsA.File.Name)..."
        $espera.Bar.Style = 'Marquee'
        [System.Windows.Forms.Application]::DoEvents()
    }
    try {
        $acts = Read-ActivitatsPlanolExcel $xlsA.File
        $ests = @()
        if ($null -ne $xlsE) {
            if ($null -ne $espera) {
                $espera.Label.Text = "Llegint $($xlsE.File.Name)..."
                [System.Windows.Forms.Application]::DoEvents()
            }
            # SENSE @(): Read-EstablimentsExcel ja torna la llista sencera (amb
            # coma), i un @() al voltant la tornava a embolcallar: el model rebia
            # UN establiment que les contenia tots, no en lligava cap amb la
            # seva activitat i TOTES sortien "sense establiment" (octubre 2026).
            $ests = Read-EstablimentsExcel $xlsE.File
        }
    } catch {
        return [pscustomobject]@{ Ok = $false; OutPath = ''; Missatge = ''; Error = "Error llegint l'Excel:`n$($_.Exception.Message)" }
    } finally {
        if ($null -ne $espera -and -not $espera.Form.IsDisposed) { $espera.Form.Close() }
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
        param($l, $p) Get-ParcelesCadastre $l $p
    }
    # El dibuix i el punt de cada parcel.la (el punt, per a la linia de punts
    # fins a l'etiqueta de l'ID).
    # Si la consulta peta, _PlanolAmbProgres torna un hashtable buit, no l'objecte.
    $geos = @{}; $puntsP = @{}
    if ($null -ne $fetP.Resultat -and $null -ne $fetP.Resultat.PSObject.Properties['Geometries']) {
        $geos = $fetP.Resultat.Geometries; $puntsP = $fetP.Resultat.Punts
    }
    $cancelP = $fetP.Cancelat
    if ($cancelU -or $cancelP) { $avisos += "Has cancel" + [char]0x00B7 + "lat les consultes al Cadastre: algunes parcel" + [char]0x00B7 + "les surten com un punt. Torna-ho a generar per completar-les (el que ja s'ha demanat queda desat)." }
    # 4b. ELS PORTALS: nomes per dir a la fitxa l'adreca del Cadastre (l'ID va a
    # la coordenada de l'Excel d'activitats). Els mateixos que fa servir
    # Coordenades i la mateixa memoria cau (portals.json): el que ja s'ha demanat
    # alli no es torna a demanar.
    $fetE = $null
    if (-not ($cancelU -or $cancelP)) {
        $fetE = _PlanolAmbProgres $rcP ("Adreces al Cadastre (" + $rcP.Count + " parcel" + [char]0x00B7 + "les)...") {
            param($l, $p) Get-PortalsPerParcelles $l $p
        }
    }
    $portals = if ($null -ne $fetE) { $fetE.Resultat } else { @{} }
    if ($null -ne $fetE -and $fetE.Cancelat) { $avisos += "Has cancel" + [char]0x00B7 + "lat la cerca de les adreces al Cadastre: on no s'ha arribat, la fitxa no les diu." }
    $senseGeo = @($rcP | Where-Object { -not $geos.ContainsKey($_) -or @($geos[$_]).Count -eq 0 }).Count
    if ($senseGeo -gt 0 -and -not ($cancelU -or $cancelP)) {
        $avisos += "$senseGeo parcel" + [char]0x00B7 + "les sense dibuix del Cadastre: surten com un punt."
    }

    # 5. El mapa.
    $dades = ConvertTo-PlanolDadesMapa $model $geos $portals $puntsP
    # L'ID va a la coordenada UTM de l'Excel d'activitats (la que corregeix
    # Coordenades); en vermell, les que no hi caben (x 0) o no en tenen (x 3).
    $nFora = @(@($dades) | ForEach-Object { @($_.e) } | Where-Object { $_.t -eq 'a' -and $_.x -eq 0 } | ForEach-Object { $_.g } | Sort-Object -Unique).Count
    $nSense = @(@($dades) | ForEach-Object { @($_.e) } | Where-Object { $_.t -eq 'a' -and $_.x -eq 3 } | ForEach-Object { $_.g } | Sort-Object -Unique).Count
    if ($nFora -gt 0) {
        $avisos += "$nFora activitats amb la coordenada UTM fora de la seva parcel" + [char]0x00B7 + "la: l'ID GIA surt en vermell al centre (es pot corregir amb l'eina Coordenades)."
    }
    if ($nSense -gt 0) {
        $avisos += "$nSense activitats sense coordenada UTM a l'Excel d'activitats: l'ID GIA surt en vermell al centre."
    }
    if (-not $Script:PlanolTeCcae) {
        $avisos += "L'Excel d'activitats no t" + [char]0x00E9 + " la columna 'CCAE Codi': el filtre d'allotjaments tur" + [char]0x00ED + "stics no amaga res."
    }
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

    $res = $model.Resum
    $msg  = "Planol generat: $(@($dades).Count) parcel" + [char]0x00B7 + "les amb $($res.Activitats) activitats.`n`n"
    if ($res.SenseEstabliment -gt 0) { $msg += "Activitats sense cap establiment (situades amb la refer" + [char]0x00E8 + "ncia de l'Excel d'activitats): $($res.SenseEstabliment)`n" }
    if ($res.NoBase -gt 0) { $msg += "Establiments amb una activitat que no " + [char]0x00E9 + "s a la base d'activitats: $($res.NoBase)`n" }
    if ($res.BuitsDuplicats -gt 0) { $msg += "Locals buits duplicats al GIA (el mateix local, buit i amb activitat; no es compten com a buits): $($res.BuitsDuplicats)`n" }
    if ($res.MarcatsBuit -gt 0) { $msg += "Activitats en un local marcat com a buit (per revisar): $($res.MarcatsBuit)`n" }
    if ($res.SensePosicio -gt 0) { $msg += "Sense refer" + [char]0x00E8 + "ncia cadastral ni coordenades (no surten): $($res.SensePosicio)`n" }
    foreach ($a in $avisos) { $msg += "`n$a" }
    $msg += "`n`nFitxer: $outPath"
    return [pscustomobject]@{ Ok = $true; Error = ''; OutPath = $outPath; Missatge = $msg }
}

# Amb el boto (o la rajola) del menu: amb finestres, i l'obre.
# L'ultim planol generat (el mes nou de local\planol-activitats\), o $null.
function Get-PlanolUltim([string]$dir = $PlanolOutputDir) {
    if ([string]::IsNullOrWhiteSpace($dir) -or -not (Test-Path -LiteralPath $dir)) { return $null }
    $f = @(Get-ChildItem -LiteralPath $dir -Filter 'Planol_*.html' -File -ErrorAction SilentlyContinue |
           Sort-Object LastWriteTime -Descending | Select-Object -First 1)
    if ($f.Count -eq 0) { return $null }
    return $f[0]
}

# El text de la pregunta "consultar l'ultim o fer-ne un de nou".
function Get-PlanolPreguntaText($ultim, [datetime]$ara = (Get-Date)) {
    $dies = [int][math]::Floor(($ara.Date - $ultim.LastWriteTime.Date).TotalDays)
    $quan = if ($dies -le 0) { 'avui' } elseif ($dies -eq 1) { 'ahir' } else { "fa $dies dies" }
    $t  = "L'" + [char]0x00FA + "ltim pl" + [char]0x00E0 + "nol " + [char]0x00E9 + "s del " + $ultim.LastWriteTime.ToString('dd/MM/yyyy') + ' a les ' + $ultim.LastWriteTime.ToString('HH:mm') + " ($quan).`n`n"
    $t += "Vols consultar-lo o fer-ne un de nou? Fer-ne un de nou torna a llegir els Excel i pot trigar una estona."
    return $t
}

function Invoke-PlanolMain {
    # No cal fer-ne un de nou cada vegada que es prem la rajola (octubre 2026):
    # si ja n'hi ha un, es pregunta. Consultar-lo nomes l'obre.
    $ultim = Get-PlanolUltim
    if ($null -ne $ultim) {
        $tria = Show-EinaTria (Get-PlanolPreguntaText $ultim) @(
            @{ Nom = 'Cancel'; Text = ('Cancel' + [char]0x00B7 + 'lar'); Esc = $true }) @(
            @{ Nom = 'Nou'; Text = 'Fer-ne un de nou' },
            @{ Nom = 'Consultar'; Text = ("Consultar l'" + [char]0x00FA + 'ltim'); Estil = 'primari'; Intro = $true })
        if ($tria -eq 'Consultar') { Start-Process $ultim.FullName; return }
        if ($tria -ne 'Nou') { return }
    }
    $r = Invoke-PlanolGenera $false
    if (-not $r.Ok) {
        if ($r.Error -ne '') { Show-EinaInfo $r.Error '' 'Warning' }
        return
    }
    Start-Process $r.OutPath
    # La mateixa copia, al Drive privat per al mobil (docs/planol.html), en
    # segon pla: no fa esperar i, si falla, ho diu pujada-mobil.log.
    $pujada = Start-ScriptSegonPla (Join-Path (Split-Path -Parent $ScriptRoot) (Join-Path 'mobil' 'PujaPlanol.ps1')) @($r.OutPath)
    # Fet a ma: el segell de sota la rajola no surt en verd (vegeu
    # PlanolAutomatic.ps1). Nomes si s'ha obert des del programa.
    if (Get-Command _PlanolAutoDesaEstat -ErrorAction SilentlyContinue) { [void](_PlanolAutoDesaEstat @{ mode = 'manual' }) }
    $msg = $r.Missatge
    $msg += if ($null -ne $pujada) { "`nPer al mobil: se'n puja una copia al Drive (Dades/planol.html) en segon pla." } else { "`nPer al mobil: no s'ha pogut llancar la pujada al Drive." }
    Show-EinaInfo $msg
}

if (-not $Script:PlanolHeadless) {
    Invoke-PlanolMain
}
