#requires -Version 5.1
<#
.SYNOPSIS
  Eina "Copiar informes" (menu, fila INFORMES): copia PLANA i INCREMENTAL dels
  informes de $InformesDir a $CopiaInformesDir, a ma (la rajola) o en AUTOMATIC
  (l'interruptor A/M, que ho fa en un proces a part: CopiaInformesAuto.ps1).

  La feina viu en quatre funcions SENSE finestra (_CopiaInformesPrepara, -Cerca,
  -Tria, -Copia); el que difereix entre el manual i l'automatic (avisar,
  preguntar, la barra) es queda a la crida. Hi ha UNA sola copia de fitxer en aquest
  fitxer, i un guard ho vigila. Abans vivia a Informes.ps1 (revisio
  d'arquitectura, setembre 2026: cada fitxer, una cosa).
#>

# ============================================================================
# Copiar informes — eina INFORMES "Copiar informes"
# ============================================================================
# Còpia PLANA i incremental dels INFORMES Word de $InformesDir a
# $CopiaInformesDir:
#  - NOMÉS es consideren informes: *.doc / *.docx (ignora els temporals ~$...)
#    AMB DATA AL PRINCIPI DEL NOM (AAAA-MM-DD, AA_MM_DD, etc.) — mateix criteri
#    que "Actualitzar base" (_ParseDataInformeFromName). Qualsevol altre Word
#    (plantilles, esborranys, documents diversos) NO es copia;
#  - guarda la data de l'última còpia (copia-informes-state.json) i només mira
#    els fitxers modificats DESPRÉS;
#  - si el fitxer ja és al destí (mateix nom), NO es torna a copiar;
#  - MAI esborra res del destí (còpia additiva).
#
# LA MATEIXA CÒPIA ES FA DE DUES MANERES, i per això la feina viu en quatre
# trossos que no saben res de cap finestra (_CopiaInformesPrepara,
# _CopiaInformesCerca, _CopiaInformesTria i _CopiaInformesCopia):
#
#   MANUAL     Invoke-CopiarInformes — la rajola del menú. Finestra de progrés
#              amb CANCEL·LAR i confirmació amb el nombre d'informes abans de
#              copiar: mai comença "a cegues".
#   AUTOMÀTIC  Invoke-CopiarInformesAuto — l'interruptor A/M del menú. Cap
#              finestra i cap pregunta, i corre en un PROCÉS A PART
#              (CopiaInformesAuto.ps1) perquè recórrer la carpeta d'informes
#              (unitat de xarxa, recursiu) no deixi el menú congelat.
#
# El que difereix —avisar, preguntar, pintar la barra— es queda a la crida, en
# scriptblocks; la còpia de debò està escrita UN sol cop.

# ----------------------------------------------------------------------------
# L'ESTAT: copia-informes-state.json
# ----------------------------------------------------------------------------
#   generat_el  quan es va crear l'estat (informatiu, ve de sempre)
#   copiat_el   l'última còpia ACABADA  ·  desti: a quina carpeta anava
#   mode        'auto' | 'manual': qui va fer aquella última còpia. El menú
#               pinta la data en VERD si va ser automàtica i en gris si no.
#   auto        l'interruptor A/M del menú (mode automàtic engegat o no)
#   auto_el     l'última PASSADA automàtica, encara que no copiés res. És el
#               que evita que es repeteixi a cada obertura del menú, i el que
#               fa que una passada perduda (PC apagat) es recuperi en obrir.
function _CopiaInformesStatePath {
    if ([string]::IsNullOrWhiteSpace($LocalActivitatsDir)) { return '' }
    return [string](Join-Path $LocalActivitatsDir 'copia-informes-state.json')
}

# Retorna SEMPRE el diccionari sencer (buit si no hi ha fitxer o està corrupte),
# així cap crider ha de comprovar si una clau hi és abans de llegir-la.
function _CopiaInformesEstat {
    $out = [ordered]@{ generat_el = ''; copiat_el = ''; desti = ''; mode = ''; auto = $false; auto_el = '' }
    $o = Read-JsonFile (_CopiaInformesStatePath)
    if ($null -ne $o) {
        foreach ($k in @('generat_el', 'copiat_el', 'desti', 'mode', 'auto_el')) {
            if ($o.PSObject.Properties[$k]) { $out[$k] = [string](Read-JsonIso $o.$k) }
        }
        if ($o.PSObject.Properties['auto']) { $out['auto'] = [bool]$o.auto }
    }
    return $out
}

# Desa NOMÉS les claus que li passes, damunt del que ja hi ha al fitxer. Mai
# llança: si la carpeta no hi és (unitat de xarxa fora de servei) el programa ha
# de seguir funcionant igual, que és el mateix criteri que _MarcaEinaUsada.
function _CopiaInformesDesaEstat($canvis) {
    try {
        $p = _CopiaInformesStatePath
        if ([string]::IsNullOrWhiteSpace($p)) { return $false }
        $est = _CopiaInformesEstat
        if ($null -ne $canvis) { foreach ($k in @($canvis.Keys)) { $est[$k] = $canvis[$k] } }
        if ([string]::IsNullOrWhiteSpace([string]$est['generat_el'])) { $est['generat_el'] = (Get-Date).ToString('o') }
        $dir = Split-Path -Parent $p
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        Write-JsonFile $p ([pscustomobject]$est) 5
        return $true
    } catch { return $false }
}

# ----------------------------------------------------------------------------
# EL NUCLI (sense cap finestra): el comparteixen el mode manual i l'automàtic
# ----------------------------------------------------------------------------
# Comprova la configuració i prepara la carpeta destí.
# Retorna @{ Ok; Error; Icona } — el manual en fa un MessageBox amb la seva
# icona; l'automàtic escriu l'error al registre i no fa res més.
function _CopiaInformesPrepara {
    if ([string]::IsNullOrWhiteSpace($CopiaInformesDir)) {
        return @{ Ok = $false; Icona = 'Information'; Error = "No s'ha configurat cap carpeta de còpia.`n`nVes a  ⚙ Configuració  i indica 'Carpeta on copiar els informes'." }
    }
    if ([string]::IsNullOrWhiteSpace($InformesDir) -or -not (Test-Path -LiteralPath $InformesDir -ErrorAction SilentlyContinue)) {
        return @{ Ok = $false; Icona = 'Warning'; Error = "No s'ha trobat la carpeta d'informes:`n$InformesDir" }
    }
    if ([System.IO.Path]::GetFullPath($CopiaInformesDir).TrimEnd('\') -ieq [System.IO.Path]::GetFullPath($InformesDir).TrimEnd('\')) {
        return @{ Ok = $false; Icona = 'Warning'; Error = "La carpeta de còpia no pot ser la mateixa que la carpeta d'informes." }
    }
    try {
        if (-not (Test-Path -LiteralPath $CopiaInformesDir)) {
            New-Item -ItemType Directory -Path $CopiaInformesDir -Force -ErrorAction Stop | Out-Null
        }
    } catch {
        return @{ Ok = $false; Icona = 'Error'; Error = "No s'ha pogut crear la carpeta destí:`n$($_.Exception.Message)" }
    }
    return @{ Ok = $true; Icona = 'Information'; Error = '' }
}

# Des de quina data (UTC) mirem els fitxers: només els modificats DESPRÉS de
# l'última còpia. Si el destí desat NO és el d'ara, es torna a mirar TOT
# ([datetime]::MinValue): la carpeta nova no té res a dins. Funció PURA.
function _CopiaInformesDesDe($est, [string]$desti) {
    if ($null -eq $est) { return [datetime]::MinValue }
    if ([string]$est['desti'] -ne [string]$desti) { return [datetime]::MinValue }
    $t = [string]$est['copiat_el']
    if ([string]::IsNullOrWhiteSpace($t)) { return [datetime]::MinValue }
    try { return ([datetime]::Parse($t)).ToUniversalTime() } catch { return [datetime]::MinValue }
}

# Enumera els INFORMES de $InformesDir. Retorna @{ Files; Cancelled }.
#
# $onProgres (opcional) es crida cada 200 fitxers explorats amb
# (explorats, trobats) i ha de tornar $false per aturar la cerca: el mode manual
# hi refresca la finestra i hi mira el botó Cancel·lar, i l'automàtic no en
# passa cap (no hi ha res per cancel·lar).
#
# EL 'do { } while ($false)' NO ES DECORATIU. Un 'break' dins d'un ForEach-Object
# atura el pipeline i, si no troba un bucle SEU, se'n va cap amunt i trenca el
# bucle de qui ha cridat la funció: cancel·lar la cerca es carregava el
# 'while ($true)' de Main i tancava el programa sencer. El bucle fals li dona
# aquí mateix un bucle per trencar. (Comprovat executant-ho, no deduït.)
function _CopiaInformesCerca($onProgres) {
    $files = New-Object System.Collections.ArrayList
    $seen = 0
    $stop = $false
    do {
        Get-ChildItem -LiteralPath $InformesDir -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
            $seen++
            if ($null -ne $onProgres -and ($seen % 200) -eq 0) {
                if (-not (& $onProgres $seen $files.Count)) { $stop = $true; break }
            }
            if ($_.Name -notlike '~$*' -and
                ($_.Extension -ieq '.docx' -or $_.Extension -ieq '.doc') -and
                $null -ne (_ParseDataInformeFromName $_.Name)) {
                [void]$files.Add($_)
            }
        }
    } while ($false)
    return @{ Files = $files; Cancelled = $stop }
}

# Quins cal copiar de veritat: els modificats després de l'última còpia i que
# encara no siguin al destí (mai recopiem ni esborrem).
# Retorna @{ ToCopy; Skipped }.
function _CopiaInformesTria($files, [datetime]$desDe, [string]$desti) {
    $toCopy = New-Object System.Collections.ArrayList
    $skipped = 0
    foreach ($f in @($files)) {
        if ($f.LastWriteTimeUtc -le $desDe) { continue }
        $dest = [System.IO.Path]::Combine($desti, $f.Name)
        if (Test-Path -LiteralPath $dest) { $skipped++; continue }
        [void]$toCopy.Add($f)
    }
    return @{ ToCopy = $toCopy; Skipped = $skipped }
}

# Copia la llista. $onProgres (opcional) es crida a cada fitxer amb
# (fets, total, copiats, errors) i ha de tornar $false per aturar.
# Retorna @{ Copied; Errors; Cancelled }.
function _CopiaInformesCopia($toCopy, [string]$desti, $onProgres) {
    $copied = 0; $errors = 0; $done = 0
    $cancelled = $false
    foreach ($f in @($toCopy)) {
        $dest = [System.IO.Path]::Combine($desti, $f.Name)
        try { Copy-Item -LiteralPath $f.FullName -Destination $dest -ErrorAction Stop; $copied++ }
        catch { $errors++ }
        $done++
        if ($null -ne $onProgres) {
            if (-not (& $onProgres $done @($toCopy).Count $copied $errors)) { $cancelled = $true; break }
        }
    }
    return @{ Copied = $copied; Errors = $errors; Cancelled = $cancelled }
}

# ----------------------------------------------------------------------------
# MODE MANUAL (la rajola del menú): finestra de progrés + confirmació
# ----------------------------------------------------------------------------
# La rajola SEMPRE fa la còpia, tant si l'interruptor està en A com en M: el
# mode automàtic no treu res a l'usuari, només afegeix una passada sola cada dia.
function Invoke-CopiarInformes {
    $prep = _CopiaInformesPrepara
    if (-not $prep.Ok) {
        [System.Windows.Forms.MessageBox]::Show([string]$prep.Error, 'Copiar informes', 'OK', [string]$prep.Icona) | Out-Null
        return
    }

    $est   = _CopiaInformesEstat
    $desDe = _CopiaInformesDesDe $est $CopiaInformesDir

    # ---- Finestra de progrés amb CANCEL·LAR --------------------------------
    # Running: mentre és cert, la X de la finestra es tracta com a "cancel·lar"
    # (no es tanca de debò fins al 'finally'), així el bucle no toca mai controls
    # ja destruïts.
    $cancel = @{ Flag = $false; Running = $true }
    $form = _NewForm
    $form.Text = 'Copiar informes'
    $form.Size = New-Object System.Drawing.Size(560, 190)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(20, 18)
    $lbl.Size = New-Object System.Drawing.Size(510, 60)
    $lbl.Text = "Cercant informes a:`n$InformesDir"
    $form.Controls.Add($lbl)
    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Location = New-Object System.Drawing.Point(20, 88)
    $bar.Size = New-Object System.Drawing.Size(510, 22)
    $bar.Style = 'Marquee'
    $form.Controls.Add($bar)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel·lar'
    $btnCancel.Size = New-Object System.Drawing.Size(120, 30)
    $btnCancel.Location = New-Object System.Drawing.Point(410, 118)
    _StyleSecondaryButton $btnCancel
    $btnCancel.add_Click({ $cancel.Flag = $true }.GetNewClosure())
    $form.Controls.Add($btnCancel)
    $form.add_FormClosing({
        param($s, $e)
        if ($cancel.Running) { $cancel.Flag = $true; $e.Cancel = $true }  # X = cancel·lar; el tancament real el fa el 'finally'
    }.GetNewClosure())
    $form.Show()
    [System.Windows.Forms.Application]::DoEvents()

    $copied = 0; $skipped = 0; $errors = 0
    $cancelled = $false
    $nothing = $false
    $trobats = 0
    try {
        # 1. Enumerar només els INFORMES. El nucli ens avisa cada 200 fitxers
        #    perquè la finestra respongui i es pugui cancel·lar.
        # $nTrobats i no $trobats: el parametre TAPARIA la variable de fora que
        # despres fa servir el missatge final, i seria una d'aquelles coses que
        # no fallen, nomes surten malament.
        $cerca = _CopiaInformesCerca {
            param($vistos, $nTrobats)
            $lbl.Text = "Cercant informes...`n$vistos fitxers explorats  ·  $nTrobats informes trobats"
            [System.Windows.Forms.Application]::DoEvents()
            return (-not $cancel.Flag)
        }.GetNewClosure()
        $trobats = @($cerca.Files).Count

        if ($cerca.Cancelled -or $cancel.Flag) {
            $cancelled = $true
        } else {
            # 2. Quins cal copiar de veritat.
            $lbl.Text = "Preparant la llista d'informes a copiar..."
            [System.Windows.Forms.Application]::DoEvents()
            $tria = _CopiaInformesTria $cerca.Files $desDe $CopiaInformesDir
            $skipped = [int]$tria.Skipped
            $toCopy = @($tria.ToCopy)

            if ($toCopy.Count -eq 0) {
                $nothing = $true
            } else {
                # 3. Confirmació explícita ABANS de copiar res.
                $rc = [System.Windows.Forms.MessageBox]::Show(
                    "Es copiaran $($toCopy.Count) informes (de $trobats trobats; $skipped ja hi són).`n`nDestí: $CopiaInformesDir`n`nVols continuar?",
                    'Copiar informes', 'YesNo', 'Question')
                if ($rc -ne [System.Windows.Forms.DialogResult]::Yes) {
                    $cancelled = $true
                } else {
                    # 4. Còpia amb barra de progrés i cancel·lació.
                    $bar.Style = 'Continuous'; $bar.Minimum = 0; $bar.Maximum = [Math]::Max(1, $toCopy.Count); $bar.Value = 0
                    $res = _CopiaInformesCopia $toCopy $CopiaInformesDir {
                        param($fets, $total, $copiats, $fallats)
                        if ($bar.Value -lt $bar.Maximum) { $bar.Value = $fets }
                        $lbl.Text = "Copiant informes...  $fets de $total`nCopiats: $copiats   Errors: $fallats"
                        [System.Windows.Forms.Application]::DoEvents()
                        return (-not $cancel.Flag)
                    }.GetNewClosure()
                    $copied = [int]$res.Copied
                    $errors = [int]$res.Errors
                    if ($res.Cancelled) { $cancelled = $true }
                }
            }
        }
    } finally {
        $cancel.Running = $false   # permet el tancament real de la finestra
        try { $form.Close() } catch { }
    }

    if ($nothing) {
        [System.Windows.Forms.MessageBox]::Show(
            "No hi ha informes nous per copiar.`n`nInformes trobats: $trobats  (ja copiats: $skipped)`nDestí: $CopiaInformesDir",
            'Copiar informes', 'OK', 'Information') | Out-Null
        return
    }

    # 5. Desem l'estat NOMÉS si s'ha completat (si s'ha cancel·lat, deixem la
    #    data com estava perquè la propera vegada es tornin a comprovar els que
    #    faltaven; els ja copiats se saltaran igualment per existència).
    if (-not $cancelled) {
        _CopiaInformesDesaEstat @{ copiat_el = (Get-Date).ToString('o'); desti = $CopiaInformesDir; mode = 'manual' } | Out-Null
    }

    $titol = if ($cancelled) { 'Còpia cancel·lada' } else { 'Còpia completada' }
    [System.Windows.Forms.MessageBox]::Show(
        "$titol`n`nInformes copiats: $copied`nJa existents (omesos): $skipped`nErrors: $errors`n`nDestí: $CopiaInformesDir",
        'Copiar informes', 'OK', 'Information') | Out-Null
}

# ----------------------------------------------------------------------------
# MODE AUTOMÀTIC: quan toca (la part PURA)
# ----------------------------------------------------------------------------
# L'usuari ho va demanar així: cada dia a les 14:30 si el programa està obert i,
# si el dia abans no es va arribar a fer, en obrir el programa.
#
# LES DUES COSES SÓN LA MATEIXA PREGUNTA i per això hi ha una sola funció: "des
# de l'últim VENCIMENT (les 14:30 que toquen), s'ha fet cap passada automàtica?"
#   · menú obert a les 14:30      -> el venciment passa a ser el d'avui i toca
#   · ahir el PC estava apagat    -> el venciment d'ahir no es va servir, toca
#   · obres a les 16:00 i el d'avui no s'ha fet -> toca (no s'espera a demà)
# Amb dues regles separades (una per al rellotge i una per a l'arrencada) el
# tercer cas es perdia fins l'endemà.
$Script:CopiaAutoHora  = 14
$Script:CopiaAutoMinut = 30

# L'últim venciment que ja hauria d'estar servit a l'hora $ara. PURA.
function _CopiaAutoVenciment([datetime]$ara) {
    $avui = New-Object datetime($ara.Year, $ara.Month, $ara.Day, $Script:CopiaAutoHora, $Script:CopiaAutoMinut, 0)
    if ($ara -lt $avui) { return $avui.AddDays(-1) }
    return $avui
}

# $ultimAuto: la marca 'auto_el' de l'estat (text ISO; buida si no s'ha fet mai).
# PURA: no llegeix el disc ni el rellotge, per poder-la provar.
function _CopiaAutoToca([datetime]$ara, $ultimAuto) {
    $venc = _CopiaAutoVenciment $ara
    $t = [string]$ultimAuto
    if ([string]::IsNullOrWhiteSpace($t)) { return $true }
    try { $fet = [datetime]::Parse($t) } catch { return $true }
    return ($fet -lt $venc)
}

# L'interruptor A/M del menú.
function _CopiaAutoActiu {
    $est = _CopiaInformesEstat
    return [bool]$est['auto']
}

function _CopiaAutoDesaActiu([bool]$on) {
    return (_CopiaInformesDesaEstat @{ auto = $on })
}

# 'auto' | 'manual' | '' : qui va fer l'última còpia (el menú hi pinta la data
# en verd o en gris).
function _CopiaInformesUltimMode {
    $est = _CopiaInformesEstat
    return [string]$est['mode']
}

# ----------------------------------------------------------------------------
# MODE AUTOMÀTIC: la passada (sense cap finestra)
# ----------------------------------------------------------------------------
# La crida CopiaInformesAuto.ps1, que és qui corre en segon pla. Retorna
# @{ Ok; Copiats; Omesos; Errors; Trobats; Motiu }.
#
# APUNTA 'auto_el' SEMPRE, fins i tot quan no hi ha res a copiar o la
# configuració no hi és: és la marca que diu que el venciment d'avui ja s'ha
# servit. Si només s'apuntés quan copia, el menú tornaria a llançar la passada
# cada minut i cada vegada que hi tornessis.
function Invoke-CopiarInformesAuto {
    $ini = @{ auto_el = (Get-Date).ToString('o') }
    $prep = _CopiaInformesPrepara
    if (-not $prep.Ok) {
        _CopiaInformesDesaEstat $ini | Out-Null
        _CopiaAutoLog ('ATURAT: ' + ([string]$prep.Error -replace "`r?`n", ' '))
        return @{ Ok = $false; Copiats = 0; Omesos = 0; Errors = 0; Trobats = 0; Motiu = [string]$prep.Error }
    }

    $est   = _CopiaInformesEstat
    $desDe = _CopiaInformesDesDe $est $CopiaInformesDir
    $cerca = _CopiaInformesCerca $null
    $tria  = _CopiaInformesTria $cerca.Files $desDe $CopiaInformesDir
    $toCopy = @($tria.ToCopy)
    $res = _CopiaInformesCopia $toCopy $CopiaInformesDir $null

    # 'mode' i 'copiat_el' NOMÉS quan ha copiat alguna cosa: si una passada que
    # no troba res reescrivís el mode, la data verda del menú deixaria de dir
    # quan es va copiar per última vegada i qui ho va fer.
    $canvis = $ini
    if ([int]$res.Copied -gt 0) {
        $canvis = @{ auto_el = $ini['auto_el']; copiat_el = (Get-Date).ToString('o'); desti = $CopiaInformesDir; mode = 'auto' }
    }
    _CopiaInformesDesaEstat $canvis | Out-Null

    _CopiaAutoLog ("Passada automatica: trobats=$(@($cerca.Files).Count) copiats=$($res.Copied) " +
                   "omesos=$($tria.Skipped) errors=$($res.Errors) desti=$CopiaInformesDir")
    return @{ Ok = $true; Copiats = [int]$res.Copied; Omesos = [int]$tria.Skipped
              Errors = [int]$res.Errors; Trobats = @($cerca.Files).Count; Motiu = '' }
}

# Registre de diagnòstic del mode automàtic (com el dels recordatoris: si no es
# veu res, l'única manera de saber què ha passat és aquest fitxer).
function _CopiaAutoLogPath {
    return [string](Join-Path (Join-Path $env:LOCALAPPDATA 'InformesCornella') 'copia-informes-log.txt')
}

function _CopiaAutoLog([string]$msg) {
    try {
        $p = _CopiaAutoLogPath
        $dir = Split-Path -Parent $p
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        Add-Content -LiteralPath $p -Value ('[' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') + '] ' + [string]$msg) -Encoding UTF8
    } catch { }
}

# Llança la passada automàtica EN SEGON PLA i torna de seguida. El menú no es
# pot quedar congelat mentre es recorre la carpeta d'informes, i tampoc no pot
# fer la còpia dins d'un tick del rellotge: amb els DoEvents de WinForms
# l'usuari podria obrir una eina a mig copiar.
#
# NOMÉS UN A LA VEGADA: si l'anterior encara corre (el menú es torna a obrir a
# cada volta de Main), no se'n llança un altre.
$Script:CopiaAutoProc = $null

function Start-CopiaInformesAuto {
    try {
        if ($null -ne $Script:CopiaAutoProc -and -not $Script:CopiaAutoProc.HasExited) { return $false }
    } catch { $Script:CopiaAutoProc = $null }
    $Script:CopiaAutoProc = Start-ScriptSegonPla 'CopiaInformesAuto.ps1'
    return ($null -ne $Script:CopiaAutoProc)
}

# El menú ho crida en obrir-se i a cada minut: si l'interruptor està en A i el
# venciment de les 14:30 encara no s'ha servit, engega la passada.
function Invoke-CopiaAutoSiToca {
    $est = _CopiaInformesEstat
    if (-not [bool]$est['auto']) { return $false }
    if (-not (_CopiaAutoToca (Get-Date) $est['auto_el'])) { return $false }
    return (Start-CopiaInformesAuto)
}
