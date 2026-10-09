#requires -Version 5.1
<#
.SYNOPSIS
  La base d'informes: llegir-ne els noms, l'edicio A MA i la pantalla
  "Editar base d'informes".

.DESCRIPTION
  ATENCIO: aqui NO hi ha l'escaner. Recorrer les carpetes, llegir cada .docx i
  muntar informes-db.json es InformesEscaneig.ps1, i decidir que diu cada
  informe es InformesClassificacio.ps1. Aquest fitxer te el que queda:

    - EL NOM DEL FITXER: la data del principi (_ParseDataInformeFromName, que
      es alhora la definicio de "aixo es un informe"), l'ID GIA, el numero
      d'expedient i la carpeta de l'activitat.
    - L'EDICIO A MA: les marques editat_a_ma / auto_* , la fusio amb el que hi
      hagi al disc si algu ha escanejat mentrestant (_FusionaEdicionsBase) i el
      desat sota el mutex (Save-BaseEditada).
    - LA PANTALLA "Editar base d'informes" (graella WinForms) i l'exportacio a
      CSV dels estats.

  Les funcions de text son PURES (operen sobre cadenes) per poder provar-les en
  headless, a Linux i sense Word.

  La ruta del fitxer surt de Get-InformesDbPath (Migracio.ps1), al costat dels
  noms de les carpetes de local\: abans estava escrita a cinc llocs.

.NOTES
  La pantalla es llanca des del menu (Pas 1), rajola "Editar base d'informes".
  "Actualitzar base d'informes" es l'altra eina, i es a InformesEscaneig.ps1.
#>

# ----------------------------------------------------------------------------
# Logica de text (funcions PURES, testejables en headless)
# ----------------------------------------------------------------------------

# Treu la data del PRINCIPI del nom del fitxer. Tolerant amb el format:
#   AAAA-MM-DD / AAAAMMDD / AAAA.MM.DD / AAAA_MM_DD  (any de 4 xifres)
#   AA-MM-DD   / AA.MM.DD / AA_MM_DD                 (any de 2 xifres, amb
#                                                     separador obligatori per no
#                                                     confondre-ho amb AAAAMMDD)
# Retorna 'yyyy-MM-dd' o $null si el nom no comenca amb una data valida.
# (El programa considera "informe" NOMES els fitxers que comencen amb data.)
function _ParseDataInformeFromName($name) {
    if ([string]::IsNullOrWhiteSpace($name)) { return $null }
    $n = [string]$name
    # Any de 4 xifres, separadors opcionals.
    if ($n -match '^(\d{4})[-_.]?(\d{2})[-_.]?(\d{2})(?:\D|$)') {
        $y = [int]$Matches[1]; $mo = [int]$Matches[2]; $d = [int]$Matches[3]
        if ($y -ge 1990 -and $y -le 2100 -and $mo -ge 1 -and $mo -le 12 -and $d -ge 1 -and $d -le 31) {
            return ('{0:D4}-{1:D2}-{2:D2}' -f $y, $mo, $d)
        }
    }
    # Any de 2 xifres, separador OBLIGATORI.
    if ($n -match '^(\d{2})[-_.](\d{2})[-_.](\d{2})(?:\D|$)') {
        $y = 2000 + [int]$Matches[1]; $mo = [int]$Matches[2]; $d = [int]$Matches[3]
        if ($mo -ge 1 -and $mo -le 12 -and $d -ge 1 -and $d -le 31) {
            return ('{0:D4}-{1:D2}-{2:D2}' -f $y, $mo, $d)
        }
    }
    return $null
}

# Cert si el valor trobat despres de "ID GIA:" es un placeholder (activitat
# encara sense GIA assignat: "-", "XXX", "N/A"...) i no un ID real. Cal
# distingir-ho perque, si no, activitats totalment diferents que encara no
# tenen GIA queden ajuntades sota una mateixa "activitat" fantasma amb
# id_gia="-" (vist a la carpeta real d'informes).
function _EsGiaPlaceholder($val) {
    if ([string]::IsNullOrWhiteSpace($val)) { return $true }
    $n = $val.Trim().ToLowerInvariant()
    return ($n -eq '-' -or $n -eq '--' -or $n -eq '---' -or $n -eq 'xxx' -or $n -eq 'n/a' -or $n -eq 'na' -or $n -eq '?')
}

# ID GIA d'una llista de linies de text del document. Busca la linia que conte
# "ID GIA" i en retorna el valor (despres dels dos punts / espais). '' si no hi
# es, o si el valor trobat es un placeholder de "encara sense GIA".
function _ExtractIdGia($lines) {
    foreach ($ln in $lines) {
        $s = [string]$ln
        if ($s -imatch 'ID\s*GIA\s*:?\s*(.+)$') {
            $val = $Matches[1].Trim()
            # Ens quedem nomes amb el primer "token" del valor (l'ID; evita
            # arrossegar text si la linia porta res mes al darrere).
            $token = $val
            if ($val -match '^([\w./-]+)') { $token = $Matches[1] }
            if (_EsGiaPlaceholder $token) { continue }
            return $token
        }
    }
    return ''
}

# Numero d'expedient d'una llista de linies. Busca la linia que comenca per
# "Exp" (p.ex. "Exp. Num: 2025/1/2563") i retorna el valor despres dels dos punts.
function _ExtractExpedient($lines) {
    foreach ($ln in $lines) {
        $s = [string]$ln
        $nt = _NormalitzaText $s
        if ($nt -match '^exp') {
            $idx = $s.IndexOf(':')
            if ($idx -ge 0) { return $s.Substring($idx + 1).Trim() }
            if ($s -match '(\S+/\S+/\S+)') { return $Matches[1] }
        }
    }
    return ''
}

# Les conclusions, la conclusio breu i l'estat de l'activitat viuen a
# InformesClassificacio.ps1 (octubre 2026: amb els formats antics i els tipus
# d'informe aquest fitxer passava de les 1.200 linies).

# ----------------------------------------------------------------------------
# EDICIONS A MA ("Editar base") -- octubre 2026, peticio de l'usuari: el que
# s'ha corregit a ma PREVAL sobre el que surti d'"Actualitzar base", i l'estat
# d'aquelles activitats surt en VERMELL a l'editor perque se'n sigui conscient.
#
# Es marca a l'INFORME (no a l'activitat): editat_a_ma = $true, i es guarda el
# valor automatic que hi havia (auto_conclusio_breu / auto_ignorat) per poder
# desfer-ho. Abans de la marca, l'escaneig ja conservava la conclusio breu i
# l'ignorat de TOTS els informes reprocessats -editats o no-, o sigui que un
# informe que es tornava a escriure mai no actualitzava la seva conclusio breu.
# ----------------------------------------------------------------------------

# Marca un informe com a editat a ma, guardant ABANS el valor automatic. Si ja
# ho estava, no toca res (el valor automatic que es guarda es el primer).
function _MarcaEditatAMa($inf) {
    if ([bool](_PropInf $inf 'editat_a_ma')) { return }
    Add-Member -InputObject $inf -NotePropertyName auto_conclusio_breu -NotePropertyValue ([string](_PropInf $inf 'conclusio_breu')) -Force
    Add-Member -InputObject $inf -NotePropertyName auto_ignorat -NotePropertyValue ([bool](_PropInf $inf 'ignorat')) -Force
    Add-Member -InputObject $inf -NotePropertyName editat_a_ma -NotePropertyValue $true -Force
}

# Torna un informe al que diu l'automatic. Retorna $true si ho ha fet.
function _DesfesEditatAMa($inf) {
    if (-not [bool](_PropInf $inf 'editat_a_ma')) { return $false }
    Add-Member -InputObject $inf -NotePropertyName conclusio_breu -NotePropertyValue ([string](_PropInf $inf 'auto_conclusio_breu')) -Force
    Add-Member -InputObject $inf -NotePropertyName ignorat -NotePropertyValue ([bool](_PropInf $inf 'auto_ignorat')) -Force
    Add-Member -InputObject $inf -NotePropertyName editat_a_ma -NotePropertyValue $false -Force
    [void]$inf.PSObject.Properties.Remove('auto_conclusio_breu')
    [void]$inf.PSObject.Properties.Remove('auto_ignorat')
    return $true
}

# Una base d'ABANS de la marca: un informe la conclusio breu del qual no es la
# que en surt del text NOMES pot ser una correccio a ma (l'automatic sempre la
# treu del text). Se li posa la marca, amb el valor automatic que en surt.
# L'ignorat, aqui, no es pot deduir (el per defecte depen de coses que la base
# no guarda); el dedueix l'escaneig, que si que les te. Un informe que ja porta
# la marca (encara que sigui $false) no es toca.
function _InferEditatAMa($inf) {
    if ($null -ne $inf.PSObject.Properties['editat_a_ma']) { return }
    $auto = _ConclusioBreu ([string](_PropInf $inf 'conclusio'))
    $actual = [string](_PropInf $inf 'conclusio_breu')
    if ($actual -ne '' -and $actual -ne $auto) {
        Add-Member -InputObject $inf -NotePropertyName auto_conclusio_breu -NotePropertyValue $auto -Force
        Add-Member -InputObject $inf -NotePropertyName auto_ignorat -NotePropertyValue ([bool](_PropInf $inf 'ignorat')) -Force
        Add-Member -InputObject $inf -NotePropertyName editat_a_ma -NotePropertyValue $true -Force
    } else {
        Add-Member -InputObject $inf -NotePropertyName editat_a_ma -NotePropertyValue $false -Force
    }
}

# Alguna conclusio/ignorat d'aquesta activitat s'ha corregit a ma?
function _ActivitatEditadaAMa($act) {
    foreach ($inf in @(_PropInf $act 'informes')) { if ($null -ne $inf -and [bool](_PropInf $inf 'editat_a_ma')) { return $true } }
    return $false
}

# La CLAU d'un informe per casar el d'una base amb el d'una altra (o amb el
# fitxer del disc): la ruta RELATIVA a la carpeta d'informes de la base
# (carpeta_arrel). Abans era la ruta absoluta, i la mateixa base feta servir amb
# la carpeta en una altra unitat (I:\...\Informes a la feina, F:\...\Informes
# fora) no casava cap informe: es reprocessava tot i es perdien TOTES les
# correccions a ma sense cap avis. Si la ruta no es dins de l'arrel (o la base
# no en diu cap), la ruta sencera. Les dues barres valen igual. PURA.
function _ClauInforme([string]$ruta, [string]$arrel) {
    $r = $ruta -replace '/', '\'
    $a = ($arrel -replace '/', '\').TrimEnd('\')
    if ($a -ne '' -and $r.Length -gt $a.Length + 1 -and $r.StartsWith($a + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
        return $r.Substring($a.Length + 1)
    }
    return $r
}

# ----------------------------------------------------------------------------
# ON es busquen els informes: NOMES el primer nivell
# ----------------------------------------------------------------------------
# Els fitxers que son DIRECTAMENT dins de cada carpeta d'activitat de la
# carpeta d'informes (<arrel>\<carpeta>\fitxer). Ni els de l'arrel ni els de
# les subcarpetes (l'usuari, octubre 2026: "nomes ha de tenir en compte els
# informes de les carpetes dins la carpeta d'informes, pero no dins les
# subcarpetes"). A les subcarpetes hi ha l'expedient del GIA descarregat
# (~25.000 fitxers) i informes vells o d'altres activitats que no han de
# decidir l'estat; abans un Get-ChildItem -Recurse els recorria tots. El fan
# servir la base d'informes i el repas de contactes (per aixo es aqui i no a
# InformesEscaneig.ps1: el repas de contactes no pot dependre del seu client).
function Get-FitxersPrimerNivell([string]$dir) {
    $out = New-Object System.Collections.ArrayList
    foreach ($d in @(Get-ChildItem -LiteralPath $dir -Directory -ErrorAction SilentlyContinue)) {
        foreach ($f in @(Get-ChildItem -LiteralPath $d.FullName -File -ErrorAction SilentlyContinue)) { [void]$out.Add($f) }
    }
    return ,$out.ToArray()
}

# La clau (_ClauInforme) es d'un fitxer de primer nivell? "carpeta\fitxer",
# exactament una barra. PURA.
function _EsDePrimerNivell([string]$clau) {
    return (([string]$clau -replace '/', '\').Trim('\').Split('\').Count -eq 2)
}


# DESAR L'EDITOR QUAN LA BASE HA CANVIAT MENTRE ERA OBERT. Amb "Actualitzar base"
# en automatic (en segon pla), la base del disc pot ser mes nova que la que
# l'editor va carregar: desar-la tal qual tornaria enrere els informes nous.
# Es posen les correccions de l'editor ($editor) damunt de la del disc ($disc),
# informe a informe (per la ruta relativa, _ClauInforme):
#   corregit a l'editor          -> la correccio (el valor automatic, el del disc)
#   desfet a l'editor ("Desfer canvi a ma") i al disc encara corregit, amb un
#   valor diferent               -> torna a l'automatic del disc
# i es recalcula l'estat de les activitats tocades. Torna $disc (modificat).
function _FusionaEdicionsBase($disc, $editor) {
    $perRuta = @{}
    $arrelEd = [string](_PropInf $editor 'carpeta_arrel')
    $arrelDisc = [string](_PropInf $disc 'carpeta_arrel')
    foreach ($act in @(_PropInf $editor 'activitats')) {
        foreach ($inf in @(_PropInf $act 'informes')) {
            $r = [string](_PropInf $inf 'ruta')
            if ($r -ne '') { $perRuta[(_ClauInforme $r $arrelEd)] = $inf }
        }
    }
    foreach ($act in @(_PropInf $disc 'activitats')) {
        if ($null -eq $act) { continue }
        $tocat = $false
        foreach ($inf in @(_PropInf $act 'informes')) {
            if ($null -eq $inf) { continue }
            $r = [string](_PropInf $inf 'ruta')
            if ($r -eq '') { continue }
            $r = _ClauInforme $r $arrelDisc
            if (-not $perRuta.ContainsKey($r)) { continue }
            $ed = $perRuta[$r]
            $edBreu = [string](_PropInf $ed 'conclusio_breu')
            $edIgn = [bool](_PropInf $ed 'ignorat')
            if ([bool](_PropInf $ed 'editat_a_ma')) {
                _MarcaEditatAMa $inf
                Add-Member -InputObject $inf -NotePropertyName conclusio_breu -NotePropertyValue $edBreu -Force
                Add-Member -InputObject $inf -NotePropertyName ignorat -NotePropertyValue $edIgn -Force
                $tocat = $true
            } elseif ([bool](_PropInf $inf 'editat_a_ma') -and
                      ($edBreu -ne [string](_PropInf $inf 'conclusio_breu') -or $edIgn -ne [bool](_PropInf $inf 'ignorat'))) {
                [void](_DesfesEditatAMa $inf)
                $tocat = $true
            }
        }
        if ($tocat) { Add-Member -InputObject $act -NotePropertyName estat_actual -NotePropertyValue (_EstatActualActivitat $act) -Force }
    }
    return $disc
}

# La marca de quan es va escriure la base (per saber si ha canviat). PURA.
function _SegellBase($db) { return [string](Read-JsonIso (_PropInf $db 'actualitzat_el')) }

# Desa la base de l'editor. Torna 'desat', 'fusionat' (la del disc havia
# canviat: s'hi han posat les correccions, _FusionaEdicionsBase) o 'ocupat'
# (s'esta actualitzant ara mateix; no s'ha escrit res). Llanca si no pot
# escriure. $state: @{ Db; Path; Segell } (el segell de quan es va carregar; no
# canvia en desar, perque l'editor segueix tenint la base vella a la memoria).
function Save-BaseEditada($state) {
    $caixa = @{ Res = 'ocupat' }
    [void](Invoke-AmbMutexUnic $Script:BaseMutexNom {
        $disc = $null
        if (Test-Path -LiteralPath $state.Path) { try { $disc = Read-JsonFile $state.Path } catch { $disc = $null } }
        if ($null -ne $disc -and (_SegellBase $disc) -ne [string]$state.Segell) {
            Write-JsonFile $state.Path (_FusionaEdicionsBase $disc $state.Db) 8
            $caixa.Res = 'fusionat'
        } else {
            Write-JsonFile $state.Path $state.Db 8
            $caixa.Res = 'desat'
        }
    })
    return $caixa.Res
}


# L'ID GIA com a NUMERO per ordenar: '9' abans que '10' (com a text, '10' anava
# abans que '9' i la llista sortia 10, 1000, 1019, 103...). Sense GIA o no
# numeric, al final.
function _GiaNumeric($gia) {
    $n = 0L
    if ([long]::TryParse(([string]$gia).Trim(), [ref]$n)) { return $n }
    return [long]::MaxValue
}

# L'ordre de les files de l'editor. PURA.
#   Sense columna triada: per activitat (ID GIA NUMERIC; les que no en tenen, al
#   final i per carpeta) i dins de cada una, per data.
#   Amb una columna triada (clic a la capcalera): AQUELLA COLUMNA MANA, i l'ID
#   GIA i la data nomes desempaten. Abans l'agrupament per activitat era sempre
#   la clau primaria i la columna nomes ordenava DINS de cada activitat, o sigui
#   que clicar 'Estat activitat' no ordenava res que es veies.
# $colExpr: hashtable index de columna -> scriptblock sobre la fila ($_).
function _OrdenaFilesBase($rows, $colExpr, [int]$sortCol, [bool]$asc) {
    $crit = @()
    if ($sortCol -ge 0 -and $null -ne $colExpr -and $colExpr.ContainsKey($sortCol)) {
        $crit += @{ Expression = $colExpr[$sortCol]; Descending = (-not $asc) }
    }
    $crit += @{ Expression = { if ([string]::IsNullOrWhiteSpace($_.Gia)) { 1 } else { 0 } } }
    $crit += @{ Expression = { _GiaNumeric $_.Gia } }
    $crit += @{ Expression = { [string]$_.Carpeta } }
    $crit += @{ Expression = { [string]$_.Data } }
    return @(@($rows) | Sort-Object -Property $crit)
}

# La data d'un informe en format de carrer: 'yyyy-MM-dd' -> 'dd/MM/yyyy'.
# Buit si no hi ha data o no te aquest format (mai peta). Funcio PURA.
function _DataInformeDdMmAaaa($data) {
    $s = [string]$data
    if ($s -match '^(\d{4})-(\d{2})-(\d{2})') {
        return ('{0}/{1}/{2}' -f $Matches[3], $Matches[2], $Matches[1])
    }
    return ''
}

# ID GIA a partir dels noms de les carpetes pare (p.ex. la carpeta de l'activitat
# "2025-1-2563 GIA 361 - ... KRICHI ..."). Retorna '' si cap carpeta no en porta.
# Parteix la ruta manualment per '\' o '/' (aixi funciona igual a Windows i a
# Linux, i es pot provar en headless).
function _GiaFromFolderName($path) {
    $segs = ([string]$path) -split '[\\/]'
    # L'ultim segment es el nom del fitxer; recorrem les carpetes de dins enfora.
    for ($i = $segs.Count - 2; $i -ge 0; $i--) {
        if ($segs[$i] -match 'GIA\s*(\d+)') { return $Matches[1] }
    }
    return ''
}

# Nom de la carpeta (activitat) on viu l'informe: la carpeta pare immediata.
function _CarpetaActivitat($path) {
    $segs = ([string]$path) -split '[\\/]' | Where-Object { $_ -ne '' }
    if ($segs.Count -ge 2) { return $segs[$segs.Count - 2] }
    return ''
}


# Construeix un mapa expedient_normalitzat -> ID GIA a partir de la cache de
# l'Excel d'activitats ($cache.ById[id] = @{ EXP_NUM; TITULAR; ... }).
function Build-ExpedientToGiaMap($cache) {
    $map = @{}
    if ($null -eq $cache -or $null -eq $cache.ById) { return $map }
    foreach ($kv in $cache.ById.GetEnumerator()) {
        $exp = _NormalitzaExpedient $kv.Value.EXP_NUM
        if ($exp -ne '' -and -not $map.ContainsKey($exp)) { $map[$exp] = [string]$kv.Key }
    }
    return $map
}


# ----------------------------------------------------------------------------
# Editor de la base d'informes (finestra amb taula)
# ----------------------------------------------------------------------------
# Estil d'una fila segons si l'informe esta ignorat: gris + tatxat si ho esta.
function _StyleInformeRow($gridRow, [bool]$ignorat, $fontNormal, $fontStrike) {
    if ($ignorat) {
        $gridRow.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gray
        $gridRow.DefaultCellStyle.Font      = $fontStrike
    } else {
        $gridRow.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
        $gridRow.DefaultCellStyle.Font      = $fontNormal
    }
}

# Obre una finestra amb la base d'informes en forma de taula: es pot veure cada
# informe (data, GIA, titular, carpeta, conclusio), OBRIR-lo (boto), marcar-lo
# com a IGNORAT (casella; reversible) i corregir-ne la CONCLUSIO BREU
# (desplegable editable). La columna "Estat activitat" es NOMES LECTURA: es
# deriva de la conclusio breu del darrer informe no ignorat de l'activitat
# (per data) i es recalcula sempre que canvia "ignorar" o "conclusio breu" de
# qualsevol dels seus informes. Tots els canvis es desen al JSON.
# Els botons que hi afegeixen altres moduls (la finestra "Contactes",
# ContactesPantalla.ps1): @{ Nom; Text; Clic }. Aixi l'editor no depen dels
# seus clients (no hi pot haver cicles entre fitxers).
$Script:EditarBaseBotonsExtra = New-Object System.Collections.ArrayList

function Invoke-InformesDbEdit {
    $outPath = Get-InformesDbPath
    if (-not (Test-Path -LiteralPath $outPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Encara no hi ha cap base d'informes.`n`nExecuta primer 'Actualitzar base d'informes'.",
            'Editar base d''informes', 'OK', 'Information') | Out-Null
        return
    }
    $db = $null
    try {
        $db = Read-JsonFile $outPath
    } catch {
        [System.Windows.Forms.MessageBox]::Show("No s'ha pogut llegir la base:`n$($_.Exception.Message)", 'Editar base d''informes', 'OK', 'Error') | Out-Null
        return
    }

    # Aplanar en files, normalitzant cada informe (assegurar ignorat/ruta/motiu/
    # conclusio_breu). Cada fila guarda una REFERENCIA a l'objecte informe del
    # JSON ($inf) i a l'activitat ($act), aixi editar "ignorar" o "conclusio
    # breu" modifica directament la base que despres desem, i podem recalcular
    # l'"estat actual" de l'activitat (sempre DERIVAT: el recomputem en carregar
    # i cada cop que canvia algun dels seus informes, mai es desa "a cegues").
    $allRows = New-Object System.Collections.ArrayList
    if ($null -ne $db.activitats) {
        foreach ($act in $db.activitats) {
            if ($null -eq $act.informes) { continue }
            foreach ($inf in $act.informes) {
                if ($null -eq $inf.PSObject.Properties['ignorat'])        { Add-Member -InputObject $inf -NotePropertyName ignorat -NotePropertyValue $false -Force }
                if ($null -eq $inf.PSObject.Properties['ruta'])           { Add-Member -InputObject $inf -NotePropertyName ruta -NotePropertyValue '' -Force }
                if ($null -eq $inf.PSObject.Properties['motiu'])          { Add-Member -InputObject $inf -NotePropertyName motiu -NotePropertyValue '' -Force }
                if ($null -eq $inf.PSObject.Properties['conclusio_breu']) { Add-Member -InputObject $inf -NotePropertyName conclusio_breu -NotePropertyValue (_ConclusioBreu ([string]$inf.conclusio)) -Force }
                _InferEditatAMa $inf
            }
            if ($null -eq $act.PSObject.Properties['estat_actual']) { Add-Member -InputObject $act -NotePropertyName estat_actual -NotePropertyValue '' -Force }
            $act.estat_actual = _EstatActualActivitat $act
            foreach ($inf in $act.informes) {
                [void]$allRows.Add([pscustomobject]@{
                    Obj           = $inf
                    Act           = $act
                    Data          = [string]$inf.data
                    Gia           = [string]$act.id_gia
                    Titular       = [string]$act.titular
                    Carpeta       = [string]$act.carpeta
                    Conclusio     = [string]$inf.conclusio
                    ConclusioBreu = [string]$inf.conclusio_breu
                    EstatActual   = [string]$act.estat_actual
                    Motiu         = [string]$inf.motiu
                    Ruta          = [string]$inf.ruta
                })
            }
        }
    }

    # Per activitat (ID GIA NUMERIC; les que no en tenen, al final i per carpeta)
    # i per data: vegeu _OrdenaFilesBase.
    $allRows = _OrdenaFilesBase $allRows @{} -1 $true

    # Estat compartit amb els gestors d'esdeveniments (hashtable per referencia).
    # 'Loading' evita que el gestor de la casella reaccioni mentre s'omple la
    # graella (Rows.Add pot disparar CellValueChanged abans d'assignar el Tag).
    # SortColIdx/SortAsc: la columna que l'usuari ha clicat (-1 = cap). Mana
    # ella, i l'ID GIA i la data nomes desempaten (_OrdenaFilesBase).
    # Segell: quan es va escriure la base que s'ha carregat (Save-BaseEditada).
    $state = @{ Dirty = $false; Db = $db; Path = $outPath; Segell = (_SegellBase $db); Loading = $false; SortColIdx = -1; SortAsc = $true }

    $form = _NewForm
    $form.Text = "Editar base d'informes"
    $form.Size = New-Object System.Drawing.Size(1000, 676)
    $form.MinimumSize = New-Object System.Drawing.Size(720, 476)

    # Graella.
    $grid = New-Object System.Windows.Forms.DataGridView
    _StyleListGrid $grid

    $cData = New-Object System.Windows.Forms.DataGridViewTextBoxColumn;  $cData.HeaderText = 'Data';      $cData.ReadOnly = $true; $cData.Width = 90
    $cGia  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn;  $cGia.HeaderText  = 'GIA';       $cGia.ReadOnly  = $true; $cGia.Width  = 60
    $cTit  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn;  $cTit.HeaderText  = 'Titular';   $cTit.ReadOnly  = $true; $cTit.Width  = 180
    $cCar  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn;  $cCar.HeaderText  = 'Carpeta';   $cCar.ReadOnly  = $true; $cCar.Width  = 190
    $cCon  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn;  $cCon.HeaderText  = 'Conclusio'; $cCon.ReadOnly  = $true; $cCon.Width  = 260
    $cBreu = New-Object System.Windows.Forms.DataGridViewComboBoxColumn; $cBreu.HeaderText = 'Conclusio breu'; $cBreu.Width = 150; $cBreu.FlatStyle = 'Flat'
    [void]$cBreu.Items.AddRange($Script:ConclusioBreuOpcions)
    $cEst  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn;  $cEst.HeaderText  = 'Estat activitat'; $cEst.ReadOnly = $true; $cEst.Width = 150
    $cMot  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn;  $cMot.HeaderText  = 'Motiu';     $cMot.ReadOnly  = $true; $cMot.Width  = 110
    $cObr  = New-Object System.Windows.Forms.DataGridViewButtonColumn;   $cObr.HeaderText  = '';          $cObr.Text = 'Obrir'; $cObr.UseColumnTextForButtonValue = $true; $cObr.Width = 64
    $cIgn  = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn; $cIgn.HeaderText  = 'Ignorar';   $cIgn.Width = 60
    [void]$grid.Columns.Add($cData)
    [void]$grid.Columns.Add($cGia)
    [void]$grid.Columns.Add($cTit)
    [void]$grid.Columns.Add($cCar)
    [void]$grid.Columns.Add($cCon)
    [void]$grid.Columns.Add($cBreu)
    [void]$grid.Columns.Add($cEst)
    [void]$grid.Columns.Add($cMot)
    [void]$grid.Columns.Add($cObr)
    [void]$grid.Columns.Add($cIgn)
    $idxConclBreu = 5
    $idxEstat     = 6
    $idxObrir     = 8
    $idxIgnorar   = 9

    # Ordenacio PROGRAMATICA: capturem el clic a la capcalera nosaltres mateixos
    # (mes avall): el DataGridView ordenaria l'ID GIA com a text i no desempataria
    # per activitat. Desactivem l'ordenacio automatica de totes les columnes.
    foreach ($col in $grid.Columns) {
        if ($col.Index -ne $idxObrir) { $col.SortMode = 'Programmatic' }
    }

    $fontNormal = $grid.Font
    $fontStrike = New-Object System.Drawing.Font($grid.Font, [System.Drawing.FontStyle]::Strikeout)
    $fontBold   = New-Object System.Drawing.Font($grid.Font, [System.Drawing.FontStyle]::Bold)
    $colEditat  = [System.Drawing.Color]::FromArgb(192, 0, 0)

    # L'ESTAT EN VERMELL quan alguna cosa de l'activitat s'ha corregit a ma: es
    # el que preval sobre "Actualitzar base", i l'usuari ho vol veure.
    $pintaEstat = {
        param($gr, $row)
        $c = $gr.Cells[$idxEstat]
        if (_ActivitatEditadaAMa $row.Act) {
            $c.Style.ForeColor = $colEditat
            $c.Style.Font = $fontBold
            $c.ToolTipText = "Corregida a m" + [char]0x00E0 + " ('Conclusi" + [char]0x00F3 + " breu' o 'Ignorar'): 'Actualitzar base' no la canvia. Per tornar al que diu l'informe, selecciona la fila i 'Desfer canvi a m" + [char]0x00E0 + "'."
        } else {
            $c.Style.ForeColor = [System.Drawing.Color]::Empty
            $c.Style.Font = $null
            $c.ToolTipText = ''
        }
    }.GetNewClosure()

    # Torna a escriure l'estat de TOTES les files d'una activitat (es el que
    # canvia quan canvia qualsevol dels seus informes).
    $refrescaActivitat = {
        param($act)
        $nouEstat = _EstatActualActivitat $act
        $act.estat_actual = $nouEstat
        foreach ($gr2 in $grid.Rows) {
            $row2 = $gr2.Tag
            if ($null -ne $row2 -and $row2.Act -eq $act) {
                $row2.EstatActual = $nouEstat
                $gr2.Cells[$idxEstat].Value = $nouEstat
                & $pintaEstat $gr2 $row2
            }
        }
        foreach ($row3 in $allRows) {
            if ($row3.Act -eq $act) { $row3.EstatActual = $nouEstat }
        }
    }.GetNewClosure()

    # Expressions d'ordenacio per index de columna (sobre la fila $_). El boto
    # "Obrir" (8) no s'ordena.
    $colExpr = @{
        0 = { [string]$_.Data }
        1 = { _GiaNumeric $_.Gia }
        2 = { [string]$_.Titular }
        3 = { [string]$_.Carpeta }
        4 = { [string]$_.Conclusio }
        5 = { [string]$_.ConclusioBreu }
        6 = { [string]$_.EstatActual }
        7 = { [string]$_.Motiu }
        9 = { [bool]$_.Obj.ignorat }
    }

    # ---- Barra superior: cerca global (fila 1) + filtres per columna (fila 2) ----
    $topPanel = New-Object System.Windows.Forms.Panel
    $topPanel.Dock = 'Top'; $topPanel.Height = 74
    # La cerca torna a omplir la graella. ($fill es defineix mes avall; ja
    # existeix quan l'usuari hi escriu.)
    $txtCerca = _AddSearchBox $topPanel 10 10 300 'Cerca:' { & $fill }

    # Etiqueta a la fila 2 dels filtres.
    $mkLbl = {
        param($text, $x)
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $text; $l.AutoSize = $true
        $l.Location = New-Object System.Drawing.Point($x, 46)
        $topPanel.Controls.Add($l)
    }

    $estatVals = @($allRows | ForEach-Object { $_.EstatActual } | Where-Object { $_ -ne '' } | Sort-Object -Unique)
    $motiuVals = @($allRows | ForEach-Object { $_.Motiu }       | Where-Object { $_ -ne '' } | Sort-Object -Unique)
    # Filtres de SELECCIO MULTIPLE (cap marcat = totes les files passen). L'accio
    # de canvi crida $fill (definit mes avall; ja existeix quan l'usuari hi toca).
    & $mkLbl 'Conclusió breu:' 10
    $mfBreu  = _MakeMultiFilter $topPanel 110 43 130 '(Totes)' $Script:ConclusioBreuOpcions { & $fill }
    & $mkLbl 'Estat:' 250
    $mfEstat = _MakeMultiFilter $topPanel 294 43 120 '(Tots)' $estatVals { & $fill }
    & $mkLbl 'Motiu:' 420
    $mfMotiu = _MakeMultiFilter $topPanel 464 43 110 '(Tots)' $motiuVals { & $fill }
    & $mkLbl 'Ignorats:' 582
    $mfIgn   = _MakeMultiFilter $topPanel 640 43 90  'Tots'    @('Actius', 'Ignorats') { & $fill }

    # (Re)omple la graella: aplica la cerca global + els filtres per columna i,
    # per acabar, ordena mantenint SEMPRE l'agrupament per activitat.
    $fill = {
        $state.Loading = $true
        $grid.Rows.Clear()
        $n = ([string]$txtCerca.Text).Trim().ToLower()
        $selBreu  = & $mfBreu.GetSelected
        $selEstat = & $mfEstat.GetSelected
        $selMotiu = & $mfMotiu.GetSelected
        $selIgn   = & $mfIgn.GetSelected

        $rows = foreach ($row in $allRows) {
            $hay = $row.Data + ' ' + $row.Gia + ' ' + $row.Titular + ' ' + $row.Carpeta + ' ' + $row.Conclusio + ' ' + $row.ConclusioBreu + ' ' + $row.EstatActual + ' ' + $row.Motiu
            if (-not (_TextMatches $hay $n)) { continue }
            # Cada filtre: cap opcio marcada = passa tot; si n'hi ha, el valor de
            # la fila ha de ser entre les marcades (unio / OR dins del filtre).
            if ($selBreu.Count  -gt 0 -and $selBreu  -notcontains $row.ConclusioBreu) { continue }
            if ($selEstat.Count -gt 0 -and $selEstat -notcontains $row.EstatActual)   { continue }
            if ($selMotiu.Count -gt 0 -and $selMotiu -notcontains $row.Motiu)         { continue }
            if ($selIgn.Count -gt 0) {
                $ignLabel = if ([bool]$row.Obj.ignorat) { 'Ignorats' } else { 'Actius' }
                if ($selIgn -notcontains $ignLabel) { continue }
            }
            $row
        }
        $rows = @($rows)

        $rows = _OrdenaFilesBase $rows $colExpr ([int]$state.SortColIdx) ([bool]$state.SortAsc)

        foreach ($row in $rows) {
            $ign = [bool]$row.Obj.ignorat
            $idx = $grid.Rows.Add(@($row.Data, $row.Gia, $row.Titular, $row.Carpeta, $row.Conclusio, $row.ConclusioBreu, $row.EstatActual, $row.Motiu, 'Obrir', $ign))
            $gr = $grid.Rows[$idx]
            $gr.Tag = $row
            $gr.Cells[4].ToolTipText = $row.Conclusio
            _StyleInformeRow $gr $ign $fontNormal $fontStrike
            & $pintaEstat $gr $row
        }
        $state.Loading = $false
    }.GetNewClosure()

    # Clic a la capcalera: tria la columna que mana i alterna asc/desc. La
    # columna "Obrir" (boto) no s'ordena.
    _EnableHeaderSort $grid $state 0 @($idxObrir) { & $fill }

    # Desa la base (retorna $true si va be). Si "Actualitzar base" (l'automatic)
    # esta escrivint ara mateix, no es desa i es diu; els canvis no es perden.
    $doSave = {
        try {
            $res = Save-BaseEditada $state
            if ($res -eq 'ocupat') {
                [System.Windows.Forms.MessageBox]::Show("La base s'est" + [char]0x00E0 + " actualitzant ara mateix (mode autom" + [char]0x00E0 + "tic).`n`nTorna a desar d'aqu" + [char]0x00ED + " a una estona: els canvis no s'han perdut.", 'Editar base d''informes', 'OK', 'Information') | Out-Null
                return $false
            }
            $state.Dirty = $false
            $state.UltimDesat = $res
            if ($res -eq 'fusionat') {
                [System.Windows.Forms.MessageBox]::Show("Canvis desats.`n`nMentre l'editor era obert, la base s'ha actualitzat (mode autom" + [char]0x00E0 + "tic). Els teus canvis s'hi han afegit; tanca i torna a obrir l'editor per veure-hi els informes nous.", 'Editar base d''informes', 'OK', 'Information') | Out-Null
            }
            return $true
        } catch {
            [System.Windows.Forms.MessageBox]::Show("No s'ha pogut desar:`n$($_.Exception.Message)", 'Editar base d''informes', 'OK', 'Error') | Out-Null
            return $false
        }
    }.GetNewClosure()

    # "Desfer canvi a ma": les files seleccionades tornen al que diu l'informe.
    # Definit ABANS del peu de botons: el boto en captura el valor en crear-se.
    $desfesAMa = {
        $files = @($grid.SelectedRows)
        if ($files.Count -eq 0 -and $null -ne $grid.CurrentRow) { $files = @($grid.CurrentRow) }
        $n = 0
        $state.Loading = $true
        try {
            foreach ($gr in $files) {
                $row = $gr.Tag
                if ($null -eq $row -or -not (_DesfesEditatAMa $row.Obj)) { continue }
                $n++
                $row.ConclusioBreu = [string]$row.Obj.conclusio_breu
                $gr.Cells[$idxConclBreu].Value = $row.ConclusioBreu
                $gr.Cells[$idxIgnorar].Value = [bool]$row.Obj.ignorat
                _StyleInformeRow $gr ([bool]$row.Obj.ignorat) $fontNormal $fontStrike
                & $refrescaActivitat $row.Act
            }
        } finally { $state.Loading = $false }
        if ($n -gt 0) { $state.Dirty = $true }
        else {
            [System.Windows.Forms.MessageBox]::Show("La fila seleccionada no t" + [char]0x00E9 + " cap canvi fet a m" + [char]0x00E0 + ".", 'Editar base d''informes', 'OK', 'Information') | Out-Null
        }
    }.GetNewClosure()

    # Barra inferior: botons.
    $botPanel = New-Object System.Windows.Forms.Panel
    $botPanel.Dock = 'Bottom'; $botPanel.Height = 48
    # Exportar a CSV els llistats d'activitats en Estat Requeriment i Precinte /
    # Cessament (usa l'estat en memoria, que ja reflecteix els canvis no desats).
    [void](_AddPeuBotons $form (@(
        @{ Nom = 'Enrere'; Text = (_TxtEnrere); Clic = { $form.Close() }.GetNewClosure() },
        @{ Nom = 'Export'; Text = 'Exportar llistats (CSV)'; Clic = { Export-EstatsActivitats $state.Db }.GetNewClosure() },
        @{ Nom = 'Desfer'; Text = ('Desfer canvi a m' + [char]0x00E0); Clic = { & $desfesAMa }.GetNewClosure() }
        ) + @($Script:EditarBaseBotonsExtra)) @(
        @{ Nom = 'Desar'; Text = 'Desar'; Estil = 'primari'; Clic = {
            if ((& $doSave) -and $state.UltimDesat -eq 'desat') {
                [System.Windows.Forms.MessageBox]::Show('Canvis desats.', 'Editar base d''informes', 'OK', 'Information') | Out-Null
            }
        }.GetNewClosure() }) 8 $botPanel)

    # Obrir l'informe en clicar el boto "Obrir".
    $grid.add_CellContentClick({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        if ($e.ColumnIndex -eq $idxObrir) {
            $row = $s.Rows[$e.RowIndex].Tag
            $ruta = [string]$row.Ruta
            if ([string]::IsNullOrWhiteSpace($ruta) -or -not (Test-Path -LiteralPath $ruta)) {
                [System.Windows.Forms.MessageBox]::Show("No s'ha trobat el fitxer:`n$ruta", 'Obrir informe', 'OK', 'Warning') | Out-Null
                return
            }
            try { Start-Process -FilePath $ruta | Out-Null } catch {
                [System.Windows.Forms.MessageBox]::Show("No s'ha pogut obrir:`n$($_.Exception.Message)", 'Obrir informe', 'OK', 'Error') | Out-Null
            }
        }
    }.GetNewClosure())

    # Forcem que la casella "Ignorar" es confirmi de seguida (no en sortir de la
    # cel-la), perque CellValueChanged salti al moment del clic.
    $grid.add_CurrentCellDirtyStateChanged({
        if ($grid.IsCurrentCellDirty) {
            $grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
    }.GetNewClosure())
    $grid.add_CellValueChanged({
        param($s, $e)
        if ($state.Loading) { return }
        if ($e.RowIndex -lt 0) { return }
        if ($e.ColumnIndex -ne $idxIgnorar -and $e.ColumnIndex -ne $idxConclBreu) { return }
        $gr = $s.Rows[$e.RowIndex]
        $row = $gr.Tag
        if ($null -eq $row) { return }

        # Un canvi a ma: es marca ABANS de canviar el valor, perque s'hi guarda
        # el que deia l'automatic (per poder-ho desfer).
        _MarcaEditatAMa $row.Obj
        if ($e.ColumnIndex -eq $idxIgnorar) {
            $val = [bool]$gr.Cells[$idxIgnorar].Value
            $row.Obj.ignorat = $val
            _StyleInformeRow $gr $val $fontNormal $fontStrike
        } else {
            $val = [string]$gr.Cells[$idxConclBreu].Value
            if ([string]::IsNullOrWhiteSpace($val)) { $val = 'Revisar' }
            $row.Obj.conclusio_breu = $val
            $row.ConclusioBreu = $val
        }
        $state.Dirty = $true

        # L'"ignorar" i la "conclusio breu" de qualsevol informe poden canviar
        # l'estat de l'activitat sencera: el recalculem i el propaguem a totes
        # les files (visibles i filtrades) d'aquesta mateixa activitat.
        & $refrescaActivitat $row.Act
    }.GetNewClosure())


    # Si es tanca amb canvis sense desar, oferim desar-los.
    $form.add_FormClosing({
        param($s, $e)
        if ($state.Dirty) {
            $r = [System.Windows.Forms.MessageBox]::Show('Hi ha canvis sense desar. Vols desar-los?', 'Editar base d''informes', 'YesNoCancel', 'Warning')
            if ($r -eq [System.Windows.Forms.DialogResult]::Yes) {
                if (-not (& $doSave)) { $e.Cancel = $true }
            } elseif ($r -eq [System.Windows.Forms.DialogResult]::Cancel) {
                $e.Cancel = $true
            }
        }
    }.GetNewClosure())

    # Ordre d'afegit: primer el Fill (queda al centre), despres Top i Bottom.
    # La banda de marca s'afegeix l'ultima perque quedi a dalt de tot.
    $form.Controls.Add($grid)
    $form.Controls.Add($topPanel)
    $form.Controls.Add($botPanel)
    [void](_AddBrandHeader $form "Editar base d'informes" 'Base de dades local de deficiencies i conclusions dels informes' 56)

    & $fill
    [void]$form.ShowDialog()
}

# ============================================================================
# Exportar llistats (CSV) — botó "Exportar llistats" de l'editor
# ============================================================================
# Genera un CSV amb les activitats en Estat "Requeriment" i "Precinte /
# Cessament" (una fila per informe). Creua amb l'Excel per obtenir l'adreça
# (si no hi ha Excel, l'adreça queda buida). S'obre a Excel en acabar.
function Export-EstatsActivitats($db) {
    if ($null -eq $db -or $null -eq $db.activitats) {
        [System.Windows.Forms.MessageBox]::Show("La base d'informes és buida.", 'Exportar llistats', 'OK', 'Information') | Out-Null
        return
    }
    $estats = @($Script:EstatRequeriment, $Script:EstatPrecinte)
    # Per com es TRACTA: el favorable pre-llicencia hi surt amb els requeriments.
    $acts = @($db.activitats | Where-Object { $estats -contains (_EstatEquivalent ([string]$_.estat_actual)) })
    if ($acts.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(("No hi ha cap activitat en Estat '" + $Script:EstatRequeriment + "' ni '" + $Script:EstatPrecinte + "'."), 'Exportar llistats', 'OK', 'Information') | Out-Null
        return
    }

    # Adreça per GIA des de l'Excel (opcional; avís suau si no es pot llegir).
    $adrByGia = @{}
    try {
        $xls = Find-LatestActivitatsExcel
        if ($null -ne $xls) {
            $cache = Initialize-ActivitatsCache $xls.File
            if ($null -ne $cache -and $null -ne $cache.ById) {
                foreach ($kv in $cache.ById.GetEnumerator()) { $adrByGia[[string]$kv.Key] = [string]$kv.Value.ADRECA }
            }
        }
    } catch { }

    $rows = New-Object System.Collections.ArrayList
    foreach ($act in $acts) {
        $gia = [string]$act.id_gia
        $adr = if ($gia -and $adrByGia.ContainsKey($gia)) { $adrByGia[$gia] } else { '' }
        foreach ($inf in @($act.informes)) {
            if ($null -eq $inf) { continue }
            [void]$rows.Add([pscustomobject]@{
                'Estat activitat'  = [string]$act.estat_actual
                'GIA'              = $gia
                'Titular'          = [string]$act.titular
                'Adreça'           = $adr
                'Expedient'        = [string]$act.expedient
                'Data informe'     = [string]$inf.data
                'Conclusió breu'   = [string]$inf.conclusio_breu
            })
        }
    }
    $sorted = @($rows | Sort-Object 'Estat activitat', 'GIA', 'Data informe')

    $dir  = _ResolveOutputDir
    $name = 'Estat activitats ' + (Get-Date).ToString('yyyy-MM-dd') + '.csv'
    $path = _GetUniqueOutputPath $dir $name
    try {
        $sorted | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    } catch {
        [System.Windows.Forms.MessageBox]::Show("No s'ha pogut escriure el CSV:`n$($_.Exception.Message)", 'Exportar llistats', 'OK', 'Error') | Out-Null
        return
    }

    $nReq = @($acts | Where-Object { (_EstatEquivalent ([string]$_.estat_actual)) -eq $Script:EstatRequeriment }).Count
    $nPre = @($acts | Where-Object { [string]$_.estat_actual -eq $Script:EstatPrecinte }).Count
    $r = [System.Windows.Forms.MessageBox]::Show(
        "CSV generat:`n$path`n`nRequeriment: $nReq activitats`nPrecinte / Cessament: $nPre activitats`n`nVols obrir-lo ara?",
        'Exportar llistats', 'YesNo', 'Information')
    if ($r -eq [System.Windows.Forms.DialogResult]::Yes) {
        try { Start-Process -FilePath $path | Out-Null } catch {
            [System.Windows.Forms.MessageBox]::Show("No s'ha pogut obrir el CSV:`n$($_.Exception.Message)", 'Exportar llistats', 'OK', 'Warning') | Out-Null
        }
    }
}
