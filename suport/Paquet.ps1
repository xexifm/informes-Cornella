#requires -Version 5.1
<#
.SYNOPSIS
  Mode "des de paquet": generar un informe SENSE assistent (ve del mobil).

.DESCRIPTION
  Genera l'informe a partir d'un paquet JSON (el mateix model que
  lastreport.json) que ha omplert el formulari web del mobil i que el vigilant
  (mobil\Vigilant.ps1) fa arribar fins aqui. Reutilitza EXACTAMENT el mateix
  motor que l'assistent (Get-ParsedCataleg, Read-Conclusions, Build-Document):
  nomes canvia QUI omple les dades, no com es genera el document.

  Ve de Motor.ps1. Es dot-sourceja des d'alli: mateix ambit, mateix
  comportament.
#>

# ============================================================================
# Mode "des de paquet" (generar des del mobil)
# ----------------------------------------------------------------------------
# Genera un informe a partir d'un paquet JSON (el mateix model que
# lastreport.json) en lloc de l'assistent WinForms. El paquet el prepara el
# formulari web del mobil i el porta fins aqui Vigilant.ps1 (o es pot passar a
# ma). Necessita Word, com el flux normal, pero es 100% no interactiu.
#
# Clau del disseny: REAPROFITA el mateix motor que el flux normal
# (Get-ParsedCataleg, Read-Conclusions, Build-Document). Nomes canvia QUI omple les
# dades: en comptes dels dialegs WinForms, surten del paquet. Les tres funcions
# Build-*FromPaquet son PURES (sense Word/UI) i es proven als tests.
# ============================================================================

# Reconstrueix l'estructura de seleccio que retorna Select-Items (Pas 3) a
# partir d'una llista de claus "Seccio::Item[::Fill]". Es l'invers de
# Get-SelectedKeysFromResult i replica EXACTAMENT la logica de construccio del
# resultat del Pas 3 (Select-Items), pero sense UI.
function Build-SelectionFromKeys($sections, $selectedKeys) {
    $checkStates = @{}
    foreach ($k in $selectedKeys) {
        if ([string]::IsNullOrWhiteSpace($k)) { continue }
        $checkStates[[string]$k] = $true
    }

    # El model el munta _SeccionsTriades (SeleccioItems.ps1), la MATEIXA que fa
    # servir la pantalla del Pas 3. Aixo estava copiat aqui linia a linia.
    return (_SeccionsTriades $sections $checkStates)
}

# Reconstrueix la llista de conclusions triades a partir dels seus titols,
# preservant l'ordre del fitxer (com fa Select-Conclusions, que itera les
# checkboxes en ordre del panell). $selectable son els objectes {Title, Body}
# de Read-Conclusions.
function Build-ConclusionsFromTitles($selectable, $titles) {
    $titleSet = New-Object System.Collections.Generic.HashSet[string]
    if ($titles) { foreach ($t in $titles) { [void]$titleSet.Add([string]$t) } }
    $out = New-Object System.Collections.ArrayList
    foreach ($c in $selectable) {
        if ($titleSet.Contains([string]$c.Title)) { [void]$out.Add($c) }
    }
    return $out.ToArray()
}

# Construeix el diccionari de camps (com els passos 4-5) i hi aplica els valors
# del paquet. $fieldValues pot ser un hashtable o un PSCustomObject (tal com el
# torna ConvertFrom-Json). Els camps sense valor al paquet queden amb el seu
# valor per defecte (per als desplegables, la primera opcio).
function Build-FieldsFromPaquet($selectedSections, $conclusions, $alwaysConcl, $fieldValues) {
    $fields = Get-FieldsFromSelection $selectedSections
    Add-FieldsFromConclusions $fields $conclusions $alwaysConcl
    foreach ($name in @($fields.Keys)) {
        $v = $null
        if ($fieldValues -is [System.Collections.IDictionary]) {
            if ($fieldValues.Contains($name)) { $v = $fieldValues[$name] }
        } elseif ($null -ne $fieldValues -and ($fieldValues.PSObject.Properties.Name -contains $name)) {
            $v = $fieldValues.$name
        }
        if ($null -ne $v) { $fields[$name].Value = [string]$v }
    }
    return $fields
}

# Orquestrador del mode paquet. Llegeix el JSON, resol el cataleg, completa la
# capcalera des de l'Excel si cal (el PC si que hi te acces), reconstrueix les
# seleccions i crida Build-Document. Torna la ruta del .docx generat.
function Invoke-GenerateFromPaquet($paquetPath) {
    if (-not (Test-Path -LiteralPath $paquetPath)) {
        throw "No s'ha trobat el paquet: $paquetPath"
    }
    $raw = Get-Content -LiteralPath $paquetPath -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($raw)) { throw "El paquet esta buit: $paquetPath" }
    $pkg = $raw | ConvertFrom-Json

    $catName = [string]$pkg.CatalegBaseName
    if ([string]::IsNullOrWhiteSpace($catName)) { throw "El paquet no indica 'CatalegBaseName'." }
    $catPath = Join-Path $EstructuralsDir ($catName + '.json')
    if (-not (Test-Path -LiteralPath $catPath)) {
        throw "No s'ha trobat el cataleg '$catName' a $EstructuralsDir."
    }

    # Capcalera: hashtable amb les claus que espera Apply-HeaderReplacements.
    $header = @{}
    foreach ($k in 'ID_GIA','EXP_NUM','ADRECA','ACTIVITAT','TITULAR','ORIGEN_TIPUS','NUM_ANOTACIO','DATA_ANOTACIO','DATA_INSPECCIO') {
        $val = ''
        if ($null -ne $pkg.Header -and ($pkg.Header.PSObject.Properties.Name -contains $k)) {
            $val = [string]$pkg.Header.$k
        }
        $header[$k] = $val
    }
    # El mobil no ofereix la tria d'origen: per defecte, documentacio aportada
    # (mateix comportament que abans, la linia "Objecte:" mostra la doc. aportada).
    if ([string]::IsNullOrWhiteSpace($header['ORIGEN_TIPUS'])) { $header['ORIGEN_TIPUS'] = 'doc' }

    # Si el paquet ve del mobil amb (gairebe) nomes l'ID GIA, completem la
    # capcalera des de l'Excel d'activitats. El PC si que hi te acces; el mobil
    # no (les dades personals no surten mai a la web). Nomes omplim els buits.
    if ([string]::IsNullOrWhiteSpace($header['TITULAR']) -and -not [string]::IsNullOrWhiteSpace($header['ID_GIA'])) {
        try {
            $xls = Find-LatestActivitatsExcel
            if ($null -ne $xls) {
                $cache = Initialize-ActivitatsCache $xls.File
                $act = Get-ActivitatFromCache $cache $header['ID_GIA']
                if ($null -ne $act) {
                    foreach ($k in 'TITULAR','ADRECA','ACTIVITAT','EXP_NUM','NUM_ANOTACIO','DATA_ANOTACIO') {
                        if ([string]::IsNullOrWhiteSpace($header[$k]) -and $act.ContainsKey($k)) {
                            $header[$k] = [string]$act[$k]
                        }
                    }
                }
            }
        } catch {
            # Sense Excel/xarxa seguim amb el que porti el paquet. No es fatal.
            Write-Host "Avis: no s'ha pogut completar la capcalera des de l'Excel ($($_.Exception.Message))."
        }
    }

    $selectedKeys     = @(); if ($pkg.SelectedKeys)    { $selectedKeys     = @($pkg.SelectedKeys) }
    $conclusionTitles = @(); if ($pkg.ConclusionTexts) { $conclusionTitles = @($pkg.ConclusionTexts) }
    $fieldValues      = $pkg.FieldValues

    $word = New-WordApp
    try {
        $parsed   = Get-ParsedCataleg -path $catPath
        $selected = Build-SelectionFromKeys $parsed.Sections $selectedKeys
        # Els informes de cos fix (p.ex. TERMINI) no seleccionen deficiencies:
        # nomes els que tenen seccions exigeixen alguna seleccio valida.
        if (-not $parsed.IsFixedBody -and $selected.Count -eq 0) {
            throw "El paquet no selecciona cap deficiencia valida per al cataleg '$catName'."
        }

        $conclAll    = Read-Conclusions -path $ConclusionsPath -reportType $catName
        $conclusions = Build-ConclusionsFromTitles $conclAll.Selectable $conclusionTitles
        $fields      = Build-FieldsFromPaquet $selected $conclusions $conclAll.Always $fieldValues

        $outPath = Build-Document -word $word -header $header `
                                  -selectedSections $selected `
                                  -fields $fields `
                                  -conclusions $conclusions `
                                  -alwaysConclusions $conclAll.Always `
                                  -catalegName $catName `
                                  -introText $parsed.IntroText `
                                  -conclusionsHeaderText $conclAll.HeaderText `
                                  -isFixedBody $parsed.IsFixedBody `
                                  -fixedBodyLines $parsed.FixedBodyLines

        Save-LastReport ([ordered]@{
            Version         = 1
            Timestamp       = (Get-Date).ToString('o')
            CatalegBaseName = $catName
            Header          = $header
            SelectedKeys    = (Get-SelectedKeysFromResult $selected)
            FieldValues     = (Get-FieldValuesForSession $fields)
            ConclusionTexts = $conclusionTitles
        })

        Write-Host "Informe generat: $outPath"
        return $outPath
    } finally {
        Close-WordApp $word
    }
}
