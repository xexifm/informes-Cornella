#requires -Version 5.1
<#
.SYNOPSIS
  El fil de l'aplicacio: el menu (Main) i l'assistent de passos.

.DESCRIPTION
  Main obre la pantalla inicial (Select-Mode, a Seguiment.ps1), que fusiona la
  tria de MODE i la de CATALEG, i despres crida el que toqui. Invoke-NouWizard
  es la maquina de passos de "Requeriment - Nou": capcalera -> deficiencies ->
  camps -> conclusions -> document, amb Enrere a cada pas.

  Ve de Motor.ps1. Es dot-sourceja des d'alli: mateix ambit, mateix
  comportament.
#>

# ----------------------------------------------------------------------------
# Main flow
# ----------------------------------------------------------------------------
function Main {
    if (-not (Test-Path $HeaderPath)) {
        [System.Windows.Forms.MessageBox]::Show("No s'ha trobat $HeaderPath",'Error','OK','Error') | Out-Null
        exit 1
    }

    # Pas 1: un sol menu (Select-Mode) que tria alhora el MODE i, per al cas
    # "nou", el CATALEG (ja no hi ha un segon pas de tria). Cada flux SEMPRE
    # torna a aquest menu quan acaba o quan es prem Enrere; aixi el programa
    # ROMAN OBERT. L'UNICA manera de sortir del programa es tancar la finestra
    # (X) d'aquest menu inicial (Select-Mode fa exit 0).
    while ($true) {
        $sel = Select-Mode
        switch ($sel.Action) {
            'seguiment'  { Invoke-SeguimentFlow }
            'actextr'    { Invoke-ActExtrFlow }
            'ruta'       { Start-RutaTool }   # llanca el planificador; torna al menu
            'controlsperiodics' { Invoke-ControlsPeriodics }   # llistat d'activitats amb control periodic (Excel)
            'informesdb'     { Invoke-InformesDbScan }   # escaneja informes -> JSON; torna al menu
            'informesdbedit' { Invoke-InformesDbEdit }   # editor de la base d'informes
            'copiarinformes' { Invoke-CopiarInformes }   # copia incremental dels Word a la carpeta de copia
            'comprovarexcel' { Invoke-ComprovarExcel }   # comprova que l'Excel reflecteix els precintes
            'seguimentgia'   { Invoke-SeguimentGia }     # llistats de seguiment del GIA (Excel o PDF)
            'revisarmobil'   { Invoke-RevisarMobil }     # revisa el mobil un sol cop; torna al menu
            'config'         { Invoke-ConfiguracioScreen }   # rutes d'aquest PC + actualitzar; torna al menu
            'editcataleg'    { Show-CatalegEditor -focusDoc ([string]$sel.Doc) }   # editor dels ESTRUCTURALS (xip del document)
            'convertirpdf'   { Invoke-ConvertirPdf }   # converteix una carpeta de Word a PDF (i signa)
            'emailtextos'    { Invoke-EmailTextos }    # edita els textos del correu del mobil
            'nou'        { [void](Invoke-NouWizard -cataleg $sel.Cataleg) }
            default      { return }
        }
        # Segell d'ultima execucio de les EINES, en un sol lloc per a totes: quan
        # l'eina torna, s'apunta la data al registre i el menu la mostra sota la
        # seva rajola. Els tipus d'informe i les pantalles de sistema no en porten
        # ($Script:AccionsSenseSegell, a Seguiment.ps1). Una rajola NOVA hi entra
        # sola, sense tocar cap llista.
        if (-not [string]::IsNullOrWhiteSpace([string]$sel.Action) -and
            [string]$sel.Action -notin $Script:AccionsSenseSegell) {
            _MarcaEinaUsada ([string]$sel.Action)
        }
        # ...i es torna a mostrar el menu (Pas 1).
    }
}

# Wizard de "generar informe nou" (Pas 2..5). El cataleg ja ve triat del Pas 1.
#
# Assistent navegable. Cada pas (dialeg) retorna un objecte amb:
#   Nav  = 'next' | 'back' | 'stay'   ·   Data = el resultat del pas (si 'next')
# El boto "Enrere" retorna 'back' i el wizard torna al pas anterior conservant
# les dades (precarrega). Enrere al Pas 2 => torna al menu inicial (retorna
# 'menu'). Tancar una finestra (X) avorta tot el programa (exit 0).
#
# $st  : dades confirmades de cada pas (es mantenen en memoria al navegar).
# $pre : precarregues per a cada pas (de l'estat o de "Recuperar ultim").
# $st.Fields es un diccionari de camps COMPARTIT: els [OPCIO:]/[CAMP:] s'omplen
# inline alla on apareixen (Pas 3 i Pas 4).
#
# Retorna 'menu' (Enrere al Pas 2) o 'done' (informe generat).
function Invoke-NouWizard {
    param($cataleg)

    $st  = @{ Cataleg=$cataleg; Parsed=$null; Header=$null; Selected=$null; Fields=[ordered]@{}; ConclAll=$null; Conclusions=$null }
    $pre = @{ Header=$null; Keys=$null; Fields=$null; Concl=$null }

    # RENDIMENT: NO obrim Word ni parsejem el cataleg aqui. El Pas 2 (dades de
    # la capcalera) es WinForms pur i no necessita Word, aixi que apareix de
    # seguida en clicar el tipus d'informe. Word (arrencada "en fred", lenta el
    # primer cop) i el parseig del cataleg es fan de forma DIFERIDA quan es
    # necessiten per primer cop (Pas 3), mentre l'usuari omple la capcalera.
    $word = $null
    try {
        $step = 2
        $dir  = 'fwd'
        while ($step -ge 2 -and $step -le 5) {
            switch ($step) {

                2 {
                    $r = Get-HeaderData -preload $pre.Header
                    if ($r.Nav -eq 'back') { return 'menu' }   # enrere Pas 2 = tornar al menu
                    else {
                        $st.Header  = $r.Data
                        $pre.Header = $r.Data
                        if ($r.Recovered) {
                            # "Recuperar dades ultim informe": precarreguem la
                            # resta de passos amb les dades de l'ultim informe.
                            $pre.Keys   = $r.Recovered.SelectedKeys
                            $pre.Fields = $r.Recovered.FieldValues
                            $pre.Concl  = $r.Recovered.ConclusionTexts
                        }
                        $step = 3; $dir = 'fwd'
                    }
                }

                3 {
                    # Primer cop que necessitem Word i el cataleg parsejat:
                    # arrenquem Word (diferit) i parsegem ara (amb cache).
                    if ($null -eq $st.Parsed) {
                        if ($null -eq $word) { $word = New-WordApp }
                        $st.Parsed = Get-ParsedCataleg -path $st.Cataleg.FullName
                    }
                    if ($st.Parsed.IsFixedBody) {
                        # Informe de cos fix (p.ex. TERMINI): no hi ha
                        # deficiencies a triar. Saltem el Pas 3.
                        $st.Selected = @()
                        if ($dir -eq 'back') { $step = 2; $dir = 'back' }
                        else                 { $step = 4; $dir = 'fwd' }
                    } else {
                        # Pas 3: triar deficiencies I omplir-ne les opcions/camps
                        # inline (al mateix panell de detall).
                        $r = Select-Items -sections $st.Parsed.Sections -preloadSelectedKeys $pre.Keys -fields $st.Fields -preloadValues $pre.Fields
                        if     ($r.Nav -eq 'back') { $step = 2; $dir = 'back' }
                        elseif ($r.Nav -eq 'stay') { }   # cap seleccio: es torna a mostrar
                        else {
                            $st.Selected = $r.Data
                            $pre.Keys    = Get-SelectedKeysFromResult $st.Selected
                            $pre.Fields  = Get-FieldValuesForSession $st.Fields
                            $step = 4; $dir = 'fwd'
                        }
                    }
                }

                # Pas 4 = CONCLUSIONS. El cos de cada conclusio es mostra sencer i
                # els seus [CAMP:]/[OPCIO:] s'omplen inline aqui mateix (ja no hi
                # ha un pas separat de camps).
                4 {
                    if ($null -eq $st.ConclAll) {
                        # Les conclusions triables depenen del tipus d'informe
                        # (BaseName del cataleg: REQ1, TERMINI...).
                        if ($null -eq $word) { $word = New-WordApp }
                        $st.ConclAll = Read-Conclusions -path $ConclusionsPath -reportType $st.Cataleg.BaseName
                    }
                    if ($st.ConclAll.Selectable.Count -eq 0) {
                        # No hi ha conclusions triables: saltem el pas.
                        $st.Conclusions = @()
                        if ($dir -eq 'back') { $step = 3; $dir = 'back' } else { $step = 5 }
                    } else {
                        $r = Select-Conclusions -conclusions $st.ConclAll.Selectable -always $st.ConclAll.Always -fields $st.Fields -preloadTitles $pre.Concl -preloadValues $pre.Fields
                        if ($r.Nav -eq 'back') { $step = 3; $dir = 'back' }
                        else {
                            $st.Conclusions = $r.Data
                            # $pre.Concl guarda nomes els TITOLS per a la
                            # precarrega de la propera vegada.
                            $pre.Concl  = @($st.Conclusions | ForEach-Object { $_.Title })
                            $pre.Fields = Get-FieldValuesForSession $st.Fields
                            $step = 5; $dir = 'fwd'
                        }
                    }
                }

                5 {
                    if ($null -eq $word) { $word = New-WordApp }
                    $outPath = Build-Document -word $word -header $st.Header `
                                              -selectedSections $st.Selected `
                                              -fields $st.Fields `
                                              -conclusions $st.Conclusions `
                                              -alwaysConclusions $st.ConclAll.Always `
                                              -catalegName $st.Cataleg.BaseName `
                                              -introText $st.Parsed.IntroText `
                                              -conclusionsHeaderText $st.ConclAll.HeaderText `
                                              -isFixedBody $st.Parsed.IsFixedBody `
                                              -fixedBodyLines $st.Parsed.FixedBodyLines

                    # Desem les dades per poder replicar aquest informe mes endavant.
                    # Per a 'ConclusionTexts' guardem els TITOLS triats (es el que
                    # fa servir Select-Conclusions per precarregar).
                    Save-LastReport ([ordered]@{
                        Version         = 1
                        Timestamp       = (Get-Date).ToString('o')
                        CatalegBaseName = $st.Cataleg.BaseName
                        Header          = $st.Header
                        SelectedKeys    = (Get-SelectedKeysFromResult $st.Selected)
                        FieldValues     = (Get-FieldValuesForSession $st.Fields)
                        ConclusionTexts = @($st.Conclusions | ForEach-Object { $_.Title })
                    })

                    [System.Windows.Forms.MessageBox]::Show(
                        "Informe generat:`n$outPath",
                        'Finalitzat', 'OK', 'Information') | Out-Null

                    # Obrim Word en primer pla per a l'usuari
                    $word.Visible = $true
                    $word.Documents.Open($outPath) | Out-Null
                    $step = 99   # surt del bucle
                }
            }
        }
        return 'done'
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)",'Error','OK','Error') | Out-Null
        throw
    }
    finally {
        # Si mai vam arribar a obrir Word (p.ex. Enrere al Pas 2), no hi ha res
        # a tancar. Si el vam obrir pero no es va fer visible (l'usuari va
        # cancel-lar abans de generar), el tanquem. Si es va fer visible (informe
        # generat i obert), el deixem obert per a l'usuari.
        if ($null -ne $word -and -not $word.Visible) { Close-WordApp $word }
    }
}
