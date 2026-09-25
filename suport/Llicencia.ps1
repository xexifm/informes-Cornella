#requires -Version 5.1
<#
.SYNOPSIS
  Informe de LLICENCIA d'activitat (Annex II de la Llei 20/2009 i llicencia
  provisional). Tres informes encadenats sobre el mateix expedient.

.DESCRIPTION
  El full de ruta d'una llicencia son TRES informes:

    1. REQUERIMENT           "Cal requerir l'esmena de les deficiencies..."
    2. FAVORABLE PRE         "S'informa favorablement a l'espera de rebre la
                              citada documentacio..." (+ condicions, opcional)
    3. FAVORABLE POST        "S'informa favorablement l'activitat i es dona per
                              tancat l'expedient." (+ condicions, opcional)

  El primer es opcional, pero es fa gairebe sempre.

  DIFERENCIA AMB UN REQUERIMENT NORMAL: aqui els punts no son deficiencies sino
  DOCUMENTACIO, i surten TANT si es te com si no. Per cada punt s'hi tria:
    - "No es disposa..."  -> NEGRETA (falta)
    - "Es disposa... (Id Firmadoc: ...)" -> sense negreta (ja hi es)
  Al Word que feia servir l'usuari sortien en verd, pero el color era una MARCA
  SEVA per veure que havia de canviar a cada informe; al document generat van
  amb el color de sempre (Format.ps1).

  D'ON SURT EL TEXT: el cos de cada punt es de REQ1, EN VIU. LLIC.json nomes hi
  afegeix el que es propi de Llicencia (els dos comentaris i el "Quan:") i una
  CLAU que apunta a l'item de REQ1 ("Seccio::Titol"). Aixi, canviar un text a
  REQ1 el canvia tambe aqui i no hi ha dues copies per mantenir. Els punts que
  no tenen equivalent a REQ1 porten el text a LLIC (no duen clau).

  ESTRUCTURA DE L'INFORME:
    DOCUMENTACIO NECESSARIA ABANS DE LA RESOLUCIO...
      (punt condicional segons Annex II / llicencia provisional)
      Autoritzacions / Informes preceptius   <- bloc ABANS de LLIC
      Projecte                               <- requeriments normals de REQ1
      Documentacio                           <- tecnic redactor + Id Firmadoc
    DOCUMENTACIO NECESSARIA DESPRES DE LA RESOLUCIO... (amb "Quan:")
    Conclusio de la fase
    ANNEX 1   <- nomes si REQUERIMENT i llicencia provisional

  Les funcions de dades son PURES (es proven en headless, sense Word); nomes
  l'assistent i la composicio del document fan servir WinForms i Word COM.
#>

# Avisa d'on ha quedat l'informe i el deixa obert al Word, com la resta de
# fluxos del programa. Es un sol lloc perque els dos camins de l'assistent
# -l'informe llarg i els dos curts- acabin exactament igual.
function _LlicObreIAvisa($word, [string]$out) {
    [System.Windows.Forms.MessageBox]::Show(
        "Informe generat:`n$out", 'Finalitzat', 'OK', 'Information') | Out-Null
    $word.Visible = $true
    $word.Documents.Open($out) | Out-Null
}

# ----------------------------------------------------------------------------
# PUNT D'ENTRADA (des del menu)
# ----------------------------------------------------------------------------
# $fases: quines fases ofereix aquest assistent. El menu principal en te DUES
# entrades -Llicencia (_LlicFases) i Modificacio NO Substancial / Traspas
# (_MnsFases)- i totes dues passen per aqui: comparteixen capcalera, tramit i
# base de dades, i el que canvia es nomes el document que en surt.
function Invoke-LlicenciaWizard($fases = $null, [string]$titol = '') {
    $llic = Read-LlicCataleg
    if ($null -eq $llic) {
        [System.Windows.Forms.MessageBox]::Show(
            ("No s'ha trobat ESTRUCTURALS\LLIC.json.`n`nAquest fitxer es la base de dades de Llicencia " +
             "(que aporta cada requeriment i que no). Fes 'Actualitzar.bat' per baixar-lo."),
            'Llicencia', 'OK', 'Error') | Out-Null
        return
    }

    $word = $null
    # $st.Fields es el diccionari de camps COMPARTIT de tot l'assistent (el
    # mateix paper que a Invoke-NouWizard): els [CAMP:]/[OPCIO:] s'hi omplen
    # alla on surten i despres la composicio els hi busca.
    # La fase inicial surt de la llista d'AQUEST assistent, no d'un literal: la
    # de MNS/Traspas no te cap 'requeriment'.
    $st = @{ Fase = (_LlicFasePerDefecte $(if ($null -ne $fases) { $fases } else { _LlicTotesLesFases }) 'requeriment')
             Prov = $false; Tecnic = @{}
             # Els actors marcats al pas de les condicions. $null = encara no
             # s'hi ha passat (i llavors se'n proposen segons el bloc ABANS).
             CondActors = $null
             Fields = [ordered]@{}
             # El que s'havia triat a cada pantalla de documentacio, per no
             # perdre-ho quan l'usuari torna ENRERE (era exactament el que
             # passava: tornaves i havies de tornar a marcar-ho tot).
             MemAbans = $null; MemDespres = $null
             # La base de dades nomes es llegeix un cop per sessio (vegeu pas 2).
             DbCarregat = $false }
    $step = 1
    try {
        while ($true) {
            switch ($step) {
                1 {
                    $r = Select-LlicFase $st.Fase $st.Prov $fases $titol
                    if ($r.Nav -ne 'fwd') { return }
                    $st.Fase = [string]$r.Fase
                    $st.Prov = [bool]$r.Prov
                    $step = 2
                }
                2 {
                    $r = Get-HeaderData -preload $st.HeaderPre
                    if ($r.Nav -eq 'back') { $step = 1; break }
                    $st.Header = $r.Data
                    $st.HeaderPre = $r.Data
                    # LA CLASSIFICACIO. La capcalera generica no en te camp (es
                    # NOMES de Llicencia), o sigui que s'omple aqui des de
                    # l'Excel, per ID GIA. Si l'Excel no en te, es demana: sortia
                    # una linia "Classificacio:" BUIDA a l'informe.
                    $st.Header['CLASSIFICACIO'] = _LlicClassificacio $st.Header
                    # LA MEMORIA D'AQUESTA LLICENCIA. Un informe de llicencia
                    # gairebe mai va sol (requeriment -> favorable pre -> post) i
                    # fins ara el segon tornava a demanar-ho TOT, Id Firmadoc i
                    # expedients inclosos. Es carrega UNA sola vegada per sessio:
                    # si l'usuari torna Enrere, el que acaba d'editar mana.
                    if (-not $st.DbCarregat) {
                        $st.DbCarregat = $true
                        $rec = Get-LlicenciaRecord (Load-LlicenciaDb) ([string]$st.Header['ID_GIA'])
                        if ($null -ne $rec) {
                            [void](Restore-LlicenciaState $rec $st)
                            $quan = Get-LlicenciaDataText $rec
                            [System.Windows.Forms.MessageBox]::Show(
                                ("S'han recuperat les dades de l'informe de llic" + [char]0x00E8 + 'ncia del ' + $quan + ".`n`n" +
                                 "Ho trobaras ja marcat i omplert als passos seguents; canvia el que calgui."),
                                'Llicencia', 'OK', 'Information') | Out-Null
                        }
                    }
                    # ELS DOS INFORMES CURTS (Modificacio NO Substancial i
                    # Traspas) no tenen ni blocs de documentacio ni deficiencies
                    # de projecte: nomes cal saber si hi ha observacions.
                    $step = if (_MnsEsFase ([string]$st.Fase)) { 20 } else { 3 }
                }
                20 {
                    if ($null -eq $st.MnsCataleg) { $st.MnsCataleg = Read-MnsCataleg }
                    if ($null -eq $st.MnsCataleg) {
                        [System.Windows.Forms.MessageBox]::Show(
                            ("No trobo el cataleg MNSTRAS.json a ESTRUCTURALS.`n`n" +
                             "Sense el text no es pot fer aquest informe."),
                            'Llicencia', 'OK', 'Error') | Out-Null
                        return
                    }
                    # QUINS PUNTS DE REQ1 S'HI ADJUNTEN. Amb la pantalla de
                    # sempre (Select-Items) i el cataleg SENCER de REQ1: aqui no
                    # hi ha cap seccio que sobri, perque no s'ha demanat res
                    # abans. -permetreBuit: no marcar-ne cap vol dir "sense mes
                    # observacions", que es un cas ben normal.
                    if ($null -eq $st.Req1) {
                        $st.Req1 = Get-ParsedCataleg -path (Join-Path $EstructuralsDir 'REQ1.json')
                        $st.IdxReq1 = _LlicIndexReq1 $st.Req1
                    }
                    $r = Select-Items -sections @($st.Req1.Sections) -preloadSelectedKeys $st.ProjKeys `
                            -fields $st.Fields -preloadValues $st.ProjVals -permetreBuit $true
                    if ($r.Nav -eq 'back') { $step = 2; break }
                    if ($r.Nav -eq 'stay') { break }
                    $st.ProjSel = $r.Data
                    $st.ProjKeys = Get-SelectedKeysFromResult $st.ProjSel
                    $st.ProjVals = Get-FieldValuesForSession $st.Fields
                    if ($null -eq $word) { $word = New-WordApp }
                    $out = Build-MnsDocument $word @{
                        Fase = [string]$st.Fase
                        Header = $st.Header
                        Fields = $st.Fields
                        Punts = @($st.ProjSel)
                        Cataleg = $st.MnsCataleg
                    }
                    _LlicObreIAvisa $word $out
                    return
                }
                3 {
                    # Aqui nomes cal REQ1 (el JSON d'on surt el text). El Word
                    # s'arrenca DIFERIT al pas 9, quan es genera de debo (mateix
                    # motiu que a Invoke-NouWizard: arrencar-lo en fred es lent
                    # i si l'usuari tira enrere no ha de quedar obert per res).
                    if ($null -eq $st.Req1) {
                        $req1Path = Join-Path $EstructuralsDir 'REQ1.json'
                        $st.Req1 = Get-ParsedCataleg -path $req1Path
                        $st.IdxReq1 = _LlicIndexReq1 $st.Req1
                    }
                    # Els punts, resolts amb el text de REQ1. Les claus ORFES
                    # s'avisen: si algu ha reanomenat un requeriment a REQ1, el
                    # punt desapareixeria de l'informe sense dir res.
                    $bAbans  = _LlicPuntsPerBloc $llic $st.IdxReq1 'ABANS' $st.Req1
                    $bDesp   = _LlicPuntsPerBloc $llic $st.IdxReq1 'DESPRES' $st.Req1
                    $bPropis = _LlicPuntsPerBloc $llic $st.IdxReq1 'PROPIS'
                    $orfes = @($bAbans.Orfes) + @($bDesp.Orfes) + @($bPropis.Orfes)
                    if ($orfes.Count -gt 0) {
                        [System.Windows.Forms.MessageBox]::Show(
                            ("Hi ha " + $orfes.Count + " punt(s) de Llicencia que apunten a un requeriment que JA NO " +
                             "existeix a REQ1:`n`n  " + ($orfes -join "`n  ") +
                             "`n`nNo sortiran a l'informe. Arregla-ho des de l'editor de catalegs."),
                            'Llicencia', 'OK', 'Warning') | Out-Null
                    }
                    # Els condicionals entren segons el tipus de llicencia.
                    $cond = @(@($bPropis.Punts) | Where-Object { _LlicCondicioEntra ([string]$_.Condicio) ([bool]$st.Prov) })
                    $st.AbansTots = @($cond) + @($bAbans.Punts)
                    # ELS TEXTOS DE LA FASE al bloc DESPRES (nomes al post: "Es
                    # disposa..." o "No es disposa de la documentacio.").
                    # S'apliquen AQUI perque la pantalla del pas 7 i el document
                    # facin servir EXACTAMENT els mateixos punts.
                    $st.DespresTots = @(_LlicPuntsAmbEstatFase $bDesp.Punts ([string]$st.Fase))
                    $step = 4
                }
                4 {
                    $r = Select-LlicDocumentacio $st.AbansTots ('Documentaci' + [char]0x00F3 + ' ABANS de la resoluci' + [char]0x00F3) `
                            ('Marca la que aplica, si ja es t' + [char]0x00E9 + ' i les seves dades') $true $false $true $st.MemAbans
                    if ($r.Nav -ne 'fwd') { $step = 2; break }
                    $st.Abans = $r.Punts
                    $st.MemAbans = $r.Memoria
                    $step = 5
                }
                5 {
                    # Projecte: els requeriments NORMALS de REQ1, amb la mateixa
                    # pantalla de sempre (no s'hi inventa res).
                    # El PROJECTE es la resta de REQ1: les seccions de
                    # documentacio ja s'han demanat al pas d'ABANS i no s'han de
                    # poder demanar dues vegades.
                    if ($null -eq $st.SeccionsProjecte) {
                        $senseAbans = @(@($st.Req1.Sections) | Where-Object { -not (_LlicEsSeccioAbans ([string]$_.Title)) })
                        # ...i fora tambe tot el que un altre bloc ja expandeix
                        # sencer (no es pot demanar dues vegades). La llista surt
                        # del PROPI cataleg, no d'aqui.
                        $st.SeccionsProjecte = @(_LlicSeccionsSenseSubseccions $senseAbans (_LlicSeccionsExpandides $llic $st.IdxReq1))
                    }
                    # -permetreBuit: pot ser que l'activitat no tingui cap
                    # deficiencia de projecte, i llavors no s'ha d'aturar res.
                    $r = Select-Items -sections $st.SeccionsProjecte -preloadSelectedKeys $st.ProjKeys -fields $st.Fields -preloadValues $st.ProjVals -permetreBuit $true
                    if ($r.Nav -eq 'back') { $step = 4; break }
                    if ($r.Nav -eq 'stay') { break }
                    $st.ProjSel = $r.Data
                    $st.ProjKeys = Get-SelectedKeysFromResult $st.ProjSel
                    $st.ProjVals = Get-FieldValuesForSession $st.Fields
                    $step = 6
                }
                6 {
                    $r = Select-LlicTecnic $st.Tecnic $st.TecnicDocs
                    if ($r.Nav -ne 'fwd') { $step = 5; break }
                    $st.Doc = @{ Text = [string]$r.Text; Items = @($r.Items) }
                    $st.Tecnic = $r.Camps
                    $st.TecnicDocs = $r.Docs
                    $step = 7
                }
                7 {
                    # LA MATEIXA PANTALLA PER A LES TRES FASES. Nomes canvia si
                    # es demana l'estat de cada punt (i les seves dades), que ho
                    # decideix _LlicEstatDespres: al requeriment i al favorable
                    # pre encara no toca dir si es te o no; al post, si.
                    $ef = _LlicEstatDespres ([string]$st.Fase)
                    $titol = 'Documentaci' + [char]0x00F3 + ' DESPR' + [char]0x00C9 + 'S de la resoluci' + [char]0x00F3
                    $sub = "Marca la que entra a l'informe"
                    # Tot marcat de sortida: el Word de l'usuari els portava tots
                    # i ell hi anava esborrant el que no tocava.
                    $r = Select-LlicDocumentacio $st.DespresTots $titol $sub ([bool]$ef.AmbEstat) $true ([bool]$ef.AmbDades) `
                            $st.MemDespres $true ([string]$ef.Estat)
                    if ($r.Nav -ne 'fwd') { $step = 6; break }
                    $st.Despres = $r.Punts
                    $st.MemDespres = $r.Memoria
                    $step = 8
                }
                8 {
                    # LES CONDICIONS, despres de tota la documentacio: la tria
                    # per defecte surt del bloc ABANS (els informes preceptius
                    # que ja es tenen), o sigui que ha d'anar darrere.
                    if (-not (_LlicAdmetCondicions ([string]$st.Fase))) { $step = 9; break }
                    $actors = @(_LlicActorsCondicions $llic)
                    $pre = if ($null -ne $st.CondActors) { @($st.CondActors) } else { @(_LlicActorsPerDefecte $actors $st.Abans) }
                    $r = Select-LlicCondicions $actors $pre $st.CondPdfs
                    if ($r.Nav -ne 'fwd') { $step = 7; break }
                    # LA COPIA LOCAL dels PDF triats (local\base-dades-llicencies\
                    # GIA <id>\a.OGAU.pdf). Si en falla alguna, es diu i es torna
                    # a la pantalla: val mes saber-ho ara que el dia de passar-ho
                    # a PDF.
                    $cp = Copy-LlicAdjunts ([string]$st.Header['ID_GIA']) @($r.Actors) $r.Pdfs
                    if (@($cp.Errors).Count -gt 0) {
                        [System.Windows.Forms.MessageBox]::Show(
                            ("No s'han pogut guardar aquests PDF:`n`n  " + (@($cp.Errors) -join "`n  ")),
                            'Condicions', 'OK', 'Warning') | Out-Null
                        $st.CondActors = @($r.Actors)
                        $st.CondPdfs = $r.Pdfs
                        break
                    }
                    $st.CondActors = @($r.Actors)
                    $st.CondPdfs = $cp.Pdfs
                    $st.CondAdjunts = @($cp.Llista)
                    $step = 9
                }
                9 {
                    if ($null -eq $word) { $word = New-WordApp }
                    # Fora dels favorables, cap actor ni cap adjunt (encara que la
                    # memoria en porti).
                    $actorsModel = if (_LlicAdmetCondicions ([string]$st.Fase)) { @($st.CondActors) } else { @() }
                    $adjuntsModel = if (_LlicAdmetCondicions ([string]$st.Fase)) { @($st.CondAdjunts) } else { @() }
                    # Els camps [CAMP: ...] dels textos triats.
                    $model = @{
                        Fase = [string]$st.Fase
                        EsProvisional = [bool]$st.Prov
                        Header = $st.Header
                        Fields = $st.Fields
                        Abans = @($st.Abans)
                        Projecte = @(_LlicPuntsDeSeleccio $st.ProjSel)
                        Despres = @($st.Despres)
                        Doc = $(if ($null -ne $st.Doc) { $st.Doc } else { @{ Text = ''; Items = @() } })
                        CondicionsActors = $actorsModel
                        Cataleg = $llic
                    }
                    $out = Build-LlicenciaDocument $word $model
                    # Es desa igual que a "Requeriment - Nou" perque el boto
                    # "Recuperar dades ultim informe" del Pas 2 hi arribi. Les
                    # claus desades son les de REQ1 (el bloc Projecte es el
                    # mateix cataleg), o sigui que serveixen als dos fluxos.
                    Save-LastReport ([ordered]@{
                        Version         = 1
                        Timestamp       = (Get-Date).ToString('o')
                        CatalegBaseName = 'LLIC'
                        Header          = $st.Header
                        SelectedKeys    = @($st.ProjKeys)
                        FieldValues     = (Get-FieldValuesForSession $st.Fields)
                        ConclusionTexts = @()
                    })
                    # ...i a la BASE DE DADES DE LLICENCIES, que es el que fa que
                    # el proper informe d'aquesta activitat surti ja omplert.
                    # Un error aqui no pot fer perdre l'informe, que ja esta fet.
                    try {
                        $db = Load-LlicenciaDb
                        $vell = Get-LlicenciaRecord $db ([string]$st.Header['ID_GIA'])
                        $hist = New-Object System.Collections.ArrayList
                        if ($null -ne $vell) { foreach ($x in @($vell.Historial)) { [void]$hist.Add($x) } }
                        # Amb els ADJUNTS d'aquest informe: "Word a PDF" els hi
                        # buscara per saber que ha d'ajuntar-hi (PdfSignar.ps1).
                        [void]$hist.Add((New-LlicenciaHistorial ([string]$st.Fase) ([string]$out) $adjuntsModel))
                        [void](Set-LlicenciaRecord $db (ConvertTo-LlicenciaRecord $st $hist.ToArray()))
                        Save-LlicenciaDb $db
                    } catch {
                        [System.Windows.Forms.MessageBox]::Show(
                            ("L'informe s'ha generat be, pero no s'han pogut desar les dades a la base de " +
                             "llicencies:`n`n" + $_.Exception.Message),
                            'Llicencia', 'OK', 'Warning') | Out-Null
                    }
                    _LlicObreIAvisa $word $out
                    return
                }
                default { return }
            }
        }
    } catch {
        # SENSE AIXO, qualsevol error aqui dins matava el programa EN SILENCI:
        # "es tanca i no passa res, tampoc es genera cap informe". Ara es diu
        # que ha passat i ON, i es torna al menu en lloc de tancar-ho tot.
        $on = ''
        try { $on = "`n`n(" + [System.IO.Path]::GetFileName([string]$_.InvocationInfo.ScriptName) + ', linia ' + [string]$_.InvocationInfo.ScriptLineNumber + ')' } catch { }
        [System.Windows.Forms.MessageBox]::Show(
            ("No s'ha pogut acabar l'informe de Llicencia:`n`n" + $_.Exception.Message + $on),
            'Llicencia', 'OK', 'Error') | Out-Null
    } finally {
        # Si l'informe s'ha generat, el Word s'ha fet visible i es deixa obert
        # per a l'usuari; si es va cancel·lar pel cami, es tanca.
        if ($null -ne $word -and -not $word.Visible) { try { Close-WordApp $word } catch { } }
    }
}
