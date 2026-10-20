#requires -Version 5.1
<#
  CorreuMobil.ps1 - El correu de l'APP DEL MOBIL (octubre 2026): una fila mes
  de la finestra Configuracio correu electronic (CorreuEinesPantalla.ps1).
  Abans era una finestra a part ("Correu del mobil..."); l'usuari: "App del
  mobil (web) no se per que te un boto, quan podria ser un apartat mes com la
  resta". Aqui hi ha el que no es pantalla: llegir, comparar i desar.

  L'usuari: "Fes que es pugui configurar el tema de correus del mobil des del
  programa i es pugui enviar o generar un mini prompt per canviar la
  configuracio d'enviar correu del mobil".

  El que s'hi pot triar des d'aqui:
    - A QUI VA: el titular i/o el representant legal. Clau 'dest' de
      docs/dades/email-textos.json, que es publica a la web amb Actualitzar
      (no es cap dada personal); sense la clau, el titular, com sempre;
    - la CCO de SEMPRE (marcades al mobil) i la A TRIAR (desmarcades). Son
      ADRECES de persones i NO van al GitHub, que es public: es desen a
      %LOCALAPPDATA%\InformesCornella\correu-mobil.json (Save-CorreuMobilCco,
      EmailTextos.ps1) i es pugen al Drive PRIVAT (Dades/correu-mobil.json,
      Publish-CorreuMobilCco), d'on les llegeix el mobil.
  Els textos ja tenen el seu editor (Textos del correu) i les claus d'EmailJS i
  el nom del remitent son a docs/config.js, que Actualitzar NO puja: per a tot
  el que no es d'aquest fitxer, el boto "Prompt..." fa un text per enganxar a
  Claude amb la configuracio d'ara i el canvi que es vol.

  Si canvia a qui va, s'ofereix publicar-ho (Actualitzar; li arriba com a
  scriptblock des de Configuracio: si el cridava directament hi hauria un
  cicle de dependencies).

  NOMES DEFINEIX FUNCIONS.
#>

# La configuracio del mobil que hi ha als textos. Torna @{ Sempre; Tria; Dest }.
# PURA.
function _CorreuMobilLlegeix($textos) {
    $b = @($textos['bcc'])
    $sempre = (@($b | Where-Object { $null -ne $_ -and [bool]$_.Default } | ForEach-Object { [string]$_.Addr })) -join '; '
    $tria = (@($b | Where-Object { $null -ne $_ -and -not [bool]$_.Default } | ForEach-Object { [string]$_.Addr })) -join '; '
    $dest = if ($textos.Contains('dest')) { @($textos['dest']) } else { @('titular') }
    return @{ Sempre = (_CorreuNetejaAdreces $sempre); Tria = ((@(_CorreuCcoTriaLlista $sempre $tria)) -join '; '); Dest = @($dest | Where-Object { $_ }) }
}

# Els textos (email-textos.json) amb a qui va: la resta de claus, tal com eren,
# i SENSE la CCO (es local, no va mai al GitHub). Torna un ordered hashtable a
# punt per a _SaveEmailTextos. PURA.
function _CorreuMobilAplica($textos, $dest) {
    $out = [ordered]@{}
    foreach ($k in @($textos.Keys)) {
        if ($k -eq 'bcc') { continue }
        $out[$k] = $textos[$k]
    }
    $out['dest'] = [string[]]@(@($dest) | ForEach-Object { [string]$_ } | Where-Object { $Script:EmailMobilDestValids -contains $_ } | Select-Object -Unique)
    return ,$out
}

# Puja la CCO del mobil al Drive privat (Dades/correu-mobil.json), d'on la
# llegeix l'app. Torna @{ Ok; Text }; mai peta.
function Publish-CorreuMobilCco {
    try {
        $json = ((_CorreuMobilCcoAJson @(Read-CorreuMobilCco)) | ConvertTo-Json -Depth 4)
        $n = _CorreuMobilCcoCompta $json
        Save-ADadesDrive 'correu-mobil.json' $json 'application/json; charset=UTF-8'
        # ES COMPROVA el que ha quedat al Drive, com ho llegira el mobil.
        if (Test-DriveApiConfigured) {
            $id = Find-DriveFileId 'correu-mobil.json' $DriveDadesId
            $nDrive = if ($id) { _CorreuMobilCcoCompta (Get-DriveFileText $id) } else { -1 }
            if ($nDrive -ne $n) { return @{ Ok = $false; Text = ("La CCO s'ha pujat al Drive, pero en tornar-la a llegir hi ha " + $nDrive + " adreces en lloc de " + $n + ".") } }
            return @{ Ok = $true; Text = ("La CCO (" + $n + " adreces) s'ha pujat al Drive i s'ha comprovat: el m" + [char]0x00F2 + "bil ja la t" + [char]0x00E9 + ".") }
        }
        return @{ Ok = $true; Text = ("La CCO (" + $n + " adreces) s'ha desat a la carpeta del Drive d'aquest PC (" + $DriveDadesDir + "): arribara al m" + [char]0x00F2 + "bil quan el Drive de l'escriptori la sincronitzi.") }
    } catch {
        return @{ Ok = $false; Text = ("No s'ha pogut pujar la CCO al Drive (" + $_.Exception.Message + "). Es tornara a provar amb Actualitzar.") }
    }
}

# Quantes adreces porta un correu-mobil.json (el text), com el llegeix el
# mobil (docs/cco.js): -1 si no es JSON. PURA.
function _CorreuMobilCcoCompta([string]$json) {
    try { $o = $json | ConvertFrom-Json } catch { return -1 }
    if ($null -eq $o) { return -1 }
    return @(@(_PropInf $o 'bcc') | Where-Object { $null -ne $_ -and -not [string]::IsNullOrWhiteSpace([string](_PropInf $_ 'addr')) }).Count
}

# EL MINI PROMPT: la configuracio d'ara i el canvi que es vol, per enganxar a
# Claude (Claude Code, al repositori). $cfgJs: el que torna Read-ConfigJs. PURA.
function _CorreuMobilPrompt($m, $cfgJs, [string]$canvi) {
    $noms = @{ titular = 'el titular'; representant = 'el representant legal' }
    $dest = (@(@($m.Dest) | ForEach-Object { [string]$noms[[string]$_] } | Where-Object { $_ })) -join ' i '
    if (-not $dest) { $dest = "ning" + [char]0x00FA + " (s'escriu a m" + [char]0x00E0 + ")" }
    $ejs = @('EMAILJS_PUBLIC_KEY', 'EMAILJS_SERVICE_ID', 'EMAILJS_TEMPLATE_ID') | Where-Object { [string]::IsNullOrWhiteSpace((Get-ConfigJsValue $cfgJs $_)) }
    $ejsTxt = if (@($ejs).Count -eq 0) { 'configurades' } else { 'en falten: ' + (@($ejs) -join ', ') }
    $c = ([string]$canvi).Trim()
    if (-not $c) { $c = '(escriu aqu' + [char]0x00ED + ' el canvi)' }
    $nl = "`r`n"
    return ("Al repositori xexifm/informes-cornella vull canviar com envia el correu l'app del m" + [char]0x00F2 + "bil (docs/)." + $nl + $nl +
        "Com est" + [char]0x00E0 + " ara:" + $nl +
        "- A qui va (per defecte): " + $dest + $nl +
        "- CCO sempre (marcades): " + $(if ($m.Sempre) { $m.Sempre } else { 'cap' }) + $nl +
        "- CCO a triar (desmarcades): " + $(if ($m.Tria) { $m.Tria } else { 'cap' }) + $nl +
        "- Envia per EmailJS; claus a docs/config.js: " + $ejsTxt + $nl +
        "- Nom del remitent (EMAIL_FROM_NAME): " + (Get-ConfigJsValue $cfgJs 'EMAIL_FROM_NAME' '(buit)') + $nl +
        "- Respondre a (Reply-To): el camp 'Reply To' de la plantilla d'EmailJS (no al codi: el repositori es public)" + $nl + $nl +
        "On viu: docs/dades/email-textos.json (textos i 'dest'), la CCO a correu-mobil.json (LOCALAPPDATA del PC i Dades del Drive; MAI al GitHub), docs/config.js (EmailJS i remitent), docs/app.js (enviarEmail, destinatariPerDefecte), docs/cco.js (la CCO: sempre marcada, a triar desmarcada), docs/drive.js i docs/email.js." + $nl + $nl +
        "El que vull canviar:" + $nl + $c + $nl + $nl +
        "Fes-ho, passa les proves (suport/tests/run-tests-all.ps1 i suport/tests/navegador/prova-mobil.mjs) i puja-ho a main.")
}

# Si la pantalla diu una altra cosa que el que hi havia (a qui va, o la CCO
# com a conjunt). Torna @{ Dest; Cco } (bool cada un). PURA.
function _CorreuMobilCanvis($abans, $ara) {
    $conj = { param($x) ((@(_CorreuLlistaAdreces ([string]$x)) | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object) -join ';') }
    $dA = (@(@($abans.Dest) | ForEach-Object { [string]$_ } | Sort-Object) -join ';')
    $dB = (@(@($ara.Dest) | ForEach-Object { [string]$_ } | Sort-Object) -join ';')
    $cco = ((& $conj $abans.Sempre) -ne (& $conj $ara.Sempre)) -or ((& $conj $abans.Tria) -ne (& $conj $ara.Tria))
    return @{ Dest = ($dA -ne $dB); Cco = [bool]$cco }
}

# Desa el que diu la pantalla ($ara: @{ Sempre; Tria; Dest }), nomes el que ha
# canviat respecte de $abans. A qui va, a email-textos.json (arriba al mobil
# amb Actualitzar); la CCO, a aquest PC i al Drive. Torna @{ Dest (si ha
# canviat a qui va); Text (que dir, '' si res) ; Ok }. Pot petar (el desa de
# la finestra ho ensenya).
function Save-CorreuMobilConfig($abans, $ara) {
    $cv = _CorreuMobilCanvis $abans $ara
    $r = @{ Dest = [bool]$cv.Dest; Text = ''; Ok = $true }
    if ($cv.Cco) { Save-CorreuMobilCco @(_CorreuCcoOpcions $ara.Sempre $ara.Tria) }
    # SEMPRE es torna a pujar, encara que no hagi canviat (octubre 2026: el
    # mobil deia "cap adreca" amb la CCO ben posada al PC; nomes es pujava si
    # canviava o a Actualitzar, i si aquell pas fallava no se'n sabia res).
    $d = Publish-CorreuMobilCco
    $r.Text = [string]$d.Text; $r.Ok = [bool]$d.Ok
    if ($cv.Dest) { _SaveEmailTextos (_CorreuMobilAplica (_LoadEmailTextos) $ara.Dest) }
    return $r
}

# EL PROMPT per a la resta (EmailJS, el remitent, com s'envia): una finestreta
# on s'escriu el canvi i es copia. $m: el que diu ara la pantalla.
function Show-CorreuMobilPrompt($m) {
    $form = _NewForm
    $form.Text = 'Prompt per a Claude'
    $form.ClientSize = New-Object System.Drawing.Size(620, 300)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = ("Per a la resta del correu de l'app del m" + [char]0x00F2 + "bil (EmailJS, el remitent, com s'envia): escriu el canvi i copia el prompt per enganxar-lo a Claude.")
    $l.Location = New-Object System.Drawing.Point(20, 70); $l.Size = New-Object System.Drawing.Size(580, 34)
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Multiline = $true; $tb.ScrollBars = 'Vertical'
    $tb.Location = New-Object System.Drawing.Point(20, 108); $tb.Size = New-Object System.Drawing.Size(580, 120)
    foreach ($c in @($l, $tb)) { [void]$form.Controls.Add($c) }
    $peu = _AddPeuBotons $form @(@{ Nom = 'Tancar'; Text = 'Tancar'; Resultat = 'Cancel'; Esc = $true }) @(@{ Nom = 'Copiar'; Text = 'Copiar el prompt'; Estil = 'primari' }) 240
    $peu.Copiar.add_Click({
        $txt = _CorreuMobilPrompt $m (Read-ConfigJs) ([string]$tb.Text)
        try { [System.Windows.Forms.Clipboard]::SetText($txt) } catch { }
        [System.Windows.Forms.MessageBox]::Show(("Copiat. Enganxa'l a Claude (Claude Code, al repositori):`n`n" + $txt), 'Copiar el prompt', 'OK', 'Information') | Out-Null
        $form.Close()
    }.GetNewClosure())
    [void](_AddBrandHeader $form 'Prompt per a Claude' ("Canviar com envia el correu l'app del m" + [char]0x00F2 + 'bil') 56)
    [void]$form.ShowDialog()
    $form.Dispose()
}
