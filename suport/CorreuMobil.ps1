#requires -Version 5.1
<#
  CorreuMobil.ps1 - "Correu del mobil" (des de Configuracio, octubre 2026).

  L'usuari: "Fes que es pugui configurar el tema de correus del mobil des del
  programa i es pugui enviar o generar un mini prompt per canviar la
  configuracio d'enviar correu del mobil".

  El mobil (docs/app.js) llegeix el que envia de docs/dades/email-textos.json,
  que es publica a la web amb Actualitzar.bat (el committeja si ha canviat). El
  que s'hi pot triar des d'aqui:
    - A QUI VA: el titular i/o el representant legal (clau 'dest'; sense, el
      titular, que es el que el mobil ha fet sempre);
    - la CCO de SEMPRE (marcades al mobil) i la A TRIAR (desmarcades): la clau
      'bcc' de sempre, amb el seu 'def'.
  Els textos ja tenen el seu editor (Textos del correu) i les claus d'EmailJS i
  el nom del remitent son a docs/config.js, que Actualitzar NO puja: per a tot
  el que no es d'aquest fitxer, el boto "Copiar el prompt" fa un text per
  enganxar a Claude amb la configuracio d'ara i el canvi que es vol.

  "Publicar ara" desa i llanca Actualitzar (li arriba com a scriptblock des de
  Configuracio: si el cridava directament hi hauria un cicle de dependencies).

  OJU: email-textos.json es al repositori, que es PUBLIC: les adreces de CCO del
  mobil ja hi eren (i la web les ha de poder llegir).

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

# Els textos amb la configuracio nova: la resta de claus, tal com eren. Torna
# un ordered hashtable a punt per a _SaveEmailTextos. PURA.
function _CorreuMobilAplica($textos, [string]$sempre, [string]$tria, $dest) {
    $out = [ordered]@{}
    foreach ($k in @($textos.Keys)) {
        if ($k -eq 'bcc') {
            $out['bcc'] = @(@(_CorreuCcoOpcions $sempre $tria) | ForEach-Object { [pscustomobject]@{ addr = [string]$_.Addr; def = [bool]$_.Default } })
        } elseif ($k -eq 'dest') {
            $out['dest'] = [string[]]@(@($dest) | ForEach-Object { [string]$_ } | Where-Object { $Script:EmailMobilDestValids -contains $_ } | Select-Object -Unique)
        } else { $out[$k] = $textos[$k] }
    }
    if (-not $out.Contains('bcc')) { $out['bcc'] = @(@(_CorreuCcoOpcions $sempre $tria) | ForEach-Object { [pscustomobject]@{ addr = [string]$_.Addr; def = [bool]$_.Default } }) }
    if (-not $out.Contains('dest')) { $out['dest'] = [string[]]@(@($dest) | Where-Object { $Script:EmailMobilDestValids -contains $_ }) }
    return ,$out
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
        "- Nom del remitent (EMAIL_FROM_NAME): " + (Get-ConfigJsValue $cfgJs 'EMAIL_FROM_NAME' '(buit)') + $nl + $nl +
        "On viu: docs/dades/email-textos.json (textos, 'dest' i 'bcc'; s'editen des del programa a Configuraci" + [char]0x00F3 + " -> Correu del m" + [char]0x00F2 + "bil i a Textos del correu), docs/config.js (EmailJS i remitent), docs/app.js (enviarEmail, muntarBcc, destinatariPerDefecte) i docs/email.js." + $nl + $nl +
        "El que vull canviar:" + $nl + $c + $nl + $nl +
        "Fes-ho, passa les proves (suport/tests/run-tests-all.ps1 i suport/tests/navegador/prova-mobil.mjs) i puja-ho a main.")
}

# La finestra. $publica: el que fa "Publicar ara" (Configuracio hi passa
# Invoke-ActualitzarPrograma).
function Show-CorreuMobilConfig([scriptblock]$publica = $null) {
    try { $textos = _LoadEmailTextos } catch {
        [System.Windows.Forms.MessageBox]::Show([string]$_.Exception.Message, 'Correu del m' + [char]0x00F2 + 'bil', 'OK', 'Error') | Out-Null
        return
    }
    $m = _CorreuMobilLlegeix $textos
    $form = _NewForm
    $form.Text = 'Correu del m' + [char]0x00F2 + 'bil'
    $form.ClientSize = New-Object System.Drawing.Size(660, 520)
    $mk = {
        param($tipus, $text, $x, $y, $w, $h)
        $c = New-Object ('System.Windows.Forms.' + $tipus)
        if ($null -ne $text) { $c.Text = $text }
        $c.Location = New-Object System.Drawing.Point($x, $y); $c.Size = New-Object System.Drawing.Size($w, $h)
        [void]$form.Controls.Add($c)
        return $c
    }
    [void](& $mk 'Label' 'A qui va el correu (el destinatari que surt escrit, es pot canviar):' 20 72 620 20)
    $cbT = & $mk 'CheckBox' 'Titular' 36 96 140 22
    $cbR = & $mk 'CheckBox' 'Representant legal' 186 96 200 22
    $cbT.Checked = (@($m.Dest) -contains 'titular'); $cbR.Checked = (@($m.Dest) -contains 'representant')
    [void](& $mk 'Label' 'CCO sempre (surten marcades):' 20 132 200 20)
    $tbS = & $mk 'TextBox' $m.Sempre 222 129 418 24
    [void](& $mk 'Label' 'CCO a triar (desmarcades):' 20 166 200 20)
    $tbT = & $mk 'TextBox' $m.Tria 222 163 418 24
    $nota = & $mk 'Label' ("Es desa a docs\dades\email-textos.json i arriba al m" + [char]0x00F2 + "bil quan es publica (Actualitzar). Compte: aquest fitxer " + [char]0x00E9 + "s al GitHub, que " + [char]0x00E9 + "s p" + [char]0x00FA + "blic. Els textos, a Textos del correu.") 20 196 620 34
    $nota.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
    [void](& $mk 'Label' ("Per a la resta (EmailJS, el remitent, com funciona l'enviament): escriu el canvi i copia el prompt per enganxar-lo a Claude.") 20 240 620 34)
    $tbCanvi = & $mk 'TextBox' '' 20 276 620 120
    $tbCanvi.Multiline = $true; $tbCanvi.ScrollBars = 'Vertical'
    $btnPrompt = & $mk 'Button' 'Copiar el prompt' 20 404 180 30
    _StyleSecondaryButton $btnPrompt
    $st = @{ Textos = $textos; S = $tbS; T = $tbT; Ti = $cbT; Re = $cbR; Canvi = $tbCanvi }
    $llegeixPantalla = {
        $d = @(); if ($st.Ti.Checked) { $d += 'titular' }; if ($st.Re.Checked) { $d += 'representant' }
        return @{ Sempre = (_CorreuNetejaAdreces ([string]$st.S.Text)); Tria = (_CorreuNetejaAdreces ([string]$st.T.Text)); Dest = $d }
    }.GetNewClosure()
    $desa = {
        $p = & $llegeixPantalla
        try { _SaveEmailTextos (_CorreuMobilAplica $st.Textos $p.Sempre $p.Tria $p.Dest); return $true } catch {
            [System.Windows.Forms.MessageBox]::Show("No s'ha pogut desar:`n" + $_.Exception.Message, 'Correu del m' + [char]0x00F2 + 'bil', 'OK', 'Error') | Out-Null
            return $false
        }
    }.GetNewClosure()
    $btnPrompt.add_Click({
        $p = & $llegeixPantalla
        $txt = _CorreuMobilPrompt $p (Read-ConfigJs) ([string]$st.Canvi.Text)
        try { [System.Windows.Forms.Clipboard]::SetText($txt) } catch { }
        [System.Windows.Forms.MessageBox]::Show(("Copiat. Enganxa'l a Claude (Claude Code, al repositori):`n`n" + $txt), 'Copiar el prompt', 'OK', 'Information') | Out-Null
    }.GetNewClosure())
    $dretaBotons = @(@{ Nom = 'Desar'; Text = 'Desar' })
    if ($null -ne $publica) { $dretaBotons += @{ Nom = 'Publicar'; Text = 'Desar i publicar ara'; Estil = 'primari' } }
    $peu = _AddPeuBotons $form @(@{ Nom = 'Tancar'; Text = 'Tancar'; Resultat = 'Cancel'; Esc = $true }) $dretaBotons 460
    $peu.Desar.add_Click({
        if (& $desa) {
            [System.Windows.Forms.MessageBox]::Show("Desat. Arribar" + [char]0x00E0 + " al m" + [char]0x00F2 + "bil la propera vegada que facis Actualitzar.", 'Correu del m' + [char]0x00F2 + 'bil', 'OK', 'Information') | Out-Null
            $form.Close()
        }
    }.GetNewClosure())
    if ($null -ne $publica) {
        $peu.Publicar.add_Click({ if (& $desa) { & $publica } }.GetNewClosure())
    }
    [void](_AddBrandHeader $form ('Correu del m' + [char]0x00F2 + 'bil') ("A qui va i la CCO del correu que envia l'app del m" + [char]0x00F2 + 'bil') 56)
    [void]$form.ShowDialog()
    $form.Dispose()
}
