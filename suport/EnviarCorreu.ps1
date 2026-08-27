#requires -Version 5.1
<#
  EnviarCorreu.ps1 - Eina "Enviar correu" (secció MÒBIL de la pantalla principal).

  Obre la MATEIXA web del mòbil (GitHub Pages) però PRECARREGADA amb l'últim
  informe generat i directament a la pantalla d'enviar. L'enviament el fa el
  navegador amb EmailJS (Public key), EXACTAMENT igual que des del mòbil, així
  que el correu surt idèntic (mateix HTML, mateix remitent) i NO cal cap
  Private key ni infraestructura de correu al PC.

  Passa l'informe al web via el fragment de la URL (#c=<json url-encoded>); el
  fragment NO s'envia mai al servidor (queda al navegador), així que les dades
  del titular no surten enlloc.
#>

# Adreça de la web del mòbil. Es pot sobreescriure a config.ps1 amb $MobilWebUrl.
function _MobilWebUrl {
    if ($MobilWebUrl) { return [string]$MobilWebUrl }
    return 'https://xexifm.github.io/informes-cornella/'
}

# Intenta obtenir l'e-mail del titular des de l'activitats.json de Drive (per
# ID GIA) per precarregar el destinatari. Fail-safe: torna '' si no pot.
function _TitularEmailPerGia($idGia) {
    try {
        if ([string]::IsNullOrWhiteSpace([string]$idGia)) { return '' }
        if (-not (Test-DriveApiConfigured) -or -not $DriveDadesId) { return '' }
        $fid = Find-DriveFileId 'activitats.json' $DriveDadesId
        if (-not $fid) { return '' }
        $data = (Get-DriveFileText $fid) | ConvertFrom-Json
        $act = $data.ById.$idGia
        if ($null -ne $act -and $act.EMAIL) { return [string]$act.EMAIL }
    } catch { }
    return ''
}

function Invoke-EnviarCorreu {
    $rep = Load-LastReport
    if ($null -eq $rep) {
        [System.Windows.Forms.MessageBox]::Show(
            "No hi ha cap informe recent desat. Genera primer un informe i torna a provar.",
            'Enviar correu', 'OK', 'Information') | Out-Null
        return
    }

    # Capçalera: partim de la desada a l'últim informe i, si el trobem, hi
    # afegim l'e-mail del titular (per precarregar el destinatari).
    $header = @{}
    if ($rep.Header) {
        foreach ($p in $rep.Header.PSObject.Properties) { $header[$p.Name] = $p.Value }
    }
    if (-not $header.ContainsKey('EMAIL') -or [string]::IsNullOrWhiteSpace([string]$header['EMAIL'])) {
        $em = _TitularEmailPerGia ([string]$header['ID_GIA'])
        if ($em) { $header['EMAIL'] = $em }
    }

    $payload = [ordered]@{
        v               = 1
        CatalegBaseName = [string]$rep.CatalegBaseName
        Header          = $header
        SelectedKeys    = @($rep.SelectedKeys)
        FieldValues     = $rep.FieldValues
        ConclusionTexts = @($rep.ConclusionTexts)
    }
    $json = $payload | ConvertTo-Json -Depth 8 -Compress
    $enc  = [uri]::EscapeDataString($json)
    $url  = (_MobilWebUrl) + '#c=' + $enc

    try {
        Start-Process $url
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "No s'ha pogut obrir el navegador automaticament. Copia aquesta adreça al navegador:`n`n$url",
            'Enviar correu', 'OK', 'Warning') | Out-Null
    }
}
