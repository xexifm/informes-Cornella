#requires -Version 5.1
<#
.SYNOPSIS
  Que una finestra HI CAPIGA sempre: scroll vertical i ajust a la pantalla.

.DESCRIPTION
  Problema real: en una pantalla mes baixa (el PC de casa, o el Windows amb
  escalat al 125%), diverses pantalles del programa son MES ALTES que l'area de
  treball. Llavors la part de baix -que es on van els botons Enrere/Seguent- no
  es pot veure ni arribar-hi: la finestra no es pot encongir perque te
  MinimumSize, i el que sobresurt queda fora de la pantalla.

  Aixo es arregla en DOS temps, i calen tots dos:

    1. AutoScroll a TOTES les finestres: si s'encongeixen i algun control queda
       per sota, surt la barra VERTICAL i s'hi arriba. Es deixa que el WinForms
       calculi sol la zona a recorrer, aixi una graella Dock='Fill' segueix
       encongint-se com sempre en lloc d'estrenar una barra que no calia.

    2. Nomes quan la finestra NO hi cap: encongir-la fins a l'area i, ABANS,
       BAIXAR-NE EL MinimumSize -si no, el Windows es nega a encongir-la-. En
       aquest cas s'hi fixa AutoScrollMinSize = l'alcada de DISSENY, que es
       l'unica manera de garantir que s'arriba a tot el que hi havia, tambe al
       que estigui ancorat a baix (que si no puja i es comprimeix). L'amplada es
       deixa a 0: els controls ancorats a la dreta ja s'estrenyen sols i posar-hi
       l'amplada de disseny trauria una barra HORITZONTAL que no cal.

  NO ES CARREGA RES EN AQUEST FITXER: nomes defineix funcions. Es a posta,
  perque el fan servir DOS PROCESSOS -el programa (via UiComuns.ps1) i el
  planificador de rutes (rutes/Ruta.ps1, que no carrega UiComuns perque aquell
  si que te efectes en carregar-se: AppUserModelID, icona...)-. Un modul amb
  efectes no es pot compartir entre processos sense arrossegar-los.
#>

# ----------------------------------------------------------------------------
# La DECISIO es pura: rebre mides i tornar què s'ha d'aplicar. Es prova a Linux.
# ----------------------------------------------------------------------------
# Entrada: la mida que vol la finestra, el seu MinimumSize i l'area de treball
# de la pantalla (la de debo, sense la barra de tasques).
# Sortida: @{ W; H; MinW; MinH; X; Y; Cal } -Cal = $false si ja hi cabia i no
# s'ha de tocar res-.
#
# Regles:
#   - el MinimumSize no pot passar de l'area: si hi passa, el Windows no deixa
#     encongir la finestra i tot plegat no serveix;
#   - la mida es retalla a l'area;
#   - i la posicio es corre perque la finestra quedi SENCERA a dins (una finestra
#     centrada que sobresurt per baix tambe sobresurt per dalt, i llavors ni la
#     barra de titol es pot agafar).
function _MidaFinestraDinsPantalla([int]$w, [int]$h, [int]$minW, [int]$minH,
                                   [int]$x, [int]$y,
                                   [int]$areaX, [int]$areaY, [int]$areaW, [int]$areaH) {
    $nMinW = if ($minW -gt $areaW) { $areaW } else { $minW }
    $nMinH = if ($minH -gt $areaH) { $areaH } else { $minH }
    $nW = if ($w -gt $areaW) { $areaW } else { $w }
    $nH = if ($h -gt $areaH) { $areaH } else { $h }
    if ($nW -lt $nMinW) { $nW = $nMinW }
    if ($nH -lt $nMinH) { $nH = $nMinH }

    $nX = $x
    $nY = $y
    if (($nX + $nW) -gt ($areaX + $areaW)) { $nX = $areaX + $areaW - $nW }
    if (($nY + $nH) -gt ($areaY + $areaH)) { $nY = $areaY + $areaH - $nH }
    if ($nX -lt $areaX) { $nX = $areaX }
    if ($nY -lt $areaY) { $nY = $areaY }

    return @{
        W = $nW; H = $nH; MinW = $nMinW; MinH = $nMinH; X = $nX; Y = $nY
        Cal = (($nW -ne $w) -or ($nH -ne $h) -or ($nMinW -ne $minW) -or
               ($nMinH -ne $minH) -or ($nX -ne $x) -or ($nY -ne $y))
    }
}

# ----------------------------------------------------------------------------
# L'APLICACIO (WinForms). Es crida des del Shown, quan la disposicio ja es
# definitiva: abans, el ClientSize encara pot canviar.
# ----------------------------------------------------------------------------
# L'alcada de DISSENY es la que te la finestra en obrir-se, ABANS de retallar-la:
# es la que ha de poder recorrer la barra de desplacament.
function _AjustaFinestraAPantalla($f) {
    if ($null -eq $f) { return }
    try {
        # L'alcada de DISSENY s'ha de llegir ABANS de tocar res.
        $altDisseny = [int]$f.ClientSize.Height

        # Scroll SEMPRE: si l'usuari encongeix la finestra i algun control queda
        # per sota, surt la barra i s'hi pot arribar. Amb AutoScrollMinSize a
        # zero, el WinForms calcula la zona a recorrer dels controls mateixos, o
        # sigui que una graella Dock='Fill' segueix ENCONGINT-SE com fins ara i
        # no apareix cap barra que abans no hi era.
        $f.AutoScroll = $true

        $area = ([System.Windows.Forms.Screen]::FromControl($f)).WorkingArea
        $r = _MidaFinestraDinsPantalla ([int]$f.Width) ([int]$f.Height) `
                                       ([int]$f.MinimumSize.Width) ([int]$f.MinimumSize.Height) `
                                       ([int]$f.Left) ([int]$f.Top) `
                                       ([int]$area.X) ([int]$area.Y) ([int]$area.Width) ([int]$area.Height)
        if (-not $r.Cal) { return }

        # Aqui si: la finestra NO hi cabia i la retallem. Llavors s'hi fixa
        # l'alcada de disseny com a zona recorrible, que es l'unica manera de
        # garantir que s'arriba a TOT el que hi havia -tambe al que estigui
        # ancorat a baix, que si no simplement pujaria i es comprimiria-.
        $f.AutoScrollMinSize = New-Object System.Drawing.Size(0, $altDisseny)
        # El MinimumSize PRIMER: si no, el Windows no deixa encongir la finestra.
        $f.MinimumSize = New-Object System.Drawing.Size([int]$r.MinW, [int]$r.MinH)
        $f.Size = New-Object System.Drawing.Size([int]$r.W, [int]$r.H)
        $f.Location = New-Object System.Drawing.Point([int]$r.X, [int]$r.Y)
    } catch { }
}
