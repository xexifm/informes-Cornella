#requires -Version 5.1
<#
.SYNOPSIS
  Mode "ACT_EXTR": seguiment d'activitats extraordinaries de caracter
  esporadic (Decret 112/2010). Comprova la documentacio segons el Decret,
  decideix QUE cal i, segons si s'ha entregat o no, genera un REQUERIMENT
  (del que falta) o un INFORME FAVORABLE.

.DESCRIPTION
  Es un mode alternatiu del programa (es tria a la pantalla inicial, com el
  "seguiment"). A diferencia del flux REQ1 (on l'usuari tria deficiencies d'un
  cataleg), aqui el PROGRAMA sap que cal segons el Decret 112/2010 a partir de
  les respostes de classificacio i de l'aforament; l'usuari nomes marca, per
  cada punt aplicable, si la documentacio esta lliurada o pendent.

  Substitueix l'Excel "DATA_Act_Extr_NOM" (basat en el Decret 112/2010):
  la logica d'aplicabilitat i el calcul dels valors (vigilants, controladors,
  lavabos/cabines, p0lissa RC, organ del Pla d'Autoproteccio...) viuen aqui;
  els TEXTOS (requeriment i informe favorable) viuen a plantilles editables
  ESTRUCTURALS\ACT_EXTR_REQ.docx i ACT_EXTR_FAV.docx.

  Flux (Invoke-ActExtrFlow):
    1. Llistat local d'activitats extraordinaries amb el seu estat (pendent /
       tancat). Es pot crear-ne una de nova o triar-ne una d'existent.
    2. (Nova) Dades de capcalera ACT_EXTR (capcalera propia, diferent de REQ1):
       ID GIA, Exp., Adreca, Activitat, Titular, Dates i AFORAMENT.
    3. Comprovacio de la documentacio: respostes de classificacio del Decret +
       aforament. El programa mostra, per cada punt, si APLICA i PER QUE, i
       l'usuari marca si esta lliurat. Des d'aqui es pot generar el requeriment
       (del que falta) o l'informe favorable (si tot esta lliurat).

  Persistencia: un registre local (JSON, ignorat per git) guarda les dades de
  cada activitat (titular, adreca, aforament, respostes, estat de cada punt i
  historial de documents generats). Aixi cada vegada que es fa un requeriment o
  l'informe favorable es te a ma la "memoria de l'estat de l'activitat".

.NOTES
  EL MODUL ES EN QUATRE FITXERS, un per cosa (revisio d'arquitectura, setembre
  2026; abans era un sol fitxer de 1.390 linies):
    ActExtrDades.ps1     logica del Decret, punts, plantilla i registre (PURES)
    ActExtrBlocs.ps1     composicio del document en blocs (PURA) + el .docx
    ActExtrPantalles.ps1 les finestres (WinForms)
    ActExtr.ps1          aquesta descripcio i l'orquestrador (Invoke-ActExtrFlow)

  Es carrega via dot-source des de GenerarInforme.ps1 (tambe en mode headless
  de proves). Les funcions PURES (logica del Decret, model, parseig de
  plantilla, inclusio de blocs) son testejables a Linux sense Word.

  Reutilitza de GenerarInforme.ps1 / Format.ps1 / MotorInforme.ps1:
  _NormalitzaText, _SplitTextAndUrls, Write-InformeDocx (obrir la plantilla,
  escriure el cos i desar), les funcions Format-* i $ReportFormatConfig.

  CONVENCIO ASCII: per evitar problemes d'encoding de PowerShell 5.1, el codi
  no porta accents. Tot el text accentuat que va als documents viu a les
  plantilles .docx (no al codi).
#>

# ----------------------------------------------------------------------------
# Orquestrador del mode ACT_EXTR
# ----------------------------------------------------------------------------
function Invoke-ActExtrFlow {
    $registry = Load-ActExtrRegistry
    $word = $null
    try {
        while ($true) {
            $sel = Show-ActExtrList $registry
            if ($sel.Action -eq 'exit') { break }

            $header = $null; $answers = $null; $delivered = $null
            if ($sel.Action -eq 'new') {
                $h = Get-ActExtrHeader
                if ($h.Nav -ne 'next') { continue }
                $header = $h.Data
            } else {
                $act = Get-ActExtrActivity $registry $sel.Id
                if ($null -eq $act) { continue }
                $header    = $act.Header
                $answers   = $act.Decret
                $delivered = $act.Punts
                # Permet revisar/editar la capcalera abans del Pas 3.
                $h = Get-ActExtrHeader -preload $header -lockId $true
                if ($h.Nav -eq 'next') { $header = $h.Data }
            }

            # Pas 3 (es pot repetir fins que es tanqui amb Enrere)
            $stay = $true
            while ($stay) {
                $r = Edit-ActExtrDocumentacio -header $header -answers $answers -delivered $delivered
                if ($r.Action -eq 'back') { $stay = $false; break }
                $answers   = $r.Answers
                $delivered = $r.Delivered

                $decret   = Build-ActExtrDecret $answers
                $computed = Get-ActExtrComputed $decret
                $estat    = Get-ActExtrActivityEstat $decret $computed $delivered

                # Desem/actualitzem l'activitat al registre. Fem servir un
                # ArrayList per a l'historial: amb '+=' sobre el resultat d'un
                # 'if' de @(...) d'un sol element, PowerShell el desempaqueta a
                # escalar i el '+=' peta ("op_Addition" sobre PSObject).
                $existing = Get-ActExtrActivity $registry ([string]$header.ID_GIA)
                $historial = New-Object System.Collections.ArrayList
                if ($existing -and $existing.Historial) {
                    foreach ($h in @($existing.Historial)) { if ($null -ne $h) { [void]$historial.Add($h) } }
                }
                $creatAt = if ($existing -and $existing.CreatAt) { [string]$existing.CreatAt } else { (Get-Date).ToString('o') }

                $outPath = $null
                if ($r.Action -in @('req','fav')) {
                    if ($null -eq $word) { $word = New-WordApp }
                    $outPath = Build-ActExtrDocument $word $header $decret $delivered $r.Action
                    $tipus = if ($r.Action -eq 'fav') { 'favorable' } else { 'requeriment' }
                    [void]$historial.Add([pscustomobject]@{
                        Data    = (Get-Date).ToString('o')
                        Tipus   = $tipus
                        Fitxer  = $outPath
                    })
                }

                $activity = [pscustomobject]@{
                    IdGia      = [string]$header.ID_GIA
                    Header     = $header
                    Decret     = $answers
                    Punts      = $delivered
                    Estat      = $estat
                    CreatAt    = $creatAt
                    ModificatAt= (Get-Date).ToString('o')
                    Historial  = @($historial.ToArray())
                }
                $registry = Set-ActExtrActivity $registry $activity
                Save-ActExtrRegistry $registry

                if ($r.Action -in @('req','fav')) {
                    [System.Windows.Forms.MessageBox]::Show("Document generat:`n$outPath",'Finalitzat','OK','Information') | Out-Null
                    $word.Visible = $true
                    $word.Documents.Open($outPath) | Out-Null
                    # En obrir Word per a l'usuari, no el tanquem; sortim al llistat.
                    $word = $null
                    $stay = $false
                } else {
                    # 'save': nomes desar; tornem al Pas 3 amb les dades desades.
                    [System.Windows.Forms.MessageBox]::Show('Dades desades al registre.','Desat','OK','Information') | Out-Null
                }
            }
        }
    } finally {
        if ($null -ne $word) { Close-WordApp $word }
    }
}
