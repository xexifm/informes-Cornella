# Correu, recordatoris i configuracio
#
# Es DOT-SOURCE des de run-tests.ps1: mateix ambit, mateixes variables i el
# mateix comptador d'asserts. No s'executa sol.

Write-Host "`n--- EmailTextos.ps1: el JSON es l'unic origen dels textos ---"
# Abans aqui es provava _DefaultEmailTextos, una copia dels textos escrita al
# codi. N'hi havia tres (aquesta, la de docs\app.js i el JSON) i van divergir
# sense que cap prova ho vegi, perque aquestes nomes miraven substrings. Ara es
# prova el que importa: que el JSON es llegeix, que porta el que ha de portar, i
# que si no hi es NO hi ha cap text de reserva a que caure.
$eload = _LoadEmailTextos
AssertEq ([bool]($eload.Contains('assumpte') -and $eload.Contains('cos') -and $eload.Contains('bcc'))) $true '_LoadEmailTextos: assumpte, cos i bcc'
AssertEq ([bool]([string]$eload['cos'] -like '*{REQUERIMENTS}*')) $true '_LoadEmailTextos: el cos porta {REQUERIMENTS}'
AssertEq ([bool]([string]$eload['assumpte'] -like '*{ID_GIA}*')) $true '_LoadEmailTextos: l''assumpte porta {ID_GIA}'
AssertEq ([bool](([string]$eload['cos']).Contains('seuelectronica'))) $true '_LoadEmailTextos: el cos porta l''enllac de la seu'
# Les DUES seus (catala i castella): l'enllac castella nomes era al JSON, i es
# justament el que les copies hardcodejades s'havien deixat.
AssertEq ([regex]::Matches([string]$eload['cos'], 'seuelectronica\.cornella\.cat').Count) 2 '_LoadEmailTextos: hi ha els dos enllacos de la seu (CA i ES)'


# Sense fitxer, PETA: es la regla que el projecte ja aplica als catalegs.
$eTmpRoot = $RepoRoot
try {
    $RepoRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('sense-textos-' + [Guid]::NewGuid().ToString('N'))
    $ePeta = $false
    try { [void](_LoadEmailTextos) } catch { $ePeta = $true }
    AssertEq $ePeta $true '_LoadEmailTextos: sense fitxer PETA (cap text de reserva)'
} finally { $RepoRoot = $eTmpRoot }

# La llista de CCO surt del mateix JSON, no del codi.
$ebcc = @(_EmailBccDeJson (Read-JsonFile (_EmailTextosPath)))
AssertEq ($ebcc.Count) 4 '_EmailBccDeJson: les 4 adreces surten del JSON'
AssertEq (@($ebcc | Where-Object { $_.Default }).Count) 1 '_EmailBccDeJson: nomes una va marcada per defecte'
AssertEq ([bool](@($ebcc)[0].Addr -like '*@*')) $true '_EmailBccDeJson: son adreces de correu'
AssertEq (@(_EmailBccDeJson $null).Count) 0 '_EmailBccDeJson: sense objecte, llista buida'
AssertEq (@(_EmailBccDeJson ([pscustomobject]@{ bcc = @() })).Count) 0 '_EmailBccDeJson: bcc buit, llista buida'

Write-Host "`n--- Cap copia dels textos ni de les adreces al codi (guard) ---"
# PER QUE. El comentari d'EmailTextos.ps1 deia "han de coincidir amb
# EMAIL_TEXTOS_DEFAULT de docs\app.js i amb el email-textos.json": tres copies
# lligades per un comentari. Aquest guard ho substitueix.
$gRoot = Split-Path -Parent (Split-Path -Parent $TestsDir)
$gFonts = @()
$gFonts += @(Get-ChildItem -Path (Join-Path $gRoot 'suport') -Recurse -Filter *.ps1 -File | Where-Object { $_.FullName -notlike '*tests*' })
$gFonts += @(Get-ChildItem -Path (Join-Path $gRoot 'docs') -Filter *.js -File)
$gAmbAdreca = @()
$gAmbTextos = @()
foreach ($gf in $gFonts) {
    $gt = Get-Content -LiteralPath $gf.FullName -Raw -Encoding UTF8
    if ($gt -match '[A-Za-z0-9._%+-]+@aj-cornella\.cat') { $gAmbAdreca += $gf.Name }
    # Un tros llarg i literal del cos del correu: si algu el torna a encastar,
    # aquesta frase hi sera.
    if ($gt.Contains('no la presenteu per parts')) { $gAmbTextos += $gf.Name }
}
AssertEq $gAmbAdreca.Count 0 ('cap adreca @aj-cornella.cat escrita al codi' + $(if ($gAmbAdreca.Count) { ' -> ' + ($gAmbAdreca -join ', ') } else { '' }))
AssertEq $gAmbTextos.Count 0 ('cap copia del cos del correu al codi' + $(if ($gAmbTextos.Count) { ' -> ' + ($gAmbTextos -join ', ') } else { '' }))

Write-Host "`n--- EnviarCorreu.ps1: diagnostic d'error d'EmailJS (pura) ---"
$e403 = _EmailJsErrorText 403 'API calls are disabled for non-browser applications' '(403) Prohibido'
AssertEq ([bool]($e403 -like '*HTTP 403*')) $true '_EmailJsErrorText: mostra l''estat HTTP'
AssertEq ([bool]($e403 -like '*non-browser applications*')) $true '_EmailJsErrorText: mostra el motiu real d''EmailJS (cos de la resposta)'
AssertEq ([bool]($e403 -like '*Security*')) $true '_EmailJsErrorText: 403 dona la guia del panell (Account -> Security)'
AssertEq ([bool]($e403 -like '*Private key*')) $true '_EmailJsErrorText: 403 recorda comprovar la Private key'
$e200 = _EmailJsErrorText 500 'boom' ''
AssertEq ([bool]($e200 -like '*Security*')) $false '_EmailJsErrorText: la guia del 403 NOMES surt en un 403'
AssertEq ([bool]($e200 -like '*boom*')) $true '_EmailJsErrorText: mostra el cos tambe en altres estats'
$eNoBody = _EmailJsErrorText 0 '' '(407) Proxy'
AssertEq ([bool]($eNoBody -like '*(407) Proxy*')) $true '_EmailJsErrorText: sense estat ni cos, cau al missatge de .NET'

Write-Host "`n--- EnviarCorreu.ps1: destinatari per defecte (Rao social + Rep. legal) ---"
$d2 = _CorreuDestinatarisPerDefecte 'rao@x.cat' 'rep@x.cat'
AssertEq $d2.Text 'rao@x.cat; rep@x.cat' '_CorreuDestinatarisPerDefecte: dues adreces diferents, totes dues'
AssertEq $d2.Compte 2 '_CorreuDestinatarisPerDefecte: compta 2 quan son diferents'
AssertEq ([bool]$d2.Duplicat) $false '_CorreuDestinatarisPerDefecte: diferents no es duplicat'
$dDup = _CorreuDestinatarisPerDefecte 'Igual@X.cat' 'igual@x.CAT'
AssertEq $dDup.Text 'Igual@X.cat' '_CorreuDestinatarisPerDefecte: mateixa adreca (ignora majuscules), nomes un cop'
AssertEq ([bool]$dDup.Duplicat) $true '_CorreuDestinatarisPerDefecte: mateixa adreca marca Duplicat'
$dRepBuit = _CorreuDestinatarisPerDefecte 'rao@x.cat' ''
AssertEq $dRepBuit.Text 'rao@x.cat' '_CorreuDestinatarisPerDefecte: nomes Rao social si falta el Rep. legal'
AssertEq ([bool]$dRepBuit.Duplicat) $false '_CorreuDestinatarisPerDefecte: una sola adreca no es duplicat'
$dRaoBuit = _CorreuDestinatarisPerDefecte '  ' 'rep@x.cat'
AssertEq $dRaoBuit.Text 'rep@x.cat' '_CorreuDestinatarisPerDefecte: nomes Rep. legal si falta la Rao social'
$dBuit = _CorreuDestinatarisPerDefecte '' ''
AssertEq $dBuit.Text '' '_CorreuDestinatarisPerDefecte: cap adreca, text buit'
AssertEq $dBuit.Compte 0 '_CorreuDestinatarisPerDefecte: cap adreca, compte 0'

Write-Host "`n--- EmailQuota.ps1: comptador d'EmailJS (pura) ---"
AssertEq (_QuotaMesActual ([datetime]'2026-09-03')) '2026-09' '_QuotaMesActual: yyyy-MM'
$qNou = _QuotaNormalitza $null '2026-09'
AssertEq $qNou.enviats 0 '_QuotaNormalitza: sense registre, comptador a 0'
AssertEq $qNou.limit 150 '_QuotaNormalitza: limit per defecte 150 (reserva de 50 sobre 200)'
$qMateix = _QuotaNormalitza ([pscustomobject]@{ mes='2026-09'; enviats=12; limit=150 }) '2026-09'
AssertEq $qMateix.enviats 12 '_QuotaNormalitza: mateix mes, es conserva el comptador'
$qAltre = _QuotaNormalitza ([pscustomobject]@{ mes='2026-08'; enviats=140; limit=150 }) '2026-09'
AssertEq $qAltre.enviats 0 '_QuotaNormalitza: mes NOU, el comptador es reinicia'
AssertEq (_QuotaRestant $qMateix) 138 '_QuotaRestant: limit menys enviats'
AssertEq (_QuotaRestant (_QuotaNormalitza ([pscustomobject]@{ mes='2026-09'; enviats=999; limit=150 }) '2026-09')) 0 '_QuotaRestant: mai negatiu'
AssertEq ((_QuotaSuma $qMateix 3).enviats) 15 '_QuotaSuma: suma els enviaments'

Write-Host "`n--- Recordatoris.ps1: campanyes i dates (pures) ---"
$rcCamps = @(_RecCampanyes)
AssertEq $rcCamps.Count 2 '_RecCampanyes: dues campanyes'
AssertEq ($rcCamps[0].Clau) 'requeriments' '_RecCampanyes: la primera es requeriments'
AssertEq ($rcCamps[1].Clau) 'precintes' '_RecCampanyes: la segona es precintes'
AssertEq (@($rcCamps[0].Estats) -join ',') 'Requeriment' '_RecCampanyes: requeriments -> estat Requeriment'
AssertEq (@($rcCamps[1].Estats) -join ',') 'Precinte / Cessament' '_RecCampanyes: precintes -> estat Precinte / Cessament'
# Els estats de les dues campanyes han de ser DISJUNTS: una activitat no pot
# rebre els dos recordatoris alhora.
$rcTots = @($rcCamps[0].Estats) + @($rcCamps[1].Estats)
AssertEq (@($rcTots | Sort-Object -Unique).Count) $rcTots.Count '_RecCampanyes: cap estat surt a dues campanyes'
AssertEq ((_RecCampanyaPerClau 'precintes').Nom) 'Precintes' '_RecCampanyaPerClau: troba per clau'
AssertEq ([string](_RecCampanyaPerClau 'inventada')) '' '_RecCampanyaPerClau: clau desconeguda -> null'

$rcAvui = [datetime]'2026-09-03'
AssertEq (_RecDiesDes '2026-09-03' $rcAvui) 0 '_RecDiesDes: avui = 0 dies'
AssertEq (_RecDiesDes '2026-08-04' $rcAvui) 30 '_RecDiesDes: 30 dies'
AssertEq (_RecDiesDes '' $rcAvui) -1 '_RecDiesDes: data buida -> -1'
AssertEq (_RecDiesDes 'demà' $rcAvui) -1 '_RecDiesDes: data il-legible -> -1'

Write-Host "`n--- Recordatoris.ps1: a qui li toca (pura) ---"
function _RcAct($gia, $estat, $data) {
    return [pscustomobject]@{
        id_gia = $gia; titular = 'Titular ' + $gia; expedient = 'EXP'; estat_actual = $estat
        informes = @([pscustomobject]@{ data = $data; conclusio_breu = $estat; ignorat = $false })
    }
}
$rcCfg = @{ periodicitatDies = 60; esperaInicialDies = 30; maxPerTanda = 15 }
$rcBuit = @{ ultim = ''; compte = 0; excloure = $false; enviaments = @() }

$t1 = _RecToca (_RcAct '' 'Requeriment' '2026-01-01') $rcCfg $rcBuit $rcAvui
AssertEq ([bool]$t1.Toca) $false '_RecToca: sense ID GIA no toca'
AssertEq $t1.Motiu 'sense ID GIA' '_RecToca: i ho diu'

$t2 = _RecToca (_RcAct '100' 'Requeriment' '2026-01-01') $rcCfg @{ ultim=''; compte=0; excloure=$true; enviaments=@() } $rcAvui
AssertEq ([bool]$t2.Toca) $false '_RecToca: activitat exclosa no toca'

$t3 = _RecToca (_RcAct '100' 'Requeriment' '2026-08-20') $rcCfg $rcBuit $rcAvui
AssertEq ([bool]$t3.Toca) $false '_RecToca: dins de l''espera inicial no toca (14 de 30 dies)'

# Just al limit de l'espera inicial: 30 dies -> SI que toca.
$t4 = _RecToca (_RcAct '100' 'Requeriment' '2026-08-04') $rcCfg $rcBuit $rcAvui
AssertEq ([bool]$t4.Toca) $true '_RecToca: just al limit de l''espera inicial, toca'

$t5 = _RecToca (_RcAct '100' 'Requeriment' '2026-01-01') $rcCfg @{ ultim='2026-08-20'; compte=1; excloure=$false; enviaments=@() } $rcAvui
AssertEq ([bool]$t5.Toca) $false '_RecToca: enviat fa 14 dies, encara no toca (cada 60)'

# Just al limit de la periodicitat: 60 dies -> torna a tocar.
$t6 = _RecToca (_RcAct '100' 'Requeriment' '2026-01-01') $rcCfg @{ ultim='2026-07-05'; compte=1; excloure=$false; enviaments=@() } $rcAvui
AssertEq ([bool]$t6.Toca) $true '_RecToca: just al limit de la periodicitat, torna a tocar'

$t7 = _RecToca (_RcAct '100' 'Requeriment' '') $rcCfg $rcBuit $rcAvui
AssertEq ([bool]$t7.Toca) $false '_RecToca: sense data d''informe no toca'

# Espera inicial 0: es pot avisar de seguida.
$t8 = _RecToca (_RcAct '100' 'Requeriment' '2026-09-03') @{ periodicitatDies=60; esperaInicialDies=0; maxPerTanda=5 } $rcBuit $rcAvui
AssertEq ([bool]$t8.Toca) $true '_RecToca: amb espera inicial 0, toca el mateix dia'

Write-Host "`n--- Recordatoris.ps1: seleccio i ordre (pura) ---"
$rcDb = [pscustomobject]@{
    actualitzat_el = '2026-09-01T10:00:00'
    activitats = @(
        (_RcAct '100' 'Requeriment' '2026-01-10')
        (_RcAct '200' 'Requeriment' '2026-02-10')
        (_RcAct '300' 'Precinte / Cessament' '2026-01-05')
        (_RcAct '400' 'Favorable' '2026-01-05')
        (_RcAct ''    'Requeriment' '2026-01-05')
    )
}
$rcHist = @{ '100' = @{ ultim='2026-06-01'; compte=2; excloure=$false; enviaments=@('2026-06-01') } }
$rcRes = _RecDueActivitats $rcDb (_RecCampanyaPerClau 'requeriments') $rcCfg $rcHist $rcAvui
$rcFiles = @($rcRes.Files)
AssertEq $rcFiles.Count 3 '_RecDueActivitats: nomes les activitats en estat Requeriment'
AssertEq $rcRes.SenseGia 1 '_RecDueActivitats: compta les que no tenen GIA'
# Ordre: primer la que no ha rebut mai cap avis (ultim buit).
AssertEq ([string]$rcFiles[0].Ultim) '' '_RecDueActivitats: els mai avisats van primer'
AssertEq ([bool](@($rcFiles | Where-Object { $_.Id -eq '400' }).Count)) $false '_RecDueActivitats: un altre estat no hi entra'
$rcPre = @(( _RecDueActivitats $rcDb (_RecCampanyaPerClau 'precintes') $rcCfg @{} $rcAvui).Files)
AssertEq $rcPre.Count 1 '_RecDueActivitats: la campanya de precintes nomes veu els precintats'
AssertEq ([string]$rcPre[0].Id) '300' '_RecDueActivitats: i es el GIA correcte'

Write-Host "`n--- Recordatoris.ps1: textos, variables i historial ---"
$rcTx = _RecDefaultTextos 'requeriments'
AssertEq ([bool]([string]$rcTx['cos']).Contains('no cal que tingueu en compte aquest correu')) $true '_RecDefaultTextos: hi ha l''avis de "ja presentada" (CA)'
AssertEq ([bool]([string]$rcTx['cos']).Contains('no tenga en cuenta este correo')) $true '_RecDefaultTextos: ...i en castella'
AssertEq ([bool]([string]$rcTx['cos']).Contains('Article 5. Condicionant per a la transmissi')) $true '_RecDefaultTextos: hi ha l''article 5 de l''Ordenanca'
AssertEq ([bool]([string]$rcTx['cos']).Contains('{ID_GIA}')) $true '_RecDefaultTextos: el cos identifica l''activitat amb {ID_GIA}'
AssertEq ([bool]([string]$rcTx['assumpte']).Contains('{ID_GIA}')) $true '_RecDefaultTextos: l''assumpte porta {ID_GIA}'
$rcTxP = _RecDefaultTextos 'precintes'
AssertEq ([bool]([string]$rcTxP['cos']).Contains('Article 5. Condicionant per a la transmissi')) $true '_RecDefaultTextos: precintes tambe porta l''article 5'
AssertEq ([bool]([string]$rcTxP['cos']).Contains('no cal que tingueu en compte aquest correu')) $true '_RecDefaultTextos: precintes tambe porta l''avis'
Assert ([string]$rcTx['cos'] -ne [string]$rcTxP['cos']) 'Cada campanya te el seu text propi'

$rcRow = [pscustomobject]@{ Id='1463'; Titular='Bar Pepe'; Adreca='C/ Major 1'; Activitat='BAR'; DataInforme='2026-01-10' }
$rcFill = _RecFillPh 'GIA {ID_GIA} / {TITULAR} / {ADRECA} / {ACTIVITAT} / {DATA_INFORME}' $rcRow
AssertEq $rcFill 'GIA 1463 / Bar Pepe / C/ Major 1 / BAR / 2026-01-10' '_RecFillPh: substitueix totes les variables'

# Anada i tornada de l'historial AMB EL JSON PEL MIG: es on es trenca sempre
# (ConvertFrom-Json torna PSCustomObjects, no hashtables).
$h1 = _RecHistorialActualitza @{} '1463' '2026-09-03'
AssertEq ([string]$h1['1463']['ultim']) '2026-09-03' '_RecHistorialActualitza: apunta la data'
AssertEq ([int]$h1['1463']['compte']) 1 '_RecHistorialActualitza: compta 1'
$h2 = _RecHistorialActualitza $h1 '1463' '2026-11-03'
AssertEq ([int]$h2['1463']['compte']) 2 '_RecHistorialActualitza: el segon avis suma'
AssertEq (@($h2['1463']['enviaments']).Count) 2 '_RecHistorialActualitza: guarda l''historial d''enviaments'
$rcJson = ([pscustomobject]@{ historial = [pscustomobject]@{ requeriments = [pscustomobject]$h2 } }) | ConvertTo-Json -Depth 12
$rcBack = $rcJson | ConvertFrom-Json
$h3 = _RecHistorialAMapa $rcBack.historial.requeriments
$e3 = _RecHistEntrada $h3 '1463'
AssertEq ([string]$e3.ultim) '2026-11-03' 'Historial: sobreviu a l''anada i tornada pel JSON'
AssertEq ([int]$e3.compte) 2 'Historial: el comptador sobreviu al JSON'
$hExc = _RecHistorialExclou $h2 '1463' $true
AssertEq ([bool](_RecHistEntrada $hExc '1463').excloure) $true '_RecHistorialExclou: marca l''exclusio'
AssertEq ([int](_RecHistEntrada $hExc '1463').compte) 2 '_RecHistorialExclou: no perd el que ja hi havia'

Write-Host "`n--- Recordatoris.ps1: configuracio i tasca programada ---"
$rcDef = _RecDefaultConfig 'requeriments'
AssertEq ([bool]$rcDef['actiu']) $false '_RecDefaultConfig: neix APAGADA (no envia res fins que l''encenguis)'
AssertEq ([string]$rcDef['mode']) 'manual' '_RecDefaultConfig: neix en manual'
$rcN = _RecNormalitzaConfig ([pscustomobject]@{ actiu=$true; mode='auto'; periodicitatDies=90 }) 'requeriments'
AssertEq ([bool]$rcN['actiu']) $true '_RecNormalitzaConfig: respecta el que ve del JSON'
AssertEq ([string]$rcN['mode']) 'auto' '_RecNormalitzaConfig: mode auto'
AssertEq ([int]$rcN['periodicitatDies']) 90 '_RecNormalitzaConfig: periodicitat del JSON'
AssertEq ([int]$rcN['maxPerTanda']) 15 '_RecNormalitzaConfig: el que no ve, del defecte'
$rcN2 = _RecNormalitzaConfig ([pscustomobject]@{ periodicitatDies=0; mode='tonteria'; cos='' }) 'requeriments'
AssertEq ([int]$rcN2['periodicitatDies']) 60 '_RecNormalitzaConfig: un valor absurd cau al defecte'
AssertEq ([string]$rcN2['mode']) 'manual' '_RecNormalitzaConfig: un mode desconegut cau a manual'
Assert ([string]$rcN2['cos'] -ne '') '_RecNormalitzaConfig: un cos buit no pot deixar el correu sense text'
$rcN3 = _RecNormalitzaConfig ([pscustomobject]@{ esperaInicialDies=0 }) 'requeriments'
AssertEq ([int]$rcN3['esperaInicialDies']) 0 '_RecNormalitzaConfig: espera inicial 0 SI que es valida'

# Les rutes han d'anar ENTRE COMETES: el clone de l'usuari te espais.
$rcTr = _RecSchtasksTr 'C:\Win\powershell.exe' 'I:\5.- Sergi Fadurdo\suport\RecordatorisAuto.ps1'
AssertEq ([bool]$rcTr.Contains('"C:\Win\powershell.exe"')) $true '_RecSchtasksTr: l''executable va entre cometes'
AssertEq ([bool]$rcTr.Contains('"I:\5.- Sergi Fadurdo\suport\RecordatorisAuto.ps1"')) $true '_RecSchtasksTr: el script (amb espais) va entre cometes'
$rcArgv = @(_RecSchtasksArgv 'InformesCornella-Recordatoris' 'C:\p.exe' 'C:\s.ps1' '09:00')
AssertEq ([string]$rcArgv[0]) '/Create' '_RecSchtasksArgv: crea la tasca'
AssertEq ([bool]($rcArgv -contains '/F')) $true '_RecSchtasksArgv: /F per sobreescriure-la'
$rcIdxSt = [array]::IndexOf($rcArgv, '/ST')
Assert ($rcIdxSt -ge 0) '_RecSchtasksArgv: hi ha l''hora d''inici (/ST)'
AssertEq ([string]$rcArgv[$rcIdxSt + 1]) '09:00' '_RecSchtasksArgv: i l''hora va just despres de /ST'

AssertEq (_RecAntiguitatDb ([pscustomobject]@{ actualitzat_el='2026-08-04T09:00:00' }) $rcAvui) 30 '_RecAntiguitatDb: 30 dies'
AssertEq (_RecAntiguitatDb ([pscustomobject]@{ }) $rcAvui) -1 '_RecAntiguitatDb: sense data -> -1'

Write-Host "`n--- ControlsCpEmail.ps1: avisos de control periodic per correu ---"
$ccdef = _DefaultControlsCpEmail
AssertEq (@($ccdef.Keys).Count) 2 '_DefaultControlsCpEmail: 2 claus (assumpte, cos)'
AssertEq ([bool]($ccdef.Contains('assumpte') -and $ccdef.Contains('cos'))) $true '_DefaultControlsCpEmail: te assumpte i cos'
AssertEq ([bool]([string]$ccdef['cos'] -like '*{ACTIVITAT}*' -and [string]$ccdef['cos'] -like '*{ADRECA}*' -and [string]$ccdef['cos'] -like '*{PROPER_CP}*')) $true '_DefaultControlsCpEmail: el cos te les variables clau'
AssertEq ([bool]([string]$ccdef['assumpte'] -like '*{ID_GIA}*')) $true '_DefaultControlsCpEmail: assumpte te {ID_GIA}'
# Destinataris: titular a To, representant a CC (i marxa enrere si en falta un).
$rc1 = _ControlsCpRecipients 'titular@x.cat' 'rep@x.cat'
AssertEq $rc1.To 'titular@x.cat' '_ControlsCpRecipients: titular a To'
AssertEq $rc1.Cc 'rep@x.cat'     '_ControlsCpRecipients: representant a CC'
AssertEq $rc1.Ok $true           '_ControlsCpRecipients: dos correus -> Ok'
$rc2 = _ControlsCpRecipients '' 'rep@x.cat'
AssertEq $rc2.To 'rep@x.cat' '_ControlsCpRecipients: sense titular -> representant a To'
AssertEq $rc2.Cc ''          '_ControlsCpRecipients: sense titular -> CC buit'
$rc3 = _ControlsCpRecipients 'titular@x.cat' ''
AssertEq $rc3.To 'titular@x.cat' '_ControlsCpRecipients: sense representant -> titular a To'
$rc4 = _ControlsCpRecipients '  ' 'no-es-un-correu'
AssertEq $rc4.Ok $false '_ControlsCpRecipients: cap correu valid -> Ok=false'
# Substitucio de variables amb una fila d'activitat.
$fila = [pscustomobject]@{ ActPrincipal='BAR'; Adreca='C/ Major 1'; Id='361'; RaoSocial='ACME SL'; ProperCP='10/01/2026'; DataControlPer='10/01/2024' }
$sub = _FillControlsCpPh 'Activitat {ACTIVITAT} a {ADRECA} (GIA {ID_GIA}), data {PROPER_CP}' $fila
AssertEq $sub 'Activitat BAR a C/ Major 1 (GIA 361), data 10/01/2026' '_FillControlsCpPh: substitueix les variables'
# HTML: escapa, negreta i enllacos.
# L'HTML d'aquest correu el fa ara _CosAHtml/_TextToHtml (EnviarCorreu.ps1),
# les mateixes que els recordatoris. Abans hi havia _ControlsCpEmailHtml
# -identica linia a linia a _RecCosHtml- i un _ControlsCpLineHtml propi.
AssertEq (_TextToHtml 'a & b < c') 'a &amp; b &lt; c' 'correu: escapa &, <'
AssertEq (_TextToHtml '**negreta**') '<b>negreta</b>' 'correu: **negreta** -> <b>'
AssertEq (_TextToHtml 'veure http://x.cat/a ok') 'veure <a href="http://x.cat/a">http://x.cat/a</a> ok' 'correu: enllac http -> <a>'
$html = _CosAHtml "linia1`n`nlinia2"
AssertEq ([bool]($html -like '*<div>linia1</div>*' -and $html -like '*<div>linia2</div>*')) $true '_CosAHtml: una linia = un <div>'
AssertEq ([bool]($html -like '*height:8px*')) $true '_CosAHtml: una linia buida es un espaiador'

Write-Host "`n--- Informes.ps1: _EstatActualActivitat (estat = conclusio breu del darrer informe fiable, per data) ---"
AssertEq (_EstatActualActivitat $null) '' '_EstatActualActivitat null -> buit'
AssertEq (_EstatActualActivitat @()) '' '_EstatActualActivitat llista buida -> buit'
$infsTotIgnorats = @(
    [pscustomobject]@{ data = '2026-01-01'; ignorat = $true; conclusio_breu = 'Requeriment' },
    [pscustomobject]@{ data = '2026-02-01'; ignorat = $true; conclusio_breu = 'Favorable' }
)
AssertEq (_EstatActualActivitat $infsTotIgnorats) '' '_EstatActualActivitat tots ignorats -> buit'
$infsNormal = @(
    [pscustomobject]@{ data = '2026-01-01'; ignorat = $false; conclusio_breu = 'Requeriment' },
    [pscustomobject]@{ data = '2026-02-01'; ignorat = $true;  conclusio_breu = 'Favorable' },
    [pscustomobject]@{ data = '2026-03-01'; ignorat = $false; conclusio_breu = 'FI Requeriment' }
)
AssertEq (_EstatActualActivitat $infsNormal) 'FI Requeriment' '_EstatActualActivitat es queda amb el darrer NO ignorat (per data)'
# QUIN informe decideix l'estat: el mateix que fa servir _EstatActualActivitat.
# Ho necessita "Comprovar Excel" per dir-ne la data (INFORME ENGINYER dd/MM/aaaa).
AssertEq ($null -eq (_InformeQueDeterminaEstat $null)) $true '_InformeQueDeterminaEstat null -> $null'
AssertEq ($null -eq (_InformeQueDeterminaEstat @())) $true '_InformeQueDeterminaEstat llista buida -> $null'
AssertEq ($null -eq (_InformeQueDeterminaEstat $infsTotIgnorats)) $true '_InformeQueDeterminaEstat tots ignorats -> $null'
AssertEq (_InformeQueDeterminaEstat $infsNormal).data '2026-03-01' '_InformeQueDeterminaEstat: el darrer NO ignorat (SALTA el del 02, ignorat)'
AssertEq (_InformeQueDeterminaEstat $infsNormal).conclusio_breu (_EstatActualActivitat $infsNormal) '_InformeQueDeterminaEstat i _EstatActualActivitat parlen del MATEIX informe'

Write-Host "`n--- Informes.ps1: _DataInformeDdMmAaaa ---"
AssertEq (_DataInformeDdMmAaaa '2022-11-11') '11/11/2022' '_DataInformeDdMmAaaa: yyyy-MM-dd -> dd/MM/yyyy'
AssertEq (_DataInformeDdMmAaaa '2026-01-05') '05/01/2026' '_DataInformeDdMmAaaa: conserva els zeros'
AssertEq (_DataInformeDdMmAaaa '2026-03-01T00:00:00') '01/03/2026' '_DataInformeDdMmAaaa: ignora el que hi hagi despres de la data'
AssertEq (_DataInformeDdMmAaaa '') '' '_DataInformeDdMmAaaa: buit -> buit'
AssertEq (_DataInformeDdMmAaaa $null) '' '_DataInformeDdMmAaaa: null -> buit'
AssertEq (_DataInformeDdMmAaaa '11/11/2022') '' '_DataInformeDdMmAaaa: format desconegut -> buit (no inventa res)'

Write-Host "`n--- Informes.ps1: _GiaFromFolderName / _CarpetaActivitat ---"
$p = 'I:\Activitats\Informes\2025-1-2563 GIA 361 - RC112- KRICHI BEJAUI HOSTELERIA, SL\20260710_Req4.docx'
AssertEq (_GiaFromFolderName $p) '361' '_GiaFromFolderName treu "GIA 361" de la carpeta'
AssertEq (_CarpetaActivitat $p) '2025-1-2563 GIA 361 - RC112- KRICHI BEJAUI HOSTELERIA, SL' '_CarpetaActivitat = carpeta pare'
AssertEq (_GiaFromFolderName 'I:\Informes\sense marca\x.docx') '' '_GiaFromFolderName sense GIA -> buit'

Write-Host "`n--- Seguiment.ps1: espaiat de les anotacions datades ---"
# Cas real: al punt 6 d'un informe hi sortia un forat entre "No s'aporta." i
# "S'aporta.", i a la resta de punts no. Motiu: NOMES aquell requeriment portava
# un w:spacing/@w:after (el que el separa del punt seguent) i l'anotacio, que
# clona el pPr del requeriment, se l'enduia; a la segona ronda hi havia doncs un
# 'after' entre les dues linies datades.
$anW = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
function New-XmlProva([string]$pPrIntern) {
    $x = New-Object System.Xml.XmlDocument
    $x.PreserveWhitespace = $true
    $x.LoadXml("<w:document xmlns:w=""$anW""><w:body><w:p><w:pPr>$pPrIntern</w:pPr><w:r><w:t>requeriment</w:t></w:r></w:p></w:body></w:document>")
    $nsm = New-Object System.Xml.XmlNamespaceManager($x.NameTable)
    $nsm.AddNamespace('w', $anW)
    return [pscustomobject]@{ Xml = $x; Ns = $nsm; Body = $x.SelectSingleNode('//w:body', $nsm) }
}
function SpacingDe($node, $xi) {
    $sp = $node.SelectSingleNode('w:pPr/w:spacing', $xi.Ns)
    if ($null -eq $sp) { return '-' }
    return ('before=' + [string]$sp.GetAttribute('before', $anW) + ' after=' + [string]$sp.GetAttribute('after', $anW))
}
# (a) Requeriment SENSE 'after' (la majoria): l'anotacio tampoc no n'ha de tenir.
$xiA = New-XmlProva '<w:pStyle w:val="Prrafodelista"/>'
$reqA = $xiA.Body.SelectSingleNode('w:p', $xiA.Ns)
$annA1 = _MakeAnnotationParagraphXml $xiA $reqA '09/06/2026' "No s'aporta." $true $true $false (_TakeSpacingAfterXml $xiA $reqA)
AssertEq (SpacingDe $annA1 $xiA) 'before=200 after=' 'anotacio 1a: nomes espai a sobre'
$annA2 = _MakeAnnotationParagraphXml $xiA $reqA '03/08/2026' "S'aporta." $false $false $false (_TakeSpacingAfterXml $xiA $annA1)
AssertEq (SpacingDe $annA2 $xiA) '-' 'anotacio 2a: sense cap espai (van seguides)'
# (b) Requeriment AMB 'after' (el cas del punt 6): l'espai s'ha de MOURE, no
#     copiar. Entre les dues linies datades no n'hi pot quedar cap.
$xiB = New-XmlProva '<w:pStyle w:val="Prrafodelista"/><w:spacing w:after="240"/>'
$reqB = $xiB.Body.SelectSingleNode('w:p', $xiB.Ns)
$afterB = _TakeSpacingAfterXml $xiB $reqB
AssertEq $afterB '240' '_TakeSpacingAfterXml: retorna l''espai de sota del requeriment'
AssertEq (SpacingDe $reqB $xiB) 'before= after=' '_TakeSpacingAfterXml: i l''hi TREU (ja no separa el requeriment de la seva anotacio)'
$annB1 = _MakeAnnotationParagraphXml $xiB $reqB '09/06/2026' "No s'aporta." $true $true $false $afterB
AssertEq (SpacingDe $annB1 $xiB) 'before=200 after=240' 'anotacio 1a: hereta l''espai que tancava el bloc'
$afterB2 = _TakeSpacingAfterXml $xiB $annB1
AssertEq $afterB2 '240' 'l''espai es torna a prendre de l''anotacio anterior'
$annB2 = _MakeAnnotationParagraphXml $xiB $reqB '03/08/2026' "S'aporta." $false $false $false $afterB2
AssertEq (SpacingDe $annB1 $xiB) 'before=200 after=' 'anotacio 1a: ES QUEDA SENSE espai a sota (aqui hi havia el forat)'
AssertEq (SpacingDe $annB2 $xiB) 'before= after=240' 'anotacio 2a: ara es ella qui tanca el bloc'
# (c) Sub-punt: l'espai de sota del sub-punt mana, pero tampoc no es pot quedar a
#     l'anotacio del mig.
$xiC = New-XmlProva '<w:pStyle w:val="Prrafodelista"/>'
$reqC = $xiC.Body.SelectSingleNode('w:p', $xiC.Ns)
$annC1 = _MakeAnnotationParagraphXml $xiC $reqC '09/06/2026' "No s'aporta." $true $true $true ''
AssertEq (SpacingDe $annC1 $xiC) 'before=200 after=240' 'sub-punt: la 1a anotacio porta espai a sobre i a sota'
$annC2 = _MakeAnnotationParagraphXml $xiC $reqC '03/08/2026' "S'aporta." $false $false $true (_TakeSpacingAfterXml $xiC $annC1)
AssertEq (SpacingDe $annC1 $xiC) 'before=200 after=' 'sub-punt: la 1a anotacio perd l''espai de sota quan en ve una altra'
AssertEq (SpacingDe $annC2 $xiC) 'before= after=240' 'sub-punt: el tanca la darrera anotacio'

Write-Host "`n--- Seguiment.ps1: segell d'ultima execucio de les eines ---"
# La marca es desa amb (Get-Date).ToString('o'), o sigui hora LOCAL amb el seu
# desplacament, i es torna a llegir en hora local: el viatge d'anada i tornada ha
# de donar la mateixa hora. Es comprova aixi i no amb una cadena fixa perque
# aquesta maquina pot anar en una zona horaria diferent de la del PC de l'usuari.
$isoLocal = ([datetime]'2026-07-30T14:22:05').ToString('o')
AssertEq (_FormatRunStamp $isoLocal) '30/07/26 14:22' '_FormatRunStamp: ISO -> dd/MM/aa HH:mm (anada i tornada en hora local)'
AssertEq (_FormatRunStamp '') '(mai)' '_FormatRunStamp: buit -> (mai)'
AssertEq (_FormatRunStamp '   ') '(mai)' '_FormatRunStamp: nomes espais -> (mai)'
AssertEq (_FormatRunStamp 'aixo no es una data') '(mai)' '_FormatRunStamp: text que no es data -> (mai)'
AssertEq (_FormatRunStamp $null) '(mai)' '_FormatRunStamp: $null -> (mai)'
# Els tipus d'informe i les pantalles de sistema NO porten segell; qualsevol
# altra accio del menu es una eina i si (aixi una rajola nova hi entra sola).
Assert ('nou' -in $Script:AccionsSenseSegell)          'AccionsSenseSegell: "nou" no porta segell'
Assert ('seguiment' -in $Script:AccionsSenseSegell)    'AccionsSenseSegell: el seguiment d''un informe tampoc'
Assert ('config' -in $Script:AccionsSenseSegell)       'AccionsSenseSegell: la configuracio tampoc'
Assert ('seguimentgia' -notin $Script:AccionsSenseSegell) 'AccionsSenseSegell: l''eina Seguiment del GIA SI que en porta'
Assert ('comprovarexcel' -notin $Script:AccionsSenseSegell) 'AccionsSenseSegell: Comprovar Excel SI que en porta'
Assert ('precintades' -notin $Script:AccionsSenseSegell)   'AccionsSenseSegell: la rajola d''enllac SI que en porta'
# Registre: escriure i tornar a llegir, i que conservi les altres eines.
$segellDir = Join-Path ([System.IO.Path]::GetTempPath()) ('segells-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $segellDir -Force)
$LocalActivitatsDirVell = $LocalActivitatsDir
$LocalActivitatsDir = $segellDir
try {
    AssertEq (_LastRunEina 'seguimentgia') '(mai)' '_LastRunEina: sense registre -> (mai)'
    _MarcaEinaUsada 'seguimentgia'
    AssertEq ([bool]((_LastRunEina 'seguimentgia') -ne '(mai)')) $true '_LastRunEina: despres de marcar-la, hi ha data'
    _MarcaEinaUsada 'comprovarexcel'
    AssertEq ([bool]((_LastRunEina 'seguimentgia') -ne '(mai)')) $true '_MarcaEinaUsada: marcar-ne una altra NO esborra la primera'
    AssertEq @(((Get-Content -LiteralPath (_EinesStatePath) -Raw | ConvertFrom-Json).PSObject.Properties)).Count 2 '_MarcaEinaUsada: un sol fitxer amb una clau per eina'
    # Les dues que tenen marca propia la fan servir si hi es.
    ([pscustomobject]@{ actualitzat_el = '2026-07-28T09:05:00Z' } | ConvertTo-Json) |
        Set-Content -LiteralPath (Join-Path $segellDir 'informes-db.json') -Encoding UTF8
    AssertEq (_LastRunEina 'informesdb') (_FormatRunStamp '2026-07-28T09:05:00Z') '_LastRunEina: informesdb fa servir actualitzat_el (mes precis)'
    AssertEq (_LastRunEina 'einaqueno') '(mai)' '_LastRunEina: una accio sense registre -> (mai)'
} finally {
    $LocalActivitatsDir = $LocalActivitatsDirVell
    Remove-Item -LiteralPath $segellDir -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "`n--- Informes.ps1: _NormalitzaExpedient / Build-ExpedientToGiaMap ---"
AssertEq (_NormalitzaExpedient '2025/1/2563')  '2025-1-2563' '_NormalitzaExpedient barres -> guions'
AssertEq (_NormalitzaExpedient '2025-01-2563') '2025-1-2563' '_NormalitzaExpedient treu zeros inicials'
AssertEq (_NormalitzaExpedient '  2025 / 1 / 2563 ') '2025-1-2563' '_NormalitzaExpedient espais'
$fakeCache = [pscustomobject]@{ ById = @{ '361' = @{ EXP_NUM = '2025/1/2563'; TITULAR = 'KRICHI' }; '99' = @{ EXP_NUM = '2024/2/10' } } }
$e2g = Build-ExpedientToGiaMap $fakeCache
AssertEq $e2g['2025-1-2563'] '361' 'Build-ExpedientToGiaMap mapa expedient->GIA'
Assert ($e2g.ContainsKey('2024-2-10')) 'Build-ExpedientToGiaMap inclou totes les files'

Write-Host "`n--- Informes.ps1: _HaDeReprocessar (escaneig incremental) ---"
$prev = [datetime]'2026-07-01T10:00:00Z'
Assert (_HaDeReprocessar ([datetime]'2026-07-05') $prev $false)            'sense entrada previa -> reprocessar (nou)'
Assert (_HaDeReprocessar ([datetime]'2026-07-05') $prev $true)             'modificat despres de l''ultima actualitzacio -> reprocessar'
Assert (-not (_HaDeReprocessar ([datetime]'2026-06-20') $prev $true))      'no tocat des de l''ultima actualitzacio -> reutilitzar'
Assert (-not (_HaDeReprocessar $prev $prev $true))                          'mateixa data exacta -> reutilitzar'

Write-Host "`n--- Informes.ps1: _FlattenInformesDb (aplanat + preservar ignorat) ---"
$fakeDb = [pscustomobject]@{
    actualitzat_el = '2026-07-01T10:00:00Z'
    activitats = @(
        [pscustomobject]@{
            id_gia='361'; expedient='2025/1/2563'; titular='KRICHI'; carpeta='2025-1-2563 GIA 361'
            informes = @(
                [pscustomobject]@{ data='2026-07-10'; fitxer='a.docx'; ruta='I:\x\a.docx'; conclusio='Vist...'; modificat='2026-07-10T00:00:00Z'; ignorat=$true;  motiu='' },
                [pscustomobject]@{ data='2025-01-14'; fitxer='b.docx'; ruta='I:\x\b.docx'; conclusio='';        modificat='2025-01-14T00:00:00Z'; ignorat=$false; motiu='sense conclusio' }
            )
        }
    )
}
$flat = _FlattenInformesDb $fakeDb
AssertEq $flat.Count 2                               '_FlattenInformesDb: 2 registres'
AssertEq $flat['I:\x\a.docx'].Gia '361'              '_FlattenInformesDb: GIA de l''activitat'
AssertEq $flat['I:\x\a.docx'].Titular 'KRICHI'       '_FlattenInformesDb: titular de l''activitat'
Assert ([bool]$flat['I:\x\a.docx'].Ignorat)          '_FlattenInformesDb: conserva ignorat=true'
Assert (-not [bool]$flat['I:\x\b.docx'].Ignorat)     '_FlattenInformesDb: conserva ignorat=false'
AssertEq $flat['I:\x\b.docx'].Motius.Count 1         '_FlattenInformesDb: motiu -> Motius'

Write-Host "`n--- Settings.ps1: _ResolveEffectiveValue (override d'aquest PC vs valor per defecte) ---"
AssertEq (_ResolveEffectiveValue 'F:\Informes' 'I:\Informes') 'F:\Informes' '_ResolveEffectiveValue amb override -> guanya l''override'
AssertEq (_ResolveEffectiveValue '' 'I:\Informes')            'I:\Informes' '_ResolveEffectiveValue buit -> per defecte'
AssertEq (_ResolveEffectiveValue '   ' 'I:\Informes')         'I:\Informes' '_ResolveEffectiveValue nomes espais -> per defecte'
AssertEq (_ResolveEffectiveValue $null 'I:\Informes')         'I:\Informes' '_ResolveEffectiveValue null -> per defecte'

Write-Host "`n--- Settings.ps1: _BuildSettingsOverrides (que es desa a settings.json) ---"
$defaults = @{ InformesDir = 'I:\Informes'; ActivitatsDir = 'I:\Activitats'; OutputDir = 'C:\Repo\Sortida' }
$valuesCanviats = @{ InformesDir = 'F:\Informes'; ActivitatsDir = 'I:\Activitats'; OutputDir = 'C:\Repo\Sortida' }
$ov1 = _BuildSettingsOverrides $valuesCanviats $defaults
AssertEq $ov1.Count 1                    '_BuildSettingsOverrides: nomes el camp canviat'
AssertEq $ov1['InformesDir'] 'F:\Informes' '_BuildSettingsOverrides: valor correcte'
Assert (-not $ov1.Contains('ActivitatsDir')) '_BuildSettingsOverrides: igual al per defecte -> no es desa'

$valuesBuits = @{ InformesDir = ''; ActivitatsDir = '   '; OutputDir = 'C:\Repo\Sortida' }
$ov2 = _BuildSettingsOverrides $valuesBuits $defaults
AssertEq $ov2.Count 0 '_BuildSettingsOverrides: camps buits -> cap override (es fa servir el per defecte)'

$valuesTots = @{ InformesDir = 'F:\Informes'; ActivitatsDir = 'F:\Activitats'; OutputDir = 'F:\Sortida' }
$ov3 = _BuildSettingsOverrides $valuesTots $defaults
AssertEq $ov3.Count 3 '_BuildSettingsOverrides: tots els camps diferents -> tots es desen'

Write-Host "`n--- _TextMatches com a filtre de les graelles (Informes / Controls periodics) ---"
# Les graelles li passen la concatenacio dels camps cercables de la fila i el
# text de cerca JA net (.Trim().ToLower()), que es el contracte de la funcio.
$filaHay = 'GIA 361 BAR PEPE Cornella de Llobregat Requeriment'
Assert (_TextMatches $filaHay 'bar')             'filtre graella: troba sense distingir majuscules'
Assert (_TextMatches $filaHay 'pepe')            'filtre graella: troba el titular'
Assert (_TextMatches $filaHay '361')             'filtre graella: troba un numero (GIA)'
Assert (_TextMatches $filaHay 'lla de llo')      'filtre graella: subcadena amb espais'
Assert (_TextMatches $filaHay '')                'filtre graella: cerca buida -> passa tot'
Assert (-not (_TextMatches $filaHay 'restaurant')) 'filtre graella: sense coincidencia -> no passa'
Assert (-not (_TextMatches '' 'x'))              'filtre graella: fila sense text i cerca amb contingut -> no passa'
