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

Write-Host "`n--- EnviarCorreu.ps1: destinatari buit = correu de prova per a un mateix ---"
$opsJo = @(@{ Addr = 'Jo@X.cat'; Default = $true }, @{ Addr = 'altre@x.cat'; Default = $false })
$rNormal = _CorreuDestinatariBuit @('tit@x.cat') @('jo@x.cat') $opsJo
AssertEq ($rNormal.To -join ',') 'tit@x.cat' '_CorreuDestinatariBuit: amb destinatari, no toca el To'
AssertEq ($rNormal.Bcc -join ',') 'jo@x.cat' '_CorreuDestinatariBuit: amb destinatari, no toca la CCO'
AssertEq ([bool]$rNormal.Prova) $false '_CorreuDestinatariBuit: amb destinatari no es prova'
$rProva = _CorreuDestinatariBuit @() @('jo@x.cat', 'altre@x.cat') $opsJo
AssertEq ($rProva.To -join ',') 'Jo@X.cat' '_CorreuDestinatariBuit: buit -> la CCO per defecte passa a destinatari'
AssertEq ($rProva.Bcc -join ',') 'altre@x.cat' '_CorreuDestinatariBuit: i surt de la CCO (no arriba dos cops); les altres CCO es queden'
AssertEq ([bool]$rProva.Prova) $true '_CorreuDestinatariBuit: buit marca Prova'
$rBlancs = _CorreuDestinatariBuit @('  ') @() $opsJo
AssertEq ($rBlancs.To -join ',') 'Jo@X.cat' '_CorreuDestinatariBuit: nomes espais compta com a buit'
AssertEq @($rBlancs.Bcc).Count 0 '_CorreuDestinatariBuit: sense CCO marcades, CCO buida'
$rSenseDef = _CorreuDestinatariBuit @() @('altre@x.cat') @(@{ Addr = 'altre@x.cat'; Default = $false })
AssertEq @($rSenseDef.To).Count 0 '_CorreuDestinatariBuit: sense CCO per defecte, To buit (el cridador ho atura)'
$rReal = _CorreuDestinatariBuit @() @() @(_CorreuBccOpcions)
AssertEq @($rReal.To).Count 1 '_CorreuDestinatariBuit: amb email-textos.json real, hi ha adreca propia'

Write-Host "`n--- EnviarCorreu.ps1: de quina activitat es l'informe (pura) ---"
# El nom real que fa _GetOutputFileName / _SeguimentOutputName.
AssertEq (_GiaDelNomFitxer '2026-09-08_Req2_GIA 1466') '1466' '_GiaDelNomFitxer: nom d''un seguiment'
AssertEq (_GiaDelNomFitxer '2026-05-29_Req1_GIA 1379') '1379' '_GiaDelNomFitxer: nom d''un requeriment nou'
# _GetUniqueOutputPath hi afegeix "_2", "_3"...: NO forma part de l'id.
AssertEq (_GiaDelNomFitxer '2026-09-08_Req2_GIA 1466_2') '1466' '_GiaDelNomFitxer: el sufix d''unicitat no entra a l''id'
AssertEq (_GiaDelNomFitxer '2026-01-02_LlicReq_GIA 900_Bar Pepe') '900' '_GiaDelNomFitxer: encara que hi vagi text al darrere'
AssertEq (_GiaDelNomFitxer 'informe antic sense res') '' '_GiaDelNomFitxer: sense GIA, cadena buida'
AssertEq (_GiaDelNomFitxer '') '' '_GiaDelNomFitxer: nom buit, cadena buida'

# La regla de decisio: el nom i la capcalera del document.
$g1 = _CorreuGiaDecideix '1466' '1466'
AssertEq $g1.Gia '1466' '_CorreuGiaDecideix: tots dos coincideixen'
AssertEq ([bool]$g1.CalPreguntar) $false '_CorreuGiaDecideix: ...i no cal preguntar'
$g2 = _CorreuGiaDecideix '1466' ''
AssertEq $g2.Gia '1466' '_CorreuGiaDecideix: nomes el nom del fitxer'
AssertEq ([bool]$g2.CalPreguntar) $false '_CorreuGiaDecideix: nomes el nom, no cal preguntar'
$g3 = _CorreuGiaDecideix '' '1466'
AssertEq $g3.Gia '1466' '_CorreuGiaDecideix: nomes la capcalera'
AssertEq ([bool]$g3.CalPreguntar) $false '_CorreuGiaDecideix: nomes la capcalera, no cal preguntar'
# Informe antic sense ID GIA enlloc -> s'ha de PREGUNTAR, mai endevinar.
$g4 = _CorreuGiaDecideix '' ''
AssertEq $g4.Gia '' '_CorreuGiaDecideix: cap dels dos, res'
AssertEq ([bool]$g4.CalPreguntar) $true '_CorreuGiaDecideix: cap dels dos -> preguntar'
# Discrepancia -> tampoc s'endevina: es pregunta i es diu que ha passat.
$g5 = _CorreuGiaDecideix '1000' '1466'
AssertEq ([bool]$g5.CalPreguntar) $true '_CorreuGiaDecideix: si no coincideixen, preguntar'
AssertEq $g5.Gia '1466' '_CorreuGiaDecideix: proposa el de la capcalera (es el contingut de l''informe)'
AssertEq ([bool]($g5.Motiu -like '*1000*' -and $g5.Motiu -like '*1466*')) $true '_CorreuGiaDecideix: el motiu diu els DOS valors'

Write-Host "`n--- EnviarCorreu.ps1: la capcalera del correu surt del DOCUMENT ---"
$ecAct = @{ TITULAR = 'Bar del GIA 1466'; ADRECA = 'C/ Nou 3'; ACTIVITAT = 'BAR'
            EMAIL = 'rao@x.cat'; EMAIL_REP = 'rep@x.cat'; EXP_NUM = 'EXP-1466' }
# L'ultim informe generat era d'UNA ALTRA activitat: no pot colar-se res seu.
$ecRepAltre = @{ ID_GIA = '1000'; TITULAR = 'Bar del GIA 1000'; ADRECA = 'C/ Vell 1'
                 EMAIL = 'altre@x.cat'; NUM_ANOTACIO = '12345' }
$hA = _CorreuHeaderMerge '1466' $ecAct $ecRepAltre
AssertEq ([string]$hA['ID_GIA']) '1466' 'Capcalera: l''ID GIA es el del document'
AssertEq ([string]$hA['TITULAR']) 'Bar del GIA 1466' 'Capcalera: el titular surt de l''Excel'
AssertEq ([string]$hA['EMAIL']) 'rao@x.cat' 'Capcalera: el correu surt de l''Excel'
AssertEq ([bool]$hA.ContainsKey('NUM_ANOTACIO')) $false 'Capcalera: NO s''agafa res de l''informe d''una ALTRA activitat'
# Mateixa activitat: si que s'aprofita el que l'Excel no te.
$ecRepMateix = @{ ID_GIA = '1466'; NUM_ANOTACIO = '999'; DATA_ANOTACIO = '01/09/2026'; TITULAR = 'No em facis cas' }
$hB = _CorreuHeaderMerge '1466' $ecAct $ecRepMateix
AssertEq ([string]$hB['NUM_ANOTACIO']) '999' 'Capcalera: del MATEIX GIA si que s''agafa el que l''Excel no te'
AssertEq ([string]$hB['TITULAR']) 'Bar del GIA 1466' 'Capcalera: pero l''Excel mana sobre l''informe anterior'
# Sense fitxa a l'Excel: com a minim l'ID GIA ha de ser correcte.
$hC = _CorreuHeaderMerge '777' $null $ecRepAltre
AssertEq ([string]$hC['ID_GIA']) '777' 'Capcalera: sense Excel, l''ID GIA segueix sent el del document'
AssertEq ([string]$hC['TITULAR']) '' 'Capcalera: sense Excel i sense informe del mateix GIA, res inventat'

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
        (_RcAct '500' $Script:EstatFavorablePre '2026-01-05')
        (_RcAct '600' $Script:EstatFavorablePost '2026-01-05')
    )
}
$rcHist = @{ '100' = @{ ultim='2026-06-01'; compte=2; excloure=$false; enviaments=@('2026-06-01') } }
$rcRes = _RecDueActivitats $rcDb (_RecCampanyaPerClau 'requeriments') $rcCfg $rcHist $rcAvui
$rcFiles = @($rcRes.Files)
AssertEq $rcFiles.Count 4 '_RecDueActivitats: nomes les activitats en estat Requeriment (i el favorable pre-llicencia, que es tracta igual)'
Assert (@($rcFiles | Where-Object { $_.Id -eq '500' }).Count -eq 1 -and @($rcFiles | Where-Object { $_.Id -eq '600' }).Count -eq 0) '_RecDueActivitats: el pre-llicencia hi es i el post-llicencia no'
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
AssertEq ([string]$rcDef['mode']) 'manual' '_RecDefaultConfig: neix en MANUAL (no envia res sol fins que la passis a Automatic)'
Assert (-not $rcDef.ContainsKey('actiu')) '_RecDefaultConfig: ja no hi ha "Campanya activa" (el mode ho decideix tot)'
$rcN = _RecNormalitzaConfig ([pscustomobject]@{ actiu=$true; mode='auto'; periodicitatDies=90 }) 'requeriments'
AssertEq ([string](_RecNormalitzaConfig ([pscustomobject]@{ actiu=$false; mode='auto' }) 'requeriments')['mode']) 'manual' '_RecNormalitzaConfig: una campanya d''abans APAGADA i en Automatic passa a Manual (no comenca a enviar per l''actualitzacio)'
AssertEq ([string](_RecNormalitzaConfig ([pscustomobject]@{ mode='auto' }) 'requeriments')['mode']) 'auto' '_RecNormalitzaConfig: sense ''actiu'' (les d''ara), el mode es el que diu'
# La preseleccio: com a molt maxPerTanda, en l'ordre de prioritat.
$rcFiles = @(1..20 | ForEach-Object { [pscustomobject]@{ Id = [string]$_; Toca = ($_ -ne 3); Excloure = $false; Sel = $true } })
_RecPreselecciona $rcFiles 15
AssertEq (@($rcFiles | Where-Object { $_.Sel }).Count) 15 '_RecPreselecciona: 19 que toquen i maxim 15 -> se''n marquen 15 (abans totes)'
AssertEq "$($rcFiles[2].Sel)|$($rcFiles[15].Sel)|$($rcFiles[16].Sel)" 'False|True|False' '_RecPreselecciona: la que no toca no es marca, i les primeres per ordre ocupen la tanda'
_RecPreselecciona $rcFiles 50
AssertEq (@($rcFiles | Where-Object { $_.Sel }).Count) 19 '_RecPreselecciona: un maxim mes gran que les que toquen -> totes les que toquen'
foreach ($k in @('cada', 'espera', 'max', 'mode')) { Assert ((_RecAjudaCamp $k).Length -gt 40) ("_RecAjudaCamp '" + $k + "': hi ha l'explicacio de la (i)") }
AssertEq ([string]$rcN['mode']) 'auto' '_RecNormalitzaConfig: mode auto'
AssertEq ([int]$rcN['periodicitatDies']) 90 '_RecNormalitzaConfig: periodicitat del JSON'
AssertEq ([int]$rcN['maxPerTanda']) 15 '_RecNormalitzaConfig: el que no ve, del defecte'
$rcN2 = _RecNormalitzaConfig ([pscustomobject]@{ periodicitatDies=0; mode='tonteria'; cos='' }) 'requeriments'
AssertEq ([int]$rcN2['periodicitatDies']) 60 '_RecNormalitzaConfig: un valor absurd cau al defecte'
AssertEq ([string]$rcN2['mode']) 'manual' '_RecNormalitzaConfig: un mode desconegut cau a manual'
Assert ([string]$rcN2['cos'] -ne '') '_RecNormalitzaConfig: un cos buit no pot deixar el correu sense text'
$rcN3 = _RecNormalitzaConfig ([pscustomobject]@{ esperaInicialDies=0 }) 'requeriments'
AssertEq ([int]$rcN3['esperaInicialDies']) 0 '_RecNormalitzaConfig: espera inicial 0 SI que es valida'

# LA TASCA EN XML: a l'hora dels modes automatics i, si es salta (PC apagat),
# en quant es pugui (StartWhenAvailable). Les rutes amb espais, entre cometes.
$rcXml = _RecTascaXml 'C:\Win\powershell.exe' 'I:\5.- Sergi Fadurdo\suport\RecordatorisAuto.ps1' (Get-AutoHoraText) ([datetime]'2026-10-05')
$rcDoc = New-Object System.Xml.XmlDocument
$rcDoc.LoadXml(($rcXml -replace '^<\?xml[^>]*\?>', ''))
$rcNs = New-Object System.Xml.XmlNamespaceManager($rcDoc.NameTable); $rcNs.AddNamespace('t', 'http://schemas.microsoft.com/windows/2004/02/mit/task')
AssertEq ([string]$rcDoc.SelectSingleNode('//t:StartBoundary', $rcNs).InnerText) '2026-10-05T13:00:00' '_RecTascaXml: cada dia a les 13:00 (l''hora dels modes automatics)'
AssertEq ([string]$rcDoc.SelectSingleNode('//t:StartWhenAvailable', $rcNs).InnerText) 'true' '_RecTascaXml: si es salta, es fa en quant es pugui'
AssertEq ([string]$rcDoc.SelectSingleNode('//t:DaysInterval', $rcNs).InnerText) '1' '_RecTascaXml: diaria'
AssertEq ([string]$rcDoc.SelectSingleNode('//t:Command', $rcNs).InnerText) 'C:\Win\powershell.exe' '_RecTascaXml: l''executable'
Assert (([string]$rcDoc.SelectSingleNode('//t:Arguments', $rcNs).InnerText).Contains('-File "I:\5.- Sergi Fadurdo\suport\RecordatorisAuto.ps1"')) '_RecTascaXml: el script (amb espais) va entre cometes'
$rcXmlAmp = _RecTascaXml 'C:\p.exe' 'C:\R&D\s.ps1' '13:00' ([datetime]'2026-10-05')
Assert ($rcXmlAmp.Contains('R&amp;D')) '_RecTascaXml: els caracters especials de la ruta s''escapen'
$rcArgv = @(_RecSchtasksArgv 'InformesCornella-Recordatoris' 'C:\t.xml')
AssertEq ($rcArgv -join ' ') '/Create /TN InformesCornella-Recordatoris /XML C:\t.xml /F' '_RecSchtasksArgv: crea (o reescriu) la tasca des de l''XML'
Assert (_RecTascaAlDia $rcXml '13:00') '_RecTascaAlDia: la d''ara esta al dia'
$rcVella = '<Task><Triggers><CalendarTrigger><StartBoundary>2026-09-01T09:00:00</StartBoundary></CalendarTrigger></Triggers><Settings><StartWhenAvailable>false</StartWhenAvailable></Settings></Task>'
Assert (-not (_RecTascaAlDia $rcVella '13:00')) '_RecTascaAlDia: la de les 09:00 sense recuperar-se s''ha de refer'
# UN COP A LA SETMANA (la programacio de 'recordatoris', Configuracio).
$rcSet = _RecTascaXml 'C:\p.exe' 'C:\s.ps1' '08:30' ([datetime]'2026-10-05') 3
[xml]$rcDocS = $rcSet
$rcNsS = New-Object System.Xml.XmlNamespaceManager($rcDocS.NameTable); $rcNsS.AddNamespace('t', 'http://schemas.microsoft.com/windows/2004/02/mit/task')
AssertEq "$([string]$rcDocS.SelectSingleNode('//t:StartBoundary', $rcNsS).InnerText)|$($null -ne $rcDocS.SelectSingleNode('//t:ScheduleByWeek/t:DaysOfWeek/t:Wednesday', $rcNsS))|$($null -eq $rcDocS.SelectSingleNode('//t:ScheduleByDay', $rcNsS))" '2026-10-05T08:30:00|True|True' '_RecTascaXml setmanal: els dimecres a les 08:30'
Assert (_RecTascaAlDia $rcSet '08:30' 3) '_RecTascaAlDia: la setmanal del dimecres, al dia'
Assert (-not (_RecTascaAlDia $rcSet '08:30' 0)) '_RecTascaAlDia: si ara toca cada dia, la setmanal s''ha de refer'
Assert (-not (_RecTascaAlDia $rcSet '08:30' 4)) '_RecTascaAlDia: si ara toca els dijous, tambe'
Assert (-not (_RecTascaAlDia $rcXml '13:00' 1)) '_RecTascaAlDia: la diaria, si ara toca setmanal, s''ha de refer'
Assert (-not (_RecTascaAlDia ($rcXml -replace 'T13:00', 'T09:00') '13:00')) '_RecTascaAlDia: una altra hora, tambe'
Assert (-not (_RecTascaAlDia '' '13:00')) '_RecTascaAlDia: sense tasca, no'
Assert (_RecTascaAlDia (($rcXml.ToCharArray() | ForEach-Object { [string]$_ + [char]0 }) -join '') '13:00') '_RecTascaAlDia: tambe si el schtasks la torna en UTF-16 llegida a 8 bits'

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

Write-Host "`n--- InformesClassificacio.ps1: _EstatActualActivitat / _InformeQueDeterminaEstat (quin informe decideix l'estat) ---"
$mkAct = { param($gia, $infs) [pscustomobject]@{ id_gia = $gia; informes = @($infs) } }
$mkInf = { param($data, $breu, $tipus = '', $ign = $false, $fitxer = 'f.docx') [pscustomobject]@{ data = $data; fitxer = $fitxer; ignorat = $ign; conclusio_breu = $breu; tipus = $tipus } }
AssertEq (_EstatActualActivitat $null) '' '_EstatActualActivitat null -> buit'
AssertEq (_EstatActualActivitat (& $mkAct '1' @())) '' '_EstatActualActivitat sense informes -> buit'
$actTotIgn = & $mkAct '1' @((& $mkInf '2026-01-01' 'Requeriment' '' $true), (& $mkInf '2026-02-01' 'Favorable' '' $true))
AssertEq (_EstatActualActivitat $actTotIgn) '' '_EstatActualActivitat tots ignorats -> buit'
$actNormal = & $mkAct '1' @((& $mkInf '2026-01-01' 'Requeriment'), (& $mkInf '2026-02-01' 'Favorable' '' $true), (& $mkInf '2026-03-01' 'FI Requeriment'))
AssertEq (_EstatActualActivitat $actNormal) 'FI Requeriment' '_EstatActualActivitat es queda amb el darrer NO ignorat (per data)'
# QUIN informe decideix l'estat: el mateix que fa servir _EstatActualActivitat.
# Ho necessita "Comprovar Excel" per dir-ne la data (INFORME ENGINYER dd/MM/aaaa).
AssertEq ($null -eq (_InformeQueDeterminaEstat $null)) $true '_InformeQueDeterminaEstat null -> $null'
AssertEq ($null -eq (_InformeQueDeterminaEstat $actTotIgn)) $true '_InformeQueDeterminaEstat tots ignorats -> $null'
AssertEq (_InformeQueDeterminaEstat $actNormal).data '2026-03-01' '_InformeQueDeterminaEstat: el darrer NO ignorat (SALTA el del 02, ignorat)'
AssertEq (_InformeQueDeterminaEstat $actNormal).conclusio_breu (_EstatActualActivitat $actNormal) '_InformeQueDeterminaEstat i _EstatActualActivitat parlen del MATEIX informe'
# Demana l'ACTIVITAT: una llista d'informes (la signatura d'abans) PETA en lloc
# de decidir sense saber el GIA.
$petaLlista = $false
try { [void](_InformeQueDeterminaEstat @((& $mkInf '2026-01-01' 'Requeriment'))) } catch { $petaLlista = $true }
Assert $petaLlista '_InformeQueDeterminaEstat amb una LLISTA d''informes peta (espera l''activitat)'
# L'ordre no depen de com venen: es torna a ordenar per data.
$actDesord = & $mkAct '1' @((& $mkInf '2026-03-01' 'FI Requeriment'), (& $mkInf '2026-01-01' 'Requeriment'))
AssertEq (_EstatActualActivitat $actDesord) 'FI Requeriment' 'els informes desordenats es tornen a ordenar per data'
# Dos informes del MATEIX dia: desempata el nom del fitxer (el Sort-Object del
# PowerShell 5.1 no es estable i l'estat podia canviar d'una passada a l'altra).
$actMateixDia = & $mkAct '1' @((& $mkInf '2026-03-01' 'FI Requeriment' '' $false 'b.docx'), (& $mkInf '2026-03-01' 'Requeriment' '' $false 'a.docx'))
AssertEq (_EstatActualActivitat $actMateixDia) 'FI Requeriment' 'mateixa data: desempata el nom del fitxer (b despres d''a)'
$actMateixDia2 = & $mkAct '1' @((& $mkInf '2026-03-01' 'Requeriment' '' $false 'b.docx'), (& $mkInf '2026-03-01' 'FI Requeriment' '' $false 'a.docx'))
AssertEq (_EstatActualActivitat $actMateixDia2) 'Requeriment' 'mateixa data: ...sigui quin sigui l''ordre d''entrada'
# Favorable de LLICENCIA: decideix sempre (abans s'ignorava per defecte i
# l'activitat es quedava en Requeriment per un informe de dos anys abans).
$actLlic = & $mkAct '1' @((& $mkInf '2024-01-01' 'Requeriment'), (& $mkInf '2026-01-01' 'Favorable' 'llicfav'))
AssertEq (_EstatActualActivitat $actLlic) 'Favorable' 'favorable de llicencia (llicfav) -> fixa Favorable'
# MNS favorable: neutre.
$actMnsReq = & $mkAct '1' @((& $mkInf '2026-01-01' 'Requeriment'), (& $mkInf '2026-02-01' 'Favorable' 'mns'))
AssertEq (_EstatActualActivitat $actMnsReq) 'Requeriment' 'MNS favorable NO tapa un Requeriment pendent'
AssertEq (_InformeQueDeterminaEstat $actMnsReq).data '2026-01-01' '...i la data que mana (Recordatoris) es la del requeriment'
$actMnsPrec = & $mkAct '1' @((& $mkInf '2026-01-01' 'Precinte / Cessament'), (& $mkInf '2026-02-01' 'Favorable' 'mns'))
AssertEq (_EstatActualActivitat $actMnsPrec) 'Precinte / Cessament' 'MNS favorable NO tapa un Precinte / Cessament'
$actMnsAmp = & $mkAct '1' @((& $mkInf '2026-01-01' ('Ampliaci' + [char]0x00F3 + ' termini')), (& $mkInf '2026-02-01' 'Favorable' 'mns'))
AssertEq (_EstatActualActivitat $actMnsAmp) ('Ampliaci' + [char]0x00F3 + ' termini') 'MNS favorable NO tapa una Ampliacio termini'
$actMnsSol = & $mkAct '1' @((& $mkInf '2026-01-01' 'Favorable' 'mns'), (& $mkInf '2026-05-01' 'Favorable' 'mns'))
AssertEq (_EstatActualActivitat $actMnsSol) 'Favorable' 'nomes MNS favorables -> Favorable (abans, l''estat buit)'
$actMnsFi = & $mkAct '1' @((& $mkInf '2026-01-01' 'FI Requeriment'), (& $mkInf '2026-02-01' 'Favorable' 'mns'))
AssertEq (_EstatActualActivitat $actMnsFi) 'Favorable' 'res pendent + MNS favorable -> Favorable'
$actMnsDespres = & $mkAct '1' @((& $mkInf '2026-01-01' 'Favorable' 'mns'), (& $mkInf '2026-02-01' 'Requeriment'))
AssertEq (_EstatActualActivitat $actMnsDespres) 'Requeriment' 'un requeriment despres d''una MNS mana'
$actMnsReqTxt = & $mkAct '1' @((& $mkInf '2026-01-01' 'FI Requeriment'), (& $mkInf '2026-02-01' 'Requeriment' 'mns'))
AssertEq (_EstatActualActivitat $actMnsReqTxt) 'Requeriment' 'una MNS que REQUEREIX no es neutra: decideix'
# Activitat extraordinaria: sota el GIA d'un establiment, ni el favorable ni el requeriment.
$actEstadi = & $mkAct '9028' @((& $mkInf '2025-01-01' 'FI Requeriment'), (& $mkInf '2026-01-01' 'Requeriment' 'actextr'), (& $mkInf '2026-02-01' 'Favorable' 'actextr'))
AssertEq (_EstatActualActivitat $actEstadi) 'FI Requeriment' 'actextr sota el GIA d''un establiment no en decideix l''estat (ni el requeriment)'
$actEstadiSol = & $mkAct '9028' @((& $mkInf '2026-01-01' 'Requeriment' 'actextr'))
AssertEq (_EstatActualActivitat $actEstadiSol) '' '...ni quan nomes te actes'
$actActe = & $mkAct '' @((& $mkInf '2026-01-01' 'Requeriment' 'actextr'), (& $mkInf '2026-02-01' 'Favorable' 'actextr'))
AssertEq (_EstatActualActivitat $actActe) 'Favorable' 'actextr SENSE GIA (per carpeta): l''activitat es l''acte, i decideix'
# Informatius (Altres): nomes si no hi ha res mes.
$actAltres = & $mkAct '1' @((& $mkInf '2026-01-01' 'Requeriment'), (& $mkInf '2026-02-01' 'Altres'))
AssertEq (_EstatActualActivitat $actAltres) 'Requeriment' 'un informe Altres no decideix si n''hi ha cap altre'
$actNomesAltres = & $mkAct '1' @((& $mkInf '2026-02-01' 'Altres'))
AssertEq (_EstatActualActivitat $actNomesAltres) 'Altres' 'si l''unic es Altres, l''estat es Altres'
# El que l'usuari ha corregit a ma MANA per damunt dels tipus.
$infEd = & $mkInf '2026-01-01' 'Requeriment' 'actextr'
Add-Member -InputObject $infEd -NotePropertyName editat_a_ma -NotePropertyValue $true
AssertEq (_EstatActualActivitat (& $mkAct '9028' @((& $mkInf '2025-01-01' 'FI Requeriment'), $infEd))) 'Requeriment' 'un actextr corregit a ma sota un GIA SI que decideix (l''usuari mana)'
$infEdMns = & $mkInf '2026-02-01' 'Favorable' 'mns'
Add-Member -InputObject $infEdMns -NotePropertyName editat_a_ma -NotePropertyValue $true
AssertEq (_EstatActualActivitat (& $mkAct '1' @((& $mkInf '2026-01-01' 'Requeriment'), $infEdMns))) 'Favorable' 'una MNS favorable corregida a ma SI que tapa el requeriment'
$infIgnLlic = & $mkInf '2026-02-01' 'Favorable' 'llicfav' $true
AssertEq (_EstatActualActivitat (& $mkAct '1' @((& $mkInf '2026-01-01' 'Requeriment'), $infIgnLlic))) 'Requeriment' 'l''ignorat a ma mana, fins i tot sobre un favorable de llicencia'

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

Write-Host "`n--- Menu.ps1: el ""?"" d'ajuda de cada rajola d'eina ---"
# Les accions de les rajoles es llegeixen del codi del menu (les linies
# "@{ Emoji = ...; Action = '...' }"): una rajola nova sense text d'ajuda fa
# petar aquesta prova en lloc de sortir amb el "?" sense res.
$srcMenuAj = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $TestsDir) 'Menu.ps1') -Raw
$accionsRajola = @([regex]::Matches($srcMenuAj, "(?m)^\s*@\{ Emoji = .*?Action = '(\w+)'") | ForEach-Object { $_.Groups[1].Value })
Assert ($accionsRajola.Count -ge 14) ('ajuda eines: es troben les rajoles del menu (' + $accionsRajola.Count + ')')
$senseAjuda = @($accionsRajola | Where-Object { [string]::IsNullOrWhiteSpace((_AjudaEina $_)) })
AssertEq ($senseAjuda -join ', ') '' 'ajuda eines: CADA rajola te el seu text'
$sobrants = @($Script:AjudaEines.Keys | Where-Object { $_ -notin $accionsRajola })
AssertEq ($sobrants -join ', ') '' 'ajuda eines: cap text d''una eina que ja no hi es'
$llargues = @($accionsRajola | Where-Object { (_AjudaEina $_).Length -gt 220 })
AssertEq ($llargues -join ', ') '' 'ajuda eines: textos BREUS (fins a 220 caracters)'
AssertEq (_AjudaEina 'einaqueno') '' '_AjudaEina: una accio desconeguda -> buit'
AssertEq (_AjudaEina '') '' '_AjudaEina: buit -> buit'

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

Write-Host "`n--- Informes.ps1: les edicions a ma ('Editar base') prevalen ---"
# Una correccio a ma es marca i guarda el que deia l'automatic.
$infM = [pscustomobject]@{ conclusio = 'Vist l''anterior, es pot donar per finalitzat el tramit.'; conclusio_breu = 'FI Requeriment'; ignorat = $false }
_MarcaEditatAMa $infM
$infM.conclusio_breu = 'Requeriment'
AssertEq "$($infM.editat_a_ma)|$($infM.auto_conclusio_breu)|$($infM.auto_ignorat)" 'True|FI Requeriment|False' 'marca: editat, amb el valor automatic guardat ABANS del canvi'
$infM.ignorat = $true
_MarcaEditatAMa $infM
AssertEq $infM.auto_conclusio_breu 'FI Requeriment' 'una segona correccio no trepitja el valor automatic guardat'
Assert (_DesfesEditatAMa $infM) 'desfer: torna $true si hi havia canvi'
AssertEq "$($infM.conclusio_breu)|$($infM.ignorat)|$($infM.editat_a_ma)|$($null -eq $infM.PSObject.Properties['auto_conclusio_breu'])" 'FI Requeriment|False|False|True' 'desfer: torna al que deia l informe i treu els auto_*'
Assert (-not (_DesfesEditatAMa $infM)) 'desfer un que no esta editat: no fa res'
# Una base d'ABANS de la marca: una conclusio breu que no surt del text nomes pot ser a ma.
$infV = [pscustomobject]@{ ruta = 'I:\x\v.docx'; conclusio = 'Vist l''anterior, es pot donar per finalitzat el tramit.'; conclusio_breu = 'Precinte / Cessament'; ignorat = $false }
_InferEditatAMa $infV
AssertEq "$($infV.editat_a_ma)|$($infV.auto_conclusio_breu)" 'True|FI Requeriment' 'base antiga: conclusio breu que no surt del text -> editada a ma'
$infV2 = [pscustomobject]@{ conclusio = 'Vist l''anterior, es pot donar per finalitzat el tramit.'; conclusio_breu = 'FI Requeriment'; ignorat = $false }
_InferEditatAMa $infV2
AssertEq $infV2.editat_a_ma $false 'base antiga: la que coincideix amb l automatic, no'
AssertEq (_ActivitatEditadaAMa ([pscustomobject]@{ informes = @($infV2, $infV) })) $true 'activitat editada si ho es algun dels seus informes'
AssertEq (_ActivitatEditadaAMa ([pscustomobject]@{ informes = @($infV2) })) $false 'i no, si cap'
# L'ESCANEIG: el de l'usuari preval; si no l'ha tocat, mana l'informe nou.
function _NouReg($cb, $ign) { [pscustomobject]@{ ConclusioBreu = $cb; Ignorat = $ign; EditatAMa = $false; AutoConclusioBreu = ''; AutoIgnorat = $false } }
$prevE = [pscustomobject]@{ ConclusioBreu = 'Requeriment'; Ignorat = $true; EditatAMa = $true; TeMarcaEdicio = $true }
$r1 = _AplicaEdicioPrevia $prevE (_NouReg 'FI Requeriment' $false)
AssertEq "$($r1.ConclusioBreu)|$($r1.Ignorat)|$($r1.EditatAMa)|$($r1.AutoConclusioBreu)" 'Requeriment|True|True|FI Requeriment' 'escaneig: la correccio a ma preval, i es guarda el nou automatic'
$prevN = [pscustomobject]@{ ConclusioBreu = 'Requeriment'; Ignorat = $false; EditatAMa = $false; TeMarcaEdicio = $true }
$r2 = _AplicaEdicioPrevia $prevN (_NouReg 'FI Requeriment' $false)
AssertEq "$($r2.ConclusioBreu)|$($r2.EditatAMa)" 'FI Requeriment|False' 'escaneig: sense correccio a ma, mana l informe nou (abans es congelava)'
$prevL = [pscustomobject]@{ ConclusioBreu = 'Requeriment'; Ignorat = $true; EditatAMa = $false; TeMarcaEdicio = $false }
$r3 = _AplicaEdicioPrevia $prevL (_NouReg 'Requeriment' $false)
AssertEq "$($r3.Ignorat)|$($r3.EditatAMa)" 'True|True' 'escaneig, base antiga: un ignorat que no es el per defecte l havia posat l usuari'
$flatM = _FlattenInformesDb ([pscustomobject]@{ activitats = @([pscustomobject]@{ id_gia = '7'; informes = @($infV) }) })
$recM = @($flatM.Values)[0]
AssertEq "$($recM.EditatAMa)|$($recM.TeMarcaEdicio)" 'True|True' '_FlattenInformesDb porta la marca d edicio'
# I es desa: els auto_* nomes si esta editat.
$jE = _InformeAJson ([pscustomobject]@{ Data = 'd'; Fitxer = 'f'; Ruta = 'r'; Conclusio = 'c'; ConclusioBreu = 'Requeriment'; Modificat = 'm'; Ignorat = $false; Motius = @('a', 'b'); EditatAMa = $true; AutoConclusioBreu = 'FI Requeriment'; AutoIgnorat = $false })
AssertEq "$($jE.editat_a_ma)|$($jE.auto_conclusio_breu)|$($jE.motiu)" 'True|FI Requeriment|a, b' '_InformeAJson: amb la marca i el valor automatic'
$jN = _InformeAJson ([pscustomobject]@{ Data = 'd'; Fitxer = 'f'; Ruta = 'r'; Conclusio = 'c'; ConclusioBreu = 'X'; Modificat = 'm'; Ignorat = $false; Motius = @(); EditatAMa = $false; AutoConclusioBreu = ''; AutoIgnorat = $false })
AssertEq "$($jN.editat_a_ma)|$($null -eq $jN.PSObject.Properties['auto_conclusio_breu'])" 'False|True' '_InformeAJson: sense edicio, sense camps auto_*'

Write-Host "`n--- Informes.ps1: _OrdenaFilesBase (ID GIA numeric, la columna clicada mana) ---"
$filesO = @(
    [pscustomobject]@{ Gia = '10'; Carpeta = ''; Data = '2026-01-02'; EstatActual = 'Requeriment' }
    [pscustomobject]@{ Gia = '1000'; Carpeta = ''; Data = '2026-01-01'; EstatActual = 'Favorable' }
    [pscustomobject]@{ Gia = '103'; Carpeta = ''; Data = '2026-01-03'; EstatActual = 'Requeriment' }
    [pscustomobject]@{ Gia = ''; Carpeta = 'B'; Data = '2026-01-01'; EstatActual = 'Favorable' }
    [pscustomobject]@{ Gia = '9'; Carpeta = ''; Data = '2026-02-01'; EstatActual = 'Favorable' }
    [pscustomobject]@{ Gia = '10'; Carpeta = ''; Data = '2025-12-01'; EstatActual = 'Requeriment' }
)
AssertEq (@(_OrdenaFilesBase $filesO @{} -1 $true | ForEach-Object { "$($_.Gia)$($_.Carpeta)" }) -join ',') '9,10,10,103,1000,B' 'sense columna: per ID GIA NUMERIC (9, 10, 103, 1000), els sense GIA al final'
AssertEq (@(_OrdenaFilesBase $filesO @{} -1 $true | Where-Object { $_.Gia -eq '10' } | ForEach-Object { $_.Data }) -join ',') '2025-12-01,2026-01-02' 'dins d una activitat, per data'
$exprO = @{ 6 = { [string]$_.EstatActual }; 1 = { _GiaNumeric $_.Gia } }
AssertEq (@(_OrdenaFilesBase $filesO $exprO 6 $true | ForEach-Object { "$($_.EstatActual):$($_.Gia)$($_.Carpeta)" }) -join ',') 'Favorable:9,Favorable:1000,Favorable:B,Requeriment:10,Requeriment:10,Requeriment:103' 'clic a Estat: l estat MANA i el GIA desempata (abans nomes ordenava dins de cada activitat)'
AssertEq (@(_OrdenaFilesBase $filesO $exprO 1 $false | ForEach-Object { $_.Gia }) -join ',') ',1000,103,10,10,9' 'clic a GIA descendent: numeric (els sense GIA, que valen el maxim, primer)'

Write-Host "`n--- ModeAutomatic.ps1: el comu dels modes automatics ---"
# Una altra hora que la de la copia: la regla es la mateixa.
Assert (-not (_AutoToca ([datetime]'2026-09-08T13:59:00') ([datetime]'2026-09-07T14:00:05').ToString('o') 14 0)) '_AutoToca 14:00: a les 13:59 encara no'
Assert (_AutoToca ([datetime]'2026-09-08T14:00:00') ([datetime]'2026-09-07T14:00:05').ToString('o') 14 0) '_AutoToca 14:00: a les 14:00 en punt, si'
# L'HORA DE TOTS ELS MODES AUTOMATICS: les 13:00 (l'usuari, octubre 2026).
AssertEq (Get-AutoHoraText) '13:00' 'modes automatics: tots a les 13:00'
AssertEq (_BaseAutoToca ([datetime]'2026-09-08T12:59:00') ([datetime]'2026-09-07T13:00:05').ToString('o')) $false 'base auto: a les 12:59, encara no'
AssertEq (_BaseAutoToca ([datetime]'2026-09-08T13:00:00') ([datetime]'2026-09-07T13:00:05').ToString('o')) $true 'base auto: a les 13:00, si'
Assert ((Get-AutoTipText 'es copia sol').Contains('13:00') -and (Get-AutoTipText 'x').Contains("l'ultima vegada que tocava")) 'l''ajuda de l''interruptor diu l''hora i que es recupera l''ultima passada perduda'
foreach ($kH in @('informesdb', 'copiarinformes')) { Assert ((_AjudaEina $kH).Contains('13:00') -and -not (_AjudaEina $kH).Contains('[HORA_AUTO]')) ("ajuda '" + $kH + "': diu les 13:00") }
# Esperar que la base s'acabi d'escriure (els Recordatoris, abans de llegir-la).
AssertEq (Wait-MutexLliure ('Global\InformesCornella.Prova.' + [guid]::NewGuid().ToString('N')) 0) $true 'Wait-MutexLliure: lliure, de seguida'
AssertEq (_AutoVenciment ([datetime]'2026-09-08T08:00:00') 7 15) ([datetime]'2026-09-08T07:15:00') '_AutoVenciment: amb els minuts de l''eina'

Write-Host "`n--- ModeAutomatic.ps1: la PROGRAMACIO de cada automatisme (setmanal i configurable) ---"
# Un cop a la setmana (el Planol activitats, l'usuari, octubre 2026). 2026-10-05 es DILLUNS.
AssertEq (_AutoVenciment ([datetime]'2026-10-07T09:00:00') 13 0 1) ([datetime]'2026-10-05T13:00:00') 'setmanal: un dimecres, el venciment es el dilluns passat'
AssertEq (_AutoVenciment ([datetime]'2026-10-05T12:59:00') 13 0 1) ([datetime]'2026-09-28T13:00:00') 'setmanal: el dilluns abans de l hora, el de la setmana passada'
AssertEq (_AutoVenciment ([datetime]'2026-10-05T13:00:00') 13 0 1) ([datetime]'2026-10-05T13:00:00') 'setmanal: el dilluns a l hora, el d avui'
AssertEq (_AutoVenciment ([datetime]'2026-10-11T23:00:00') 13 0 7) ([datetime]'2026-10-11T13:00:00') 'setmanal: el diumenge (7) tambe'
Assert (-not (_AutoToca ([datetime]'2026-10-09T10:00:00') ([datetime]'2026-10-05T13:00:10').ToString('o') 13 0 1)) 'setmanal: fet dilluns, el divendres ja no toca'
Assert (_AutoToca ([datetime]'2026-10-13T08:00:00') ([datetime]'2026-10-02T13:00:10').ToString('o') 13 0 1) 'setmanal: el dilluns es va perdre (PC apagat): toca en obrir, encara que sigui dimarts'
$setPr = [pscustomobject]@{ Automatismes = [pscustomobject]@{
    copiarinformes = [pscustomobject]@{ Freq = 'setmana'; Dia = 3; Hora = '9:05' }
    informesdb     = [pscustomobject]@{ Freq = 'cada hora'; Dia = 1; Hora = '13:00' } } }
$prC = Get-ProgramacioAuto 'copiarinformes' $setPr
AssertEq "$($prC.Freq)|$($prC.Dia)|$($prC.Hora)|$($prC.Propia)" 'setmana|3|09:05|True' 'la programacio d aquest PC (settings.json) mana, amb l hora normalitzada'
AssertEq (Get-ProgramacioText $prC) 'cada dimecres a les 09:05' 'i es diu en catala'
AssertEq "$((Get-ProgramacioAuto 'informesdb' $setPr).Freq)|$((Get-ProgramacioAuto 'informesdb' $setPr).Propia)" 'dia|False' 'una programacio que no es valida (feta a ma) no compta: la per defecte'
AssertEq (Get-ProgramacioText (Get-ProgramacioAuto 'planol' ([pscustomobject]@{}))) 'cada dilluns a les 13:00' 'el Planol, per defecte: cada dilluns a les 13:00'
Assert (Test-ProgramacioToca 'copiarinformes' ([datetime]'2026-10-07T09:06:00') ([datetime]'2026-10-06T13:00:00').ToString('o') $setPr) 'amb la programacio propia: el dimecres a les 09:06 toca'
Assert (-not (Test-ProgramacioToca 'copiarinformes' ([datetime]'2026-10-07T09:04:00') ([datetime]'2026-10-06T13:00:00').ToString('o') $setPr)) '...a les 09:04, encara no'
foreach ($pv in @(@{ Freq = 'dia'; Dia = 1; Hora = '24:00' }, @{ Freq = 'setmana'; Dia = 8; Hora = '10:00' }, @{ Freq = 'mes'; Dia = 1; Hora = '10:00' }, $null)) {
    Assert (-not (Test-ProgramacioValida $pv)) ("programacio no valida: " + ($pv | ConvertTo-Json -Compress))
}
$aDesar = ConvertTo-AutomatismesSettings @{
    copiarinformes = @{ Freq = 'dia'; Dia = 4; Hora = '13:00' }      # igual que per defecte (el dia no compta)
    planol         = @{ Freq = 'setmana'; Dia = 5; Hora = '08:30' }  # diferent
    informesdb     = @{ Freq = 'dia'; Dia = 1; Hora = '99:00' } }    # no valida
AssertEq (@($aDesar.Keys) -join ',') 'planol' 'a settings.json nomes hi va el que difereix del per defecte (i es valid)'
Assert ((Get-AutoTipText 'x' 'planol').Contains('cada dilluns a les 13:00') -and (Get-AutoTipText 'x' 'planol').Contains('Configuracio')) 'l ajuda de l interruptor diu la programacio de l eina i que es canvia a Configuracio'
AssertEq (_AjudaEina 'planol').Contains('cada dilluns a les 13:00') $true 'l ajuda de la rajola del Planol diu quan es fa sol'
$tmpMA = Join-Path ([System.IO.Path]::GetTempPath()) ('mode-auto-' + [guid]::NewGuid().ToString('N'))
try {
    $plMA = [ordered]@{ auto = $false; auto_el = ''; mode = '' }
    $pMA = Join-Path (Join-Path $tmpMA 'sub') 'estat.json'
    $e0MA = Read-EstatAuto $pMA $plMA
    AssertEq "$($e0MA['auto'])|$($e0MA['auto_el'])|$(@($e0MA.Keys).Count)" 'False||3' 'Read-EstatAuto: sense fitxer, la plantilla sencera'
    Assert (Save-EstatAuto $pMA @{ auto = $true } $plMA) 'Save-EstatAuto: crea la carpeta i desa'
    [void](Save-EstatAuto $pMA @{ mode = 'auto' } $plMA)
    $e1MA = Read-EstatAuto $pMA $plMA
    AssertEq "$($e1MA['auto'])|$($e1MA['mode'])" 'True|auto' 'Save-EstatAuto: nomes les claus que es donen (l''interruptor no s''ha perdut)'
    Assert ($e1MA['auto'] -is [bool]) 'Read-EstatAuto: les claus [bool] de la plantilla es llegeixen com a bool'
    Assert ($null -eq (Read-JsonFile $pMA).PSObject.Properties['generat_el']) 'Save-EstatAuto: no s''inventa generat_el si la plantilla no en te'
    Set-Content -LiteralPath $pMA -Value '{ aixo no es json'
    AssertEq ([bool](Read-EstatAuto $pMA $plMA)['auto']) $false 'Read-EstatAuto: un fitxer corrupte val com si no n''hi hagues'
    AssertEq (Save-EstatAuto '' @{ auto = $true } $plMA) $false 'Save-EstatAuto: sense ruta, $false (mai llanca)'
} finally { Remove-Item -LiteralPath $tmpMA -Recurse -Force -ErrorAction SilentlyContinue }
$fetMA = @{ N = 0 }
AssertEq (Invoke-AmbMutexUnic ('Global\InformesCornella.Prova.' + [guid]::NewGuid().ToString('N')) { $fetMA.N++ }) $true 'Invoke-AmbMutexUnic: lliure, fa la feina i torna $true'
AssertEq $fetMA.N 1 'Invoke-AmbMutexUnic: ...una sola vegada'
$errMA = ''
try { [void](Invoke-AmbMutexUnic ('Global\InformesCornella.Prova.' + [guid]::NewGuid().ToString('N')) { throw 'peta' }) } catch { $errMA = [string]$_.Exception.Message }
AssertEq $errMA 'peta' 'Invoke-AmbMutexUnic: l''error de la feina es propaga'

# Un ALTRE PROCES que reté un mutex amb nom fins que se li diu: dins del mateix
# fil el mutex es recursiu i no es podria veure mai "ocupat".
$retenirMutex = {
    param([string]$nom, [string]$dir)
    $ok = Join-Path $dir 'agafat'; $prou = Join-Path $dir 'prou'
    $cmd = "`$m = New-Object System.Threading.Mutex(`$false, '$nom'); [void]`$m.WaitOne(); Set-Content -LiteralPath '$ok' -Value 1; " +
           "for (`$k = 0; `$k -lt 600 -and -not (Test-Path -LiteralPath '$prou'); `$k++) { Start-Sleep -Milliseconds 100 }; `$m.ReleaseMutex()"
    $enc = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($cmd))
    $spl = @{ FilePath = (Get-Process -Id $PID).Path; ArgumentList = @('-NoProfile', '-NonInteractive', '-EncodedCommand', $enc); PassThru = $true }
    if ($PSVersionTable.PSEdition -eq 'Desktop' -or $IsWindows) { $spl.WindowStyle = 'Hidden' }
    $p = Start-Process @spl
    for ($k = 0; $k -lt 300 -and -not (Test-Path -LiteralPath $ok); $k++) { Start-Sleep -Milliseconds 100 }
    return @{ Proc = $p; Prou = $prou; Agafat = (Test-Path -LiteralPath $ok) }
}
$deixarMutex = {
    param($h)
    Set-Content -LiteralPath $h.Prou -Value 1
    try { [void]$h.Proc.WaitForExit(20000) } catch { }
}

Write-Host "`n--- InformesEscaneig.ps1: Actualitzar base en AUTOMATIC (d'extrem a extrem) ---"
# .docx de veritat (un zip amb word/document.xml): l'escaneig els obre igual
# que els de la feina.
Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
$nouDocx = {
    param([string]$path, [string[]]$paras)
    $cos = ($paras | ForEach-Object { '<w:p><w:r><w:t xml:space="preserve">' + [System.Security.SecurityElement]::Escape($_) + '</w:t></w:r></w:p>' }) -join ''
    $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + $cos + '</w:body></w:document>'
    $fs = [System.IO.File]::Create($path)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        $sw = New-Object System.IO.StreamWriter($zip.CreateEntry('word/document.xml').Open(), (New-Object System.Text.UTF8Encoding($false)))
        $sw.Write($xml); $sw.Dispose(); $zip.Dispose()
    } finally { $fs.Dispose() }
}
$ba = Join-Path ([System.IO.Path]::GetTempPath()) ('base-auto-' + [guid]::NewGuid().ToString('N'))
$baInf = Join-Path $ba 'informes'; $baLoc = Join-Path $ba 'local'; $baAct = Join-Path $ba 'activitats'; $baApp = Join-Path $ba 'appdata'
foreach ($d in @((Join-Path $baInf 'GIA 361'), (Join-Path $baInf 'GIA 362'), $baLoc, $baAct, $baApp)) { [void](New-Item -ItemType Directory -Path $d -Force) }
$vellsBA = @{ Inf = $InformesDir; Loc = $LocalActivitatsDir; Act = $ActivitatsDir; App = $env:LOCALAPPDATA }
$InformesDir = $baInf; $LocalActivitatsDir = $baLoc; $ActivitatsDir = $baAct; $env:LOCALAPPDATA = $baApp
$req = "Vist l'anterior, no es pot donar per finalitzat el tr" + [char]0x00E0 + "mit."
$fi  = "Vist l'anterior, es pot donar per finalitzat el tr" + [char]0x00E0 + "mit."
$doc361 = Join-Path (Join-Path $baInf 'GIA 361') '2026-09-01_Req_GIA_361.docx'
$doc362 = Join-Path (Join-Path $baInf 'GIA 362') '2026-09-02_Seg_GIA_362.docx'
$dbBA = Join-Path $baLoc 'informes-db.json'
try {
    & $nouDocx $doc361 @('ID GIA: 361', 'Antecedents', $req, 'Ho poso al seu coneixement.')
    & $nouDocx $doc362 @('ID GIA: 362', $fi, 'Ho poso al seu coneixement.')
    $r1BA = Invoke-InformesDbAuto
    AssertEq "$($r1BA.Ok)|$($r1BA.NInformes)|$($r1BA.NActivitats)" 'True|2|2' 'base auto: la passada llegeix els dos informes'
    $eBA = _BaseAutoEstat
    AssertEq ([string]$eBA['mode']) 'auto' 'base auto: consta que l''ha feta l''automatic (data en verd al menu)'
    Assert (-not (_BaseAutoToca (Get-Date) $eBA['auto_el'])) 'base auto: el venciment queda servit (no es torna a llancar)'
    $logBA = Get-Content -LiteralPath (Get-AutoLogPath 'informes-db-log.txt') -Raw
    Assert ($logBA.Contains('informes=2')) 'base auto: el registre diu que ha fet'
    $estats = { param($db) (@($db.activitats | ForEach-Object { "$($_.id_gia)=$($_.estat_actual)" }) -join ',') }
    AssertEq (& $estats (Read-JsonFile $dbBA)) '361=Requeriment,362=FI Requeriment' 'base auto: l''estat de cada activitat'

    # Una correccio a ma ("Editar base") a la 361, i els dos informes es tornen
    # a escriure: la 361 diu ara "finalitzat" i la 362 "no finalitzat".
    $db1 = Read-JsonFile $dbBA
    $inf361 = @($db1.activitats | Where-Object { $_.id_gia -eq '361' })[0].informes[0]
    _MarcaEditatAMa $inf361
    $inf361.conclusio_breu = 'Precinte / Cessament'
    Write-JsonFile $dbBA $db1 8
    & $nouDocx $doc361 @('ID GIA: 361', $fi, 'Ho poso al seu coneixement.')
    & $nouDocx $doc362 @('ID GIA: 362', $req, 'Ho poso al seu coneixement.')
    foreach ($f in @($doc361, $doc362)) { (Get-Item -LiteralPath $f).LastWriteTimeUtc = (Get-Date).ToUniversalTime().AddMinutes(5) }
    $r2BA = Invoke-InformesDbAuto
    AssertEq ([int]$r2BA.Reprocessats) 2 'base auto: els dos informes tocats es tornen a llegir'
    $db2 = Read-JsonFile $dbBA
    AssertEq (& $estats $db2) '361=Precinte / Cessament,362=Requeriment' 'base auto: la correccio a ma PREVAL; la resta, el que diu l''informe ara'
    $inf361b = @($db2.activitats | Where-Object { $_.id_gia -eq '361' })[0].informes[0]
    AssertEq "$($inf361b.editat_a_ma)|$($inf361b.auto_conclusio_breu)" 'True|FI Requeriment' 'base auto: ...i es guarda el nou automatic per poder desfer'

    # L'EDITOR obert mentre l'automatic actualitza: desar no pot tornar enrere.
    $stE = @{ Db = (Read-JsonFile $dbBA); Path = $dbBA }
    $stE.Segell = _SegellBase $stE.Db
    $infE = @($stE.Db.activitats | Where-Object { $_.id_gia -eq '362' })[0].informes[0]
    _MarcaEditatAMa $infE
    $infE.ignorat = $true
    $infE361 = @($stE.Db.activitats | Where-Object { $_.id_gia -eq '361' })[0].informes[0]
    [void](_DesfesEditatAMa $infE361)
    AssertEq (Save-BaseEditada $stE) 'desat' 'Save-BaseEditada: la base no ha canviat -> es desa tal qual'
    # Ara si: l'automatic la reescriu (un informe nou) mentre l'editor es obert.
    & $nouDocx (Join-Path (Join-Path $baInf 'GIA 362') '2026-09-03_Seg_GIA_362.docx') @('ID GIA: 362', $fi, 'Ho poso al seu coneixement.')
    $stE.Db = Read-JsonFile $dbBA; $stE.Segell = _SegellBase $stE.Db
    $infE2 = @($stE.Db.activitats | Where-Object { $_.id_gia -eq '361' })[0].informes[0]
    _MarcaEditatAMa $infE2
    $infE2.conclusio_breu = 'Ampliaci' + [char]0x00F3 + ' termini'
    Start-Sleep -Milliseconds 50
    [void](Invoke-InformesDbAuto)
    AssertEq (Save-BaseEditada $stE) 'fusionat' 'Save-BaseEditada: la base ha canviat mentre era obert -> s''hi fusiona'
    $db3 = Read-JsonFile $dbBA
    AssertEq ([int]$db3.n_informes) 3 'fusio: l''informe nou de l''automatic no es perd'
    AssertEq (& $estats $db3) ('361=Ampliaci' + [char]0x00F3 + ' termini,362=FI Requeriment') 'fusio: la correccio de l''editor hi es, i l''estat es recalcula'

    # Ocupat: un altre proces esta escanejant (el boto, o l'automatic d'una
    # altra finestra del programa).
    $hBA = & $retenirMutex $Script:BaseMutexNom $ba
    try {
        Assert $hBA.Agafat 'prova: l''altre proces te el mutex de la base'
        $ocMA = @{ N = 0 }
        AssertEq (Invoke-AmbMutexUnic $Script:BaseMutexNom { throw 'no hauria de correr' } { $ocMA.N++ }) $false 'Invoke-AmbMutexUnic: ocupat -> no fa la feina i torna $false'
        AssertEq $ocMA.N 1 'Invoke-AmbMutexUnic: ...i avisa ($siOcupat)'
        [void](_BaseAutoDesaEstat @{ auto_el = '' })
        $rOc = Invoke-InformesDbAuto
        AssertEq ([string]$rOc.Error) 'ocupat' 'base auto: si ja s''esta actualitzant, no en fa un altre al damunt'
        Assert (-not [string]::IsNullOrWhiteSpace([string](_BaseAutoEstat)['auto_el'])) 'base auto: ...pero el venciment queda servit'
        AssertEq (Save-BaseEditada $stE) 'ocupat' 'Save-BaseEditada: mentre s''actualitza, no s''escriu (i es diu)'
        AssertEq (Wait-MutexLliure $Script:BaseMutexNom 1) $false 'Wait-MutexLliure: si no s''allibera, s''acaba cansant d''esperar'
    } finally { & $deixarMutex $hBA }

    # Sense la carpeta d'informes (fora de la feina): no es error, s'apunta.
    $InformesDir = Join-Path $ba 'no-hi-es'
    [void](_BaseAutoDesaEstat @{ auto_el = ''; mode = 'auto' })
    $r4BA = Invoke-InformesDbAuto
    AssertEq "$($r4BA.Ok)|$((_BaseAutoEstat)['mode'])" 'False|auto' 'base auto: sense la carpeta no fa res (i no toca el mode)'
    Assert (-not [string]::IsNullOrWhiteSpace([string](_BaseAutoEstat)['auto_el'])) 'base auto: ...pero apunta la passada (si no, cada minut)'
    Assert ((Get-Content -LiteralPath (Get-AutoLogPath 'informes-db-log.txt') -Raw).Contains('ATURAT')) 'base auto: ...i ho diu al registre'
} finally {
    $InformesDir = $vellsBA.Inf; $LocalActivitatsDir = $vellsBA.Loc; $ActivitatsDir = $vellsBA.Act; $env:LOCALAPPDATA = $vellsBA.App
    Remove-Item -LiteralPath $ba -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "`n--- InformesEscaneig.ps1: ruta relativa (una altra unitat), germans de carpeta, GIA diferent, correccions que ja no ho son ---"
# D'extrem a extrem amb .docx de debo ($nouDocx, del bloc d'abans). Textos i
# GIA inventats.
$rb = Join-Path ([System.IO.Path]::GetTempPath()) ('base-rel-' + [guid]::NewGuid().ToString('N'))
$rbI = Join-Path (Join-Path $rb 'I') 'Informes'; $rbF = Join-Path (Join-Path $rb 'F') 'Informes'; $rbZ = Join-Path (Join-Path $rb 'Z') 'Informes'
$rbLoc = Join-Path $rb 'local'; $rbAct = Join-Path $rb 'activitats'; $rbApp = Join-Path $rb 'appdata'
$vellsRB = @{ Inf = $InformesDir; Loc = $LocalActivitatsDir; Act = $ActivitatsDir; App = $env:LOCALAPPDATA }
try {
    foreach ($d in @((Join-Path $rbI 'GIA 501'), (Join-Path $rbI 'Expedient 2025-1-9999'), (Join-Path $rbI 'GIA 503'), $rbZ, $rbLoc, $rbAct, $rbApp)) { [void](New-Item -ItemType Directory -Path $d -Force) }
    $LocalActivitatsDir = $rbLoc; $ActivitatsDir = $rbAct; $env:LOCALAPPDATA = $rbApp
    $reqT = "Vist l'anterior, s'inicia d'ofici el procediment d'esmena, disposant d'un termini d'un mes."
    $fiT  = "Vist l'anterior s'informa que es pot donar per finalitzat el procediment d'esmena."
    $llicT = "S'informa favorablement l'activitat i es d" + [char]0x00F3 + "na per tancat l'expedient."
    $fi = 'Ho poso al seu coneixement als efectes oportuns,'
    & $nouDocx (Join-Path (Join-Path $rbI 'GIA 501') '2026-01-10_Req.docx') @('ID GIA: 501', $reqT, $fi)
    & $nouDocx (Join-Path (Join-Path $rbI 'GIA 501') '2026-02-10_Seg.docx') @('ID GIA: 501', $fiT, $fi)
    & $nouDocx (Join-Path (Join-Path $rbI 'Expedient 2025-1-9999') '2026-03-01_Req.docx') @('ID GIA: 502', $reqT, $fi)
    & $nouDocx (Join-Path (Join-Path $rbI 'Expedient 2025-1-9999') '2026-04-01_Llic.docx') @('ID GIA: -', $llicT, $fi)
    & $nouDocx (Join-Path (Join-Path $rbI 'GIA 503') '2026-05-01_X.docx') @('ID GIA: 530', $fiT, $fi)
    $vell = (Get-Date).ToUniversalTime().AddHours(-2)
    Get-ChildItem -LiteralPath $rbI -Recurse -File | ForEach-Object { $_.LastWriteTimeUtc = $vell }
    $dbRB = Join-Path $rbLoc 'informes-db.json'
    $actDe = { param($db, $gia) @($db.activitats | Where-Object { [string]$_.id_gia -eq $gia })[0] }

    $InformesDir = $rbI
    $resI = Invoke-InformesDbEscaneig
    $dbI = Read-JsonFile $dbRB
    AssertEq "$($resI.Ok)|$($resI.NInformes)|$($dbI.versio_classificador)" ("True|5|" + $Script:ClassificadorVersio) 'escaneig: llegeix els 5 informes i desa la versio del classificador'
    $a502 = & $actDe $dbI '502'
    AssertEq @($a502.informes).Count 2 'germans: l''informe sense GIA va amb el GIA dels altres de la seva carpeta'
    AssertEq "$($a502.estat_actual)|$(@($a502.informes)[1].tipus)" ($Script:EstatFavorablePost + '|llicfav') 'germans: ...i el favorable (post) de llicencia decideix l''estat'
    Assert (-not (@($dbI.a_revisar | ForEach-Object { $_.motiu }) -join '|').Contains('sense ID GIA')) 'germans: ...i ja no surt "sense ID GIA" a revisar'
    $a530 = & $actDe $dbI '530'
    AssertEq ([string]@($a530.informes)[0].motiu) 'GIA del document diferent del de la carpeta' 'GIA 530 al document i GIA 503 a la carpeta -> a revisar, amb el motiu'
    AssertEq ([string]$dbI.carpeta_arrel) $rbI 'la base diu de quina carpeta es (carpeta_arrel)'

    # Dues correccions a ma a la 501: una de debo (Precinte) i una que ja
    # coincideix amb el que diu l'automatic (com les 179 d'octubre de 2026).
    $a501 = & $actDe $dbI '501'
    $infR = @($a501.informes)[0]; _MarcaEditatAMa $infR; $infR.conclusio_breu = 'Precinte / Cessament'
    $infS = @($a501.informes)[1]; _MarcaEditatAMa $infS; $infS.auto_conclusio_breu = 'Revisar'
    Write-JsonFile $dbRB $dbI 8

    # LA MATEIXA BASE AMB LA CARPETA EN UNA ALTRA UNITAT (F: fora de la feina).
    [void](New-Item -ItemType Directory -Path (Split-Path -Parent $rbF) -Force)
    Copy-Item -LiteralPath $rbI -Destination (Split-Path -Parent $rbF) -Recurse
    Get-ChildItem -LiteralPath $rbF -Recurse -File | ForEach-Object { $_.LastWriteTimeUtc = $vell }
    $InformesDir = $rbF
    $resF = Invoke-InformesDbEscaneig
    $dbF = Read-JsonFile $dbRB
    AssertEq "$($resF.Ok)|$($resF.Reprocessats)" 'True|0' 'una altra unitat: tots els informes casen per la ruta relativa (cap es reprocessa)'
    $infRF = @((& $actDe $dbF '501').informes)[0]
    AssertEq "$($infRF.conclusio_breu)|$($infRF.editat_a_ma)" 'Precinte / Cessament|True' 'una altra unitat: la correccio a ma NO es perd'
    Assert ([string]$infRF.ruta).StartsWith($rbF) 'una altra unitat: la ruta desada passa a ser la d''ara'
    AssertEq ([string]$dbF.carpeta_arrel) $rbF '...i la carpeta_arrel tambe'

    # Una VERSIO NOVA del classificador: es tornen a llegir TOTS, les correccions
    # es queden, i la que ja coincideix amb l'automatic deixa de ser-ho.
    $dbF.versio_classificador = 'vella'
    Write-JsonFile $dbRB $dbF 8
    $resV = Invoke-InformesDbEscaneig
    $dbV = Read-JsonFile $dbRB
    AssertEq ([int]$resV.Reprocessats) 5 'versio del classificador diferent: es tornen a llegir tots els informes'
    $infsV = @((& $actDe $dbV '501').informes)
    AssertEq "$($infsV[0].conclusio_breu)|$($infsV[0].editat_a_ma)|$($infsV[0].auto_conclusio_breu)" 'Precinte / Cessament|True|Requeriment' 'la correccio de debo es queda (amb el nou automatic al costat)'
    AssertEq "$($infsV[1].conclusio_breu)|$($infsV[1].editat_a_ma)|$($null -eq $infsV[1].PSObject.Properties['auto_conclusio_breu'])" 'FI Requeriment|False|True' 'la "correccio" que ja coincideix amb l''automatic deixa de ser-ho (sense auto_*)'

    # UNA ALTRA CARPETA on no casa res: preguntar abans d'escriure.
    [void](New-Item -ItemType Directory -Path (Join-Path $rbZ 'GIA 777') -Force)
    & $nouDocx (Join-Path (Join-Path $rbZ 'GIA 777') '2026-06-01_Altre.docx') @('ID GIA: 777', $fiT, $fi)
    $InformesDir = $rbZ
    $preg = @{ Text = '' }
    $resNo = Invoke-InformesDbEscaneig $null { param($p) $preg.Text = $p; $false }
    AssertEq "$($resNo.Ok)|$($resNo.Cancelat)" 'False|True' 'una altra carpeta que no casa: si es diu que no, no s''escriu'
    Assert ($preg.Text.Contains($rbF) -and $preg.Text.Contains($rbZ) -and $preg.Text.Contains('perdrien 1 correccions')) 'la pregunta diu de quina carpeta es la base, quina s''escanejaria i quantes correccions es perdrien'
    AssertEq ([int](Read-JsonFile $dbRB).n_informes) 5 '...i la base es queda com era'
    $resAuto = Invoke-InformesDbEscaneig
    AssertEq "$($resAuto.Ok)|$($resAuto.Cancelat)" 'False|True' 'sense ningu a qui preguntar (l''automatic): tampoc no s''escriu'
    $resSi = Invoke-InformesDbEscaneig $null { param($p) $true }
    AssertEq "$($resSi.Ok)|$([int](Read-JsonFile $dbRB).n_informes)" 'True|1' 'si es diu que si, es fa'
} catch {
    Assert $false ('bloc ruta relativa / germans: ' + $_.Exception.Message + ' @ ' + $_.InvocationInfo.ScriptLineNumber)
} finally {
    $InformesDir = $vellsRB.Inf; $LocalActivitatsDir = $vellsRB.Loc; $ActivitatsDir = $vellsRB.Act; $env:LOCALAPPDATA = $vellsRB.App
    Remove-Item -LiteralPath $rb -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "`n--- InformesEscaneig.ps1: nomes el PRIMER NIVELL (ni l'arrel ni les subcarpetes) ---"
AssertEq (_EsDePrimerNivell 'GIA 1\a.docx') $true '_EsDePrimerNivell: carpeta\fitxer -> si'
AssertEq (_EsDePrimerNivell 'a.docx') $false '_EsDePrimerNivell: a l''arrel -> no'
AssertEq (_EsDePrimerNivell 'GIA 1\Expedient\a.docx') $false '_EsDePrimerNivell: dins d''una subcarpeta -> no'
$pn = Join-Path ([System.IO.Path]::GetTempPath()) ('primer-nivell-' + [guid]::NewGuid().ToString('N'))
$pnI = Join-Path $pn 'Informes'; $vellsPN = @{ Inf = $InformesDir; Loc = $LocalActivitatsDir; Act = $ActivitatsDir; App = $env:LOCALAPPDATA }
try {
    foreach ($d in @((Join-Path $pnI 'GIA 801'), (Join-Path (Join-Path $pnI 'GIA 801') 'Expedient GIA'), (Join-Path $pn 'local'), (Join-Path $pn 'act'), (Join-Path $pn 'app'))) { [void](New-Item -ItemType Directory -Path $d -Force) }
    $LocalActivitatsDir = Join-Path $pn 'local'; $ActivitatsDir = Join-Path $pn 'act'; $env:LOCALAPPDATA = Join-Path $pn 'app'; $InformesDir = $pnI
    $fiPN = "Vist l'anterior s'informa que es pot donar per finalitzat el procediment d'esmena."
    & $nouDocx (Join-Path (Join-Path $pnI 'GIA 801') '2026-01-10_Seg.docx') @('ID GIA: 801', $fiPN, 'Ho poso al seu coneixement als efectes oportuns,')
    & $nouDocx (Join-Path (Join-Path (Join-Path $pnI 'GIA 801') 'Expedient GIA') '2026-02-10_Vell.docx') @('ID GIA: 801', "Vist l'anterior, cal requerir l'esmena de les deficiències indicades.", 'Ho poso al seu coneixement als efectes oportuns,')
    & $nouDocx (Join-Path $pnI '2026-03-10_Solt.docx') @('ID GIA: 802', $fiPN, 'Ho poso al seu coneixement als efectes oportuns,')
    AssertEq (@(Get-FitxersPrimerNivell $pnI | ForEach-Object { $_.Name }) -join '|') '2026-01-10_Seg.docx' 'Get-FitxersPrimerNivell: nomes els de dins de cada carpeta (ni l''arrel ni les subcarpetes)'
    $resPN = Invoke-InformesDbEscaneig
    $dbPN = Read-JsonFile (Join-Path $LocalActivitatsDir 'informes-db.json')
    AssertEq "$($resPN.NInformes)|$(@($dbPN.activitats)[0].estat_actual)" '1|FI Requeriment' 'Actualitzar base: el requeriment de la subcarpeta (mes nou) no decideix l''estat, i el de l''arrel no hi entra'
} catch {
    Assert $false ('bloc primer nivell: ' + $_.Exception.Message + ' @ ' + $_.InvocationInfo.ScriptLineNumber)
} finally {
    $InformesDir = $vellsPN.Inf; $LocalActivitatsDir = $vellsPN.Loc; $ActivitatsDir = $vellsPN.Act; $env:LOCALAPPDATA = $vellsPN.App
    Remove-Item -LiteralPath $pn -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "`n--- InformesEscaneig.ps1: l'ID GIA assignat a ma (gia-assignats_*.json) ---"
# Peces pures. Rutes i GIA inventats.
$gaObj = [pscustomobject]@{
    informes = @(
        [pscustomobject]@{ ruta_relativa = 'Antic/2019-01-01_Req.docx'; id_gia = 610; confianca = 'segur'; motiu = 'm'; forcar = $false },
        [pscustomobject]@{ ruta_relativa = '\GIA 620\2019-02-01_Req.docx'; id_gia = '293'; confianca = 'segur'; motiu = 'capcalera equivocada'; forcar = $true },
        [pscustomobject]@{ ruta_relativa = ''; id_gia = '1' },
        [pscustomobject]@{ ruta_relativa = 'x.docx'; id_gia = '' })
    sense_gia = @([pscustomobject]@{ ruta_relativa = 'Sense/2018-01-01_Z.docx'; motiu = 'no se sap' })
}
$gaM = _GiaAssignatsMapa $gaObj
AssertEq (@($gaM.Mapa.Keys | Sort-Object) -join '|') 'Antic\2019-01-01_Req.docx|GIA 620\2019-02-01_Req.docx' '_GiaAssignatsMapa: clau com _ClauInforme (barres i "\" inicial), sense les entrades buides'
AssertEq "$($gaM.Mapa['antic\2019-01-01_req.docx'].Gia)|$($gaM.Mapa['GIA 620\2019-02-01_Req.docx'].Forcar)" '610|True' '_GiaAssignatsMapa: id_gia numeric -> text, forcar, i la clau no distingeix majuscules'
AssertEq (@($gaM.SenseGia.Keys) -join '|') 'Sense\2018-01-01_Z.docx' '_GiaAssignatsMapa: sense_gia'
AssertEq (@(_GiaAssignatsNoTrobats $gaM @{ 'Antic\2019-01-01_Req.docx' = $true }) -join '|') 'GIA 620\2019-02-01_Req.docx|Sense\2018-01-01_Z.docx' '_GiaAssignatsNoTrobats: les que no casen amb cap informe (informes i sense_gia)'
AssertEq (@(_GiaAssignatsMapa $null).Count) 1 '_GiaAssignatsMapa: un fitxer buit no peta'

# D'extrem a extrem, amb .docx de debo ($nouDocx).
$ga = Join-Path ([System.IO.Path]::GetTempPath()) ('gia-assignats-' + [guid]::NewGuid().ToString('N'))
$gaI = Join-Path $ga 'Informes'; $gaLoc = Join-Path $ga 'local'; $gaAct = Join-Path $ga 'activitats'; $gaApp = Join-Path $ga 'appdata'
$vellsGA = @{ Inf = $InformesDir; Loc = $LocalActivitatsDir; Act = $ActivitatsDir; App = $env:LOCALAPPDATA }
try {
    foreach ($d in @((Join-Path $gaI 'Antic'), (Join-Path $gaI 'GIA 620'), (Join-Path $gaI 'GIA 630'), (Join-Path $gaI 'GIA 650'), $gaLoc, $gaAct, $gaApp)) { [void](New-Item -ItemType Directory -Path $d -Force) }
    $LocalActivitatsDir = $gaLoc; $ActivitatsDir = $gaAct; $env:LOCALAPPDATA = $gaApp; $InformesDir = $gaI
    $reqG = "Vist l'anterior, s'inicia d'ofici el procediment d'esmena, disposant d'un termini d'un mes."
    $fiG = 'Ho poso al seu coneixement als efectes oportuns,'
    & $nouDocx (Join-Path (Join-Path $gaI 'Antic') '2019-01-01_Req.docx') @('Sense capcalera', $reqG, $fiG)
    & $nouDocx (Join-Path (Join-Path $gaI 'GIA 620') '2019-02-01_Req.docx') @('ID GIA: 239', $reqG, $fiG)
    & $nouDocx (Join-Path (Join-Path $gaI 'GIA 630') '2019-03-01_X.docx') @('ID GIA: 631', $reqG, $fiG)
    & $nouDocx (Join-Path (Join-Path $gaI 'GIA 650') '2019-04-01_Y.docx') @('Sense capcalera', $reqG, $fiG)
    $vellG = (Get-Date).ToUniversalTime().AddHours(-2)
    Get-ChildItem -LiteralPath $gaI -Recurse -File | ForEach-Object { $_.LastWriteTimeUtc = $vellG }
    $gaFitxer = { param($nom, $entrades, $sense)
        $p = Join-Path $gaLoc $nom
        Write-JsonFile $p ([pscustomobject]@{ informes = @($entrades); sense_gia = @($sense) }) 5
        (Get-Item -LiteralPath $p).LastWriteTimeUtc = $vellG
        $p
    }
    $gaE = { param($ruta, $gia, $forcar) [pscustomobject]@{ ruta_relativa = $ruta; id_gia = $gia; confianca = 'segur'; motiu = 'prova'; forcar = $forcar } }
    [void](& $gaFitxer 'gia-assignats_2026-10-01.json' @(
        (& $gaE 'Antic\2019-01-01_Req.docx' '610' $false),
        (& $gaE 'GIA 620\2019-02-01_Req.docx' '293' $true),
        (& $gaE 'GIA 630\2019-03-01_X.docx' '640' $false),
        (& $gaE 'GIA 650\2019-04-01_Y.docx' '660' $false),
        (& $gaE 'Ja no hi es\2019-05-01_Q.docx' '700' $false)) @([pscustomobject]@{ ruta_relativa = 'Tampoc\2018-01-01_Z.docx'; motiu = 'x' }))
    $dbGA = Join-Path $gaLoc 'informes-db.json'
    $infDe = { param($db, $fitxer) foreach ($a in @($db.activitats)) { foreach ($i in @($a.informes)) { if ([string]$i.fitxer -eq $fitxer) { return [pscustomobject]@{ Gia = [string]$a.id_gia; Inf = $i } } } } }

    $resG = Invoke-InformesDbEscaneig
    $dbG = Read-JsonFile $dbGA
    AssertEq "$($resG.Ok)|$($resG.NInformes)" 'True|4' 'GIA assignat: l''escaneig llegeix els 4 informes'
    $x = & $infDe $dbG '2019-01-01_Req.docx'
    AssertEq "$($x.Gia)|$($x.Inf.gia_font)|$($x.Inf.motiu)" '610|assignat|' 'sense GIA al document ni a la carpeta -> el del fitxer, marcat "assignat" i sense "sense ID GIA"'
    $x = & $infDe $dbG '2019-02-01_Req.docx'
    AssertEq "$($x.Gia)|$($x.Inf.gia_font)|$($x.Inf.motiu)" '293|assignat|' 'forcar=true: mana sobre la capcalera equivocada (239), i sense l''avis "GIA del document diferent"'
    $x = & $infDe $dbG '2019-03-01_X.docx'
    AssertEq "$($x.Gia)|$($x.Inf.gia_font)|$($x.Inf.motiu)" '631|document|GIA del document diferent del de la carpeta' 'sense forcar, el document mana sobre l''assignacio (i l''avis de la carpeta es queda)'
    $x = & $infDe $dbG '2019-04-01_Y.docx'
    AssertEq "$($x.Gia)|$($x.Inf.gia_font)" '660|assignat' 'l''assignacio mana sobre la carpeta "GIA 650"'
    AssertEq (@($resG.GiaAssignatsNoTrobats) -join '|') 'Ja no hi es\2019-05-01_Q.docx|Tampoc\2018-01-01_Z.docx' 'les entrades que ja no troben l''informe no peten: es retornen per avisar'
    Assert ([string]$dbG.gia_assignats).StartsWith('gia-assignats_2026-10-01.json|') 'la base apunta quin fitxer d''assignacions s''ha fet servir'

    # Sense canvis: res es reprocessa i la marca es conserva.
    $resG2 = Invoke-InformesDbEscaneig
    $x = & $infDe (Read-JsonFile $dbGA) '2019-01-01_Req.docx'
    AssertEq "$($resG2.Reprocessats)|$($x.Gia)|$($x.Inf.gia_font)" '0|610|assignat' 'sense canvis: no es reprocessa res, i el GIA i la marca es queden'

    # Un fitxer NOU (mes recent): es tornen a llegir tots i mana el nou.
    [void](& $gaFitxer 'gia-assignats_2026-10-08.json' @((& $gaE 'Antic\2019-01-01_Req.docx' '611' $false)) @())
    $resG3 = Invoke-InformesDbEscaneig
    $dbG3 = Read-JsonFile $dbGA
    $x = & $infDe $dbG3 '2019-01-01_Req.docx'
    AssertEq "$($resG3.Reprocessats)|$($x.Gia)" '4|611' 'un fitxer d''assignacions mes recent: es tornen a llegir tots i mana el nou'
    $x = & $infDe $dbG3 '2019-02-01_Req.docx'
    AssertEq "$($x.Gia)|$($x.Inf.gia_font)" '239|document' '...i el que el nou ja no porta torna al que diu el document'

    # Un fitxer que no es pot llegir: l'escaneig no peta, ho diu.
    [System.IO.File]::WriteAllText((Join-Path $gaLoc 'gia-assignats_2026-10-09.json'), '{ aixo no es json')
    $resG4 = Invoke-InformesDbEscaneig
    Assert ([bool]$resG4.Ok -and ([string]$resG4.GiaAssignatsError).Contains('gia-assignats_2026-10-09.json')) 'un fitxer d''assignacions trencat: l''escaneig es fa igual i retorna l''error per dir-ho'
} catch {
    Assert $false ('bloc GIA assignat: ' + $_.Exception.Message + ' @ ' + $_.InvocationInfo.ScriptLineNumber)
} finally {
    $InformesDir = $vellsGA.Inf; $LocalActivitatsDir = $vellsGA.Loc; $ActivitatsDir = $vellsGA.Act; $env:LOCALAPPDATA = $vellsGA.App
    Remove-Item -LiteralPath $ga -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "`n--- InformesEscaneig.ps1: peces pures (_ClauInforme, _AvisCanviArrel, _GiaDelsGermans) ---"
AssertEq (_ClauInforme 'I:\Act\Informes\GIA 1\a.docx' 'I:\Act\Informes') 'GIA 1\a.docx' '_ClauInforme: la ruta relativa a l''arrel'
AssertEq (_ClauInforme 'F:\FEINA\Informes\GIA 1\a.docx' 'F:\FEINA\Informes\') 'GIA 1\a.docx' '_ClauInforme: una altra unitat, la mateixa clau (i la barra final no compta)'
AssertEq (_ClauInforme 'i:\act\informes\GIA 1\a.docx' 'I:\Act\Informes') 'GIA 1\a.docx' '_ClauInforme: sense distingir majuscules a l''arrel'
AssertEq (_ClauInforme '/tmp/inf/GIA 1/a.docx' '/tmp/inf') 'GIA 1\a.docx' '_ClauInforme: les dues barres valen igual'
AssertEq (_ClauInforme 'X:\altre\a.docx' 'I:\Act\Informes') 'X:\altre\a.docx' '_ClauInforme: fora de l''arrel, la ruta sencera'
AssertEq (_ClauInforme 'I:\x\a.docx' '') 'I:\x\a.docx' '_ClauInforme: sense arrel (bases antigues), la ruta sencera'
AssertEq (_AvisCanviArrel 'I:\Inf' 'i:\inf\' 100 0 5) '' '_AvisCanviArrel: la mateixa arrel -> no pregunta'
AssertEq (_AvisCanviArrel 'I:\Inf' 'F:\Inf' 100 60 5) '' '_AvisCanviArrel: una altra arrel pero casa la majoria -> no pregunta'
AssertEq (_AvisCanviArrel 'I:\Inf' 'F:\Altre' 100 2 0) '' '_AvisCanviArrel: no casa pero no es perd cap correccio -> no pregunta'
Assert ((_AvisCanviArrel 'I:\Inf' 'F:\Altre' 100 2 7).Contains('perdrien 7 correccions')) '_AvisCanviArrel: no casa i es perdrien correccions -> pregunta'
AssertEq (_AvisCanviArrel '' 'F:\Altre' 100 0 7) '' '_AvisCanviArrel: base sense carpeta_arrel (antiga) -> no pregunta'
$germ = @(
    [pscustomobject]@{ Gia = '7'; GiaFont = 'document'; Ruta = 'I:\Inf\Exp A\1.docx'; Motius = @() },
    [pscustomobject]@{ Gia = ''; GiaFont = 'document'; Ruta = 'I:\Inf\Exp A\2.docx'; Motius = @('sense ID GIA', 'sense conclusio') },
    [pscustomobject]@{ Gia = '8'; GiaFont = 'document'; Ruta = 'I:\Inf\Exp B\1.docx'; Motius = @() },
    [pscustomobject]@{ Gia = '9'; GiaFont = 'document'; Ruta = 'I:\Inf\Exp B\2.docx'; Motius = @() },
    [pscustomobject]@{ Gia = ''; GiaFont = 'document'; Ruta = 'I:\Inf\Exp B\3.docx'; Motius = @('sense ID GIA') },
    [pscustomobject]@{ Gia = ''; GiaFont = 'document'; Ruta = 'I:\Inf\Exp C\1.docx'; Motius = @('sense ID GIA') }
)
AssertEq (_GiaDelsGermans $germ) 1 '_GiaDelsGermans: en resol un'
AssertEq "$($germ[1].Gia)|$($germ[1].GiaFont)|$(@($germ[1].Motius) -join ',')" '7|germans|sense conclusio' '_GiaDelsGermans: tots els germans del mateix GIA -> aquell GIA (i treu "sense ID GIA")'
AssertEq $germ[4].Gia '' '_GiaDelsGermans: germans de GIA diferents -> no en tria cap'
AssertEq $germ[5].Gia '' '_GiaDelsGermans: sense germans amb GIA -> res'

Write-Host "`n--- Informes.ps1: _FusionaEdicionsBase (desar l'editor damunt d'una base nova) ---"
$mkInfF = { param($ruta, $breu, $ign, $ed) $o = [pscustomobject]@{ ruta = $ruta; data = '2026-01-01'; conclusio = ''; conclusio_breu = $breu; ignorat = $ign; editat_a_ma = $false }; if ($ed) { _MarcaEditatAMa $o; $o.conclusio_breu = $ed }; $o }
$discF = [pscustomobject]@{ activitats = @(
    [pscustomobject]@{ id_gia = '1'; estat_actual = 'Requeriment'; informes = @((& $mkInfF 'a' 'Requeriment' $false $null)) },
    [pscustomobject]@{ id_gia = '2'; estat_actual = 'Precinte / Cessament'; informes = @((& $mkInfF 'b' 'Favorable' $false 'Precinte / Cessament')) },
    [pscustomobject]@{ id_gia = '3'; estat_actual = 'Favorable'; informes = @((& $mkInfF 'c' 'Favorable' $false $null), (& $mkInfF 'nou' 'Requeriment' $false $null)) }) }
$editorF = [pscustomobject]@{ activitats = @(
    [pscustomobject]@{ id_gia = '1'; informes = @((& $mkInfF 'a' 'Requeriment' $false 'FI Requeriment')) },
    [pscustomobject]@{ id_gia = '2'; informes = @((& $mkInfF 'b' 'Favorable' $false $null)) },
    [pscustomobject]@{ id_gia = '3'; informes = @((& $mkInfF 'c' 'Favorable' $false $null)) }) }
$fF = _FusionaEdicionsBase $discF $editorF
AssertEq (@($fF.activitats | ForEach-Object { "$($_.id_gia)=$($_.estat_actual)" }) -join ',') '1=FI Requeriment,2=Favorable,3=Favorable' 'fusio: la correccio entra, el desfet es desfa, la resta igual'
AssertEq "$($fF.activitats[0].informes[0].editat_a_ma)|$($fF.activitats[0].informes[0].auto_conclusio_breu)" 'True|Requeriment' 'fusio: el valor automatic, el del disc'
AssertEq "$($fF.activitats[1].informes[0].editat_a_ma)|$($null -eq $fF.activitats[1].informes[0].PSObject.Properties['auto_ignorat'])" 'False|True' 'fusio: el desfet torna a l''automatic del disc'
AssertEq @($fF.activitats[2].informes).Count 2 'fusio: l''informe nou del disc hi es'
# L'editor i el disc amb la carpeta en unitats diferents: casen per la ruta relativa.
$discU = [pscustomobject]@{ carpeta_arrel = 'F:\Inf'; activitats = @([pscustomobject]@{ id_gia = '1'; estat_actual = 'Requeriment'; informes = @((& $mkInfF 'F:\Inf\GIA 1\a.docx' 'Requeriment' $false $null)) }) }
$editorU = [pscustomobject]@{ carpeta_arrel = 'I:\Inf'; activitats = @([pscustomobject]@{ id_gia = '1'; informes = @((& $mkInfF 'I:\Inf\GIA 1\a.docx' 'Requeriment' $false 'FI Requeriment')) }) }
AssertEq ((_FusionaEdicionsBase $discU $editorU).activitats[0].estat_actual) 'FI Requeriment' 'fusio: editor i disc en unitats diferents casen per la ruta relativa'

Write-Host "`n--- ModeAutomatic.ps1: el registre dels interruptors A/M ---"
foreach ($kMA in @('copiarinformes', 'informesdb', 'planol')) {
    Assert ($Script:ModesAuto.Contains($kMA)) "registre: '$kMA' hi es"
    foreach ($cMA in @('Titol', 'Actiu', 'DesaActiu', 'UltimMode', 'SiToca', 'Requisit', 'TipA', 'TipM')) {
        Assert ($null -ne $Script:ModesAuto[$kMA][$cMA]) "registre: '$kMA' te $cMA"
    }
}
$rgDir = Join-Path ([System.IO.Path]::GetTempPath()) ('registre-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $rgDir -Force)
$vellsRG = @{ Loc = $LocalActivitatsDir; Inf = $InformesDir; Cop = $CopiaInformesDir }
$LocalActivitatsDir = $rgDir
try {
    $mB = $Script:ModesAuto['informesdb']
    AssertEq ([bool](& $mB.Actiu)) $false 'registre base: l''interruptor arrenca apagat'
    [void](& $mB.DesaActiu $true)
    AssertEq ([bool](& $mB.Actiu)) $true 'registre base: DesaActiu / Actiu'
    [void](_BaseAutoDesaEstat @{ mode = 'manual' })
    AssertEq ([string](& $mB.UltimMode)) 'manual' 'registre base: UltimMode'
    AssertEq ([bool](& $mB.Actiu)) $true 'registre base: desar el mode no apaga l''interruptor'
    $InformesDir = ''
    Assert (([string](& $mB.Requisit)).Contains('Configuraci')) 'registre base: sense carpeta d''informes no es pot engegar'
    $InformesDir = 'X:\Informes'
    AssertEq ([string](& $mB.Requisit)) '' 'registre base: amb carpeta, si'
    $CopiaInformesDir = ''
    Assert (([string](& $Script:ModesAuto['copiarinformes'].Requisit)).Contains('Configuraci')) 'registre copia: sense carpeta de copia no es pot engegar'
    # EL PLANOL (setmanal): l'estat a planol-auto.json, com els altres.
    $mP = $Script:ModesAuto['planol']
    $vellRepo = $RepoRoot
    $RepoRoot = $rgDir
    AssertEq ([bool](& $mP.Actiu)) $false 'registre planol: arrenca apagat'
    [void](& $mP.DesaActiu $true)
    AssertEq "$([bool](& $mP.Actiu))|$([string](& $mP.UltimMode))" 'True|' 'registre planol: DesaActiu / Actiu'
    Assert ((_PlanolAutoStatePath).EndsWith('planol-auto.json')) 'registre planol: l estat, a planol-auto.json'
    $RepoRoot = $vellRepo
    Assert (([string](& $mP.TipA)).Contains('cada dilluns')) 'registre planol: l ajuda de l interruptor diu cada dilluns'
} finally {
    $LocalActivitatsDir = $vellsRG.Loc; $InformesDir = $vellsRG.Inf; $CopiaInformesDir = $vellsRG.Cop
    Remove-Item -LiteralPath $rgDir -Recurse -Force -ErrorAction SilentlyContinue
}

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

Write-Host "`n--- CorreuVia.ps1: EmailJS o l'Outlook (octubre 2026) ---"
try {
    AssertEq (_CorreuViaValida 'outlook') 'outlook' 'via: outlook es valida'
    AssertEq (_CorreuViaValida ' Outlook-Esborrany ') 'outlook-esborrany' 'via: sense distingir majuscules ni espais'
    AssertEq (_CorreuViaValida 'gmail') 'emailjs' 'via desconeguda (settings escrit a ma) -> EmailJS, no deixa el PC sense correu'
    AssertEq (_CorreuViaValida $null) 'emailjs' 'via buida -> EmailJS'
    Assert (_CorreuViaEsOutlook 'outlook-esborrany') 'via: el mode esborrany tambe es Outlook'
    Assert (-not (_CorreuViaEsOutlook 'emailjs')) 'via: EmailJS no es Outlook'

    # Desar la via no pot trepitjar la resta de settings.json.
    $cvSet = [pscustomobject]@{ InformesDir = 'F:\Informes'; Automatismes = @{ planol = 'x' } }
    $cvH = _SettingsAmbCorreuVia $cvSet 'outlook'
    AssertEq $cvH['CorreuVia'] 'outlook' 'Set-CorreuVia: desa la via triada'
    AssertEq $cvH['InformesDir'] 'F:\Informes' 'Set-CorreuVia: conserva les carpetes'
    Assert ($cvH.Contains('Automatismes')) 'Set-CorreuVia: conserva els automatismes'
    $cvH2 = _SettingsAmbCorreuVia $cvH 'emailjs'
    Assert (-not $cvH2.Contains('CorreuVia')) 'Set-CorreuVia: la per defecte no s''escriu (nomes el que difereix)'
    AssertEq $cvH2['InformesDir'] 'F:\Informes' 'Set-CorreuVia: tornar a EmailJS tampoc trepitja res'
    $cvH3 = _SettingsAmbCorreuVia $null 'outlook-esborrany'
    AssertEq $cvH3['CorreuVia'] 'outlook-esborrany' 'Set-CorreuVia: sense settings.json previ'

    AssertEq (_OutlookAdreces 'a@x.cat, b@x.cat;c@x.cat ;') 'a@x.cat; b@x.cat; c@x.cat' 'Outlook: adreces separades per ";"'
    AssertEq (_OutlookAdreces '') '' 'Outlook: cap adreca -> buit'

    AssertEq (Test-CorreuViaLlest 'outlook') '' 'Outlook: no necessita les claus d''EmailJS'
    AssertEq (Test-CorreuViaLlest 'emailjs') (Test-CorreuLlest) 'EmailJS: la mateixa comprovacio de les claus'

    # L'Outlook amb un doble de COM: que el mode esborrany DESA i no envia.
    $cvLog = New-Object System.Collections.ArrayList
    $cvNouItem = {
        $it = [pscustomobject]@{ To = ''; BCC = ''; Subject = ''; HTMLBody = '' }
        $it | Add-Member ScriptMethod Save { [void]$cvLog.Add('save:' + $this.To + '|' + $this.BCC) }
        $it | Add-Member ScriptMethod Send { [void]$cvLog.Add('send:' + $this.To) }
        return $it
    }
    $cvOl = [pscustomobject]@{}
    $cvOl | Add-Member ScriptMethod CreateItem { param($t) & $cvNouItem }
    $cvSes = @{ Via = 'outlook-esborrany'; Cfg = $null; Outlook = $cvOl; JaObert = $true; Desats = 0; Enviats = 0 }
    Send-CorreuSessio $cvSes 'a@x.cat,b@x.cat' 'c@x.cat' 'Assumpte' '<p>hola</p>'
    AssertEq ($cvLog -join ',') 'save:a@x.cat; b@x.cat|c@x.cat' 'Outlook esborrany: desa, no envia'
    AssertEq $cvSes.Desats 1 'Outlook esborrany: compta els desats'
    AssertEq $cvSes.Enviats 0 'Outlook esborrany: cap enviat'
    $cvLog.Clear()
    $cvSes.Via = 'outlook'
    Send-CorreuSessio $cvSes 'a@x.cat' '' 'Assumpte' '<p>hola</p>'
    AssertEq ($cvLog -join ',') 'send:a@x.cat' 'Outlook: envia'
    AssertEq $cvSes.Enviats 1 'Outlook: compta els enviats'
    Assert ((_CorreuSessioError $cvSes ([System.Management.Automation.ErrorRecord]::new((New-Object Exception 'bloquejat'), 'x', 'NotSpecified', $null))) -like '*Outlook*bloquejat*') 'Outlook: l''error diu que ha fallat l''Outlook i per que'

    # --- El remitent (octubre 2026) ---
    $cvS = _SettingsAmbClau ([pscustomobject]@{ CorreuVia = 'outlook'; InformesDir = 'F:\x' }) 'CorreuRemitent' ' activitats@x.cat '
    AssertEq $cvS['CorreuRemitent'] 'activitats@x.cat' 'remitent: es desa net'
    AssertEq $cvS['CorreuVia'] 'outlook' 'remitent: no trepitja la via'
    AssertEq $cvS['InformesDir'] 'F:\x' 'remitent: no trepitja les carpetes'
    Assert (-not (_SettingsAmbClau $cvS 'CorreuRemitent' '').Contains('CorreuRemitent')) 'remitent buit -> no s''escriu (el compte per defecte)'
    Assert (_CorreuRemitentValid '') 'remitent buit es valid (el per defecte)'
    Assert (_CorreuRemitentValid 'activitats@cornella.cat') 'remitent: una adreca'
    Assert (-not (_CorreuRemitentValid 'activitats')) 'remitent: sense @ no'
    Assert (-not (_CorreuRemitentValid 'a@x.cat; b@x.cat')) 'remitent: nomes UNA adreca'
    AssertEq (_CorreuRemitentText 'emailjs' 'a@x.cat') '' 'remitent: amb EmailJS no es diu (no compta)'
    AssertEq (_CorreuRemitentText 'outlook' 'a@x.cat') 'des de a@x.cat' 'remitent: amb l''Outlook es diu'
    Assert ((_CorreuRemitentText 'outlook-esborrany' '') -like '*per defecte*') 'remitent buit: el compte per defecte'

    # Un compte de l'Outlook -> SendUsingAccount; una altra adreca (bustia
    # compartida) -> SentOnBehalfOfName.
    $cvCompte = [pscustomobject]@{ SmtpAddress = 'Activitats@X.cat' }
    $cvOl2 = [pscustomobject]@{ Session = [pscustomobject]@{ Accounts = @([pscustomobject]@{ SmtpAddress = 'jo@x.cat' }, $cvCompte) } }
    Assert ([object]::ReferenceEquals((_OutlookTriaCompte $cvOl2 'activitats@x.cat'), $cvCompte)) 'remitent: troba el compte (sense distingir majuscules)'
    Assert ($null -eq (_OutlookTriaCompte $cvOl2 'compartida@x.cat')) 'remitent: una bustia compartida no es un compte'
    Assert ($null -eq (_OutlookTriaCompte $cvOl2 '')) 'remitent buit: cap compte'
    $cvM = [pscustomobject]@{ SendUsingAccount = $null; SentOnBehalfOfName = '' }
    _OutlookPosaRemitent $cvM $cvCompte 'activitats@x.cat'
    Assert ([object]::ReferenceEquals($cvM.SendUsingAccount, $cvCompte)) 'remitent: compte -> SendUsingAccount'
    AssertEq $cvM.SentOnBehalfOfName '' 'remitent: compte -> no en nom d''altri'
    $cvM = [pscustomobject]@{ SendUsingAccount = $null; SentOnBehalfOfName = '' }
    _OutlookPosaRemitent $cvM $null 'compartida@x.cat'
    AssertEq $cvM.SentOnBehalfOfName 'compartida@x.cat' 'remitent: bustia compartida -> SentOnBehalfOfName'
    $cvM = [pscustomobject]@{ SendUsingAccount = $null; SentOnBehalfOfName = '' }
    _OutlookPosaRemitent $cvM $null ''
    Assert ($null -eq $cvM.SendUsingAccount -and $cvM.SentOnBehalfOfName -eq '') 'remitent buit: no es toca res'

    # La sessio hi posa el remitent i el CC (Controls periodics) a cada correu.
    $cvLog.Clear()
    $cvNouItem = {
        $it = [pscustomobject]@{ To = ''; CC = ''; BCC = ''; Subject = ''; HTMLBody = ''; SendUsingAccount = $null; SentOnBehalfOfName = '' }
        $it | Add-Member ScriptMethod Save { [void]$cvLog.Add('save:' + $this.To + '|cc=' + $this.CC + '|de=' + $this.SentOnBehalfOfName) }
        $it | Add-Member ScriptMethod Send { [void]$cvLog.Add('send:' + $this.To) }
        return $it
    }
    $cvSes = @{ Via = 'outlook-esborrany'; Cfg = $null; Outlook = $cvOl; JaObert = $true; Desats = 0; Enviats = 0; Remitent = 'compartida@x.cat'; Compte = $null }
    Send-CorreuSessio $cvSes 'titular@x.cat' '' 'Assumpte' '<p>hola</p>' 'rep@x.cat'
    AssertEq ($cvLog -join ',') 'save:titular@x.cat|cc=rep@x.cat|de=compartida@x.cat' 'sessio: remitent i CC a l''esborrany'

    # Tancar: si l'Outlook l'hem obert nosaltres, envia i espera la Safata de sortida.
    $cvLog.Clear()
    $cvSessio = [pscustomobject]@{}
    $cvSessio | Add-Member ScriptMethod SendAndReceive { param($b) [void]$cvLog.Add('sendreceive') }
    $cvSessio | Add-Member ScriptMethod GetDefaultFolder { param($n) [void]$cvLog.Add("carpeta:$n"); return [pscustomobject]@{ Items = [pscustomobject]@{ Count = 0 } } }
    $cvSes = @{ Via = 'outlook'; Outlook = [pscustomobject]@{ Session = $cvSessio }; JaObert = $false; Enviats = 1; Compte = $null }
    Close-CorreuSessio $cvSes 2
    AssertEq ($cvLog -join ',') 'sendreceive,carpeta:4' 'tancar: envia i mira la Safata de sortida'
    Assert ($null -eq $cvSes.Outlook) 'tancar: allibera l''Outlook'
    $cvLog.Clear()
    $cvSes = @{ Via = 'outlook-esborrany'; Outlook = [pscustomobject]@{ Session = $cvSessio }; JaObert = $false; Enviats = 0; Compte = $null }
    Close-CorreuSessio $cvSes 2
    AssertEq ($cvLog -join ',') '' 'tancar: nomes esborranys -> no envia res'

    # --- Els esborranys pendents (els de la tasca automatica) ---
    $cvQuan = [datetime]'2026-10-08 13:00'
    $cvP = _EsborranysAmbNous @() 'Recordatoris (Requeriments)' @('GIA 1 - A (a@x.cat)', 'GIA 2 - B (b@x.cat)') $cvQuan
    AssertEq @($cvP).Count 1 'esborranys: una tanda'
    AssertEq @($cvP[0].correus).Count 2 'esborranys: els dos correus'
    AssertEq @(_EsborranysAmbNous $cvP 'x' @() $cvQuan).Count 1 'esborranys: una tanda buida no s''apunta'
    $cvP2 = _EsborranysAmbNous $cvP 'Recordatoris (Precintes)' @('GIA 3 - C (c@x.cat)') $cvQuan.AddDays(1)
    AssertEq @($cvP2).Count 2 'esborranys: s''acumulen fins que es diu que s''han enviat'
    $cvTxt = _EsborranysAvisText $cvP2
    Assert ($cvTxt.Contains('Hi ha 3 correus')) 'avis esborranys: quants'
    Assert ($cvTxt.Contains('2026-10-09 13:00 - Recordatoris (Precintes): 1')) 'avis esborranys: quan i de quina campanya'
    Assert ($cvTxt.Contains('GIA 2 - B (b@x.cat)')) 'avis esborranys: a qui'

    # Anada i tornada AMB EL JSON PEL MIG (on aquest projecte s'ha trencat sempre).
    $cvTmp = Join-Path ([System.IO.Path]::GetTempPath()) ('cv-' + [guid]::NewGuid().ToString('N'))
    $cvAbans = $env:LOCALAPPDATA
    try {
        $env:LOCALAPPDATA = $cvTmp
        AssertEq (Get-CorreuEsborranysPendents).Count 0 'esborranys: sense fitxer, cap'
        Add-CorreuEsborranysPendents 'Recordatoris (Requeriments)' @('GIA 1 - A (a@x.cat)')
        Add-CorreuEsborranysPendents 'Recordatoris (Precintes)' @('GIA 2 - B (b@x.cat)', 'GIA 3 - C (c@x.cat)')
        $cvG = Get-CorreuEsborranysPendents
        AssertEq $cvG.Count 2 'esborranys (JSON): les dues tandes'
        AssertEq @($cvG[1].correus).Count 2 'esborranys (JSON): els correus de la segona'
        AssertEq @($cvG[0].correus).Count 1 'esborranys (JSON): una tanda d''UN correu segueix sent una llista'
        Clear-CorreuEsborranysPendents
        AssertEq (Get-CorreuEsborranysPendents).Count 0 'esborranys: "ja els he enviat" les buida'
    } finally {
        if ($null -eq $cvAbans) { Remove-Item Env:LOCALAPPDATA -ErrorAction SilentlyContinue } else { $env:LOCALAPPDATA = $cvAbans }
        Remove-Item -LiteralPath $cvTmp -Recurse -Force -ErrorAction SilentlyContinue
    }
} catch {
    Assert $false ("bloc CorreuVia: excepcio " + $_.Exception.Message)
}
