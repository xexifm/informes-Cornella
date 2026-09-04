# Composicio del document, textos fixos i base de llicencies
#
# Es DOT-SOURCE des de run-tests.ps1: mateix ambit, mateixes variables i el
# mateix comptador d'asserts. No s'executa sol.

Write-Host "`n--- SeguimentGia.ps1: els cinc llistats del GIA (funcions pures) ---"
# Taula d'exemple amb la MATEIXA forma que la fulla Estes de la base de dades:
# capceleres amb 'Classificacio general annex' REPETIDA (a la base hi surt dues
# vegades: la 1a es la que es MOSTRA, la 2a la que fa de CRITERI a ANNEX II) i
# dues parelles de Camp Info.
$sgH = @(
    'ID Activitat', ('N' + [char]0x00FA + 'm. expedient '), ('Ra' + [char]0x00F3 + ' social'),
    ('Ra' + [char]0x00F3 + ' soc. M' + [char]0x00F2 + 'bil'), 'Representant legal',
    ('Rep. Leg. M' + [char]0x00F2 + 'bil'), 'Nom comercial activitat ',
    'Emp. Tipus via', 'Emp. Carrer', ('Emp. N' + [char]0x00FA + 'mero'),
    ('Classificaci' + [char]0x00F3 + ' general annex'), ('Classificaci' + [char]0x00F3 + ' general Apartat'),
    ('Data control inicial/verificaci' + [char]0x00F3), ('Data control peri' + [char]0x00F2 + 'dic'),
    'Activitat principal', ('Classificaci' + [char]0x00F3 + ' general annex'), ('Descripci' + [char]0x00F3 + ' lliure'),
    'Camp Info 1 - Nom', 'Camp Info 1 - Valor', 'Camp Info 2 - Nom', 'Camp Info 2 - Valor'
)
# SENSE @(): _FindCampInfoPairs protegeix el seu retorn amb una coma (per al cas
# d'UNA sola parella) i un @() al voltant hi torna a posar la capa. Aixo es
# exactament el que feia petar l'eina Seguiment amb "No se puede convertir el
# valor System.Object[] ... al tipo System.Int32" (vegeu _SgAplanaPairs).
$sgPairs = _FindCampInfoPairs $sgH
AssertEq @($sgPairs).Count 2 'SeguimentGia: es reaprofita _FindCampInfoPairs (2 parelles)'
AssertEq @(_FindCampInfoPairs $sgH | Select-Object -First 1).Count 1 '_FindCampInfoPairs: el retorn es una llista de parelles, no una llista d''una llista'
# La forma EMBOLCALLADA (com si algu hi tornes a posar un @()) i la neta han de
# donar el mateix: _SgAplanaPairs ho aplana al punt d'entrada.
$sgPairsEmbolcallat = @(_FindCampInfoPairs $sgH)
AssertEq @($sgPairsEmbolcallat).Count 1 '_FindCampInfoPairs amb @(): queda embolcallat (per aixo no s''hi posa)'
AssertEq @(_SgAplanaPairs $sgPairs).Count 2 '_SgAplanaPairs: la forma neta es queda igual'
AssertEq @(_SgAplanaPairs $sgPairsEmbolcallat).Count 2 '_SgAplanaPairs: la forma embolcallada s''aplana'
AssertEq @(_SgAplanaPairs $null).Count 0 '_SgAplanaPairs: $null -> llista buida'
AssertEq @(_SgAplanaPairs @()).Count 0 '_SgAplanaPairs: llista buida -> llista buida'
AssertEq ([int](@(_SgAplanaPairs $sgPairsEmbolcallat)[0].NomCol)) 18 '_SgAplanaPairs: els elements son parelles de debo (NomCol es un enter)'
AssertEq (_SgColIndex $sgH 'ID Activitat') 1 '_SgColIndex: 1-based'
AssertEq (_SgColIndex $sgH ('classificacio general annex')) 11 '_SgColIndex: la PRIMERA que coincideix (com el MATCH de la plantilla)'
AssertEq (_SgColIndex $sgH 'no existeix') 0 '_SgColIndex: si no hi es -> 0'
#            1      2         3        4       5      6      7     8     9    10    11   12       13           14           15      16   17        18                      19                20              21
# ATENCIO als PARENTESIS de ('DEN' + [char]0x00DA + 'NCIA?'): dins d'un literal
# @(...) la COMA lliga mes fort que el '+', o sigui que sense parentesis
# 'DEN' + [char]0x00DA + 'NCIA?' es converteix en TRES elements ('DEN', 'U',
# 'NCIA?') i desalinea tota la fila. Va passar: la prova de DENUNCIES no trobava
# res perque la fila tenia 23 columnes en lloc de 21.
$sgFiles = @(
  @('100','E-1','TITULAR U','600','REP U','601','BAR U','C','MAJOR','1','II','5.17.b','2019-03-15 00:00:00.0','2025-07-08 00:00:00.0','BAR','II','Text llarg de descripcio','PRECINTE ACTIVITAT?','SI, precintada',('DEN' + [char]0x00DA + 'NCIA?'),'SI, soroll'),
  @('101','E-2','TITULAR D','602','','','BAR D','AV','PICASSO','6','III','','','','REST','III','','SONOMETRIA?','SI, queixa','',''),
  @('102','E-3','TITULAR T','603','','','BAR T','PG','FERRO','24','II','1.1','','','IND','II','','REQUERIT PER DECRET?','PROCEDIMENT ESMENA','',''),
  @('','E-4','SENSE ID','','','','','','','','II','','','','','II','hi ha descripcio','PRECINTE ACTIVITAT?','SI',''),
  @('103','E-5','TITULAR C','604','','','BAR C','C','LLUNA','9','II','2.2','','','COM','II','','','','','')
)
$sgDefs = @(_SgFullesDef)
AssertEq $sgDefs.Count 5 '_SgFullesDef: 5 pestanyes'

# --- Tria de pestanyes (les caselles de la finestra) ---------------------------
# La llista de caselles i el que construeix el llibre surten de la MATEIXA font,
# _SgOpcionsExport, per no poder-se desincronitzar.
$sgOps = @(_SgOpcionsExport)
AssertEq $sgOps.Count 6 '_SgOpcionsExport: la copia de la base + els 5 llistats'
AssertEq $sgOps[0] (_SgNomEstes) '_SgOpcionsExport: la copia de la base va primer'
AssertEq $sgOps[1] 'PRECINTES'   '_SgOpcionsExport: despres els llistats, en ordre'
AssertEq $sgOps[5] 'ANNEX II'    '_SgOpcionsExport: l ultim es ANNEX II'
Assert (_SgSeleccioTeEstes @((_SgNomEstes), 'PRECINTES')) '_SgSeleccioTeEstes: la troba'
Assert (_SgSeleccioTeEstes @('estes'))                    '_SgSeleccioTeEstes: sense accents i en minuscules tambe'
Assert (-not (_SgSeleccioTeEstes @('PRECINTES')))         '_SgSeleccioTeEstes: si no hi es, no'
Assert (-not (_SgSeleccioTeEstes @()))                    '_SgSeleccioTeEstes: seleccio buida'
AssertEq @(_SgFullesTriades @('PRECINTES', 'ANNEX II')).Count 2 '_SgFullesTriades: nomes les triades'
AssertEq @(_SgFullesTriades @('PRECINTES', 'ANNEX II'))[0].Nom 'PRECINTES' '_SgFullesTriades: en l ordre de _SgFullesDef, no en el de la tria'
AssertEq @(_SgFullesTriades @('ANNEX II', 'PRECINTES'))[0].Nom 'PRECINTES' '_SgFullesTriades: l ordre de la tria no compta'
AssertEq @(_SgFullesTriades @((_SgNomEstes))).Count 0 '_SgFullesTriades: la copia de la base no es cap llistat'
AssertEq @(_SgFullesTriades @()).Count 0              '_SgFullesTriades: res triat -> cap llistat'
AssertEq @(_SgFullesTriades $null).Count 0            '_SgFullesTriades: $null -> cap llistat'
AssertEq @(_SgFullesTriades $sgOps).Count 5           '_SgFullesTriades: totes marcades -> els 5 llistats'
AssertEq @(_SgFullesTriades @('PRECINTES')).Count 1   '_SgFullesTriades: una sola tria no es desenrotlla'
AssertEq ($sgDefs[0].Nom) 'PRECINTES' '_SgFullesDef: la primera es PRECINTES'
$sgPre = @(_SgFilesPerFulla $sgDefs[0] $sgH $sgFiles $sgPairs)
AssertEq $sgPre.Count 1 'PRECINTES: nomes l activitat 100 (la fila sense ID Activitat NO compta)'
AssertEq ($sgPre[0][0]) 1 'PRECINTES: la columna N numera 1..N (no es el numero de fila de l Excel)'
AssertEq ($sgPre[0][1]) '100' 'PRECINTES: ID Activitat'
AssertEq ($sgPre[0][11]) 'PRECINTE ACTIVITAT?' 'PRECINTES: penultima columna = el Camp Info que ha coincidit'
AssertEq ($sgPre[0][12]) 'SI, precintada' 'PRECINTES: ultima columna = el seu valor'
# El criteri es NOMES tenir el camp: el valor pot dir qualsevol cosa.
$sgReq = @(_SgFilesPerFulla $sgDefs[2] $sgH $sgFiles $sgPairs)
AssertEq $sgReq.Count 1 'REQUERIT DECRET: hi entra tot i que el valor no comenci per SI'
AssertEq ($sgReq[0][12]) 'PROCEDIMENT ESMENA' 'REQUERIT DECRET: el valor surt tal qual (aqui NO es demana SI)'
# La 2a parella de Camp Info tambe val (a la plantilla el MATCH mirava tota la fila).
$sgDen = @(_SgFilesPerFulla $sgDefs[1] $sgH $sgFiles $sgPairs)
AssertEq $sgDen.Count 1 ('DEN' + [char]0x00DA + 'NCIES: troba el criteri a la 2a parella de Camp Info')
AssertEq ($sgDen[0][12]) 'SI, soroll' ('DEN' + [char]0x00DA + 'NCIES: valor de la 2a parella')
$sgSon = @(_SgFilesPerFulla $sgDefs[3] $sgH $sgFiles $sgPairs)
AssertEq $sgSon.Count 1 'SONOMETRIA: nomes l activitat 101'
# ANNEX II: classificacio II (2a columna repetida) I descripcio lliure amb text.
$sgAnx = @(_SgFilesPerFulla $sgDefs[4] $sgH $sgFiles $sgPairs)
AssertEq $sgAnx.Count 1 'ANNEX II: nomes la 100 (la 102 i la 103 son II pero sense descripcio; la 101 es III)'
AssertEq ($sgAnx[0][11]) 'II' 'ANNEX II: mostra la PRIMERA columna de classificacio'
AssertEq ($sgAnx[0][13]) '15/03/2019' 'ANNEX II: la data surt com a dd/MM/aaaa, no en cru'
AssertEq ($sgAnx[0][14]) '08/07/2025' 'ANNEX II: la 2a data, igual'
AssertEq ($sgAnx[0][16]) 'Text llarg de descripcio' ('ANNEX II: ultima columna = Descripci' + [char]0x00F3 + ' lliure')
AssertEq ([bool](_SgAnnexCoincideix 'II' 'text')) $true '_SgAnnexCoincideix: II + descripcio -> si'
AssertEq ([bool](_SgAnnexCoincideix 'II' '')) $false '_SgAnnexCoincideix: sense descripcio -> no'
AssertEq ([bool](_SgAnnexCoincideix 'III' 'text')) $false '_SgAnnexCoincideix: annex III -> no'
AssertEq ([bool](_SgAnnexCoincideix 'ii' '  x ')) $true '_SgAnnexCoincideix: no distingeix majuscules'
# Titols i nom del fitxer
$sgData = Get-Date '2026-07-30'
AssertEq (_SgTitolFulla $sgDefs[0] $sgData) 'ACTIVITAT PRECINTADA? 30/07/2026' '_SgTitolFulla: titol + data de la fila 1'
AssertEq (_SgTitolFulla $sgDefs[4] $sgData) 'ANNEXOS II 30/07/2026' '_SgTitolFulla: ANNEX II te el seu propi titol'
AssertEq (_SgNomFitxer $sgData) '2026-07-30 Seguiment GIA.xlsx' '_SgNomFitxer: xlsx per defecte'
AssertEq (_SgNomFitxer $sgData 'pdf') '2026-07-30 Seguiment GIA.pdf' '_SgNomFitxer: pdf'
# Missatge d'error: ha de dir el PAS i la LINIA, no nomes el missatge pelat. Amb
# COM, un "Unable to cast object of type X to type Y" no diu on ha estat, i
# aquesta eina nomes es pot provar amb l'Excel de debo.
$sgErr = $null
try { throw 'petada de prova' } catch { $sgErr = $_ }
$sgTxt = _SgTextError $sgErr "bolcant les dades a 'PRECINTES'"
Assert ($sgTxt -like '*petada de prova*')                 '_SgTextError: hi surt el missatge original'
Assert ($sgTxt -like "*bolcant les dades a 'PRECINTES'*") '_SgTextError: hi surt el pas on estavem'
Assert ($sgTxt -like '*SeguimentGia.ps1, linia*')         '_SgTextError: hi surt la linia del codi'
Assert ((_SgTextError $sgErr '') -like '*petada de prova*') '_SgTextError: sense pas, encara dona el missatge'

# _SgEscriuMatriu: si l'escriptura EN BLOC falla, ha d'escriure cel·la a cel·la i
# deixar les dades igualment. Amb l'Excel de l'usuari l'assignacio en bloc peta
# ("Unable to cast object of type 'System.Object[,]' to type 'System.String'"),
# aixi que aquest cami alternatiu es el que ha de salvar el fitxer.
# La fulla falsa no te InvokeMember de COM, o sigui que el bloc falla sol i es
# prova exactament el que interessa: el respatller.
$sgEscrites = @{}
$sgFullaFalsa = [pscustomobject]@{}
$sgFullaFalsa | Add-Member -MemberType ScriptProperty -Name Cells -Value {
    [pscustomobject]@{} | Add-Member -MemberType ScriptMethod -Name Item -Value {
        param($r, $c)
        $cel = [pscustomobject]@{ R = $r; C = $c }
        $cel | Add-Member -MemberType ScriptProperty -Name Value2 -Value { '' } `
                          -SecondValue { param($v) $script:sgEscrites["$($this.R),$($this.C)"] = $v } -PassThru
    } -PassThru
}
$sgFullaFalsa | Add-Member -MemberType ScriptMethod -Name Range -Value { param($a, $b) throw 'sense COM' }
$sgMat = New-Object 'object[,]' 2, 3
$sgMat[0, 0] = 'a'; $sgMat[0, 1] = 'b'; $sgMat[0, 2] = 'c'
$sgMat[1, 0] = 'd'; $sgMat[1, 1] = 'e'; $sgMat[1, 2] = 'f'
$sgVia = _SgEscriuMatriu $sgFullaFalsa 3 $sgMat
AssertEq $sgVia ('cel' + [char]0x00B7 + 'la') '_SgEscriuMatriu: si el bloc falla, passa a cel·la a cel·la'
AssertEq $sgEscrites.Count 6      '_SgEscriuMatriu: escriu totes les cel·les de la matriu'
AssertEq $sgEscrites['3,1'] 'a'   '_SgEscriuMatriu: la matriu comenca a la fila indicada, columna 1'
AssertEq $sgEscrites['3,3'] 'c'   '_SgEscriuMatriu: ultima columna de la 1a fila'
AssertEq $sgEscrites['4,1'] 'd'   '_SgEscriuMatriu: 2a fila'
AssertEq $sgEscrites['4,3'] 'f'   '_SgEscriuMatriu: ultima cel·la'
# A l'Excel de l'usuari, Value2 nomes accepta CADENES: un Int32 peta igual que
# la matriu. Per aixo tot el que s'escriu passa per _SgValorCella.
AssertEq (_SgValorCella 1) '1'            '_SgValorCella: un enter surt com a text (Value2 nomes accepta cadenes)'
AssertEq (_SgValorCella $null) ''         '_SgValorCella: $null -> cadena buida'
AssertEq (_SgValorCella '') ''            '_SgValorCella: buit -> buit'
AssertEq (_SgValorCella 'SI, precintat') 'SI, precintat' '_SgValorCella: el text no es toca'
Assert ((_SgValorCella 1) -is [string])   '_SgValorCella: retorna SEMPRE un String'
Assert ((_SgValorCella $null) -is [string]) '_SgValorCella: fins i tot amb $null'
# La columna N de veritat es un enter: es el cas que va fer petar l'eina.
$sgMatN = New-Object 'object[,]' 1, 2
$sgMatN[0, 0] = 1; $sgMatN[0, 1] = 'TEXT'
$sgEscrites = @{}
[void](_SgEscriuMatriu $sgFullaFalsa 3 $sgMatN)
AssertEq $sgEscrites['3,1'] '1' '_SgEscriuMatriu: la columna N (enter) s''escriu com a text'
Assert ($sgEscrites['3,1'] -is [string]) '_SgEscriuMatriu: i arriba a la cel·la com a String'
# Si CAP de les tres maneres funciona, ha de petar dient que ha passat a cada una.
$sgFullaMorta = [pscustomobject]@{}
$sgFullaMorta | Add-Member -MemberType ScriptProperty -Name Cells -Value { throw 'sense Cells' }
$sgFullaMorta | Add-Member -MemberType ScriptMethod -Name Range -Value { param($a, $b) throw 'sense Range' }
$sgErrTot = ''
try { [void](_SgEscriuMatriu $sgFullaMorta 3 $sgMat) } catch { $sgErrTot = [string]$_.Exception.Message }
Assert ($sgErrTot -like '*de cap de les tres maneres*') '_SgEscriuMatriu: si tot falla, peta (no es queda mut)'
Assert ($sgErrTot -like '*bloc-invoke:*')               '_SgEscriuMatriu: l''error diu que ha passat a cada intent'
# Amplades: tantes com columnes (N + les de dades).
foreach ($d in $sgDefs) {
    AssertEq (@($d.Amplades).Count) ((@($d.Cols).Count) + 1) ("Amplades de " + $d.Nom + ": una per columna, incloent-hi la 'N'")
}

Write-Host "`n--- Cap dada d'una activitat CREMADA al cataleg de Llicencia ---"
# Al LLIC.json hi havia sis expedients i referencies d'UNA activitat concreta
# ("Expedient: FUE-2023-03018882", "Referencia 24/2022/000104"...), que sortien
# a l'informe de TOTHOM. Han de ser camps. Aquesta prova vigila tambe el que
# s'hi pugui tornar a escriure des de l'editor de catalegs.
$llicPathC = Join-Path $EstructuralsDir 'LLIC.json'
if (Test-Path -LiteralPath $llicPathC) {
    $llicC = Read-LlicCataleg $llicPathC
    $cremades = New-Object System.Collections.ArrayList
    foreach ($secC in @($llicC.nodes)) {
        foreach ($itC in @($secC.fills)) {
            $bC = _LlicFill $itC 'sidisposa'
            if ($null -eq $bC) { continue }
            foreach ($lnC in @(_LlicCos $bC)) {
                # Les dades van al PARENTESI del final: "(Id Firmadoc: X; Expedient: Y)".
                $parC = @([regex]::Matches([string]$lnC, '\(([^)]*)\)'))
                if ($parC.Count -eq 0) { continue }
                foreach ($pecaC in (([string]$parC[$parC.Count - 1].Groups[1].Value) -split ';')) {
                    # Trec els [CAMP: ...]; el que quedi despres dels dos punts
                    # ha de ser NOMES espais. Si hi queda text, es un valor cremat.
                    $netC = $pecaC -replace '\[CAMP:[^\]]*\]', ''
                    $iC = $netC.IndexOf(':')
                    $restaC = if ($iC -ge 0) { $netC.Substring($iC + 1) } else { $netC }
                    if ([string]::IsNullOrWhiteSpace($restaC)) { continue }
                    [void]$cremades.Add(([string]$itC.titol) + ' -> ' + $pecaC.Trim())
                }
            }
        }
    }
    AssertEq $cremades.Count 0 ('LLIC.json: cap expedient/referencia d''una activitat concreta (' + ($cremades -join ' | ') + ')')
}

Write-Host "`n--- Tots els punts d'ABANS poden dir que es tenen ---"
# La llista d'ABANS surt de les 4 seccions de REQ1 (mes de 40 punts) pero LLIC
# nomes en descriu 15: als altres, triar "Es disposa" no ensenyava res i a
# l'informe no s'hi escrivia res.
$defT = _LlicTextosPerDefecte
Assert ([bool](@($defT.NoDisposa).Count -gt 0)) '_LlicTextosPerDefecte: hi ha el "No es disposa"'
Assert ([bool](@($defT.SiDisposa) -join ' ') -like '*[CAMP: Id Firmadoc]*') '_LlicTextosPerDefecte: el "Es disposa" demana l''Id Firmadoc'
if ((Test-Path -LiteralPath $llicPathC) -and (Test-Path -LiteralPath (Join-Path $EstructuralsDir 'REQ1.json'))) {
    $req1D = Read-CatalegJson (Join-Path $EstructuralsDir 'REQ1.json')
    $llicD = Read-LlicCataleg $llicPathC
    $pAb = @((_LlicPuntsPerBloc $llicD (_LlicIndexReq1 $req1D) 'ABANS' $req1D).Punts)
    Assert ([bool]($pAb.Count -gt 25)) 'ABANS: hi ha tots els punts de les 4 seccions'
    AssertEq (@($pAb | Where-Object { @($_.NoDisposa).Count -eq 0 }).Count) 0 'ABANS: cap punt sense "No es disposa"'
    AssertEq (@($pAb | Where-Object { @($_.SiDisposa).Count -eq 0 }).Count) 0 'ABANS: cap punt sense "Es disposa"'
    AssertEq (@($pAb | Where-Object { (@($_.SiDisposa) -join ' ') -notmatch 'Id Firmadoc' }).Count) 0 'ABANS: tots demanen l''Id Firmadoc'
    # ...i els que SI que consten a LLIC conserven la seva redaccio.
    $pSanD = @($pAb | Where-Object { $_.Titol -eq 'Sanitat' })[0]
    Assert ([bool]((@($pSanD.SiDisposa) -join ' ') -like '*autoritzaci*')) 'ABANS: el text propi de LLIC mana sobre el de defecte'
}

Write-Host "`n--- El text del cos, sense el marcador [[URL]] ---"
$cosU = @('Zona inundable. Cal l''informe.', '[[URL]] https://aca.gencat.cat/x')
$plaU = _LlicTextPlaDelCos $cosU
Assert ($plaU -notmatch '\[\[URL\]\]') '_LlicTextPlaDelCos: el marcador intern no es veu'
Assert ($plaU.Contains('https://aca.gencat.cat/x')) '_LlicTextPlaDelCos: ...pero l''adreca si'
AssertEq (_LlicTextPlaDelCos @()) '' '_LlicTextPlaDelCos: sense cos, buit'

Write-Host "`n--- Base de dades de llicencies ---"
# Insercio, substitucio per ID GIA i esborrat.
$dbT = New-LlicenciaDb
AssertEq (@($dbT.Llicencies).Count) 0 'Base nova: buida'
Assert ($null -eq (Get-LlicenciaRecord $dbT '1463')) 'Base buida: no troba res'
[void](Set-LlicenciaRecord $dbT ([ordered]@{ IdGia = '1463'; Titular = 'A' }))
[void](Set-LlicenciaRecord $dbT ([ordered]@{ IdGia = '999';  Titular = 'B' }))
AssertEq (@($dbT.Llicencies).Count) 2 'Set-LlicenciaRecord: dues fitxes'
[void](Set-LlicenciaRecord $dbT ([ordered]@{ IdGia = '1463'; Titular = 'A bis' }))
AssertEq (@($dbT.Llicencies).Count) 2 'Set-LlicenciaRecord: el mateix GIA NO duplica'
AssertEq ([string](Get-LlicenciaRecord $dbT '1463').Titular) 'A bis' 'Set-LlicenciaRecord: substitueix'
AssertEq ([string](Get-LlicenciaRecord $dbT '999').Titular) 'B' 'Set-LlicenciaRecord: ...i no toca les altres'
[void](Remove-LlicenciaRecord $dbT '1463')
AssertEq (@($dbT.Llicencies).Count) 1 'Remove-LlicenciaRecord: en treu una'
Assert ($null -eq (Get-LlicenciaRecord $dbT '1463')) 'Remove-LlicenciaRecord: ...la seva'
Assert ($null -ne (Get-LlicenciaRecord $dbT '999'))  'Remove-LlicenciaRecord: ...i nomes la seva'

# ANADA I TORNADA amb ConvertTo-Json pel mig, que es com viura de debo: el JSON
# torna PSCustomObjects i les claus numeriques dels sub-punts tornen com a text.
$stT = @{
    Fase = 'requeriment'; Prov = $true; Condicions = 'les de sempre'
    Header = @{ ID_GIA = '1463'; TITULAR = 'ZEROCATORZE'; ADRECA = 'CAMI 12'; ACTIVITAT = 'TALLER'; CLASSIFICACIO = 'Llei 20/2009' }
    MemAbans = @{
        'Autoritzacions / Informes preceptius::Sanitat' = @{ Marcat = $true;  Estat = 'si'; Valors = @{ 'Id Firmadoc' = 'FD-777' }; Subs = @{} }
        'Registres::RASIC'                              = @{ Marcat = $false; Estat = 'no'; Valors = @{}; Subs = @{} }
    }
    MemDespres = @{ '#PAU' = @{ Marcat = $true; Estat = 'no'; Valors = @{}; Subs = @{ 0 = $true; 1 = $false } } }
    ProjKeys = @('Projecte::A', 'Projecte::B'); ProjVals = @{ 'Epigraf' = '12.3' }
    Tecnic = @{ Tecnic = 'Nom'; NumCol = '123' }
}
$recT = ConvertTo-LlicenciaRecord $stT @((New-LlicenciaHistorial 'requeriment' 'C:\a.docx'))
AssertEq ([string]$recT.IdGia) '1463' 'ConvertTo-LlicenciaRecord: la clau es l''ID GIA'
AssertEq ([string]$recT.Titular) 'ZEROCATORZE' 'ConvertTo-LlicenciaRecord: el titular, per llistar-la'
AssertEq (@($recT.Historial).Count) 1 'ConvertTo-LlicenciaRecord: l''historial'
# El pas per JSON (i tornada), que es el que trenca les coses.
$recJ = ($recT | ConvertTo-Json -Depth 20) | ConvertFrom-Json
$stR = @{ Fase = 'x'; Prov = $false; Header = @{ ID_GIA = '1463' } }
[void](Restore-LlicenciaState $recJ $stR)
AssertEq ([bool]$stR.MemAbans['Autoritzacions / Informes preceptius::Sanitat'].Marcat) $true 'anada i tornada: el marcat'
AssertEq ([string]$stR.MemAbans['Autoritzacions / Informes preceptius::Sanitat'].Estat) 'si' 'anada i tornada: l''estat'
AssertEq ([string]$stR.MemAbans['Autoritzacions / Informes preceptius::Sanitat'].Valors['Id Firmadoc']) 'FD-777' 'anada i tornada: l''Id Firmadoc'
AssertEq ([bool]$stR.MemAbans['Registres::RASIC'].Marcat) $false 'anada i tornada: el NO marcat tambe es recorda'
AssertEq ([bool]$stR.MemDespres['#PAU'].Subs[0]) $true  'anada i tornada: el sub-punt triat'
AssertEq ([bool]$stR.MemDespres['#PAU'].Subs[1]) $false 'anada i tornada: i el no triat'
Assert ((@($stR.MemDespres['#PAU'].Subs.Keys)[0]) -is [int]) 'anada i tornada: les claus dels sub-punts tornen a ser NUMERIQUES'
AssertEq (@($stR.ProjKeys) -join '|') 'Projecte::A|Projecte::B' 'anada i tornada: els punts del projecte'
AssertEq ([string]$stR.ProjVals['Epigraf']) '12.3' 'anada i tornada: els camps del projecte'
AssertEq ([string]$stR.Tecnic['Tecnic']) 'Nom' 'anada i tornada: el tecnic redactor'
AssertEq ([string]$stR.Condicions) 'les de sempre' 'anada i tornada: les condicions'
# La capcalera NO es toca: l'omple l'Excel per ID GIA i la de la base pot ser vella.
AssertEq ([string]$stR.Header['ID_GIA']) '1463' 'Restore-LlicenciaState: no toca la capcalera'
# El resum per a la llista.
$resT = Get-LlicenciaResum $recJ
AssertEq ([string]$resT.IdGia) '1463' 'Get-LlicenciaResum: ID GIA'
AssertEq ([int]$resT.Punts) 2 'Get-LlicenciaResum: nomes compta els punts MARCATS'
AssertEq ([int]$resT.Informes) 1 'Get-LlicenciaResum: els informes fets'
Assert (-not [string]::IsNullOrWhiteSpace([string]$resT.Data)) 'Get-LlicenciaResum: la data, llegible'

# Una fitxa mig buida no pot petar (una base vella, o feta a ma).
$stB = @{ Header = @{ ID_GIA = '7' } }
$recB = ConvertTo-LlicenciaRecord $stB
AssertEq ([string]$recB.IdGia) '7' 'ConvertTo-LlicenciaRecord: sense memoria, no peta'
$stB2 = @{}
[void](Restore-LlicenciaState (($recB | ConvertTo-Json -Depth 20) | ConvertFrom-Json) $stB2)
AssertEq (@($stB2.MemAbans.Keys).Count) 0 'Restore-LlicenciaState: sense memoria, memoria buida'
[void](Restore-LlicenciaState $null $stB2) | Out-Null
Assert $true 'Restore-LlicenciaState: amb $null tampoc peta'

# EL CAMI SENCER: base de dades -> memoria -> text de l'informe. Els valors
# recuperats han de SUBSTITUIR els [CAMP: ...] al document.
$valsR = $stR.MemAbans['Autoritzacions / Informes preceptius::Sanitat'].Valors
$liniesR = @(_LlicAplicaCamps @('Es disposa de l''autoritzacio (Id Firmadoc: [CAMP: Id Firmadoc])') $valsR)
Assert ($liniesR[0] -notmatch '\[CAMP:') 'De la base a l''informe: cap marcador literal'
Assert ($liniesR[0].Contains('FD-777')) 'De la base a l''informe: hi surt el valor recuperat'
# ...i un punt del qual la base no en sap res queda buit, no amb el marcador.
$liniesR2 = @(_LlicAplicaCamps @('Es disposa del document (Id Firmadoc: [CAMP: Id Firmadoc])') @{})
Assert ($liniesR2[0] -notmatch '\[CAMP:') 'Sense valor a la base: tampoc queda el marcador'

# ELS PUNTS EDITABLES DE LA PANTALLA DE CONSULTA. El defecte: la pantalla nomes
# sabia pintar les dades JA DESADES, o sigui que d'un punt que no s'havia
# omplert mai no en sortia cap casella i no es podia escriure res.
$llicEd = Read-LlicCataleg (Join-Path $Global:EstructuralsDir 'LLIC.json')
$req1Ed = Get-ParsedCataleg -path (Join-Path $Global:EstructuralsDir 'REQ1.json')
if ($null -ne $llicEd -and $null -ne $req1Ed) {
    $edit = Get-LlicenciaPuntsEditables $llicEd $req1Ed
    $abEd = @($edit['Abans']); $deEd = @($edit['Despres'])
    Assert ($abEd.Count -ge 30) ('Get-LlicenciaPuntsEditables: tots els punts d''ABANS (' + $abEd.Count + ')')
    Assert ($deEd.Count -ge 40) ('Get-LlicenciaPuntsEditables: i els de DESPRES (' + $deEd.Count + ')')
    Assert (-not (@($abEd) | Where-Object { @($_.Camps).Count -eq 0 })) 'Get-LlicenciaPuntsEditables: CAP punt d''ABANS es queda sense camps'
    Assert (-not (@($abEd) | Where-Object { -not (@($_.Camps) -contains 'Id Firmadoc') })) 'Get-LlicenciaPuntsEditables: i tots demanen l''Id Firmadoc'
    $clausEd = @(@($abEd) | ForEach-Object { [string]$_.Clau })
    AssertEq (@($clausEd | Select-Object -Unique).Count) $clausEd.Count 'Get-LlicenciaPuntsEditables: cap clau repetida'
    Assert (-not (@($abEd) | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.Titol) })) 'Get-LlicenciaPuntsEditables: tots tenen titol'
    # El punt amb dades propies (SDR) demana els SEUS camps, no nomes el firmadoc.
    $sdrEd = @(@($abEd) | Where-Object { @($_.Camps) -contains 'NIMA' })
    Assert ($sdrEd.Count -ge 1) 'Get-LlicenciaPuntsEditables: els camps propis del punt (NIMA) tambe hi surten'
}
AssertEq (@((Get-LlicenciaPuntsEditables $null $null)['Abans']).Count) 0 'Get-LlicenciaPuntsEditables: sense cataleg, cap punt'

# _LlicDbMemEditable: el que ve del JSON son PSCustomObject i s'hi ha de poder
# escriure a sobre sense Add-Member a cada nivell.
$memEd = _LlicDbMemEditable ((@{
    'S::A' = @{ Marcat = $true; Estat = 'si'; Valors = @{ 'Id Firmadoc' = 'FD-1' }; Subs = @{ '0' = $true } }
} | ConvertTo-Json -Depth 10) | ConvertFrom-Json)
AssertEq ([bool]$memEd['S::A']['Marcat']) $true '_LlicDbMemEditable: el marcat'
AssertEq ([string]$memEd['S::A']['Valors']['Id Firmadoc']) 'FD-1' '_LlicDbMemEditable: els valors'
$memEd['S::A']['Valors']['Expedient'] = '901417/23'
$memEd['S::A']['Estat'] = 'no'
AssertEq ([string]$memEd['S::A']['Valors']['Expedient']) '901417/23' '_LlicDbMemEditable: s''hi pot AFEGIR un camp nou'
# ...i tornar a la forma desable sense perdre res.
$memDes = ConvertTo-LlicenciaMemoria $memEd
AssertEq ([string]$memDes['S::A']['Estat']) 'no' '_LlicDbMemEditable: i tornar a desar-ho'
AssertEq ([string]$memDes['S::A']['Valors']['Expedient']) '901417/23' '_LlicDbMemEditable: amb el camp nou inclos'
AssertEq (@((_LlicDbMemEditable $null).Keys).Count) 0 '_LlicDbMemEditable: amb $null, mapa buit'

# L'ANNEX 1 NOMES SI ENCARA S'HA DE DEMANAR L'AUTORITZACIO. L'annex diu quina
# documentacio cal per demanar-la: si ja se'n disposa, no pinta res.
$pAnx = @(
    [pscustomobject]@{ Titol = 'LL Prov Compatibilitat'; Condicio = 'provisional'; Estat = 'no' },
    [pscustomobject]@{ Titol = 'Incendis';               Condicio = '';            Estat = 'si' })
AssertEq (_LlicCalAnnex1 $pAnx $true) $true '_LlicCalAnnex1: provisional i NO es disposa -> hi va'
$pAnx2 = @(
    [pscustomobject]@{ Titol = 'LL Prov Compatibilitat'; Condicio = 'provisional'; Estat = 'si' },
    [pscustomobject]@{ Titol = 'Incendis';               Condicio = '';            Estat = 'no' })
AssertEq (_LlicCalAnnex1 $pAnx2 $true) $false '_LlicCalAnnex1: es disposa de la LL Prov -> FORA'
AssertEq (_LlicCalAnnex1 $pAnx $false) $false '_LlicCalAnnex1: sense llicencia provisional, mai'
AssertEq (_LlicCalAnnex1 @() $true) $true '_LlicCalAnnex1: si el punt no s''ha marcat, hi va'
# La condicio es compara sense distingir majuscules ni espais.
AssertEq (_LlicCalAnnex1 @([pscustomobject]@{ Condicio = ' Provisional '; Estat = 'si' }) $true) $false '_LlicCalAnnex1: la condicio, tolerant'

# ELS DOCUMENTS SIGNATS pel tecnic redactor: es desen i es tornen a llegir.
$docsT = [ordered]@{
    'Projecte' = @{ Marcat = $true;  Id = '9741790' }
    'Annexos'  = @{ Marcat = $true;  Id = '' }
}
$itemsT = @(_LlicItemsDocsSignats $docsT)
AssertEq $itemsT.Count 2 '_LlicItemsDocsSignats: nomes els marcats'
Assert ([bool]($itemsT[0] -like 'Projecte (Id Firmadoc: 9741790)')) '_LlicItemsDocsSignats: amb l''Id Firmadoc'
AssertEq ([string]$itemsT[1]) 'Annexos' '_LlicItemsDocsSignats: sense Id, nomes el nom'
# ...i EN L'ORDRE DEL CATALEG, que un hashtable no en te.
$ordreT = @(_LlicItemsDocsSignats ([ordered]@{
    'Annexos'  = @{ Marcat = $true; Id = '' }
    'Projecte' = @{ Marcat = $true; Id = '' } }))
AssertEq ($ordreT -join '|') 'Projecte|Annexos' '_LlicItemsDocsSignats: ordre del cataleg, no del mapa'
AssertEq (@(_LlicItemsDocsSignats @{ 'Projecte' = @{ Marcat = $false; Id = 'x' } }).Count) 0 '_LlicItemsDocsSignats: el no marcat no hi surt'
AssertEq (@(_LlicItemsDocsSignats $null).Count) 0 '_LlicItemsDocsSignats: amb $null, cap'
AssertEq (@((_LlicDocsBuits).Keys).Count) (@(_LlicDocsSignats).Count) '_LlicDocsBuits: un per document del cataleg'
# El pas per JSON (PSCustomObjects) no els pot perdre.
$docsJ = ($docsT | ConvertTo-Json -Depth 10) | ConvertFrom-Json
AssertEq (@(_LlicItemsDocsSignats $docsJ).Count) 2 '_LlicItemsDocsSignats: sobreviu el pas per JSON'
$stD = @{ Header = @{ ID_GIA = '1457' }; TecnicDocs = $docsT }
$recD = ConvertTo-LlicenciaRecord $stD
$stD2 = @{}
[void](Restore-LlicenciaState (($recD | ConvertTo-Json -Depth 20) | ConvertFrom-Json) $stD2)
AssertEq ([bool](_LlicDbAMapa $stD2['TecnicDocs'])['Projecte']['Marcat']) $true 'base de dades: els documents signats es recorden'
AssertEq ([string](_LlicDbAMapa $stD2['TecnicDocs'])['Projecte']['Id']) '9741790' 'base de dades: ...amb el seu Id Firmadoc'

# ---------------------------------------------------------------------------
# MODIFICACIO NO SUBSTANCIAL i TRASPAS (MnsTraspas.ps1 + MNSTRAS.json)
# ---------------------------------------------------------------------------
$fasesM = @(_MnsFases)
AssertEq $fasesM.Count 2 '_MnsFases: els dos informes curts'
Assert ([bool](_MnsEsFase 'mns'))     '_MnsEsFase: la modificacio no substancial'
Assert ([bool](_MnsEsFase 'traspas')) '_MnsEsFase: el traspas'
Assert (-not (_MnsEsFase 'requeriment')) '_MnsEsFase: el requeriment NO hi es'
AssertEq (@(_LlicTotesLesFases).Count) 5 '_LlicTotesLesFases: 3 informes llargs + 2 curts'
$clausF = @(@(_LlicTotesLesFases) | ForEach-Object { [string]$_.Clau })
AssertEq (@($clausF | Select-Object -Unique).Count) $clausF.Count '_LlicTotesLesFases: cap clau repetida'

# Quin node hi entra segons si hi ha punts de REQ1 marcats.
Assert ([bool](_MnsNodeEntra '' $true))  '_MnsNodeEntra: sense clau, sempre'
Assert ([bool](_MnsNodeEntra 'amb-observacions' $true))     '_MnsNodeEntra: amb-observacions quan n''hi ha'
Assert (-not (_MnsNodeEntra 'amb-observacions' $false))     '_MnsNodeEntra: ...i no quan no n''hi ha'
Assert ([bool](_MnsNodeEntra 'sense-observacions' $false))  '_MnsNodeEntra: sense-observacions quan no n''hi ha'
Assert (-not (_MnsNodeEntra 'sense-observacions' $true))    '_MnsNodeEntra: ...i no quan n''hi ha'
Assert ([bool](_MnsNodeEntra ' AMB-OBSERVACIONS ' $true))   '_MnsNodeEntra: la clau, tolerant'

AssertEq (_MnsNomFitxer ([datetime]'2026-08-21') 'mns' '1457') '2026-08-21_LlicMNS_GIA 1457.docx' '_MnsNomFitxer: modificacio no substancial'
Assert ([bool]((_MnsNomFitxer ([datetime]'2026-08-21') 'traspas' '1') -like '*LlicTraspas*')) '_MnsNomFitxer: traspas'

# QUINES CONCLUSIONS. Funcio PURA: es prova sense cataleg ni Word.
$cMns  = @([pscustomobject]@{ Title = 'Actes dels controls periodics'; Body = '59.1.d...' })
$cReq1 = @([pscustomobject]@{ Title = 'Requeriment'; Body = 'Vist l''anterior, cal requerir...' },
           [pscustomobject]@{ Title = 'Desistiment'; Body = 'una altra' })
$k1 = @(_MnsTriaConclusions $cMns $cReq1 'mns' $true)
AssertEq $k1.Count 2 'MNS amb punts: 59.1.d + la de REQUERIMENT de REQ1'
AssertEq ([string]$k1[0].Title) 'Actes dels controls periodics' 'MNS amb punts: el 59.1.d va primer'
AssertEq ([string]$k1[1].Title) 'Requeriment' 'MNS amb punts: ...i despres la de REQ1'
$k2 = @(_MnsTriaConclusions $cMns $cReq1 'mns' $false)
AssertEq $k2.Count 1 'MNS sense punts: nomes el 59.1.d'
AssertEq ([string]$k2[0].Title) 'Actes dels controls periodics' 'MNS sense punts: ...i es aquell'
$k3 = @(_MnsTriaConclusions $cMns $cReq1 'traspas' $true)
AssertEq $k3.Count 1 'Traspas amb punts: nomes la de REQUERIMENT'
AssertEq ([string]$k3[0].Title) 'Requeriment' 'Traspas amb punts: ...i es aquella'
AssertEq (@(_MnsTriaConclusions $cMns $cReq1 'traspas' $false).Count) 0 'Traspas sense punts: CAP conclusio (ja es al text fix)'
# La de REQ1 NO es una copia: surt del mateix grup que els requeriments normals.
$reqCat = Read-Conclusions $Global:ConclusionsPath 'REQ1'
Assert ([bool](@($reqCat.Selectable) | Where-Object { [string]$_.Title -eq 'Requeriment' })) 'cataleg: la conclusio de REQUERIMENT es la de REQ1'
$mnsCat = Read-Conclusions $Global:ConclusionsPath 'MNS'
AssertEq (@($mnsCat.Selectable).Count) 1 'cataleg: el grup MNS porta l''avis del 59.1.d'
Assert ([bool]((@($mnsCat.Selectable)[0].Body) -like '*59.1.d*')) 'cataleg: ...i es aquell'

# EL CATALEG DE TEXT.
$mnsPath = Join-Path $Global:EstructuralsDir 'MNSTRAS.json'
Assert (Test-Path -LiteralPath $mnsPath) 'MNSTRAS.json: hi es'
$catM = Read-MnsCataleg $mnsPath
Assert ($null -ne $catM) 'MNSTRAS.json: es valid'
AssertEq ([string]$catM.familia) 'mnstraspas' 'MNSTRAS.json: la familia'
Assert ($null -eq (_MnsSeccio $catM 'no-existeix')) '_MnsSeccio: una clau desconeguda -> $null'

foreach ($fM in @('mns', 'traspas')) {
    $ambM   = @(_MnsParagrafs $catM $fM $true)
    $senseM = @(_MnsParagrafs $catM $fM $false)
    Assert ($ambM.Count -gt 0)   ($fM + ': amb punts, hi ha paragrafs')
    Assert ($senseM.Count -gt 0) ($fM + ': sense punts, tambe')
    $txtAmb   = (@($ambM   | ForEach-Object { @($_.Linies) }) -join ' ')
    $txtSense = (@($senseM | ForEach-Object { @($_.Linies) }) -join ' ')
    Assert ([bool]($txtAmb   -like '*amb les seg*ents observacions:*')) ($fM + ': amb punts ho diu')
    Assert (-not ($txtAmb    -like '*sense m*s observacions*'))         ($fM + ': ...i no diu tambe el contrari')
    Assert ([bool]($txtSense -like '*sense m*s observacions*'))         ($fM + ': sense punts ho diu')
    Assert (-not ($txtSense  -like '*amb les seg*ents observacions*'))  ($fM + ': ...i no diu tambe el contrari')
    # La frase de les observacions va en NEGRETA (ve del **...** del cataleg).
    Assert ([bool]($txtAmb -like '*`*`*S*informa FAVORABLEMENT*' -or $txtAmb -like '*`*`*En relaci*')) ($fM + ': la frase de les observacions va en negreta')
    Assert (-not ($txtAmb -match '\[\[URL\]\]|\[CAMP:|\[OPCIO:')) ($fM + ': cap marcador intern al text')
    # FORA l'enumeracio buida d'observacions: ara alli hi van els punts de REQ1.
    AssertEq (@($ambM | Where-Object { [string]$_.Tipus -eq 'llista' }).Count) `
             (@($senseM | Where-Object { [string]$_.Tipus -eq 'llista' }).Count) `
             ($fM + ': els punts de REQ1 no afegeixen cap llista buida')
}
# La MNS porta SEMPRE la llista de les modificacions justificades (aquella si
# que l'escriu l'usuari a ma); el Traspas no en porta cap.
AssertEq (@(@(_MnsParagrafs $catM 'mns' $false) | Where-Object { [string]$_.Tipus -eq 'llista' }).Count) 1 'mns: la llista de modificacions hi es sempre'
AssertEq (@(@(_MnsParagrafs $catM 'traspas' $true) | Where-Object { [string]$_.Tipus -eq 'llista' }).Count) 0 'traspas: cap llista'
# EL PARAGRAF DELS CRITERIS DE SUBSTANCIALITAT HI ES, i va DESPRES de la llista
# de modificacions i ABANS de la frase d'observacions. (El vaig treure quan el
# _BE de l'agost no el portava; l'usuari l'ha tornat a posar, o sigui que ara
# la prova diu el contrari i el text torna a viure al cataleg.)
foreach ($ambObsM in @($true, $false)) {
    $parasM = @(_MnsParagrafs $catM 'mns' $ambObsM)
    $textM = @($parasM | ForEach-Object { @($_.Linies) -join ' ' })
    $iCri = [Array]::FindIndex([string[]]$textM, [Predicate[string]]{ param($x) $x -like '*CRITERIS DE SUBSTANCIALITAT*' })
    Assert ($iCri -ge 0) ('mns: hi ha el paragraf dels criteris de substancialitat (obs=' + $ambObsM + ')')
    $iLli = [Array]::FindIndex([string[]]@($parasM | ForEach-Object { [string]$_.Tipus }), [Predicate[string]]{ param($x) $x -eq 'llista' })
    Assert ($iLli -ge 0 -and $iCri -gt $iLli) 'mns: ...i va DESPRES de la llista de modificacions'
    $iObs = [Array]::FindIndex([string[]]$textM, [Predicate[string]]{ param($x) $x -like '*S*informa FAVORABLEMENT*' })
    Assert ($iObs -ge 0 -and $iCri -lt $iObs) 'mns: ...i ABANS de la frase d''observacions'
}
# Al TRASPAS no hi va: es una regla de la modificacio no substancial.
Assert (-not ((@(@(_MnsParagrafs $catM 'traspas' $true) | ForEach-Object { @($_.Linies) }) -join ' ') -like '*CRITERIS DE SUBSTANCIALITAT*')) 'traspas: cap paragraf de criteris de substancialitat'
AssertEq (@(_MnsParagrafs $null 'mns' $true).Count) 0 '_MnsParagrafs: sense cataleg, cap paragraf'

# GENERACIO SENCERA amb el Word simulat.
if ((Test-Path -LiteralPath $mnsPath) -and (Test-Path -LiteralPath (Join-Path $Global:EstructuralsDir 'REQ1.json'))) {
    . (Join-Path $TestsDir 'FormatDoubles.ps1')
    $selM = [pscustomobject]@{ Range = [pscustomobject]@{ Start = 0; End = 0 } }
    $selM | Add-Member ScriptMethod EndKey { param($u) } -Force
    $selM | Add-Member ScriptMethod InsertBreak { param($b) } -Force
    $docM = [pscustomobject]@{}
    $docM | Add-Member ScriptMethod Activate {} -Force
    $docM | Add-Member ScriptMethod Save {} -Force
    $docM | Add-Member ScriptMethod Close { param($x) } -Force
    $wordM = [pscustomobject]@{ Selection = $selM }
    function _ResolveOutputDir { return ([System.IO.Path]::GetTempPath()) }
    function _GetUniqueOutputPath($d, $b) { return (Join-Path $d $b) }
    function _OpenOutputDocument($w, $p) { return $script:_docMprova }
    function Select-CapcaleraBlock($d, $w) { [void]$global:emitCalls.Add("CAPCALERA|$w") }
    function Apply-HeaderReplacements { param($doc, $header) }
    $script:_docMprova = $docM
    if ([string]::IsNullOrWhiteSpace($env:TEMP)) { $env:TEMP = [System.IO.Path]::GetTempPath() }

    # Uns quants punts de REQ1 de debo, per veure que surten amb el format de
    # REQ1. La tria es munta com la munta el programa: per CLAU (la mateixa
    # Build-SelectionFromKeys que fan servir el mode mobil i Controls periodics).
    $req1M = Get-ParsedCataleg -path (Join-Path $Global:EstructuralsDir 'REQ1.json')
    $sec0M = @($req1M.Sections)[0]
    $clauM = @(@($sec0M.Items) | Where-Object { [string]$_.Kind -eq 'item' } |
               ForEach-Object { _ItemKey $sec0M.Title $_.Short } | Select-Object -First 2)
    $secM = @(Build-SelectionFromKeys @($req1M.Sections) $clauM)
    Assert ($clauM.Count -ge 1) 'proves MNS: hi ha punts de REQ1 per triar'
    foreach ($fM in @('mns', 'traspas')) {
        foreach ($ambM in @($true, $false)) {
            $global:emitCalls.Clear()
            $petaM = $false
            try {
                [void](Build-MnsDocument $wordM @{
                    Fase = $fM; Header = @{ ID_GIA = '1483'; TITULAR = 'PROVA SL' }
                    Fields = [ordered]@{}
                    Punts = $(if ($ambM) { $secM } else { @() })
                    Cataleg = $catM })
            } catch { $petaM = $true; Write-Host ("    EXCEPCIO ($fM/$ambM): " + $_.Exception.Message) -ForegroundColor Red }
            AssertEq $petaM $false ($fM + '/' + $ambM + ': genera sense petar')
            $emM = @($global:emitCalls)
            Assert ([bool]($emM -contains 'CAPCALERA|LLIC')) ($fM + ': fa servir la capcalera de Llicencia')
            # El tancament, SEMPRE, i del cataleg (va com a conclusio, com a REQ1).
            Assert ([bool]($emM | Where-Object { $_ -like 'CONCL|Ho poso al seu coneixement*' })) ($fM + '/' + $ambM + ': porta el tancament')
            Assert ([bool]($emM | Where-Object { $_ -like 'CONCL|Cornell* de Llobregat,' })) ($fM + '/' + $ambM + ': ...i la linia de Cornella')
            # ELS PUNTS DE REQ1: amb el format de REQ1 (seccio + items numerats).
            $itemsM = @($emM | Where-Object { $_ -like 'ITEM|*' })
            if ($ambM) {
                Assert ($itemsM.Count -gt 0) ($fM + ': amb punts, hi surten els items de REQ1')
                Assert ([bool](($itemsM[0] -split '\|')[1] -match '^\d+\.$')) ($fM + ': ...numerats com a REQ1')
                Assert ([bool]($emM | Where-Object { $_ -like 'SECT|*' })) ($fM + ': ...amb el titol de la seccio de REQ1')
            } else {
                AssertEq $itemsM.Count 0 ($fM + ': sense punts, cap item')
            }
            # El bloc de CONCLUSIONS: nomes quan te alguna linia.
            $capM = @($emM | Where-Object { $_ -like 'CONCLCAP|*' })
            $volCap = ($fM -eq 'mns') -or $ambM
            AssertEq ($capM.Count -gt 0) $volCap ($fM + '/' + $ambM + ': el bloc CONCLUSIONS surt nomes si te linies')
            if ($fM -eq 'mns') {
                Assert ([bool]($emM | Where-Object { $_ -like 'CONCL|*59.1.d*' })) ($fM + '/' + $ambM + ': la MNS porta sempre l''avis del 59.1.d')
            }
            if ($ambM) {
                Assert ([bool]($emM | Where-Object { $_ -like 'CONCL|*cal requerir l*esmena*' })) ($fM + ': amb punts, la conclusio de REQUERIMENT')
            } else {
                Assert (-not ($emM | Where-Object { $_ -like 'CONCL|*cal requerir l*esmena*' })) ($fM + ': sense punts, cap conclusio de requeriment')
            }
        }
    }
}

# ---------------------------------------------------------------------------
# LA CAPCALERA EN JSON (CapcaleraJson.ps1) - sobre el .docx REAL
# ---------------------------------------------------------------------------
# El .docx mana en el FORMAT (escut, taula, tabulacions) i el JSON en el TEXT.
# Aqui es comprova que llegir-lo, escriure'l i tornar-lo a llegir no en canvia
# res mes: aquest fitxer es l'unic que no es pot regenerar.
$capDocx = Join-Path $Global:EstructuralsDir '0 CAPCALERA.docx'
if (Test-Path -LiteralPath $capDocx) {
    $capXml = _CapLlegeixDocumentXml $capDocx
    Assert (-not [string]::IsNullOrWhiteSpace($capXml)) 'capcalera: es llegeix el document.xml'
    $capBlocs = @(_CapBlocsDelXml $capXml)
    AssertEq $capBlocs.Count 3 'capcalera: els tres blocs (generic, ACT_EXTR, LLIC)'
    AssertEq ([string]$capBlocs[0].Clau) ''         'capcalera: el primer bloc es el generic'
    AssertEq ([string]$capBlocs[1].Clau) 'ACT_EXTR' 'capcalera: el segon, activitats extraordinaries'
    AssertEq ([string]$capBlocs[2].Clau) 'LLIC'     'capcalera: el tercer, llicencia'
    # Cada bloc ha de portar els seus <<PLACEHOLDER>>: si un dia en desapareix
    # un, l'informe surt amb una linia BUIDA i ningu se n'assabenta.
    foreach ($bC in $capBlocs) {
        $totC = (@($bC.Linies) | ForEach-Object { [string]$_.Etiqueta + [string]$_.Valor }) -join ' '
        foreach ($ph in @('<<ID_GIA>>', '<<EXP_NUM>>', '<<ADRECA>>', '<<ACTIVITAT>>', '<<TITULAR>>')) {
            Assert ($totC.Contains($ph)) ('capcalera [' + [string]$bC.Clau + ']: hi ha ' + $ph)
        }
    }
    # La classificacio NOMES al bloc de llicencia.
    $txtLlic = ((@($capBlocs[2].Linies) | ForEach-Object { [string]$_.Valor }) -join ' ')
    $txtGen  = ((@($capBlocs[0].Linies) | ForEach-Object { [string]$_.Valor }) -join ' ')
    Assert ($txtLlic.Contains('<<CLASSIFICACIO>>')) 'capcalera: la classificacio es al bloc de llicencia'
    Assert (-not $txtGen.Contains('<<CLASSIFICACIO>>')) 'capcalera: ...i NO al generic'
    # FORA EL CAIXETI DE LA NOTA de l'Ordenanca, de TOTES les capcaleres: era
    # una taula (<w:tbl>) a la generica i una altra a la de llicencia.
    Assert (-not ($capXml.Contains('<w:tbl>'))) 'capcalera: cap taula al document (fora el caixeti de la Nota)'
    foreach ($bN in $capBlocs) {
        $totN = ((@($bN.Linies) | ForEach-Object { [string]$_.Etiqueta + [string]$_.Valor }) -join ' ')
        Assert (-not ($totN -like '*Nota:*'))     ('capcalera [' + [string]$bN.Clau + ']: cap "Nota:"')
        Assert (-not ($totN -like '*Ordenan*a relativa*')) ('capcalera [' + [string]$bN.Clau + ']: ...ni el text de l''Ordenanca')
    }
    # LES TRES CAPCALERES SON LA MATEIXA amb desviacions: totes porten el
    # departament, les linies de camp i l'INFORME, i les uniques diferencies
    # son ASSUMPTE/Dates (act. extraordinaries) i Classificacio (llicencia).
    foreach ($bN in $capBlocs) {
        $texts = @(@($bN.Linies) | Where-Object { [string]$_.Tipus -eq 'text' } | ForEach-Object { [string]$_.Valor })
        Assert ([bool]($texts -contains 'INFORME')) ('capcalera [' + [string]$bN.Clau + ']: hi ha l''INFORME')
        Assert ([bool](@($texts) | Where-Object { $_ -like 'Activitats i Ordenances*' })) ('capcalera [' + [string]$bN.Clau + ']: ...i el departament')
    }
    # Cap etiqueta sense marcador: una linia "Camp:" sense <<...>> nomes pot
    # sortir BUIDA a l'informe (ja va passar amb "Classificacio:").
    foreach ($bC in $capBlocs) {
        foreach ($lC in @($bC.Linies)) {
            if ([string]$lC.Tipus -ne 'etiqueta') { continue }
            if (([string]$lC.Etiqueta).Trim() -eq 'Nota:') { continue }
            $v = [string]$lC.Valor
            Assert ([bool]($v -match '<<[A-Z_]+>>' -or $v.Length -gt 20)) ('capcalera: "' + ([string]$lC.Etiqueta).Trim() + '" te valor o marcador')
        }
    }
    # A quin informe s'aplica cada bloc (el que l'usuari no podia saber).
    AssertEq (@(_CapAplicaDe 'ACT_EXTR').Count) 1 '_CapAplicaDe: act. extraordinaries, un tipus'
    AssertEq (@(_CapAplicaDe 'LLIC').Count) 3 '_CapAplicaDe: llicencia, els seus tres informes'
    Assert ([bool]((_CapAplicaDe '') -contains 'Requeriment - Nou (REQ1)')) '_CapAplicaDe: el generic, el requeriment'
    # ...i sense la trampa de la coma dins d'un @(): cap element esmicolat.
    foreach ($aC in @(@(_CapAplicaDe '') + @(_CapAplicaDe 'LLIC') + @(_CapAplicaDe 'ACT_EXTR'))) {
        Assert (([string]$aC).Length -gt 3) ('_CapAplicaDe: "' + $aC + '" no esta esmicolat')
    }

    # ANADA I TORNADA: el JSON generat, aplicat sobre el mateix XML, no en canvia
    # ni un byte.
    $capJson = (_CapModelAJson $capBlocs | ConvertTo-Json -Depth 20) | ConvertFrom-Json
    AssertEq (_CapAplicaAlXml $capXml $capJson) $capXml 'capcalera: sense canvis, el document no es toca'

    # ...i un canvi de text nomes toca aquella linia (no refa la regio: si la
    # refes, el .docx sortiria diferent a cada desat encara que no s'hagues
    # canviat res, i 'Actualitzar.bat' committejaria una capcalera "nova" cada
    # vegada).
    foreach ($nC in @($capJson.nodes)) {
        foreach ($fC in @($nC.fills)) {
            if ([string]$fC.clau -eq 'p2') { $fC.titol = 'Expedient: ' }
        }
    }
    $capNou = _CapAplicaAlXml $capXml $capJson
    Assert (($capNou.Length - $capXml.Length) -lt 200) 'capcalera: canviar un text nomes toca aquella linia'
    Assert ($capNou -ne $capXml) 'capcalera: un canvi de text si que el toca'
    $blocsNous = @(_CapBlocsDelXml $capNou)
    $l2 = @(@($blocsNous[0].Linies) | Where-Object { [int]$_.Para -eq 2 })[0]
    AssertEq ([string]$l2.Etiqueta) 'Expedient: ' 'capcalera: l''etiqueta nova hi es'
    AssertEq ([string]$l2.Valor) '<<EXP_NUM>>' 'capcalera: ...i el marcador no s''ha tocat'
    # La resta de linies, intactes.
    $abansTxt = ((@($capBlocs[0].Linies) | Where-Object { [int]$_.Para -ne 2 } | ForEach-Object { [string]$_.Etiqueta + '|' + [string]$_.Valor }) -join '###')
    $despresTxt = ((@($blocsNous[0].Linies) | Where-Object { [int]$_.Para -ne 2 } | ForEach-Object { [string]$_.Etiqueta + '|' + [string]$_.Valor }) -join '###')
    AssertEq $despresTxt $abansTxt 'capcalera: cap altra linia no s''ha mogut'
    # El document segueix sent un XML de Word ben format: els espais de noms de
    # l'arrel, el mc:Ignorable i cap prefix ns0: (aixo ja va corrompre el fitxer
    # una vegada, vegeu CLAUDE.md).
    Assert ($capNou.Contains('mc:Ignorable')) 'capcalera: el mc:Ignorable hi segueix'
    Assert (-not ($capNou -match '\bns\d+:')) 'capcalera: cap prefix ns0:/ns1: inventat'
    Assert ((@([regex]::Matches($capNou.Substring(0, [Math]::Min(2000, $capNou.Length)), 'xmlns:')).Count) -ge 15) 'capcalera: hi son tots els espais de noms'
    $petaXml = $false
    try { [void]([xml]$capNou) } catch { $petaXml = $true }
    AssertEq $petaXml $false 'capcalera: el XML resultant es valid'

    # I EL JSON QUE HI HA AL REPOSITORI diu el mateix que el .docx.
    $capJsonPath = Join-Path $Global:EstructuralsDir '0 CAPCALERA.json'
    if (Test-Path -LiteralPath $capJsonPath) {
        $capDelDisc = Get-Content -LiteralPath $capJsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
        AssertEq (_CapAplicaAlXml $capXml $capDelDisc) $capXml 'capcalera: el JSON del repositori quadra amb el .docx'
        AssertEq ([string]$capDelDisc.familia) 'capcalera' 'capcalera: la familia del JSON'
        AssertEq (@($capDelDisc.nodes).Count) 3 'capcalera: el JSON porta els tres blocs'
    }
}
# EL JSON MANA I EL .docx ES DERIVA: afegir una linia al JSON l'ha de fer
# apareixer al document, amb el format de les altres. Aixi les tres capcaleres
# deixen de ser tres copies a mantenir a ma.
if (Test-Path -LiteralPath $capDocx) {
    $capX2 = _CapLlegeixDocumentXml $capDocx
    $capJ2 = (_CapModelAJson (_CapBlocsDelXml $capX2) | ConvertTo-Json -Depth 20) | ConvertFrom-Json
    $genN = @($capJ2.nodes)[0]
    $nCampsAbans = @(@($genN.fills) | Where-Object { [string]$_.tipus -eq 'etiqueta' }).Count
    # Una linia nova, calcada d'una que ja hi es.
    $novaN = ((@($genN.fills)[4]) | ConvertTo-Json -Depth 10) | ConvertFrom-Json
    $novaN.titol = 'Classificaci' + [char]0x00F3 + ': '
    $novaN.cos[0].runs[0].t = '<<CLASSIFICACIO>>'
    $genN.fills = @(@($genN.fills)[0..4]) + @($novaN) + @(@($genN.fills)[5..(@($genN.fills).Count - 1)])
    $capX3 = _CapAplicaAlXml $capX2 $capJ2 'Bookman Old Style'
    $bloc3 = @(_CapBlocsDelXml $capX3)[0]
    $camps3 = @(@($bloc3.Linies) | Where-Object { [string]$_.Tipus -eq 'etiqueta' })
    AssertEq $camps3.Count ($nCampsAbans + 1) 'capcalera: afegir una linia al JSON l''afegeix al document'
    Assert ([bool](@($camps3) | Where-Object { [string]$_.Valor -eq '<<CLASSIFICACIO>>' })) 'capcalera: ...amb el seu marcador'
    # ...i al lloc que li tocava: just despres de l'Activitat, com a la de
    # llicencia. Es mira per POSICIO RELATIVA, no per index absolut.
    $etq3 = @(@($camps3) | ForEach-Object { ([string]$_.Etiqueta).Trim() })
    $iAct = [Array]::IndexOf($etq3, 'Activitat:')
    $iCla = [Array]::IndexOf($etq3, ('Classificaci' + [char]0x00F3 + ':'))
    Assert ($iAct -ge 0 -and $iCla -eq ($iAct + 1)) 'capcalera: ...i al lloc que li tocava (just sota Activitat)'
    # ...i la resta del bloc no s'ha mogut.
    $texts3 = @(@($bloc3.Linies) | Where-Object { [string]$_.Tipus -eq 'text' } | ForEach-Object { [string]$_.Valor })
    Assert ([bool]($texts3 -contains 'INFORME')) 'capcalera: ...i l''INFORME segueix al seu lloc'
    $petaX3 = $false
    try { [void]([xml]$capX3) } catch { $petaX3 = $true }
    AssertEq $petaX3 $false 'capcalera: el document amb la linia nova segueix sent XML valid'
    # LA LLETRA surt de la configuracio: era l'unic tros de l'informe que no
    # obeia BodyFontName (la porta escrita a cada <w:r>).
    $capX4 = _CapAplicaAlXml $capX2 $capJ2 'Arial'
    Assert ($capX4.Contains('w:ascii="Arial"')) 'capcalera: la lletra surt de la configuracio'
}

# L'editor ensenya a quin informe s'aplica cada seccio (capcalera i conclusions).
AssertEq (_Ed_AplicaText 'capcalera' @{ tipus='seccio'; clau='LLIC'; titol='x' }) ((_CapAplicaDe 'LLIC') -join ', ') '_Ed_AplicaText: capcalera de llicencia'
AssertEq (_Ed_AplicaText 'conclusions' @{ tipus='seccio'; clau=''; titol='REQ1' }) 'Requeriment - Nou (REQ1)' '_Ed_AplicaText: conclusions de REQ1'
AssertEq (_Ed_AplicaText 'cataleg' @{ tipus='seccio'; clau=''; titol='X' }) '' '_Ed_AplicaText: als catalegs normals, res'
AssertEq (_Ed_AplicaText 'capcalera' @{ tipus='etiqueta'; clau='p1'; titol='ID GIA:' }) '' '_Ed_AplicaText: nomes a les seccions'

# ---------------------------------------------------------------------------
# EL MENU: els handlers que trien opcio han de veure $result
# ---------------------------------------------------------------------------
# Un scriptblock SENSE .GetNewClosure() NO veu els locals de la funcio que el
# crea (nomes l'ambit de l'script). Els enllacos 'Capcalera'/'Conclusions' es
# van posar ABANS de declarar $result i sense closure: "$result.Choice = ..."
# queia sobre $null, el menu es tancava i el programa SORTIA sense fer res.
$srcMenu2 = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $TestsDir) 'Menu.ps1') -Raw
$iDecl = $srcMenu2.IndexOf('$result = @{ Choice = $null }')
Assert ($iDecl -ge 0) 'menu: $result es declara a Select-Mode'
# Nomes les linies de CODI (una menció dins d'un comentari no compta).
$nUsos = 0
foreach ($mR in [regex]::Matches($srcMenu2, '(?m)^[ \t]*\$result\.Choice\s*=')) {
    $nUsos++
    Assert ($mR.Index -gt $iDecl) 'menu: cap handler no toca $result abans de declarar-lo'
    # ...i el bloc on va ha de ser una CLOSURE (si no, no el veu).
    $tros = $srcMenu2.Substring($mR.Index, [Math]::Min(600, $srcMenu2.Length - $mR.Index))
    Assert ($tros.Contains('.GetNewClosure()')) 'menu: el handler que tria opcio es una closure'
}
Assert ($nUsos -ge 3) 'menu: hi ha els handlers de tria (rajoles + enllacos)'
# UN LinkLabel TE UNA SOLA LLETRA per a tot el text: amb la Segoe UI del
# programa, un emoji hi surt com un quadrat. Als xips de les rajoles si que n'hi
# ha perque alla es dibuixa a part, amb Segoe UI Emoji.
$iLL = $srcMenu2.IndexOf('New-Object System.Windows.Forms.LinkLabel')
if ($iLL -ge 0) {
    $trosLL = $srcMenu2.Substring($iLL, [Math]::Min(700, $srcMenu2.Length - $iLL))
    Assert (-not ($trosLL -match '\$ll\.Text\s*=\s*\[string\]\$pencil')) 'menu: els enllacos no porten emoji (el LinkLabel no en sap)'
}

# ---------------------------------------------------------------------------
# 'continue' DINS D'UN 'switch' NO CONTINUA EL FOREACH DE FORA
# ---------------------------------------------------------------------------
# Comprovacio del propi llenguatge, perque la regla quedi escrita i provada: el
# desat de la base de llicencies hi va caure (les entrades de la documentacio
# del projecte queien al codi dels blocs de punts i petaven).
$provaSw = New-Object System.Collections.ArrayList
foreach ($x in 1..3) {
    switch ("$x") {
        '1' { continue }
    }
    [void]$provaSw.Add($x)
}
AssertEq ($provaSw -join ',') '1,2,3' "'continue' dins d'un switch NO continua el foreach de fora"

# ---------------------------------------------------------------------------
# L'ACCES DIRECTE (AccesDirecte.ps1) i el boto de la carpeta d'informes
# ---------------------------------------------------------------------------
# Windows NO deixa ancorar un .bat a la barra de tasques: l'acces directe ha
# d'apuntar a un EXECUTABLE. Per aixo va a wscript.exe amb el .vbs com a
# argument, que es el mateix que fa GenerarInforme.bat.
$objAD = Get-AccesDirecteObjectiu 'C:\clone\informes' 'C:\Windows'
Assert ([string]$objAD.Desti -like '*\wscript.exe') 'acces directe: el desti es un EXECUTABLE (wscript.exe)'
Assert (-not ([string]$objAD.Desti -like '*.bat')) 'acces directe: mai un .bat (no es pot ancorar)'
AssertEq ([string]$objAD.Arguments) '"C:\clone\informes\suport\GenerarInforme.vbs"' 'acces directe: apunta al llancador sense consola'
Assert ([string]$objAD.Arguments).StartsWith('"') 'acces directe: l''argument va entre cometes (la ruta pot portar espais)'
AssertEq ([string]$objAD.Carpeta) 'C:\clone\informes' 'acces directe: la carpeta de treball es el clone'
AssertEq ([string]$objAD.Icona) 'C:\clone\informes\suport\cornella.ico' 'acces directe: porta l''escut'
# LA COPIA LOCAL DE L'ESCUT. El clone de l'usuari viu en una unitat de XARXA i
# l'explorador no es de fiar carregant icones d'alli per a un element ancorat:
# es queda amb la generica, que es el que li sortia.
$objAD2 = Get-AccesDirecteObjectiu 'C:\clone\informes' 'C:\Windows' 'C:\Users\x\AppData\Local\InformesCornella'
AssertEq ([string]$objAD2.Icona) 'C:\Users\x\AppData\Local\InformesCornella\cornella.ico' 'acces directe: l''escut, del disc local'
AssertEq ([string]$objAD2.IconaOrigen) 'C:\clone\informes\suport\cornella.ico' 'acces directe: ...copiat del clone'
AssertEq ([string]$objAD2.Arguments) ([string]$objAD.Arguments) 'acces directe: la copia de l''escut no canvia res mes'
AssertEq ([string](Get-AccesDirecteObjectiu 'C:\x' 'C:\Windows' 'C:\y\').Icona) 'C:\y\cornella.ico' 'acces directe: barra final de la carpeta de l''escut'

# LA ICONA DE LA FINESTRA. El .ico de l'Ajuntament porta TOTES les mides
# comprimides en PNG i el GDI+ no les sap descomprimir: "new Icon(path)" dona
# una icona BUIDA. Amb l'AppUserModelID propi, la barra de tasques va passar a
# fer servir la icona de la finestra... i va quedar sense escut.
$icoRepo = Join-Path (Split-Path -Parent $TestsDir) 'cornella.ico'
if (Test-Path -LiteralPath $icoRepo) {
    $rawIco = [System.IO.File]::ReadAllBytes($icoRepo)
    $frIco = _IcoTriaFrame $rawIco 32
    Assert ($null -ne $frIco) 'cornella.ico: es un .ico valid'
    Assert ([bool]$frIco.EsPng) 'cornella.ico: les imatges van comprimides en PNG (per aixo cal llegir-lo a ma)'
    AssertEq ([int]$frIco.Amplada) 32 'cornella.ico: per a 32 px s''agafa el frame de 32'
    Assert ([int](_IcoTriaFrame $rawIco 16).Amplada -eq 16) 'cornella.ico: ...i per a 16 px, el de 16'
    Assert ([int](_IcoTriaFrame $rawIco 1000).Amplada -eq 256) 'cornella.ico: si cap no hi arriba, la mes gran'
    AssertEq (_IcoTriaFrame @(1,2,3) 32) $null '_IcoTriaFrame: uns bytes que no son un .ico -> $null'
    # ...i la funcio que en fa la icona ha d'existir i venir de UiComuns.
    $srcUi2 = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $TestsDir) 'UiComuns.ps1') -Raw
    Assert ($srcUi2.Contains('function _IconaDeIco')) 'UiComuns: hi ha _IconaDeIco'
    Assert ($srcUi2.Contains('_IconaDeIco $iconPath 32')) 'UiComuns: la icona de l''app es fa amb ella (no amb new Icon)'
    Assert (-not ($srcUi2 -match 'AppIcon\s*=\s*New-Object System\.Drawing\.Icon')) 'UiComuns: ja no es fa servir el constructor pelat'
    $srcPdf2 = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $TestsDir) 'PdfSignar.ps1') -Raw
    Assert (-not ($srcPdf2.Contains('function _IcoTriaFrame'))) 'PdfSignar: _IcoTriaFrame ha passat a UiComuns (no hi es dues vegades)'
    Assert ($srcPdf2.Contains('_IcoTriaFrame $raw')) 'PdfSignar: ...i la segueix fent servir'
}
# L'IDENTIFICADOR D'APLICACIO. Es el que lliga la icona ANCORADA amb la finestra
# del programa: sense ell, la drecera ancorada surt sense icona i en obrir-la
# apareix un SEGON boto a la barra de tasques (va passar de debo).
Assert (-not [string]::IsNullOrWhiteSpace([string]$Script:AppUserModelId)) 'acces directe: hi ha un AppUserModelID'
# ...i esta escrit en UN SOL LLOC de tot suport/ (el proces i la drecera n'han
# de fer servir EXACTAMENT el mateix).
$dirSup = Split-Path -Parent $TestsDir
$nLit = 0
foreach ($fAD in (Get-ChildItem -LiteralPath $dirSup -Filter '*.ps1')) {
    $cAD = Get-Content -LiteralPath $fAD.FullName -Raw
    $nLit += @([regex]::Matches($cAD, [regex]::Escape("'" + [string]$Script:AppUserModelId + "'"))).Count
}
AssertEq $nLit 1 'acces directe: l''AppUserModelID nomes esta escrit una vegada'
# El proces se'l posa (UiComuns.ps1) i la drecera tambe (AccesDirecte.ps1).
$srcUi = Get-Content -LiteralPath (Join-Path $dirSup 'UiComuns.ps1') -Raw
Assert ($srcUi.Contains('SetCurrentProcessExplicitAppUserModelID')) 'acces directe: el proces es posa l''identificador'
Assert ($srcUi.Contains('$Script:AppUserModelId')) 'acces directe: ...i el llegeix del lloc unic'
$srcAD = Get-Content -LiteralPath (Join-Path $dirSup 'AccesDirecte.ps1') -Raw
Assert ($srcAD.Contains('9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3')) 'acces directe: PKEY_AppUserModel_ID'
Assert ($srcAD.Contains('Set-AccesDirecteAppId $ruta')) 'acces directe: es posa a CADA drecera creada'
# El C# de les interficies COM ha de COMPILAR (aqui tambe: nomes es compila,
# no es crida res del shell).
Assert ([bool](_AccesDirecteCarregaTipus)) 'acces directe: el C# de les interficies COM compila'
if ('CornellaApp.PropertyKey' -as [type]) {
    AssertEq ([System.Runtime.InteropServices.Marshal]::SizeOf([type]'CornellaApp.PropertyKey')) 20 'acces directe: PROPERTYKEY = GUID + int'
}
# Sense fitxer no peta i no s'inventa res.
AssertEq (Set-AccesDirecteAppId (Join-Path ([System.IO.Path]::GetTempPath()) 'no-hi-es-mai.lnk')) $false 'acces directe: si el .lnk no hi es, retorna fals'
# Una barra final al clone no ha de duplicar-se.
AssertEq ([string](Get-AccesDirecteObjectiu 'C:\clone\informes\' 'C:\Windows').Icona) 'C:\clone\informes\suport\cornella.ico' 'acces directe: la barra final del clone no es duplica'
# ...i sense SystemRoot es cau a C:\Windows (mai una ruta buida).
Assert ([string](Get-AccesDirecteObjectiu 'C:\x' '').Desti -like 'C:\Windows\*') 'acces directe: sense SystemRoot, C:\Windows'
# Es deixa a l'escriptori I al menu Inici (des d'alli es pot ancorar).
$destAD = @(Get-AccesDirecteDestins 'C:\Users\x\Desktop' 'C:\Users\x\Programs')
AssertEq $destAD.Count 2 'acces directe: escriptori i menu Inici'
Assert (-not (@($destAD) | Where-Object { -not ([string]$_).EndsWith('.lnk') })) 'acces directe: tots dos son .lnk'
AssertEq (@(Get-AccesDirecteDestins '' '').Count) 0 'acces directe: sense carpetes, cap desti'
# El .vbs al qual apunta ha d'existir de debo al repositori.
Assert (Test-Path -LiteralPath (Join-Path (Split-Path -Parent $TestsDir) 'GenerarInforme.vbs')) 'acces directe: el llancador .vbs hi es'
Assert (Test-Path -LiteralPath (Join-Path (Split-Path -Parent $TestsDir) 'cornella.ico')) 'acces directe: l''escut hi es'
# El .bat que el crea: ASCII pur (els .bat amb accents es trenquen segons la
# codepage) i sense cap '^' dins de cometes -dins de cometes el cmd el deixa
# passar LITERAL i el que arriba al PowerShell ja no es el que havies escrit-.
$batAD = Join-Path (Split-Path -Parent $TestsDir) 'Crear-acces-directe.bat'
Assert (Test-Path -LiteralPath $batAD) 'acces directe: hi ha suport\Crear-acces-directe.bat'
if (Test-Path -LiteralPath $batAD) {
    $txtAD = Get-Content -LiteralPath $batAD -Raw
    Assert (-not (@([char[]]$txtAD) | Where-Object { [int]$_ -gt 127 })) 'Crear-acces-directe.bat: ASCII pur'
    foreach ($lnAD in @($txtAD -split "`n")) {
        $iQ = $lnAD.IndexOf('"')
        if ($iQ -lt 0) { continue }
        $jQ = $lnAD.LastIndexOf('"')
        if ($jQ -le $iQ) { continue }
        Assert (-not ($lnAD.Substring($iQ, $jQ - $iQ).Contains('^'))) 'Crear-acces-directe.bat: cap ^ dins de cometes'
    }
    Assert ($txtAD.Contains('AccesDirecte.ps1')) 'Crear-acces-directe.bat: la feina la fa el .ps1'
}

# EL BOTO DE LA CARPETA del menu: la ruta NO pot estar escrita al codi, ha de
# sortir de _ResolveOutputDir (que es el que mana la Configuracio).
$srcMenu3 = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $TestsDir) 'Menu.ps1') -Raw
$iBtnC = $srcMenu3.IndexOf('$btnCarpeta = New-Object')
Assert ($iBtnC -ge 0) 'menu: hi ha el boto de la carpeta dels informes'
if ($iBtnC -ge 0) {
    $trosC = $srcMenu3.Substring($iBtnC, [Math]::Min(1800, $srcMenu3.Length - $iBtnC))
    Assert ($trosC.Contains('_ResolveOutputDir')) 'menu: la carpeta surt de la CONFIGURACIO, no del codi'
    Assert (-not ($trosC -match '[A-Z]:\\')) 'menu: cap ruta escrita al codi'
    Assert ($trosC.Contains('ConvertFromUtf32')) 'menu: l''emoji de carpeta es astral i va amb ConvertFromUtf32'
    # No tanca el menu: obrir una carpeta no es triar cap opcio.
    Assert (-not ($trosC.Substring(0, $trosC.IndexOf('$btnConfig')).Contains('$form.Close()'))) 'menu: obrir la carpeta NO tanca el menu'
}

# ---------------------------------------------------------------------------
# EL FORMAT DEL REQUERIMENT DE LLICENCIA (mesurat sobre el .docx fet a ma)
# ---------------------------------------------------------------------------
# F2: cap titol de seccio de Llicencia no acaba amb punt.
Assert (-not ((_LlicTitolAbans).TrimEnd().EndsWith('.')))   'F2: el titol del bloc ABANS no acaba amb punt'
Assert (-not ((_LlicTitolDespres).TrimEnd().EndsWith('.'))) 'F2: ...ni el del bloc DESPRES'
Assert ((_LlicTitolAbans).Contains('AMBIENTAL'))            'F2: ...pero el titol hi es sencer'

# F4: cap linia "No es disposa / Es disposa" no es queda sense punt final. Aixo
# vigila tambe el que s'hi escrigui de nou des de l'editor de catalegs.
$llicF4 = Read-LlicCataleg (Join-Path $Global:EstructuralsDir 'LLIC.json')
if ($null -ne $llicF4) {
    $senseFi = New-Object System.Collections.ArrayList
    foreach ($secF4 in @($llicF4.nodes)) {
        foreach ($itF4 in @($secF4.fills)) {
            foreach ($chF4 in @($itF4.fills)) {
                if ([string]$chF4.tipus -notin @('nodisposa', 'sidisposa')) { continue }
                foreach ($lF4 in @(_LlicCos $chF4)) {
                    $tF4 = ([string]$lF4).Trim()
                    if ([string]::IsNullOrWhiteSpace($tF4)) { continue }
                    if ($tF4.StartsWith('[[URL]]') -or $tF4.StartsWith('http')) { continue }
                    if ($tF4.EndsWith('.') -or $tF4.EndsWith(':') -or $tF4.EndsWith(')')) { continue }
                    [void]$senseFi.Add([string]$itF4.titol + ' / ' + [string]$chF4.tipus)
                }
            }
        }
    }
    AssertEq ($senseFi -join ' | ') '' 'F4: cap linia d''estat de LLIC.json es queda sense punt final'
}
Assert ((@((_LlicTextosPerDefecte).NoDisposa) -join ' ').Trim().EndsWith('.')) 'F4: ...ni la de per defecte'

# F1 i F3, sobre el punt sencer.
. (Join-Path $TestsDir 'FormatDoubles.ps1')
$puntF = [pscustomobject]@{
    Clau = 'S::A'; Subseccio = ''; Titol = 'Un punt'; Condicio = ''
    Cos = @('El cos del punt.')
    NoDisposa = @('No es disposa del document.')
    SiDisposa = @()
    Quan = @()
    # Un sub-punt que NOMES porta un enllac: no en surt cap pic.
    Subs = @(, @('[[URL]] https://exemple.cat/nomes-enllac'))
}
$global:emitCalls.Clear()
_LlicEscriuPunt $null $puntF '1.' ([ordered]@{}) 'no' $false
$emF = @($global:emitCalls)
# F1: la linia d'estat va SEPARADA (i en negreta, que ja hi era).
Assert ([bool]($emF | Where-Object { $_ -like 'BODY/N/SEP|No es disposa*' })) 'F1: la linia d''estat va separada del cos del punt'
# ...i nomes la PRIMERA linia del comentari.
$puntF2 = $puntF | Select-Object *
$puntF2.NoDisposa = @('Primera linia.', 'Segona linia.')
$global:emitCalls.Clear()
_LlicEscriuPunt $null $puntF2 '1.' ([ordered]@{}) 'no' $false
$sepF = @(@($global:emitCalls) | Where-Object { $_ -like '*/SEP|*' })
AssertEq $sepF.Count 1 'F1: nomes la primera linia del comentari va separada'
# F3: l'enllac d'un sub-punt SENSE text no va sagnat.
Assert (-not ($emF | Where-Object { $_ -like 'BULLET*' })) 'F3: un sub-punt sense text no emet cap pic'
Assert ([bool]($emF | Where-Object { $_ -eq 'URL|https://exemple.cat/nomes-enllac' })) 'F3: ...i el seu enllac NO va sagnat'
Assert (-not ($emF | Where-Object { $_ -like 'URL/CH|*' })) 'F3: ...cap enllac de fill'
# ...pero si el sub-punt TE text, el pic hi es i l'enllac si que va sagnat.
$puntF3 = $puntF | Select-Object *
$puntF3.Subs = @(, @('Text del sub-punt', '[[URL]] https://exemple.cat/amb-pic'))
$global:emitCalls.Clear()
_LlicEscriuPunt $null $puntF3 '1.' ([ordered]@{}) 'no' $false
$emF3 = @($global:emitCalls)
Assert ([bool]($emF3 | Where-Object { $_ -like 'BULLET/CH/1r|Text del sub-punt' })) 'F3: amb text, el sub-punt emet el seu pic'
Assert ([bool]($emF3 | Where-Object { $_ -eq 'URL/CH|https://exemple.cat/amb-pic' })) 'F3: ...i llavors l''enllac SI que va sagnat'

# ---------------------------------------------------------------------------
# Write-InformeDocx: la seqüencia d'obrir/escriure/desar, un sol cop
# ---------------------------------------------------------------------------
# Abans estava copiada a les quatre families (Build-Document, Build-ActExtrDocument,
# Build-LlicenciaDocument, Build-MnsDocument). Aqui es comprova el CONTRACTE:
# quin bloc de capcalera demana, que el cos rebi la Selection, i -l'unica cosa
# que canvia de comportament- que si el cos peta el document es TANCA (abans
# nomes ho feia ACT_EXTR; les altres tres deixaven el Word amb un document obert
# i el %TEMP% brut).
Write-Host "`n--- Write-InformeDocx (la seqüencia compartida) ---"
{
    $script:_widBloc = $null
    $script:_widTancat = 0
    $script:_widDesat = 0
    $selW = [pscustomobject]@{}
    $selW | Add-Member ScriptMethod EndKey { param($u) } -Force
    $docW = [pscustomobject]@{}
    $docW | Add-Member ScriptMethod Activate {} -Force
    $docW | Add-Member ScriptMethod Save { $script:_widDesat++ } -Force
    $docW | Add-Member ScriptMethod Close { param($x) $script:_widTancat++ } -Force
    $wordW = [pscustomobject]@{ Selection = $selW }
    function _ResolveOutputDir { return ([System.IO.Path]::GetTempPath()) }
    function _GetUniqueOutputPath($d, $b) { return (Join-Path $d $b) }
    function _OpenOutputDocument($w, $p) { return $script:_widDocProva }
    function Select-CapcaleraBlock($d, $w) { $script:_widBloc = [string]$w }
    function Apply-HeaderReplacements { param($doc, $header) }
    $script:_widDocProva = $docW
    $tempAbans = $env:TEMP
    if ([string]::IsNullOrWhiteSpace($env:TEMP)) { $env:TEMP = [System.IO.Path]::GetTempPath() }

    # Cami bo: el cos rep la Selection i el document es desa i es tanca un cop.
    $vist = $null
    $ruta = Write-InformeDocx $wordW 'prova-motor.docx' 'LLIC' @{ ID_GIA = '1' } {
        param($sel)
        $script:_widSel = $sel
    }
    $vist = $script:_widSel
    AssertEq $script:_widBloc 'LLIC' 'Write-InformeDocx: demana el bloc de capcalera que se li diu'
    Assert ($null -ne $vist) 'Write-InformeDocx: el cos rep la Selection'
    AssertEq $script:_widDesat 1 'Write-InformeDocx: desa el document un sol cop'
    AssertEq $script:_widTancat 1 'Write-InformeDocx: ...i el tanca un sol cop'
    Assert ([bool]([string]$ruta).EndsWith('prova-motor.docx')) 'Write-InformeDocx: retorna la ruta del document'

    # LA REGRESSIO DE DEBO: si el cos peta, el document s'ha de tancar igualment
    # i l'error ha d'arribar a qui ha cridat (no es pot empassar en silenci).
    $script:_widTancat = 0
    $script:_widDesat = 0
    $petada = $null
    try {
        [void](Write-InformeDocx $wordW 'prova-motor-2.docx' '' @{ ID_GIA = '2' } {
            param($sel)
            throw 'el cos ha petat'
        })
    } catch { $petada = $_ }
    Assert ($null -ne $petada) 'Write-InformeDocx: si el cos peta, l''error arriba a qui ha cridat'
    AssertEq $script:_widTancat 1 'Write-InformeDocx: ...i el document es tanca igualment'
    AssertEq $script:_widDesat 0 'Write-InformeDocx: ...sense desar-lo'
    $env:TEMP = $tempAbans
}.Invoke() | Out-Null

# I QUE NO TORNI A APAREIXER LA COPIA: cap fitxer fora de MotorInforme.ps1 pot
# obrir el document de sortida pel seu compte. Es la prova que evita que la
# seqüencia es torni a duplicar a la quinta familia d'informe.
$srcMotorDir = Join-Path (Split-Path -Parent $TestsDir) ''
$obren = @()
foreach ($f in @(Get-ChildItem -Path $srcMotorDir -Filter '*.ps1' -File)) {
    if ($f.Name -eq 'MotorInforme.ps1') { continue }
    $txt = Get-Content -LiteralPath $f.FullName -Raw
    # Nomes les CRIDES, no les mencions als comentaris.
    foreach ($ln in ($txt -split "`r?`n")) {
        if ($ln.TrimStart().StartsWith('#')) { continue }
        if ($ln -match '_OpenOutputDocument\s+\$') { $obren += ($f.Name + ': ' + $ln.Trim()) }
    }
}
AssertEq $obren.Count 0 ('Cap informe obre el document pel seu compte (nomes Write-InformeDocx)' + $(if ($obren.Count) { ' -> ' + ($obren -join ' | ') } else { '' }))

# ---------------------------------------------------------------------------
# L'aire entre blocs: una bandera, un sol lloc
# ---------------------------------------------------------------------------
# Abans hi havia TRENTA-QUATRE "if ($cfg.SpacerAfterX) { Format-Spacer $sel }"
# escampats per Document / Llicencia / MnsTraspas / VistaWord. Ara es
# Format-Aire $sel '<clau>' i la bandera es resol en un sol lloc.
Write-Host "`n--- Format-Aire (l'aire entre blocs) ---"
Assert (Test-FormatAire 'seccio')        'Aire: la clau "seccio" mira SpacerAfterSection'
Assert (Test-FormatAire 'subseccio')     'Aire: "subseccio" -> SpacerAfterSubsection'
Assert (Test-FormatAire 'item')          'Aire: "item" -> SpacerAfterItem'
Assert (Test-FormatAire 'intro')         'Aire: "intro" -> SpacerAfterIntro'
Assert (Test-FormatAire 'introparagraf') 'Aire: "introparagraf" -> SpacerAfterIntroParagraph'
Assert (Test-FormatAire 'conclusions')   'Aire: "conclusions" -> SpacerBeforeConclusionsBlock'
Assert (Test-FormatAire 'SECCIO')        'Aire: la clau no distingeix majuscules'
# Una clau desconeguda NO ha de posar aire "per si de cas" ni petar: un nom mal
# escrit no pot afegir una linia en blanc a un informe sense que ningu ho vegi.
Assert (-not (Test-FormatAire 'aixo-no-existeix')) 'Aire: una clau desconeguda no posa aire'
Assert (-not (Test-FormatAire ''))                 'Aire: una clau buida tampoc'
# I la bandera MANA: si es posa a fals, l'aire desapareix.
$aireAbans = $Script:ReportFormatConfig.SpacerAfterItem
try {
    $Script:ReportFormatConfig.SpacerAfterItem = $false
    Assert (-not (Test-FormatAire 'item')) 'Aire: amb la bandera a fals, no hi va aire'
} finally { $Script:ReportFormatConfig.SpacerAfterItem = $aireAbans }
Assert (Test-FormatAire 'item') 'Aire: ...i en tornar-la a posar, si'

# CAP INFORME LLEGEIX LES BANDERES PEL SEU COMPTE. Es la prova que impedeix que
# tornin a apareixer els 34 "if ($cfg.SpacerAfterX)" escampats: si algu n'escriu
# un de nou, aqui salta.
$srcAireDir = Split-Path -Parent $TestsDir
$llegeixen = @()
foreach ($f in @(Get-ChildItem -Path $srcAireDir -Filter '*.ps1' -File)) {
    if ($f.Name -eq 'Format.ps1') { continue }   # es qui les defineix
    foreach ($ln in ((Get-Content -LiteralPath $f.FullName -Raw) -split "`r?`n")) {
        if ($ln.TrimStart().StartsWith('#')) { continue }
        if ($ln -match '\$cfg\.Spacer|ReportFormatConfig\.Spacer') { $llegeixen += ($f.Name + ': ' + $ln.Trim()) }
    }
}
AssertEq $llegeixen.Count 0 ('Cap informe llegeix les banderes d''aire pel seu compte' + $(if ($llegeixen.Count) { ' -> ' + ($llegeixen -join ' | ') } else { '' }))

# ---------------------------------------------------------------------------
# Write-Linia: escriure una linia de cataleg, un sol cop
# ---------------------------------------------------------------------------
# N'hi havia tres copies ($emitLine a Document.ps1, _LlicEmetLinia a
# Llicencia.ps1 i _VLine a VistaWord.ps1). Les dues primeres nomes es
# diferenciaven en si deduplicaven els enllacos.
Write-Host "`n--- Write-Linia (text + enllacos) ---"
{
    . (Join-Path $TestsDir 'FormatDoubles.ps1')
    $selL = [pscustomobject]@{}
    $U1 = 'https://exemple.cat/u1'
    $U2 = 'https://exemple.cat/u2'

    # Text i enllac: el text va com a cos i l'enllac en paragraf propi.
    # (El prefix '[[URL]] ' nomes val quan es TOTA la linia; enmig d'un text,
    # l'URL es detecta pel contingut. Vegeu _SplitTextAndUrls.)
    Write-Linia $selL ('Cal presentar el document. ' + $U1)
    $e = @($global:emitCalls)
    AssertEq $e.Count 2 'Write-Linia: text + enllac son dos paragrafs'
    Assert ([bool]($e[0] -eq 'BODY|Cal presentar el document.')) 'Write-Linia: el text va com a cos'
    Assert ([bool]($e[1] -eq ('URL|' + $U1)))                    'Write-Linia: ...i l''enllac despres'

    # -IsChild: cos i enllac sagnats.
    $global:emitCalls.Clear()
    Write-Linia $selL ('Sub-punt. ' + $U1) -IsChild
    $e = @($global:emitCalls)
    Assert ([bool]($e[0] -eq 'BODY/CH|Sub-punt.')) 'Write-Linia: -IsChild sagna el cos'
    Assert ([bool]($e[1] -eq ('URL/CH|' + $U1)))   'Write-Linia: ...i l''enllac'

    # Una linia buida no escriu res (ni un paragraf en blanc).
    $global:emitCalls.Clear()
    Write-Linia $selL '   '
    AssertEq @($global:emitCalls).Count 0 'Write-Linia: una linia buida no escriu res'

    # Un enllac SOL (sense text) no emet cap cos buit.
    $global:emitCalls.Clear()
    Write-Linia $selL ('[[URL]] ' + $U1)
    $e = @($global:emitCalls)
    AssertEq $e.Count 1 'Write-Linia: un enllac sol (prefix [[URL]]) no emet cap cos buit'
    Assert ([bool]($e[0] -eq ('URL|' + $U1))) 'Write-Linia: ...nomes l''enllac'

    # DEDUPLICACIO: nomes quan se li passa el conjunt $vistos (Llicencia). Amb
    # $null (REQ1) no es dedupa res, que es el comportament de sempre.
    $global:emitCalls.Clear()
    Write-Linia $selL ('A. ' + $U1)
    Write-Linia $selL ('B. ' + $U1)
    AssertEq @(@($global:emitCalls) | Where-Object { $_ -like 'URL|*' }).Count 2 'Write-Linia: sense $vistos, l''enllac repetit surt dues vegades'

    $vistos = New-Object System.Collections.Generic.HashSet[string]
    $global:emitCalls.Clear()
    Write-Linia $selL ('A. ' + $U1) $vistos
    Write-Linia $selL ('B. ' + $U1) $vistos
    Write-Linia $selL ('C. ' + $U2) $vistos
    $urls = @(@($global:emitCalls) | Where-Object { $_ -like 'URL|*' })
    AssertEq $urls.Count 2 'Write-Linia: amb $vistos, l''enllac repetit NO es torna a escriure'
    Assert ([bool]($urls[1] -eq ('URL|' + $U2))) 'Write-Linia: ...i el nou si'
    # ...pero els tres COSSOS hi son: dedupliquem enllacos, no text.
    AssertEq @(@($global:emitCalls) | Where-Object { $_ -like 'BODY|*' }).Count 3 'Write-Linia: el text no es dedupica mai'

    # $emesos: on s'apunta el que s'ha arribat a escriure (Llicencia el fa servir
    # per saber quins enllacos ha de deixar per DESPRES del comentari).
    $vistos2 = New-Object System.Collections.Generic.HashSet[string]
    $emesos = New-Object System.Collections.Generic.HashSet[string]
    $global:emitCalls.Clear()
    Write-Linia $selL ('A. ' + $U1) $vistos2 $emesos
    Assert ($emesos.Contains($U1)) 'Write-Linia: $emesos recull els enllacos escrits'
}.Invoke() | Out-Null

# ---------------------------------------------------------------------------
# L'alineat i els espais surten de la CONFIGURACIO, no de literals
# ---------------------------------------------------------------------------
# Hi havia tres "Alignment = 3" escrits a Format-Note, Format-Label i
# Format-Conclusion, que _Apply-Indent ja hi posava: amb ells, canviar
# BodyAlignment no els tocava. L'unic literal que hi pot quedar es el centrat
# del titol CONCLUSIONS, que es una decisio propia d'aquell paragraf.
Write-Host "`n--- Format.ps1: cap alineat ni espai escrit a pel ---"
$fmtSrc = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $TestsDir) 'Format.ps1') -Raw
$literals = @()
foreach ($ln in ($fmtSrc -split "`r?`n")) {
    if ($ln.TrimStart().StartsWith('#')) { continue }
    if ($ln -notmatch 'ParagraphFormat\.Alignment\s*=') { continue }
    # 1 = centrat: nomes el titol de conclusions, i hi ha de ser.
    if ($ln -match 'ParagraphFormat\.Alignment\s*=\s*1\b') { continue }
    if ($ln -match 'ParagraphFormat\.Alignment\s*=\s*\d') { $literals += $ln.Trim() }
}
AssertEq $literals.Count 0 ('Format.ps1: cap alineat escrit a pel' + $(if ($literals.Count) { ' -> ' + ($literals -join ' | ') } else { '' }))
AssertEq ([int]$Script:ReportFormatConfig.ConclusionHeaderSpaceAfterPt) 12 'L''espai sota CONCLUSIONS es configurable (i val 12 pt)'
# ...i va A PART del de cada conclusio: son dues decisions diferents.
Assert ($null -ne $Script:ReportFormatConfig.ConclusionSpaceAfterPt) 'L''espai entre conclusions segueix sent el seu'

# Les anotacions de SEGUIMENT llegeixen els valors de Format.ps1 i les seves
# conversions, no una copia. Abans en tenien dues (els valors per defecte i els
# factors 1440/2,54 i 20) i es podien desfasar en silenci.
Write-Host "`n--- Seguiment: el format de l'anotacio ve de Format.ps1 ---"
$fmtAnot = _AnnotationFormatTwips
AssertEq $fmtAnot.Indent      (_CmToTwips $Script:ReportFormatConfig.AnnotationIndentCm)      'Anotacio: la sangria surt de la configuracio'
AssertEq $fmtAnot.SpaceBefore (_PtToTwips $Script:ReportFormatConfig.AnnotationSpaceBeforePt) 'Anotacio: l''espai de sobre tambe'
AssertEq $fmtAnot.SpaceAfter  (_PtToTwips $Script:ReportFormatConfig.AnnotationSpaceAfterPt)  'Anotacio: i el de sota'
# ...i canviar-la a Format.ps1 hi arriba.
$indAbans = $Script:ReportFormatConfig.AnnotationIndentCm
try {
    $Script:ReportFormatConfig.AnnotationIndentCm = 2.0
    AssertEq (_AnnotationFormatTwips).Indent (_CmToTwips 2.0) 'Anotacio: canviar-ho a Format.ps1 hi arriba'
} finally { $Script:ReportFormatConfig.AnnotationIndentCm = $indAbans }

# ---------------------------------------------------------------------------
# NOMES Format.ps1 toca el Word per FORMAT
# ---------------------------------------------------------------------------
# Es la invariant que fa que "canviar-ho en un lloc" sigui veritat: si un altre
# fitxer posa una sagnia, un salt de pagina o un nivell d'esquema pel seu
# compte, aquell tros de format queda fora de $ReportFormatConfig i ja no es
# pot canviar des d'un sol lloc. Ja va passar amb el salt de pagina de l'ANNEX 1
# i amb l'OutlineLevel de les vistes.
Write-Host "`n--- Nomes Format.ps1 toca el Word per format ---"
$dirFmt = Split-Path -Parent $TestsDir
$toquen = @()
foreach ($f in @(Get-ChildItem -Path $dirFmt -Filter '*.ps1' -File)) {
    if ($f.Name -eq 'Format.ps1') { continue }
    $i = 0
    foreach ($ln in ((Get-Content -LiteralPath $f.FullName -Raw) -split "`r?`n")) {
        $i++
        if ($ln.TrimStart().StartsWith('#')) { continue }
        if ($ln -match 'ParagraphFormat\.|\.OutlineLevel|InsertBreak\(') {
            $toquen += ($f.Name + ':' + $i + ': ' + $ln.Trim())
        }
    }
}
AssertEq $toquen.Count 0 ('Cap fitxer fora de Format.ps1 toca el Word per format' + $(if ($toquen.Count) { ' -> ' + ($toquen -join ' | ') } else { '' }))

# ---------------------------------------------------------------------------
# El motor de blocs: Build-CatalegBlocs (pura) i Write-Informe
# ---------------------------------------------------------------------------
Write-Host "`n--- Build-CatalegBlocs (pura: sense Word ni dobles) ---"
# Un cataleg de mentida amb tot el que compta: una seccio, una subseccio buida
# (que NO ha de sortir), una amb item, un item amb fills i un item no triat.
function _NouEl([string]$kind, [string]$short, $linies, $fills, [bool]$triat) {
    return [pscustomobject]@{
        Kind = $kind; Short = $short
        BodyLines = @($linies); Children = @($fills); Selected = $triat
    }
}
$secP = [pscustomobject]@{
    Title = 'Instal·lacions'
    Items = @(
        (_NouEl 'subsection' 'Subseccio BUIDA' @() @() $false),
        (_NouEl 'subsection' 'Legalitzacions'  @() @() $false),
        (_NouEl 'item' 'A' @('Primer punt.') @() $true),
        (_NouEl 'item' 'B' @('Punt amb fills.') @(
            (_NouEl 'subitem' 'f1' @('Fill u.') @() $true),
            (_NouEl 'subitem' 'f2' @('Fill dos.') @() $true)
        ) $true),
        (_NouEl 'item' 'C' @('No triat.') @() $false)
    )
}
$bl = @(Build-CatalegBlocs @($secP) $null '' $false @() -SenseCamps)
$tipus = @($bl | ForEach-Object { [string]$_.T })
AssertEq ($tipus -join ',') 'seccio,aire,subseccio,aire,unitat,unitat' 'Blocs: seccio + subseccio pendent + dues unitats'
# LA SUBSECCIO BUIDA NO HI ES: nomes surt la que va seguida d'un item de debo.
Assert (-not (@($bl) | Where-Object { [string]$_.Text -eq 'Subseccio BUIDA' })) 'Blocs: una subseccio sense items no surt'
Assert ([bool](@($bl) | Where-Object { [string]$_.Text -eq 'Legalitzacions' })) 'Blocs: la que en te, si'
# L'item no triat i sense fills tampoc.
$nums = @(@($bl) | Where-Object { $_.T -eq 'unitat' } | ForEach-Object { [string]@($_.Blocs)[0].Num })
AssertEq ($nums -join ',') '1.,2.' 'Blocs: la numeracio va seguida i salta el que no es tria'
# Els fills son PICS, no items numerats.
$dins2 = @(@($bl)[-1].Blocs | ForEach-Object { [string]$_.T })
AssertEq ($dins2 -join ',') 'item,pic,pic' 'Blocs: els fills van amb pic, no numerats'

# Write-Informe: el -First del primer sub-punt el posa EL MOTOR.
Write-Host "`n--- Write-Informe (l'aire i el -First els decideix el motor) ---"
{
    . (Join-Path $TestsDir 'FormatDoubles.ps1')
    $selM = [pscustomobject]@{}
    [void](Write-Informe $selM $bl)
    $em = @($global:emitCalls)
    Assert ([bool]($em | Where-Object { $_ -eq 'BULLET/CH/1r|Fill u.' })) 'Motor: el primer sub-punt d''una unitat va a 12 pt'
    Assert ([bool]($em | Where-Object { $_ -eq 'BULLET/CH|Fill dos.' }))  'Motor: el segon es queda a 6 pt'
    # L'aire d'item va DESPRES de la unitat sencera (amb els seus fills).
    $iFill = [Array]::IndexOf([string[]]$em, 'BULLET/CH|Fill dos.')
    AssertEq $em[$iFill + 1] 'AIRE|item' 'Motor: l''aire va despres de la unitat SENCERA'

    # Una unitat que no escriu res NO deixa aire darrere.
    $global:emitCalls.Clear()
    [void](Write-Informe $selM @(@{ T = 'unitat'; Blocs = @() }))
    AssertEq @($global:emitCalls).Count 0 'Motor: una unitat buida no deixa aire'

    # Un tipus de bloc desconegut PETA: val mes que generar un document al qual
    # li falta un tros sense que ho digui ningu.
    $petat = $false
    try { [void](Write-Informe $selM @(@{ T = 'aixo-no-existeix' })) } catch { $petat = $true }
    Assert $petat 'Motor: un tipus de bloc desconegut peta (no s''ho empassa)'
}.Invoke() | Out-Null

# ---------------------------------------------------------------------------
# Una sola linia en blanc entre la capcalera i el cos
# ---------------------------------------------------------------------------
# La plantilla no en porta les mateixes a cada bloc (el generic i el de LLIC en
# tenien DUES despres d'"INFORME", el d'ACT_EXTR una), i per aixo uns informes
# sortien amb un forat mes gros que els altres.
Write-Host "`n--- Capcalera: una sola linia en blanc al final ---"
AssertEq (_CapBlancsQueSobren @('INFORME', '', '')) 1 'Capcalera: de dos blancs, en sobra un'
AssertEq (_CapBlancsQueSobren @('INFORME', '')) 0 'Capcalera: amb un, no en sobra cap'
AssertEq (_CapBlancsQueSobren @('INFORME')) 0 'Capcalera: sense cap, tampoc'
AssertEq (_CapBlancsQueSobren @('INFORME', '', '', '')) 2 'Capcalera: de tres, en sobren dos'
# El Word acaba cada paragraf amb \r (i les cel·les amb \a): han de comptar com a buits.
AssertEq (_CapBlancsQueSobren @("INFORME`r", "`r", "   `r")) 1 'Capcalera: el \r del Word no fa que un blanc sembli text'
AssertEq (_CapBlancsQueSobren @()) 0 'Capcalera: sense paragrafs, res a fer'
# I que no es mengi text de debo.
AssertEq (_CapBlancsQueSobren @('', 'INFORME')) 0 'Capcalera: si l''ultim porta text, no en sobra cap'

# SOBRE LA PLANTILLA REAL: cada bloc ha d'acabar amb un sol blanc DESPRES de
# normalitzar. Es la comprovacio que lliga la regla al fitxer de debo.
$capBlancPath = Join-Path $EstructuralsDir '0 CAPCALERA.docx'
if (Test-Path -LiteralPath $capBlancPath) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $zipB = [System.IO.Compression.ZipFile]::OpenRead($capBlancPath)
    try {
        $entB = $zipB.GetEntry('word/document.xml')
        $srB = New-Object System.IO.StreamReader($entB.Open())
        $xmlB = $srB.ReadToEnd(); $srB.Close()
    } finally { $zipB.Dispose() }
    # Els paragrafs, en ordre, amb el seu text.
    $parasB = @([regex]::Matches($xmlB, '<w:p\b(?:[^>]*/>|[^>]*>.*?</w:p>)', 'Singleline') | ForEach-Object {
        (([regex]::Matches($_.Value, '<w:t[^>]*>(.*?)</w:t>', 'Singleline') | ForEach-Object { $_.Groups[1].Value }) -join '')
    })
    # Cada bloc va d'un marcador [[CAP:x]] al seguent; el generic, del principi
    # al primer marcador.
    $tallsB = @(0)
    for ($i = 0; $i -lt $parasB.Count; $i++) {
        if ((_CapMarcador $parasB[$i]) -ne '') { $tallsB += ($i + 1) }
    }
    $blocsB = @()
    for ($k = 0; $k -lt $tallsB.Count; $k++) {
        $ini = $tallsB[$k]
        $fi = if ($k -lt ($tallsB.Count - 1)) { $tallsB[$k + 1] - 1 } else { $parasB.Count }
        $blocsB += ,@($parasB[$ini..($fi - 1)])
    }
    AssertEq $blocsB.Count 3 'Capcalera: la plantilla te tres blocs (generic, ACT_EXTR, LLIC)'
    foreach ($b in $blocsB) {
        $q = @($b | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
        $darrer = if ($q.Count) { [string]$q[$q.Count - 1] } else { '' }
        # Cada bloc acaba amb "INFORME" i despres nomes blancs.
        Assert ([bool]($darrer.Trim() -eq 'INFORME')) 'Capcalera: el bloc acaba amb INFORME'
    }
    # I despres de normalitzar, cap bloc no pot quedar amb mes d'un blanc.
    foreach ($b in $blocsB) {
        $sobren = _CapBlancsQueSobren $b
        $blancs = 0
        for ($i = @($b).Count - 1; $i -ge 0; $i--) {
            if ([string]::IsNullOrWhiteSpace([string]@($b)[$i])) { $blancs++ } else { break }
        }
        AssertEq ($blancs - $sobren) 1 'Capcalera: el bloc queda amb UNA sola linia en blanc'
    }
}

# ---------------------------------------------------------------------------
# EL TEXT FIX D'UNA SUBSECCIO SURT ENCARA QUE NO ES TRIÏ EL PRIMER PUNT
# ---------------------------------------------------------------------------
# Defecte real (GIA 1484): l'intro de Instal·lacions / Legalitzacions -"Segons
# l'article 4 de l'Ordenanca..."- penja del primer item del cataleg
# ("RITSIC - fotovoltaica"). En un informe que nomes demanava la baixa tensio i
# el PCI, el text fix NO sortia. A REQ1 no passa perque alli l'intro queda
# PENDENT fins que surt un item, sigui quin sigui.
Write-Host "`n--- Llicencia: el text fix d'una subseccio, com a REQ1 ---"
# _LlicItemsAmbUbicacio: TOTS els items del grup es porten el mateix intro.
$secX = [pscustomobject]@{
    Title = 'Instal·lacions'
    Items = @(
        [pscustomobject]@{ Kind = 'subsection'; Short = 'Legalitzacions'; BodyLines = @(); Children = @() },
        [pscustomobject]@{ Kind = 'intro'; Short = ''; BodyLines = @('Segons l''article 4:'); Children = @() },
        [pscustomobject]@{ Kind = 'item'; Short = 'A'; BodyLines = @('Punt A.'); Children = @() },
        [pscustomobject]@{ Kind = 'item'; Short = 'B'; BodyLines = @('Punt B.'); Children = @() },
        [pscustomobject]@{ Kind = 'subsection'; Short = 'Inspeccions'; BodyLines = @(); Children = @() },
        [pscustomobject]@{ Kind = 'item'; Short = 'C'; BodyLines = @('Punt C.'); Children = @() }
    )
}
$ubX = @(_LlicItemsAmbUbicacio @($secX))
AssertEq $ubX.Count 3 'Ubicacio: tres items'
AssertEq (@($ubX[0].Intro) -join '') 'Segons l''article 4:' 'Ubicacio: el PRIMER item del grup porta l''intro'
AssertEq (@($ubX[1].Intro) -join '') 'Segons l''article 4:' 'Ubicacio: ...i el SEGON tambe (si no es tria el primer, es perdria)'
AssertEq (@($ubX[2].Intro) -join '') '' 'Ubicacio: una subseccio nova buida l''intro'

# I l'escriptura: amb NOMES el segon punt, el text fix ha de sortir igualment.
{
    . (Join-Path $TestsDir 'FormatDoubles.ps1')
    $selX = [pscustomobject]@{}
    $puntB = [pscustomobject]@{
        Clau = 'Instal·lacions::B'; Seccio = 'Instal·lacions'; Subseccio = 'Legalitzacions'
        Intro = @('Segons l''article 4:'); Titol = 'B'; Condicio = ''
        Cos = @('Punt B.'); NoDisposa = @(); SiDisposa = @(); Quan = @(); Subs = @(); Estat = ''
    }
    $nX = 0
    _LlicEscriuPunts $selX @($puntB) ([ref]$nX) ([ordered]@{}) $false
    $eX = @($global:emitCalls)
    Assert ([bool]($eX | Where-Object { $_ -eq 'SECT|Instal·lacions' }))   'Text fix: hi surt la seccio'
    Assert ([bool]($eX | Where-Object { $_ -eq 'SUB|Legalitzacions' }))    'Text fix: hi surt la subseccio'
    Assert ([bool]($eX | Where-Object { $_ -eq 'BODY|Segons l''article 4:' })) 'Text fix: ...I EL TEXT FIX, encara que no s''hagi triat el primer punt'
    # I en l'ordre bo: seccio, subseccio, text fix, punt.
    $ordre = @($eX | Where-Object { $_ -notlike 'AIRE|*' })
    AssertEq $ordre[0] 'SECT|Instal·lacions' 'Text fix: primer la seccio'
    AssertEq $ordre[1] 'SUB|Legalitzacions'  'Text fix: despres la subseccio'
    AssertEq $ordre[2] 'BODY|Segons l''article 4:' 'Text fix: despres el text fix'
    Assert ([bool]($ordre[3] -like 'ITEM|*')) 'Text fix: i despres el punt'

    # NO es repeteix a cada punt del grup.
    $global:emitCalls.Clear()
    $puntA = $puntB | Select-Object *
    $puntA.Titol = 'A'; $puntA.Cos = @('Punt A.')
    $nX = 0
    _LlicEscriuPunts $selX @($puntA, $puntB) ([ref]$nX) ([ordered]@{}) $false
    AssertEq @(@($global:emitCalls) | Where-Object { $_ -eq 'BODY|Segons l''article 4:' }).Count 1 'Text fix: surt UN sol cop per grup'
}.Invoke() | Out-Null

# ---------------------------------------------------------------------------
# TOTS els informes que llegeixen REQ1 treuen els seus TEXTOS FIXOS
# ---------------------------------------------------------------------------
# El defecte del GIA 1484 (l'intro d'una subseccio que no sortia si no es
# triava el PRIMER punt del grup) es va comprovar nomes a Llicencia, i encara
# hi havia una familia mes amb el mateix problema: la VISTA de LLIC.
#
# Per aixo aquesta prova no mira UN informe: recorre TOTES les families que
# poden portar punts de REQ1 i, a cada una, tria a posta el SEGON punt de cada
# grup amb text fix i MAI el primer. Si algun dia s'afegeix una familia nova,
# nomes cal afegir-la a la llista d'aqui sota.
Write-Host "`n--- Els textos fixos de REQ1, a TOTES les families ---"
$req1TF = $null
try { $req1TF = Read-CatalegJson (Join-Path $EstructuralsDir 'REQ1.json') } catch { }
$llicTF = $null
try { $llicTF = Read-LlicCataleg (Join-Path $EstructuralsDir 'LLIC.json') } catch { }
if ($null -ne $req1TF -and $null -ne $llicTF) {
    . (Join-Path $TestsDir 'FormatDoubles.ps1')

    # Els grups amb text fix i com a minim DOS punts, i la clau del SEGON.
    $grupsTF = New-Object System.Collections.ArrayList
    foreach ($sec in $req1TF.Sections) {
        $sub = ''; $intro = @(); $items = @()
        $tanca = {
            if (@($intro).Count -gt 0 -and @($items).Count -ge 2) {
                [void]$grupsTF.Add(@{ Sec = [string]$sec.Title; Sub = $sub; Intro = $intro; Segon = $items[1] })
            }
        }
        foreach ($el in $sec.Items) {
            if ([string]$el.Kind -eq 'subsection') { & $tanca; $sub = [string]$el.Short; $intro = @(); $items = @(); continue }
            if ([string]$el.Kind -eq 'intro')      { $intro = @($el.BodyLines); continue }
            if ([string]$el.Kind -eq 'item')       { $items += $el }
        }
        & $tanca
    }
    Assert ($grupsTF.Count -ge 1) ('Textos fixos: REQ1 en te ' + $grupsTF.Count + ' de comprovables')
    $clausTF = @($grupsTF | ForEach-Object { _ItemKey $_.Sec $_.Segon.Short })
    $selTF   = Build-SelectionFromKeys $req1TF.Sections $clausTF

    # Hi es, el text fix de cada grup, a la seqüencia emesa?
    #
    # $grups es OPCIONAL i per defecte son TOTS: les families que componen amb
    # _WriteCatalegBody porten qualsevol punt de REQ1. Llicencia, en canvi, NO
    # es queda totes les subseccions -nomes expandeix les que digui LLIC.json-,
    # o sigui que alli s'hi passen NOMES els grups que aquell bloc porta. Si
    # s'exigissin tots, afegir una subseccio nova a REQ1 faria petar aquesta
    # prova sense que hi hagues cap defecte.
    $comprova = {
        param($nom, $crides, $grups = $null)
        if ($null -eq $grups) { $grups = $grupsTF }
        foreach ($g in $grups) {
            $t = ([string]@($g.Intro)[0])
            $cap = $t.Substring(0, [Math]::Min(30, $t.Length))
            $hi = [bool](@($crides) | Where-Object { $_ -like ('*' + $cap + '*') })
            Assert $hi ($nom + ': hi surt el text fix de "' + $g.Sub + '"')
        }
    }

    # Word i entorn de mentida (nomes el que demanen els Build-*).
    $sdTF = [pscustomobject]@{ Range = [pscustomobject]@{ Start = 0; End = 0 } }
    $sdTF | Add-Member ScriptMethod EndKey { param($u) } -Force
    $sdTF | Add-Member ScriptMethod InsertBreak { param($b) } -Force
    $ddTF = [pscustomobject]@{}
    $ddTF | Add-Member ScriptMethod Activate {} -Force
    $ddTF | Add-Member ScriptMethod Save {} -Force
    $ddTF | Add-Member ScriptMethod Close { param($x) } -Force
    $wdTF = [pscustomobject]@{ Selection = $sdTF }
    $script:_docTF = $ddTF
    function _ResolveOutputDir { return ([System.IO.Path]::GetTempPath()) }
    function _GetUniqueOutputPath($d, $b) { return (Join-Path $d $b) }
    function _OpenOutputDocument($w, $p) { return $script:_docTF }
    function Select-CapcaleraBlock($d, $w) { }
    function Apply-HeaderReplacements { param($doc, $header) }
    $tmpTF = $env:TEMP
    if ([string]::IsNullOrWhiteSpace($env:TEMP)) { $env:TEMP = [System.IO.Path]::GetTempPath() }
    $hdrTF = @{ ID_GIA = '1'; TITULAR = 'X'; CLASSIFICACIO = 'Y' }

    # 1) REQ1 / TERMINI / Controls periodics / Paquet: tots componen amb
    #    _WriteCatalegBody, o sigui que provant-lo els cobreix tots.
    $global:emitCalls.Clear()
    _WriteCatalegBody $sdTF $Script:ReportFormatConfig $selTF ([ordered]@{}) ''
    & $comprova 'REQ1 (i Controls periodics i el paquet del mobil)' $global:emitCalls

    # 2) MNS i TRASPAS
    $mnsTF = $null
    try { $mnsTF = _LoadEstructuralJson (Join-Path $EstructuralsDir 'MNSTRAS.json') } catch { }
    if ($null -ne $mnsTF) {
        foreach ($f in @('mns', 'traspas')) {
            $global:emitCalls.Clear()
            [void](Build-MnsDocument $wdTF @{ Fase = $f; Header = $hdrTF; Fields = [ordered]@{}; Cataleg = $mnsTF; Punts = $selTF })
            & $comprova ('MNS/' + $f) $global:emitCalls
        }
    }

    # 3) LLICENCIA, les tres fases. Es tria el SEGON punt de cada grup del bloc
    #    DESPRES, mai el primer -que es exactament el cas que fallava.
    $idxTF = _LlicIndexReq1 $req1TF
    $totsTF = @((_LlicPuntsPerBloc $llicTF $idxTF 'DESPRES' $req1TF).Punts)
    $vistTF = @{}; $despTF = @()
    foreach ($p in $totsTF) {
        if (@($p.Intro).Count -eq 0) { continue }
        $k = [string]$p.Seccio + '::' + [string]$p.Subseccio
        if (-not $vistTF.ContainsKey($k)) { $vistTF[$k] = 1; continue }   # SALTA el primer
        if ($vistTF[$k] -eq 1) { $vistTF[$k] = 2; $despTF += $p }
    }
    Assert (@($despTF).Count -ge 1) 'Textos fixos: hi ha punts de Llicencia per comprovar (mai el primer del grup)'
    # Els grups que el bloc DESPRES porta de debo: Llicencia nomes expandeix les
    # subseccions que li diu LLIC.json, i exigir-hi la resta seria demanar un
    # text que aquell informe no ha d'escriure.
    $ubiTF = @{}
    foreach ($p in $despTF) { $ubiTF[([string]$p.Seccio + '::' + [string]$p.Subseccio)] = $true }
    $grupsLlicTF = @($grupsTF | Where-Object { $ubiTF.ContainsKey([string]$_.Sec + '::' + [string]$_.Sub) })
    Assert (@($grupsLlicTF).Count -ge 1) 'Textos fixos: el bloc DESPRES de Llicencia en porta algun'
    foreach ($f in @('requeriment', 'favorable-pre', 'favorable-post')) {
        $dTF = @(@(_LlicPuntsAmbEstatFase $despTF $f) | ForEach-Object { $_ | Add-Member NoteProperty Estat 'no' -PassThru -Force })
        $global:emitCalls.Clear()
        [void](Build-LlicenciaDocument $wdTF @{
            Fase = $f; EsProvisional = $false; Header = $hdrTF; Fields = [ordered]@{}
            Abans = @(); Projecte = @(); Despres = $dTF
            Doc = @{ Text = ''; Items = @() }; Condicions = ''; Cataleg = $llicTF
        })
        & $comprova ('Llicencia/' + $f) $global:emitCalls $grupsLlicTF
    }

    # 4) LES VISTES en Word. Han d'ensenyar el mateix que el document: aqui hi
    #    havia el defecte que la comprovacio d'un sol informe no va veure.
    $global:emitCalls.Clear()
    _VistaCataleg $sdTF (Join-Path $EstructuralsDir 'REQ1.json') 'REQ1'
    & $comprova 'Vista REQ1' $global:emitCalls
    $global:emitCalls.Clear()
    _VistaLlicencia $sdTF (Join-Path $EstructuralsDir 'LLIC.json')
    & $comprova 'Vista LLIC' $global:emitCalls

    $env:TEMP = $tmpTF
}

# ---------------------------------------------------------------------------
# Llicencia: PROJECTE va el primer i amb LLETRES; la doc del projecte, nomes
# als favorables
# ---------------------------------------------------------------------------
# Les lletres NO son estetica: quan els requeriments de projecte queden
# resolts, el bloc desapareix i la resta de la documentacio ha de conservar la
# MATEIXA numeracio. Amb tot numerat, l'"1." passaria a ser una altra cosa i el
# titular no ho podria comparar amb el que ja tenia.
function _GdTriaPrimersOrd($parsed) {
    $c = New-Object System.Collections.ArrayList
    foreach ($sec in @($parsed.Sections)) {
        foreach ($el in @($sec.Items)) {
            if ([string]$el.Kind -eq 'subsection' -or [string]$el.Kind -eq 'intro') { continue }
            [void]$c.Add((_ItemKey ([string]$sec.Title) ([string]$el.Short)))
            break
        }
    }
    return $c.ToArray()
}
Write-Host "`n--- Llicencia: PROJECTE amb lletres i la doc nomes als favorables ---"
AssertEq (_LlicLletra 1) 'A'   '_LlicLletra: 1 -> A'
AssertEq (_LlicLletra 4) 'D'   '_LlicLletra: 4 -> D'
AssertEq (_LlicLletra 26) 'Z'  '_LlicLletra: 26 -> Z'
AssertEq (_LlicLletra 27) 'AA' '_LlicLletra: 27 -> AA (com les columnes de l''Excel, sense topall)'
AssertEq (_LlicLletra 28) 'AB' '_LlicLletra: 28 -> AB'
AssertEq (_LlicLletra 52) 'AZ' '_LlicLletra: 52 -> AZ'
AssertEq (_LlicLletra 53) 'BA' '_LlicLletra: 53 -> BA'
AssertEq (_LlicLletra 0) ''    '_LlicLletra: 0 -> res'
AssertEq (_LlicMarca 3 'numero') '3.' '_LlicMarca: numero'
AssertEq (_LlicMarca 3 'lletra') 'C.' '_LlicMarca: lletra'
AssertEq (_LlicMarca 3 '') '3.'       '_LlicMarca: per defecte, numero'
Assert (_LlicPortaDocProjecte 'favorable-pre')  'Doc projecte: al favorable pre, si'
Assert (_LlicPortaDocProjecte 'favorable-post') 'Doc projecte: al favorable post, tambe'
Assert (-not (_LlicPortaDocProjecte 'requeriment')) 'Doc projecte: al requeriment, NO'
Assert (-not (_LlicPortaDocProjecte 'mns')) 'Doc projecte: a la MNS tampoc'
# I el bloc PROJECTE es EL CONTRARI: nomes al requeriment. Son complementaris.
AssertEq (_LlicTitolProjecte) 'REQUERIMENTS PROJECTE' 'El bloc es diu REQUERIMENTS PROJECTE (no nomes PROJECTE)'
Assert (_LlicPortaProjecte 'requeriment') 'Bloc PROJECTE: al requeriment, si'
Assert (-not (_LlicPortaProjecte 'favorable-pre'))  'Bloc PROJECTE: al favorable pre, NO'
Assert (-not (_LlicPortaProjecte 'favorable-post')) 'Bloc PROJECTE: al favorable post, NO'
foreach ($fx in @('requeriment', 'favorable-pre', 'favorable-post')) {
    Assert ((_LlicPortaProjecte $fx) -ne (_LlicPortaDocProjecte $fx)) ('Complementaris a ' + $fx + ': o el bloc o la documentacio, mai tots dos')
}

# I sobre el document sencer: l'ordre i els dos comptadors.
$llicOrd = $null
try { $llicOrd = Read-LlicCataleg (Join-Path $EstructuralsDir 'LLIC.json') } catch { }
$req1Ord = $null
try { $req1Ord = Read-CatalegJson (Join-Path $EstructuralsDir 'REQ1.json') } catch { }
if ($null -ne $llicOrd -and $null -ne $req1Ord) {
    . (Join-Path $TestsDir 'FormatDoubles.ps1')
    $sdO = [pscustomobject]@{ Range = [pscustomobject]@{ Start = 0; End = 0 } }
    $sdO | Add-Member ScriptMethod EndKey { param($u) } -Force
    $sdO | Add-Member ScriptMethod InsertBreak { param($b) } -Force
    $ddO = [pscustomobject]@{}
    $ddO | Add-Member ScriptMethod Activate {} -Force
    $ddO | Add-Member ScriptMethod Save {} -Force
    $ddO | Add-Member ScriptMethod Close { param($x) } -Force
    $wdO = [pscustomobject]@{ Selection = $sdO }
    $script:_docO = $ddO
    function _ResolveOutputDir { return ([System.IO.Path]::GetTempPath()) }
    function _GetUniqueOutputPath($d, $b) { return (Join-Path $d $b) }
    function _OpenOutputDocument($w, $p) { return $script:_docO }
    function Select-CapcaleraBlock($d, $w) { }
    function Apply-HeaderReplacements { param($doc, $header) }
    $tmpO = $env:TEMP
    if ([string]::IsNullOrWhiteSpace($env:TEMP)) { $env:TEMP = [System.IO.Path]::GetTempPath() }

    $idxO = _LlicIndexReq1 $req1Ord
    $abansO = @(@((_LlicPuntsPerBloc $llicOrd $idxO 'ABANS' $req1Ord).Punts) | Select-Object -First 2 |
                ForEach-Object { $_ | Select-Object * | Add-Member NoteProperty Estat 'no' -PassThru -Force })
    $projO = @(_LlicPuntsDeSeleccio (Build-SelectionFromKeys $req1Ord.Sections (@(_GdTriaPrimersOrd $req1Ord) | Select-Object -First 3)))
    $modelO = @{
        Fase = 'requeriment'; EsProvisional = $false
        Header = @{ ID_GIA = '1'; TITULAR = 'X'; CLASSIFICACIO = 'Y' }; Fields = [ordered]@{}
        Abans = $abansO; Projecte = $projO; Despres = @()
        Doc = @{ Text = 'Signada pel tecnic:'; Items = @('Projecte') }; Condicions = ''; Cataleg = $llicOrd
    }
    $global:emitCalls.Clear()
    [void](Build-LlicenciaDocument $wdO $modelO)
    $emO = @($global:emitCalls)
    $blocsO = @($emO | Where-Object { $_ -like 'BLOC|*' })
    AssertEq $blocsO[0] ('BLOC|' + (_LlicTitolProjecte)) 'Ordre: REQUERIMENTS PROJECTE va el PRIMER de tot al requeriment'
    Assert (-not ($emO | Where-Object { $_ -like 'BLOC|DOCUMENTACI* PROJECTE' })) 'Ordre: al requeriment no hi ha la doc del projecte'
    # PROJECTE amb lletres, ABANS amb numeros que comencen per 1.
    $marquesO = @($emO | Where-Object { $_ -like 'ITEM|*' } | ForEach-Object { ($_ -split '\|')[1] })
    Assert ([bool]($marquesO[0] -eq 'A.')) 'Marques: el bloc PROJECTE comenca per A.'
    $numsO = @($marquesO | Where-Object { $_ -match '^\d+\.$' })
    Assert ([bool]($numsO[0] -eq '1.')) 'Marques: la numeracio de la documentacio comenca per 1. (les lletres no hi compten)'
    $lletresO = @($marquesO | Where-Object { $_ -match '^[A-Z]+\.$' })
    AssertEq ($lletresO -join ',') 'A.,B.,C.' 'Marques: A., B., C. seguides'

    # I al FAVORABLE, la doc del projecte hi es i va DALT DE TOT.
    $modelO.Fase = 'favorable-pre'
    $global:emitCalls.Clear()
    [void](Build-LlicenciaDocument $wdO $modelO)
    $emP = @($global:emitCalls)
    Assert ([bool]($emP[0] -like 'BLOC|DOCUMENTACI* PROJECTE')) 'Favorable: la doc del projecte va dalt de tot'
    # I el bloc PROJECTE NO hi surt: aquells requeriments ja estan resolts.
    Assert (-not ($emP | Where-Object { $_ -eq ('BLOC|' + (_LlicTitolProjecte)) })) 'Favorable: el bloc REQUERIMENTS PROJECTE no hi surt'
    Assert (-not ($emP | Where-Object { $_ -match '^ITEM\|[A-Z]+\.\|' })) 'Favorable: ...i per tant cap punt amb lletra'
    $env:TEMP = $tmpO
}

# ---------------------------------------------------------------------------
# L'ORDRE DELS INFORMES AL MENU, i que MNS/Traspas hi te entrada propia
# ---------------------------------------------------------------------------
# L'ordre el decideix l'usuari i es una llista, no una casualitat del codi.
# Select-Mode es WinForms i les proves no el criden mai (vegeu CLAUDE.md), o
# sigui que aixo es una prova de FONT: mira en quin ordre s'afegeixen les
# entrades a $menu.
