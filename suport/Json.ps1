#requires -Version 5.1
<#
.SYNOPSIS
  Llegir i escriure JSON. Una sola manera, per a tot el programa.

.DESCRIPTION
  Abans, cada modul s'ho feia pel seu compte i aixo tenia tres consequencies
  que es veien al disc:

  1. DOS ENCODINGS INCOMPATIBLES. 17 llocs escrivien amb
     '| Set-Content -Encoding UTF8', que al Windows PowerShell 5.1 posa BOM, i 7
     amb '[IO.File]::WriteAllText(..., UTF8Encoding($false))', que no en posa.
     Dins de la MATEIXA carpeta hi havia les dues coses: ESTRUCTURALS\0
     CAPCALERA.json amb BOM i ESTRUCTURALS\REQ1.json sense. Tres moduls fins i
     tot documentaven "UTF-8 sense BOM" mentre els altres feien el contrari
     sense dir-ho.

  2. CAP ESCRIPTURA ERA ATOMICA. Un desat interromput a mitges (l'usuari tanca
     la sessio, es queda sense bateria, el Drive sincronitza) deixava el fitxer
     TRUNCAT. I com que TOTS els lectors tracten un JSON corrupte igual que un
     fitxer que no hi es -tornen el valor per defecte dins d'un catch buit-, la
     base de llicencies o la de recordatoris es podia perdre SENSE CAP AVIS.
     Aqui s'escriu a un temporal i es MOU a sobre, que es el mateix patro que ja
     feia el .docx (MotorInforme.ps1) i el PDF signat (PdfSignar.ps1).

  3. EL MATEIX ESQUELET COPIAT QUATRE COPS. Load-AppSettings, Load-LastReport,
     Load-ActExtrRegistry i Load-LlicenciaDb feien exactament la mateixa
     seqüencia (Test-Path -> Get-Content -Raw -> IsNullOrWhiteSpace ->
     ConvertFrom-Json -> catch) i nomes es diferenciaven en QUE tornen quan
     falla. Per aixo Read-JsonFile torna $null i prou: el valor per defecte el
     posa el crider, que es l'unic que sap quin ha de ser.

.NOTES
  Es carrega dels PRIMERS a Motor.ps1: Settings.ps1 el fa servir mentre es
  carrega (Load-AppSettings es crida durant el bloc de configuracio), i
  rutes\Ruta.ps1, que corre en un proces a part, tambe l'ha de carregar.
#>

# Llegeix un JSON i el torna com a objecte.
#
# Torna $null en els quatre casos en que no hi ha res utilitzable: el fitxer no
# hi es, es buit, nomes te espais, o no es un JSON valid. NO decideix cap valor
# per defecte: aixo es del crider (uns volen un objecte buit, altres un
# [pscustomobject] amb Version i una llista, i altres $null de debo).
function Read-JsonFile([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
        return ($raw | ConvertFrom-Json)
    } catch {
        return $null
    }
}

# Escriu un objecte com a JSON: UTF-8 SENSE BOM i de manera ATOMICA.
#
# -Depth es el mateix del ConvertTo-Json i cada crider hi posa el seu (les bases
# de dades en volen molta; un settings.json, poca). No te valor per defecte a
# posta: el de ConvertTo-Json es 2, que trunca en silenci qualsevol cosa una
# mica imbricada, i val mes haver-hi pensat.
#
# Crea la carpeta de desti si cal, que abans repetia cada _Save* pel seu compte.
function Write-JsonFile([string]$Path, $Object, [int]$Depth) {
    Write-JsonText $Path ($Object | ConvertTo-Json -Depth $Depth)
}

# El mateix, pero amb el JSON JA SERIALITZAT. Hi ha dos llocs que el tenen fet
# molt abans d'escriure'l -l'editor de catalegs, que primer el valida tornant-lo
# a llegir, i l'exportacio d'activitats, que el mateix text va tambe a Drive- i
# passar-lo per ConvertTo-Json una segona vegada l'escaparia sencer dins d'una
# cadena. Aquesta es la funcio que fa la feina; Write-JsonFile nomes serialitza
# i delega, de manera que l'encoding i l'atomicitat son els mateixos per a tots.
function Write-JsonText([string]$Path, [string]$Json) {
    $dir = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    # ATOMIC: primer el temporal sencer, despres el moviment. Si el proces mor a
    # mig escriure, el que es fa malbe es el .tmp i el fitxer bo continua intacte.
    # El temporal va a LA MATEIXA CARPETA perque Move-Item entre volums no es
    # atomic (i %TEMP% sol ser en un altre volum que una unitat de xarxa).
    # Amb SALT DE LINIA final. Set-Content n'hi posava un i WriteAllText no, o
    # sigui que en migrar-ho tot aqui els fitxers el perdien i el git els marcava
    # amb "\ No newline at end of file" a cada diff. Un fitxer de text acaba amb
    # un salt de linia.
    $tmp = $Path + '.tmp'
    [System.IO.File]::WriteAllText($tmp, ($Json + [Environment]::NewLine), (New-Object System.Text.UTF8Encoding($false)))
    Move-Item -LiteralPath $tmp -Destination $Path -Force
}
