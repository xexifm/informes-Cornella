#requires -Version 5.1
<#
.SYNOPSIS
  Crea l'acces directe del programa, per poder-lo ANCORAR A LA BARRA DE TASQUES.

.DESCRIPTION
  Windows NO deixa ancorar un .bat (ni un .vbs) a la barra de tasques: nomes hi
  admet accessos directes que apuntin a un EXECUTABLE. Per aixo aqui no es fa un
  acces directe al GenerarInforme.bat, sino un que apunta a:

      wscript.exe  "<clone>\suport\GenerarInforme.vbs"

  ...que es exactament el que ja fa el .bat (el .vbs arrenca el PowerShell sense
  cap finestra de consola). Com que el desti es un .exe, l'acces directe SI que
  es pot ancorar.

  Es deixa a DOS llocs:
    - l'ESCRIPTORI, per tenir-lo a ma;
    - el MENU INICI (%APPDATA%\...\Start Menu\Programs), que es el que fa que
      surti buscant "Generador d'informes" i que es pugui ancorar des d'alli.

  I li posa l'ESCUT (suport\cornella.ico), perque a la barra de tasques es
  reconegui d'un cop d'ull i no surti la icona generica del Windows.

  PER QUE NO S'ANCORA SOL: des del Windows 10, el verb "Ancorar a la barra de
  tasques" ja no es pot invocar per codi (Microsoft el va treure de la interficie
  del shell). El que es troba per internet son trucs que escriuen al registre i
  reinicien l'explorer: son fragils i es poden carregar la barra de tasques de
  l'usuari. Val mes deixar l'acces directe fet i que l'usuari faci un clic dret.

.NOTES
  Nomes te sentit a Windows (WScript.Shell). Les funcions de RUTA son pures i es
  proven a Linux; la creacio de l'acces directe, no.

  CONVENCIO ASCII: el codi no porta accents.
#>

# El nom del fitxer de l'acces directe. Sense accents ni apostrof tipografic:
# es un nom de fitxer i ha de sobreviure a qualsevol codepage.
$Script:AccesDirecteNom = "Generador d'informes Cornella.lnk"

# L'IDENTIFICADOR D'APLICACIO (AppUserModelID), i nomes es escrit AQUI.
#
# Es el que lliga la icona ancorada amb la finestra del programa. Windows agrupa
# la barra de tasques per aquest identificador: si la drecera no en te, l'ancora
# i la finestra son DUES coses diferents -surt un segon boto quan obres el
# programa, i l'ancorada es queda sense icona-. Va a la drecera
# (Set-AccesDirecteAppId) i al proces (UiComuns.ps1), i han de ser EL MATEIX.
$Script:AppUserModelId = 'Cornella.Informes.Generador'

# ON APUNTA l'acces directe, a partir de l'arrel del clone. Funcio PURA: no toca
# el disc, o sigui que es pot provar a qualsevol plataforma.
# $carpetaIcona: on ha de viure la COPIA LOCAL de l'escut (opcional). Vegeu
# _AccesDirecteCarpetaIcona: el clone de l'usuari viu en una unitat de XARXA i
# l'explorador de Windows no es de fiar carregant icones d'alli per a un element
# ancorat -es queda amb la generica-. Amb una copia al disc de sempre, no falla.
function Get-AccesDirecteObjectiu([string]$repoRoot, [string]$systemRoot = '', [string]$carpetaIcona = '') {
    if ([string]::IsNullOrWhiteSpace($systemRoot)) { $systemRoot = 'C:\Windows' }
    $sep = '\'
    $rr = ([string]$repoRoot).TrimEnd('\', '/')
    $icoOrigen = ($rr + $sep + 'suport' + $sep + 'cornella.ico')
    $ico = $icoOrigen
    if (-not [string]::IsNullOrWhiteSpace($carpetaIcona)) {
        $ico = (([string]$carpetaIcona).TrimEnd('\', '/') + $sep + 'cornella.ico')
    }
    return @{
        Desti       = ($systemRoot.TrimEnd('\') + $sep + 'System32' + $sep + 'wscript.exe')
        Arguments   = ('"' + $rr + $sep + 'suport' + $sep + 'GenerarInforme.vbs"')
        Carpeta     = $rr
        Icona       = $ico
        IconaOrigen = $icoOrigen
        Nom         = $Script:AccesDirecteNom
    }
}

# On va la copia local de l'escut. El mateix lloc que la resta d'estat d'aquest
# ordinador (%LOCALAPPDATA%\InformesCornella), que sobreviu a tornar a clonar.
function _AccesDirecteCarpetaIcona {
    if ($AppDataDir) { return [string]$AppDataDir }
    if ([string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) { return '' }
    return [string](Join-Path $env:LOCALAPPDATA 'InformesCornella')
}

# Els llocs on es deixa. Funcio PURA (rep les carpetes ja resoltes).
function Get-AccesDirecteDestins([string]$escriptori, [string]$menuInici) {
    $out = New-Object System.Collections.ArrayList
    foreach ($d in @($escriptori, $menuInici)) {
        if ([string]::IsNullOrWhiteSpace($d)) { continue }
        [void]$out.Add(($d.TrimEnd('\') + '\' + $Script:AccesDirecteNom))
    }
    return $out.ToArray()
}

# ----------------------------------------------------------------------------
# L'AppUserModelID DE LA DRECERA
# ----------------------------------------------------------------------------
# El WScript.Shell sap fer una drecera pero NO sap posar-li aquesta propietat:
# cal anar a l'IShellLink i demanar-li l'IPropertyStore. Es el mateix patro que
# _PickFolderModern (UiComuns.ps1): les interficies COM es declaren amb Add-Type
# i es compilen EN VIU el primer cop.
#
# PER QUE CAL. Sense aixo, la icona ancorada i la finestra del programa son dues
# aplicacions diferents per a Windows:
#   - l'ancorada surt SENSE ICONA (el desti es wscript.exe, i Windows no sap
#     que aquella drecera es "una aplicacio");
#   - i en obrir-la surt un SEGON boto a la barra de tasques, en lloc
#     d'il-luminar-se el que ja hi havia.
# Amb el mateix identificador a la drecera i al proces, Windows els ajunta.
#
# ATENCIO: el C# ha de ser de PowerShell 5.1 (C# 5): res de 'nameof', ni
# membres amb cos d'expressio, ni 'out var'.
$Script:AccesDirecteTipusCarregats = $false
function _AccesDirecteCarregaTipus {
    if ($Script:AccesDirecteTipusCarregats) { return $true }
    if ('CornellaApp.Lnk' -as [type]) { $Script:AccesDirecteTipusCarregats = $true; return $true }
    $codi = @'
using System;
using System.Runtime.InteropServices;

namespace CornellaApp {

    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct PropertyKey {
        public Guid fmtid;
        public int pid;
    }

    // El PROPVARIANT nomes s'omple aqui per a una cadena (VT_LPWSTR = 31). Les
    // dades comencen al byte 8: 2 del tipus i 6 de reservats, tant a 32 com a
    // 64 bits.
    [StructLayout(LayoutKind.Explicit)]
    public struct PropVariant {
        [FieldOffset(0)] public ushort vt;
        [FieldOffset(8)] public IntPtr p;
    }

    [ComImport, Guid("00021401-0000-0000-C000-000000000046")]
    public class ShellLink { }

    [ComImport, Guid("0000010b-0000-0000-C000-000000000046"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPersistFile {
        void GetClassID(out Guid pClassID);
        [PreserveSig] int IsDirty();
        void Load([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, uint dwMode);
        void Save([MarshalAs(UnmanagedType.LPWStr)] string pszFileName,
                  [MarshalAs(UnmanagedType.Bool)] bool fRemember);
        void SaveCompleted([MarshalAs(UnmanagedType.LPWStr)] string pszFileName);
        void GetCurFile([MarshalAs(UnmanagedType.LPWStr)] out string ppszFileName);
    }

    [ComImport, Guid("886d8eeb-8cf2-4446-8d02-cdba1dbdcf99"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPropertyStore {
        void GetCount(out uint cProps);
        void GetAt(uint iProp, out PropertyKey pkey);
        void GetValue(ref PropertyKey key, out PropVariant pv);
        void SetValue(ref PropertyKey key, ref PropVariant pv);
        void Commit();
    }

    public static class Lnk {
        // PKEY_AppUserModel_ID
        static readonly Guid FMTID = new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3");
        const int PID = 5;
        const ushort VT_LPWSTR = 31;

        [DllImport("ole32.dll")]
        static extern int PropVariantClear(ref PropVariant pvar);

        public static void SetAppId(string lnkPath, string appId) {
            object o = new ShellLink();
            try {
                IPersistFile pf = (IPersistFile)o;
                pf.Load(lnkPath, 2);            // 2 = STGM_READWRITE
                IPropertyStore ps = (IPropertyStore)o;
                PropertyKey key = new PropertyKey();
                key.fmtid = FMTID;
                key.pid = PID;
                PropVariant pv = new PropVariant();
                pv.vt = VT_LPWSTR;
                pv.p = Marshal.StringToCoTaskMemUni(appId);
                try {
                    ps.SetValue(ref key, ref pv);
                    ps.Commit();
                } finally {
                    PropVariantClear(ref pv);
                }
                pf.Save(lnkPath, true);
            } finally {
                Marshal.ReleaseComObject(o);
            }
        }
    }
}
'@
    try {
        Add-Type -TypeDefinition $codi -ErrorAction Stop
        $Script:AccesDirecteTipusCarregats = $true
        return $true
    } catch {
        return $false
    }
}

# Posa l'AppUserModelID a una drecera ja creada. Retorna $true si ho ha fet.
function Set-AccesDirecteAppId([string]$lnkPath, [string]$appId = '') {
    if ([string]::IsNullOrWhiteSpace($appId)) { $appId = [string]$Script:AppUserModelId }
    if (-not (Test-Path -LiteralPath $lnkPath)) { return $false }
    if (-not (_AccesDirecteCarregaTipus)) { return $false }
    try {
        [CornellaApp.Lnk]::SetAppId((Resolve-Path -LiteralPath $lnkPath).Path, $appId)
        return $true
    } catch {
        return $false
    }
}

# Crea (o refresca) l'acces directe. Retorna @{ Ok; Fets; Errors }.
function New-AccesDirecteInformes([string]$repoRoot = '') {
    if ([string]::IsNullOrWhiteSpace($repoRoot)) { $repoRoot = $RepoRoot }
    $obj = Get-AccesDirecteObjectiu $repoRoot $env:SystemRoot (_AccesDirecteCarpetaIcona)
    # LA COPIA LOCAL DE L'ESCUT. Si no es pot fer, es fa servir el del clone (a
    # una unitat local funciona igual de be).
    try {
        $carpIco = Split-Path -Parent ([string]$obj.Icona)
        if (-not [string]::IsNullOrWhiteSpace($carpIco)) {
            if (-not (Test-Path -LiteralPath $carpIco)) { New-Item -ItemType Directory -Path $carpIco -Force | Out-Null }
            Copy-Item -LiteralPath ([string]$obj.IconaOrigen) -Destination ([string]$obj.Icona) -Force
        }
    } catch { $obj.Icona = [string]$obj.IconaOrigen }
    if (-not (Test-Path -LiteralPath ([string]$obj.Icona))) { $obj.Icona = [string]$obj.IconaOrigen }
    $fets = New-Object System.Collections.ArrayList
    $errs = New-Object System.Collections.ArrayList

    $vbs = $obj.Arguments.Trim('"')
    if (-not (Test-Path -LiteralPath $vbs)) {
        return @{ Ok = $false; Fets = @(); Errors = @(("No trobo " + $vbs)) }
    }

    $sh = $null
    try { $sh = New-Object -ComObject WScript.Shell } catch {
        return @{ Ok = $false; Fets = @(); Errors = @('No s''ha pogut crear l''acces directe (WScript.Shell).') }
    }
    $destins = Get-AccesDirecteDestins ([Environment]::GetFolderPath('Desktop')) ([Environment]::GetFolderPath('Programs'))
    foreach ($ruta in $destins) {
        try {
            $carpeta = Split-Path -Parent $ruta
            if (-not (Test-Path -LiteralPath $carpeta)) { New-Item -ItemType Directory -Path $carpeta -Force | Out-Null }
            $lnk = $sh.CreateShortcut($ruta)
            $lnk.TargetPath = [string]$obj.Desti
            $lnk.Arguments = [string]$obj.Arguments
            $lnk.WorkingDirectory = [string]$obj.Carpeta
            $lnk.Description = "Generador d'informes - Ajuntament de Cornella de Llobregat"
            # AMB L'INDEX (",0"): es la forma que espera el shell per a un fitxer
            # d'icones, i sense ell hi ha Windows que es queden amb la generica.
            if (Test-Path -LiteralPath ([string]$obj.Icona)) { $lnk.IconLocation = ([string]$obj.Icona + ',0') }
            $lnk.Save()
            # ...i l'identificador d'aplicacio, que es el que fa que l'ancorada i
            # la finestra del programa siguin LA MATEIXA cosa per a Windows.
            if (-not (Set-AccesDirecteAppId $ruta ([string]$Script:AppUserModelId))) {
                [void]$errs.Add(($ruta + ": la drecera s'ha creat, pero no s'hi ha pogut posar l'identificador d'aplicacio (sortira sense icona a la barra de tasques)."))
            }
            [void]$fets.Add($ruta)
        } catch {
            [void]$errs.Add(($ruta + ': ' + $_.Exception.Message))
        }
    }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sh) | Out-Null } catch { }
    return @{ Ok = ($fets.Count -gt 0); Fets = $fets.ToArray(); Errors = $errs.ToArray() }
}

# El que crida Crear-acces-directe.bat: crea l'acces directe i ho explica per
# pantalla. Va aqui i no al .bat perque una ordre de PowerShell escrita dins
# d'un .bat, amb cometes i canonades, es un niu d'errors (vegeu CLAUDE.md: dins
# de cometes dobles el cmd NO interpreta el '|', pero si que deixa passar el '^'
# literal, i el que arriba al PowerShell ja no es el que havies escrit).
function Invoke-CrearAccesDirecte([string]$repoRoot = '') {
    if ([string]::IsNullOrWhiteSpace($repoRoot)) { $repoRoot = (Get-Location).Path }
    $r = New-AccesDirecteInformes $repoRoot
    Write-Host ''
    if ([bool]$r.Ok) {
        Write-Host "Acces directe creat a:" -ForegroundColor Green
        foreach ($f in @($r.Fets)) { Write-Host ('   ' + $f) }
    } else {
        Write-Host "No s'ha pogut crear l'acces directe." -ForegroundColor Red
    }
    foreach ($e in @($r.Errors)) { Write-Host ('   ' + $e) -ForegroundColor Yellow }
    if ([bool]$r.Ok) {
        # Windows es queda la drecera ANCORADA tal com era el dia que es va
        # ancorar: si ja hi era, s'ha de treure i tornar-hi a posar perque
        # agafi la icona i l'identificador nous.
        Write-Host ''
        Write-Host 'Si ja el tenies ancorat a la barra de tasques, treu-lo i torna-hi:' -ForegroundColor Cyan
        Write-Host '   Windows es queda la copia del dia que el vas ancorar.'
    }
    return $r
}
