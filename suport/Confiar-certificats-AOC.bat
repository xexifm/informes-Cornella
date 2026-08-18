@echo off
REM Confiar en els certificats de l'AOC en AQUEST ordinador, perque les
REM signatures fetes amb el certificat TCAT (Consorci AOC) surtin com a
REM VALIDES a l'Adobe i a la resta de programes.
REM
REM PER QUE CAL: l'Adobe, per defecte, nomes es refia de les seves llistes
REM (AATL i la llista europea EUTL). Si l'entitat que ha emes el certificat
REM del signant no hi es (o la llista no s'ha descarregat), la signatura surt
REM com a "validesa DESCONEGUDA" encara que sigui perfectament correcta. La
REM solucio oficial (la mateixa que documenten el BOE i els ministeris per als
REM seus documents) es donar confianca a l'entitat emissora en aquest PC.
REM
REM QUE FA: instal.la els DOS certificats PUBLICS del Consorci AOC (l'arrel
REM "CA CONSORCI AOC (G3) ROOT-A" i la intermedia "SubCA SECTOR PUBLIC Q (G3)
REM A.1") al magatzem de certificats de Windows DE L'USUARI (no cal ser
REM administrador). Son certificats publics: no hi ha cap secret aqui dins.
REM
REM Doble clic per executar. Windows demanara confirmacio per a l'ARREL:
REM respon SI.

title Confiar en els certificats de l'AOC
setlocal
set "TMPDIR=%TEMP%\aoc-certs"
mkdir "%TMPDIR%" 2>nul

> "%TMPDIR%\aoc-root.cer" (
  echo -----BEGIN CERTIFICATE-----
  echo MIICpTCCAiqgAwIBAgIQBnQ8jTRW14AiMtM9T+k5OjAKBggqhkjOPQQDAzCBgjEL
  echo MAkGA1UEBhMCRVMxMzAxBgNVBAoMKkNPTlNPUkNJIEFETUlOSVNUUkFDSU8gT0JF
  echo UlRBIERFIENBVEFMVU5ZQTEYMBYGA1UEYQwPVkFURVMtUTA4MDExNzVBMSQwIgYD
  echo VQQDDBtDQSBDT05TT1JDSSBBT0MgKEczKSBST09ULUEwHhcNMjMwMTI2MTAyMTE0
  echo WhcNNDgwMTIwMTAyMTE0WjCBgjELMAkGA1UEBhMCRVMxMzAxBgNVBAoMKkNPTlNP
  echo UkNJIEFETUlOSVNUUkFDSU8gT0JFUlRBIERFIENBVEFMVU5ZQTEYMBYGA1UEYQwP
  echo VkFURVMtUTA4MDExNzVBMSQwIgYDVQQDDBtDQSBDT05TT1JDSSBBT0MgKEczKSBS
  echo T09ULUEwdjAQBgcqhkjOPQIBBgUrgQQAIgNiAATqiBDqgC81r92gxsFn70ROHbtK
  echo XMc1+EbmXZe2jbUmn8p1876dNFRG0IJ/zeDlgliIQKbeHlfw/vpL1JT2HFZlUVKX
  echo 0N4Ne8V+459WC0NotnZudCGswqFuudkp6fAh9FijYzBhMA8GA1UdEwEB/wQFMAMB
  echo Af8wHwYDVR0jBBgwFoAUSSOWCR8lLL4StugPA/NrmyflnoEwHQYDVR0OBBYEFEkj
  echo lgkfJSy+ErboDwPza5sn5Z6BMA4GA1UdDwEB/wQEAwIBBjAKBggqhkjOPQQDAwNp
  echo ADBmAjEAhzcSYQwGjJQHWaA/JVox+uWq5e/JEdWdJL4Gl+WNLGd96NW7FV1b/FD/
  echo hRb7tY6wAjEAnx6C1ClLMFZ63WmN7PGbG4PUGjVf9qoToRnBKB6qMK0jl+q+3LLd
  echo MxmaI0FR92TO
  echo -----END CERTIFICATE-----
)

> "%TMPDIR%\aoc-subca.cer" (
  echo -----BEGIN CERTIFICATE-----
  echo MIIDtDCCAzugAwIBAgIQWcbauO6RG0+NoKtdwcr+ujAKBggqhkjOPQQDAzCBgjEL
  echo MAkGA1UEBhMCRVMxMzAxBgNVBAoMKkNPTlNPUkNJIEFETUlOSVNUUkFDSU8gT0JF
  echo UlRBIERFIENBVEFMVU5ZQTEYMBYGA1UEYQwPVkFURVMtUTA4MDExNzVBMSQwIgYD
  echo VQQDDBtDQSBDT05TT1JDSSBBT0MgKEczKSBST09ULUEwHhcNMjMwMzAyMTEzNjUx
  echo WhcNNDgwMTIwMTAyMTE0WjCBhTELMAkGA1UEBhMCRVMxMzAxBgNVBAoMKkNPTlNP
  echo UkNJIEFETUlOSVNUUkFDSU8gT0JFUlRBIERFIENBVEFMVU5ZQTEYMBYGA1UEYQwP
  echo VkFURVMtUTA4MDExNzVBMScwJQYDVQQDDB5TdWJDQSBTRUNUT1IgUFVCTElDIFEg
  echo KEczKSBBLjEwdjAQBgcqhkjOPQIBBgUrgQQAIgNiAASnYFoQg73R9amVtEbaFnm2
  echo EyOpYtGnuQhdSgWfOLzL7boQbmt6Cbfl0c4KdxzggsTDF13wEYtYz6NQqYLOWYil
  echo LAxMVfKUJDPCz5rCUJfEtGkwsDev6ybHE1uuq5l0FMujggFvMIIBazASBgNVHRMB
  echo Af8ECDAGAQH/AgEAMB8GA1UdIwQYMBaAFEkjlgkfJSy+ErboDwPza5sn5Z6BMGkG
  echo CCsGAQUFBwEBBF0wWzA4BggrBgEFBQcwAoYsaHR0cDovL2Vwc2NkLmFvYy5jYXQv
  echo ZGVzY2FycmVnYS9jYXJvb3QtYS5jcnQwHwYIKwYBBQUHMAGGE2h0dHA6Ly9vY3Nw
  echo LmFvYy5jYXQwNgYDVR0gBC8wLTArBgRVHSAAMCMwIQYIKwYBBQUHAgEWFWh0dHBz
  echo Oi8vZXBzY2QuYW9jLmNhdDAqBgNVHSUEIzAhBggrBgEFBQcDAgYKKwYBBAGCNxQC
  echo AgYJKoZIhvcvAQEFMDYGA1UdHwQvMC0wK6ApoCeGJWh0dHA6Ly9lcHNjZC5hb2Mu
  echo Y2F0L2NybC9jYXJvb3QtYS5jcmwwHQYDVR0OBBYEFB3+KWURvz9XZyK7l9R4YTSb
  echo dlPmMA4GA1UdDwEB/wQEAwIBBjAKBggqhkjOPQQDAwNnADBkAjB9vKmYrgB39aWI
  echo 135Fsj+AMTWsK5u3GV62Dn8q8IFAPYZc3YjngjaLzHTdoirgOEICMCa2Zj5K5pNC
  echo LQxf3xCgnMRYMKF77uN9FxLkEzxNFKi7LR+aBBiWcSI8Ao1n772SYA==
  echo -----END CERTIFICATE-----
)

echo.
echo Instal.lant el certificat ARREL de l'AOC (Windows demanara confirmacio:
echo respon SI al quadre que sortira)...
certutil -user -addstore Root "%TMPDIR%\aoc-root.cer"
if errorlevel 1 (
    echo.
    echo  No s'ha pogut instal.lar l'arrel ^(potser has dit que NO al quadre^).
    echo    Torna a executar aquest fitxer i respon SI.
)
echo.
echo Instal.lant el certificat INTERMEDI...
certutil -user -addstore CA "%TMPDIR%\aoc-subca.cer"

echo.
echo ============================================================
echo  I ara, a l'ADOBE (nomes cal fer-ho UN cop en aquest PC):
echo.
echo   1. Edicion ^> Preferencias ^> Firmas ^> Verificacion ^> boto "Mas..."
echo   2. A "Integracion con Windows", marca les dues caselles de
echo      confiar en els certificats del magatzem de Windows
echo      ("Validando firmas" i "Validando documentos certificados").
echo   3. Accepta-ho tot i TORNA A OBRIR el PDF signat.
echo.
echo  Alternativa sense tocar preferencies: obre el PDF signat, panell
echo  de firmes ^> boto dret a la firma ^> Propiedades de la firma ^>
echo  Mostrar certificado del firmante ^> pestanya Confianza ^>
echo  "Agregar a certificados de confianza" sobre l'ARREL de la ruta.
echo ============================================================
echo.
pause
