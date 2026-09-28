@echo off
set PATH=%PATH%;%cd%\openssl\bin

echo [1] Opening GUI to paste CRT content...
cscript //nologo openssl\CreateCRT.vbs

if %errorlevel% neq 0 (
    echo Error: No certificate data provided.
    timeout 5
    exit /b
)

echo [2] Generating PFX file... 

set/p x=<openssl\CN
set "%x: =%"
if "%cn:~,1%"=="*" set cn=star%cn:~1%
set "fn=%cn:.=_%-%date:~4,2%-%date:~-2%"

for /f %%p in (pwd.txt) do (
openssl pkcs12 -export -out certificate.pfx -inkey private.key -in certificate.crt -passout pass:%%p -name "%fn%"||goto ERR
openssl pkcs12 -export -out certificate_PBE-SHA1-3DES.pfx -inkey private.key -in certificate.crt -certpbe PBE-SHA1-3DES -keypbe PBE-SHA1-3DES -macalg sha1 -passout pass:%%p -name "%fn%"||goto ERR
)

echo.
echo ---------------------------------------------------
echo SUCCESS! 
echo Files created: certificate.pfx, certificate_PBE-SHA1-3DES.pfx
echo ---------------------------------------------------

copy openssl\mail.msg .
chcp 65001
powershell -ExecutionPolicy Bypass -File "openssl\ems.ps1" "%cd%\mail.msg" "%cd%" "תעודת SSL [%cn:star=*%]"
start mail.msg

timeout 5
rd/q /s openssl
del/q *.bat
exit

:ERR
echo.
echo ERROR: Failed to generate PFX. 
echo Make sure 'private.key' exists in this folder and matches the CRT.
pause&exit
