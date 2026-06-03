@echo off
setlocal enabledelayedexpansion
set PATH=%PATH%;%cd%\openssl\bin

echo [1] Opening GUI to paste CRT content...
cscript //nologo openssl\CreateCRT.vbs

if %errorlevel% neq 0 (
    echo Error: No certificate data provided.
    pause
    exit /b
)

set /p PFX_PWD=<pwd.txt

echo [2] Generating PFX file...

openssl pkcs12 -export -out certificate.pfx -inkey private.key -in certificate.crt -passout pass:%PFX_PWD%

if %errorlevel% equ 0 (
    echo.
    echo ---------------------------------------------------
    echo SUCCESS! 
    echo File created: certificate.pfx
    echo ---------------------------------------------------
) else (
    echo.
    echo ERROR: Failed to generate PFX. 
    echo Make sure 'private.key' exists in this folder and matches the CRT.
)


pause