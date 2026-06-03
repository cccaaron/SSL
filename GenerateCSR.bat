@echo off
setlocal
set PATH=%PATH%;%cd%\openssl\bin

echo [1] Opening GUI for CSR details...
cscript //nologo openssl\GUI.vbs
set SELECTED_BITS=%errorlevel%
if not exist "temp_openssl.conf" (
    echo Error: Configuration file was not created.
    pause
    exit /b
)

echo [2] Generating CSR with %SELECTED_BITS% bits...
openssl req -new -out CSR.txt -newkey rsa:%SELECTED_BITS% -nodes -keyout private.key -config temp_openssl.conf
del/q temp_openssl.conf

echo.
echo ---------------------------------------------------
echo DONE! 
echo Files created: 
echo - private.key (Keep this secret!)
echo - CSR.txt (Send this to your CA)
echo ---------------------------------------------------
echo.

start/max notepad CSR.txt

pause