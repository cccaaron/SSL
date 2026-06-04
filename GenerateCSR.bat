@echo off
setlocal
set PATH=%PATH%;%cd%\openssl\bin

cscript //nologo openssl\oldCSR.vbs
if exist old_CSR.txt call:getvals args

echo.%args%

echo [1] Opening GUI for CSR details...
cscript //nologo openssl\GUI.vbs %args%
set SELECTED_BITS=%errorlevel%
if not exist "temp_openssl.conf" (
    echo Error: Configuration file was not created.
    timeout 5
    exit
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

start/max CSR.txt

timeout 5
exit



:getvals
setlocal
openssl req -in old_CSR.txt -noout -text >deCSR.txt
for /f "tokens=1,2 delims=:" %%s in ('findstr Subject: deCSR.txt') do set "%%s=%%t"
for /f "tokens=2 delims=(" %%s in ('findstr Public-Key: deCSR.txt') do set/a k=%%s 2>nul
for /f "tokens=*" %%s in ('findstr DNS: deCSR.txt') do set "s=%%s"
del/q deCSR.txt old_CSR.txt
set "s=%s:DNS:=%"
call:cut "%Subject:, =" "%"
set "sub="%C%" "%ST%" "%L%" "%O%" "%OU%" "%CN%""
call:cutSAN "%s:, =" "%"
endlocal&set "%~1=%sub% "%k%" %san:" "=, %"
exit/b

:cut
if "%~1"=="" exit/b
set %1
shift 
goto cut

:cutSAN
if "%~1"=="" exit/b
if "%~1" neq "%CN%" set "san=%san% %1"
shift 
goto cutSAN


