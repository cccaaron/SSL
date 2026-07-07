@echo off
title Install PFX to IIS (Auto mode)
if "%~1" neq "" set arg="""%~1"""
net session >nul 2>&1||(powershell start-process ""%0"" %arg% -Verb RunAs&exit)

setlocal enabledelayedexpansion
set PATH=%PATH%;%systemroot%\system32\inetsrv\

set "pfxFile=%*"
if not defined pfxFile set/p pfxFile=PFX: 
set "pfxFile=%pfxFile:"=%"
cls
set pfxFile

:trypwd
set/p pfxPwd=Password: 
certutil -p "%pfxPwd%" -dump "%pfxFile%">nul||(
echo Wrong PFX Password "%pfxPwd%" try again.&goto trypwd)

for /f "tokens=1,2 delims=:" %%i in ('certutil -p "%pfxPwd%" -dump "%pfxFile%"^|findstr /c:"Subject" /c:"Cert Hash"') do (set "str=%%j"
if "%%i"=="Subject" (call:getvar "!str:, =" "!") else set "t=!str: =!")
set "domain=%CN: =%"

certutil -f -p "%pfxPwd%" -importpfx "%pfxFile%" NoRoot

set n=0
for /f tokens^=2^ delims^=^" %%i in ('appcmd list site') do set/a n+=1&set "s!n!=%%i"

if %n%==1 (set k=1) else (for /l %%i in (1,1,%n%) do echo %%i - !s%%i!
echo.&set/p k=Select number of WebSite: )
set webSite=!s%k%!

appcmd set site /site.name:"%webSite%" /-bindings.[protocol='https',bindingInformation='*:443:%domain%'] >nul 2>&1
netsh http delete sslcert ipport=0.0.0.0:443 >nul 2>&1

netsh http add sslcert ipport=0.0.0.0:443 certhash=%t% appid={4dc3e181-e14b-4a21-b022-59fc669b0914}
appcmd set site /site.name:"%webSite%" /+bindings.[protocol='https',bindingInformation='*:443:%domain%']

endlocal
timeout 20
exit/b

:getvar
if "%~1"=="" exit/b
set %1&shift &goto getvar
