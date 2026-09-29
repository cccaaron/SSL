@echo off
setlocal enabledelayedexpansion
echo.&set/p "d=Enter new dirctory work: "
set "d=%d:"=%"&call:clear
set "d=%username%-%d%"
if not exist %date:~-4% md %date:~-4%
cd %date:~-4%
md %d%\openssl\bin
md %d%\openssl\etc\ssl
start %d%
set url=https://cccaaron.github.io/SSL
for /f %%f in ('curl -Lk %url%/files.txt') do (
set l=%%f&start/b curl -sLk %url%/!l:\=/! -o !d!\%%f)
exit
:clear
for /f "tokens=1,* delims=*" %%i in ("!d!") do (
set d=%%i%%j
if "%%j" neq "" goto clear
)
for %%j in (\ / : ? ^< ^> ^|) do set d=!d:%%j=!
