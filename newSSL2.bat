@echo off
setlocal enabledelayedexpansion
echo.&set/p d=Enter new dirctory work: 
md %d%\openssl\bin
md %d%\openssl\etc\ssl
start %d%
set url=https://cccaaron.github.io/SSL
for /f %%f in ('curl -Lk %url%/files.txt') do (
set l=%%f&start/b curl -sLk %url%/!l:\=/! -o !d!\%%f)





