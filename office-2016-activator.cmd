@echo off
title Activate Microsoft Office 2016 ALL versions for FREE!&cls&echo =====================================&echo #Copyright: MSGuides.com&echo =====================================&echo.&echo #Supported products:&echo - Microsoft Office Standard 2016&echo - Microsoft Office Professional Plus 2016&echo.&echo.
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"
for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
echo.&echo ====================================&echo Activating your Office...&set i=1&cscript //nologo ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 >nul
:server
if %i%==1 set KMS_Sev=kms.MSGuides.com
if %i%==2 set KMS_Sev=kms2.MSGuides.com
if %i%==3 set KMS_Sev=kms3.MSGuides.com
if %i%==4 goto unsupported
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul
echo ------------------------------------&echo.&echo.
cscript //nologo ospp.vbs /act | find /i "successful" && (echo.& echo ====================================== & echo. & choice /n /c YN /m "Pressione N pra sair [YN]!" & if errorlevel 2 exit) || (echo The connection to the server failed! Trying to connect to another one... & echo Please wait... & echo. & echo. & set /a i+=1 & goto server)
explorer "http://MSGuides.com"&goto halt
:unsupported
echo.&echo ======================================&echo Sorry! Your version is not supported.&echo Please try installing the latest version here: bit.ly/getmsps
:halt
pause