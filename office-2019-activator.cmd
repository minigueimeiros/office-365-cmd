@echo off
title Ativador MicrosoftOffice 2019 TODAS as versoes sem custo! Traduzido por LGR HARDWARE&cls&echo ============================================================================&echo #Project: Ativando programas da Microsoft gratuitamente sem software&echo ============================================================================&echo.&echo #Programas Suportados:&echo - Microsoft Office Standard 2019&echo - Microsoft Office Professional Plus 2019&echo.&echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&echo.&echo ============================================================================&echo Ativando o seu Office...&cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:6MWKP >nul&cscript //nologo ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul&set i=1
:server
if %i%==1 set KMS_Sev=kms7.MSGuides.com
if %i%==2 set KMS_Sev=kms8.MSGuides.com
if %i%==3 set KMS_Sev=kms9.MSGuides.com
if %i%==4 goto notsupported
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul&echo ============================================================================&echo.&echo.
cscript //nologo ospp.vbs /act | find /i "successful" && (echo.&echo ============================================================================&echo.&echo #Blog Oficial do Criador: RafaelPantojaSuperTutoriais.com&echo.&echo #Como funciona: bit.ly/kms-server&echo.&echo #Fique a vontade para contactar o criador em msguides.com@gmail.com se voce tiver alguma duvida ou angustia em utilizar esse metodo.&echo.&echo #Por favor considere ajudar essa iniciativa: donate.msguides.com&echo.&echo #Seu suporte faz com que eu consiga manter meus servidores em funcionamento todos os dias!&echo.&echo ============================================================================&choice /n /c SN /m "Voce Gostaria de Visitar o Blog do Criador [S,N]?" & if errorlevel 2 exit) || (echo A conexao com o meu servidor KMS Falhou! Tentando conectar a outro... & echo Por favor aguarde... & echo. & echo. & set /a i+=1 & goto server)
explorer "http://MSGuides.com"&goto halt
:notsupported
echo.&echo ============================================================================&echo Desculpa! Nao ha suporte a sua versao do Office.&echo Porfavor tente instalar a ultima disponivel: bit.ly/aiomsp
:halt
pause >nul
