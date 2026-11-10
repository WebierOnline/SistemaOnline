@echo off
setlocal
rem ============================================================================
rem CompartilharExportarXML.bat
rem Cria (uma unica vez, no SERVIDOR) o compartilhamento de rede "ExportarXML" usado pelo
rem envio mensal de XML para a contabilidade (os terminais PDV copiam o .rar deles para ca).
rem Criar um compartilhamento exige administrador: de duplo clique neste arquivo e confirme a
rem janela do Windows (UAC). Nao precisa abrir o Prompt de Comando.
rem
rem Uso:  CompartilharExportarXML.bat            -> pasta padrao: NFE\ExportarXML (ao lado deste arquivo)
rem       CompartilharExportarXML.bat "C:\pasta" -> outra pasta (a mesma do campo Diretorio XML da empresa + \ExportarXML)
rem ============================================================================

set "PASTA=%~1"

net session >nul 2>&1
if %errorlevel% equ 0 goto :admin

echo Solicitando permissao de administrador...
powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -ArgumentList ('\"' + '%PASTA%' + '\"') -Verb RunAs"
goto :fim

:admin
if "%PASTA%"=="" set "PASTA=%~dp0NFE\ExportarXML"
if not exist "%PASTA%" mkdir "%PASTA%"
if not exist "%PASTA%" goto :semPasta

echo Pasta: %PASTA%
echo.

rem refaz o compartilhamento (garante que aponta para a pasta certa)
net share ExportarXML /delete >nul 2>&1
net share ExportarXML="%PASTA%" /GRANT:Everyone,FULL
if errorlevel 1 goto :falhou

rem permissao NTFS de modificar para todos (a mesma que os terminais precisam para gravar o .rar)
icacls "%PASTA%" /grant *S-1-1-0:(OI)(CI)M >nul 2>&1

rem libera "Compartilhamento de Arquivo e Impressora" no firewall (nome varia conforme o idioma do Windows)
netsh advfirewall firewall set rule group="Compartilhamento de Arquivo e Impressora" new enable=Yes >nul 2>&1
netsh advfirewall firewall set rule group="File and Printer Sharing" new enable=Yes >nul 2>&1

echo.
echo ============================================================
echo  PRONTO! Compartilhamento criado: \%COMPUTERNAME%\ExportarXML
echo ============================================================
goto :fim

:semPasta
echo ERRO: nao foi possivel criar/encontrar a pasta "%PASTA%".
goto :fim

:falhou
echo.
echo ERRO: nao foi possivel criar o compartilhamento. Verifique se a pasta existe e se o servico
echo "Servidor" (LanmanServer) do Windows esta em execucao.

:fim
echo.
pause
endlocal
