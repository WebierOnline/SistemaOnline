@echo off
chcp 1252 >nul
setlocal EnableDelayedExpansion

REM ============================================================
REM Executa os scripts .sql de C:\Projeto\Arquivos\scripts na ordem
REM definida pelo manifesto _manifesto.txt (NNN|arquivo.sql|CATEGORIA).
REM
REM Versao sem PowerShell (para maquinas Windows 7 sem PowerShell ou
REM com politica de execucao bloqueada). Equivalente ao
REM Executar_Scripts.ps1, usando cmd.exe + sqlcmd.
REM
REM CATEGORIA pode ser:
REM   GERAL - roda sempre
REM   OS    - so roda se configuracao.config_valor = 1 onde config_nome = 'OS'
REM
REM Uso:
REM   Executar_Scripts.bat
REM   Executar_Scripts.bat "SERVIDOR\INSTANCIA" "cyber_base" "sa" "senha"
REM   Executar_Scripts.bat "SERVIDOR\INSTANCIA" "cyber_base" "sa" "senha" 44
REM       (retoma a partir do script 045, depois de corrigir um erro no 044)
REM   Executar_Scripts.bat "SERVIDOR\INSTANCIA" "cyber_base" "sa" "senha" 0 NOBACKUP
REM       (pula o backup automatico - uso local/teste repetido, NAO usar em cliente)
REM ============================================================

set "SERVER=%~1"
if "%SERVER%"=="" set "SERVER=.\SQLEXPRESS2008"
set "DATABASE=%~2"
if "%DATABASE%"=="" set "DATABASE=cyber_base"
set "DBUSER=%~3"
if "%DBUSER%"=="" set "DBUSER=sa"
set "DBPASS=%~4"
if "%DBPASS%"=="" set "DBPASS=190106web"
set "CONTINUARAPOS=%~5"
if "%CONTINUARAPOS%"=="" set "CONTINUARAPOS=0"
set "PULARBACKUP=%~6"

set "SCRIPTDIR=%~dp0"
set "MANIFESTO=%SCRIPTDIR%_manifesto.txt"

REM Timestamp sem depender de locale (usa WMIC quando existe; senao, cai no %date%/%time%)
set "TS="
for /f "skip=1 tokens=1-2" %%a in ('wmic os get localdatetime /format:table 2^>nul') do (
    if not "%%a"=="" if "!TS!"=="" set "TS=%%a"
)
if "%TS%"=="" (
    set "TS=%date:~-4%%date:~3,2%%date:~0,2%_%time:~0,2%%time:~3,2%%time:~6,2%"
    set "TS=!TS: =0!"
) else (
    set "TS=%TS:~0,8%_%TS:~8,6%"
)
set "LOGFILE=%SCRIPTDIR%execucao_%TS%.log"

where sqlcmd >nul 2>&1
if errorlevel 1 (
    echo ERRO: sqlcmd nao encontrado no PATH. Instale o "sqlcmd Utility" ^(SQL Server Command Line Utilities^) nesta maquina antes de rodar este script.
    exit /b 1
)

if not exist "%MANIFESTO%" (
    echo ERRO: manifesto nao encontrado: %MANIFESTO%
    exit /b 1
)

call :log "=== Iniciando execucao - Servidor: %SERVER% - Banco: %DATABASE% ==="

if /i "%PULARBACKUP%"=="NOBACKUP" (
    call :log "AVISO: backup de seguranca PULADO. Nao usar essa opcao em banco de cliente."
) else (
    call :log "Localizando pasta de dados do banco para gerar backup de seguranca..."
    set "DATAPATH="
    for /f "usebackq delims=" %%p in (`sqlcmd -S "%SERVER%" -d "%DATABASE%" -U "%DBUSER%" -P "%DBPASS%" -h -1 -Q "SET NOCOUNT ON; SELECT TOP 1 physical_name FROM sys.master_files WHERE database_id = DB_ID('%DATABASE%') AND type_desc = 'ROWS'" 2^>nul`) do (
        if not "%%p"=="" set "DATAPATH=%%p"
    )
    if "!DATAPATH!"=="" (
        call :log "ERRO: nao foi possivel localizar a pasta de dados do banco '%DATABASE%'. Abortando ANTES de rodar qualquer script, por seguranca."
        exit /b 1
    )
    for %%f in ("!DATAPATH!") do set "DATAFOLDER=%%~dpf"
    set "BACKUPFILE=!DATAFOLDER!%DATABASE%_pre_scripts_%TS%.bak"
    call :log "Gerando backup de seguranca em: !BACKUPFILE!"
    sqlcmd -S "%SERVER%" -U "%DBUSER%" -P "%DBPASS%" -Q "BACKUP DATABASE [%DATABASE%] TO DISK = N'!BACKUPFILE!' WITH INIT, STATS = 10" >> "%LOGFILE%" 2>&1
    if errorlevel 1 (
        call :log "ERRO: falha ao gerar o backup de seguranca. Abortando ANTES de rodar qualquer script, por seguranca. Veja o log: %LOGFILE%"
        exit /b 1
    )
    call :log "Backup de seguranca criado com sucesso: !BACKUPFILE!"
)

set "USAOS=1"
for /f "usebackq delims=" %%v in (`sqlcmd -S "%SERVER%" -d "%DATABASE%" -U "%DBUSER%" -P "%DBPASS%" -h -1 -Q "SET NOCOUNT ON; SELECT config_valor FROM configuracao WHERE config_nome = 'OS'" 2^>nul`) do (
    set "TMPVAL=%%v"
    set "TMPVAL=!TMPVAL: =!"
    if not "!TMPVAL!"=="" set "USAOS=!TMPVAL!"
)
call :log "Empresa usa modulo de Ordem de Servico: %USAOS%"

set "FALHAS="
set "EXECUTADOS=0"
set "PULADOS=0"

for /f "usebackq tokens=1,2,3 delims=|" %%N in ("%MANIFESTO%") do (
    set "NUM=%%N"
    set "ARQ=%%O"
    set "CAT=%%P"
    if not "!NUM!"=="" (
        set /a NUMVAL=1!NUM! - 1000
        if !NUMVAL! GTR %CONTINUARAPOS% (
            set "SKIP=0"
            if /i "!CAT!"=="OS" if not "%USAOS%"=="1" set "SKIP=1"
            if "!SKIP!"=="1" (
                call :log "!NUM! - PULADO ^(empresa nao usa OS^): !ARQ!"
                set /a PULADOS+=1
            ) else (
                if not exist "%SCRIPTDIR%!ARQ!" (
                    call :log "!NUM! - PULADO ^(arquivo nao encontrado^): !ARQ!"
                    set "FALHAS=!FALHAS! !NUM!"
                ) else (
                    call :log "!NUM! - Executando [!CAT!]: !ARQ!"
                    sqlcmd -S "%SERVER%" -d "%DATABASE%" -U "%DBUSER%" -P "%DBPASS%" -i "%SCRIPTDIR%!ARQ!" -f 1252 -b >> "%LOGFILE%" 2>&1
                    if errorlevel 1 (
                        call :log "!NUM! - ERRO ao executar !ARQ!. Veja o log: %LOGFILE%"
                        set "FALHAS=!FALHAS! !NUM!"
                        call :log "Parando na primeira falha. Para retomar depois de corrigir, rode: Executar_Scripts.bat "%SERVER%" "%DATABASE%" "%DBUSER%" "%DBPASS%" !NUM!"
                        goto :fim
                    ) else (
                        set /a EXECUTADOS+=1
                    )
                )
            )
        )
    )
)

:fim
call :log "=== Fim da execucao - %EXECUTADOS% script(s) aplicados com sucesso, %PULADOS% pulado(s) ==="
if not "%FALHAS%"=="" (
    call :log "Scripts com falha:%FALHAS%"
    exit /b 1
) else (
    call :log "Nenhuma falha."
    exit /b 0
)

:log
set "MSG=%~1"
echo [%time%] %MSG%
echo [%time%] %MSG% >> "%LOGFILE%"
exit /b 0
