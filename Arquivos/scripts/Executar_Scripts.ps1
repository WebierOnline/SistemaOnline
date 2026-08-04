<#
.SINOPSE
    Executa os scripts .sql de C:\Projeto\Arquivos\scripts na ordem definida
    pelo manifesto _manifesto.txt (formato "NNN|nome_do_arquivo.sql|CATEGORIA").

    CATEGORIA pode ser:
      GERAL - roda sempre
      OS    - so roda se configuracao.config_valor = 1 onde config_nome = 'OS'
              (empresa usa o modulo de Ordem de Servico)

.EXEMPLOS
    # Rodar tudo no banco padrao (mesma maquina desta sessao)
    .\Executar_Scripts.ps1

    # Instalar num cliente novo
    .\Executar_Scripts.ps1 -Server "SERVIDOR-CLIENTE\SQLEXPRESS" -Database "cyber_base" -User "sa" -Password "senha_do_cliente"

    # Retomar a partir do script 045 (por exemplo, depois de corrigir um erro no 044)
    .\Executar_Scripts.ps1 -ContinuarApos 44

    # Nao parar no primeiro erro (roda tudo e reporta as falhas no final)
    .\Executar_Scripts.ps1 -PararNoPrimeiroErro:$false

    # Pular o backup de seguranca automatico (uso local/teste repetido, NAO usar em cliente)
    .\Executar_Scripts.ps1 -PularBackup
#>

param(
    [string]$Server = ".\SQLEXPRESS2008",
    [string]$Database = "cyber_base",
    [string]$User = "sa",
    [string]$Password = "190106web",
    [int]$ContinuarApos = 0,
    [bool]$PararNoPrimeiroErro = $true,
    [switch]$PularBackup
)

$PastaScripts = $PSScriptRoot
$Manifesto = Join-Path $PastaScripts "_manifesto.txt"
$LogFile = Join-Path $PastaScripts ("execucao_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".log")

if (-not (Test-Path $Manifesto)) {
    Write-Error "Manifesto nao encontrado: $Manifesto"
    exit 1
}

$sqlcmdPath = (Get-Command sqlcmd -ErrorAction SilentlyContinue)
if (-not $sqlcmdPath) {
    Write-Error "sqlcmd nao encontrado no PATH. Instale o 'sqlcmd Utility' (parte do SQL Server Command Line Utilities) na maquina do cliente antes de rodar este script."
    exit 1
}

function Log($msg) {
    $linha = "[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $msg
    Write-Host $linha
    Add-Content -Path $LogFile -Value $linha -Encoding UTF8
}

function ObterUsaOS {
    param($Server, $Database, $User, $Password)
    try {
        $conn = New-Object -ComObject ADODB.Connection
        $connStr = "Provider=SQLOLEDB.1;Data Source=$Server;Initial Catalog=$Database;User ID=$User;Password=$Password;"
        $conn.Open($connStr)
        $rs = New-Object -ComObject ADODB.Recordset
        $rs.Open("SELECT config_valor FROM configuracao WHERE config_nome = 'OS'", $conn)
        $resultado = $true
        if (-not $rs.EOF) {
            $valor = $rs.Fields.Item("config_valor").Value
            $resultado = ([string]$valor -eq "1")
        } else {
            Log "AVISO: nao existe linha config_nome='OS' em configuracao - assumindo que a empresa USA OS (scripts OS nao serao pulados)."
        }
        if ($rs.State -ne 0) { $rs.Close() }
        $conn.Close()
        return $resultado
    } catch {
        Log ("AVISO: nao foi possivel consultar configuracao.config_valor (OS): {0} - assumindo que a empresa USA OS (scripts OS nao serao pulados)." -f $_.Exception.Message)
        return $true
    }
}

function Fazer-BackupSeguranca {
    param($Server, $Database, $User, $Password, $LogFile)
    try {
        $conn = New-Object -ComObject ADODB.Connection
        $connStr = "Provider=SQLOLEDB.1;Data Source=$Server;Initial Catalog=$Database;User ID=$User;Password=$Password;"
        $conn.Open($connStr)
        $rs = New-Object -ComObject ADODB.Recordset
        $rs.Open("SELECT TOP 1 physical_name FROM sys.master_files WHERE database_id = DB_ID('$Database') AND type_desc = 'ROWS'", $conn)
        if ($rs.EOF) {
            if ($rs.State -ne 0) { $rs.Close() }
            $conn.Close()
            return $null
        }
        $caminhoDados = $rs.Fields.Item("physical_name").Value
        $rs.Close()
        $conn.Close()
        return (Split-Path $caminhoDados -Parent)
    } catch {
        Log ("ERRO ao localizar a pasta de dados do banco '$Database': {0}" -f $_.Exception.Message)
        return $null
    }
}

Log "=== Iniciando execucao - Servidor: $Server | Banco: $Database ==="

if ($PularBackup) {
    Log "AVISO: backup de seguranca PULADO (-PularBackup). Nao usar essa opcao em banco de cliente."
} else {
    Log "Localizando pasta de dados do banco para gerar backup de seguranca..."
    $pastaDados = Fazer-BackupSeguranca -Server $Server -Database $Database -User $User -Password $Password -LogFile $LogFile
    if (-not $pastaDados) {
        Log "ERRO: nao foi possivel localizar a pasta de dados do banco '$Database'. Abortando execucao ANTES de rodar qualquer script, por seguranca."
        exit 1
    }

    $timestampBackup = Get-Date -Format "yyyyMMdd_HHmmss"
    $caminhoBackup = Join-Path $pastaDados ("$Database" + "_pre_scripts_$timestampBackup.bak")
    Log ("Gerando backup de seguranca em: $caminhoBackup")

    $saidaBackup = & sqlcmd -S $Server -U $User -P $Password -Q "BACKUP DATABASE [$Database] TO DISK = N'$caminhoBackup' WITH INIT, STATS = 10"
    $saidaBackup | ForEach-Object { Add-Content -Path $LogFile -Value $_ -Encoding UTF8 }

    if ($LASTEXITCODE -ne 0) {
        Log "ERRO: falha ao gerar o backup de seguranca (sqlcmd saiu com codigo $LASTEXITCODE). Abortando execucao ANTES de rodar qualquer script, por seguranca."
        Log "Se o erro for de espaco em disco ou permissao, resolva e rode de novo. Para pular o backup (nao recomendado em cliente), use -PularBackup."
        exit 1
    }
    Log ("Backup de seguranca criado com sucesso: $caminhoBackup")
}

$usaOS = ObterUsaOS -Server $Server -Database $Database -User $User -Password $Password
Log ("Empresa usa modulo de Ordem de Servico: {0}" -f $usaOS)

$linhas = Get-Content -Path $Manifesto -Encoding Default | Where-Object { $_.Trim() -ne "" }

$falhas = @()
$pulados = @()
$executados = 0

foreach ($linha in $linhas) {
    $partes = $linha -split '\|'
    if ($partes.Count -lt 2) {
        Log "AVISO: linha do manifesto ignorada (formato invalido): $linha"
        continue
    }
    $numero = [int]$partes[0]
    $nomeArquivo = $partes[1].Trim()
    $categoria = if ($partes.Count -ge 3) { $partes[2].Trim().ToUpper() } else { "GERAL" }

    if ($numero -le $ContinuarApos) {
        continue
    }

    if ($categoria -eq "OS" -and -not $usaOS) {
        Log ("{0:D3} | PULADO (empresa nao usa OS): {1}" -f $numero, $nomeArquivo)
        $pulados += $numero
        continue
    }

    $caminhoScript = Join-Path $PastaScripts $nomeArquivo
    if (-not (Test-Path $caminhoScript)) {
        Log ("{0:D3} | PULADO (arquivo nao encontrado): {1}" -f $numero, $nomeArquivo)
        $falhas += $numero
        if ($PararNoPrimeiroErro) { break }
        continue
    }

    Log ("{0:D3} | Executando [{1}]: {2}" -f $numero, $categoria, $nomeArquivo)

    $saida = & sqlcmd -S $Server -d $Database -U $User -P $Password -i "$caminhoScript" -b
    $saida | ForEach-Object { Add-Content -Path $LogFile -Value $_ -Encoding UTF8 }
    $saida | ForEach-Object { Write-Host $_ }

    if ($LASTEXITCODE -ne 0) {
        Log ("{0:D3} | ERRO ao executar {1} (sqlcmd saiu com codigo {2}). Veja o log: {3}" -f $numero, $nomeArquivo, $LASTEXITCODE, $LogFile)
        $falhas += $numero
        if ($PararNoPrimeiroErro) {
            Log ("Parando na primeira falha. Para retomar depois de corrigir, rode:  .\Executar_Scripts.ps1 -Server '$Server' -Database '$Database' -ContinuarApos {0}" -f ($numero - 1))
            break
        }
    } else {
        $executados++
    }
}

Log ("=== Fim da execucao - $executados script(s) aplicados com sucesso, $($pulados.Count) pulado(s) ===")
if ($pulados.Count -gt 0) {
    Log ("Scripts pulados (categoria OS, empresa nao usa OS): " + ($pulados -join ", "))
}
if ($falhas.Count -gt 0) {
    Log ("Scripts com falha: " + ($falhas -join ", "))
    exit 1
} else {
    Log "Nenhuma falha."
    exit 0
}
