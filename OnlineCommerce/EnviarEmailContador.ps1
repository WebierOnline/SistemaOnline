param(
    [Parameter(Mandatory=$true)][string]$Anexos,
    [Parameter(Mandatory=$true)][int]$Mes,
    [Parameter(Mandatory=$true)][int]$Ano
)

# Envia o email mensal para o contador usando a DLL fiscal snfe.Util.
# Existe como script SEPARADO (em vez de fazer isso direto no
# ExportarXMLServidor.vbs) porque o VBScript nao consegue passar um array
# ByRef fortemente tipado (String()) pro EmailEnviar - da erro "Tipos
# incompativeis". O PowerShell consegue, usando [string[]] + [ref].
# Chamado pelo ExportarXMLServidor.vbs (Function EnviarEmailContador).

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$logPath = Join-Path $scriptDir "EnviarEmailContador_log.txt"

function Log($msg) {
    Add-Content -Path $logPath -Value "$(Get-Date) - $msg"
}

function LerIni($caminho, $secao, $chave) {
    if (-not (Test-Path $caminho)) { return "" }
    $secaoAtual = ""
    foreach ($linha in Get-Content $caminho) {
        $l = $linha.Trim()
        if ($l.StartsWith("[") -and $l.EndsWith("]")) {
            $secaoAtual = $l.Substring(1, $l.Length - 2)
        } elseif ($secaoAtual -ieq $secao) {
            $pos = $l.IndexOf("=")
            if ($pos -gt 0) {
                $chaveAtual = $l.Substring(0, $pos).Trim()
                if ($chaveAtual -ieq $chave) {
                    return $l.Substring($pos + 1).Trim()
                }
            }
        }
    }
    return ""
}

try {
    Log "----- Inicio envio email contador (Mes=$Mes Ano=$Ano) -----"
    Log "Anexos recebidos: $Anexos"

    $vIP = LerIni (Join-Path $scriptDir "config.ini") "IP_MAQUINA" "ip"
    if ([string]::IsNullOrEmpty($vIP)) { $vIP = "localhost\SQLEXPRESS2008" }

    $conn = New-Object -ComObject ADODB.Connection
    $conn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;DRIVER={Sql Server};SERVER=$vIP;uid=sa;pwd=190106web;DATABASE=cyber_base;Connect Timeout=600;TRUSTED_CONNECTION=NO"
    $conn.Open()

    $rs = $conn.Execute("SELECT TOP 1 DiretorioXML, CNPJ, Razao, Estado, AmbienteNF, CertificadoDigital, NFCeIDToken, NFCeCSC, LicencaDLL, Email, caminho FROM empresa ORDER BY fantasia")
    if ($rs.EOF) { throw "Nenhuma empresa cadastrada" }

    $vDiretorioXML = $rs.Fields.Item("DiretorioXML").Value
    $vCNPJ = $rs.Fields.Item("CNPJ").Value
    $vRazao = $rs.Fields.Item("Razao").Value
    $vEstado = $rs.Fields.Item("Estado").Value
    $vAmbienteNF = $rs.Fields.Item("AmbienteNF").Value
    $vCertificadoDigital = $rs.Fields.Item("CertificadoDigital").Value
    $vNFCeIDToken = $rs.Fields.Item("NFCeIDToken").Value
    $vNFCeCSC = $rs.Fields.Item("NFCeCSC").Value
    $vLicencaDLL = $rs.Fields.Item("LicencaDLL").Value
    $vEmailEmpresa = $rs.Fields.Item("Email").Value
    $vCaminhoDANFe = $rs.Fields.Item("caminho").Value
    $rs.Close()

    $rsCont = $conn.Execute("SELECT TOP 1 Email FROM TbContabilista")
    if ($rsCont.EOF) { throw "Nenhum contador cadastrado em TbContabilista" }
    $vEmailContador = $rsCont.Fields.Item("Email").Value
    $rsCont.Close()
    $conn.Close()

    if ([string]::IsNullOrEmpty($vEmailContador)) { throw "Email do contador esta vazio" }

    if ($vDiretorioXML.Substring($vDiretorioXML.Length - 1) -ne "\") { $vDiretorioXML += "\" }

    $obj = New-Object -ComObject snfe.Util
    $vNFCeIDTokenPad = $vNFCeIDToken.ToString().PadLeft(6, '0')

    $obj.ConfigurarDLL("", $vCertificadoDigital, "", 1, "$($vDiretorioXML)nfe\", "$($vDiretorioXML)nfe\schemas", $vEstado, 55, "1", "30000", $vAmbienteNF.ToString().Substring(0,1), "1", 4, $vCNPJ, $vNFCeIDTokenPad, $vNFCeCSC, "02.382.419/0001-80", "OnLine Info", $false, $vLicencaDLL) | Out-Null
    $obj.ConfigurarEmail("mail.ekklesiasoft.com.br", 587, 100000, $true, "dev@ekklesiasoft.com.br", "Ekk29639780", $true, $vEmailEmpresa, $vRazao) | Out-Null
    $obj.ConfigurarDANFe($vCaminhoDANFe, $true, $false, $true, $false) | Out-Null
    $obj.certificadoAvisaVencimento = $false
    $obj.certificadoDiasAviso = 10
    $obj.exibirAvisos = $false
    $obj.CarregarConfiguracoes() | Out-Null

    [string[]] $arrAnexos = $Anexos -split '\|'
    [string[]] $arrCC = @($vEmailContador)

    $assunto = "Arquivos XML ref. Mes " + $Mes.ToString("00") + "/" + $Ano
    $corpo = "Segue em anexo os arquivos XML/RAR de NFe e NFCe emitidos no periodo. <br><br>Atenciosamente, <br><br>$vRazao"

    $ret = $obj.EmailEnviar($vEmailContador, $assunto, $corpo, ([ref]$arrAnexos), ([ref]$arrCC))

    # O envio real parece ser assincrono internamente na DLL - matar o processo logo
    # apos a chamada corta a entrega antes de completar. Da tempo antes de sair.
    Start-Sleep -Seconds 25

    try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($obj) | Out-Null } catch {}

    if ($ret) {
        Log "Email enviado com sucesso para $vEmailContador ($($arrAnexos.Count) anexo(s)) - a entrega real pode demorar alguns minutos"
        exit 0
    } else {
        Log "EmailEnviar retornou False"
        exit 1
    }
} catch {
    Log "ERRO: $($_.Exception.Message)"
    exit 1
}
