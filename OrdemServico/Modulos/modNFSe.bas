Attribute VB_Name = "modNFSe"
Option Explicit

'---------------------------------------------------------------------------
' NFS-e Nacional (API direta do Sefin) - NFSeEmissor.exe (pasta NFSe do sistema, .NET Framework 4.8)
'---------------------------------------------------------------------------
'O exe faz tudo (gerar a nota a partir das OS, transmitir, consultar, cancelar e gerar a DANFSe) e grava nas
'tabelas NFSe/NFSeItens (script 178). A tela so grava o pedido "chave=valor", roda o exe e le o .resultado
'(mesmo padrao do SpedGerador). O pedido tem a senha do banco: o exe apaga assim que le.
'Situacao da nota: 0 Em digitacao, 1 Autorizada, 2 Rejeitada, 3 Cancelada, 4 Pendente (envio sem resposta).
'Autorizada com AnaliseFiscalSolicitada preenchida (script 179) = cancelamento pedido a prefeitura, aguardando resposta.

Public Const NFSE_DIGITACAO = 0
Public Const NFSE_AUTORIZADA = 1
Public Const NFSE_REJEITADA = 2
Public Const NFSE_CANCELADA = 3
Public Const NFSE_PENDENTE = 4

'Roda o NFSeEmissor com o pedido (a conexao e o usuario sao incluidos aqui). Devolve a 1a linha do resultado
'("OK" ou "ERRO: motivo"); as demais linhas "chave=valor" ficam em oDados (Scripting.Dictionary: idnfse, situacao,
'numero, chave, xml, pdf, mensagem, valor, desconto).
Public Function NFSe_Executar(ByVal sParametros As String, ByRef oDados As Object) As String
Dim sExe As String, sPedido As String, sLinha As String, sStatus As String
Dim nArq As Integer, p As Integer
Dim bAbriu As Boolean
Dim oShell As Object

On Error GoTo deuErro

Set oDados = CreateObject("Scripting.Dictionary")
oDados.CompareMode = 1   'TextCompare
sExe = appPathApp & "NFSe\NFSeEmissor.exe"
If Dir$(sExe) = "" Then
   NFSe_Executar = "ERRO: " & sExe & " nao encontrado (atualize o sistema)."
   Exit Function
End If

sPedido = appPathApp & "NFSe\pedido_" & CStr(vCodFunc) & "_" & Format$(Now, "hhnnss") & ".txt"
If Dir$(sPedido & ".resultado") <> "" Then Kill sPedido & ".resultado"
nArq = FreeFile
Open sPedido For Output As #nArq
Print #nArq, "conexao=" & gConexaoBD
Print #nArq, "usuario=" & NFSe_Usuario()
Print #nArq, sParametros
Close #nArq

Screen.MousePointer = vbHourglass
Set oShell = CreateObject("WScript.Shell")
oShell.Run Chr(34) & sExe & Chr(34) & " " & Chr(34) & sPedido & Chr(34), 0, True
Set oShell = Nothing
Screen.MousePointer = vbDefault
If Dir$(sPedido) <> "" Then Kill sPedido

If Dir$(sPedido & ".resultado") = "" Then
   NFSe_Executar = "ERRO: o emissor de NFS-e nao devolveu resultado (ver NFSe\NFSeEmissor_log.txt)."
   Exit Function
End If

nArq = FreeFile
Open sPedido & ".resultado" For Input As #nArq
bAbriu = True
Line Input #nArq, sStatus
Do While Not EOF(nArq)
   Line Input #nArq, sLinha
   p = InStr(sLinha, "=")
   If p > 1 Then oDados(Left$(sLinha, p - 1)) = Mid$(sLinha, p + 1)
Loop
Close #nArq
bAbriu = False
Kill sPedido & ".resultado"
NFSe_Executar = sStatus
Exit Function

deuErro:
NFSe_Executar = "ERRO: " & Err.Description
Err.Clear
On Error Resume Next
Screen.MousePointer = vbDefault
If bAbriu Then Close #nArq
If sPedido <> "" Then
   If Dir$(sPedido) <> "" Then Kill sPedido
End If
End Function

'Valor de uma chave do resultado ("" se nao veio)
Public Function NFSe_Valor(oDados As Object, ByVal sChave As String) As String
If oDados Is Nothing Then Exit Function
If oDados.Exists(sChave) Then NFSe_Valor = CStr(oDados(sChave))
End Function

'Tabelas do script 178 existem? (empresas que ainda nao atualizaram o banco)
Public Function NFSe_BancoPronto() As Boolean
NFSe_BancoPronto = (Val(SQLExecutaRetorno("SELECT CASE WHEN OBJECT_ID('dbo.NFSe', 'U') IS NULL THEN 0 ELSE 1 END AS ok", "ok", "0")) = 1)
End Function

'Script 179: coluna NFSe.AnaliseFiscalSolicitada (pedido de cancelamento a prefeitura ainda sem resposta)
Public Function NFSe_TemAnaliseFiscal() As Boolean
NFSe_TemAnaliseFiscal = (Val(SQLExecutaRetorno("SELECT CASE WHEN COL_LENGTH('dbo.NFSe', 'AnaliseFiscalSolicitada') IS NULL THEN 0 ELSE 1 END AS ok", "ok", "0")) = 1)
End Function

'Script 180: tabela NFSeADN (NFS-e recebidas/emitidas e eventos baixados do Ambiente de Dados Nacional)
Public Function NFSe_TemADN() As Boolean
NFSe_TemADN = (Val(SQLExecutaRetorno("SELECT CASE WHEN OBJECT_ID('dbo.NFSeADN', 'U') IS NULL THEN 0 ELSE 1 END AS ok", "ok", "0")) = 1)
End Function

Public Function NFSe_SituacaoTexto(ByVal nSituacao As Integer) As String
Select Case nSituacao
   Case NFSE_DIGITACAO: NFSe_SituacaoTexto = "EM DIGITAÇÃO"
   Case NFSE_AUTORIZADA: NFSe_SituacaoTexto = "AUTORIZADA"
   Case NFSE_REJEITADA: NFSe_SituacaoTexto = "REJEITADA"
   Case NFSE_CANCELADA: NFSe_SituacaoTexto = "CANCELADA"
   Case NFSE_PENDENTE: NFSe_SituacaoTexto = "PENDENTE"
   Case Else: NFSe_SituacaoTexto = CStr(nSituacao)
End Select
End Function

'Texto para SQL entre aspas simples
Public Function NFSe_Q(ByVal s As String) As String
NFSe_Q = "'" & Replace(s, "'", "''") & "'"
End Function

'Decimal para SQL (ponto)
Public Function NFSe_N(ByVal v As Variant) As String
If Trim$(CStr(v)) = "" Then
   NFSe_N = "NULL"
Else
   NFSe_N = Replace(Format$(CDbl(v), "0.00"), ",", ".")
End If
End Function

'Abre um arquivo (PDF da DANFSe) no programa padrao do Windows
Public Sub NFSe_AbrirArquivo(ByVal sArquivo As String)
Dim oShell As Object
On Error GoTo deuErro
If sArquivo = "" Then Exit Sub
If Dir$(sArquivo) = "" Then
   MsgBox "Arquivo não encontrado:" & vbCrLf & sArquivo, vbExclamation, "NFS-e"
   Exit Sub
End If
Set oShell = CreateObject("WScript.Shell")
oShell.Run Chr(34) & sArquivo & Chr(34), 1, False
Set oShell = Nothing
Exit Sub
deuErro:
MsgBox "Não foi possível abrir o arquivo:" & vbCrLf & sArquivo & vbCrLf & Err.Description, vbExclamation, "NFS-e"
End Sub

Private Function NFSe_Usuario() As String
On Error Resume Next
If gModoAutomatico Then NFSe_Usuario = "AUTOMATICO": Exit Function
NFSe_Usuario = SQLExecutaRetorno("SELECT nome FROM funcionario WHERE (codigo = " & vCodFunc & ");", "nome", "")
If NFSe_Usuario = "" Then NFSe_Usuario = CStr(vCodFunc)
End Function

'Usado no Reabrir OS: a OS esta numa NFS-e que nao foi cancelada?
Public Function NFSe_DaOS(ByVal lCodOS As Long) As String
Dim s As String
s = "IF OBJECT_ID('dbo.NFSe', 'U') IS NULL SELECT '' AS r ELSE " & _
    "SELECT TOP 1 CASE WHEN n.NumeroNFSe IS NULL THEN 'interna ' + CAST(n.IdNFSe AS varchar(10)) ELSE 'nº ' + n.NumeroNFSe END AS r " & _
    "FROM NFSeItens i INNER JOIN NFSe n ON n.IdNFSe = i.IdNFSe WHERE (i.CodOS = " & lCodOS & ") AND (n.Situacao IN (0, 1, 2, 4))"
NFSe_DaOS = CStr(SQLExecutaRetorno(s, "r", ""))
End Function

'---------------------------------------------------------------------------
' Busca automatica das NFS-e no ADN (so no servidor), igual ao monitor de NF-e destinadas do OnlineCommerce:
' tarefa agendada "OrdemServico_BuscarNFSe" (de hora em hora) roda "Ordem de Servico.exe /buscarnfse", que chama
' o NFSeEmissor (acao=distribuir) e encerra. Assim as recebidas ja estao no banco quando o .rar do mes vai para o
' contador. So e registrada onde a NFS-e esta em uso (empresas sem NFS-e nao ganham a tarefa).
' Desligar numa empresa: configuracao MONITOR_NFSE = 0. Log em BuscarNFSe_log.txt na pasta do sistema.
'---------------------------------------------------------------------------
Private Sub NFSe_LogBusca(ByVal sMsg As String)
Dim nLog As Integer
On Error Resume Next
nLog = FreeFile
Open appPathApp & "BuscarNFSe_log.txt" For Append As #nLog
Print #nLog, Format$(Now, "dd/mm/yyyy hh:nn:ss") & " - " & sMsg
Close #nLog
End Sub

'NFS-e em uso: configuracao OS = 1 e nota ja autorizada/cancelada pelo sistema ou documento ja baixado do ADN
Public Function NFSe_EmUso() As Boolean
On Error Resume Next
If Trim$(CStr(SQLExecutaRetorno("SELECT config_valor FROM configuracao WHERE (config_nome = 'OS')", "config_valor", "1"))) <> "1" Then Exit Function
If Not NFSe_TemADN() Then Exit Function
NFSe_EmUso = (Val(SQLExecutaRetorno("SELECT CASE WHEN EXISTS (SELECT 1 FROM NFSe WHERE Situacao IN (1, 3)) " & _
                                    "OR EXISTS (SELECT 1 FROM NFSeADN) THEN 1 ELSE 0 END AS ok", "ok", "0")) = 1)
End Function

Private Function NFSe_BuscaLigada() As Boolean
NFSe_BuscaLigada = (Trim$(CStr(SQLExecutaRetorno("SELECT config_valor FROM configuracao WHERE (config_nome = 'MONITOR_NFSE')", "config_valor", "1"))) <> "0")
End Function

'Chamado pelo Senha_OS.Form_Load quando o sistema e iniciado com /buscarnfse
Public Sub NFSe_BuscarADNAuto()
Dim oRes As Object
Dim sStatus As String
On Error GoTo deuErro
If Not NFSe_BuscaLigada() Then
   NFSe_LogBusca "Busca desligada (configuracao MONITOR_NFSE = 0)."
   Exit Sub
End If
If Not NFSe_EmUso() Then
   NFSe_LogBusca "NFS-e nao esta em uso nesta empresa (sem script 180 ou sem notas): nada a fazer."
   Exit Sub
End If
sStatus = NFSe_Executar("acao=distribuir", oRes)
If Left$(sStatus, 2) = "OK" Then
   NFSe_LogBusca "OK: " & NFSe_Valor(oRes, "mensagem")
Else
   NFSe_LogBusca sStatus
End If
Exit Sub
deuErro:
NFSe_LogBusca "Erro na busca automatica: " & Err.Description
Err.Clear
End Sub

'Chamado no login (Senha_OS): cria a tarefa de hora em hora, so no servidor e so com a NFS-e em uso
Public Sub NFSe_RegistrarTarefaADN()
Dim oShell As Object
Dim sExe As String
On Error Resume Next
If Not EhServidorLocal() Then Exit Sub
If Not NFSe_EmUso() Then Exit Sub
If Not NFSe_BuscaLigada() Then Exit Sub
Set oShell = CreateObject("WScript.Shell")
If oShell.Run("schtasks /Query /TN ""OrdemServico_BuscarNFSe""", 0, True) <> 0 Then
   sExe = """\""" & appPathApp & appEXEName & "\"" /buscarnfse"""
   oShell.Run "schtasks /Create /TN ""OrdemServico_BuscarNFSe"" /TR " & sExe & " /SC HOURLY /F", 0, True
End If
Set oShell = Nothing
End Sub
