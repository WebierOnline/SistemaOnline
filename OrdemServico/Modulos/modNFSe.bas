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

'Nota sem tomador identificado (cliente sem CPF/CNPJ válido no cadastro quando a nota foi gerada):
'"" = tem CPF/CNPJ; "PJ" = sem documento e o cliente é pessoa JURÍDICA; "PF" = sem documento (pessoa física)
Public Function NFSe_SemTomador(ByVal lId As Long) As String
Dim r As ADODB.Recordset
Dim sDoc As String, sTipo As String
Set r = dbData.OpenRecordset("SELECT NFSe.TomadorCpfCnpj, cliente.Tipo FROM NFSe LEFT JOIN cliente ON (cliente.CODIGO = NFSe.CodCliente) " & _
                             "WHERE (NFSe.IdNFSe = " & lId & ");")
If Not r.EOF Then
   sDoc = Trim$(ValidateNull(r("TomadorCpfCnpj")))
   sTipo = UCase$(Trim$(ValidateNull(r("Tipo"))))
   If Len(sDoc) <> 11 And Len(sDoc) <> 14 Then
      If Left$(sTipo, 3) = "JUR" Then NFSe_SemTomador = "PJ" Else NFSe_SemTomador = "PF"
   End If
End If
r.Close
End Function

'Antes de transmitir: nota sem tomador só para pessoa física que não pediu o CPF na nota.
'Empresa (CNPJ) é obrigada a ser identificada: avisa forte e o padrão é NÃO transmitir.
Public Function NFSe_ConfirmarSemTomador(ByVal lId As Long) As Boolean
Select Case NFSe_SemTomador(lId)
   Case ""
      NFSe_ConfirmarSemTomador = True
   Case "PJ"
      NFSe_ConfirmarSemTomador = (MsgBox("ATENÇÃO: o cliente desta nota está cadastrado como EMPRESA (pessoa jurídica) e está SEM CNPJ." & vbCrLf & vbCrLf & _
         "A NFS-e vai sair SEM TOMADOR IDENTIFICADO. Empresa é obrigada a ser identificada na nota (é com ela que a empresa " & _
         "lança a despesa na contabilidade) e a nota autorizada não pode ser corrigida depois." & vbCrLf & vbCrLf & _
         "O certo: corrija o CNPJ no cadastro do cliente, exclua este rascunho e gere a NFS-e de novo." & vbCrLf & vbCrLf & _
         "Transmitir assim mesmo, sem o tomador?", vbExclamation + vbYesNo + vbDefaultButton2, "NFS-e") = vbYes)
   Case Else
      NFSe_ConfirmarSemTomador = (MsgBox("A NFS-e vai sair SEM TOMADOR IDENTIFICADO: o cliente está sem CPF no cadastro." & vbCrLf & vbCrLf & _
         "Isso só vale para pessoa física que não pediu o CPF na nota. Se o cliente quer a nota no nome dele, " & _
         "informe o CPF no cadastro, exclua este rascunho e gere a NFS-e de novo." & vbCrLf & vbCrLf & _
         "Transmitir sem o tomador?", vbQuestion + vbYesNo + vbDefaultButton1, "NFS-e") = vbYes)
End Select
End Function

'O que falta no cadastro do cliente para a NFS-e sair completa ("" = nada). Diferente da NF-e, nada disso impede
'a NFS-e: sem CPF/CNPJ a nota sai sem tomador identificado; sem endereço/CEP/IBGE, o tomador sai sem endereço
'(regras do NFSeEmissor: TomadorCliente.Ler + Servico, endereço só com município IBGE, logradouro e CEP de 8 dígitos)
Public Function NFSe_ClienteFaltando(ByVal lCod As Long) As String
Dim r As ADODB.Recordset
Dim sDoc As String, sTipo As String, sFalta As String
Dim bJur As Boolean, bSemEnd As Boolean
NFSe_ClienteFaltando = ""
If lCod <= 0 Then Exit Function
Set r = dbData.OpenRecordset("SELECT * FROM cliente WHERE (CODIGO = " & lCod & ");")
If r.EOF Then
   r.Close
   Exit Function
End If
sDoc = RemoverFormato(ValidateNull(r("CPF")) & "")
sTipo = UCase$(Trim$(ValidateNull(r("Tipo")) & ""))
bJur = (Left$(sTipo, 3) = "JUR")

If Vazio(r("Nome")) Then sFalta = sFalta & vbCrLf & "  - Nome"
If bJur Then
   If Len(sDoc) <> 14 Then
      sFalta = sFalta & vbCrLf & "  - CNPJ (empresa precisa ser identificada na nota)"
   ElseIf Not Validar_CNPJ(sDoc) Then
      sFalta = sFalta & vbCrLf & "  - CNPJ inválido"
   End If
Else
   If Len(sDoc) <> 11 Then
      sFalta = sFalta & vbCrLf & "  - CPF (sem ele a nota sai sem tomador identificado)"
   ElseIf Not Validar_CPF(sDoc) Then
      sFalta = sFalta & vbCrLf & "  - CPF inválido"
   End If
End If

If Vazio(r("Endereco")) Then sFalta = sFalta & vbCrLf & "  - Endereço": bSemEnd = True
If Len(RemoverFormato(ValidateNull(r("CEP")) & "")) <> 8 Then sFalta = sFalta & vbCrLf & "  - CEP": bSemEnd = True
If Len(RemoverFormato(ValidateNull(r("CodigoIBGE")) & "")) <> 7 Then sFalta = sFalta & vbCrLf & "  - Código IBGE da cidade": bSemEnd = True
If bSemEnd Then sFalta = sFalta & vbCrLf & "    (sem endereço, CEP e código IBGE o tomador sai sem endereço na nota)"

r.Close
Set r = Nothing
If Len(sFalta) > 0 Then NFSe_ClienteFaltando = Mid$(sFalta, Len(vbCrLf) + 1)
End Function

'Antes de gerar a NFS-e da OS (botão NFSe): avisa o que falta no cadastro do cliente.
'Sim = abre o Cadastro de Clientes e confere de novo; Não = gera assim mesmo; Cancelar = não gera.
Public Function NFSe_ConferirCliente(ByVal lCodOS As Long) As Boolean
Dim lCli As Long
Dim sFalta As String
Dim iResp As VbMsgBoxResult
NFSe_ConferirCliente = True
lCli = Val(SQLExecutaRetorno("SELECT TOP 1 COD_CLIENTE AS c FROM OS WHERE (COD_OS = " & lCodOS & ")", "c", "0"))
Do
   sFalta = NFSe_ClienteFaltando(lCli)
   If sFalta = "" Then Exit Function
   iResp = MsgBox("O cadastro do cliente da OS " & lCodOS & " está incompleto para a NFS-e. Falta:" & vbCrLf & vbCrLf & sFalta & vbCrLf & vbCrLf & _
                  "Depois de autorizada, a NFS-e não pode ser corrigida." & vbCrLf & vbCrLf & _
                  "Sim = abrir o cadastro do cliente para completar" & vbCrLf & _
                  "Não = gerar a NFS-e assim mesmo" & vbCrLf & _
                  "Cancelar = não gerar agora", vbExclamation + vbYesNoCancel, "NFS-e")
   Select Case iResp
      Case vbYes
         NFeOS_AbrirCadastroCliente lCli
      Case vbNo
         Exit Function
      Case Else
         NFSe_ConferirCliente = False
         Exit Function
   End Select
Loop
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
