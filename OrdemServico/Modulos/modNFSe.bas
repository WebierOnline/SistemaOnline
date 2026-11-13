Attribute VB_Name = "modNFSe"
Option Explicit

'---------------------------------------------------------------------------
' NFS-e Nacional (API direta do Sefin) - NFSeEmissor.exe (pasta NFSe do sistema, .NET Framework 4.8)
'---------------------------------------------------------------------------
'O exe faz tudo (gerar a nota a partir das OS, transmitir, consultar, cancelar e gerar a DANFSe) e grava nas
'tabelas NFSe/NFSeItens (script 178). A tela so grava o pedido "chave=valor", roda o exe e le o .resultado
'(mesmo padrao do SpedGerador). O pedido tem a senha do banco: o exe apaga assim que le.
'Situacao da nota: 0 Em digitacao, 1 Autorizada, 2 Rejeitada, 3 Cancelada, 4 Pendente (envio sem resposta).

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
