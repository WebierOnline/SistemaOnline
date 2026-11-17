Attribute VB_Name = "modNFeOS"
Option Explicit
'---------------------------------------------------------------------------
' NF-e (modelo 55) das PEÇAS da OS (botão NFe da tela OS_Recapadora).
' A NF-e continua sendo feita na tela NF-e do Online Commerce (NFe_Completa: natureza, CFOP, transporte,
' ICMS interestadual/ST, duplicatas). O botão só leva o pedido da OS até lá:
'   1. grava o pedido em NFe_OS.txt na pasta do Online Commerce;
'   2. abre o OnlineCommerce.exe /nfeos (se ele já está aberto nesta máquina, a nova instância sai calada);
'   3. o timer tmrNFeOS da Tela_Principal do Online Commerce lê o arquivo e abre a NF-e com o pedido
'      convertido (NFe_Completa.CarregarPedidoOS = mesmo caminho do botão Converter da aba PEDIDOS).
' Os itens da NF-e vêm de pedidos_itens (só as peças, com o desconto já rateado pela OS); os serviços vão na NFS-e.
' Pasta do Online Commerce: config.ini [ONLINE_COMMERCE] caminho=, senão a pasta da OS, senão ..\OnlineCommerce.
'---------------------------------------------------------------------------

Private Declare Function AllowSetForegroundWindow Lib "user32" (ByVal dwProcessId As Long) As Long
Private Const ASFW_ANY As Long = -1

Private Function NFeOS_TemTabela() As Boolean
NFeOS_TemTabela = (Val(SQLExecutaRetorno("SELECT CASE WHEN OBJECT_ID('dbo.NotaFiscal', 'U') IS NULL OR OBJECT_ID('dbo.NotaFiscalItens', 'U') IS NULL THEN 0 ELSE 1 END AS ok", "ok", "0")) = 1)
End Function

'OS terminadas com NF-e autorizada (peças do pedido da OS numa NF-e enviada, não cancelada): "|12|15|" (coluna NFe do Grid_OS)
Public Function NFeOS_ComNFe() As String
Dim r As ADODB.Recordset
Dim s As String
s = "|"
If NFeOS_TemTabela() Then
   Set r = dbData.OpenRecordset("SELECT DISTINCT o.COD_OS FROM OS o " & _
                                "WHERE (o.STATUS_OS = 1) AND (o.COD_PEDIDO > 0) AND EXISTS (SELECT 1 FROM NotaFiscalItens i INNER JOIN NotaFiscal n ON (n.CodigoNota = i.CodigoNota) " & _
                                "WHERE (i.Cod_Pedido = o.COD_PEDIDO) AND (ISNULL(n.Enviada, 0) = 1) AND (ISNULL(n.Cancelada, 0) = 0) AND (ISNULL(n.Denegada, 0) = 0) AND (ISNULL(n.Inutilizada, 0) = 0));")
   Do While Not r.EOF
      s = s & r("COD_OS") & "|"
      r.MoveNext
   Loop
   r.Close
   Set r = Nothing
End If
NFeOS_ComNFe = s
End Function

'NF-e com as peças do pedido (cancelada, denegada ou inutilizada não conta): "nº 123 autorizada" / "nº 123 em digitação" / ""
Public Function NFeOS_NotaDoPedido(ByVal lPed As Long) As String
Dim r As ADODB.Recordset
NFeOS_NotaDoPedido = ""
If lPed <= 0 Then Exit Function
If Not NFeOS_TemTabela() Then Exit Function
Set r = dbData.OpenRecordset("SELECT TOP 1 n.NumeroNota, CAST(ISNULL(n.Enviada, 0) AS int) AS env FROM NotaFiscal n " & _
                             "WHERE EXISTS (SELECT 1 FROM NotaFiscalItens i WHERE (i.CodigoNota = n.CodigoNota) AND (i.Cod_Pedido = " & lPed & ")) " & _
                             "AND (ISNULL(n.Cancelada, 0) = 0) AND (ISNULL(n.Denegada, 0) = 0) AND (ISNULL(n.Inutilizada, 0) = 0) " & _
                             "ORDER BY CAST(ISNULL(n.Enviada, 0) AS int) DESC, n.CodigoNota DESC;")
If Not r.EOF Then
   NFeOS_NotaDoPedido = "nº " & ValidateNull(r("NumeroNota")) & IIf(Val(ValidateNull(r("env"))) <> 0, " autorizada", " em digitação no Online Commerce")
End If
r.Close
Set r = Nothing
End Function

'NFC-e ativa do pedido (cancelada ou inutilizada não conta): "nº 45 autorizada" / "nº 45 ainda não autorizada" / ""
Private Function NFeOS_NFCeDoPedido(ByVal lPed As Long) As String
Dim r As ADODB.Recordset
NFeOS_NFCeDoPedido = ""
If Val(SQLExecutaRetorno("SELECT CASE WHEN OBJECT_ID('dbo.TbNFCe', 'U') IS NULL THEN 0 ELSE 1 END AS ok", "ok", "0")) <> 1 Then Exit Function
Set r = dbData.OpenRecordset("SELECT TOP 1 NumeNota, CAST(ISNULL(NFCeEnviada, 0) AS int) AS env FROM TbNFCe " & _
                             "WHERE (Num_OS_VD_Origem = " & lPed & ") AND (ISNULL(NFCeCancelada, 0) = 0) AND (ISNULL(Inutilizada, 0) = 0) ORDER BY IdNFProd DESC;")
If Not r.EOF Then
   NFeOS_NFCeDoPedido = "nº " & ValidateNull(r("NumeNota")) & IIf(Val(ValidateNull(r("env"))) <> 0, " autorizada", " ainda não autorizada")
End If
r.Close
Set r = Nothing
End Function

'Pasta do Online Commerce (com "\" no fim) ou "" se não achar o OnlineCommerce.exe
Private Function NFeOS_PastaOnline() As String
Dim oIni As Ini
Dim sCam As String
Dim aPastas(2) As String
Dim i As Integer
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
sCam = Trim$(oIni.LerTexto("ONLINE_COMMERCE", "caminho"))
Set oIni = Nothing
If LCase$(Right$(sCam, 4)) = ".exe" Then sCam = Left$(sCam, InStrRev(sCam, "\"))
aPastas(0) = sCam
aPastas(1) = appPathApp
aPastas(2) = appPathApp & "..\OnlineCommerce\"
For i = 0 To 2
   sCam = aPastas(i)
   If sCam <> "" Then
      If Right$(sCam, 1) <> "\" Then sCam = sCam & "\"
      If Dir$(sCam & "OnlineCommerce.exe") <> "" Then
         NFeOS_PastaOnline = sCam
         Exit Function
      End If
   End If
Next
NFeOS_PastaOnline = ""
End Function

'Leva o pedido da OS terminada para a NF-e do Online Commerce (veja o cabeçalho)
Public Function NFeOS_Abrir(ByVal lCodOS As Long) As Boolean
Dim r As ADODB.Recordset
Dim lPed As Long
Dim sNota As String, sPasta As String
Dim curPecas As Currency
Dim f As Integer

On Error GoTo deuErro

If Not NFeOS_TemTabela() Then
   MsgBox "O banco de dados não tem as tabelas da NF-e (NotaFiscal).", vbExclamation, "NF-e"
   Exit Function
End If

'OS terminada, com pedido
Set r = dbData.OpenRecordset("SELECT o.COD_PEDIDO, CAST(ISNULL(o.STATUS_OS, 0) AS int) AS st, CAST(ISNULL(p.cancelado, 0) AS int) AS canc " & _
                             "FROM OS o LEFT JOIN pedidos p ON (p.COD_PEDIDO = o.COD_PEDIDO) WHERE (o.COD_OS = " & lCodOS & ");")
If r.EOF Then
   r.Close
   MsgBox "OS não encontrada.", vbExclamation, "NF-e"
   Exit Function
End If
lPed = Val(ValidateNull(r("COD_PEDIDO")))
If Val(ValidateNull(r("st"))) <> 1 Or lPed = 0 Then
   r.Close
   MsgBox "A OS " & lCodOS & " precisa estar terminada (financeiro fechado) para ter NF-e.", vbExclamation, "NF-e"
   Exit Function
End If
If Val(ValidateNull(r("canc"))) <> 0 Then
   r.Close
   MsgBox "O pedido da OS " & lCodOS & " está cancelado.", vbExclamation, "NF-e"
   Exit Function
End If
r.Close
Set r = Nothing

'as peças saem numa nota só: NF-e ou NFC-e
sNota = NFeOS_NotaDoPedido(lPed)
If sNota <> "" Then
   If InStr(1, sNota, "autorizada") > 0 Then
      MsgBox "A OS " & lCodOS & " já tem a NF-e " & sNota & ".", vbInformation, "NF-e"
      Exit Function
   End If
   'em digitação: segue e abre o Online Commerce (a tela de NF-e avisa onde ela está)
End If
sNota = NFeOS_NFCeDoPedido(lPed)
If sNota <> "" Then
   MsgBox "As peças da OS " & lCodOS & " já estão na NFC-e " & sNota & "." & vbCrLf & _
          "Para fazer NF-e, a NFC-e precisa ser cancelada (ou inutilizada, se não foi autorizada).", vbExclamation, "NF-e"
   Exit Function
End If

'só as peças (itens do pedido); serviço vai na NFS-e
curPecas = CCur(SQLExecutaRetorno("SELECT ISNULL(SUM(Total), 0) AS t FROM pedidos_itens WHERE (cod_pedido = " & lPed & ")", "t", "0"))
If curPecas <= 0 Then
   MsgBox "A OS " & lCodOS & " não tem peças/produtos: a NF-e é só das peças (os serviços vão na NFS-e).", vbInformation, "NF-e"
   Exit Function
End If

sPasta = NFeOS_PastaOnline()
If sPasta = "" Then
   MsgBox "Não achei o Online Commerce (OnlineCommerce.exe) neste computador: a NF-e é feita na tela de NF-e dele." & vbCrLf & vbCrLf & _
          "Informe a pasta no config.ini da OS:" & vbCrLf & "[ONLINE_COMMERCE]" & vbCrLf & "caminho=C:\pasta\do\OnlineCommerce", vbExclamation, "NF-e"
   Exit Function
End If

'pedido para o Online Commerce (lido pelo timer da Tela_Principal dele)
f = FreeFile
Open sPasta & "NFe_OS.txt" For Output As #f
Print #f, "pedido=" & lPed
Print #f, "os=" & lCodOS
Print #f, "hora=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
Close #f
f = 0

'deixa o Online Commerce (já aberto ou não) vir para a frente
AllowSetForegroundWindow ASFW_ANY
Shell """" & sPasta & "OnlineCommerce.exe"" /nfeos", vbNormalFocus
NFeOS_Abrir = True
Exit Function

deuErro:
If f <> 0 Then Close #f
MsgBox "Erro ao abrir a NF-e da OS no Online Commerce: " & Err.Description, vbCritical, "NF-e"
End Function
