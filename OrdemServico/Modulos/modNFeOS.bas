Attribute VB_Name = "modNFeOS"
Option Explicit
'---------------------------------------------------------------------------
' NF-e (modelo 55) das PEÇAS da OS (botão NFe da tela OS_Recapadora).
' A NF-e continua sendo feita na tela NF-e do Online Commerce (NFe_Completa: natureza, CFOP, transporte,
' ICMS interestadual/ST, duplicatas). O botão só leva o pedido da OS até lá:
'   1. grava o pedido em NFe_OS.txt na pasta do Online Commerce;
'   2. abre o OnlineCommerce.exe /nfeos (se ele já está aberto nesta máquina, a nova instância sai calada);
'   (antes: cadastro do cliente completo para NF-e, senão oferece abrir o Cadastro de Clientes e só segue completo)
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

'O que falta no cadastro do cliente para a NF-e ("" = completo). MESMA regra do ClienteCamposFaltando da
'NFe_Completa do Online Commerce (se mudar lá, mudar aqui): assim a OS já avisa antes de abrir o Online Commerce.
Private Function NFeOS_ClienteFaltando(ByVal lCod As Long) As String
Dim r As ADODB.Recordset
Dim bSN As Boolean, bJur As Boolean, sDoc As String, iTipoCtr As Integer, sTipo As String, sFalta As String
NFeOS_ClienteFaltando = ""
If lCod <= 0 Then NFeOS_ClienteFaltando = "  - Cliente não informado no pedido": Exit Function
Set r = dbData.OpenRecordset("SELECT *, tipo AS vTipo, nome AS vNome FROM cliente WHERE codigo = " & lCod)
If r.EOF Then
   r.Close
   NFeOS_ClienteFaltando = "  - Cliente não encontrado"
   Exit Function
End If
bSN = (ValidateNull(r("SN")) = True)
sDoc = RemoverFormato(IIf(IsNull(r("cpf")), "", r("cpf")))
iTipoCtr = Val(IIf(IsNull(r("TipoContribuinte")), 0, r("TipoContribuinte")))
sTipo = UCase(Trim(IIf(IsNull(r("vTipo")), "", r("vTipo"))))

If Vazio(r("vNome")) Then sFalta = sFalta & vbCrLf & "  - Nome"
If Vazio(r("endereco")) Then sFalta = sFalta & vbCrLf & "  - Endereço"
If Vazio(r("numero")) And Not bSN Then sFalta = sFalta & vbCrLf & "  - Número"
If Vazio(r("bairro")) Then
   sFalta = sFalta & vbCrLf & "  - Bairro"
ElseIf Len(r("bairro") & "") < 4 Then
   sFalta = sFalta & vbCrLf & "  - Bairro (mínimo 4 caracteres)"
End If
If Vazio(r("cidade")) Then sFalta = sFalta & vbCrLf & "  - Cidade"
If Vazio(r("estado")) Then sFalta = sFalta & vbCrLf & "  - Estado"
If Len(RemoverFormato(r("CEP") & "")) <> 8 Then sFalta = sFalta & vbCrLf & "  - CEP"
If Vazio(r("CodigoIBGE")) Then
   sFalta = sFalta & vbCrLf & "  - Código IBGE da cidade"
ElseIf Len(CStr(r("CodigoIBGE"))) <> 7 Then
   sFalta = sFalta & vbCrLf & "  - Código IBGE da cidade (precisa ter 7 dígitos)"
End If

'documento: JURIDICA = CNPJ (14); FISICA/RURAL = CPF (11); sem tipo cadastrado, usa o contribuinte (9 = CPF)
If sTipo = "" Then bJur = (iTipoCtr <> 9) Else bJur = (sTipo = "JURÍDICA")
If bJur Then
   If Len(sDoc) <> 14 Then
      sFalta = sFalta & vbCrLf & "  - CNPJ"
   ElseIf Not Validar_CNPJ(sDoc) Then
      sFalta = sFalta & vbCrLf & "  - CNPJ inválido"
   End If
Else
   If Len(sDoc) <> 11 Then
      sFalta = sFalta & vbCrLf & "  - CPF"
   ElseIf Not Validar_CPF(sDoc) Then
      sFalta = sFalta & vbCrLf & "  - CPF inválido"
   End If
End If

If iTipoCtr = 0 Then sFalta = sFalta & vbCrLf & "  - Tipo de contribuinte"
If iTipoCtr = 1 And Vazio(r("ie")) Then sFalta = sFalta & vbCrLf & "  - Inscrição Estadual (IE)"

If (sTipo = "FÍSICA" And iTipoCtr <> 9 And iTipoCtr <> 0) Or (sTipo = "JURÍDICA" And iTipoCtr = 9) Or _
   (sTipo = "RURAL" And (iTipoCtr = 9 Or iTipoCtr = 2)) Then
   sFalta = sFalta & vbCrLf & "  - Tipo de contribuinte incompatível com o tipo de pessoa (" & sTipo & ")"
End If

r.Close
Set r = Nothing
If Len(sFalta) > 0 Then NFeOS_ClienteFaltando = Mid$(sFalta, Len(vbCrLf) + 1)
End Function

'Cadastro de Clientes (Compartilhado) aberto no cliente, forçando o tipo CADASTRO (todos os campos), como a NFe_Completa faz
Public Sub NFeOS_AbrirCadastroCliente(ByVal lCod As Long)
Load Clientes_Cadastro
Clientes_Cadastro.bForcarCadastroNFe = True
Clientes_Cadastro.SSTab1.Tab = 0
Clientes_Cadastro.cmdNovo.Enabled = False
Clientes_Cadastro.cmdSalvar.Enabled = False
Clientes_Cadastro.cmdCancelar.Enabled = False
Clientes_Cadastro.txtCodigo.Text = CStr(lCod)
Clientes_Cadastro.Show vbModal
End Sub

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
Dim lPed As Long, lCli As Long
Dim sNota As String, sPasta As String, sFalta As String
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

'cadastro do cliente completo para NF-e: só segue com ele completo (o Online Commerce também confere na conversão)
lCli = Val(SQLExecutaRetorno("SELECT TOP 1 cod_cliente AS c FROM pedidos WHERE (cod_pedido = " & lPed & ")", "c", "0"))
sFalta = NFeOS_ClienteFaltando(lCli)
If sFalta <> "" Then
   If lCli <= 0 Then
      MsgBox "O pedido da OS " & lCodOS & " não tem cliente: a NF-e precisa do cliente com cadastro completo.", vbExclamation, "NF-e"
      Exit Function
   End If
   If MsgBox("O cadastro do cliente da OS " & lCodOS & " está incompleto para a NF-e. Falta:" & vbCrLf & vbCrLf & sFalta & vbCrLf & vbCrLf & _
             "Só é possível prosseguir com a NF-e completando o cadastro do cliente." & vbCrLf & _
             "Abrir o cadastro do cliente agora?", vbExclamation + vbYesNo, "NF-e") = vbNo Then Exit Function
   NFeOS_AbrirCadastroCliente lCli
   sFalta = NFeOS_ClienteFaltando(lCli)
   If sFalta <> "" Then
      MsgBox "O cadastro do cliente ainda está incompleto. Falta:" & vbCrLf & vbCrLf & sFalta & vbCrLf & vbCrLf & _
             "A NF-e não foi aberta: complete o cadastro e clique em NFe de novo.", vbExclamation, "NF-e"
      Exit Function
   End If
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
