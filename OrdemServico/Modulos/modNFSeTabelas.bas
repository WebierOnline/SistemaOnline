Attribute VB_Name = "modNFSeTabelas"
Option Explicit

'Tabelas oficiais da NFS-e (arquivos NFSe_*.txt da pasta NFSe, os mesmos que o NFSeEmissor usa para conferir antes
'de transmitir), usadas no Configurar NFS-e (padrão da empresa) e no Editar NFS-e (cada nota):
'serviço (Anexo B) > NBS > indicador da operação (Anexo VII) > cClassTrib, cada lista filtrada pela anterior (Anexo VIII).
'Os combos guardam "código - descrição"; NFSeTab_CodigoDoCombo devolve só os dígitos do código escolhido.

'Anexo B: serviços (código de tributação nacional)
Private mSrvCod() As String, mSrvDesc() As String, mSrvQtd As Long
'Anexo VII: indicadores da operação usados na NFS-e
Private mIndCod() As String, mIndTxt() As String, mIndImovel() As Boolean, mIndQtd As Long
'Anexo VIII: item da LC 116 x NBS x indOp x cClassTrib
Private mCorItem() As String, mCorNBS() As String, mCorNBSDesc() As String, mCorInd() As String
Private mCorClass() As String, mCorClassDesc() As String, mCorQtd As Long
Private mTabelasOk As Boolean

Private Const CB_SETDROPPEDWIDTH = &H160
Private Declare Function NFSeTab_SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'---------------------------------------------------------------------------------------------------- carga

'Lê as tabelas uma vez por sessão (se faltarem os arquivos, tenta de novo na próxima chamada)
Public Function NFSeTab_Carregar() As Boolean
Dim c As Collection
Dim v As Variant, p() As String
Dim i As Long

If mTabelasOk Then NFSeTab_Carregar = True: Exit Function
On Error GoTo deuErro
If Not LerTabela("NFSe_Servicos.txt", 2, c) Then Exit Function
mSrvQtd = c.Count
ReDim mSrvCod(1 To mSrvQtd): ReDim mSrvDesc(1 To mSrvQtd)
i = 0
For Each v In c
   i = i + 1
   p = Split(v, ";")
   mSrvCod(i) = p(0)
   mSrvDesc(i) = p(1)
Next

If Not LerTabela("NFSe_IndOp.txt", 4, c) Then Exit Function
mIndQtd = c.Count
ReDim mIndCod(1 To mIndQtd): ReDim mIndTxt(1 To mIndQtd): ReDim mIndImovel(1 To mIndQtd)
i = 0
For Each v In c
   i = i + 1
   p = Split(v, ";")
   mIndCod(i) = p(0)
   mIndTxt(i) = p(1) & " - " & p(2)
   mIndImovel(i) = (p(3) = "1")
Next

If Not LerTabela("NFSe_Correlacao.txt", 6, c) Then Exit Function
mCorQtd = c.Count
ReDim mCorItem(1 To mCorQtd): ReDim mCorNBS(1 To mCorQtd): ReDim mCorNBSDesc(1 To mCorQtd)
ReDim mCorInd(1 To mCorQtd): ReDim mCorClass(1 To mCorQtd): ReDim mCorClassDesc(1 To mCorQtd)
i = 0
For Each v In c
   i = i + 1
   p = Split(v, ";")
   mCorItem(i) = p(0)
   mCorNBS(i) = p(1)
   mCorNBSDesc(i) = p(2)
   mCorInd(i) = p(3)
   mCorClass(i) = p(4)
   mCorClassDesc(i) = p(5)
Next
mTabelasOk = True
NFSeTab_Carregar = True
Exit Function

deuErro:
mTabelasOk = False
mSrvQtd = 0: mIndQtd = 0: mCorQtd = 0
End Function

Public Function NFSeTab_Ok() As Boolean
NFSeTab_Ok = mTabelasOk
End Function

Private Function LerTabela(ByVal sArquivo As String, ByVal nCampos As Integer, ByRef c As Collection) As Boolean
Dim nArq As Integer
Dim sLinha As String, sCaminho As String
sCaminho = appPathApp & "NFSe\" & sArquivo
Set c = New Collection
If Dir$(sCaminho) = "" Then Exit Function
nArq = FreeFile
Open sCaminho For Input As #nArq
Do While Not EOF(nArq)
   Line Input #nArq, sLinha
   If UBound(Split(sLinha, ";")) >= nCampos - 1 Then c.Add sLinha
Loop
Close #nArq
LerTabela = (c.Count > 0)
End Function

'---------------------------------------------------------------------------------------------------- combos

'Lista aberta mais larga que a caixa (descrições longas), em pixels
Public Sub NFSeTab_LarguraLista(cbo As ComboBox, ByVal nPixels As Long)
NFSeTab_SendMessage cbo.hwnd, CB_SETDROPPEDWIDTH, nPixels, ByVal 0&
End Sub

'Código no começo do texto do combo ("140101 - ...", "1.2001.31.10 - ..."); "(...)" = nada escolhido
Public Function NFSeTab_CodigoDoCombo(cbo As ComboBox) As String
Dim t As String
If cbo.ListIndex < 0 Then Exit Function
t = cbo.List(cbo.ListIndex)
If Left$(t, 1) = "(" Then Exit Function
If InStr(t, " - ") > 0 Then t = Left$(t, InStr(t, " - ") - 1)
NFSeTab_CodigoDoCombo = SoDigitos(t)
End Function

Public Sub NFSeTab_SelecionarPorCodigo(cbo As ComboBox, ByVal sCodigo As String)
Dim i As Integer
If sCodigo = "" Then Exit Sub
For i = 0 To cbo.ListCount - 1
   If SoDigitos(Split(cbo.List(i) & " - ", " - ")(0)) = sCodigo Then
      cbo.ListIndex = i
      Exit Sub
   End If
Next
End Sub

Public Function NFSeTab_JaNaLista(cbo As ComboBox, ByVal sCodigo As String) As Boolean
Dim i As Integer
For i = 0 To cbo.ListCount - 1
   If SoDigitos(Split(cbo.List(i) & " - ", " - ")(0)) = sCodigo Then NFSeTab_JaNaLista = True: Exit Function
Next
End Function

Public Sub NFSeTab_PreencherServicos(cboServ As ComboBox, ByVal sPreferido As String)
Dim i As Long
cboServ.Clear
For i = 1 To mSrvQtd
   cboServ.AddItem mSrvCod(i) & " - " & Curto(mSrvDesc(i), 160)
Next
If sPreferido <> "" And Not NFSeTab_JaNaLista(cboServ, sPreferido) Then
   cboServ.AddItem sPreferido & " - (código gravado, fora da lista nacional: escolha outro)", 0
End If
NFSeTab_SelecionarPorCodigo cboServ, sPreferido
End Sub

'NBS do serviço escolhido. Devolve quantos NBS a tabela tem para ele (0 = serviço sem NBS: sem grupo IBS/CBS).
Public Function NFSeTab_PreencherNBS(cboServ As ComboBox, cboNBS As ComboBox, ByVal sPreferido As String) As Long
Dim i As Long, sItem As String, n As Long
cboNBS.Clear
cboNBS.AddItem "(sem NBS - não informar)"
sItem = ItemDoServico(NFSeTab_CodigoDoCombo(cboServ))
For i = 1 To mCorQtd
   If mCorItem(i) = sItem Then
      If Not NFSeTab_JaNaLista(cboNBS, mCorNBS(i)) Then
         cboNBS.AddItem NBSFormatado(mCorNBS(i)) & " - " & Curto(mCorNBSDesc(i), 150)
         n = n + 1
      End If
   End If
Next
'sem tabelas (ou NBS gravado fora dela): mantém o gravado para não perder
If sPreferido <> "" And Not NFSeTab_JaNaLista(cboNBS, sPreferido) And (Not mTabelasOk Or n = 0) Then
   cboNBS.AddItem NBSFormatado(sPreferido) & " - (NBS gravado)"
End If
cboNBS.ListIndex = 0
NFSeTab_SelecionarPorCodigo cboNBS, sPreferido
If cboNBS.ListIndex = 0 And n = 1 Then cboNBS.ListIndex = 1
NFSeTab_PreencherNBS = n
End Function

'indOp do Anexo VIII para serviço+NBS, os do mesmo grupo (050101 -> 050102) e os da regra de operação não onerosa/
'exterior (100101/100301/100501); os de imóvel ficam de fora (exigem os dados do imóvel, que a OS não tem)
Public Sub NFSeTab_PreencherIndOp(cboServ As ComboBox, cboNBS As ComboBox, cboInd As ComboBox, ByVal sPreferido As String)
Dim i As Long, j As Long, sItem As String, sNBS As String, sPrimeiro As String
cboInd.Clear
sItem = ItemDoServico(NFSeTab_CodigoDoCombo(cboServ))
sNBS = NFSeTab_CodigoDoCombo(cboNBS)
If sNBS <> "" Then
   For i = 1 To mCorQtd
      If mCorItem(i) = sItem And mCorNBS(i) = sNBS Then
         If sPrimeiro = "" Then sPrimeiro = mCorInd(i)
         For j = 1 To mIndQtd
            If Left$(mIndCod(j), 4) = Left$(mCorInd(i), 4) Then AdicionarIndOp cboInd, mIndCod(j)
         Next
         Select Case mCorInd(i)
            Case "100101": AdicionarIndOp cboInd, "100102": AdicionarIndOp cboInd, "100201"
            Case "100301": AdicionarIndOp cboInd, "100302": AdicionarIndOp cboInd, "100401"
            Case "100501": AdicionarIndOp cboInd, "100502": AdicionarIndOp cboInd, "100601"
         End Select
      End If
   Next
End If
If sPreferido <> "" And Not NFSeTab_JaNaLista(cboInd, sPreferido) And Not mTabelasOk Then
   cboInd.AddItem sPreferido & " - (indicador gravado)"
End If
NFSeTab_SelecionarPorCodigo cboInd, sPreferido
If cboInd.ListIndex < 0 Then NFSeTab_SelecionarPorCodigo cboInd, sPrimeiro
If cboInd.ListIndex < 0 And cboInd.ListCount > 0 Then cboInd.ListIndex = 0
End Sub

Private Sub AdicionarIndOp(cboInd As ComboBox, ByVal sCod As String)
Dim k As Long
k = IndiceIndOp(sCod)
If k = 0 Then Exit Sub
If mIndImovel(k) Then Exit Sub
If NFSeTab_JaNaLista(cboInd, sCod) Then Exit Sub
cboInd.AddItem sCod & " - " & Curto(mIndTxt(k), 200)
End Sub

Public Sub NFSeTab_PreencherClassTrib(cboServ As ComboBox, cboNBS As ComboBox, cboClass As ComboBox, ByVal sPreferido As String)
Dim i As Long, sItem As String, sNBS As String
cboClass.Clear
sItem = ItemDoServico(NFSeTab_CodigoDoCombo(cboServ))
sNBS = NFSeTab_CodigoDoCombo(cboNBS)
If sNBS <> "" Then
   For i = 1 To mCorQtd
      If mCorItem(i) = sItem And mCorNBS(i) = sNBS Then
         If Not NFSeTab_JaNaLista(cboClass, mCorClass(i)) Then cboClass.AddItem mCorClass(i) & " - " & Curto(mCorClassDesc(i), 150)
      End If
   Next
End If
If sPreferido <> "" And Not NFSeTab_JaNaLista(cboClass, sPreferido) And Not mTabelasOk Then
   cboClass.AddItem sPreferido & " - (classificação gravada)"
End If
NFSeTab_SelecionarPorCodigo cboClass, sPreferido
If cboClass.ListIndex < 0 And cboClass.ListCount > 0 Then cboClass.ListIndex = 0
End Sub

'Serviço feito no estabelecimento x em outro local: no grupo do indicador escolhido (4 primeiros dígitos) troca para o
'final 01 (estabelecimento do fornecedor) ou 02 (local diverso), quando o par existir na lista. Ex.: 050101 <-> 050102.
Public Sub NFSeTab_IndOpPorLocal(cboInd As ComboBox, ByVal bForaDoEstabelecimento As Boolean)
Dim sCod As String
sCod = NFSeTab_CodigoDoCombo(cboInd)
If Len(sCod) <> 6 Then Exit Sub
If Right$(sCod, 2) <> "01" And Right$(sCod, 2) <> "02" Then Exit Sub
NFSeTab_SelecionarPorCodigo cboInd, Left$(sCod, 4) & IIf(bForaDoEstabelecimento, "02", "01")
End Sub

'---------------------------------------------------------------------------------------------------- apoio

Private Function SoDigitos(ByVal s As String) As String
Dim i As Integer
For i = 1 To Len(s)
   If Mid$(s, i, 1) Like "[0-9]" Then SoDigitos = SoDigitos & Mid$(s, i, 1)
Next
End Function

Private Function Curto(ByVal s As String, ByVal n As Integer) As String
Curto = IIf(Len(s) > n, Left$(s, n - 3) & "...", s)
End Function

Private Function NBSFormatado(ByVal s As String) As String
If Len(s) = 9 Then
   NBSFormatado = Left$(s, 1) & "." & Mid$(s, 2, 4) & "." & Mid$(s, 6, 2) & "." & Mid$(s, 8, 2)
Else
   NBSFormatado = s
End If
End Function

'Item do Anexo VIII do serviço: desdobro ("99.03.02") se a tabela tiver, senão subitem ("14.01")
Private Function ItemDoServico(ByVal sCod As String) As String
Dim i As Long, sDesd As String
If Len(sCod) <> 6 Then Exit Function
sDesd = Left$(sCod, 2) & "." & Mid$(sCod, 3, 2) & "." & Mid$(sCod, 5, 2)
For i = 1 To mCorQtd
   If mCorItem(i) = sDesd Then ItemDoServico = sDesd: Exit Function
Next
ItemDoServico = Left$(sCod, 2) & "." & Mid$(sCod, 3, 2)
End Function

Private Function IndiceIndOp(ByVal sCod As String) As Long
Dim i As Long
For i = 1 To mIndQtd
   If mIndCod(i) = sCod Then IndiceIndOp = i: Exit Function
Next
End Function
