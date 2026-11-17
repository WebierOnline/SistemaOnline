Attribute VB_Name = "modNFCeOS"
Option Explicit
'---------------------------------------------------------------------------
' NFC-e das PEÇAS da OS (botão NFCe da tela OS_Recapadora).
' Mesmo procedimento fiscal do botão NFCe do Estornar do PDV (PDV.GerarEImprimirNFCeParaPedido):
' procedure NFCeIncluir do pedido da OS + TransmitirNFCe/DANFCe do modNFe (Compartilhado) - o PDV não é alterado.
' Diferença da OS: o pedido guarda o total da OS inteira (peças + serviços) e as parcelas cobrem tudo,
' mas os itens do pedido são só as peças. Então, depois da NFCeIncluir:
'   - o cabeçalho (valor dos produtos, desconto, outras, frete) passa a ser a soma dos itens;
'   - os pagamentos (parcelas da OS) são reduzidos na proporção das peças (soma = valor da NFC-e).
' O desconto geral da OS já vem rateado em cada peça/serviço pela tela da OS (RatearDescontoProdutosServicosOS,
' opção Produtos/Serviços/Tudo); o acréscimo geral da OS não entra na NFC-e.
' Os serviços vão na NFS-e (modNFSe).
'---------------------------------------------------------------------------

Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long

Private Function NFCeOS_TemInternet() As Boolean
Dim vFlags As Long
NFCeOS_TemInternet = (InternetGetConnectedState(vFlags, 0) <> 0)
End Function

Private Function NFCeOS_Num(ByVal v As Variant) As String
NFCeOS_Num = Replace(CStr(CCur(v)), ",", ".")
End Function

'Valor (campo "v") lido pela conexão do dbData: dentro da transação o SQLExecutaRetorno (outra conexão, vgDb)
'ficaria esperando as linhas ainda não confirmadas
Private Function NFCeOS_Valor(ByVal sSQL As String) As Variant
Dim r As ADODB.Recordset
NFCeOS_Valor = 0
Set r = dbData.OpenRecordset(sSQL)
If Not r.EOF Then
   If Not IsNull(r("v")) Then NFCeOS_Valor = r("v")
End If
r.Close
Set r = Nothing
End Function

Private Function NFCeOS_TemTabela() As Boolean
NFCeOS_TemTabela = (Val(SQLExecutaRetorno("SELECT CASE WHEN OBJECT_ID('dbo.TbNFCe', 'U') IS NULL THEN 0 ELSE 1 END AS ok", "ok", "0")) = 1)
End Function

'OS terminadas com NFC-e autorizada (pelo pedido da OS): "|12|15|" (coluna NFCe do Grid_OS)
Public Function NFCeOS_ComNFCe() As String
Dim r As ADODB.Recordset
Dim s As String
s = "|"
If NFCeOS_TemTabela() Then
   Set r = dbData.OpenRecordset("SELECT DISTINCT o.COD_OS FROM OS o INNER JOIN TbNFCe n ON (n.Num_OS_VD_Origem = o.COD_PEDIDO) " & _
                                "WHERE (o.STATUS_OS = 1) AND (ISNULL(n.NFCeEnviada, 0) = 1) AND (ISNULL(n.NFCeCancelada, 0) = 0) AND (ISNULL(n.Inutilizada, 0) = 0);")
   Do While Not r.EOF
      s = s & r("COD_OS") & "|"
      r.MoveNext
   Loop
   r.Close
   Set r = Nothing
End If
NFCeOS_ComNFCe = s
End Function

'Gera, transmite e imprime a NFC-e das peças da OS. Se a NFC-e já existe e não foi autorizada, oferece transmitir de novo.
Public Function NFCeOS_Gerar(ByVal lCodOS As Long) As Boolean
Dim r As ADODB.Recordset
Dim rItens As ADODB.Recordset
Dim lPed As Long, lId As Long
Dim sTipoPgto As String, sNum As String
Dim curPecas As Currency, curNF As Currency, curParc As Currency, curDif As Currency
Dim bTrans As Boolean, bContingencia As Boolean
Dim sErro As String

On Error GoTo deuErro

If Not NFCeOS_TemTabela() Then
   MsgBox "O banco de dados não tem as tabelas da NFC-e.", vbExclamation, "NFC-e"
   Exit Function
End If

'OS terminada, com pedido
Set r = dbData.OpenRecordset("SELECT o.COD_PEDIDO, CAST(ISNULL(o.STATUS_OS, 0) AS int) AS st, p.TIPO_PAGAMENTO, CAST(ISNULL(p.cancelado, 0) AS int) AS canc " & _
                             "FROM OS o LEFT JOIN pedidos p ON (p.COD_PEDIDO = o.COD_PEDIDO) WHERE (o.COD_OS = " & lCodOS & ");")
If r.EOF Then
   r.Close
   MsgBox "OS não encontrada.", vbExclamation, "NFC-e"
   Exit Function
End If
lPed = Val(ValidateNull(r("COD_PEDIDO")))
sTipoPgto = UCase$(Trim$(ValidateNull(r("TIPO_PAGAMENTO"))))
If Val(ValidateNull(r("st"))) <> 1 Or lPed = 0 Then
   r.Close
   MsgBox "A OS " & lCodOS & " precisa estar terminada (financeiro fechado) para ter NFC-e.", vbExclamation, "NFC-e"
   Exit Function
End If
If Val(ValidateNull(r("canc"))) <> 0 Then
   r.Close
   MsgBox "O pedido da OS " & lCodOS & " está cancelado.", vbExclamation, "NFC-e"
   Exit Function
End If
r.Close

'NFC-e ativa do pedido da OS (cancelada ou inutilizada não conta: libera emitir uma nova)
Set r = dbData.OpenRecordset("SELECT TOP 1 IdNFProd, NumeNota, CAST(ISNULL(NFCeEnviada, 0) AS int) AS env " & _
                             "FROM TbNFCe WHERE (Num_OS_VD_Origem = " & lPed & ") AND (ISNULL(NFCeCancelada, 0) = 0) AND (ISNULL(Inutilizada, 0) = 0) ORDER BY IdNFProd DESC;")
If Not r.EOF Then
   lId = Val(ValidateNull(r("IdNFProd")))
   sNum = ValidateNull(r("NumeNota"))
   If Val(ValidateNull(r("env"))) <> 0 Then
      r.Close
      MsgBox "A OS " & lCodOS & " já tem a NFC-e nº " & sNum & " autorizada.", vbInformation, "NFC-e"
      Exit Function
   End If
   r.Close
   If MsgBox("A NFC-e nº " & sNum & " da OS " & lCodOS & " foi criada mas ainda não foi autorizada pela SEFAZ." & vbCrLf & vbCrLf & _
             "Transmitir agora?", vbQuestion + vbYesNo, "NFC-e") <> vbYes Then Exit Function
   If Not NFCeOS_MaquinaPronta() Then Exit Function
   bContingencia = NFCeOS_Contingencia()
   NFCeOS_Gerar = NFCeOS_Transmitir(lId, lCodOS, bContingencia)
   Exit Function
End If
r.Close
Set r = Nothing

'as peças saem numa nota só: NFC-e ou NF-e (a NF-e é feita no Online Commerce, modNFeOS)
sNum = NFeOS_NotaDoPedido(lPed)
If sNum <> "" Then
   MsgBox "As peças da OS " & lCodOS & " já estão na NF-e " & sNum & "." & vbCrLf & _
          "Para fazer NFC-e, a NF-e precisa ser cancelada (ou excluída, se ainda está em digitação).", vbExclamation, "NFC-e"
   Exit Function
End If

'só as peças (itens do pedido); serviço vai na NFS-e
curPecas = CCur(SQLExecutaRetorno("SELECT ISNULL(SUM(Total), 0) AS t FROM pedidos_itens WHERE (cod_pedido = " & lPed & ")", "t", "0"))
If curPecas <= 0 Then
   MsgBox "A OS " & lCodOS & " não tem peças/produtos: a NFC-e é só das peças (os serviços vão na NFS-e).", vbInformation, "NFC-e"
   Exit Function
End If

If Not NFCeOS_MaquinaPronta() Then Exit Function

If (sTipoPgto = "À PRAZO" Or sTipoPgto = "A PRAZO") And _
   UCase$(Trim$(CStr(SQLExecutaRetorno("SELECT config_valor FROM configuracao WHERE (config_nome = 'CONFPRAZONFCE')", "config_valor", "")))) <> "SIM" Then
   MsgBox "A emissão de NFC-e para vendas à prazo está desativada nas configurações.", vbExclamation, "NFC-e"
   Exit Function
End If

If MsgBox("Emitir a NFC-e das peças da OS " & lCodOS & " (R$ " & Format$(curPecas, ocMONEY) & ")?", vbQuestion + vbYesNo, "NFC-e") <> vbYes Then Exit Function

bContingencia = NFCeOS_Contingencia()

dbData.Execute "BEGIN TRANSACTION"
bTrans = True

dbData.Execute "EXEC NFCeIncluir " & lPed
lId = Val(NFCeOS_Valor("SELECT ISNULL(MAX(IdNFProd), 0) AS v FROM TbNFCe WHERE (Num_OS_VD_Origem = " & lPed & ")"))
If lId = 0 Then
   dbData.Execute "ROLLBACK TRANSACTION"
   bTrans = False
   MsgBox "Não foi possível gerar a NFC-e da OS " & lCodOS & ".", vbExclamation, "NFC-e"
   Exit Function
End If

'cabeçalho = soma dos itens (o pedido da OS tem o total com os serviços)
dbData.Execute "UPDATE TbNFCe SET " & _
   "Valor_NF_Prod = (SELECT ISNULL(SUM(ROUND(i.QtdeMov * i.ValorUnit, 2)), 0) FROM TbNFCe_Itens i WHERE (i.IdNFProd = " & lId & ")), " & _
   "DescontoPromocional = (SELECT ISNULL(SUM(i.Desconto), 0) FROM TbNFCe_Itens i WHERE (i.IdNFProd = " & lId & ")), " & _
   "OutrasDespesasAces = (SELECT ISNULL(SUM(i.ValorOutras), 0) FROM TbNFCe_Itens i WHERE (i.IdNFProd = " & lId & ")), " & _
   "Valor_Frete = (SELECT ISNULL(SUM(i.Valor_Frete), 0) FROM TbNFCe_Itens i WHERE (i.IdNFProd = " & lId & ")) " & _
   "WHERE (IdNFProd = " & lId & ");"
curNF = CCur(NFCeOS_Valor("SELECT (ISNULL(Valor_NF_Prod, 0) - ISNULL(DescontoPromocional, 0) + ISNULL(OutrasDespesasAces, 0) + ISNULL(Valor_Frete, 0)) AS v " & _
                               "FROM TbNFCe WHERE (IdNFProd = " & lId & ")"))

PreencherReformaTributariaNFCe lId, dbData.ActiveConnection

'pagamentos: parcelas da OS na proporção das peças (a diferença do arredondamento fica na maior)
curParc = CCur(NFCeOS_Valor("SELECT ISNULL(SUM(VALOR), 0) AS v FROM parcelas WHERE (COD_PEDIDO = " & lPed & ")"))
If curParc > 0 And curNF > 0 Then
   dbData.Execute "INSERT INTO TbNFCe_Faturas (IdNFProd, IDParcela, TipoPgto, Vencimento, Valor, IdBandeira, CartaoNumeroAutorizacao) " & _
                  "SELECT " & lId & ", NUMERO, dbo.NFCeFormaPagto(FORMA_PGTO, TIPO_CARTAO), DATA, " & _
                  "ROUND(VALOR * " & NFCeOS_Num(curNF) & " / " & NFCeOS_Num(curParc) & ", 2), '01', '' " & _
                  "FROM parcelas WHERE (COD_PEDIDO = " & lPed & ")"
   curDif = curNF - CCur(NFCeOS_Valor("SELECT ISNULL(SUM(Valor), 0) AS v FROM TbNFCe_Faturas WHERE (IdNFProd = " & lId & ")"))
   If curDif <> 0 Then
      dbData.Execute "UPDATE TbNFCe_Faturas SET Valor = Valor + (" & NFCeOS_Num(curDif) & ") " & _
                     "WHERE (IdNFProd = " & lId & ") AND (IDParcela = (SELECT TOP 1 IDParcela FROM TbNFCe_Faturas WHERE (IdNFProd = " & lId & ") ORDER BY Valor DESC))"
   End If
   dbData.Execute "DELETE FROM TbNFCe_Faturas WHERE (IdNFProd = " & lId & ") AND (Valor <= 0)"
End If

'dado fiscal dos produtos: com erro desfaz tudo (corrige o cadastro e clica de novo)
Set rItens = dbData.OpenRecordset("SELECT IdNFProd, IdNFProd_Item, IDProduto, CodBarras, DescricaoProduto, CodNcm, CFOP, Bc_Icms, ICMSCST, IPICST, COFINSCST, PISCST, UN " & _
                                  "FROM TbNFCe_Itens WHERE (IdNFProd = " & lId & ");")
If NFCeOS_ItensComErro(rItens) Then
   rItens.Close
   dbData.Execute "ROLLBACK TRANSACTION"
   bTrans = False
   MsgBox "Algum produto da OS " & lCodOS & " está com dado fiscal incompleto (NCM, CFOP, CST, unidade ou código de barras)." & vbCrLf & _
          "Corrija o cadastro do produto e clique em NFCe de novo.", vbExclamation, "NFC-e"
   Exit Function
End If
rItens.Close
Set rItens = Nothing

dbData.Execute "COMMIT TRANSACTION"
bTrans = False

NFCeOS_Gerar = NFCeOS_Transmitir(lId, lCodOS, bContingencia)
Exit Function

deuErro:
sErro = Err.Description
On Error Resume Next
If bTrans Then dbData.Execute "ROLLBACK TRANSACTION"
On Error GoTo 0
MsgBox "Erro ao gerar a NFC-e: " & sErro, vbCritical, "NFC-e"
End Function

'Máquina configurada para emitir NFC-e (config.ini [IMPRIMIR_NFCE] resposta=SIM) + impressora da NFC-e
Private Function NFCeOS_MaquinaPronta() As Boolean
Dim oIni As Ini
Dim Prt As Printer
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
If UCase$(Trim$(oIni.LerTexto("IMPRIMIR_NFCE", "resposta"))) <> "SIM" Then
   Set oIni = Nothing
   MsgBox "Esse computador não está configurado para emitir NFC-e." & vbCrLf & vbCrLf & _
          "No config.ini:" & vbCrLf & "[IMPRIMIR_NFCE]" & vbCrLf & "resposta=SIM" & vbCrLf & vbCrLf & _
          "[IMPRESSORA_NFCE]" & vbCrLf & "impressora=(nome da impressora da NFC-e)", vbExclamation, "NFC-e"
   Exit Function
End If
var_ImpNFCe = oIni.LerTexto("IMPRESSORA_NFCE", "impressora")
Set oIni = Nothing
For Each Prt In Printers
   If Prt.DeviceName = var_ImpNFCe Then
      Set Printer = Prt
      Exit For
   End If
Next
NFCeOS_MaquinaPronta = True
End Function

'Contingência off-line: a da empresa, ou sem internet nesta máquina (igual ao PDV: grava NFCeOffline = 1)
Private Function NFCeOS_Contingencia() As Boolean
NFCeOS_Contingencia = (Val(SQLExecutaRetorno("SELECT TOP 1 CAST(ISNULL(NFCeOffline, 0) AS int) AS v FROM empresa ORDER BY fantasia", "v", "0")) <> 0)
If Not NFCeOS_TemInternet() Then
   NFCeOS_Contingencia = True
   dbData.Execute "UPDATE Empresa SET NFCeOffline = 1"
   MsgBox "Sem conexão com a internet neste computador." & vbCr & _
          "A NFC-e será gerada em modo CONTINGÊNCIA (offline) e transmitida à SEFAZ assim que a conexão voltar.", vbExclamation, "NFC-e - Sem Internet"
End If
End Function

'Transmite e imprime (DANFC-e na impressora da NFC-e), como no PDV
Private Function NFCeOS_Transmitir(ByVal lId As Long, ByVal lCodOS As Long, ByVal bContingencia As Boolean) As Boolean
Dim bOk As Boolean
Dim sistNFe As snfe.Util

DoEvents
bOk = TransmitirNFCe(lId, "1", Not bContingencia, "65", True)
If Not bOk And Not bContingencia Then NFCeOS_SugerirContingencia

If bOk Then
   Set sistNFe = New snfe.Util
   ConfiguraDLLNFeNFCe 65, "1", sistNFe
   If Not bContingencia Then
      Call sistNFe.DANFCeImprimir(xCaminhoXML, True, var_ImpNFCe, True, xCaminhoPDF, 0, False, False, "")
   Else
      Call sistNFe.DANFCeOFFImprimir(xCaminhoXML, True, var_ImpNFCe, True, xCaminhoPDF, 0, False, False, "")
   End If
   Set sistNFe = Nothing
   NFCeOS_Transmitir = True
Else
   MsgBox "A NFC-e da OS " & lCodOS & " foi criada mas não foi autorizada pela SEFAZ." & vbCrLf & _
          "Clique em NFCe de novo para tentar transmitir.", vbExclamation, "NFC-e"
End If
End Function

'Transmissão falhou com internet: certificado ou SEFAZ fora do ar (mesmas mensagens do PDV)
Private Sub NFCeOS_SugerirContingencia()
Dim sMsg As String
If Not NFCeOS_TemInternet() Then Exit Sub
sMsg = MsgCertificadoVencido()
If sMsg = "" Then sMsg = MsgCertificadoAusente()
If sMsg <> "" Then
   MsgBox sMsg, vbExclamation, "NFC-e - Certificado digital"
   Exit Sub
End If
NFeResposta = ""
ConsultaStatus 65, True
If NFeResposta = "" Then
   MsgBox "A transmissão da NFC-e falhou mesmo com internet neste computador, e não foi possível consultar o status da SEFAZ." & vbCr & vbCr & _
          "Pode ser um problema de rede local ou da SEFAZ. Se persistir, considere ativar o modo CONTINGÊNCIA (ContigenciaNFCe) no cadastro da empresa.", _
          vbExclamation, "NFC-e - Falha de Comunicação"
ElseIf Left$(NFeResposta, 3) <> "107" Then
   MsgBox "A transmissão da NFC-e falhou mesmo com internet neste computador." & vbCr & vbCr & _
          "Status da SEFAZ: " & NFeResposta & vbCr & vbCr & _
          "Pode ser que a SEFAZ esteja fora do ar. Se confirmar, ative o modo CONTINGÊNCIA (ContigenciaNFCe) no cadastro da empresa.", _
          vbExclamation, "NFC-e - Possível Contingência SEFAZ"
End If
End Sub

'NCM/CFOP/CST ICMS/PIS/COFINS/UN/GTIN de cada item (mesma regra do ValidarItensNFCe do PDV). True = achou erro.
Private Function NFCeOS_ItensComErro(ByVal rItens As ADODB.Recordset) As Boolean
Dim bErro As Boolean
Do While Not rItens.EOF And Not bErro
   If ValidateNull(rItens!CodBarras) <> "SEM GTIN" Then
      If Len(ValidateNull(rItens!CodBarras)) > 13 Or Len(ValidateNull(rItens!CodBarras)) < 8 Then bErro = True
   End If
   If ValidateNull(rItens!CFOP) <> "" Then
      If Len(ValidateNull(rItens!CFOP)) <> 4 Then bErro = True
   End If
   If ValidateNull(rItens!ICMSCST) <> "" Then
      If Len(ValidateNull(rItens!ICMSCST)) <> 3 Then bErro = True
   End If
   If ValidateNull(rItens!PISCST) <> "" Then
      If Len(ValidateNull(rItens!PISCST)) <> 2 Then bErro = True
   End If
   If ValidateNull(rItens!COFINSCST) <> "" Then
      If Len(ValidateNull(rItens!COFINSCST)) <> 2 Then bErro = True
   End If
   If ValidateNull(rItens!CodNcm) <> "" Then
      If Len(ValidateNull(rItens!CodNcm)) <> 8 Then bErro = True
   End If
   If ValidateNull(rItens!UN) <> "" Then
      If Len(ValidateNull(rItens!UN)) > 2 Then bErro = True
   End If
   rItens.MoveNext
Loop
NFCeOS_ItensComErro = bErro
End Function
