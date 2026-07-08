path = 'C:/projeto/OnlineCommerce/Forms/Entrada_Estoque.Frm'
data = open(path, 'rb').read().decode('windows-1252')

# ---- R1: expand rFisc2 to also read CST ----
old1 = (
"Dim sNCMEnt As String, sCFOPEnt As String\r\n"
"Dim sPISCSTEnt As String, sCOFINSCSTEnt As String\r\n"
"Dim rFisc2 As New ADODB.Recordset\r\n"
"RsOpen rFisc2, \"SELECT ISNULL(NCM,'') AS NCM, ISNULL(CAST(CFOP AS VARCHAR(10)),'0') AS CFOP, \" & _\r\n"
"               \"ISNULL(pisCST,'') AS pisCST, ISNULL(cofinsCST,'') AS cofinsCST \" & _\r\n"
"               \"FROM EntradaEstoqueItens WHERE CodigoNota = \" & txtCodNota.Text & \" AND Item = \" & txtItem.Text\r\n"
"If Not rFisc2.EOF Then\r\n"
"   sNCMEnt = CStr(rFisc2!NCM)\r\n"
"   sCFOPEnt = CStr(rFisc2!CFOP)\r\n"
"   sPISCSTEnt = CStr(rFisc2!PISCST)\r\n"
"   sCOFINSCSTEnt = CStr(rFisc2!COFINSCST)\r\n"
"End If\r\n"
)

new1 = (
"Dim sNCMEnt As String, sCFOPEnt As String\r\n"
"Dim sCSTEnt As String, sPISCSTEnt As String, sCOFINSCSTEnt As String\r\n"
"Dim rFisc2 As New ADODB.Recordset\r\n"
"RsOpen rFisc2, \"SELECT ISNULL(NCM,'') AS NCM, ISNULL(CAST(CFOP AS VARCHAR(10)),'0') AS CFOP, \" & _\r\n"
"               \"ISNULL(CST,'') AS CST, ISNULL(PISCST,'') AS PISCST, ISNULL(COFINSCST,'') AS COFINSCST \" & _\r\n"
"               \"FROM EntradaEstoqueItens WHERE CodigoNota = \" & txtCodNota.Text & \" AND Item = \" & txtItem.Text\r\n"
"If Not rFisc2.EOF Then\r\n"
"   sNCMEnt = CStr(rFisc2!NCM)\r\n"
"   sCFOPEnt = CStr(rFisc2!CFOP)\r\n"
"   sCSTEnt = CStr(rFisc2!CST)\r\n"
"   sPISCSTEnt = CStr(rFisc2!PISCST)\r\n"
"   sCOFINSCSTEnt = CStr(rFisc2!COFINSCST)\r\n"
"End If\r\n"
)

# ---- R2: fix CFOP conversion + add CST->CSOSN ----
old2 = (
"Dim sCFOPSaida2 As String\r\n"
"Select Case Mid(sCFOPEnt, 2, 1)\r\n"
"   Case \"1\": sCFOPSaida2 = \"5102\"\r\n"
"   Case \"4\": sCFOPSaida2 = \"5405\"\r\n"
"   Case Else: sCFOPSaida2 = \"\"\r\n"
"End Select\r\n"
"\r\n"
"Dim sICMSCST2 As String\r\n"
"Select Case sCFOPSaida2\r\n"
"   Case \"5102\": sICMSCST2 = \"102\"\r\n"
"   Case \"5405\": sICMSCST2 = \"500\"\r\n"
"   Case Else:   sICMSCST2 = \"\"\r\n"
"End Select\r\n"
)

new2 = (
"'-- CFOP: converte saida->entrada (5xxx->1xxx, 6xxx->2xxx)\r\n"
"Dim sCFOPEntrada2 As String\r\n"
"If Left(sCFOPEnt, 1) = \"5\" Then\r\n"
"   sCFOPEntrada2 = \"1\" & Mid(sCFOPEnt, 2)\r\n"
"ElseIf Left(sCFOPEnt, 1) = \"6\" Then\r\n"
"   sCFOPEntrada2 = \"2\" & Mid(sCFOPEnt, 2)\r\n"
"Else\r\n"
"   sCFOPEntrada2 = sCFOPEnt\r\n"
"End If\r\n"
"\r\n"
"'-- CST->CSOSN quando regime Simples Nacional (CRT=1)\r\n"
"Dim sCSTEntrada2 As String\r\n"
"If sCRT2 = \"1\" Then\r\n"
"   Select Case sCSTEnt\r\n"
"      Case \"000\": sCSTEntrada2 = \"102\"\r\n"
"      Case \"010\": sCSTEntrada2 = \"201\"\r\n"
"      Case \"020\": sCSTEntrada2 = \"101\"\r\n"
"      Case \"030\": sCSTEntrada2 = \"202\"\r\n"
"      Case \"040\": sCSTEntrada2 = \"300\"\r\n"
"      Case \"041\": sCSTEntrada2 = \"400\"\r\n"
"      Case \"060\": sCSTEntrada2 = \"500\"\r\n"
"      Case \"070\": sCSTEntrada2 = \"203\"\r\n"
"      Case \"090\": sCSTEntrada2 = \"900\"\r\n"
"      Case Else:  sCSTEntrada2 = sCSTEnt\r\n"
"   End Select\r\n"
"Else\r\n"
"   sCSTEntrada2 = sCSTEnt\r\n"
"End If\r\n"
"\r\n"
"Dim sICMSCST2 As String\r\n"
"sICMSCST2 = sCSTEntrada2\r\n"
)

# ---- R3: fix sCFOPSaida2 -> sCFOPEntrada2 in sSetFisc ----
old3 = (
"If sCFOPSaida2 <> \"\" And sCFOPSaida2 <> Trim(sCFOPAtual2) Then _\r\n"
"   sSetFisc = sSetFisc & \"CFOP = '\" & sCFOPSaida2 & \"', \"\r\n"
)

new3 = (
"If sCFOPEntrada2 <> \"\" And sCFOPEntrada2 <> Trim(sCFOPAtual2) Then _\r\n"
"   sSetFisc = sSetFisc & \"CFOP = '\" & sCFOPEntrada2 & \"', \"\r\n"
)

# ---- R4: after INSERT execute, add 5 UPDATE statements to copy fiscal fields ----
old4 = (
"'Adiciona o registro\r\n"
"dbData.Execute sSQL\r\n"
"\r\n"
"Preco_Entrada\r\n"
)

new4 = (
"'Adiciona o registro\r\n"
"dbData.Execute sSQL\r\n"
"\r\n"
"'-- Copia campos fiscais de EntradaEstoqueItens -> produtos_entrada_itens\r\n"
"Dim sJoinEnt As String\r\n"
"sJoinEnt = \" FROM produtos_entrada_itens p INNER JOIN EntradaEstoqueItens e\" & _\r\n"
"           \" ON e.CodigoNota = \" & txtCodNota.Text & \" AND e.Item = \" & txtItem.Text & _\r\n"
"           \" WHERE p.codigo = \" & var_COD_ITENS\r\n"
"\r\n"
"'-- Identificacao, qtd comercial, totais item, CFOP e CST convertidos\r\n"
"dbData.Execute \"UPDATE p SET\" & _\r\n"
"   \" p.CodigoNota=e.CodigoNota, p.Item=e.Item, p.InformacoesAdicionaisProduto=e.InformacoesAdicionaisProduto,\" & _\r\n"
"   \" p.NCM=e.NCM, p.Genero=e.Genero, p.Referencia=e.Referencia, p.Adicionada=1,\" & _\r\n"
"   \" p.UnidadeComercial=e.UnidadeComercial, p.QuantidadeComercial=e.QuantidadeComercial,\" & _\r\n"
"   \" p.ValorUnitarioComercializacao=e.ValorUnitarioComercializacao, p.ValorTotalBruto=e.ValorTotalBruto,\" & _\r\n"
"   \" p.EANTrib=e.EANTrib, p.UnidadeTributavel=e.UnidadeTributavel, p.ValorUnitarioTributario=e.ValorUnitarioTributario,\" & _\r\n"
"   \" p.ValorFrete=e.ValorFrete, p.ValorSeguro=e.ValorSeguro, p.TipoDesconto=e.TipoDesconto,\" & _\r\n"
"   \" p.Desconto=e.Desconto, p.ValorDesconto=e.ValorDesconto, p.ValorTributos=e.ValorTributos,\" & _\r\n"
"   \" p.vOutro=e.vOutro, p.vItem=e.vItem, p.indTot=e.indTot,\" & _\r\n"
"   \" p.CFOP=\" & sCFOPEntrada2 & \", p.CST='\" & sCSTEntrada2 & \"'\" & _\r\n"
"   sJoinEnt\r\n"
"\r\n"
"'-- ICMS detalhado\r\n"
"dbData.Execute \"UPDATE p SET\" & _\r\n"
"   \" p.modBC=e.modBC, p.pRedBC=e.pRedBC, p.vBC=e.vBC, p.pICMS=e.pICMS, p.vICMS=e.vICMS,\" & _\r\n"
"   \" p.modBCST=e.modBCST, p.pMVAST=e.pMVAST, p.pRedBCST=e.pRedBCST, p.vBCST=e.vBCST,\" & _\r\n"
"   \" p.pICMSST=e.pICMSST, p.vICMSST=e.vICMSST,\" & _\r\n"
"   \" p.vICMSOp=e.vICMSOp, p.pDif=e.pDif, p.vICMSDeson=e.vICMSDeson, p.motDesICMS=e.motDesICMS,\" & _\r\n"
"   \" p.vBCFCPST=e.vBCFCPST, p.pFCPST=e.pFCPST, p.vFCPST=e.vFCPST,\" & _\r\n"
"   \" p.vBCSTRet=e.vBCSTRet, p.pST=e.pST, p.vICMSSubstituto=e.vICMSSubstituto, p.vICMSSTRet=e.vICMSSTRet,\" & _\r\n"
"   \" p.vBCFCPSTRet=e.vBCFCPSTRet, p.pFCPSTRet=e.pFCPSTRet, p.vFCPSTRet=e.vFCPSTRet,\" & _\r\n"
"   \" p.CEST=e.CEST, p.indEscala=e.indEscala, p.nFCI=e.nFCI, p.TipoProduto=e.TipoProduto\" & _\r\n"
"   sJoinEnt\r\n"
"\r\n"
"'-- IPI, II, PIS, COFINS, ISS\r\n"
"dbData.Execute \"UPDATE p SET\" & _\r\n"
"   \" p.IPIclEnq=e.IPIclEnq, p.IPICNPJProd=e.IPICNPJProd, p.IPIcSelo=e.IPIcSelo,\" & _\r\n"
"   \" p.IPIqSelo=e.IPIqSelo, p.IPIcEnq=e.IPIcEnq, p.IPICST=e.IPICST,\" & _\r\n"
"   \" p.IPIvIPI=e.IPIvIPI, p.IPIvBC=e.IPIvBC, p.IPIpIPI=e.IPIpIPI,\" & _\r\n"
"   \" p.IPIvUnid=e.IPIvUnid, p.IPIqUnid=e.IPIqUnid,\" & _\r\n"
"   \" p.IIvBC=e.IIvBC, p.IIvDespAdu=e.IIvDespAdu, p.IIvII=e.IIvII, p.IIvIOF=e.IIvIOF,\" & _\r\n"
"   \" p.PISCST=e.PISCST, p.PISvBC=e.PISvBC, p.PISpPIS=e.PISpPIS,\" & _\r\n"
"   \" p.PISvPIS=e.PISvPIS, p.PISqBCProd=e.PISqBCProd, p.PISvAliqProd=e.PISvAliqProd,\" & _\r\n"
"   \" p.COFINSCST=e.COFINSCST, p.COFINSvBC=e.COFINSvBC, p.COFINSpCOFINS=e.COFINSpCOFINS,\" & _\r\n"
"   \" p.COFINSvCOFINS=e.COFINSvCOFINS, p.COFINSqBCProd=e.COFINSqBCProd, p.COFINSvAliqProd=e.COFINSvAliqProd,\" & _\r\n"
"   \" p.ISSvBC=e.ISSvBC, p.ISSvAliq=e.ISSvAliq, p.ISSvISSQN=e.ISSvISSQN,\" & _\r\n"
"   \" p.ISScMunFG=e.ISScMunFG, p.ISScListServ=e.ISScListServ\" & _\r\n"
"   sJoinEnt\r\n"
"\r\n"
"'-- IBS/CBS\r\n"
"dbData.Execute \"UPDATE p SET\" & _\r\n"
"   \" p.IBSCBSCST=e.IBSCBSCST, p.IBSCBScClassTrib=e.IBSCBScClassTrib, p.IBSCBSvBC=e.IBSCBSvBC,\" & _\r\n"
"   \" p.IBSUFpAliq=e.IBSUFpAliq, p.IBSUFpRedAliq=e.IBSUFpRedAliq, p.IBSUFpAliqEfet=e.IBSUFpAliqEfet, p.IBSUFvIBS=e.IBSUFvIBS,\" & _\r\n"
"   \" p.IBSMunpAliq=e.IBSMunpAliq, p.IBSMunpRedAliq=e.IBSMunpRedAliq, p.IBSMunpAliqEfet=e.IBSMunpAliqEfet,\" & _\r\n"
"   \" p.IBSMunvIBS=e.IBSMunvIBS, p.IBSvIBS=e.IBSvIBS,\" & _\r\n"
"   \" p.CBSpAliq=e.CBSpAliq, p.CBSpRedAliq=e.CBSpRedAliq, p.CBSpAliqEfet=e.CBSpAliqEfet, p.CBSvCBS=e.CBSvCBS,\" & _\r\n"
"   \" p.ISCST=e.ISCST, p.IScClassTrib=e.IScClassTrib, p.ISvBCIS=e.ISvBCIS, p.ISpIS=e.ISpIS, p.ISvIS=e.ISvIS,\" & _\r\n"
"   \" p.IBSUFpDif=e.IBSUFpDif, p.IBSUFvDif=e.IBSUFvDif, p.IBSUFvDevTrib=e.IBSUFvDevTrib,\" & _\r\n"
"   \" p.IBSMunpDif=e.IBSMunpDif, p.IBSMunvDif=e.IBSMunvDif, p.IBSMunvDevTrib=e.IBSMunvDevTrib,\" & _\r\n"
"   \" p.CBSpDif=e.CBSpDif, p.CBSvDif=e.CBSvDif, p.CBSvDevTrib=e.CBSvDevTrib\" & _\r\n"
"   sJoinEnt\r\n"
"\r\n"
"'-- IBS/CBS Mono\r\n"
"dbData.Execute \"UPDATE p SET\" & _\r\n"
"   \" p.IBSCBSMonoqBCMono=e.IBSCBSMonoqBCMono, p.IBSCBSMonoadRemIBS=e.IBSCBSMonoadRemIBS,\" & _\r\n"
"   \" p.IBSCBSMonoadRemCBS=e.IBSCBSMonoadRemCBS, p.IBSCBSMonovIBSMono=e.IBSCBSMonovIBSMono,\" & _\r\n"
"   \" p.IBSCBSMonovCBSMono=e.IBSCBSMonovCBSMono,\" & _\r\n"
"   \" p.IBSCBSMonoqBCMonoReten=e.IBSCBSMonoqBCMonoReten, p.IBSCBSMonoadRemIBSReten=e.IBSCBSMonoadRemIBSReten,\" & _\r\n"
"   \" p.IBSCBSMonovIBSMonoReten=e.IBSCBSMonovIBSMonoReten, p.IBSCBSMonoadRemCBSReten=e.IBSCBSMonoadRemCBSReten,\" & _\r\n"
"   \" p.IBSCBSMonovCBSMonoReten=e.IBSCBSMonovCBSMonoReten,\" & _\r\n"
"   \" p.IBSCBSMonoqBCMonoRet=e.IBSCBSMonoqBCMonoRet, p.IBSCBSMonoadRemIBSRet=e.IBSCBSMonoadRemIBSRet,\" & _\r\n"
"   \" p.IBSCBSMonovIBSMonoRet=e.IBSCBSMonovIBSMonoRet, p.IBSCBSMonoadRemCBSRet=e.IBSCBSMonoadRemCBSRet,\" & _\r\n"
"   \" p.IBSCBSMonovCBSMonoRet=e.IBSCBSMonovCBSMonoRet,\" & _\r\n"
"   \" p.IBSCBSMonovTotIBSMonoItem=e.IBSCBSMonovTotIBSMonoItem, p.IBSCBSMonovTotCBSMonoItem=e.IBSCBSMonovTotCBSMonoItem\" & _\r\n"
"   sJoinEnt\r\n"
"\r\n"
"Preco_Entrada\r\n"
)

print('r1 found:', data.count(old1))
print('r2 found:', data.count(old2))
print('r3 found:', data.count(old3))
print('r4 found:', data.count(old4))

data2 = data
data2 = data2.replace(old1, new1, 1)
data2 = data2.replace(old2, new2, 1)
data2 = data2.replace(old3, new3, 1)
data2 = data2.replace(old4, new4, 1)

print('changed:', data2 != data)
raw = data2.encode('windows-1252')
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(path, 'wb').write(raw)
print('ok')
