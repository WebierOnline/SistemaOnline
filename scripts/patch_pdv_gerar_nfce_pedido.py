# -*- coding: utf-8 -*-
"""PDV.frm - nova funcao Public GerarEImprimirNFCeParaPedido(pCodPedido), chamada pelo
cmdNFCe do Estonar. Extrai/simplifica o bloco de emissao de NFCe que ja existe (repetido)
dentro de cmdFinalizar_Click - mesmos passos fiscais (EXEC NFCeIncluir, TbNFCe_Faturas,
ValidarItensNFCe, TransmitirNFCe, impressao DANFCe), mas:

- 100% independente do estado da TELA do PDV: nunca le nem escreve txtCodPedido/
  txtCodCliente/cboTipoPgto - o operador pode estar com uma venda diferente em andamento
  na MESMA maquina quando alguem aciona o cmdNFCe pelo Estonar (Estonar e PDV rodam no
  mesmo processo/instancia). So usa parametro + variaveis locais + as globais de config
  (var_ImpNFCe, vConfImprimeNFCeLocal, vNFCeConfPrazo, NFCeContingencia) que ja eram usadas
  do mesmo jeito no fluxo original.
- SEM pop-up de confirmacao (pedido explicito do usuario, "gera e imprime direto"): nao
  chama DecidirGeracaoNFCe (que pergunta "Impressora Pronta?") nem ValidarCPFCNPJCliente
  (que pode abrir InputBox pedindo CPF do cliente) - usa as configs (vConfImprimeNFCeLocal/
  vNFCeConfPrazo) direto pra decidir, e usa o CPF que ja estiver salvo no cadastro do
  cliente (sem pedir um novo). Mantém o aviso de "sem internet -> contingencia"
  (VerificarConexaoParaNFCe) e o de falha de transmissao (SugerirContingenciaSeSefazCaida) -
  sao avisos informativos de erro/estado, nao perguntas bloqueando o fluxo.
- Reusa NFCeJaExisteParaPedido/ValidarItensNFCe/SugerirContingenciaSeSefazCaida (Private,
  ja em PDV.frm - por isso a funcao nova mora no MESMO form, nao num modulo separado) e
  TransmitirNFCe/ConfiguraDLLNFeNFCe/PreencherReformaTributariaNFCe (Public, ja em
  Compartilhado/Modulos/modNFe.bas).

.frm cp1252 -> edicao binaria + normaliza CRLF.
"""
import sys

p = r"C:\projeto\PDV\Forms\PDV.frm"
d = open(p, "rb").read().replace(b"\r\n", b"\n").replace(b"\r", b"\n")

o = b'"Marcar Todos"\nEnd Sub\n\nPrivate Function NFCeJaExisteParaPedido'
n = (
b'"Marcar Todos"\nEnd Sub\n\n'
b"Public Function GerarEImprimirNFCeParaPedido(ByVal pCodPedido As Long) As Boolean\n"
b"'gera, transmite e imprime a NFCe de um pedido ja fechado, sem pedir confirmacao (chamado\n"
b"'pelo Estonar - cmdNFCe). Mesmo procedimento fiscal de cmdFinalizar_Click, mas independente\n"
b"'do estado da tela do PDV (nao mexe em txtCodPedido/txtCodCliente/cboTipoPgto).\n"
b"On Error GoTo ErrHandlerGerarNFCe\n"
b"\n"
b"Dim sSQL As String\n"
b"Dim r As ADODB.Recordset, rNFCe As ADODB.Recordset, rNFCeItens As ADODB.Recordset\n"
b"Dim bTrans As Boolean\n"
b"Dim iRetorno As Boolean\n"
b"Dim EncontroErroNFCe As Boolean\n"
b"Dim vTipoPgtoPedido As String\n"
b"Dim NFCeContingencia As Boolean\n"
b"\n"
b"GerarEImprimirNFCeParaPedido = False\n"
b"\n"
b"If NFCeJaExisteParaPedido(CStr(pCodPedido)) Then\n"
b'    MsgBox "NFCe para esse pedido j\xe1 foi criada.", vbInformation, "Aviso do Sistema"\n'
b"    Exit Function\n"
b"End If\n"
b"\n"
b'sSQL = "SELECT tipo_pagamento FROM pedidos WHERE (cod_pedido = " & pCodPedido & ")"\n'
b"Set r = dbData.OpenRecordset(sSQL)\n"
b"If r.EOF Then\n"
b'    MsgBox "Pedido n\xe3o encontrado.", vbExclamation, "Aviso do Sistema"\n'
b"    Exit Function\n"
b"End If\n"
b'vTipoPgtoPedido = ValidateNull(r("tipo_pagamento"))\n'
b"If r.State <> 0 Then r.Close\n"
b"Set r = Nothing\n"
b"\n"
b'If vConfImprimeNFCeLocal <> "SIM" Then\n'
b'    MsgBox "Essa m\xe1quina n\xe3o est\xe1 configurada para emitir NFCe localmente.", vbExclamation, "Aviso do Sistema"\n'
b"    Exit Function\n"
b"End If\n"
b"\n"
b'If (vTipoPgtoPedido = "\xe0 Prazo" Or vTipoPgtoPedido = "\xe0 PRAZO") And vNFCeConfPrazo <> "SIM" Then\n'
b'    MsgBox "A emiss\xe3o de NFCe para vendas \xe0 Prazo est\xe1 desativada nas configura\xe7\xf5es.", vbExclamation, "Aviso do Sistema"\n'
b"    Exit Function\n"
b"End If\n"
b"\n"
b"Dim oIni As Ini\n"
b"Set oIni = New Ini\n"
b'oIni.Arquivo = appPathApp & "config.ini"\n'
b'var_ImpNFCe = oIni.LerTexto("IMPRESSORA_NFCE", "impressora")\n'
b"Set oIni = Nothing\n"
b"\n"
b"Dim Prt As Printer\n"
b"For Each Prt In Printers\n"
b"   If Prt.DeviceName = var_ImpNFCe Then\n"
b"      Set Printer = Prt\n"
b"      Exit For\n"
b"   End If\n"
b"Next\n"
b"\n"
b'sSQL = "SELECT TOP 1 NFCeOffline FROM empresa ORDER BY fantasia;"\n'
b"Set r = dbData.OpenRecordset(sSQL)\n"
b"NFCeContingencia = r!NFCeOffline\n"
b"If r.State <> 0 Then r.Close\n"
b"Set r = Nothing\n"
b"\n"
b"VerificarConexaoParaNFCe NFCeContingencia\n"
b"\n"
b'dbData.Execute "BEGIN TRANSACTION"\n'
b"bTrans = True\n"
b"\n"
b'sSQL = "EXEC NFCeIncluir " & pCodPedido\n'
b"dbData.Execute sSQL\n"
b"\n"
b'sSQL = "SELECT IdNFProd FROM TbNFCe WHERE Num_OS_VD_Origem = " & pCodPedido\n'
b"Set rNFCe = dbData.OpenRecordset(sSQL)\n"
b"\n"
b"If rNFCe.RecordCount = 0 Then\n"
b'    dbData.Execute "ROLLBACK TRANSACTION"\n'
b"    bTrans = False\n"
b'    MsgBox "N\xe3o foi poss\xedvel gerar a NFCe para esse pedido.", vbExclamation, "Aviso do Sistema"\n'
b"    Exit Function\n"
b"End If\n"
b"\n"
b"PreencherReformaTributariaNFCe CLng(rNFCe!IdNFProd), dbData.ActiveConnection\n"
b"\n"
b'sSQL = "INSERT INTO [TbNFCe_Faturas] ([IdNFProd],[IDParcela],[TipoPgto],[Vencimento],[Valor],[IdBandeira],[CartaoNumeroAutorizacao]) " & _\n'
b'       "SELECT " & rNFCe!IdNFProd & ", NUMERO, dbo.NFCeFormaPagto(FORMA_PGTO, TIPO_CARTAO), DATA, VALOR, \'01\', \'\' " & _\n'
b'       "FROM [parcelas] WHERE COD_PEDIDO = " & pCodPedido\n'
b"dbData.Execute sSQL\n"
b"\n"
b'sSQL = "SELECT IdNFProd, IdNFProd_Item, IDProduto, CodBarras, DescricaoProduto, CodNcm, CFOP, Bc_Icms, ICMSCST, IPICST, COFINSCST, PISCST, UN " & _\n'
b'       "FROM TbNFCe_Itens WHERE (IdNFProd = " & rNFCe!IdNFProd & ");"\n'
b"Set rNFCeItens = dbData.OpenRecordset(sSQL)\n"
b"\n"
b"EncontroErroNFCe = ValidarItensNFCe(rNFCeItens)\n"
b"\n"
b'dbData.Execute "COMMIT TRANSACTION"\n'
b"bTrans = False\n"
b"\n"
b"If EncontroErroNFCe Then\n"
b'    MsgBox "A NFCe foi criada, mas algum item tem dado fiscal incompleto (NCM/CFOP/CST). Corrija o cadastro do produto antes de transmitir.", vbExclamation, "Aviso do Sistema"\n'
b"    Exit Function\n"
b"End If\n"
b"\n"
b"DoEvents\n"
b'iRetorno = TransmitirNFCe(rNFCe!IdNFProd, "1", Not NFCeContingencia, "65", True)\n'
b"If Not iRetorno And Not NFCeContingencia Then SugerirContingenciaSeSefazCaida\n"
b"\n"
b"If iRetorno Then\n"
b"    Dim sistNFe As snfe.Util\n"
b"    Set sistNFe = New snfe.Util\n"
b'    ConfiguraDLLNFeNFCe 65, "1", sistNFe\n'
b"    If Not NFCeContingencia Then\n"
b'       Call sistNFe.DANFCeImprimir(xCaminhoXML, True, var_ImpNFCe, True, xCaminhoPDF, 0, False, False, "")\n'
b"    Else\n"
b'       Call sistNFe.DANFCeOFFImprimir(xCaminhoXML, True, var_ImpNFCe, True, xCaminhoPDF, 0, False, False, "")\n'
b"    End If\n"
b"    GerarEImprimirNFCeParaPedido = True\n"
b"Else\n"
b'    MsgBox "A NFCe foi criada mas n\xe3o foi autorizada pela SEFAZ.", vbExclamation, "Aviso do Sistema"\n'
b"End If\n"
b"\n"
b"Exit Function\n"
b"\n"
b"ErrHandlerGerarNFCe:\n"
b"If bTrans Then\n"
b'    dbData.Execute "ROLLBACK TRANSACTION"\n'
b"End If\n"
b'MsgBox "Erro ao gerar a NFCe: " & Err.Description, vbCritical, "Erro"\n'
b"End Function\n"
b"\n"
b"Private Function NFCeJaExisteParaPedido"
)

if n in d and o not in d:
    print("[ja] GerarEImprimirNFCeParaPedido")
else:
    c = d.count(o)
    if c != 1:
        print("[ABORTA] -- %d ocorrencias de old (esperava 1)" % c)
        sys.exit(1)
    d = d.replace(o, n, 1)
    print("[ok] GerarEImprimirNFCeParaPedido inserida")

open(p, "wb").write(d.replace(b"\n", b"\r\n"))
print("gravado " + p)
