# -*- coding: utf-8 -*-
# Reescreve cmdDuplicar_Click para copiar nota idêntica:
# - Usa Load_Controls para carregar TODOS os campos da NF original
# - Sobrescreve apenas: número (próximo), código (próximo), hora (agora)
# - DataEmissao/DataSaida já são forçadas para hoje em LerDadosInserir
# - Copia itens com TODAS as colunas via INFORMATION_SCHEMA (exclui PK CodigoItem)
PATH = r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm'
data = open(PATH, 'rb').read()
text = data.decode('windows-1252')

i = text.find('Private Sub cmdDuplicar_Click()')
end = text.find('\r\nEnd Sub\r\n', i) + len('\r\nEnd Sub\r\n')
OLD = text[i:end]

NEW = (
    'Private Sub cmdDuplicar_Click()\r\n'
    'If GridNotas.Row = 0 Then MsgBox "Selecione uma nota fiscal na lista!", vbInformation, "Aviso do Sistema": Exit Sub\r\n'
    '\r\n'
    'Dim varCodNotaOrigem As Long\r\n'
    'varCodNotaOrigem = CLng(GridNotas.TextMatrix(GridNotas.Row, 1))\r\n'
    '\r\n'
    "'autonumera\xe7\xe3o\r\n"
    'Dim tbNota As ADODB.Recordset\r\n'
    'Set tbNota = dbData.OpenRecordset("SELECT ISNULL(MAX(numeronota), 0) AS Maior_nota FROM NotaFiscal")\r\n'
    '\r\n'
    'LimparObjetosNota\r\n'
    'LimparObjetosDestinatario\r\n'
    'LimparObjetosProduto\r\n'
    'LimparObjetosNotaTotais\r\n'
    'LimparObjestosNotaOutros\r\n'
    'LimparGridItensNota\r\n'
    '\r\n'
    "'calcular novo n\xfamero e c\xf3digo da nota\r\n"
    'RsOpen TbNotas, "SELECT * FROM NotaFiscal"\r\n'
    'Dim novoNumNota As Long, novoCodNota As Long\r\n'
    'If TbNotas.EOF And TbNotas.BOF Then\r\n'
    '    novoNumNota = tbNota("Maior_nota") + 1\r\n'
    '    novoCodNota = 1\r\n'
    'Else\r\n'
    '    TbNotas.MoveLast\r\n'
    '    novoNumNota = tbNota("Maior_nota") + 1\r\n'
    '    novoCodNota = TbNotas("CodigoNota") + 1\r\n'
    'End If\r\n'
    '\r\n'
    "'carregar todos os dados da nota original no form\r\n"
    'vTipoEdicaoNFe = "Edicao"\r\n'
    'RsOpen TbNotas, "SELECT * FROM NotaFiscal WHERE CodigoNota = " & varCodNotaOrigem\r\n'
    'Load_Controls\r\n'
    '\r\n'
    "'sobrescrever apenas n\xfamero, c\xf3digo e hora (datas s\xe3o for\xe7adas para hoje em LerDadosInserir)\r\n"
    'txtNumNota.Text = novoNumNota\r\n'
    'txtCodNota.Text = novoCodNota\r\n'
    'mskHora = Format(Time(), "HH:MM:ss")\r\n'
    '\r\n'
    "'abrir recordset vazio para AddNew do cabe\xe7alho\r\n"
    'RsOpen TbNotas, "SELECT * FROM NotaFiscal"\r\n'
    'TbNotas.AddNew\r\n'
    '\r\n'
    "'copiar itens com TODAS as colunas (exceto PK CodigoItem) via INFORMATION_SCHEMA\r\n"
    'Dim rCols As ADODB.Recordset\r\n'
    'RsOpen rCols, "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS " & _\r\n'
    '              "WHERE TABLE_NAME = \'NotaFiscalItens\' AND COLUMN_NAME NOT IN (\'CodigoItem\', \'CodigoNota\') " & _\r\n'
    '              "ORDER BY ORDINAL_POSITION"\r\n'
    'Dim colList As String\r\n'
    'colList = ""\r\n'
    'Do While Not rCols.EOF\r\n'
    '    colList = colList & "[" & rCols!COLUMN_NAME & "], "\r\n'
    '    rCols.MoveNext\r\n'
    'Loop\r\n'
    'If Len(colList) > 2 Then colList = Left(colList, Len(colList) - 2)\r\n'
    '\r\n'
    'sSQL = "INSERT INTO NotaFiscalItens (CodigoNota, " & colList & ") " & _\r\n'
    '       "SELECT " & novoCodNota & ", " & colList & " " & _\r\n'
    '       "FROM NotaFiscalItens WHERE CodigoNota = " & varCodNotaOrigem\r\n'
    'dbData.Execute sSQL\r\n'
    '\r\n'
    'Exibir_Itens\r\n'
    'AtualizarTotaisNota\r\n'
    '\r\n'
    'LerDadosInserir\r\n'
    'TbNotas.Update\r\n'
    '\r\n'
    'cmdNovo.Enabled = False\r\n'
    'cmdSalvar.Enabled = True\r\n'
    'cmdCancelar.Enabled = True\r\n'
    'frmNota.Enabled = True\r\n'
    'frmDestinatario.Enabled = True\r\n'
    'Tab_Totais.Enabled = True\r\n'
    'Tab_Produtos.Enabled = True\r\n'
    'frmItens.Enabled = True\r\n'
    'cmdDuplicar.Enabled = False\r\n'
    '\r\n'
    'Frm_NF.Tab = 0\r\n'
    'Tab_Produtos.Tab = 0\r\n'
    'Tab_Totais.Tab = 0\r\n'
    '\r\n'
    'Call cmdRecalcular_Click\r\n'
    '\r\n'
    'End Sub\r\n'
)

c = text.count(OLD)
print(f'OLD count={c}')
if c == 1:
    text = text.replace(OLD, NEW)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(PATH, 'wb').write(out)
    print('OK — cmdDuplicar_Click reescrito')
else:
    print('ABORTADO')
