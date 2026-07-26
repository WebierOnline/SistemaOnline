path = 'C:/projeto/OnlineCommerce/Forms/Entrada_Estoque.Frm'
data = open(path, 'rb').read().decode('windows-1252')

# Mover leitura de QUANT para ANTES do BEGIN TRANSACTION (ADO nao pode ficar dentro de transacao DAO)
# e garantir que o DELETE ocorre apenas APOS a leitura

old = (
    '   Dim bTransRem As Boolean\r\n'
    '   bTransRem = False\r\n'
    '   On Error GoTo ErrRemover2\r\n'
    '\r\n'
    '   dbData.Execute "BEGIN TRANSACTION"\r\n'
    '   bTransRem = True\r\n'
    '\r\n'
    '   dbData.Execute "UPDATE EntradaEstoqueItens SET adicionada = 0 WHERE (codigonota = " & txtCodNota.Text & ") and item = " & GridEntradaItens.TextMatrix(GridEntradaItens.RowSel, 17) & ";"\r\n'
    '   dbData.Execute "DELETE FROM produtos_entrada_itens WHERE (codigo = " & GridEntradaItens.TextMatrix(GridEntradaItens.Row, 1) & ") AND (codigo_entrada = " & txtCodEntrada.Text & ");"\r\n'
    '   Dim dQtdRem As Double\r\n'
    '   dQtdRem = CDbl(SQLExecutaRetorno("SELECT ISNULL(QUANT,0) r FROM produtos_entrada_itens WHERE codigo = " & GridEntradaItens.TextMatrix(GridEntradaItens.RowSel, 1), "r", "0"))\r\n'
    '   dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque - " & Replace(dQtdRem, ",", ".") & " WHERE (codigo = " & Val(GridEntradaItens.TextMatrix(GridEntradaItens.RowSel, 4)) & ");"\r\n'
)
new = (
    '   Dim bTransRem As Boolean\r\n'
    '   bTransRem = False\r\n'
    '   On Error GoTo ErrRemover2\r\n'
    '\r\n'
    '   Dim dQtdRem As Double\r\n'
    '   dQtdRem = CDbl(SQLExecutaRetorno("SELECT ISNULL(QUANT,0) r FROM produtos_entrada_itens WHERE codigo = " & GridEntradaItens.TextMatrix(GridEntradaItens.RowSel, 1), "r", "0"))\r\n'
    '\r\n'
    '   dbData.Execute "BEGIN TRANSACTION"\r\n'
    '   bTransRem = True\r\n'
    '\r\n'
    '   dbData.Execute "UPDATE EntradaEstoqueItens SET adicionada = 0 WHERE (codigonota = " & txtCodNota.Text & ") and item = " & GridEntradaItens.TextMatrix(GridEntradaItens.RowSel, 17) & ";"\r\n'
    '   dbData.Execute "DELETE FROM produtos_entrada_itens WHERE (codigo = " & GridEntradaItens.TextMatrix(GridEntradaItens.RowSel, 1) & ") AND (codigo_entrada = " & txtCodEntrada.Text & ");"\r\n'
    '   dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque - " & Replace(dQtdRem, ",", ".") & " WHERE (codigo = " & Val(GridEntradaItens.TextMatrix(GridEntradaItens.RowSel, 4)) & ");"\r\n'
)

print('found:', data.count(old))
data2 = data.replace(old, new, 1)
print('changed:', data2 != data)
raw = data2.encode('windows-1252')
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(path, 'wb').write(raw)
print('ok')
