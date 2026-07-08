# -*- coding: utf-8 -*-
"""
Patch v7: cmdExibirPedidos — corrige passagem de parametros ao loadPedidos
  1. sBase (POR SERVICOS SQL) — substitui '0 AS var_CodOS' por OS.COD_PEDIDO
  2. FormatarGrid_Servicos col 10 — grava COD_PEDIDO em vez de "0"
  3. cmdExibirPedidos — branch por cboTipo, passa parametros corretos
"""
import sys

FILE = r'C:\Projeto\OnlineCommerce\Forms\Vendas_Consulta_PorProdutos.frm'

with open(FILE, 'rb') as f:
    raw = f.read()

raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n')
text = raw.decode('windows-1252')

changes = []

# ------------------------------------------------------------------
# 1. sBase em cmdLocalizar_Click: 0 AS var_CodOS -> OS.COD_PEDIDO
# ------------------------------------------------------------------
old1 = (
    '   sBase = "SELECT s.codigo, OS.COD_OS AS varCodPed, s.data AS varData, s.descricao AS varNome, " & _\n'
    '           "s.preco AS varValor, s.quantidade AS varQuant, s.subtotal AS varSubtotal, " & _\n'
    '           "s.desconto AS varDesc, s.total AS varTotal, 0 AS var_CodOS " & _\n'
    '           "FROM OS_Servicos_Auto s INNER JOIN OS ON s.cod_os = OS.COD_OS "\n'
)
new1 = (
    '   sBase = "SELECT s.codigo, OS.COD_OS AS varCodPed, s.data AS varData, s.descricao AS varNome, " & _\n'
    '           "s.preco AS varValor, s.quantidade AS varQuant, s.subtotal AS varSubtotal, " & _\n'
    '           "s.desconto AS varDesc, s.total AS varTotal, ISNULL(OS.COD_PEDIDO, 0) AS var_CodOS " & _\n'
    '           "FROM OS_Servicos_Auto s INNER JOIN OS ON s.cod_os = OS.COD_OS "\n'
)
changes.append((old1, new1, '1 - sBase inclui OS.COD_PEDIDO como var_CodOS'))

# ------------------------------------------------------------------
# 2. FormatarGrid_Servicos col 10: "0" -> rTabela("var_CodOS")
# ------------------------------------------------------------------
old2 = (
    '            .TextMatrix(.rows - 1, 9) = Format(rTabela("varTotal"), ocMONEY)\n'
    '            .TextMatrix(.rows - 1, 10) = "0"\n'
)
new2 = (
    '            .TextMatrix(.rows - 1, 9) = Format(rTabela("varTotal"), ocMONEY)\n'
    '            .TextMatrix(.rows - 1, 10) = rTabela("var_CodOS")\n'
)
changes.append((old2, new2, '2 - FormatarGrid_Servicos col10 grava COD_PEDIDO'))

# ------------------------------------------------------------------
# 3. cmdExibirPedidos_Click — branch por cboTipo
# ------------------------------------------------------------------
old3 = (
    "If Not IsNumeric(Grid.TextMatrix(Grid.Row, 1)) = True Then Exit Sub\n"
    "If Grid.TextMatrix(Grid.Row, 1) = \"\" Or Grid.TextMatrix(Grid.Row, 10) = \"\" Then Exit Sub\n"
    "\n"
    "If Grid.TextMatrix(Grid.Row, 10) <> \"0\" Then\n"
    "   Parcelas_Consulta_Produtos.loadPedidos Grid.TextMatrix(Grid.Row, 1), \"OS\"\n"
    "Else\n"
    "   Parcelas_Consulta_Produtos.loadPedidos Grid.TextMatrix(Grid.Row, 1), Grid.TextMatrix(Grid.Row, 7)\n"
    "End If\n"
    "\n"
    "\n"
    "Parcelas_Consulta_Produtos.Show 1\n"
    "End Sub\n"
)
new3 = (
    "If Grid.TextMatrix(Grid.Row, 1) = \"\" Then Exit Sub\n"
    "\n"
    "If cboTipo.Text = \"POR SERVIÇOS\" Then\n"
    "   ' col1 = COD_OS, col10 = COD_PEDIDO (gravado pelo FormatarGrid_Servicos)\n"
    "   If Not IsNumeric(Grid.TextMatrix(Grid.Row, 10)) Then Exit Sub\n"
    "   If CLng(Grid.TextMatrix(Grid.Row, 10)) = 0 Then Exit Sub\n"
    "   Parcelas_Consulta_Produtos.loadPedidos CLng(Grid.TextMatrix(Grid.Row, 10)), \"OS\"\n"
    "Else\n"
    "   ' POR PRODUTOS: col1 = cod_pedido, col10 = COD_OS (0 = sem OS)\n"
    "   If Not IsNumeric(Grid.TextMatrix(Grid.Row, 1)) Then Exit Sub\n"
    "   If Grid.TextMatrix(Grid.Row, 10) <> \"0\" Then\n"
    "      Parcelas_Consulta_Produtos.loadPedidos CLng(Grid.TextMatrix(Grid.Row, 1)), \"OS\"\n"
    "   Else\n"
    "      Parcelas_Consulta_Produtos.loadPedidos CLng(Grid.TextMatrix(Grid.Row, 1)), \"VENDA\"\n"
    "   End If\n"
    "End If\n"
    "\n"
    "Parcelas_Consulta_Produtos.Show 1\n"
    "End Sub\n"
)
changes.append((old3, new3, '3 - cmdExibirPedidos branch por cboTipo'))

# ------------------------------------------------------------------
# Aplicar e verificar
# ------------------------------------------------------------------
for old, new, label in changes:
    count = text.count(old)
    if count != 1:
        print(f'ERRO [{label}]: encontrado {count} ocorrencias (esperado 1)')
        sys.exit(1)
    text = text.replace(old, new)
    print(f'OK: {label}')

text = text.replace('\r\n', '\n').replace('\r', '\n')
out = text.encode('windows-1252')
out = out.replace(b'\n', b'\r\n')

with open(FILE, 'wb') as f:
    f.write(out)

print('\nArquivo gravado com sucesso.')
