# -*- coding: utf-8 -*-
"""
Fix cmdExibirParcelas_Click em Vendas_Consulta_PorProdutos.frm:
  POR SERVICOS: col 1 = COD_OS; col 11 = COD_PEDIDO da venda vinculada.
  loadInformacoes busca parcelas por cod_pedido, entao usar col 11.
  Se col 11 = 0 (sem venda), nao ha parcelas.
"""
import sys

FILE = r'C:\Projeto\OnlineCommerce\Forms\Vendas_Consulta_PorProdutos.frm'

with open(FILE, 'rb') as f:
    raw = f.read()
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n')
text = raw.decode('windows-1252')

old = (
    'Sub cmdExibirParcelas_Click()\n'
    'If Grid.Col = 0 Then Exit Sub\n'
    '   If IsNumeric(Grid.TextMatrix(Grid.Row, 1)) = True Then\n'
    '         Vendas_Consulta_Geral_Parcelas.loadInformacoes (Grid.TextMatrix(Grid.Row, 1))\n'
    '         Vendas_Consulta_Geral_Parcelas.Show 1\n'
    '   End If\n'
    'End Sub'
)
new = (
    'Sub cmdExibirParcelas_Click()\n'
    'If Grid.Col = 0 Then Exit Sub\n'
    '   Dim lPedido As Long\n'
    '   If cboTipo.Text = "POR SERVI\xc7OS" Then\n'
    '      If Not IsNumeric(Grid.TextMatrix(Grid.Row, 11)) Then Exit Sub\n'
    '      lPedido = CLng(Grid.TextMatrix(Grid.Row, 11))\n'
    '      If lPedido = 0 Then Exit Sub\n'
    '   ElseIf IsNumeric(Grid.TextMatrix(Grid.Row, 1)) Then\n'
    '      lPedido = CLng(Grid.TextMatrix(Grid.Row, 1))\n'
    '   Else\n'
    '      Exit Sub\n'
    '   End If\n'
    '   Vendas_Consulta_Geral_Parcelas.loadInformacoes lPedido\n'
    '   Vendas_Consulta_Geral_Parcelas.Show 1\n'
    'End Sub'
)

count = text.count(old)
if count != 1:
    print(f'ERRO: {count} ocorrencias (esperado 1)')
    sys.exit(1)

text = text.replace(old, new)
print('OK (1x): cmdExibirParcelas_Click — usa col 11 (COD_PEDIDO) para POR SERVICOS')

out = text.encode('windows-1252')
out = out.replace(b'\n', b'\r\n')
with open(FILE, 'wb') as f:
    f.write(out)
print('\nArquivo gravado com sucesso.')
