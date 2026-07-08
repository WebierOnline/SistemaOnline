# -*- coding: utf-8 -*-
"""
Patch v5: PRODUTO/MENSAL e PRODUTO/PERIODO
  1. cboCriterioPrinc_LostFocus — adiciona casos para mostrar cboDescricao+datas
  2. cmdLocalizar_Click — adiciona codigo de consulta para esses 2 criterios
"""
import sys

FILE = r'C:\Projeto\OnlineCommerce\Forms\Vendas_Consulta_PorProdutos.frm'

with open(FILE, 'rb') as f:
    raw = f.read()

raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n')
text = raw.decode('windows-1252')

changes = []

# ------------------------------------------------------------------
# 1. cboCriterioPrinc_LostFocus — inserir PRODUTO/MENSAL e PRODUTO/PERIODO
#    entre o caso DATA e o caso SERVICOS
# ------------------------------------------------------------------
old1 = (
    "ElseIf cboCriterioPrinc.Text = \"DATA\" Then\n"
    "    lblInicio.Visible = True\n"
    "    lblInicio.Caption = \"Data\"\n"
    "    mskInicio.Visible = True\n"
    "    lblFim.Visible = False\n"
    "    mskFim.Visible = False\n"
    "    lblAte.Visible = False\n"
    "    cmdCalendario1.Visible = True\n"
    "    cmdCalendario2.Visible = False\n"
    "    lblMes.Visible = False\n"
    "    cboMes.Visible = False\n"
    "    lblAno.Visible = False\n"
    "    cboAno.Visible = False\n"
    "ElseIf cboCriterioPrinc.Text = \"SERVIÇOS\" Then\n"
)
new1 = (
    "ElseIf cboCriterioPrinc.Text = \"DATA\" Then\n"
    "    lblInicio.Visible = True\n"
    "    lblInicio.Caption = \"Data\"\n"
    "    mskInicio.Visible = True\n"
    "    lblFim.Visible = False\n"
    "    mskFim.Visible = False\n"
    "    lblAte.Visible = False\n"
    "    cmdCalendario1.Visible = True\n"
    "    cmdCalendario2.Visible = False\n"
    "    lblMes.Visible = False\n"
    "    cboMes.Visible = False\n"
    "    lblAno.Visible = False\n"
    "    cboAno.Visible = False\n"
    "ElseIf cboCriterioPrinc.Text = \"PRODUTO/MENSAL\" Then\n"
    "    lblInicio.Visible = False\n"
    "    mskInicio.Visible = False\n"
    "    lblFim.Visible = False\n"
    "    mskFim.Visible = False\n"
    "    lblAte.Visible = False\n"
    "    cmdCalendario1.Visible = False\n"
    "    cmdCalendario2.Visible = False\n"
    "    lblMes.Visible = True\n"
    "    cboMes.Visible = True\n"
    "    lblAno.Visible = True\n"
    "    cboAno.Visible = True\n"
    "    lblDescricao.Caption = \"Produto\"\n"
    "    lblDescricao.Visible = True\n"
    "    cboDescricao.Visible = True\n"
    "    txtCodBarra.Visible = False\n"
    "    LimparObjetos_Consulta\n"
    "    Exit Sub\n"
    "ElseIf cboCriterioPrinc.Text = \"PRODUTO/PERÍODO\" Then\n"
    "    lblInicio.Visible = True\n"
    "    lblInicio.Caption = \"Inicio\"\n"
    "    mskInicio.Visible = True\n"
    "    lblFim.Visible = True\n"
    "    mskFim.Visible = True\n"
    "    lblAte.Visible = True\n"
    "    cmdCalendario1.Visible = True\n"
    "    cmdCalendario2.Visible = True\n"
    "    lblMes.Visible = False\n"
    "    cboMes.Visible = False\n"
    "    lblAno.Visible = False\n"
    "    cboAno.Visible = False\n"
    "    lblDescricao.Caption = \"Produto\"\n"
    "    lblDescricao.Visible = True\n"
    "    cboDescricao.Visible = True\n"
    "    txtCodBarra.Visible = False\n"
    "    LimparObjetos_Consulta\n"
    "    Exit Sub\n"
    "ElseIf cboCriterioPrinc.Text = \"SERVIÇOS\" Then\n"
)
changes.append((old1, new1, '1 - PRODUTO/MENSAL e PRODUTO/PERIODO em LostFocus'))

# ------------------------------------------------------------------
# 2. cmdLocalizar_Click — adiciona casos PRODUTO/MENSAL e PRODUTO/PERIODO
#    antes do End If do bloco POR PRODUTOS
# ------------------------------------------------------------------
old2 = (
    "                sSQL = sSQL & \" and produtos.CATEGORIA = '\" & cboDescricao.Text & \"' and (pedidos_itens.data = CONVERT(DATETIME, '\" & Format(mskInicio.Text, ocDATA) & \"', 103)) \" & _\n"
    "                       \"ORDER BY \" & INDICE\n"
    "            End If\n"
    "            \n"
    "    'Debug.Print sSQL\n"
)
new2 = (
    "                sSQL = sSQL & \" and produtos.CATEGORIA = '\" & cboDescricao.Text & \"' and (pedidos_itens.data = CONVERT(DATETIME, '\" & Format(mskInicio.Text, ocDATA) & \"', 103)) \" & _\n"
    "                       \"ORDER BY \" & INDICE\n"
    "            'PRODUTO/MENSAL\n"
    "             ElseIf cboCriterioPrinc.Text = \"PRODUTO/MENSAL\" Then\n"
    "                If txtCodProduto.Text = \"\" Then Exit Sub\n"
    "                If cboMes.Text = \"\" Or cboAno.Text = \"\" Then Exit Sub\n"
    "                sSQL = sSQL & \" and produtos.codigo = \" & txtCodProduto.Text & \" and (MONTH(pedidos_itens.data) = \" & cboMes.ListIndex + 1 & \") AND (YEAR(pedidos_itens.data) = \" & cboAno & \") \" & _\n"
    "                       \"ORDER BY \" & INDICE\n"
    "            'PRODUTO/PERÍODO\n"
    "             ElseIf cboCriterioPrinc.Text = \"PRODUTO/PERÍODO\" Then\n"
    "                If txtCodProduto.Text = \"\" Then Exit Sub\n"
    "                If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then Exit Sub\n"
    "                sSQL = sSQL & \" and produtos.codigo = \" & txtCodProduto.Text & \" and (pedidos_itens.data >= CONVERT(DATETIME, '\" & Format(mskInicio.Text, ocDATA) & \"', 103)) AND (pedidos_itens.data <= CONVERT(DATETIME, '\" & Format(mskFim.Text, ocDATA) & \"', 103)) \" & _\n"
    "                       \"ORDER BY \" & INDICE\n"
    "            End If\n"
    "            \n"
    "    'Debug.Print sSQL\n"
)
changes.append((old2, new2, '2 - PRODUTO/MENSAL e PRODUTO/PERIODO em cmdLocalizar_Click'))

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

# Re-encode com CRLF
text = text.replace('\r\n', '\n').replace('\r', '\n')
out = text.encode('windows-1252')
out = out.replace(b'\n', b'\r\n')

with open(FILE, 'wb') as f:
    f.write(out)

print('\nArquivo gravado com sucesso.')
