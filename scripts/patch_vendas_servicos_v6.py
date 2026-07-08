# -*- coding: utf-8 -*-
"""
Patch v6: cboDescricao_GotFocus — popula produtos quando cboCriterioPrinc
          for PRODUTO/MENSAL ou PRODUTO/PERIODO (independente do cboCriterioSec)
"""
import sys

FILE = r'C:\Projeto\OnlineCommerce\Forms\Vendas_Consulta_PorProdutos.frm'

with open(FILE, 'rb') as f:
    raw = f.read()

raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n')
text = raw.decode('windows-1252')

old = (
    "   moCombo.AttachTo cboDescricao\n"
    "   Exit Sub\n"
    "End If\n"
    "\n"
    "If cboCriterioSec.Text = \"DESCRIÇÃO\" Then\n"
)
new = (
    "   moCombo.AttachTo cboDescricao\n"
    "   Exit Sub\n"
    "End If\n"
    "\n"
    "If cboCriterioPrinc.Text = \"PRODUTO/MENSAL\" Or cboCriterioPrinc.Text = \"PRODUTO/PERÍODO\" Then\n"
    "   sSQL = \"SELECT DISTINCT descricao, codigo FROM produtos ORDER BY descricao;\"\n"
    "   Set r = dbData.OpenRecordset(sSQL)\n"
    "   Do While Not r.EOF\n"
    "      cboDescricao.AddItem r(\"descricao\")\n"
    "      cboDescricao.ItemData(cboDescricao.NewIndex) = r(\"codigo\")\n"
    "      r.MoveNext\n"
    "   Loop\n"
    "   If r.State <> 0 Then r.Close\n"
    "   Set r = Nothing\n"
    "   moCombo.AttachTo cboDescricao\n"
    "   Exit Sub\n"
    "End If\n"
    "\n"
    "If cboCriterioSec.Text = \"DESCRIÇÃO\" Then\n"
)

count = text.count(old)
if count != 1:
    print(f'ERRO: encontrado {count} ocorrencias (esperado 1)')
    sys.exit(1)

text = text.replace(old, new)
print('OK: cboDescricao_GotFocus — PRODUTO/MENSAL e PRODUTO/PERIODO carregam produtos')

text = text.replace('\r\n', '\n').replace('\r', '\n')
out = text.encode('windows-1252')
out = out.replace(b'\n', b'\r\n')

with open(FILE, 'wb') as f:
    f.write(out)

print('Arquivo gravado com sucesso.')
