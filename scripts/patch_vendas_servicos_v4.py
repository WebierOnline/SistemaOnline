# -*- coding: utf-8 -*-
"""
Patch v4: SERVIÇOS/* tornam-se opcionais — sem serviço selecionado traz
todos os registros de OS_Servicos_Auto com o critério de data.
"""
import sys

FILE = r'C:\Projeto\OnlineCommerce\Forms\Vendas_Consulta_PorProdutos.frm'

with open(FILE, 'rb') as f:
    raw = f.read()

raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n')
text = raw.decode('windows-1252')

# ------------------------------------------------------------------
# Único patch: SERVIÇOS/MENSAL, SERVIÇOS/PERÍODO e SERVIÇOS
# ------------------------------------------------------------------
old = (
    "   ElseIf cboCriterioPrinc.Text = \"SERVIÇOS/MENSAL\" Then\n"
    "      If cboDescricao.Text = \"\" Or txtCodProduto.Text = \"\" Then Exit Sub\n"
    "      If cboMes.Text = \"\" Or cboAno.Text = \"\" Then Exit Sub\n"
    "      sSQL = sBase & \"WHERE s.cod_servico = \" & txtCodProduto.Text & \" AND MONTH(s.data) = \" & cboMes.ListIndex + 1 & \" AND YEAR(s.data) = \" & cboAno & \" ORDER BY \" & INDICE\n"
    "   ElseIf cboCriterioPrinc.Text = \"SERVIÇOS/PERÍODO\" Then\n"
    "      If cboDescricao.Text = \"\" Or txtCodProduto.Text = \"\" Then Exit Sub\n"
    "      If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then Exit Sub\n"
    "      sSQL = sBase & \"WHERE s.cod_servico = \" & txtCodProduto.Text & \" AND (s.data >= CONVERT(DATETIME, '\" & Format(mskInicio.Text, ocDATA) & \"', 103)) AND (s.data <= CONVERT(DATETIME, '\" & Format(mskFim.Text, ocDATA) & \"', 103)) ORDER BY \" & INDICE\n"
    "   ElseIf cboCriterioPrinc.Text = \"SERVIÇOS\" Then\n"
    "      If cboDescricao.Text = \"\" Or txtCodProduto.Text = \"\" Then Exit Sub\n"
    "      sSQL = sBase & \"WHERE s.cod_servico = \" & txtCodProduto.Text & \" ORDER BY \" & INDICE\n"
    "   End If\n"
    "End If\n"
)

new = (
    "   ElseIf cboCriterioPrinc.Text = \"SERVIÇOS/MENSAL\" Then\n"
    "      If cboMes.Text = \"\" Or cboAno.Text = \"\" Then Exit Sub\n"
    "      If txtCodProduto.Text <> \"\" Then\n"
    "         sSQL = sBase & \"WHERE s.cod_servico = \" & txtCodProduto.Text & \" AND MONTH(s.data) = \" & cboMes.ListIndex + 1 & \" AND YEAR(s.data) = \" & cboAno & \" ORDER BY \" & INDICE\n"
    "      Else\n"
    "         sSQL = sBase & \"WHERE MONTH(s.data) = \" & cboMes.ListIndex + 1 & \" AND YEAR(s.data) = \" & cboAno & \" ORDER BY \" & INDICE\n"
    "      End If\n"
    "   ElseIf cboCriterioPrinc.Text = \"SERVIÇOS/PERÍODO\" Then\n"
    "      If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then Exit Sub\n"
    "      If txtCodProduto.Text <> \"\" Then\n"
    "         sSQL = sBase & \"WHERE s.cod_servico = \" & txtCodProduto.Text & \" AND (s.data >= CONVERT(DATETIME, '\" & Format(mskInicio.Text, ocDATA) & \"', 103)) AND (s.data <= CONVERT(DATETIME, '\" & Format(mskFim.Text, ocDATA) & \"', 103)) ORDER BY \" & INDICE\n"
    "      Else\n"
    "         sSQL = sBase & \"WHERE (s.data >= CONVERT(DATETIME, '\" & Format(mskInicio.Text, ocDATA) & \"', 103)) AND (s.data <= CONVERT(DATETIME, '\" & Format(mskFim.Text, ocDATA) & \"', 103)) ORDER BY \" & INDICE\n"
    "      End If\n"
    "   ElseIf cboCriterioPrinc.Text = \"SERVIÇOS\" Then\n"
    "      If txtCodProduto.Text <> \"\" Then\n"
    "         sSQL = sBase & \"WHERE s.cod_servico = \" & txtCodProduto.Text & \" ORDER BY \" & INDICE\n"
    "      Else\n"
    "         sSQL = sBase & \"ORDER BY \" & INDICE\n"
    "      End If\n"
    "   End If\n"
    "End If\n"
)

count = text.count(old)
if count != 1:
    print(f'ERRO: encontrado {count} ocorrencias (esperado 1)')
    sys.exit(1)

text = text.replace(old, new)
print('OK: SERVICOS/* tornam-se opcionais')

# Re-encode com CRLF
text = text.replace('\r\n', '\n').replace('\r', '\n')
out = text.encode('windows-1252')
out = out.replace(b'\n', b'\r\n')

with open(FILE, 'wb') as f:
    f.write(out)

print('Arquivo gravado com sucesso.')
