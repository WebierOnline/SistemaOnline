# -*- coding: utf-8 -*-
"""
Patch v2: POR SERVIÇOS — novos critérios SERVIÇOS/MENSAL e SERVIÇOS/PERÍODO,
cboDescricao via OS_Servicos, suporte a CÓD. OS.
"""
import sys

FILE = r'C:\Projeto\OnlineCommerce\Forms\Vendas_Consulta_PorProdutos.frm'

with open(FILE, 'rb') as f:
    raw = f.read()

raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n')
text = raw.decode('windows-1252')

changes = []

# ------------------------------------------------------------------
# 1. cboCriterioPrinc_LostFocus — substitui ESPECIFICO/MENSAL e ESPECIFICO
#    pelos novos: SERVIÇOS, SERVIÇOS/MENSAL, SERVIÇOS/PERÍODO (com Exit Sub)
# ------------------------------------------------------------------
old1 = (
    "ElseIf cboCriterioPrinc.Text = \"ESPECIFICO/MENSAL\" Then\n"
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
    "ElseIf cboCriterioPrinc.Text = \"ESPECIFICO\" Then\n"
    "    lblInicio.Visible = False\n"
    "    mskInicio.Visible = False\n"
    "    lblFim.Visible = False\n"
    "    mskFim.Visible = False\n"
    "    lblAte.Visible = False\n"
    "    cmdCalendario1.Visible = False\n"
    "    cmdCalendario2.Visible = False\n"
    "    lblMes.Visible = False\n"
    "    cboMes.Visible = False\n"
    "    lblAno.Visible = False\n"
    "    cboAno.Visible = False\n"
    "End If\n"
)
new1 = (
    "ElseIf cboCriterioPrinc.Text = \"SERVIÇOS\" Then\n"
    "    lblInicio.Visible = False\n"
    "    mskInicio.Visible = False\n"
    "    lblFim.Visible = False\n"
    "    mskFim.Visible = False\n"
    "    lblAte.Visible = False\n"
    "    cmdCalendario1.Visible = False\n"
    "    cmdCalendario2.Visible = False\n"
    "    lblMes.Visible = False\n"
    "    cboMes.Visible = False\n"
    "    lblAno.Visible = False\n"
    "    cboAno.Visible = False\n"
    "    lblDescricao.Caption = \"Serviço\"\n"
    "    lblDescricao.Visible = True\n"
    "    cboDescricao.Visible = True\n"
    "    txtCodBarra.Visible = False\n"
    "    LimparObjetos_Consulta\n"
    "    Exit Sub\n"
    "ElseIf cboCriterioPrinc.Text = \"SERVIÇOS/MENSAL\" Then\n"
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
    "    lblDescricao.Caption = \"Serviço\"\n"
    "    lblDescricao.Visible = True\n"
    "    cboDescricao.Visible = True\n"
    "    txtCodBarra.Visible = False\n"
    "    LimparObjetos_Consulta\n"
    "    Exit Sub\n"
    "ElseIf cboCriterioPrinc.Text = \"SERVIÇOS/PERÍODO\" Then\n"
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
    "    lblDescricao.Caption = \"Serviço\"\n"
    "    lblDescricao.Visible = True\n"
    "    cboDescricao.Visible = True\n"
    "    txtCodBarra.Visible = False\n"
    "    LimparObjetos_Consulta\n"
    "    Exit Sub\n"
    "End If\n"
)
changes.append((old1, new1, '1 - SERVIÇOS/MENSAL e SERVIÇOS/PERÍODO em cboCriterioPrinc_LostFocus'))

# ------------------------------------------------------------------
# 2. cboCriterioPrinc_LostFocus segundo bloco — adiciona "CÓD. OS"
# ------------------------------------------------------------------
old2 = (
    "ElseIf cboCriterioSec.Text = \"CÓD. BARRA\" Then\n"
    "    lblDescricao.Caption = \"Cód. Barra\"\n"
    "    lblDescricao.Visible = True\n"
    "    cboDescricao.Visible = False\n"
    "    txtCodBarra.Visible = True\n"
    "Else\n"
    "End If\n"
    "\n"
    "\n"
    "LimparObjetos_Consulta\n"
    "End Sub\n"
)
new2 = (
    "ElseIf cboCriterioSec.Text = \"CÓD. BARRA\" Then\n"
    "    lblDescricao.Caption = \"Cód. Barra\"\n"
    "    lblDescricao.Visible = True\n"
    "    cboDescricao.Visible = False\n"
    "    txtCodBarra.Visible = True\n"
    "ElseIf cboCriterioSec.Text = \"CÓD. OS\" Then\n"
    "    lblDescricao.Caption = \"Cód. OS\"\n"
    "    lblDescricao.Visible = True\n"
    "    cboDescricao.Visible = False\n"
    "    txtCodBarra.Visible = True\n"
    "Else\n"
    "End If\n"
    "\n"
    "\n"
    "LimparObjetos_Consulta\n"
    "End Sub\n"
)
changes.append((old2, new2, '2 - CÓD. OS em cboCriterioPrinc_LostFocus segundo bloco'))

# ------------------------------------------------------------------
# 3. cboCriterioSec_LostFocus — adiciona "CÓD. OS"
# ------------------------------------------------------------------
old3 = (
    "ElseIf cboCriterioSec.Text = \"CÓD. BARRA\" Then\n"
    "    lblDescricao.Caption = \"Cód. Barra\"\n"
    "    lblDescricao.Visible = True\n"
    "    cboDescricao.Visible = False\n"
    "    txtCodBarra.Visible = True\n"
    "Else\n"
    "End If\n"
    "End Sub\n"
)
new3 = (
    "ElseIf cboCriterioSec.Text = \"CÓD. BARRA\" Then\n"
    "    lblDescricao.Caption = \"Cód. Barra\"\n"
    "    lblDescricao.Visible = True\n"
    "    cboDescricao.Visible = False\n"
    "    txtCodBarra.Visible = True\n"
    "ElseIf cboCriterioSec.Text = \"CÓD. OS\" Then\n"
    "    lblDescricao.Caption = \"Cód. OS\"\n"
    "    lblDescricao.Visible = True\n"
    "    cboDescricao.Visible = False\n"
    "    txtCodBarra.Visible = True\n"
    "Else\n"
    "End If\n"
    "End Sub\n"
)
changes.append((old3, new3, '3 - CÓD. OS em cboCriterioSec_LostFocus'))

# ------------------------------------------------------------------
# 4. cboDescricao_GotFocus — carrega OS_Servicos (com ItemData = CODIGO)
# ------------------------------------------------------------------
old4 = (
    "If cboTipo.Text = \"POR SERVIÇOS\" Then\n"
    "   sSQL = \"SELECT DISTINCT descricao FROM OS_Servicos_Auto ORDER BY descricao;\"\n"
    "   Set r = dbData.OpenRecordset(sSQL)\n"
    "   Do While Not r.EOF\n"
    "      cboDescricao.AddItem r(\"descricao\")\n"
    "      r.MoveNext\n"
    "   Loop\n"
    "   If r.State <> 0 Then r.Close\n"
    "   Set r = Nothing\n"
    "   moCombo.AttachTo cboDescricao\n"
    "   Exit Sub\n"
    "End If\n"
)
new4 = (
    "If cboTipo.Text = \"POR SERVIÇOS\" Then\n"
    "   sSQL = \"SELECT SERVICO, CODIGO FROM OS_Servicos ORDER BY SERVICO;\"\n"
    "   Set r = dbData.OpenRecordset(sSQL)\n"
    "   Do While Not r.EOF\n"
    "      cboDescricao.AddItem r(\"SERVICO\")\n"
    "      cboDescricao.ItemData(cboDescricao.NewIndex) = r(\"CODIGO\")\n"
    "      r.MoveNext\n"
    "   Loop\n"
    "   If r.State <> 0 Then r.Close\n"
    "   Set r = Nothing\n"
    "   moCombo.AttachTo cboDescricao\n"
    "   Exit Sub\n"
    "End If\n"
)
changes.append((old4, new4, '4 - cboDescricao_GotFocus carrega OS_Servicos'))

# ------------------------------------------------------------------
# 5. cmdLocalizar_Click bloco POR SERVIÇOS — atualiza critérios e SQL
# ------------------------------------------------------------------
old5 = (
    "ElseIf cboTipo.Text = \"POR SERVIÇOS\" Then\n"
    "   Dim sBase As String\n"
    "   sBase = \"SELECT s.codigo, OS.COD_OS AS varCodPed, s.data AS varData, s.descricao AS varNome, \" & _\n"
    "           \"s.preco AS varValor, s.quantidade AS varQuant, s.subtotal AS varSubtotal, \" & _\n"
    "           \"s.desconto AS varDesc, s.total AS varTotal, 0 AS var_CodOS \" & _\n"
    "           \"FROM OS_Servicos_Auto s INNER JOIN OS ON s.cod_os = OS.COD_OS \"\n"
    "\n"
    "   If cboCriterioPrinc.Text = \"TODOS\" Then\n"
    "      If cboCriterioSec.Text = \"DESCRIÇÃO\" Then\n"
    "         If cboDescricao.Text = \"\" Then Exit Sub\n"
    "         sSQL = sBase & \"WHERE s.descricao = '\" & cboDescricao.Text & \"' ORDER BY \" & INDICE\n"
    "      Else\n"
    "         sSQL = sBase & \"ORDER BY \" & INDICE\n"
    "      End If\n"
    "   ElseIf cboCriterioPrinc.Text = \"MENSAL\" Then\n"
    "      If cboMes.Text = \"\" Or cboAno.Text = \"\" Then Exit Sub\n"
    "      If cboCriterioSec.Text = \"DESCRIÇÃO\" Then\n"
    "         If cboDescricao.Text = \"\" Then Exit Sub\n"
    "         sSQL = sBase & \"WHERE s.descricao = '\" & cboDescricao.Text & \"' AND MONTH(s.data) = \" & cboMes.ListIndex + 1 & \" AND YEAR(s.data) = \" & cboAno & \" ORDER BY \" & INDICE\n"
    "      Else\n"
    "         sSQL = sBase & \"WHERE MONTH(s.data) = \" & cboMes.ListIndex + 1 & \" AND YEAR(s.data) = \" & cboAno & \" ORDER BY \" & INDICE\n"
    "      End If\n"
    "   ElseIf cboCriterioPrinc.Text = \"ESPECIFICO/MENSAL\" Then\n"
    "      If cboDescricao.Text = \"\" Then Exit Sub\n"
    "      If cboMes.Text = \"\" Or cboAno.Text = \"\" Then Exit Sub\n"
    "      sSQL = sBase & \"WHERE s.descricao = '\" & cboDescricao.Text & \"' AND MONTH(s.data) = \" & cboMes.ListIndex + 1 & \" AND YEAR(s.data) = \" & cboAno & \" ORDER BY \" & INDICE\n"
    "   ElseIf cboCriterioPrinc.Text = \"ESPECIFICO\" Then\n"
    "      If cboDescricao.Text = \"\" Then Exit Sub\n"
    "      sSQL = sBase & \"WHERE s.descricao = '\" & cboDescricao.Text & \"' ORDER BY \" & INDICE\n"
    "   End If\n"
    "End If\n"
)
new5 = (
    "ElseIf cboTipo.Text = \"POR SERVIÇOS\" Then\n"
    "   Dim sBase As String\n"
    "   sBase = \"SELECT s.codigo, OS.COD_OS AS varCodPed, s.data AS varData, s.descricao AS varNome, \" & _\n"
    "           \"s.preco AS varValor, s.quantidade AS varQuant, s.subtotal AS varSubtotal, \" & _\n"
    "           \"s.desconto AS varDesc, s.total AS varTotal, 0 AS var_CodOS \" & _\n"
    "           \"FROM OS_Servicos_Auto s INNER JOIN OS ON s.cod_os = OS.COD_OS \"\n"
    "\n"
    "   If cboCriterioPrinc.Text = \"TODOS\" Then\n"
    "      If cboCriterioSec.Text = \"DESCRIÇÃO\" Then\n"
    "         If cboDescricao.Text = \"\" Or txtCodProduto.Text = \"\" Then Exit Sub\n"
    "         sSQL = sBase & \"WHERE s.cod_servico = \" & txtCodProduto.Text & \" ORDER BY \" & INDICE\n"
    "      ElseIf cboCriterioSec.Text = \"CÓD. OS\" Then\n"
    "         If txtCodBarra.Text = \"\" Then Exit Sub\n"
    "         sSQL = sBase & \"WHERE OS.COD_OS = \" & Val(txtCodBarra.Text) & \" ORDER BY \" & INDICE\n"
    "      Else\n"
    "         sSQL = sBase & \"ORDER BY \" & INDICE\n"
    "      End If\n"
    "   ElseIf cboCriterioPrinc.Text = \"MENSAL\" Then\n"
    "      If cboMes.Text = \"\" Or cboAno.Text = \"\" Then Exit Sub\n"
    "      If cboCriterioSec.Text = \"DESCRIÇÃO\" Then\n"
    "         If cboDescricao.Text = \"\" Or txtCodProduto.Text = \"\" Then Exit Sub\n"
    "         sSQL = sBase & \"WHERE s.cod_servico = \" & txtCodProduto.Text & \" AND MONTH(s.data) = \" & cboMes.ListIndex + 1 & \" AND YEAR(s.data) = \" & cboAno & \" ORDER BY \" & INDICE\n"
    "      ElseIf cboCriterioSec.Text = \"CÓD. OS\" Then\n"
    "         If txtCodBarra.Text = \"\" Then Exit Sub\n"
    "         sSQL = sBase & \"WHERE OS.COD_OS = \" & Val(txtCodBarra.Text) & \" AND MONTH(s.data) = \" & cboMes.ListIndex + 1 & \" AND YEAR(s.data) = \" & cboAno & \" ORDER BY \" & INDICE\n"
    "      Else\n"
    "         sSQL = sBase & \"WHERE MONTH(s.data) = \" & cboMes.ListIndex + 1 & \" AND YEAR(s.data) = \" & cboAno & \" ORDER BY \" & INDICE\n"
    "      End If\n"
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
changes.append((old5, new5, '5 - cmdLocalizar_Click bloco POR SERVIÇOS completo'))

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
