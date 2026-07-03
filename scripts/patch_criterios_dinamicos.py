# -*- coding: utf-8 -*-
"""
Patch: Vendas_Consulta_PorProdutos.frm
1. cboCriterioPrinc_GotFocus: PRODUTO/* e SERVICOS/* so aparecem quando cboCriterioSec=TODOS;
   adiciona PERIODO para POR SERVICOS.
2. cboTipo_Change: adiciona PERIODO para POR SERVICOS.
3. cboCriterioSec_LostFocus: recarrega cboCriterioPrinc + limpa cboDescricao/cboCriterioPrinc.
4. cmdLocalizar_Click: SQL para PERIODO em POR SERVICOS.
5. cmdImprimir_Click: rfCons para PERIODO em POR SERVICOS.
"""
import sys

FILE = r'C:\Projeto\OnlineCommerce\Forms\Vendas_Consulta_PorProdutos.frm'

with open(FILE, 'rb') as f:
    raw = f.read()
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n')
text = raw.decode('windows-1252')

changes = []

# -----------------------------------------------------------------------
# 1. cboCriterioPrinc_GotFocus
# -----------------------------------------------------------------------
old1 = (
    'Sub cboCriterioPrinc_GotFocus()\n'
    'cboCriterioPrinc.Clear\n'
    '   \n'
    'If cboTipo.Text = "POR PRODUTOS" Then\n'
    '   If cboCriterioSec.Text <> "TODOS" Then cboCriterioPrinc.AddItem "TODOS"\n'
    '   cboCriterioPrinc.AddItem "MENSAL"\n'
    '   cboCriterioPrinc.AddItem "PER\xcdODO"\n'
    '   cboCriterioPrinc.AddItem "DATA"\n'
    '   cboCriterioPrinc.AddItem "PRODUTO/MENSAL"\n'
    '   cboCriterioPrinc.AddItem "PRODUTO/PER\xcdODO"\n'
    'ElseIf cboTipo.Text = "POR SERVI\xc7OS" Then\n'
    '   If cboCriterioSec.Text <> "TODOS" Then cboCriterioPrinc.AddItem "TODOS"\n'
    '   cboCriterioPrinc.AddItem "MENSAL"\n'
    '   cboCriterioPrinc.AddItem "SERVI\xc7OS"\n'
    '   cboCriterioPrinc.AddItem "SERVI\xc7OS/MENSAL"\n'
    '   cboCriterioPrinc.AddItem "SERVI\xc7OS/PER\xcdODO"\n'
    'End If\n'
    '   \n'
    'moCombo.AttachTo cboCriterioPrinc\n'
    'End Sub'
)
new1 = (
    'Sub cboCriterioPrinc_GotFocus()\n'
    'cboCriterioPrinc.Clear\n'
    '   \n'
    'If cboTipo.Text = "POR PRODUTOS" Then\n'
    '   If cboCriterioSec.Text <> "TODOS" Then cboCriterioPrinc.AddItem "TODOS"\n'
    '   cboCriterioPrinc.AddItem "MENSAL"\n'
    '   cboCriterioPrinc.AddItem "PER\xcdODO"\n'
    '   cboCriterioPrinc.AddItem "DATA"\n'
    '   If cboCriterioSec.Text = "TODOS" Then\n'
    '      cboCriterioPrinc.AddItem "PRODUTO/MENSAL"\n'
    '      cboCriterioPrinc.AddItem "PRODUTO/PER\xcdODO"\n'
    '   End If\n'
    'ElseIf cboTipo.Text = "POR SERVI\xc7OS" Then\n'
    '   If cboCriterioSec.Text <> "TODOS" Then cboCriterioPrinc.AddItem "TODOS"\n'
    '   cboCriterioPrinc.AddItem "MENSAL"\n'
    '   cboCriterioPrinc.AddItem "PER\xcdODO"\n'
    '   cboCriterioPrinc.AddItem "SERVI\xc7OS"\n'
    '   If cboCriterioSec.Text = "TODOS" Then\n'
    '      cboCriterioPrinc.AddItem "SERVI\xc7OS/MENSAL"\n'
    '      cboCriterioPrinc.AddItem "SERVI\xc7OS/PER\xcdODO"\n'
    '   End If\n'
    'End If\n'
    '   \n'
    'moCombo.AttachTo cboCriterioPrinc\n'
    'End Sub'
)
changes.append((old1, new1, 'cboCriterioPrinc_GotFocus — condicionar PRODUTO/* e SERVICOS/*, add PERIODO', False, 1))

# -----------------------------------------------------------------------
# 2. cboTipo_Change — adiciona PERIODO para POR SERVICOS
# -----------------------------------------------------------------------
old2 = (
    'ElseIf cboTipo.Text = "POR SERVI\xc7OS" Then\n'
    '   cboCriterioPrinc.AddItem "MENSAL"\n'
    '   cboCriterioPrinc.AddItem "SERVI\xc7OS"\n'
    '   cboCriterioPrinc.AddItem "SERVI\xc7OS/MENSAL"\n'
    '   cboCriterioPrinc.AddItem "SERVI\xc7OS/PER\xcdODO"\n'
    'End If\n'
    'cboCriterioPrinc.ListIndex = 0'
)
new2 = (
    'ElseIf cboTipo.Text = "POR SERVI\xc7OS" Then\n'
    '   cboCriterioPrinc.AddItem "MENSAL"\n'
    '   cboCriterioPrinc.AddItem "PER\xcdODO"\n'
    '   cboCriterioPrinc.AddItem "SERVI\xc7OS"\n'
    '   cboCriterioPrinc.AddItem "SERVI\xc7OS/MENSAL"\n'
    '   cboCriterioPrinc.AddItem "SERVI\xc7OS/PER\xcdODO"\n'
    'End If\n'
    'cboCriterioPrinc.ListIndex = 0'
)
changes.append((old2, new2, 'cboTipo_Change — adiciona PERIODO para POR SERVICOS', False, 1))

# -----------------------------------------------------------------------
# 3. cboCriterioSec_LostFocus — recarregar cboCriterioPrinc ao mudar criterio
# -----------------------------------------------------------------------
old3 = (
    'ElseIf cboCriterioSec.Text = "TODOS" Then\n'
    '    lblDescricao.Visible = False\n'
    '    cboDescricao.Visible = False\n'
    '    txtCodBarra.Visible = False\n'
    'Else\n'
    'End If\n'
    'End Sub'
)
new3 = (
    'ElseIf cboCriterioSec.Text = "TODOS" Then\n'
    '    lblDescricao.Visible = False\n'
    '    cboDescricao.Visible = False\n'
    '    txtCodBarra.Visible = False\n'
    'Else\n'
    'End If\n'
    '\n'
    'cboCriterioPrinc.Clear\n'
    'If cboTipo.Text = "POR PRODUTOS" Then\n'
    '   If cboCriterioSec.Text <> "TODOS" Then cboCriterioPrinc.AddItem "TODOS"\n'
    '   cboCriterioPrinc.AddItem "MENSAL"\n'
    '   cboCriterioPrinc.AddItem "PER\xcdODO"\n'
    '   cboCriterioPrinc.AddItem "DATA"\n'
    '   If cboCriterioSec.Text = "TODOS" Then\n'
    '      cboCriterioPrinc.AddItem "PRODUTO/MENSAL"\n'
    '      cboCriterioPrinc.AddItem "PRODUTO/PER\xcdODO"\n'
    '   End If\n'
    'ElseIf cboTipo.Text = "POR SERVI\xc7OS" Then\n'
    '   If cboCriterioSec.Text <> "TODOS" Then cboCriterioPrinc.AddItem "TODOS"\n'
    '   cboCriterioPrinc.AddItem "MENSAL"\n'
    '   cboCriterioPrinc.AddItem "PER\xcdODO"\n'
    '   cboCriterioPrinc.AddItem "SERVI\xc7OS"\n'
    '   If cboCriterioSec.Text = "TODOS" Then\n'
    '      cboCriterioPrinc.AddItem "SERVI\xc7OS/MENSAL"\n'
    '      cboCriterioPrinc.AddItem "SERVI\xc7OS/PER\xcdODO"\n'
    '   End If\n'
    'End If\n'
    'cboDescricao.Text = ""\n'
    'cboCriterioPrinc.ListIndex = 0\n'
    'cboCriterioPrinc_LostFocus\n'
    'End Sub'
)
changes.append((old3, new3, 'cboCriterioSec_LostFocus — recarregar cboCriterioPrinc + limpar', False, 1))

# -----------------------------------------------------------------------
# 4. cmdLocalizar_Click — SQL PERIODO para POR SERVICOS
# Insere antes de ElseIf SERVICOS/MENSAL
# -----------------------------------------------------------------------
old4 = (
    '   ElseIf cboCriterioPrinc.Text = "SERVI\xc7OS/MENSAL" Then\n'
    '      If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub\n'
)
new4 = (
    '   ElseIf cboCriterioPrinc.Text = "PER\xcdODO" Then\n'
    '      If Not IsDate(mskInicio.Text) Or Not IsDate(mskFim.Text) Then Exit Sub\n'
    '      If cboCriterioSec.Text = "DESCRI\xc7\xc3O" Then\n'
    '         If cboDescricao.Text = "" Or txtCodProduto.Text = "" Then Exit Sub\n'
    '         sSQL = sBase & "WHERE s.cod_servico = " & txtCodProduto.Text & " AND (s.data >= CONVERT(DATETIME, \'" & Format(mskInicio.Text, ocDATA) & "\', 103)) AND (s.data <= CONVERT(DATETIME, \'" & Format(mskFim.Text, ocDATA) & "\', 103)) ORDER BY " & INDICE\n'
    '      ElseIf cboCriterioSec.Text = "C\xd3D. OS" Then\n'
    '         If txtCodBarra.Text = "" Then Exit Sub\n'
    '         sSQL = sBase & "WHERE OS.COD_OS = " & Val(txtCodBarra.Text) & " AND (s.data >= CONVERT(DATETIME, \'" & Format(mskInicio.Text, ocDATA) & "\', 103)) AND (s.data <= CONVERT(DATETIME, \'" & Format(mskFim.Text, ocDATA) & "\', 103)) ORDER BY " & INDICE\n'
    '      Else\n'
    '         sSQL = sBase & "WHERE (s.data >= CONVERT(DATETIME, \'" & Format(mskInicio.Text, ocDATA) & "\', 103)) AND (s.data <= CONVERT(DATETIME, \'" & Format(mskFim.Text, ocDATA) & "\', 103)) ORDER BY " & INDICE\n'
    '      End If\n'
    '   ElseIf cboCriterioPrinc.Text = "SERVI\xc7OS/MENSAL" Then\n'
    '      If cboMes.Text = "" Or cboAno.Text = "" Then Exit Sub\n'
)
changes.append((old4, new4, 'cmdLocalizar_Click — SQL PERIODO para POR SERVICOS', False, 1))

# -----------------------------------------------------------------------
# 5. cmdImprimir_Click — rfCons PERIODO em POR SERVICOS
# Insere antes de ElseIf SERVICOS (usa ancora com rfCons2 = Servico)
# -----------------------------------------------------------------------
old5 = (
    '    ElseIf cboCriterioPrinc.Text = "SERVI\xc7OS" Then\n'
    '        If txtCodProduto.Text <> "" Then\n'
    '            REL_Cons_Venda_Prod.rfCons2.Caption = "SERVI\xe7o = " & cboDescricao.Text\n'
    '        End If\n'
)
new5 = (
    '    ElseIf cboCriterioPrinc.Text = "PER\xcdODO" Then\n'
    '        If cboCriterioSec.Text = "DESCRI\xc7\xc3O" Then\n'
    '            REL_Cons_Venda_Prod.rfCons2.Caption = "DESCRI\xc7\xc3O = " & cboDescricao.Text\n'
    '            REL_Cons_Venda_Prod.rfCons3.Caption = "Inicio/Final = " & mskInicio.Text & " at\xe9 " & mskFim.Text\n'
    '        ElseIf cboCriterioSec.Text = "C\xd3D. OS" Then\n'
    '            REL_Cons_Venda_Prod.rfCons2.Caption = "C\xd3D. OS = " & txtCodBarra.Text\n'
    '            REL_Cons_Venda_Prod.rfCons3.Caption = "Inicio/Final = " & mskInicio.Text & " at\xe9 " & mskFim.Text\n'
    '        Else\n'
    '            REL_Cons_Venda_Prod.rfCons2.Caption = "Inicio/Final = " & mskInicio.Text & " at\xe9 " & mskFim.Text\n'
    '        End If\n'
    '\n'
    '    ElseIf cboCriterioPrinc.Text = "SERVI\xc7OS" Then\n'
    '        If txtCodProduto.Text <> "" Then\n'
    '            REL_Cons_Venda_Prod.rfCons2.Caption = "SERVI\xe7o = " & cboDescricao.Text\n'
    '        End If\n'
)
changes.append((old5, new5, 'cmdImprimir_Click — rfCons PERIODO em POR SERVICOS', False, 1))

# -----------------------------------------------------------------------
# Aplicar
# -----------------------------------------------------------------------
for old, new, label, replace_all, expected in changes:
    count = text.count(old)
    if count != expected:
        print(f'ERRO [{label}]: {count} ocorrencias (esperado {expected})')
        sys.exit(1)
    text = text.replace(old, new)
    print(f'OK (1x): {label}')

out = text.encode('windows-1252')
out = out.replace(b'\n', b'\r\n')
with open(FILE, 'wb') as f:
    f.write(out)
print('\nArquivo gravado com sucesso.')
