# -*- coding: utf-8 -*-
"""
Patch 1: cboCriterioPrinc_LostFocus — adiciona cboDescricao.Text = "" nos 3 blocos SERVICOS*.
Patch 2: cmdImprimir_Click — reescreve bloco POR SERVICOS com preenchimento sequencial
         (sem gaps: 1 dado=rfCons1, 2 dados=rfCons1+rfCons2, 3 dados=todos).
"""
import sys

FILE = r'C:\Projeto\OnlineCommerce\Forms\Vendas_Consulta_PorProdutos.frm'

with open(FILE, 'rb') as f:
    raw = f.read()
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n')
text = raw.decode('windows-1252')

changes = []

# -----------------------------------------------------------------------
# PATCH 1: cboDescricao.Text = "" antes de LimparObjetos nos blocos SERVICOS*
# Ancora unica: lblDescricao.Caption = "Servi\xc7o" (so SERVICOS*, nao PRODUTO*)
# replace_all esperando 3 ocorrencias
# -----------------------------------------------------------------------
old_limpar = (
    '    lblDescricao.Caption = "Servi\xe7o"\n'
    '    lblDescricao.Visible = True\n'
    '    cboDescricao.Visible = True\n'
    '    txtCodBarra.Visible = False\n'
    '    LimparObjetos_Consulta\n'
    '    Exit Sub\n'
)
new_limpar = (
    '    lblDescricao.Caption = "Servi\xe7o"\n'
    '    lblDescricao.Visible = True\n'
    '    cboDescricao.Visible = True\n'
    '    txtCodBarra.Visible = False\n'
    '    cboDescricao.Text = ""\n'
    '    LimparObjetos_Consulta\n'
    '    Exit Sub\n'
)
changes.append((old_limpar, new_limpar, 'cboCriterioPrinc_LostFocus — limpa cboDescricao nos SERVICOS*', True, 3))

# -----------------------------------------------------------------------
# PATCH 2: cmdImprimir_Click — bloco POR SERVICOS: preenchimento sequencial
# -----------------------------------------------------------------------
old_servicos = (
    'ElseIf cboTipo.Text = "POR SERVI\xc7OS" Then\n'
    '\n'
    '    If cboCriterioPrinc.Text = "TODOS" Then\n'
    '        REL_Cons_Venda_Prod.rfCons1.Caption = "TODOS"\n'
    '        REL_Cons_Venda_Prod.rfCons3.Caption = ""\n'
    '    ElseIf cboCriterioPrinc.Text = "MENSAL" Then\n'
    '        REL_Cons_Venda_Prod.rfCons1.Caption = "MENSAL"\n'
    '        REL_Cons_Venda_Prod.rfCons3.Caption = "M\xeas/Ano = " & cboMes.Text & "/" & cboAno.Text\n'
    '    ElseIf cboCriterioPrinc.Text = "SERVI\xc7OS/MENSAL" Then\n'
    '        REL_Cons_Venda_Prod.rfCons1.Caption = "SERVI\xc7OS/MENSAL"\n'
    '        REL_Cons_Venda_Prod.rfCons3.Caption = "M\xeas/Ano = " & cboMes.Text & "/" & cboAno.Text\n'
    '    ElseIf cboCriterioPrinc.Text = "SERVI\xc7OS/PER\xcdODO" Then\n'
    '        REL_Cons_Venda_Prod.rfCons1.Caption = "SERVI\xc7OS/PER\xcdODO"\n'
    '        REL_Cons_Venda_Prod.rfCons3.Caption = "Inicio/Final = " & mskInicio.Text & " at\xe9 " & mskFim.Text\n'
    '    ElseIf cboCriterioPrinc.Text = "SERVI\xc7OS" Then\n'
    '        REL_Cons_Venda_Prod.rfCons1.Caption = "SERVI\xc7OS"\n'
    '        REL_Cons_Venda_Prod.rfCons3.Caption = ""\n'
    '    End If\n'
    '\n'
    '    If cboCriterioPrinc.Text = "SERVI\xc7OS/MENSAL" Or cboCriterioPrinc.Text = "SERVI\xc7OS/PER\xcdODO" Or cboCriterioPrinc.Text = "SERVI\xc7OS" Then\n'
    '        If txtCodProduto.Text <> "" Then\n'
    '            REL_Cons_Venda_Prod.rfCons2.Caption = "C\xd3D. SERV. = " & txtCodProduto.Text\n'
    '        Else\n'
    '            REL_Cons_Venda_Prod.rfCons2.Caption = ""\n'
    '        End If\n'
    '    ElseIf cboCriterioSec.Text = "DESCRI\xc7\xc3O" Then\n'
    '        REL_Cons_Venda_Prod.rfCons2.Caption = "DESCRI\xc7\xc3O = " & cboDescricao.Text\n'
    '    ElseIf cboCriterioSec.Text = "C\xd3D. OS" Then\n'
    '        REL_Cons_Venda_Prod.rfCons2.Caption = "C\xd3D. OS = " & txtCodBarra.Text\n'
    '    Else\n'
    '        REL_Cons_Venda_Prod.rfCons2.Caption = ""\n'
    '    End If\n'
    '\n'
    'End If\n'
)

new_servicos = (
    'ElseIf cboTipo.Text = "POR SERVI\xc7OS" Then\n'
    '\n'
    '    REL_Cons_Venda_Prod.rfCons1.Caption = cboCriterioPrinc.Text\n'
    '    REL_Cons_Venda_Prod.rfCons2.Caption = ""\n'
    '    REL_Cons_Venda_Prod.rfCons3.Caption = ""\n'
    '\n'
    '    If cboCriterioPrinc.Text = "TODOS" Then\n'
    '        If cboCriterioSec.Text = "DESCRI\xc7\xc3O" Then\n'
    '            REL_Cons_Venda_Prod.rfCons2.Caption = "DESCRI\xc7\xc3O = " & cboDescricao.Text\n'
    '        ElseIf cboCriterioSec.Text = "C\xd3D. OS" Then\n'
    '            REL_Cons_Venda_Prod.rfCons2.Caption = "C\xd3D. OS = " & txtCodBarra.Text\n'
    '        End If\n'
    '\n'
    '    ElseIf cboCriterioPrinc.Text = "MENSAL" Then\n'
    '        If cboCriterioSec.Text = "DESCRI\xc7\xc3O" Then\n'
    '            REL_Cons_Venda_Prod.rfCons2.Caption = "DESCRI\xc7\xc3O = " & cboDescricao.Text\n'
    '            REL_Cons_Venda_Prod.rfCons3.Caption = "M\xeas/Ano = " & cboMes.Text & "/" & cboAno.Text\n'
    '        ElseIf cboCriterioSec.Text = "C\xd3D. OS" Then\n'
    '            REL_Cons_Venda_Prod.rfCons2.Caption = "C\xd3D. OS = " & txtCodBarra.Text\n'
    '            REL_Cons_Venda_Prod.rfCons3.Caption = "M\xeas/Ano = " & cboMes.Text & "/" & cboAno.Text\n'
    '        Else\n'
    '            REL_Cons_Venda_Prod.rfCons2.Caption = "M\xeas/Ano = " & cboMes.Text & "/" & cboAno.Text\n'
    '        End If\n'
    '\n'
    '    ElseIf cboCriterioPrinc.Text = "SERVI\xc7OS" Then\n'
    '        If txtCodProduto.Text <> "" Then\n'
    '            REL_Cons_Venda_Prod.rfCons2.Caption = "C\xd3D. SERV. = " & txtCodProduto.Text\n'
    '        End If\n'
    '\n'
    '    ElseIf cboCriterioPrinc.Text = "SERVI\xc7OS/MENSAL" Then\n'
    '        If txtCodProduto.Text <> "" Then\n'
    '            REL_Cons_Venda_Prod.rfCons2.Caption = "C\xd3D. SERV. = " & txtCodProduto.Text\n'
    '            REL_Cons_Venda_Prod.rfCons3.Caption = "M\xeas/Ano = " & cboMes.Text & "/" & cboAno.Text\n'
    '        Else\n'
    '            REL_Cons_Venda_Prod.rfCons2.Caption = "M\xeas/Ano = " & cboMes.Text & "/" & cboAno.Text\n'
    '        End If\n'
    '\n'
    '    ElseIf cboCriterioPrinc.Text = "SERVI\xc7OS/PER\xcdODO" Then\n'
    '        If txtCodProduto.Text <> "" Then\n'
    '            REL_Cons_Venda_Prod.rfCons2.Caption = "C\xd3D. SERV. = " & txtCodProduto.Text\n'
    '            REL_Cons_Venda_Prod.rfCons3.Caption = "Inicio/Final = " & mskInicio.Text & " at\xe9 " & mskFim.Text\n'
    '        Else\n'
    '            REL_Cons_Venda_Prod.rfCons2.Caption = "Inicio/Final = " & mskInicio.Text & " at\xe9 " & mskFim.Text\n'
    '        End If\n'
    '\n'
    '    End If\n'
    '\n'
    'End If\n'
)
changes.append((old_servicos, new_servicos, 'cmdImprimir_Click — POR SERVICOS preenchimento sequencial', False, 1))

# -----------------------------------------------------------------------
# Aplicar
# -----------------------------------------------------------------------
for old, new, label, replace_all, expected in changes:
    count = text.count(old)
    if replace_all:
        if count != expected:
            print(f'ERRO [{label}]: {count} ocorrencias (esperado {expected})')
            sys.exit(1)
        text = text.replace(old, new)
        print(f'OK ({count}x): {label}')
    else:
        if count != 1:
            print(f'ERRO [{label}]: {count} ocorrencias (esperado 1)')
            sys.exit(1)
        text = text.replace(old, new)
        print(f'OK (1x): {label}')

out = text.encode('windows-1252')
out = out.replace(b'\n', b'\r\n')
with open(FILE, 'wb') as f:
    f.write(out)
print('\nArquivo gravado com sucesso.')
