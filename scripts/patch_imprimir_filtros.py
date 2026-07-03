# -*- coding: utf-8 -*-
"""
Atualiza bloco de filtros do cmdImprimir_Click em Vendas_Consulta_PorProdutos.frm.
Adiciona: DATA, PRODUTO/MENSAL, PRODUTO/PERIODO (POR PRODUTOS); CATEGORIA (cboCriterioSec);
e todo o bloco POR SERVICOS (TODOS/MENSAL/SERVICOS/MENSAL/SERVICOS/PERIODO/SERVICOS).
"""
import sys

FILE = r'C:\Projeto\OnlineCommerce\Forms\Vendas_Consulta_PorProdutos.frm'

with open(FILE, 'rb') as f:
    raw = f.read()
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n')
text = raw.decode('windows-1252')

old_block = (
    'If cboCriterioPrinc.Text = "TODOS" Then\n'
    '    REL_Cons_Venda_Prod.rfCons1.Caption = "TODOS"\n'
    '    REL_Cons_Venda_Prod.rfCons3.Caption = ""\n'
    'ElseIf cboCriterioPrinc.Text = "MENSAL" Then\n'
    '    REL_Cons_Venda_Prod.rfCons1.Caption = "MENSAL"\n'
    '    REL_Cons_Venda_Prod.rfCons3.Caption = "M\xeas/Ano = " & cboMes.Text & "/" & cboAno.Text\n'
    'ElseIf cboCriterioPrinc.Text = "PER\xcdODO" Then\n'
    '    REL_Cons_Venda_Prod.rfCons1.Caption = "PER\xcdODO"\n'
    '    REL_Cons_Venda_Prod.rfCons3.Caption = "Inicio/Final = " & mskInicio.Text & " at\xe9 " & mskFim.Text\n'
    'End If\n'
    '\n'
    'If cboCriterioSec.Text = "DESCRI\xc7\xc3O" Then\n'
    '    REL_Cons_Venda_Prod.rfCons2.Caption = "DESCRI\xc7\xc3O = " & cboDescricao.Text & ""\n'
    'ElseIf cboCriterioSec.Text = "C\xd3D. BARRA" Then\n'
    '    REL_Cons_Venda_Prod.rfCons2.Caption = "C\xd3D. BARRA = " & txtCodBarra.Text & ""\n'
    'ElseIf cboCriterioSec.Text = "REFER\xcaNCIA" Then\n'
    '    REL_Cons_Venda_Prod.rfCons2.Caption = "REFER\xcaNCIA = " & cboDescricao.Text & ""\n'
    'ElseIf cboCriterioSec.Text = "FABRICANTE" Then\n'
    '    REL_Cons_Venda_Prod.rfCons2.Caption = "FABRICANTE = " & cboDescricao.Text & ""\n'
    'End If\n'
)

new_block = (
    'If cboTipo.Text = "POR PRODUTOS" Then\n'
    '\n'
    '    If cboCriterioPrinc.Text = "TODOS" Then\n'
    '        REL_Cons_Venda_Prod.rfCons1.Caption = "TODOS"\n'
    '        REL_Cons_Venda_Prod.rfCons3.Caption = ""\n'
    '    ElseIf cboCriterioPrinc.Text = "MENSAL" Then\n'
    '        REL_Cons_Venda_Prod.rfCons1.Caption = "MENSAL"\n'
    '        REL_Cons_Venda_Prod.rfCons3.Caption = "M\xeas/Ano = " & cboMes.Text & "/" & cboAno.Text\n'
    '    ElseIf cboCriterioPrinc.Text = "PER\xcdODO" Then\n'
    '        REL_Cons_Venda_Prod.rfCons1.Caption = "PER\xcdODO"\n'
    '        REL_Cons_Venda_Prod.rfCons3.Caption = "Inicio/Final = " & mskInicio.Text & " at\xe9 " & mskFim.Text\n'
    '    ElseIf cboCriterioPrinc.Text = "DATA" Then\n'
    '        REL_Cons_Venda_Prod.rfCons1.Caption = "DATA"\n'
    '        REL_Cons_Venda_Prod.rfCons3.Caption = "Data = " & mskInicio.Text\n'
    '    ElseIf cboCriterioPrinc.Text = "PRODUTO/MENSAL" Then\n'
    '        REL_Cons_Venda_Prod.rfCons1.Caption = "PRODUTO/MENSAL"\n'
    '        REL_Cons_Venda_Prod.rfCons3.Caption = "M\xeas/Ano = " & cboMes.Text & "/" & cboAno.Text\n'
    '    ElseIf cboCriterioPrinc.Text = "PRODUTO/PER\xcdODO" Then\n'
    '        REL_Cons_Venda_Prod.rfCons1.Caption = "PRODUTO/PER\xcdODO"\n'
    '        REL_Cons_Venda_Prod.rfCons3.Caption = "Inicio/Final = " & mskInicio.Text & " at\xe9 " & mskFim.Text\n'
    '    End If\n'
    '\n'
    '    If cboCriterioPrinc.Text = "PRODUTO/MENSAL" Or cboCriterioPrinc.Text = "PRODUTO/PER\xcdODO" Then\n'
    '        REL_Cons_Venda_Prod.rfCons2.Caption = "C\xd3D. PRODUTO = " & txtCodProduto.Text\n'
    '    ElseIf cboCriterioSec.Text = "DESCRI\xc7\xc3O" Then\n'
    '        REL_Cons_Venda_Prod.rfCons2.Caption = "DESCRI\xc7\xc3O = " & cboDescricao.Text & ""\n'
    '    ElseIf cboCriterioSec.Text = "C\xd3D. BARRA" Then\n'
    '        REL_Cons_Venda_Prod.rfCons2.Caption = "C\xd3D. BARRA = " & txtCodBarra.Text & ""\n'
    '    ElseIf cboCriterioSec.Text = "REFER\xcaNCIA" Then\n'
    '        REL_Cons_Venda_Prod.rfCons2.Caption = "REFER\xcaNCIA = " & cboDescricao.Text & ""\n'
    '    ElseIf cboCriterioSec.Text = "FABRICANTE" Then\n'
    '        REL_Cons_Venda_Prod.rfCons2.Caption = "FABRICANTE = " & cboDescricao.Text & ""\n'
    '    ElseIf cboCriterioSec.Text = "CATEGORIA" Then\n'
    '        REL_Cons_Venda_Prod.rfCons2.Caption = "CATEGORIA = " & cboDescricao.Text & ""\n'
    '    End If\n'
    '\n'
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

count = text.count(old_block)
if count != 1:
    print(f'ERRO: {count} ocorrencias do bloco antigo (esperado 1)')
    sys.exit(1)

text = text.replace(old_block, new_block)
print('OK (1x): bloco de filtros cmdImprimir_Click atualizado')

out = text.encode('windows-1252')
out = out.replace(b'\n', b'\r\n')
with open(FILE, 'wb') as f:
    f.write(out)
print('\nArquivo gravado com sucesso.')
