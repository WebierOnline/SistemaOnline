# -*- coding: utf-8 -*-
"""
Patch v3: POR SERVIÇOS/PRODUTOS — recarrega combos ao mudar tipo,
adiciona 'TODOS' em cboCriterioSec, remove 'TODOS' de cboCriterioPrinc
quando cboCriterioSec = TODOS; trata visibilidade dos campos.
"""
import sys

FILE = r'C:\Projeto\OnlineCommerce\Forms\Vendas_Consulta_PorProdutos.frm'

with open(FILE, 'rb') as f:
    raw = f.read()

raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n')
text = raw.decode('windows-1252')

changes = []

# ------------------------------------------------------------------
# 1. cboCriterioSec_GotFocus — adiciona "TODOS" como primeiro item
# ------------------------------------------------------------------
old1 = (
    "Private Sub cboCriterioSec_GotFocus()\n"
    "cboCriterioSec.Clear\n"
    "\n"
    "If cboTipo.Text = \"POR SERVIÇOS\" Then\n"
    "   cboCriterioSec.AddItem \"DESCRIÇÃO\"\n"
    "   cboCriterioSec.AddItem \"CÓD. OS\"\n"
    "Else\n"
    "   cboCriterioSec.AddItem \"DESCRIÇÃO\"\n"
    "   cboCriterioSec.AddItem \"CÓD. BARRA\"\n"
    "   cboCriterioSec.AddItem \"REFERÊNCIA\"\n"
    "   cboCriterioSec.AddItem \"FABRICANTE\"\n"
    "   cboCriterioSec.AddItem \"CATEGORIA\"\n"
    "End If\n"
    "\n"
    "moCombo.AttachTo cboCriterioSec\n"
    "End Sub\n"
)
new1 = (
    "Private Sub cboCriterioSec_GotFocus()\n"
    "cboCriterioSec.Clear\n"
    "\n"
    "If cboTipo.Text = \"POR SERVIÇOS\" Then\n"
    "   cboCriterioSec.AddItem \"TODOS\"\n"
    "   cboCriterioSec.AddItem \"DESCRIÇÃO\"\n"
    "   cboCriterioSec.AddItem \"CÓD. OS\"\n"
    "Else\n"
    "   cboCriterioSec.AddItem \"TODOS\"\n"
    "   cboCriterioSec.AddItem \"DESCRIÇÃO\"\n"
    "   cboCriterioSec.AddItem \"CÓD. BARRA\"\n"
    "   cboCriterioSec.AddItem \"REFERÊNCIA\"\n"
    "   cboCriterioSec.AddItem \"FABRICANTE\"\n"
    "   cboCriterioSec.AddItem \"CATEGORIA\"\n"
    "End If\n"
    "\n"
    "moCombo.AttachTo cboCriterioSec\n"
    "End Sub\n"
)
changes.append((old1, new1, '1 - TODOS em cboCriterioSec_GotFocus'))

# ------------------------------------------------------------------
# 2. cboCriterioPrinc_GotFocus — "TODOS" condicional ao cboCriterioSec
# ------------------------------------------------------------------
old2 = (
    "Private Sub cboCriterioPrinc_GotFocus()\n"
    "cboCriterioPrinc.Clear\n"
    "   \n"
    "If cboTipo.Text = \"POR PRODUTOS\" Then\n"
    "   cboCriterioPrinc.AddItem \"TODOS\"\n"
    "   cboCriterioPrinc.AddItem \"MENSAL\"\n"
    "   cboCriterioPrinc.AddItem \"PERÍODO\"\n"
    "   cboCriterioPrinc.AddItem \"DATA\"\n"
    "   cboCriterioPrinc.AddItem \"PRODUTO/MENSAL\"\n"
    "   cboCriterioPrinc.AddItem \"PRODUTO/PERÍODO\"\n"
    "ElseIf cboTipo.Text = \"POR SERVIÇOS\" Then\n"
    "   cboCriterioPrinc.AddItem \"TODOS\"\n"
    "   cboCriterioPrinc.AddItem \"MENSAL\"\n"
    "   cboCriterioPrinc.AddItem \"SERVIÇOS\"\n"
    "   cboCriterioPrinc.AddItem \"SERVIÇOS/MENSAL\"\n"
    "   cboCriterioPrinc.AddItem \"SERVIÇOS/PERÍODO\"\n"
    "End If\n"
    "   \n"
    "moCombo.AttachTo cboCriterioPrinc\n"
    "End Sub\n"
)
new2 = (
    "Private Sub cboCriterioPrinc_GotFocus()\n"
    "cboCriterioPrinc.Clear\n"
    "   \n"
    "If cboTipo.Text = \"POR PRODUTOS\" Then\n"
    "   If cboCriterioSec.Text <> \"TODOS\" Then cboCriterioPrinc.AddItem \"TODOS\"\n"
    "   cboCriterioPrinc.AddItem \"MENSAL\"\n"
    "   cboCriterioPrinc.AddItem \"PERÍODO\"\n"
    "   cboCriterioPrinc.AddItem \"DATA\"\n"
    "   cboCriterioPrinc.AddItem \"PRODUTO/MENSAL\"\n"
    "   cboCriterioPrinc.AddItem \"PRODUTO/PERÍODO\"\n"
    "ElseIf cboTipo.Text = \"POR SERVIÇOS\" Then\n"
    "   If cboCriterioSec.Text <> \"TODOS\" Then cboCriterioPrinc.AddItem \"TODOS\"\n"
    "   cboCriterioPrinc.AddItem \"MENSAL\"\n"
    "   cboCriterioPrinc.AddItem \"SERVIÇOS\"\n"
    "   cboCriterioPrinc.AddItem \"SERVIÇOS/MENSAL\"\n"
    "   cboCriterioPrinc.AddItem \"SERVIÇOS/PERÍODO\"\n"
    "End If\n"
    "   \n"
    "moCombo.AttachTo cboCriterioPrinc\n"
    "End Sub\n"
)
changes.append((old2, new2, '2 - TODOS condicional em cboCriterioPrinc_GotFocus'))

# ------------------------------------------------------------------
# 3. cboTipo_Change — recarrega cboCriterioSec e cboCriterioPrinc
# ------------------------------------------------------------------
old3 = (
    "Private Sub cboTipo_Change()\n"
    "If cboTipo.Text = \"POR PRODUTOS\" Then\n"
    "'cmdExibirParcelas.Visible = False\n"
    "   cmdExibirPedidos.Visible = True\n"
    "ElseIf cboTipo.Text = \"POR SERVIÇOS\" Then\n"
    "'cmdExibirParcelas.Visible = False\n"
    "   cmdExibirPedidos.Visible = True\n"
    "Else\n"
    "   Exit Sub\n"
    "End If\n"
    "End Sub\n"
)
new3 = (
    "Private Sub cboTipo_Change()\n"
    "If cboTipo.Text = \"POR PRODUTOS\" Then\n"
    "'cmdExibirParcelas.Visible = False\n"
    "   cmdExibirPedidos.Visible = True\n"
    "ElseIf cboTipo.Text = \"POR SERVIÇOS\" Then\n"
    "'cmdExibirParcelas.Visible = False\n"
    "   cmdExibirPedidos.Visible = True\n"
    "Else\n"
    "   Exit Sub\n"
    "End If\n"
    "\n"
    "' Recarrega cboCriterioSec e seleciona TODOS\n"
    "cboCriterioSec.Clear\n"
    "If cboTipo.Text = \"POR SERVIÇOS\" Then\n"
    "   cboCriterioSec.AddItem \"TODOS\"\n"
    "   cboCriterioSec.AddItem \"DESCRIÇÃO\"\n"
    "   cboCriterioSec.AddItem \"CÓD. OS\"\n"
    "Else\n"
    "   cboCriterioSec.AddItem \"TODOS\"\n"
    "   cboCriterioSec.AddItem \"DESCRIÇÃO\"\n"
    "   cboCriterioSec.AddItem \"CÓD. BARRA\"\n"
    "   cboCriterioSec.AddItem \"REFERÊNCIA\"\n"
    "   cboCriterioSec.AddItem \"FABRICANTE\"\n"
    "   cboCriterioSec.AddItem \"CATEGORIA\"\n"
    "End If\n"
    "cboCriterioSec.ListIndex = 0\n"
    "\n"
    "' Recarrega cboCriterioPrinc sem TODOS (cboCriterioSec = TODOS)\n"
    "cboCriterioPrinc.Clear\n"
    "If cboTipo.Text = \"POR PRODUTOS\" Then\n"
    "   cboCriterioPrinc.AddItem \"MENSAL\"\n"
    "   cboCriterioPrinc.AddItem \"PERÍODO\"\n"
    "   cboCriterioPrinc.AddItem \"DATA\"\n"
    "   cboCriterioPrinc.AddItem \"PRODUTO/MENSAL\"\n"
    "   cboCriterioPrinc.AddItem \"PRODUTO/PERÍODO\"\n"
    "ElseIf cboTipo.Text = \"POR SERVIÇOS\" Then\n"
    "   cboCriterioPrinc.AddItem \"MENSAL\"\n"
    "   cboCriterioPrinc.AddItem \"SERVIÇOS\"\n"
    "   cboCriterioPrinc.AddItem \"SERVIÇOS/MENSAL\"\n"
    "   cboCriterioPrinc.AddItem \"SERVIÇOS/PERÍODO\"\n"
    "End If\n"
    "cboCriterioPrinc.ListIndex = 0\n"
    "\n"
    "cboCriterioSec_LostFocus\n"
    "cboCriterioPrinc_LostFocus\n"
    "End Sub\n"
)
changes.append((old3, new3, '3 - cboTipo_Change recarrega combos'))

# ------------------------------------------------------------------
# 4. cboCriterioSec_LostFocus — trata "TODOS" (oculta controles)
# ------------------------------------------------------------------
old4 = (
    "ElseIf cboCriterioSec.Text = \"CÓD. OS\" Then\n"
    "    lblDescricao.Caption = \"Cód. OS\"\n"
    "    lblDescricao.Visible = True\n"
    "    cboDescricao.Visible = False\n"
    "    txtCodBarra.Visible = True\n"
    "Else\n"
    "End If\n"
    "End Sub\n"
)
new4 = (
    "ElseIf cboCriterioSec.Text = \"CÓD. OS\" Then\n"
    "    lblDescricao.Caption = \"Cód. OS\"\n"
    "    lblDescricao.Visible = True\n"
    "    cboDescricao.Visible = False\n"
    "    txtCodBarra.Visible = True\n"
    "ElseIf cboCriterioSec.Text = \"TODOS\" Then\n"
    "    lblDescricao.Visible = False\n"
    "    cboDescricao.Visible = False\n"
    "    txtCodBarra.Visible = False\n"
    "Else\n"
    "End If\n"
    "End Sub\n"
)
changes.append((old4, new4, '4 - TODOS em cboCriterioSec_LostFocus'))

# ------------------------------------------------------------------
# 5. cboCriterioPrinc_LostFocus segundo bloco — trata "TODOS"
# ------------------------------------------------------------------
old5 = (
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
new5 = (
    "ElseIf cboCriterioSec.Text = \"CÓD. OS\" Then\n"
    "    lblDescricao.Caption = \"Cód. OS\"\n"
    "    lblDescricao.Visible = True\n"
    "    cboDescricao.Visible = False\n"
    "    txtCodBarra.Visible = True\n"
    "ElseIf cboCriterioSec.Text = \"TODOS\" Then\n"
    "    lblDescricao.Visible = False\n"
    "    cboDescricao.Visible = False\n"
    "    txtCodBarra.Visible = False\n"
    "Else\n"
    "End If\n"
    "\n"
    "\n"
    "LimparObjetos_Consulta\n"
    "End Sub\n"
)
changes.append((old5, new5, '5 - TODOS em cboCriterioPrinc_LostFocus segundo bloco'))

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
