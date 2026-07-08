# -*- coding: cp1252 -*-
"""
Patch Funcionario_Cadastro.frm:
- Renomeia controles do Frame7 "Vendas a prazo" para o grupo AP
  (txtComAPAlvo1/2/3, txtComAP1/2/3)
- Renomeia controles do Frame8 "Servicos" para o grupo Serv
  (txtComServAlvo1/2/3, txtComServ1/2/3)
- Estende Inserir_Dados / Atualizar_Dados / Campos_Brancos / Mostrar_Dados
  para os tiers 2 e 3 (novos campos no banco)
- Renomeia campo Valor_Comissao1/2/3 -> Valor_ComissaoAV1/2/3
- Adiciona os LostFocus handlers que faltam
"""
import re

PATH = r"C:\projeto\OnlineCommerce\Forms\Funcionario_Cadastro.frm"

with open(PATH, "rb") as f:
    raw = f.read()

text = raw.decode("cp1252")
original = text


def replace_once(old, new):
    global text
    n = text.count(old)
    if n != 1:
        raise SystemExit(f"ERRO: esperado 1 ocorrencia, encontrado {n}: {old!r}")
    text = text.replace(old, new, 1)


def replace_all(old, new, expected):
    global text
    n = text.count(old)
    if n != expected:
        raise SystemExit(f"ERRO: esperado {expected} ocorrencias, encontrado {n}: {old!r}")
    text = text.replace(old, new)


# ---------------------------------------------------------------
# 1) Renomeia declaracoes de controles (Frame7 -> grupo AP)
# ---------------------------------------------------------------
replace_once("Begin VB.TextBox Text2 \r\n", "Begin VB.TextBox txtComAPAlvo1 \r\n")
replace_once("Begin VB.TextBox Text4 \r\n", "Begin VB.TextBox txtComAPAlvo2 \r\n")
replace_once("Begin VB.TextBox Text3 \r\n", "Begin VB.TextBox txtComAP2 \r\n")
replace_once("Begin VB.TextBox Text6 \r\n", "Begin VB.TextBox txtComAPAlvo3 \r\n")
replace_once("Begin VB.TextBox Text5 \r\n", "Begin VB.TextBox txtComAP3 \r\n")

# ---------------------------------------------------------------
# 2) Renomeia declaracoes de controles (Frame8 -> grupo Serv)
# ---------------------------------------------------------------
replace_once("Begin VB.TextBox Text8 \r\n", "Begin VB.TextBox txtComServAlvo1 \r\n")
replace_once("Begin VB.TextBox Text10 \r\n", "Begin VB.TextBox txtComServAlvo2 \r\n")
replace_once("Begin VB.TextBox Text9 \r\n", "Begin VB.TextBox txtComServ2 \r\n")
replace_once("Begin VB.TextBox Text12 \r\n", "Begin VB.TextBox txtComServAlvo3 \r\n")
replace_once("Begin VB.TextBox Text11 \r\n", "Begin VB.TextBox txtComServ3 \r\n")

# ---------------------------------------------------------------
# 3) Renomeia todos os usos (declaracao + codigo + LostFocus sub name)
#    txtComPrazo1 -> txtComAP1, txtComServicos1 -> txtComServ1
# ---------------------------------------------------------------
replace_all("txtComPrazo1", "txtComAP1", 10)
replace_all("txtComServicos1", "txtComServ1", 10)

# ---------------------------------------------------------------
# 4) INSERT (Inserir_Dados) - lista de colunas
# ---------------------------------------------------------------
replace_once(
    "Comissao_Avista1, Comissao_Avista2, Comissao_Avista3, Valor_Comissao1, Valor_Comissao2, Valor_Comissao3, "
    "Comissao_Recebido1, Comissao_Recebido2, Comissao_Recebido3, Valor_ComissaoRec1, Valor_ComissaoRec2, Valor_ComissaoRec3, "
    "Comissao_Prazo1, Comissao_Servico1) VALUES (",
    "Comissao_Avista1, Comissao_Avista2, Comissao_Avista3, Valor_ComissaoAV1, Valor_ComissaoAV2, Valor_ComissaoAV3, "
    "Comissao_Recebido1, Comissao_Recebido2, Comissao_Recebido3, Valor_ComissaoRec1, Valor_ComissaoRec2, Valor_ComissaoRec3, "
    "Comissao_Prazo1, Comissao_Prazo2, Comissao_Prazo3, Valor_ComissaoAP1, Valor_ComissaoAP2, Valor_ComissaoAP3, "
    "Comissao_Servico1, Comissao_Servico2, Comissao_Servico3, Valor_ComissaoServ1, Valor_ComissaoServ2, Valor_ComissaoServ3) VALUES (",
)

# ---------------------------------------------------------------
# 5) INSERT (Inserir_Dados) - lista de valores (fim da VALUES)
#    quebrada em varias linhas fisicas (continuacao "& _") para nao
#    estourar o limite de ~1023 caracteres por linha do VB6
# ---------------------------------------------------------------
replace_once(
    '" & Replace(CDbl(txtComAP1.Text), ",", ".") & ", " & Replace(CDbl(txtComServ1.Text), ",", ".") & ");"',
    '" & Replace(CDbl(txtComAP1.Text), ",", ".") & ", " & Replace(CDbl(txtComAP2.Text), ",", ".") & ", " & Replace(CDbl(txtComAP3.Text), ",", ".") & ", " & _\r\n'
    '      Replace(CDbl(txtComAPAlvo1.Text), ",", ".") & ", " & Replace(CDbl(txtComAPAlvo2.Text), ",", ".") & ", " & Replace(CDbl(txtComAPAlvo3.Text), ",", ".") & ", " & _\r\n'
    '      Replace(CDbl(txtComServ1.Text), ",", ".") & ", " & Replace(CDbl(txtComServ2.Text), ",", ".") & ", " & Replace(CDbl(txtComServ3.Text), ",", ".") & ", " & _\r\n'
    '      Replace(CDbl(txtComServAlvo1.Text), ",", ".") & ", " & Replace(CDbl(txtComServAlvo2.Text), ",", ".") & ", " & Replace(CDbl(txtComServAlvo3.Text), ",", ".") & ");"',
)

# ---------------------------------------------------------------
# 6) UPDATE (Atualizar_Dados) - Comissao_Prazo1 / Comissao_Servico1 + Valor_Comissao1/2/3
#    idem, quebrada em varias linhas fisicas
# ---------------------------------------------------------------
replace_once(
    'Comissao_Prazo1 = " & Replace(CDbl(txtComAP1.Text), ",", ".") & ", Comissao_Servico1 = " & Replace(CDbl(txtComServ1.Text), ",", ".") & ", " & _',
    'Comissao_Prazo1 = " & Replace(CDbl(txtComAP1.Text), ",", ".") & ", Comissao_Prazo2 = " & Replace(CDbl(txtComAP2.Text), ",", ".") & ", Comissao_Prazo3 = " & Replace(CDbl(txtComAP3.Text), ",", ".") & ", " & _\r\n'
    '      "Valor_ComissaoAP1 = " & Replace(CDbl(txtComAPAlvo1.Text), ",", ".") & ", Valor_ComissaoAP2 = " & Replace(CDbl(txtComAPAlvo2.Text), ",", ".") & ", Valor_ComissaoAP3 = " & Replace(CDbl(txtComAPAlvo3.Text), ",", ".") & ", " & _\r\n'
    '      "Comissao_Servico1 = " & Replace(CDbl(txtComServ1.Text), ",", ".") & ", Comissao_Servico2 = " & Replace(CDbl(txtComServ2.Text), ",", ".") & ", Comissao_Servico3 = " & Replace(CDbl(txtComServ3.Text), ",", ".") & ", " & _\r\n'
    '      "Valor_ComissaoServ1 = " & Replace(CDbl(txtComServAlvo1.Text), ",", ".") & ", Valor_ComissaoServ2 = " & Replace(CDbl(txtComServAlvo2.Text), ",", ".") & ", Valor_ComissaoServ3 = " & Replace(CDbl(txtComServAlvo3.Text), ",", ".") & ", " & _',
)

replace_once(
    'Valor_Comissao1 = " & Replace(CDbl(txtComVistaAlvo1.Text), ",", ".") & ", Valor_Comissao2 = " & Replace(CDbl(txtComVistaAlvo2.Text), ",", ".") & ", Valor_Comissao3 = " & Replace(CDbl(txtComVistaAlvo3.Text), ",", ".") & ", " & _',
    'Valor_ComissaoAV1 = " & Replace(CDbl(txtComVistaAlvo1.Text), ",", ".") & ", Valor_ComissaoAV2 = " & Replace(CDbl(txtComVistaAlvo2.Text), ",", ".") & ", Valor_ComissaoAV3 = " & Replace(CDbl(txtComVistaAlvo3.Text), ",", ".") & ", " & _',
)

# ---------------------------------------------------------------
# 7) Campos_Brancos - limpar campos
# ---------------------------------------------------------------
replace_once(
    "txtComAP1.Text = Format(0, ocMONEY)\r\ntxtComServ1.Text = Format(0, ocMONEY)\r\n",
    "txtComAP1.Text = Format(0, ocMONEY)\r\n"
    "txtComAP2.Text = Format(0, ocMONEY)\r\n"
    "txtComAP3.Text = Format(0, ocMONEY)\r\n"
    "txtComAPAlvo1.Text = Format(0, ocMONEY)\r\n"
    "txtComAPAlvo2.Text = Format(0, ocMONEY)\r\n"
    "txtComAPAlvo3.Text = Format(0, ocMONEY)\r\n"
    "txtComServ1.Text = Format(0, ocMONEY)\r\n"
    "txtComServ2.Text = Format(0, ocMONEY)\r\n"
    "txtComServ3.Text = Format(0, ocMONEY)\r\n"
    "txtComServAlvo1.Text = Format(0, ocMONEY)\r\n"
    "txtComServAlvo2.Text = Format(0, ocMONEY)\r\n"
    "txtComServAlvo3.Text = Format(0, ocMONEY)\r\n",
)

# ---------------------------------------------------------------
# 8) Mostrar_Dados - Valor_Comissao1/2/3 -> Valor_ComissaoAV1/2/3
# ---------------------------------------------------------------
replace_once(
    'txtComVistaAlvo1.Text = Format(ValidateNull(rTabela("Valor_Comissao1")), ocMONEY)\r\n'
    'txtComVistaAlvo2.Text = Format(ValidateNull(rTabela("Valor_Comissao2")), ocMONEY)\r\n'
    'txtComVistaAlvo3.Text = Format(ValidateNull(rTabela("Valor_Comissao3")), ocMONEY)\r\n',
    'txtComVistaAlvo1.Text = Format(ValidateNull(rTabela("Valor_ComissaoAV1")), ocMONEY)\r\n'
    'txtComVistaAlvo2.Text = Format(ValidateNull(rTabela("Valor_ComissaoAV2")), ocMONEY)\r\n'
    'txtComVistaAlvo3.Text = Format(ValidateNull(rTabela("Valor_ComissaoAV3")), ocMONEY)\r\n',
)

# ---------------------------------------------------------------
# 9) Mostrar_Dados - Comissao_Prazo1 / Comissao_Servico1 + novos tiers
# ---------------------------------------------------------------
replace_once(
    'txtComAP1.Text = Format(ValidateNull(rTabela("Comissao_Prazo1")), ocMONEY)\r\n'
    "\r\n"
    'txtComServ1.Text = Format(ValidateNull(rTabela("Comissao_Servico1")), ocMONEY)\r\n',
    'txtComAP1.Text = Format(ValidateNull(rTabela("Comissao_Prazo1")), ocMONEY)\r\n'
    'txtComAP2.Text = Format(ValidateNull(rTabela("Comissao_Prazo2")), ocMONEY)\r\n'
    'txtComAP3.Text = Format(ValidateNull(rTabela("Comissao_Prazo3")), ocMONEY)\r\n'
    'txtComAPAlvo1.Text = Format(ValidateNull(rTabela("Valor_ComissaoAP1")), ocMONEY)\r\n'
    'txtComAPAlvo2.Text = Format(ValidateNull(rTabela("Valor_ComissaoAP2")), ocMONEY)\r\n'
    'txtComAPAlvo3.Text = Format(ValidateNull(rTabela("Valor_ComissaoAP3")), ocMONEY)\r\n'
    "\r\n"
    'txtComServ1.Text = Format(ValidateNull(rTabela("Comissao_Servico1")), ocMONEY)\r\n'
    'txtComServ2.Text = Format(ValidateNull(rTabela("Comissao_Servico2")), ocMONEY)\r\n'
    'txtComServ3.Text = Format(ValidateNull(rTabela("Comissao_Servico3")), ocMONEY)\r\n'
    'txtComServAlvo1.Text = Format(ValidateNull(rTabela("Valor_ComissaoServ1")), ocMONEY)\r\n'
    'txtComServAlvo2.Text = Format(ValidateNull(rTabela("Valor_ComissaoServ2")), ocMONEY)\r\n'
    'txtComServAlvo3.Text = Format(ValidateNull(rTabela("Valor_ComissaoServ3")), ocMONEY)\r\n',
)

# ---------------------------------------------------------------
# 10) LostFocus handlers - substitui os 2 subs existentes (agora AP1/Serv1)
#     por blocos completos com os tiers 2 e 3 e os campos Alvo
# ---------------------------------------------------------------
def lostfocus_block(name):
    return (
        f"Private Sub {name}_LostFocus()\r\n"
        f'If {name}.Text = "" Then\r\n'
        f"   {name}.Text = Format(0, ocMONEY)\r\n"
        f"Else\r\n"
        f"   {name}.Text = Format({name}, ocMONEY)\r\n"
        f"End If\r\n"
        f"End Sub\r\n"
    )


ap_names = ["txtComAP1", "txtComAP2", "txtComAP3", "txtComAPAlvo1", "txtComAPAlvo2", "txtComAPAlvo3"]
serv_names = ["txtComServ1", "txtComServ2", "txtComServ3", "txtComServAlvo1", "txtComServAlvo2", "txtComServAlvo3"]

old_ap_sub = lostfocus_block("txtComAP1")
new_ap_subs = "\r\n\r\n".join(lostfocus_block(n) for n in ap_names)
replace_once(old_ap_sub, new_ap_subs)

old_serv_sub = lostfocus_block("txtComServ1")
new_serv_subs = "\r\n\r\n".join(lostfocus_block(n) for n in serv_names)
replace_once(old_serv_sub, new_serv_subs)

# ---------------------------------------------------------------
# Grava o arquivo normalizando quebras de linha
# ---------------------------------------------------------------
out = text.encode("cp1252")
out = out.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

with open(PATH, "wb") as f:
    f.write(out)

print("OK - patch aplicado")
print("bytes originais:", len(raw), "bytes finais:", len(out))
