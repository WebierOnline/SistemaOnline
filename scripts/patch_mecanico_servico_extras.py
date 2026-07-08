# -*- coding: utf-8 -*-
"""
Patch OS_Recapadora.frm - 3 pedidos do usuario:

1) txtCodServicoAuto_Change: ao escolher um servico do catalogo,
   cboMecanicoServ passa a copiar o nome/codigo de cboMecanico/txtCodMecanico
   (o "responsavel" da OS). Se cboMecanico estiver vazio, cboMecanicoServ
   fica em branco.

2) MostrarGrid_Servicos / FormatarGrid_Servicos: acrescenta uma nova
   coluna VISIVEL (indice 12, no final) com o NOME do mecanico. Como as
   colunas ocultas atuais (9=ITEM,10=PROD,11=cod_mecanico) tem largura 0,
   essa nova coluna aparece visualmente logo apos "Total", sem precisar
   deslocar nenhum indice usado em outras subs (remover, ajustar estoque
   etc.) - risco zero de quebrar codigo existente.

3) LimparObjetos_ServicosAuto: agora tambem limpa cboMecanicoServ/
   vCodMecanicoServ - jah eh chamada por cmdAdicionarServicosAuto_Click e
   cmdEditarServicosAuto_Click, entao os dois ganham esse comportamento.
"""

PATH = r"C:\projeto\OrdemServico\Forms\OS_Recapadora.frm"

with open(PATH, "rb") as f:
    raw = f.read()

text = raw.decode("cp1252")
lines = text.split("\r\n")


def find_line_exact(s, start=0, end=None):
    end = end if end is not None else len(lines)
    for i in range(start, end):
        if lines[i] == s:
            return i
    raise SystemExit(f"ERRO: linha exata nao encontrada: {s!r}")


def find_line(substr, start=0, end=None):
    end = end if end is not None else len(lines)
    for i in range(start, end):
        if substr in lines[i]:
            return i
    raise SystemExit(f"ERRO: ancora nao encontrada: {substr!r}")


def find_sub(name, start=0):
    s = find_line_exact(f"Private Sub {name}()", start)
    e = find_line_exact("End Sub", s)
    return s, e


# ---------------------------------------------------------------
# 1) txtCodServicoAuto_Change - copiar cboMecanico -> cboMecanicoServ
# ---------------------------------------------------------------
start, end = find_sub("txtCodServicoAuto_Change")
i = find_line_exact("    Set r = Nothing", start, end)
lines[i] = lines[i] + (
    "\r\n"
    'If cboMecanico.Text = "" Then\r\n'
    '    cboMecanicoServ.Text = ""\r\n'
    '    vCodMecanicoServ = ""\r\n'
    "Else\r\n"
    "    cboMecanicoServ.Text = cboMecanico.Text\r\n"
    "    vCodMecanicoServ = txtCodMecanico.Text\r\n"
    "End If"
)

# ---------------------------------------------------------------
# 2a) MostrarGrid_Servicos - adiciona subquery do nome do mecanico
# ---------------------------------------------------------------
start_mgs, end_mgs = find_sub("MostrarGrid_Servicos")

i = find_line(
    "CODIGO AS var_CODITEM, '' as var_CODPROD, cod_mecanico FROM OS_Servicos_Auto",
    start_mgs,
    end_mgs,
)
lines[i] = lines[i].replace(
    "CODIGO AS var_CODITEM, '' as var_CODPROD, cod_mecanico FROM OS_Servicos_Auto",
    "CODIGO AS var_CODITEM, '' as var_CODPROD, cod_mecanico, "
    "(SELECT nome FROM funcionario WHERE funcionario.codigo = OS_Servicos_Auto.cod_mecanico) AS var_nomemecanico "
    "FROM OS_Servicos_Auto",
)

i = find_line("pedidos_itens.COD_PRODUTO as var_CODPROD, NULL as cod_mecanico", start_mgs, end_mgs)
lines[i] = lines[i].replace(
    "pedidos_itens.COD_PRODUTO as var_CODPROD, NULL as cod_mecanico",
    "pedidos_itens.COD_PRODUTO as var_CODPROD, NULL as cod_mecanico, NULL as var_nomemecanico",
)

# ---------------------------------------------------------------
# 2b) FormatarGrid_Servicos - nova coluna visivel 12 = nome do mecanico
# ---------------------------------------------------------------
start_fmt = find_line_exact("Private Sub FormatarGrid_Servicos(rTabela As ADODB.Recordset)")
end_fmt = find_line_exact("End Sub", start_fmt)

i = find_line(".Cols = 12", start_fmt, end_fmt)
lines[i] = lines[i].replace(".Cols = 12", ".Cols = 13")

i = find_line(".ColWidth(11) = 0", start_fmt, end_fmt)
lines[i] = lines[i] + '\r\n       .ColWidth(12) = 1300'

i = find_line('.TextMatrix(0, 11) = "MEC"', start_fmt, end_fmt)
lines[i] = lines[i] + '\r\n       .TextMatrix(0, 12) = "MECÂNICO"'

i = find_line('.TextMatrix(.Rows - 1, 11) = ValidateNull(rTabela("cod_mecanico"))', start_fmt, end_fmt)
lines[i] = lines[i] + '\r\n             .TextMatrix(.Rows - 1, 12) = ValidateNull(rTabela("var_nomemecanico"))'

# ---------------------------------------------------------------
# 3) LimparObjetos_ServicosAuto - limpar cboMecanicoServ/vCodMecanicoServ
# ---------------------------------------------------------------
i = find_line_exact('txtObsServ.Text = ""')
lines[i] = lines[i] + '\r\ncboMecanicoServ.Text = ""\r\nvCodMecanicoServ = ""'

# ---------------------------------------------------------------
# Grava
# ---------------------------------------------------------------
out_text = "\r\n".join(lines)
out = out_text.encode("cp1252")
with open(PATH, "wb") as f:
    f.write(out)

print("OK - patch aplicado")
