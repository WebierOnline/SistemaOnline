# -*- coding: utf-8 -*-
"""
Corrige cod_servico ficando NULL em OS_Servicos_Auto:

1) INSERT (cmdAdicionarServicosAuto_Click, 2 branches) - adiciona
   cod_servico = txtCodServicoAuto.Text.
2) UPDATE (cmdEditarServicosAuto_Click) - idem.
3) MostrarGrid_Servicos - adiciona cod_servico na SELECT (as duas metades
   do UNION), para o grid conseguir devolver o valor original ao editar.
4) FormatarGrid_Servicos - nova coluna oculta (13) com cod_servico.
5) Grid_Servicos_DblClick - seta txtCodServicoAuto.Text a partir dessa
   coluna ANTES dos outros campos (para que os valores reais salvos,
   setados logo em seguida no mesmo Sub, sobrescrevam o efeito colateral
   de txtCodServicoAuto_Change que reseta preco/mecanico para o padrao
   do catalogo).
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
# 1) INSERT - 2 ocorrencias (Automoveis/Motocicletas e Informatica/Celular)
# ---------------------------------------------------------------
insert_col_old = "(codigo, cod_os, descricao, preco, quantidade, subtotal, desconto, total, data, cod_mecanico) VALUES ("
insert_col_new = "(codigo, cod_os, descricao, preco, quantidade, subtotal, desconto, total, data, cod_mecanico, cod_servico) VALUES ("
occurrences = [i for i, l in enumerate(lines) if insert_col_old in l]
assert len(occurrences) == 2, occurrences
for i in occurrences:
    lines[i] = lines[i].replace(insert_col_old, insert_col_new)

close_old = ', " & vCodMecanicoServ & ")"'
close_new = ', " & vCodMecanicoServ & ", " & txtCodServicoAuto.Text & ")"'
occ2 = [i for i, l in enumerate(lines) if l.rstrip().endswith(close_old)]
assert len(occ2) == 2, occ2
for i in occ2:
    lines[i] = lines[i][: lines[i].rfind(close_old)] + close_new

# ---------------------------------------------------------------
# 2) UPDATE (cmdEditarServicosAuto_Click)
# ---------------------------------------------------------------
i = find_line(', cod_mecanico = " & vCodMecanicoServ & " WHERE (codigo = ')
lines[i] = lines[i].replace(
    ', cod_mecanico = " & vCodMecanicoServ & " WHERE (codigo = ',
    ', cod_mecanico = " & vCodMecanicoServ & ", cod_servico = " & txtCodServicoAuto.Text & " WHERE (codigo = ',
)

# ---------------------------------------------------------------
# 3) MostrarGrid_Servicos - adiciona cod_servico na SELECT
# ---------------------------------------------------------------
start_mgs, end_mgs = find_sub("MostrarGrid_Servicos")

i = find_line(
    "cod_mecanico, (SELECT nome FROM funcionario WHERE funcionario.codigo = OS_Servicos_Auto.cod_mecanico) AS var_nomemecanico FROM OS_Servicos_Auto",
    start_mgs,
    end_mgs,
)
lines[i] = lines[i].replace(
    "AS var_nomemecanico FROM OS_Servicos_Auto",
    "AS var_nomemecanico, cod_servico FROM OS_Servicos_Auto",
)

i = find_line("NULL as cod_mecanico, NULL as var_nomemecanico", start_mgs, end_mgs)
lines[i] = lines[i].replace(
    "NULL as cod_mecanico, NULL as var_nomemecanico",
    "NULL as cod_mecanico, NULL as var_nomemecanico, NULL as cod_servico",
)

# ---------------------------------------------------------------
# 4) FormatarGrid_Servicos - coluna oculta 13 = cod_servico
# ---------------------------------------------------------------
start_fmt = find_line_exact("Private Sub FormatarGrid_Servicos(rTabela As ADODB.Recordset)")
end_fmt = find_line_exact("End Sub", start_fmt)

i = find_line(".Cols = 13", start_fmt, end_fmt)
lines[i] = lines[i].replace(".Cols = 13", ".Cols = 14")

i = find_line(".ColWidth(12) = 1300", start_fmt, end_fmt)
lines[i] = lines[i] + "\r\n       .ColWidth(13) = 0"

i = find_line('.TextMatrix(0, 12) = "MEC', start_fmt, end_fmt)
lines[i] = lines[i] + '\r\n       .TextMatrix(0, 13) = "CODSERV"'

i = find_line('.TextMatrix(.Rows - 1, 12) = ValidateNull(rTabela("var_nomemecanico"))', start_fmt, end_fmt)
lines[i] = lines[i] + '\r\n             .TextMatrix(.Rows - 1, 13) = ValidateNull(rTabela("cod_servico"))'

# ---------------------------------------------------------------
# 5) Grid_Servicos_DblClick - carregar cod_servico ANTES dos outros campos
# ---------------------------------------------------------------
start_dc, end_dc = find_sub("Grid_Servicos_DblClick")
i = find_line_exact("vCodItemServicoEditando = Grid_Servicos.TextMatrix(Grid_Servicos.Row, 9)", start_dc, end_dc)
lines[i] = lines[i] + (
    "\r\ntxtCodServicoAuto.Text = Grid_Servicos.TextMatrix(Grid_Servicos.Row, 13)"
)

# ---------------------------------------------------------------
# Grava
# ---------------------------------------------------------------
out_text = "\r\n".join(lines)
out = out_text.encode("cp1252")
with open(PATH, "wb") as f:
    f.write(out)

print("OK - patch aplicado")
