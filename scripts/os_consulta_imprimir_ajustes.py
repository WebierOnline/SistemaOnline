# -*- coding: utf-8 -*-
"""
1) OS_Consulta.frm: adiciona coluna calculada nome_completo (mesma
   concatenacao usada no grid: nome / fabricante / modelo / ano para
   veiculos; nome / equipamento / fabricante / modelo p/ demais) nas
   19 queries (18 de MostrarGrid_OS + 1 de MostrarGrid_OS_Refinado).
2) REL_OS_Consulta.frm: dfNome passa a usar Campo = "nome_completo"
   em vez de "nome".
"""

PATH_CONSULTA = r"C:\projeto\OrdemServico\Forms\OS_Consulta.frm"
PATH_REL = r"C:\projeto\Compartilhado\Forms\REL_OS_Consulta.frm"

with open(PATH_CONSULTA, "rb") as f:
    raw = f.read()
text = raw.decode("cp1252")

auto_old = "os.cod_pedido, OS_Equipamento_Auto.fabricante, OS_Equipamento_Auto.ano, OS_Equipamento_Auto.modelo, os.status AS var_status"
auto_new = (
    "os.cod_pedido, OS_Equipamento_Auto.fabricante, OS_Equipamento_Auto.ano, OS_Equipamento_Auto.modelo, "
    "(cliente.Nome + ' / ' + ISNULL(OS_Equipamento_Auto.fabricante,'') + ' / ' + ISNULL(OS_Equipamento_Auto.modelo,'') + ' / ' + ISNULL(CAST(OS_Equipamento_Auto.ano AS VARCHAR(10)),'')) AS nome_completo, "
    "os.status AS var_status"
)
n_auto = text.count(auto_old)
assert n_auto == 7, n_auto
text = text.replace(auto_old, auto_new)

info_old = "os.cod_pedido, OS_Equipamento.fabricante, OS_Equipamento.equipamento, OS_Equipamento.modelo, os.status AS var_status"
info_new = (
    "os.cod_pedido, OS_Equipamento.fabricante, OS_Equipamento.equipamento, OS_Equipamento.modelo, "
    "(cliente.Nome + ' / ' + ISNULL(OS_Equipamento.equipamento,'') + ' / ' + ISNULL(OS_Equipamento.fabricante,'') + ' / ' + ISNULL(OS_Equipamento.modelo,'')) AS nome_completo, "
    "os.status AS var_status"
)
n_info = text.count(info_old)
assert n_info == 12, n_info
text = text.replace(info_old, info_new)

text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH_CONSULTA, "wb") as f:
    f.write(text.encode("cp1252"))
print(f"OK - OS_Consulta.frm: {n_auto} (auto) + {n_info} (info/comvisual) = {n_auto+n_info} queries atualizadas")

# ---------------------------------------------------------------
# REL_OS_Consulta.frm: dfNome Campo = "nome_completo"
# ---------------------------------------------------------------
with open(PATH_REL, "rb") as f:
    raw2 = f.read()
text2 = raw2.decode("cp1252")
lines2 = text2.split("\r\n")


def find_line_exact(lines, s, start=0, end=None):
    end = end if end is not None else len(lines)
    for i in range(start, end):
        if lines[i] == s:
            return i
    raise SystemExit(f"ERRO: linha exata nao encontrada: {s!r}")


i = find_line_exact(lines2, "      Begin ReportX.ReportField dfNome ")
j = find_line_exact(lines2, '         Campo           =   "nome"', i, i + 15)
lines2[j] = '         Campo           =   "nome_completo"'

out2 = "\r\n".join(lines2)
out2 = out2.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH_REL, "wb") as f:
    f.write(out2.encode("cp1252"))
print("OK - REL_OS_Consulta.frm: dfNome.Campo = nome_completo")
