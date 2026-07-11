# -*- coding: utf-8 -*-
"""
cmdImprimir_Click (OS_Consulta.frm): preenche REL_OS_Consulta.rfData com
data e hora da impressao, formato "Data: 11/07/26 as 17:18hs".
"""

PATH = r"C:\projeto\OrdemServico\Forms\OS_Consulta.frm"

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


anchor = find_line_exact("REL_OS_Consulta.lblTitulo.Caption = \"RELAT\xd3RIO - CONSULTA DE ORDEM DE SERVI\xc7OS\"")

new_line = "REL_OS_Consulta.rfData.Caption = \"Data: \" & Format(Now, \"dd/mm/yy\") & \" \xe0s \" & Format(Now, \"hh:nn\") & \"hs\""

lines.insert(anchor + 1, new_line)

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print(f"OK - rfData.Caption inserido apos linha {anchor}")
