# -*- coding: utf-8 -*-
"""
Corrige cmdAlterar_Click: quando cboStatus.Text = "EM EXECUÇÃO", o codigo
forcava stProdSer.Visible = True (nas 4 branches de vTipoOS) - o oposto do
pedido do usuario. Esse era o verdadeiro motivo do stProdSer continuar
visivel: cmdEditarOS_Click so age na abertura da OS, mas cmdAlterar_Click
(botao "Alterar", usado ao salvar) reafirmava True toda vez que o usuario
salvava uma OS com status EM EXECUÇÃO.
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


def find_sub(name, start=0):
    s = find_line_exact(f"Private Sub {name}()", start)
    e = find_line_exact("End Sub", s)
    return s, e


start, end = find_sub("cmdAlterar_Click")

count = 0
for i in range(start, end):
    if lines[i].strip() == "stProdSer.Visible = True":
        indent = lines[i][: len(lines[i]) - len(lines[i].lstrip())]
        lines[i] = indent + "stProdSer.Visible = False"
        count += 1

assert count == 4, f"esperado 4 substituicoes, feito {count}"

out_text = "\r\n".join(lines)
out = out_text.encode("cp1252")
out = out.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

with open(PATH, "wb") as f:
    f.write(out)

print("OK - corrigido,", count, "ocorrencias trocadas para False")
