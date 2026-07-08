# -*- coding: utf-8 -*-
"""
cmdFinalizar_Click (bloco A PRAZO com entrada): a "baixa" da parcela de
entrada sobrescrevia o campo TIPO para 'PARCELA', quando deveria manter
'OS' (mesmo valor usado no INSERT e em todas as outras parcelas da OS).
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


start = find_line_exact("Private Sub cmdFinalizar_Click()")
end = find_line_exact("End Sub", start)

i = find_line_exact(
    '                      "tipo = \'PARCELA\', tipo_cartao = " & varTipoCartaoEntrada & ", " & _',
    start, end,
)
lines[i] = '                      "tipo = \'OS\', tipo_cartao = " & varTipoCartaoEntrada & ", " & _'

out_text = "\r\n".join(lines)
out = out_text.encode("cp1252")
with open(PATH, "wb") as f:
    f.write(out)

print("OK - patch aplicado")
