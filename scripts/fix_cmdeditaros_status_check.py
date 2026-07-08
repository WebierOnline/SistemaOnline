# -*- coding: utf-8 -*-
"""
Corrige cmdEditarOS_Click: a checagem de EM EXECUÇÃO deve usar
cboStatus.Text (ja carregado nesse ponto, pois txtCodOS.Text foi setado
logo acima e disparou txtCodOS_Change -> Mostrar_Entrada -> cboStatus.Text),
em vez de Grid_OS.TextMatrix(posit, 1) — coluna cujo indice de status varia
entre os branches de vTipoOS (as vezes é a coluna 1, as vezes a 2), tornando
essa leitura pouco confiavel.
"""

PATH = r"C:\projeto\OrdemServico\Forms\OS_Recapadora.frm"

with open(PATH, "rb") as f:
    raw = f.read()

text = raw.decode("cp1252")

OLD = 'ElseIf (Trim(Grid_OS.TextMatrix(posit, 1))) = ("EM EXECUÇÃO") Then'
NEW = 'ElseIf cboStatus.Text = "EM EXECUÇÃO" Then'

assert text.count(OLD) == 1, text.count(OLD)
text = text.replace(OLD, NEW, 1)

out = text.encode("cp1252")
out = out.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

with open(PATH, "wb") as f:
    f.write(out)

print("OK - corrigido")
print("bytes originais:", len(raw), "bytes finais:", len(out))
