"""
patch_chkDadosBancarios_v3.py

Corrige Error 5 no uncheck: VB6 And nao faz short-circuit, entao
Mid(sTmp, iPos-2, 2) era avaliado mesmo com iPos=1, causando argumento invalido.
Substituido por If aninhado para evitar o Mid quando iPos < 3.
"""
import sys, shutil

FRM = r"C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm"
BAK = FRM + ".bak_chkDadosBancarios_v3"

shutil.copy2(FRM, BAK)
print(f"Backup: {BAK}")

with open(FRM, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

OLD = (
    b"            If iPos > 0 Then\r\n"
    b"                If iPos >= 3 And Mid(sTmp, iPos - 2, 2) = vbCrLf Then\r\n"
    b"                    txtInfComple.Text = Trim(Left(sTmp, iPos - 3))\r\n"
    b"                Else\r\n"
    b"                    txtInfComple.Text = \"\"\r\n"
    b"                End If\r\n"
    b"            End If\r\n"
)

NEW = (
    b"            If iPos > 0 Then\r\n"
    b"                If iPos >= 3 Then\r\n"
    b"                    If Mid(sTmp, iPos - 2, 2) = vbCrLf Then\r\n"
    b"                        txtInfComple.Text = Trim(Left(sTmp, iPos - 3))\r\n"
    b"                    Else\r\n"
    b"                        txtInfComple.Text = \"\"\r\n"
    b"                    End If\r\n"
    b"                Else\r\n"
    b"                    txtInfComple.Text = \"\"\r\n"
    b"                End If\r\n"
    b"            End If\r\n"
)

cnt = data.count(OLD)
if cnt != 1:
    print(f"ERRO: trecho encontrado {cnt}x (esperado 1). Arquivo NAO alterado.")
    sys.exit(1)

data = data.replace(OLD, NEW)
data = norm(data)

with open(FRM, "wb") as f:
    f.write(data)

print("OK: If aninhado substitui And para evitar Mid com argumento invalido")
print("Arquivo salvo com sucesso.")
