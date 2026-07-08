"""
patch_difal_locale_fix.py

Corrige parsing de aliquotas em frmDIFAL_Cadastro.frm:
CDbl(Replace(s, ",", ".")) em pt-BR trata "." como milhar → "19.50" vira 1950.

Fix: Val(Replace(Replace(s, ".", ""), ",", ".")) é locale-neutro.
"""
import sys, shutil

FRM = r"C:\Projeto\OnlineCommerce\Forms\frmDIFAL_Cadastro.frm"
BAK = FRM + ".bak_locale_fix"

shutil.copy2(FRM, BAK)
print(f"Backup: {BAK}")

with open(FRM, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

patches = []

# AliquotaInterna
patches.append((
    b"    If Not IsNumeric(Replace(Trim(txtAliqInterna.Text), \",\", \".\")) Then\r\n"
    b"        MsgBox \"Al\xedquota Interna inv\xe1lida!\", vbExclamation, \"Aviso\": Exit Sub\r\n"
    b"    End If\r\n"
    b"    dAliqInt = CDbl(Replace(Trim(txtAliqInterna.Text), \",\", \".\"))\r\n",

    b"    If Not IsNumeric(Replace(Replace(Trim(txtAliqInterna.Text), \".\", \"\"), \",\", \".\")) Then\r\n"
    b"        MsgBox \"Al\xedquota Interna inv\xe1lida!\", vbExclamation, \"Aviso\": Exit Sub\r\n"
    b"    End If\r\n"
    b"    dAliqInt = Val(Replace(Replace(Trim(txtAliqInterna.Text), \".\", \"\"), \",\", \".\"))\r\n",

    "AliqInterna: CDbl -> Val locale-safe"
))

# AliquotaFCP
patches.append((
    b"    If Not IsNumeric(Replace(Trim(txtAliqFCP.Text), \",\", \".\")) Then\r\n"
    b"        MsgBox \"Al\xedquota FCP inv\xe1lida!\", vbExclamation, \"Aviso\": Exit Sub\r\n"
    b"    End If\r\n"
    b"    dAliqFCP = CDbl(Replace(Trim(txtAliqFCP.Text), \",\", \".\"))\r\n",

    b"    If Not IsNumeric(Replace(Replace(Trim(txtAliqFCP.Text), \".\", \"\"), \",\", \".\")) Then\r\n"
    b"        MsgBox \"Al\xedquota FCP inv\xe1lida!\", vbExclamation, \"Aviso\": Exit Sub\r\n"
    b"    End If\r\n"
    b"    dAliqFCP = Val(Replace(Replace(Trim(txtAliqFCP.Text), \".\", \"\"), \",\", \".\"))\r\n",

    "AliqFCP: CDbl -> Val locale-safe"
))

errors = 0
for old, new, desc in patches:
    cnt = data.count(old)
    if cnt != 1:
        print(f"ERRO {desc}: count={cnt}")
        errors += 1
    else:
        data = data.replace(old, new)
        print(f"OK   {desc}")

if errors:
    print(f"\n{errors} erro(s). Arquivo NAO salvo.")
    sys.exit(1)

data = norm(data)
with open(FRM, "wb") as f:
    f.write(data)
print("\nArquivo salvo com sucesso.")
