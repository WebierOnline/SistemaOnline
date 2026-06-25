"""
patch_imprimir_xml_outra_maquina.py

Em cmdImprimir_Click, substitui o Exit Sub silencioso (quando o XML nao existe
mesmo apos consultaNFe) pela mensagem informando que a nota foi gerada em outro
computador.
"""
import sys, shutil

FRM = r"C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm"
BAK = FRM + ".bak_imprimir_xml_outra_maquina"

shutil.copy2(FRM, BAK)
print(f"Backup: {BAK}")

with open(FRM, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

# "Você" = Voc\xea, "conseguirá" = conseguir\xe1, "impressão" = impress\xe3o
OLD = (
    b"     If Not Existe(xCaminhoXML) Then Exit Sub\r\n"
)

NEW = (
    b"     If Not Existe(xCaminhoXML) Then\r\n"
    b"         MsgBox \"Essa nota fiscal foi gerada em outro computador!\" & vbCrLf & _\r\n"
    b"                \"Voc\xea somente conseguir\xe1 gerar a impress\xe3o no computador que foi gerado a nota fiscal\", vbExclamation, \"Aviso do Sistema\"\r\n"
    b"         Exit Sub\r\n"
    b"     End If\r\n"
)

cnt = data.count(OLD)
if cnt != 1:
    print(f"ERRO: trecho encontrado {cnt}x (esperado 1). Arquivo NAO alterado.")
    sys.exit(1)

data = data.replace(OLD, NEW)
data = norm(data)

with open(FRM, "wb") as f:
    f.write(data)

print("OK: mensagem 'gerada em outro computador' adicionada em cmdImprimir_Click")
print("Arquivo salvo com sucesso.")
