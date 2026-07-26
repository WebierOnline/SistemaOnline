"""
Patch: substitui AddItem hardcoded do cboISClasse por query em tbISClassTrib
filtrando registros vigentes hoje (GETDATE() BETWEEN dIniVig AND dFimVig)
"""
import sys

FILE = r"C:\Projeto\Compartilhado\Forms\Produtos_Cadastro.frm"

with open(FILE, "rb") as f:
    data = f.read()

data = data.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

# Bloco antigo: AddItem hardcoded (cp1252: \xe7=ç, \xe3=ã, \xed=í, \xfa=ú, \xe9=é, \xf3=ó)
old = (
    b"    ' Imposto Seletivo - Classifica\xe7\xe3o\r\n"
    b"    cboISClasse.AddItem \"900001 - Vinhos, Espumantes e Cacha\xe7as\"\r\n"
    b"    cboISClasse.AddItem \"900002 - Cervejas e Chopes\"\r\n"
    b"    cboISClasse.AddItem \"900003 - U\xedsque, Gin, Vodka (Destilados)\"\r\n"
    b"    cboISClasse.AddItem \"900010 - Refrigerantes e Sucos com A\xe7\xfacar\"\r\n"
    b"    cboISClasse.AddItem \"900011 - Energ\xe9ticos e Bebidas Esportivas\"\r\n"
    b"    cboISClasse.AddItem \"900020 - Ve\xedculos (Autom\xf3veis poluentes)\"\r\n"
    b"    cboISClasse.AddItem \"900040 - Cigarros e fumo\"\r\n"
    b"End Sub"
)

# Bloco novo: query em tbISClassTrib vigentes hoje
new = (
    b"    ' Imposto Seletivo - Classifica\xe7\xe3o (tbISClassTrib vigentes hoje)\r\n"
    b"    Dim rIS As ADODB.Recordset\r\n"
    b"    cboISClasse.Clear\r\n"
    b"    RsOpen rIS, \"SELECT cClassTrib_IS, Descricao FROM tbISClassTrib \" _\r\n"
    b"        & \"WHERE GETDATE() BETWEEN dIniVig AND ISNULL(dFimVig,'9999-12-31') \" _\r\n"
    b"        & \"ORDER BY cClassTrib_IS\"\r\n"
    b"    Do While Not rIS.EOF\r\n"
    b"        cboISClasse.AddItem rIS(\"cClassTrib_IS\") & \" - \" & rIS(\"Descricao\")\r\n"
    b"        rIS.MoveNext\r\n"
    b"    Loop\r\n"
    b"    If rIS.State <> 0 Then rIS.Close\r\n"
    b"End Sub"
)

c = data.count(old)
if c != 1:
    print(f"ERRO: padrao encontrado {c}x (esperado 1)")
    # Debug: mostra os bytes ao redor das linhas esperadas
    idx = data.find(b"cboISClasse.AddItem")
    if idx >= 0:
        print("Bytes encontrados:")
        print(repr(data[idx:idx+200]))
    sys.exit(1)

data = data.replace(old, new, 1)
print("OK: cboISClasse agora carrega de tbISClassTrib")

data = data.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")
with open(FILE, "wb") as f:
    f.write(data)
print("Arquivo gravado.")
