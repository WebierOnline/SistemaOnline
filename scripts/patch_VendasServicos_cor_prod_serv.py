"""
patch_VendasServicos_cor_prod_serv.py

1. VendasServicos_Consulta.frm  — fundo amarelo claro (&H80FFFF = RGB 255,255,128)
   nas colunas PRODUTOS (4) e SERVICOS (5) do grid.

2. Parcelas_Consulta_Produtos.frm — fonte vermelho escuro (&H80 = RGB 128,0,0)
   em todas as celulas de linhas cujo tipo_item = "SERVICO".
"""
import sys, shutil

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

errors = 0

# ===========================================================================
# 1. VendasServicos_Consulta.frm
# ===========================================================================
FRM1 = r"C:\Projeto\OnlineCommerce\Forms\VendasServicos_Consulta.frm"
BAK1 = FRM1 + ".bak_cor_prod_serv"
shutil.copy2(FRM1, BAK1)
print(f"Backup: {BAK1}")

with open(FRM1, "rb") as f:
    d1 = f.read()

OLD1 = (
    b"      'MUDAR COR DE FONTE DA COLUNA\r\n"
    b"      For i = 1 To .rows - 1\r\n"
    b"         .Row = i\r\n"
    b"         .Col = 9\r\n"
    b"         .CellForeColor = &HC0&\r\n"
    b"         .CellFontBold = True\r\n"
    b"      Next\r\n"
    b"      \r\n"
    b"      .rows = .rows - 1\r\n"
    b"      Grid.Redraw = True\r\n"
    b"   End With\r\n"
)

NEW1 = (
    b"      'MUDAR COR DE FONTE DA COLUNA\r\n"
    b"      For i = 1 To .rows - 1\r\n"
    b"         .Row = i\r\n"
    b"         .Col = 9\r\n"
    b"         .CellForeColor = &HC0&\r\n"
    b"         .CellFontBold = True\r\n"
    b"      Next\r\n"
    b"      \r\n"
    b"      'COR DE FUNDO AMARELO CLARO PARA PRODUTOS E SERVICOS\r\n"
    b"      For i = 1 To .rows - 1\r\n"
    b"         .Row = i\r\n"
    b"         .Col = 4\r\n"
    b"         .CellBackColor = &H80FFFF\r\n"
    b"         .Col = 5\r\n"
    b"         .CellBackColor = &H80FFFF\r\n"
    b"      Next\r\n"
    b"      \r\n"
    b"      .rows = .rows - 1\r\n"
    b"      Grid.Redraw = True\r\n"
    b"   End With\r\n"
)

cnt = d1.count(OLD1)
if cnt != 1:
    print(f"ERRO VendasServicos P1: count={cnt} (esperado 1)")
    errors += 1
else:
    d1 = d1.replace(OLD1, NEW1)
    print("OK   VendasServicos: fundo amarelo cols 4 e 5")

if not errors:
    d1 = norm(d1)
    with open(FRM1, "wb") as f:
        f.write(d1)
    print("     Arquivo salvo.")

# ===========================================================================
# 2. Parcelas_Consulta_Produtos.frm
# ===========================================================================
FRM2 = r"C:\Projeto\Compartilhado\Forms\Parcelas_Consulta_Produtos.frm"
BAK2 = FRM2 + ".bak_cor_servico"
shutil.copy2(FRM2, BAK2)
print(f"Backup: {BAK2}")

with open(FRM2, "rb") as f:
    d2 = f.read()

# \xc7 = Ç (cp1252) — SERVIÇO
OLD2 = (
    b"            .TextMatrix(.rows - 1, 7) = Format(rTabela(\"total\"), ocMONEY)\r\n"
    b"\r\n"
    b"            rTabela.MoveNext\r\n"
    b"            .rows = .rows + 1\r\n"
    b"         Loop\r\n"
)

NEW2 = (
    b"            .TextMatrix(.rows - 1, 7) = Format(rTabela(\"total\"), ocMONEY)\r\n"
    b"\r\n"
    b"            If rTabela(\"tipo_item\") = \"SERVI\xc7O\" Then\r\n"
    b"               Dim j As Integer\r\n"
    b"               For j = 0 To .Cols - 1\r\n"
    b"                  .Row = .rows - 1\r\n"
    b"                  .Col = j\r\n"
    b"                  .CellForeColor = &H80&\r\n"
    b"               Next j\r\n"
    b"            End If\r\n"
    b"\r\n"
    b"            rTabela.MoveNext\r\n"
    b"            .rows = .rows + 1\r\n"
    b"         Loop\r\n"
)

cnt = d2.count(OLD2)
if cnt != 1:
    print(f"ERRO Parcelas P2: count={cnt} (esperado 1)")
    errors += 1
else:
    d2 = d2.replace(OLD2, NEW2)
    print("OK   Parcelas: fonte vermelho escuro em linhas SERVICO")

if errors == 0 or (errors > 0 and d2 != open(FRM2, "rb").read()):
    # salva apenas se houve mudanca no Parcelas
    if d2.count(OLD2) == 0:  # ja substituiu
        d2 = norm(d2)
        with open(FRM2, "wb") as f:
            f.write(d2)
        print("     Arquivo salvo.")

if errors:
    print(f"\n{errors} erro(s) encontrado(s).")
    sys.exit(1)
else:
    print("\nTudo aplicado com sucesso.")
