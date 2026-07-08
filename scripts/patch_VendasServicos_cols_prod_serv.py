"""
patch_VendasServicos_cols_prod_serv.py

Adiciona colunas PRODUTOS (col4) e SERVICOS (col5) antes de SUBTOTAL em
VendasServicos_Consulta.frm, deslocando colunas 4-9 para 6-11.

Estrategia: busca por linha no grid de itens via per-row queries em
FormatarGrid_Vendas (evita tocar nas queries complexas de cmdLocalizar_Click).

- PRODUTOS: SUM(pedidos_itens.total) WHERE cod_pedido = X
- SERVICOS: SUM(OS_Servicos_Auto.total) JOIN OS WHERE OS.COD_PEDIDO = X
"""
import sys, shutil

FRM = r"C:\Projeto\OnlineCommerce\Forms\VendasServicos_Consulta.frm"
BAK = FRM + ".bak_cols_prod_serv"

shutil.copy2(FRM, BAK)
print(f"Backup: {BAK}")

with open(FRM, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

errors = 0
patches = []

# ---------------------------------------------------------------------------
# P1: Grid setup — Cols + ColWidths + cabecalhos
# ---------------------------------------------------------------------------
patches.append((
    b"      .Clear\r\n"
    b"      .Cols = 10\r\n"
    b"      .rows = 2\r\n"
    b"      \r\n"
    b"      .ColWidth(0) = 150\r\n"
    b"      .ColWidth(1) = 800\r\n"
    b"      .ColWidth(2) = 900\r\n"
    b"      .ColWidth(3) = 4000\r\n"
    b"      .ColWidth(4) = 1100\r\n"
    b"      .ColWidth(5) = 800\r\n"
    b"      .ColWidth(6) = 800\r\n"
    b"      .ColWidth(7) = 1000\r\n"
    b"      .ColWidth(8) = 850\r\n"
    b"      .ColWidth(9) = 900\r\n"
    b"      '.ColWidth(10) = 1300\r\n"
    b"     \r\n"
    b"      .TextMatrix(0, 1) = \"PEDIDO\"\r\n"
    b"      .TextMatrix(0, 2) = \"DATA\"\r\n"
    b"      .TextMatrix(0, 3) = \"NOME DO CLIENTE\"\r\n"
    b"      .TextMatrix(0, 4) = \"SUBTOTAL\"\r\n"
    b"      .TextMatrix(0, 5) = \"DESC.\"\r\n"
    b"      .TextMatrix(0, 6) = \"ACRE.\"\r\n"
    b"      .TextMatrix(0, 7) = \"VALOR\"\r\n"
    b"      .TextMatrix(0, 8) = \"FORMA\"\r\n"
    b"      .TextMatrix(0, 9) = \"TIPO\"\r\n"
    b"      '.TextMatrix(0, 10) = \"FORMA\"\r\n",

    b"      .Clear\r\n"
    b"      .Cols = 12\r\n"
    b"      .rows = 2\r\n"
    b"      \r\n"
    b"      .ColWidth(0)  = 150\r\n"
    b"      .ColWidth(1)  = 800\r\n"
    b"      .ColWidth(2)  = 900\r\n"
    b"      .ColWidth(3)  = 3000\r\n"
    b"      .ColWidth(4)  = 1100\r\n"
    b"      .ColWidth(5)  = 1100\r\n"
    b"      .ColWidth(6)  = 1000\r\n"
    b"      .ColWidth(7)  = 700\r\n"
    b"      .ColWidth(8)  = 700\r\n"
    b"      .ColWidth(9)  = 1000\r\n"
    b"      .ColWidth(10) = 850\r\n"
    b"      .ColWidth(11) = 900\r\n"
    b"     \r\n"
    b"      .TextMatrix(0, 1)  = \"PEDIDO\"\r\n"
    b"      .TextMatrix(0, 2)  = \"DATA\"\r\n"
    b"      .TextMatrix(0, 3)  = \"NOME DO CLIENTE\"\r\n"
    b"      .TextMatrix(0, 4)  = \"PRODUTOS\"\r\n"
    b"      .TextMatrix(0, 5)  = \"SERVI\xc7OS\"\r\n"
    b"      .TextMatrix(0, 6)  = \"SUBTOTAL\"\r\n"
    b"      .TextMatrix(0, 7)  = \"DESC.\"\r\n"
    b"      .TextMatrix(0, 8)  = \"ACRE.\"\r\n"
    b"      .TextMatrix(0, 9)  = \"VALOR\"\r\n"
    b"      .TextMatrix(0, 10) = \"FORMA\"\r\n"
    b"      .TextMatrix(0, 11) = \"TIPO\"\r\n",

    "P1: grid setup 10->12 cols + cabecalhos PRODUTOS/SERVICOS"
))

# ---------------------------------------------------------------------------
# P2: Loop de dados — adiciona cols PRODUTOS/SERVICOS e desloca 4-9 -> 6-11
# ---------------------------------------------------------------------------
patches.append((
    b"            .TextMatrix(.rows - 1, 1) = Format(rTabela(\"var_codped\"), \"000000\")\r\n"
    b"            .TextMatrix(.rows - 1, 2) = Format(rTabela(\"data_compra\"), \"dd/mm/yy\")\r\n"
    b"            .TextMatrix(.rows - 1, 3) = ValidateNull(rTabela(\"nome\"))\r\n"
    b"            .TextMatrix(.rows - 1, 4) = Format(rTabela(\"subtotal\"), ocMONEY)\r\n"
    b"            .TextMatrix(.rows - 1, 5) = Format(rTabela(\"ValorDescReal\"), ocMONEY)\r\n"
    b"            .TextMatrix(.rows - 1, 6) = Format(rTabela(\"ValorAcrescReal\"), ocMONEY)\r\n"
    b"            .TextMatrix(.rows - 1, 7) = Format(rTabela(\"var_total\"), ocMONEY)\r\n"
    b"            .TextMatrix(.rows - 1, 8) = rTabela(\"tipo_pagamento\")\r\n"
    b"            .TextMatrix(.rows - 1, 9) = rTabela(\"TIPO_PEDIDO\")\r\n",

    b"            .TextMatrix(.rows - 1, 1) = Format(rTabela(\"var_codped\"), \"000000\")\r\n"
    b"            .TextMatrix(.rows - 1, 2) = Format(rTabela(\"data_compra\"), \"dd/mm/yy\")\r\n"
    b"            .TextMatrix(.rows - 1, 3) = ValidateNull(rTabela(\"nome\"))\r\n"
    b"            Dim rSub As ADODB.Recordset\r\n"
    b"            Dim lCodPed As Long\r\n"
    b"            lCodPed = CLng(rTabela(\"var_codped\"))\r\n"
    b"            RsOpen rSub, \"SELECT ISNULL(SUM(total),0) AS vProd FROM pedidos_itens WHERE cod_pedido=\" & lCodPed\r\n"
    b"            .TextMatrix(.rows - 1, 4) = Format(rSub(\"vProd\"), ocMONEY)\r\n"
    b"            rSub.Close\r\n"
    b"            RsOpen rSub, \"SELECT ISNULL(SUM(sv.total),0) AS vServ FROM OS_Servicos_Auto sv INNER JOIN OS ON sv.cod_os=OS.COD_OS WHERE OS.COD_PEDIDO=\" & lCodPed\r\n"
    b"            .TextMatrix(.rows - 1, 5) = Format(rSub(\"vServ\"), ocMONEY)\r\n"
    b"            rSub.Close\r\n"
    b"            Set rSub = Nothing\r\n"
    b"            .TextMatrix(.rows - 1, 6) = Format(rTabela(\"subtotal\"), ocMONEY)\r\n"
    b"            .TextMatrix(.rows - 1, 7) = Format(rTabela(\"ValorDescReal\"), ocMONEY)\r\n"
    b"            .TextMatrix(.rows - 1, 8) = Format(rTabela(\"ValorAcrescReal\"), ocMONEY)\r\n"
    b"            .TextMatrix(.rows - 1, 9) = Format(rTabela(\"var_total\"), ocMONEY)\r\n"
    b"            .TextMatrix(.rows - 1, 10) = rTabela(\"tipo_pagamento\")\r\n"
    b"            .TextMatrix(.rows - 1, 11) = rTabela(\"TIPO_PEDIDO\")\r\n",

    "P2: loop dados - PRODUTOS col4, SERVICOS col5, demais 6-11"
))

# ---------------------------------------------------------------------------
# P3: Cor da coluna VALOR: col 7 -> col 9
# ---------------------------------------------------------------------------
patches.append((
    b"      'MUDAR COR DE FONTE DA COLUNA\r\n"
    b"      For i = 1 To .rows - 1\r\n"
    b"         .Row = i\r\n"
    b"         .Col = 7\r\n"
    b"         .CellForeColor = &HC0&\r\n"
    b"         .CellFontBold = True\r\n"
    b"      Next\r\n",

    b"      'MUDAR COR DE FONTE DA COLUNA\r\n"
    b"      For i = 1 To .rows - 1\r\n"
    b"         .Row = i\r\n"
    b"         .Col = 9\r\n"
    b"         .CellForeColor = &HC0&\r\n"
    b"         .CellFontBold = True\r\n"
    b"      Next\r\n",

    "P3: cor VALOR col 7 -> col 9"
))

# ---------------------------------------------------------------------------
# P4: SomaGrid — ajustar indices de coluna
# ---------------------------------------------------------------------------
patches.append((
    b"    lblSubtotal = Format(SomaGrid(Grid, 4), ocMONEY)\r\n"
    b"    lblTotalDesc = Format(SomaGrid(Grid, 5), ocMONEY)\r\n"
    b"    lblTotalAcresc = Format(SomaGrid(Grid, 6), ocMONEY)\r\n"
    b"    lblSubtotalBruto = Format(SomaGrid(Grid, 7), ocMONEY)\r\n",

    b"    lblSubtotal = Format(SomaGrid(Grid, 6), ocMONEY)\r\n"
    b"    lblTotalDesc = Format(SomaGrid(Grid, 7), ocMONEY)\r\n"
    b"    lblTotalAcresc = Format(SomaGrid(Grid, 8), ocMONEY)\r\n"
    b"    lblSubtotalBruto = Format(SomaGrid(Grid, 9), ocMONEY)\r\n",

    "P4: SomaGrid indices 4-7 -> 6-9"
))

# ---------------------------------------------------------------------------
# P5: cmdExibirProdutos_Click — TIPO_PEDIDO de col 9 -> col 11
# ---------------------------------------------------------------------------
patches.append((
    b"      Parcelas_Consulta_Produtos.loadPedidos Grid.TextMatrix(Grid.Row, 1), Grid.TextMatrix(Grid.Row, 9)\r\n",
    b"      Parcelas_Consulta_Produtos.loadPedidos Grid.TextMatrix(Grid.Row, 1), Grid.TextMatrix(Grid.Row, 11)\r\n",
    "P5: cmdExibirProdutos TIPO_PEDIDO col 9 -> col 11"
))

# ---------------------------------------------------------------------------
for old, new, desc in patches:
    cnt = data.count(old)
    if cnt != 1:
        print(f"ERRO {desc}: count={cnt} (esperado 1)")
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
