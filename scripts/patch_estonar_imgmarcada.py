# -*- coding: utf-8 -*-
"""Estonar.frm - troca o texto "SIM" pela imagem ImgMarcada (ja existe no form, Estonar.frx)
nas colunas Reaberto(14)/Cancel.(15)/NFCe(16) do Grid principal.

Problema: essas 3 colunas nao sao so visuais - o valor "SIM"/"" e lido de volta via
Grid.TextMatrix em ~12 lugares diferentes (guards de cmdExcluirPedido/cmdPedidoAbrir/
cmdModificar/cmdPDF/cmdPedidoImprimir/cmdReabrir, habilitar cmdReaberturas, e o loop de
totais de Mostrar_Pedido). Se so limpasse o texto e pusesse a imagem, todos esses guards
iam parar de funcionar (texto vazio <> "SIM" sempre).

Fix: separa ARMAZENAMENTO de EXIBICAO.
- 3 colunas novas e ocultas no fim do grid (20/21/22, ColWidth=0) guardam o "SIM"/"" cru,
  exatamente como antes - NADA na logica de negocio muda de comportamento.
- As colunas visiveis (14/15/16) ficam com o texto em branco e ganham `CellPicture =
  ImgMarcada.Picture` (mesmo padrao ja usado em Produtos_Estoque_Simples/Etiquetas_Impressao -
  ver [[reference_msflexgrid_cellpicture_performance]]) quando o valor da coluna oculta
  correspondente for "SIM". So marca (nao desenha nada quando nao e "SIM"), Redraw=False
  ja envolve o loop inteiro - ganho de performance ja garantido.
- Todo lugar que lia col 14/15/16 (Grid.TextMatrix(Grid.Row, ...) ou o loop de totais)
  passa a ler das colunas ocultas 20/21/22.

.frm cp1252 -> edicao binaria + normaliza CRLF.
"""
import sys

p = r"C:\projeto\PDV\Forms\Estonar.frm"
d = open(p, "rb").read().replace(b"\r\n", b"\n").replace(b"\r", b"\n")

reps = []  # (nome, old, new, esperado)

# ---------------------------------------------------- .Cols = 20 -> 23
reps.append(("Cols 20->23", b"      .Cols = 20\n", b"      .Cols = 23\n", 1))

# ---------------------------------------------------- 3 novas colunas ocultas
reps.append((
    "ColWidth novas colunas ocultas (20/21/22)",
    b"      .ColWidth(19) = 800\n",
    b"      .ColWidth(19) = 800\n      .ColWidth(20) = 0\n      .ColWidth(21) = 0\n      .ColWidth(22) = 0\n",
    1,
))

# ---------------------------------------------------- fill loop: separa armazenamento (oculto) de exibicao (imagem)
o = (
b'            .TextMatrix(.Rows - 1, 14) = rTabela("Var_StatusREABERTO")\n'
b'            .TextMatrix(.Rows - 1, 15) = rTabela("Var_StatusCANCELADO")\n'
b'            .TextMatrix(.Rows - 1, 16) = ValidateNull(rTabela("Var_StatusNFCE"))\n'
)
n = (
b'            .TextMatrix(.Rows - 1, 20) = rTabela("Var_StatusREABERTO")\n'
b'            .TextMatrix(.Rows - 1, 21) = rTabela("Var_StatusCANCELADO")\n'
b'            .TextMatrix(.Rows - 1, 22) = ValidateNull(rTabela("Var_StatusNFCE"))\n'
b'            If .TextMatrix(.Rows - 1, 20) = "SIM" Then\n'
b'                .Row = .Rows - 1\n'
b'                .Col = 14\n'
b'                Set .CellPicture = ImgMarcada.Picture\n'
b'                .CellPictureAlignment = 4\n'
b'            End If\n'
b'            If .TextMatrix(.Rows - 1, 21) = "SIM" Then\n'
b'                .Row = .Rows - 1\n'
b'                .Col = 15\n'
b'                Set .CellPicture = ImgMarcada.Picture\n'
b'                .CellPictureAlignment = 4\n'
b'            End If\n'
b'            If .TextMatrix(.Rows - 1, 22) = "SIM" Then\n'
b'                .Row = .Rows - 1\n'
b'                .Col = 16\n'
b'                Set .CellPicture = ImgMarcada.Picture\n'
b'                .CellPictureAlignment = 4\n'
b'            End If\n'
)
reps.append(("fill loop: coluna oculta + CellPicture", o, n, 1))

# ---------------------------------------------------- loop de totais (Mostrar_Pedido)
reps.append((
    "loop de totais: col 15 -> col 21 (oculta)",
    b'If .TextMatrix(i, 15) = "SIM" Then',
    b'If .TextMatrix(i, 21) = "SIM" Then',
    1,
))

# ---------------------------------------------------- ~12 guards espalhados pelo form (Grid.Row)
reps.append((
    'Grid.Row col 15 (Cancel.) -> col 21, todas as ocorrencias',
    b'Grid.TextMatrix(Grid.Row, 15) = "SIM"',
    b'Grid.TextMatrix(Grid.Row, 21) = "SIM"',
    6,
))
reps.append((
    'Grid.Row col 16 (NFCe) -> col 22, todas as ocorrencias',
    b'Grid.TextMatrix(Grid.Row, 16) = "SIM"',
    b'Grid.TextMatrix(Grid.Row, 22) = "SIM"',
    4,
))
reps.append((
    'Grid.Row col 14 (Reaberto) <> "SIM" -> col 20',
    b'Grid.TextMatrix(Grid.Row, 14) <> "SIM"',
    b'Grid.TextMatrix(Grid.Row, 20) <> "SIM"',
    1,
))
reps.append((
    'Grid.Row col 14 (Reaberto) = "SIM" -> col 20',
    b'Grid.TextMatrix(Grid.Row, 14) = "SIM"',
    b'Grid.TextMatrix(Grid.Row, 20) = "SIM"',
    1,
))

for nome, old, new, esperado in reps:
    if new in d and old not in d:
        print("[ja] " + nome); continue
    c = d.count(old)
    if c != esperado:
        print("[ABORTA] %s -- %d ocorrencias de old (esperava %d)" % (nome, c, esperado))
        sys.exit(1)
    if esperado > 1:
        d = d.replace(old, new)
    else:
        d = d.replace(old, new, 1)
    print("[ok]  " + nome + (" (%dx)" % c if esperado > 1 else ""))

open(p, "wb").write(d.replace(b"\n", b"\r\n"))
print("gravado " + p)
