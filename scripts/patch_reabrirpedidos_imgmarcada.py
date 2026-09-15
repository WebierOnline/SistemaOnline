# -*- coding: utf-8 -*-
"""Estorno_ReabrirPedidos.frm - troca o texto "SIM" pela imagem ImgMarcada (usuario ja
adicionou o controle no form) nas colunas CANCEL.(6)/ABERTO(7).

Como o valor cru "SIM"/"" so e lido dentro dessa mesma Sub (nenhum outro lugar do form le
essas colunas - form pequeno, so leitura + fechar), da pra guardar em 2 colunas ocultas
novas (8/9) sem precisar mapear referencias externas, igual foi feito no Estonar.frm.

O loop antigo de "MUDAR COR DE FONTE" (vermelho/preto no texto) sai - nao faz mais sentido
sem texto pra colorir, a imagem substitui.

.frm cp1252 -> edicao binaria + normaliza CRLF.
"""
import sys

p = r"C:\projeto\PDV\Forms\Estorno_ReabrirPedidos.frm"
d = open(p, "rb").read().replace(b"\r\n", b"\n").replace(b"\r", b"\n")

o = (
b'Private Sub FormatarGrid(rTabela As ADODB.Recordset)\n'
b'   Dim i As Integer\n'
b'   \n'
b'   With Grid\n'
b'      .Clear\n'
b'      .Cols = 8\n'
b'      .Rows = 2\n'
b'      \n'
b'      .ColWidth(0) = 0\n'
b'      .ColWidth(1) = 1300\n'
b'      .ColWidth(2) = 2000\n'
b'      .ColWidth(3) = 1300\n'
b'      .ColWidth(4) = 1000\n'
b'      .ColWidth(5) = 1000\n'
b'      .ColWidth(6) = 1000\n'
b'      .ColWidth(7) = 1000\n'
b'      \n'
b'      .TextMatrix(0, 1) = "COD_PEDIDO"\n'
b'      .TextMatrix(0, 2) = "USUARIO"\n'
b'      .TextMatrix(0, 3) = "VALOR"\n'
b'      .TextMatrix(0, 4) = "DATA"\n'
b'      .TextMatrix(0, 5) = "HORA"\n'
b'      .TextMatrix(0, 6) = "CANCEL."\n'
b'      .TextMatrix(0, 7) = "ABERTO"\n'
b'      \n'
b"      'colocar os cabe\xe7alho em negrito\n"
b'      For i = 0 To .Cols - 1\n'
b'         .Col = i\n'
b'         .Row = 0\n'
b'         .CellFontBold = True\n'
b'      Next\n'
b'      \n'
b"      'ALINHAMENTO\n"
b"      '.ColAlignment(2) = 1\n"
b'      \n'
b"      'centralizar o titulo\n"
b'      For i = 0 To .Cols - 1\n'
b'         .Row = 0\n'
b'         .Col = i\n'
b'         .CellAlignment = flexAlignCenterCenter\n'
b'      Next\n'
b'      \n'
b'      If Not rTabela Is Nothing Then\n'
b'         Do While Not rTabela.EOF\n'
b'            .TextMatrix(.Rows - 1, 1) = rTabela("COD_PEDIDO")\n'
b'            .TextMatrix(.Rows - 1, 2) = rTabela("LOGIN")\n'
b'            .TextMatrix(.Rows - 1, 3) = Format(rTabela("VLR_PEDIDO"), ocMONEY)\n'
b'            .TextMatrix(.Rows - 1, 4) = Format(rTabela("DATA"), "DD/MM/YY")\n'
b'            .TextMatrix(.Rows - 1, 5) = Format(rTabela("HORA"), ocHORA)\n'
b'            .TextMatrix(.Rows - 1, 6) = ValidateNull(rTabela("vCancelado"))\n'
b'            .TextMatrix(.Rows - 1, 7) = ValidateNull(rTabela("vStatus"))\n'
b'            rTabela.MoveNext\n'
b'            .Rows = .Rows + 1\n'
b'         Loop\n'
b'      End If\n'
b'      \n'
b"      'MUDAR COR DE FONTE DA COLUNA\n"
b'      For i = 1 To .Rows - 1\n'
b'         .Row = i\n'
b'         .Col = 6\n'
b'         If .TextMatrix(i, 6) = "SIM" Then\n'
b'            .CellForeColor = vbRed\n'
b'         Else\n'
b'            .CellForeColor = vbBlack\n'
b'         End If\n'
b'         .CellFontBold = True\n'
b'      Next\n'
b'\n'
b'      For i = 1 To .Rows - 1\n'
b'         .Row = i\n'
b'         .Col = 7\n'
b'         If .TextMatrix(i, 7) = "SIM" Then\n'
b'            .CellForeColor = vbRed\n'
b'         Else\n'
b'            .CellForeColor = vbBlack\n'
b'         End If\n'
b'         .CellFontBold = True\n'
b'      Next\n'
b'      \n'
b'      .Rows = .Rows - 1\n'
b'   End With\n'
b'End Sub'
)
n = (
b'Private Sub FormatarGrid(rTabela As ADODB.Recordset)\n'
b'   Dim i As Integer\n'
b'   \n'
b'   With Grid\n'
b'      .Clear\n'
b'      .Cols = 10\n'
b'      .Rows = 2\n'
b'      \n'
b'      .ColWidth(0) = 0\n'
b'      .ColWidth(1) = 1300\n'
b'      .ColWidth(2) = 2000\n'
b'      .ColWidth(3) = 1300\n'
b'      .ColWidth(4) = 1000\n'
b'      .ColWidth(5) = 1000\n'
b'      .ColWidth(6) = 1000\n'
b'      .ColWidth(7) = 1000\n'
b'      .ColWidth(8) = 0\n'
b'      .ColWidth(9) = 0\n'
b'      \n'
b'      .TextMatrix(0, 1) = "COD_PEDIDO"\n'
b'      .TextMatrix(0, 2) = "USUARIO"\n'
b'      .TextMatrix(0, 3) = "VALOR"\n'
b'      .TextMatrix(0, 4) = "DATA"\n'
b'      .TextMatrix(0, 5) = "HORA"\n'
b'      .TextMatrix(0, 6) = "CANCEL."\n'
b'      .TextMatrix(0, 7) = "ABERTO"\n'
b'      \n'
b"      'colocar os cabe\xe7alho em negrito\n"
b'      For i = 0 To .Cols - 1\n'
b'         .Col = i\n'
b'         .Row = 0\n'
b'         .CellFontBold = True\n'
b'      Next\n'
b'      \n'
b"      'ALINHAMENTO\n"
b"      '.ColAlignment(2) = 1\n"
b'      \n'
b"      'centralizar o titulo\n"
b'      For i = 0 To .Cols - 1\n'
b'         .Row = 0\n'
b'         .Col = i\n'
b'         .CellAlignment = flexAlignCenterCenter\n'
b'      Next\n'
b'      \n'
b'      If Not rTabela Is Nothing Then\n'
b'         Do While Not rTabela.EOF\n'
b'            .TextMatrix(.Rows - 1, 1) = rTabela("COD_PEDIDO")\n'
b'            .TextMatrix(.Rows - 1, 2) = rTabela("LOGIN")\n'
b'            .TextMatrix(.Rows - 1, 3) = Format(rTabela("VLR_PEDIDO"), ocMONEY)\n'
b'            .TextMatrix(.Rows - 1, 4) = Format(rTabela("DATA"), "DD/MM/YY")\n'
b'            .TextMatrix(.Rows - 1, 5) = Format(rTabela("HORA"), ocHORA)\n'
b'            .TextMatrix(.Rows - 1, 8) = ValidateNull(rTabela("vCancelado"))\n'
b'            .TextMatrix(.Rows - 1, 9) = ValidateNull(rTabela("vStatus"))\n'
b'            If .TextMatrix(.Rows - 1, 8) = "SIM" Then\n'
b'                .Row = .Rows - 1\n'
b'                .Col = 6\n'
b'                Set .CellPicture = ImgMarcada.Picture\n'
b'                .CellPictureAlignment = 4\n'
b'            End If\n'
b'            If .TextMatrix(.Rows - 1, 9) = "SIM" Then\n'
b'                .Row = .Rows - 1\n'
b'                .Col = 7\n'
b'                Set .CellPicture = ImgMarcada.Picture\n'
b'                .CellPictureAlignment = 4\n'
b'            End If\n'
b'            rTabela.MoveNext\n'
b'            .Rows = .Rows + 1\n'
b'         Loop\n'
b'      End If\n'
b'      \n'
b'      .Rows = .Rows - 1\n'
b'   End With\n'
b'End Sub'
)

if n in d and o not in d:
    print("[ja] FormatarGrid com ImgMarcada")
else:
    c = d.count(o)
    if c != 1:
        print("[ABORTA] -- %d ocorrencias de old (esperava 1)" % c)
        sys.exit(1)
    d = d.replace(o, n, 1)
    print("[ok] FormatarGrid: CANCEL./ABERTO viram ImgMarcada")

open(p, "wb").write(d.replace(b"\n", b"\r\n"))
print("gravado " + p)
