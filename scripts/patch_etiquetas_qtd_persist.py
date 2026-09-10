# -*- coding: utf-8 -*-
"""Etiquetas_Impressao.frm: persistir tambem a QUANTIDADE de etiquetas por produto marcado
(mapa xQtdImpressao "|cod=qtd|..."), nao so o check. .frm cp1252 -> edicao binaria + CRLF."""
import sys

p = r"C:\projeto\Compartilhado\Forms\Etiquetas_Impressao.frm"
d = open(p, "rb").read().replace(b"\r\n", b"\n")
Q = b"'"

reps = [
 # R1: var de modulo
 (b"Dim xProdutosSelecionados As String\n",
  b"Dim xProdutosSelecionados As String\n"
  b'Dim xQtdImpressao As String   ' + Q + b' qtde de etiquetas por produto marcado, formato "|cod=qtd|cod=qtd"\n'),

 # R2: troca ImpressoesDoGrid pelos helpers do mapa
 (b"Private Function ImpressoesDoGrid(ByVal sCod As String) As Integer\n"
  b"    'qtde de etiquetas: se o produto esta numa linha visivel do grid, respeita o valor editado; senao 1\n"
  b"    Dim k As Integer\n"
  b"    ImpressoesDoGrid = 1\n"
  b"    For k = 1 To Grid.rows - 1\n"
  b"        If Grid.TextMatrix(k, 2) = sCod Then\n"
  b"            If IsNumeric(Grid.TextMatrix(k, 4)) Then ImpressoesDoGrid = CInt(Grid.TextMatrix(k, 4))\n"
  b"            Exit Function\n"
  b"        End If\n"
  b"    Next k\n"
  b"End Function\n",
  b"Private Function GetQtdImpr(ByVal sCod As String) As Long\n"
  b"    'qtde de etiquetas do produto, persistida no mapa xQtdImpressao; default 1\n"
  b"    Dim p1 As Long, p2 As Long, sVal As String\n"
  b"    GetQtdImpr = 1\n"
  b'    p1 = InStr(1, xQtdImpressao, "|" & sCod & "=")\n'
  b"    If p1 = 0 Then Exit Function\n"
  b"    p1 = p1 + Len(sCod) + 2\n"
  b'    p2 = InStr(p1, xQtdImpressao, "|")\n'
  b"    If p2 = 0 Then p2 = Len(xQtdImpressao) + 1\n"
  b"    sVal = Mid$(xQtdImpressao, p1, p2 - p1)\n"
  b"    If IsNumeric(sVal) Then GetQtdImpr = CLng(sVal)\n"
  b"    If GetQtdImpr < 1 Then GetQtdImpr = 1\n"
  b"End Function\n"
  b"\n"
  b"Private Sub RemoveQtdImpr(ByVal sCod As String)\n"
  b"    Dim p1 As Long, p2 As Long\n"
  b'    p1 = InStr(1, xQtdImpressao, "|" & sCod & "=")\n'
  b"    If p1 = 0 Then Exit Sub\n"
  b'    p2 = InStr(p1 + 1, xQtdImpressao, "|")\n'
  b"    If p2 = 0 Then p2 = Len(xQtdImpressao) + 1\n"
  b"    xQtdImpressao = Left$(xQtdImpressao, p1 - 1) & Mid$(xQtdImpressao, p2)\n"
  b"End Sub\n"
  b"\n"
  b"Private Sub SetQtdImpr(ByVal sCod As String, ByVal nQtd As Long)\n"
  b"    RemoveQtdImpr sCod\n"
  b"    If nQtd < 1 Then nQtd = 1\n"
  b'    xQtdImpressao = xQtdImpressao & "|" & sCod & "=" & nQtd\n'
  b"End Sub\n"
  b"\n"
  b"Private Sub AtualizarQtdDaLinha(ByVal nRow As Long)\n"
  b"    'chamado quando o usuario edita a coluna IMPRESSOES - so grava se a linha estiver marcada\n"
  b"    If nRow < 1 Or nRow > Grid.rows - 1 Then Exit Sub\n"
  b"    Dim sCod As String\n"
  b"    sCod = Grid.TextMatrix(nRow, 2)\n"
  b"    If Not SelProcuraValor(xProdutosSelecionados, sCod) Then Exit Sub\n"
  b"    If IsNumeric(Grid.TextMatrix(nRow, 4)) Then SetQtdImpr sCod, CLng(Grid.TextMatrix(nRow, 4))\n"
  b"End Sub\n"),

 # R3: cmdImprimirEtiqueta usa o mapa
 (b'        arrayDeDados(idx, 0) = ImpressoesDoGrid(CStr(rP("codigo")))\n',
  b'        arrayDeDados(idx, 0) = GetQtdImpr(CStr(rP("codigo")))\n'),

 # R4: Grid_Click - marca/desmarca sincroniza o mapa
 (b"    If Grid.Col = 3 And editandoEtiqueta Then\n"
  b"        SelAdicionaValor xProdutosSelecionados, Grid.TextMatrix(Grid.Row, 2)\n"
  b"        If Grid.CellPicture = picChecked Then\n"
  b"            Set Grid.CellPicture = picUnchecked\n"
  b'            Grid.TextMatrix(Grid.Row, 4) = ""   ' + Q + b"senao SomaGrid ainda conta esta linha\n"
  b"            lblQuantSelecionada.Caption = SomaGrid(Grid, 4)\n"
  b"        Else\n"
  b"            Set Grid.CellPicture = picChecked\n"
  b"            Grid.TextMatrix(Grid.Row, 4) = 1\n"
  b"            lblQuantSelecionada.Caption = SomaGrid(Grid, 4)\n"
  b"        End If\n"
  b"    End If\n",
  b"    If Grid.Col = 3 And editandoEtiqueta Then\n"
  b"        SelAdicionaValor xProdutosSelecionados, Grid.TextMatrix(Grid.Row, 2)\n"
  b"        If Grid.CellPicture = picChecked Then\n"
  b"            Set Grid.CellPicture = picUnchecked\n"
  b'            Grid.TextMatrix(Grid.Row, 4) = ""   ' + Q + b"senao SomaGrid ainda conta esta linha\n"
  b"            RemoveQtdImpr Grid.TextMatrix(Grid.Row, 2)\n"
  b"            lblQuantSelecionada.Caption = SomaGrid(Grid, 4)\n"
  b"        Else\n"
  b"            Set Grid.CellPicture = picChecked\n"
  b"            Grid.TextMatrix(Grid.Row, 4) = 1\n"
  b"            SetQtdImpr Grid.TextMatrix(Grid.Row, 2), 1\n"
  b"            lblQuantSelecionada.Caption = SomaGrid(Grid, 4)\n"
  b"        End If\n"
  b"    End If\n"),

 # R5: Formatar_Grid_Etiquetas - restaura a qtde do mapa (nao mais fixo 1)
 (b"      If SelProcuraValor(xProdutosSelecionados, .TextMatrix(i, 2)) Then\n"
  b"         Set .CellPicture = picChecked.Picture\n"
  b"         If Not IsNumeric(.TextMatrix(i, 4)) Then .TextMatrix(i, 4) = 1\n"
  b"      Else\n",
  b"      If SelProcuraValor(xProdutosSelecionados, .TextMatrix(i, 2)) Then\n"
  b"         Set .CellPicture = picChecked.Picture\n"
  b"         .TextMatrix(i, 4) = GetQtdImpr(.TextMatrix(i, 2))\n"
  b"      Else\n"),

 # R6: txtEdit_LostFocus grava a qtde editada no mapa
 (b'Private Sub txtEdit_LostFocus()\n'
  b'Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)\n'
  b'lblQuantSelecionada.Caption = SomaGrid(Grid, 4)\n',
  b'Private Sub txtEdit_LostFocus()\n'
  b'Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)\n'
  b'AtualizarQtdDaLinha iRow\n'
  b'lblQuantSelecionada.Caption = SomaGrid(Grid, 4)\n'),

 # R7: txtEdit_KeyUp (2 pontos de escrita)
 (b'      Grid.Row = iRow - 1\n'
  b'      Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)\n'
  b'      Grid_Click\n',
  b'      Grid.Row = iRow - 1\n'
  b'      Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)\n'
  b'      AtualizarQtdDaLinha iRow\n'
  b'      Grid_Click\n'),
 (b'      Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)\n'
  b'      Grid.Row = iRow + 1\n'
  b'      Grid_Click\n',
  b'      Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)\n'
  b'      AtualizarQtdDaLinha iRow\n'
  b'      Grid.Row = iRow + 1\n'
  b'      Grid_Click\n'),

 # R8: cmdCancelar limpa tambem o mapa
 (b'xProdutosSelecionados = ""   ' + Q + b"so aqui (e ao fechar a janela) a selecao e zerada\n",
  b'xProdutosSelecionados = ""   ' + Q + b"so aqui (e ao fechar a janela) a selecao e zerada\n"
  b'xQtdImpressao = ""\n'),
]

for i, (o, n) in enumerate(reps, 1):
    if n in d and o not in d:
        print(f"#{i} ja"); continue
    if d.count(o) != 1:
        print(f"#{i}: {d.count(o)} ocorrencias -- ABORTA"); sys.exit(1)
    d = d.replace(o, n, 1); print(f"#{i} OK")

open(p, "wb").write(d.replace(b"\n", b"\r\n"))
print("gravado")
