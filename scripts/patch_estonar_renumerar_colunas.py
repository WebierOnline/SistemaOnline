# -*- coding: utf-8 -*-
"""Estonar.frm - CORRECAO: o ColPosition (patch_estonar_frete_maquina.py) nao funcionou como
esperado - o usuario mandou print mostrando FRETE ausente e MAQUINA no fim (depois de C�D/CX),
nao antes de CAIXA. Abandona ColPosition (comportamento nao confiavel/nao verificado nesse
projeto) e faz a coisa certa: RENUMERA as colunas de verdade, deslocando os indices logicos.

Mapa OLD -> NEW:
  13 VALOR -> 14
  14 Reaberto(img) -> 15
  15 Cancel.(img) -> 16
  16 NFCe(img) -> 17
  17 INUT(oculta,morta) -> 18
  18 CAIXA -> 20
  19 C�D/CX -> 21
  20 Reaberto(oculta crua) -> 22
  21 Cancel.(oculta crua) -> 23
  22 NFCe(oculta crua) -> 24
  NOVO: 13 = FRETE, NOVO: 19 = MAQUINA
  .Cols continua 25 (mesmo total, so remapeado)

Levantamento completo ANTES de mexer (ver conversa) - contei toda ocorrencia de cada padrao
por familia (Grid.Row, .Rows-1, (i,..), iLinha, .Col=, ColWidth, TextMatrix(0,..)) pra saber
exatamente quantas trocar e nao deixar nenhuma pra tras. FormatarGrid_Pedido e
chkMostrarCaixa_Click sao reescritos inteiros (concentram quase toda a mudanca estrutural);
o resto (~35 ocorrencias espalhadas por outros 10 Subs) e trocado por padrao literal exato,
em ORDEM DESCENDENTE de numero antigo (22,21,20,19,18,13) pra nunca uma troca esbarrar no
resultado de outra troca (13->14 e 18->20 nao colidem, mas 18->20 e 20->22 colidiriam se
trocasse 18 antes de 20 - por isso a ordem descendente).

.frm cp1252 -> edicao binaria + normaliza CRLF.
"""
import sys

p = r"C:\projeto\PDV\Forms\Estonar.frm"
d = open(p, "rb").read().replace(b"\r\n", b"\n").replace(b"\r", b"\n")

reps = []  # (nome, old, new, esperado)

# ==================================================================================
# 1) FormatarGrid_Pedido - reescrito inteiro com a numeracao final
# ==================================================================================
o_fmt = (
b'Private Sub FormatarGrid_Pedido(rTabela As ADODB.Recordset)\n'
b'Dim i As Integer\n'
b'Dim j As Integer\n'
b'\n'
b'   With Grid\n'
b'      .Redraw = False\n'
b'      .Clear\n'
b'      .Cols = 25\n'
b'      .Rows = 2\n'
b'      \n'
b'      .ColWidth(0) = 0\n'
b'      .ColWidth(1) = 1050\n'
b'      .ColWidth(2) = 630\n'
b'      .ColWidth(3) = 0\n'
b'      .ColWidth(4) = 750\n'
b'      .ColWidth(5) = 0\n'
b'      .ColWidth(6) = 270\n'
b'      .ColWidth(7) = 750\n'
b'      .ColWidth(8) = 1450\n'
b'      .ColWidth(9) = 3000\n'
b'      .ColWidth(10) = 900\n'
b'      .ColWidth(11) = 900\n'
b'      .ColWidth(12) = 900\n'
b'      .ColWidth(13) = 900\n'
b'      .ColWidth(14) = 800\n'
b'      .ColWidth(15) = 800\n'
b'      .ColWidth(16) = 800\n'
b'      .ColWidth(17) = 0\n'
b'      .ColWidth(18) = IIf(chkMostrarCaixa.Value = Checked, 800, 0)\n'
b'      .ColWidth(19) = IIf(chkMostrarCaixa.Value = Checked, 800, 0)\n'
b'      .ColWidth(20) = 0\n'
b'      .ColWidth(21) = 0\n'
b'      .ColWidth(22) = 0\n'
b'      .ColWidth(23) = 900\n'
b'      .ColWidth(24) = IIf(chkMostrarCaixa.Value = Checked, 900, 0)\n'
b'      \n'
b"      'FRETE entra na posicao 13 (entre ACRESC. e VALOR); MAQUINA na 19 (antes de CAIXA) -\n"
b"      'ColPosition so muda onde aparece na tela, TextMatrix/Col continuam pelo indice logico\n"
b'      .ColPosition(23) = 13\n'
b'      .ColPosition(24) = 19\n'
b'      \n'
b'      .TextMatrix(0, 1) = "TIPO"\n'
b'      .TextMatrix(0, 2) = "PEDIDO"\n'
b'      .TextMatrix(0, 3) = "STATUS"\n'
b'      .TextMatrix(0, 4) = "EMISS\xc3O"\n'
b'      .TextMatrix(0, 5) = "TIPO"\n'
b'      .TextMatrix(0, 6) = "V"\n'
b'      .TextMatrix(0, 7) = "FORMA"\n'
b'      .TextMatrix(0, 8) = "TIPO"\n'
b'      .TextMatrix(0, 9) = "CLIENTE"\n'
b'      .TextMatrix(0, 10) = "SUBTOT."\n'
b'      .TextMatrix(0, 11) = "DESC."\n'
b'      .TextMatrix(0, 12) = "ACRESC."\n'
b'      .TextMatrix(0, 13) = "VALOR"\n'
b'      .TextMatrix(0, 14) = "Reaberto"\n'
b'      .TextMatrix(0, 15) = "Cancel."\n'
b'      .TextMatrix(0, 16) = "NFCe"\n'
b'      .TextMatrix(0, 17) = "INUT"\n'
b'      .TextMatrix(0, 18) = "CAIXA"\n'
b'      .TextMatrix(0, 19) = "C\xd3D/CX"\n'
b'      .TextMatrix(0, 23) = "FRETE"\n'
b'      .TextMatrix(0, 24) = "MAQUINA"\n'
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
b'      Next i\n'
b'\n'
b'      If Not rTabela Is Nothing Then\n'
b'         Do While Not rTabela.EOF\n'
b'            .TextMatrix(.Rows - 1, 1) = rTabela("var_TipoPedido")\n'
b'            .TextMatrix(.Rows - 1, 2) = rTabela("var_CodPedido")\n'
b'            .TextMatrix(.Rows - 1, 3) = ValidateNull(rTabela("Var_StatusPEDIDO"))\n'
b'            .TextMatrix(.Rows - 1, 4) = Format(rTabela("var_Data"), "dd/mm/yy")\n'
b'            .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("var_TipoPedido"))\n'
b'            .TextMatrix(.Rows - 1, 6) = ValidateNull(rTabela("varCod_Func"))\n'
b'            .TextMatrix(.Rows - 1, 7) = ValidateNull(rTabela("var_TipoPagamento"))\n'
b'            If chkIncompleto.Value = Unchecked Then\n'
b'                .TextMatrix(.Rows - 1, 8) = ValidateNull(rTabela("var_Pagamento"))\n'
b'            End If\n'
b'        If cboStatus.Text <> "VAZIO" Then\n'
b'            .TextMatrix(.Rows - 1, 9) = ValidateNull(rTabela("var_Cliente"))\n'
b'        End If\n'
b'            .TextMatrix(.Rows - 1, 10) = Format(rTabela("var_SUBTOTAL"), ocMONEY)\n'
b'            .TextMatrix(.Rows - 1, 11) = Format(rTabela("var_DESC"), ocMONEY)\n'
b'            .TextMatrix(.Rows - 1, 12) = Format(rTabela("var_ACRESC"), ocMONEY)\n'
b'            .TextMatrix(.Rows - 1, 13) = Format(rTabela("var_Total"), ocMONEY)\n'
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
b"'            .TextMatrix(.Rows - 1, 17) = rTabela(\"Var_NFCEInutilizada\")\n"
b'            .TextMatrix(.Rows - 1, 18) = ValidateNull(rTabela("VarPEDCAIXA"))\n'
b'            .TextMatrix(.Rows - 1, 19) = ValidateNull(rTabela("VarPEDCODCAIXA"))\n'
b'            .TextMatrix(.Rows - 1, 23) = Format(rTabela("varPedFrete"), ocMONEY)\n'
b'            .TextMatrix(.Rows - 1, 24) = ValidateNull(rTabela("varPedMaquina"))\n'
b'            rTabela.MoveNext\n'
b'            .Rows = .Rows + 1\n'
b'         Loop\n'
b'      End If\n'
b'      \n'
b'   FlexCores &HFFFFFF, &HE0E0E0\n'
b'\n'
b'      .Rows = .Rows - 1\n'
b'      .Redraw = True\n'
b'   End With\n'
b'End Sub'
)
n_fmt = (
b'Private Sub FormatarGrid_Pedido(rTabela As ADODB.Recordset)\n'
b'Dim i As Integer\n'
b'Dim j As Integer\n'
b'\n'
b'   With Grid\n'
b'      .Redraw = False\n'
b'      .Clear\n'
b'      .Cols = 25\n'
b'      .Rows = 2\n'
b'      \n'
b'      .ColWidth(0) = 0\n'
b'      .ColWidth(1) = 1050\n'
b'      .ColWidth(2) = 630\n'
b'      .ColWidth(3) = 0\n'
b'      .ColWidth(4) = 750\n'
b'      .ColWidth(5) = 0\n'
b'      .ColWidth(6) = 270\n'
b'      .ColWidth(7) = 750\n'
b'      .ColWidth(8) = 1450\n'
b'      .ColWidth(9) = 3000\n'
b'      .ColWidth(10) = 900\n'
b'      .ColWidth(11) = 900\n'
b'      .ColWidth(12) = 900\n'
b'      .ColWidth(13) = 900\n'
b'      .ColWidth(14) = 900\n'
b'      .ColWidth(15) = 800\n'
b'      .ColWidth(16) = 800\n'
b'      .ColWidth(17) = 800\n'
b'      .ColWidth(18) = 0\n'
b'      .ColWidth(19) = IIf(chkMostrarCaixa.Value = Checked, 900, 0)\n'
b'      .ColWidth(20) = IIf(chkMostrarCaixa.Value = Checked, 800, 0)\n'
b'      .ColWidth(21) = IIf(chkMostrarCaixa.Value = Checked, 800, 0)\n'
b'      .ColWidth(22) = 0\n'
b'      .ColWidth(23) = 0\n'
b'      .ColWidth(24) = 0\n'
b'      \n'
b'      .TextMatrix(0, 1) = "TIPO"\n'
b'      .TextMatrix(0, 2) = "PEDIDO"\n'
b'      .TextMatrix(0, 3) = "STATUS"\n'
b'      .TextMatrix(0, 4) = "EMISS\xc3O"\n'
b'      .TextMatrix(0, 5) = "TIPO"\n'
b'      .TextMatrix(0, 6) = "V"\n'
b'      .TextMatrix(0, 7) = "FORMA"\n'
b'      .TextMatrix(0, 8) = "TIPO"\n'
b'      .TextMatrix(0, 9) = "CLIENTE"\n'
b'      .TextMatrix(0, 10) = "SUBTOT."\n'
b'      .TextMatrix(0, 11) = "DESC."\n'
b'      .TextMatrix(0, 12) = "ACRESC."\n'
b'      .TextMatrix(0, 13) = "FRETE"\n'
b'      .TextMatrix(0, 14) = "VALOR"\n'
b'      .TextMatrix(0, 15) = "Reaberto"\n'
b'      .TextMatrix(0, 16) = "Cancel."\n'
b'      .TextMatrix(0, 17) = "NFCe"\n'
b'      .TextMatrix(0, 18) = "INUT"\n'
b'      .TextMatrix(0, 19) = "MAQUINA"\n'
b'      .TextMatrix(0, 20) = "CAIXA"\n'
b'      .TextMatrix(0, 21) = "C\xd3D/CX"\n'
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
b'      Next i\n'
b'\n'
b'      If Not rTabela Is Nothing Then\n'
b'         Do While Not rTabela.EOF\n'
b'            .TextMatrix(.Rows - 1, 1) = rTabela("var_TipoPedido")\n'
b'            .TextMatrix(.Rows - 1, 2) = rTabela("var_CodPedido")\n'
b'            .TextMatrix(.Rows - 1, 3) = ValidateNull(rTabela("Var_StatusPEDIDO"))\n'
b'            .TextMatrix(.Rows - 1, 4) = Format(rTabela("var_Data"), "dd/mm/yy")\n'
b'            .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("var_TipoPedido"))\n'
b'            .TextMatrix(.Rows - 1, 6) = ValidateNull(rTabela("varCod_Func"))\n'
b'            .TextMatrix(.Rows - 1, 7) = ValidateNull(rTabela("var_TipoPagamento"))\n'
b'            If chkIncompleto.Value = Unchecked Then\n'
b'                .TextMatrix(.Rows - 1, 8) = ValidateNull(rTabela("var_Pagamento"))\n'
b'            End If\n'
b'        If cboStatus.Text <> "VAZIO" Then\n'
b'            .TextMatrix(.Rows - 1, 9) = ValidateNull(rTabela("var_Cliente"))\n'
b'        End If\n'
b'            .TextMatrix(.Rows - 1, 10) = Format(rTabela("var_SUBTOTAL"), ocMONEY)\n'
b'            .TextMatrix(.Rows - 1, 11) = Format(rTabela("var_DESC"), ocMONEY)\n'
b'            .TextMatrix(.Rows - 1, 12) = Format(rTabela("var_ACRESC"), ocMONEY)\n'
b'            .TextMatrix(.Rows - 1, 13) = Format(rTabela("varPedFrete"), ocMONEY)\n'
b'            .TextMatrix(.Rows - 1, 14) = Format(rTabela("var_Total"), ocMONEY)\n'
b'            .TextMatrix(.Rows - 1, 22) = rTabela("Var_StatusREABERTO")\n'
b'            .TextMatrix(.Rows - 1, 23) = rTabela("Var_StatusCANCELADO")\n'
b'            .TextMatrix(.Rows - 1, 24) = ValidateNull(rTabela("Var_StatusNFCE"))\n'
b'            If .TextMatrix(.Rows - 1, 22) = "SIM" Then\n'
b'                .Row = .Rows - 1\n'
b'                .Col = 15\n'
b'                Set .CellPicture = ImgMarcada.Picture\n'
b'                .CellPictureAlignment = 4\n'
b'            End If\n'
b'            If .TextMatrix(.Rows - 1, 23) = "SIM" Then\n'
b'                .Row = .Rows - 1\n'
b'                .Col = 16\n'
b'                Set .CellPicture = ImgMarcada.Picture\n'
b'                .CellPictureAlignment = 4\n'
b'            End If\n'
b'            If .TextMatrix(.Rows - 1, 24) = "SIM" Then\n'
b'                .Row = .Rows - 1\n'
b'                .Col = 17\n'
b'                Set .CellPicture = ImgMarcada.Picture\n'
b'                .CellPictureAlignment = 4\n'
b'            End If\n'
b"'            .TextMatrix(.Rows - 1, 18) = rTabela(\"Var_NFCEInutilizada\")\n"
b'            .TextMatrix(.Rows - 1, 19) = ValidateNull(rTabela("varPedMaquina"))\n'
b'            .TextMatrix(.Rows - 1, 20) = ValidateNull(rTabela("VarPEDCAIXA"))\n'
b'            .TextMatrix(.Rows - 1, 21) = ValidateNull(rTabela("VarPEDCODCAIXA"))\n'
b'            rTabela.MoveNext\n'
b'            .Rows = .Rows + 1\n'
b'         Loop\n'
b'      End If\n'
b'      \n'
b'   FlexCores &HFFFFFF, &HE0E0E0\n'
b'\n'
b'      .Rows = .Rows - 1\n'
b'      .Redraw = True\n'
b'   End With\n'
b'End Sub'
)
reps.append(("FormatarGrid_Pedido inteiro reescrito", o_fmt, n_fmt, 1))

# ==================================================================================
# 2) chkMostrarCaixa_Click - reescrito com a numeracao final (19/20/21)
# ==================================================================================
o_chk = (
b'Private Sub chkMostrarCaixa_Click()\n'
b"'mostra/esconde MAQUINA, CAIXA e C\xd3D/CX (colunas 24, 18, 19) sem precisar reconsultar\n"
b'If chkMostrarCaixa.Value = Checked Then\n'
b'    Grid.ColWidth(24) = 900\n'
b'    Grid.ColWidth(18) = 800\n'
b'    Grid.ColWidth(19) = 800\n'
b'Else\n'
b'    Grid.ColWidth(24) = 0\n'
b'    Grid.ColWidth(18) = 0\n'
b'    Grid.ColWidth(19) = 0\n'
b'End If\n'
b'End Sub'
)
n_chk = (
b'Private Sub chkMostrarCaixa_Click()\n'
b"'mostra/esconde MAQUINA, CAIXA e C\xd3D/CX (colunas 19, 20, 21) sem precisar reconsultar\n"
b'If chkMostrarCaixa.Value = Checked Then\n'
b'    Grid.ColWidth(19) = 900\n'
b'    Grid.ColWidth(20) = 800\n'
b'    Grid.ColWidth(21) = 800\n'
b'Else\n'
b'    Grid.ColWidth(19) = 0\n'
b'    Grid.ColWidth(20) = 0\n'
b'    Grid.ColWidth(21) = 0\n'
b'End If\n'
b'End Sub'
)
reps.append(("chkMostrarCaixa_Click reescrito", o_chk, n_chk, 1))

# ==================================================================================
# 3) totais (Mostrar_Pedido): col 13(VALOR)->14, col 21(cancelado oculta)->23
# ==================================================================================
reps.append(("totais: If .TextMatrix(i, 21)->23", b'If .TextMatrix(i, 21) = "SIM" Then', b'If .TextMatrix(i, 23) = "SIM" Then', 1))
reps.append(("totais: CCur(.TextMatrix(i, 13))->14", b'CCur(.TextMatrix(i, 13))', b'CCur(.TextMatrix(i, 14))', 4))

# ==================================================================================
# 4) FlexCores: col 21(cancelado oculta)->23
# ==================================================================================
reps.append(("FlexCores: .TextMatrix(iLinha, 21)->23", b'.TextMatrix(iLinha, 21) = "SIM"', b'.TextMatrix(iLinha, 23) = "SIM"', 1))

# ==================================================================================
# 5) familia "Grid.Row, N)" espalhada por ~10 Subs - ordem DESCENDENTE de N antigo
# ==================================================================================
reps.append(("Grid.Row: 22->24 (NFCe oculta)", b"Grid.Row, 22)", b"Grid.Row, 24)", 4))
reps.append(("Grid.Row: 21->23 (Cancelado oculta)", b"Grid.Row, 21)", b"Grid.Row, 23)", 6))
reps.append(("Grid.Row: 20->22 (Reaberto oculta)", b"Grid.Row, 20)", b"Grid.Row, 22)", 2))
reps.append(("Grid.Row: 19->21 (C\xd3D/CX)", b"Grid.Row, 19)", b"Grid.Row, 21)", 6))
reps.append(("Grid.Row: 18->20 (CAIXA)", b"Grid.Row, 18)", b"Grid.Row, 20)", 5))
reps.append(("Grid.Row: 13->14 (VALOR)", b"Grid.Row, 13)", b"Grid.Row, 14)", 18))

# ==================================================================================
# aplica
# ==================================================================================
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
