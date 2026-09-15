# -*- coding: utf-8 -*-
"""Estonar.frm - pedido do usuario 2026-09-15:
1. Coluna FRETE (pedidos.ValorFreteReal) depois de ACRESC., antes de VALOR.
2. Coluna MAQUINA (pedidos.MAQUINA) antes de CAIXA.
3. chkMostrarCaixa (ja existia no form, sem handler) mostra/esconde MAQUINA/CAIXA/CODCAIXA.
4. chkMostrarCaixa comeca desmarcado (False) ao abrir o form.

Decisao de implementacao: NAO renumerei as colunas existentes (14 Reaberto...22 NFCe oculta
ficariam 16...24, tocando ~15 lugares diferentes que ja leem essas colunas por indice -
risco alto de esquecer um). Em vez disso, as 2 colunas novas entram no FIM (indices logicos
23=FRETE, 24=MAQUINA) e usam `Grid.ColPosition` (propriedade padrao do MSFlexGrid, doc da
propria Microsoft) pra aparecer visualmente no lugar certo - TextMatrix/CellPicture/Col
sempre endercam por indice LOGICO, nao por posicao visual, entao nenhum dos ~15 lugares que
ja leem colunas 14-22 precisa mudar. CAIXA(18)/C�D-CX(19) tambem mantem o indice logico -
so a POSICAO visual delas desloca (efeito automatico do ColPosition), o codigo que hoje
mexe nelas (nenhum, so leitura de FormatarGrid_Pedido) nao precisa de ajuste.

.frm cp1252 -> edicao binaria + normaliza CRLF.
"""
import sys

p = r"C:\projeto\PDV\Forms\Estonar.frm"
d = open(p, "rb").read().replace(b"\r\n", b"\n").replace(b"\r", b"\n")

reps = []  # (nome, old, new, esperado)

# ---------------------------------------------------- .Cols 23 -> 25
reps.append(("Cols 23->25", b"      .Cols = 23\n", b"      .Cols = 25\n", 1))

# ---------------------------------------------------- ColWidth: 18/19 viram condicionais + 23/24 novas
o = (
b"      .ColWidth(18) = 800\n"
b"      .ColWidth(19) = 800\n"
b"      .ColWidth(20) = 0\n"
b"      .ColWidth(21) = 0\n"
b"      .ColWidth(22) = 0\n"
)
n = (
b"      .ColWidth(18) = IIf(chkMostrarCaixa.Value = Checked, 800, 0)\n"
b"      .ColWidth(19) = IIf(chkMostrarCaixa.Value = Checked, 800, 0)\n"
b"      .ColWidth(20) = 0\n"
b"      .ColWidth(21) = 0\n"
b"      .ColWidth(22) = 0\n"
b"      .ColWidth(23) = 900\n"
b"      .ColWidth(24) = IIf(chkMostrarCaixa.Value = Checked, 900, 0)\n"
b"      \n"
b"      'FRETE entra na posicao 13 (entre ACRESC. e VALOR); MAQUINA na 19 (antes de CAIXA) -\n"
b"      'ColPosition so muda onde aparece na tela, TextMatrix/Col continuam pelo indice logico\n"
b"      .ColPosition(23) = 13\n"
b"      .ColPosition(24) = 19\n"
)
reps.append(("ColWidth 18/19 condicional + novas colunas 23/24 + ColPosition", o, n, 1))

# ---------------------------------------------------- headers
o = b'      .TextMatrix(0, 19) = "C\xd3D/CX"\n'
n = b'      .TextMatrix(0, 19) = "C\xd3D/CX"\n      .TextMatrix(0, 23) = "FRETE"\n      .TextMatrix(0, 24) = "MAQUINA"\n'
reps.append(("headers FRETE/MAQUINA", o, n, 1))

# ---------------------------------------------------- fill loop
o = b'            .TextMatrix(.Rows - 1, 19) = ValidateNull(rTabela("VarPEDCODCAIXA"))\n'
n = (
b'            .TextMatrix(.Rows - 1, 19) = ValidateNull(rTabela("VarPEDCODCAIXA"))\n'
b'            .TextMatrix(.Rows - 1, 23) = Format(rTabela("varPedFrete"), ocMONEY)\n'
b'            .TextMatrix(.Rows - 1, 24) = ValidateNull(rTabela("varPedMaquina"))\n'
)
reps.append(("fill loop FRETE/MAQUINA", o, n, 1))

# ---------------------------------------------------- SQL: novos campos em todos os ramos (13 vivos + 2 comentario, inofensivo)
o = b"pedidos.caixa as varPedCaixa"
n = b"pedidos.ValorFreteReal as varPedFrete, pedidos.MAQUINA as varPedMaquina, pedidos.caixa as varPedCaixa"
reps.append(("SQL: ValorFreteReal/MAQUINA em todos os ramos", o, n, 15))

# ---------------------------------------------------- chkMostrarCaixa_Click (novo handler)
# ancora: logo depois de FormatarGrid_Pedido, antes de SomaGrid
o = b"\nPublic Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Currency\n"
n = (
b"\nPrivate Sub chkMostrarCaixa_Click()\n"
b"'mostra/esconde MAQUINA, CAIXA e C\xd3D/CX (colunas 24, 18, 19) sem precisar reconsultar\n"
b"If chkMostrarCaixa.Value = Checked Then\n"
b"    Grid.ColWidth(24) = 900\n"
b"    Grid.ColWidth(18) = 800\n"
b"    Grid.ColWidth(19) = 800\n"
b"Else\n"
b"    Grid.ColWidth(24) = 0\n"
b"    Grid.ColWidth(18) = 0\n"
b"    Grid.ColWidth(19) = 0\n"
b"End If\n"
b"End Sub\n"
b"\nPublic Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Currency\n"
)
reps.append(("novo Sub chkMostrarCaixa_Click", o, n, 1))

# ---------------------------------------------------- Form_Load: chkMostrarCaixa comeca desmarcado
o = b'CAIXA_FECHADO = True\ntxtCodPedidoCerto.Text = ""\n'
n = b'CAIXA_FECHADO = True\nchkMostrarCaixa.Value = Unchecked\ntxtCodPedidoCerto.Text = ""\n'
reps.append(("Form_Load: chkMostrarCaixa default False", o, n, 1))

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
