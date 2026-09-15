# -*- coding: utf-8 -*-
"""PDV.frm - pedido do usuario 2026-09-15: reabrir uma venda com frete nao trazia o valor
do frete de volta pro campo txtFrete.

Investigacao achou o mesmo bug tambem pro ACRESCIMO (txtAcresc), nos MESMOS 5 lugares -
todos os 3 handlers que recarregam os campos ao reabrir um pedido (`cmdFinalizaravista_Click`
x2 ramos, `cmdFinalizarPrazo_Click` x2 ramos, `cmdOr�amento_Click` x1) fazem:

    txtAcresc.Text = FormatNumber(0, 2)   ' reset ANTES de saber se e reabertura
    txtFrete.Text = FormatNumber(0, 2)
    ...
    If lblEstornar.Caption = "ESTORNO" Then   ' e reabertura
        ...
        txtDesc.Text = FormatNumber(r("VALOR_DESC"), 2)   ' desconto recarrega...
        ...                                                ' ...acrescimo e frete NUNCA
    End If

Fix: logo apos cada `txtDesc.Text = FormatNumber(r("VALOR_DESC"), N)`, recarrega tambem
TIPO_ACRESCIMO/VALOR_ACRESCIMO (mesmo padrao ja usado pro TIPO_DESC/VALOR_DESC no mesmo
bloco) e ValorFreteReal (sem opt de tipo - frete no PDV e sempre R$, sem % - ver comentario
em cmdFinalizar_Click "aplica o frete (txtFrete, sempre R$, sem opcao de %)"). ValidateNull
pra NULL virar 0 em vez de estourar erro 13 no FormatNumber - pedido antigo pode nao ter
ValorFreteReal setado.

Trabalha por NUMERO DE LINHA (nao por padrao de texto) porque 3 dos 5 blocos sao
byte-identicos entre si - insercao de baixo pra cima pra nao invalidar os indices.

.frm cp1252 -> edicao binaria + normaliza CRLF.
"""
import sys

p = r"C:\projeto\PDV\Forms\PDV.frm"
raw = open(p, "rb").read().replace(b"\r\n", b"\n").replace(b"\r", b"\n")
text = raw.decode("cp1252")
lines = text.split("\n")

# (numero da linha 1-based, indentacao, precisao do FormatNumber ja existente - so validacao)
alvos = [
    (9176, "                ", 2),
    (9296, "                ", 3),
    (9494, "                ", 2),
    (9624, "                ", 2),
    (9926, "        ", 2),
]

def bloco_novo(indent):
    return [
        indent + 'If r("TIPO_ACRESCIMO") = "R" Then',
        indent + "    optAscrescRS.Value = True",
        indent + "Else",
        indent + "    optAscrescPorc.Value = True",
        indent + "End If",
        indent + 'txtAcresc.Text = FormatNumber(ValidateNull(r("VALOR_ACRESCIMO")), 2)',
        indent + 'txtFrete.Text = FormatNumber(ValidateNull(r("ValorFreteReal")), 2)',
    ]

# processa de baixo pra cima pra os indices dos alvos anteriores nao mudarem
for lineno, indent, precisao in sorted(alvos, key=lambda t: -t[0]):
    idx = lineno - 1
    esperado = '%stxtDesc.Text = FormatNumber(r("VALOR_DESC"), %d)' % (indent, precisao)
    atual = lines[idx]
    if atual != esperado:
        print("[ABORTA] linha %d nao bate.\n  esperado: %r\n  atual:    %r" % (lineno, esperado, atual))
        sys.exit(1)
    novas = bloco_novo(indent)
    # ja aplicado? (checagem simples: proxima linha nao-vazia ja e o bloco novo)
    if lines[idx+1] == novas[0]:
        print("[ja] linha %d" % lineno)
        continue
    lines[idx+1:idx+1] = novas
    print("[ok] linha %d (+%d linhas)" % (lineno, len(novas)))

novo_texto = "\n".join(lines)
open(p, "wb").write(novo_texto.encode("cp1252").replace(b"\n", b"\r\n"))
print("gravado " + p)
