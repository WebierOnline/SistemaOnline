# -*- coding: utf-8 -*-
"""
Remove o debito duplicado de estoque em cmdFinalizar_Click (2 ocorrencias
do mesmo loop). Esse Sub e' o commit real disparado por
cmdFinalizarAV_Click/cmdFinalizarAP_Click (via frmVendaFechamento +
cmdFinalizar ou Enter em txtRecebido/cboQuantForma) - NAO e' codigo
morto. Como agora o estoque ja e debitado no momento de adicionar a peca
(cmdAdicionarPecas_Click) e devolvido ao remover (cmdRemoverPecas_Click),
este loop debitaria de novo ao finalizar, duplicando o desconto.

Comenta o bloco (mesmo padrao ja usado no arquivo para codigo desativado,
com marcador explicando o motivo), sem apagar a logica original.
"""

PATH = r"C:\projeto\OrdemServico\Forms\OS_Recapadora.frm"

with open(PATH, "rb") as f:
    raw = f.read()

text = raw.decode("cp1252")
lines = text.split("\r\n")

OLD_BLOCK = [
    "        'Retirar da tabela PRODUTOS as QUANTIDADES mencionadas no grid",
    "        For i = 1 To Grid_Servicos.Rows - 1  'analizar essa linha",
    '            If Grid_Servicos.TextMatrix(i, 2) = "PRODUTO" Then',
    '                dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque - " & Replace(CDbl(Grid_Servicos.TextMatrix(i, 5)), ",", ".") & " WHERE (codigo = " & Grid_Servicos.TextMatrix(i, 10) & ");"',
    "            End If",
    "        Next",
]

NEW_BLOCK = [
    "        'DESATIVEI 2026-07-07: estoque ja e debitado em cmdAdicionarPecas_Click (e devolvido em cmdRemoverPecas_Click);",
    "        'manter esse loop aqui causaria debito duplicado ao finalizar.",
    "        'Retirar da tabela PRODUTOS as QUANTIDADES mencionadas no grid",
    "        'For i = 1 To Grid_Servicos.Rows - 1  'analizar essa linha",
    '        \'    If Grid_Servicos.TextMatrix(i, 2) = "PRODUTO" Then',
    '        \'        dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque - " & Replace(CDbl(Grid_Servicos.TextMatrix(i, 5)), ",", ".") & " WHERE (codigo = " & Grid_Servicos.TextMatrix(i, 10) & ");"',
    "        '    End If",
    "        'Next",
]

# encontra as 2 ocorrencias do bloco (6 linhas cada) e substitui pelas 8 linhas comentadas
found_positions = []
n = len(OLD_BLOCK)
i = 0
while i <= len(lines) - n:
    if lines[i : i + n] == OLD_BLOCK:
        found_positions.append(i)
        i += n
    else:
        i += 1

assert len(found_positions) == 2, found_positions

# substitui de tras pra frente pra nao bagunçar os indices
for pos in reversed(found_positions):
    lines[pos : pos + n] = NEW_BLOCK

out_text = "\r\n".join(lines)
out = out_text.encode("cp1252")
with open(PATH, "wb") as f:
    f.write(out)

print("OK - patch aplicado, blocos comentados:", len(found_positions))
