# -*- coding: utf-8 -*-
"""
Implementa as correcoes encontradas na analise de cmdFinalizar_Click:

1) Remove validacao duplicada de txtCodPedido.Text = "" (a 2a nunca era
   alcancada, pois a 1a ja sai da sub antes).
2) Corrige espaco faltando em "status_pedido = 1" no bloco A VISTA (evita
   depender do tokenizer do SQL Server para separar "1" de "WHERE").
3) Adiciona "cmdFinalizar.Enabled = True" antes dos Exit Sub que o
   usuario pode escolher (cancelar confirmacao, cliente nao identificado)
   - sem isso o botao ficava aparentemente travado ate o usuario clicar
   de novo em Finalizar AV/AP.
4) Envolve os 2 blocos de escrita (A PRAZO e A VISTA) em transacao
   (BEGIN/COMMIT/ROLLBACK), com On Error Goto + captura de Err.Description
   antes de qualquer outra coisa no handler, seguindo o padrao do projeto
   (CLAUDE.md).
5) Pre-busca os codigos de parcela (Autonumeracao_Parcelas, que abre um
   OpenRecordset) ANTES de iniciar a transacao, substituindo as chamadas
   dentro dos loops por aritmetica simples (lNovoCodBase + N) - evita o
   deadlock de abrir recordset ADODB dentro de transacao DAO que o
   CLAUDE.md alerta.
"""

PATH = r"C:\projeto\OrdemServico\Forms\OS_Recapadora.frm"

with open(PATH, "rb") as f:
    raw = f.read()

text = raw.decode("cp1252")
lines = text.split("\r\n")


def find_line_exact(s, start=0, end=None):
    end = end if end is not None else len(lines)
    for i in range(start, end):
        if lines[i] == s:
            return i
    raise SystemExit(f"ERRO: linha exata nao encontrada: {s!r}")


def find_line(substr, start=0, end=None):
    end = end if end is not None else len(lines)
    for i in range(start, end):
        if substr in lines[i]:
            return i
    raise SystemExit(f"ERRO: ancora nao encontrada: {substr!r}")


start = find_line_exact("Private Sub cmdFinalizar_Click()")
end = find_line_exact("End Sub", start)

# ---------------------------------------------------------------
# 1) Remove validacao duplicada de txtCodPedido (nunca alcancada)
# ---------------------------------------------------------------
i = find_line_exact('If txtCodPedido.Text = "" Then MsgBox "Cód. Pedido em Branco": Exit Sub', start, end)
del lines[i]
end = find_line_exact("End Sub", start)  # end mudou (1 linha a menos)

# ---------------------------------------------------------------
# 2) Adiciona Dim bTrans / lNovoCodBase
# ---------------------------------------------------------------
i = find_line_exact("Dim lNovoCod As Long", start, end)
lines[i] = lines[i] + "\r\nDim lNovoCodBase As Long\r\nDim bTrans As Boolean"
end = find_line_exact("End Sub", start)

# ---------------------------------------------------------------
# 3) On Error Goto antes do bloco A PRAZO / A VISTA
# ---------------------------------------------------------------
i = find_line_exact('If cboTipoPgto.Text = "À PRAZO" Then', start, end)
lines[i - 1] = lines[i - 1] + "\r\nbTrans = False\r\nOn Error GoTo TrataErroFinalizar"
end = find_line_exact("End Sub", start)

# ---------------------------------------------------------------
# 4) Reenable do botao antes dos Exit Sub "manuais" (cancelamento do usuario)
# ---------------------------------------------------------------
i = find_line('If txtCodCliente = "1" Then MsgBox "IDENTIFIQUE O CLIENTE DA COMPRA!"', start, end)
lines[i] = lines[i].replace(": Exit Sub", ": cmdFinalizar.Enabled = True: Exit Sub")

# ha 2 ocorrencias identicas do ShowMsg de confirmacao (A Prazo e A Vista) - tratar as duas
occ = [
    k
    for k in range(start, end)
    if 'If ShowMsg("Deseja finalizar essa compra?"' in lines[k] and "Then Exit Sub" in lines[k]
]
assert len(occ) == 2, occ
for k in occ:
    lines[k] = lines[k].replace("Then Exit Sub", "Then cmdFinalizar.Enabled = True: Exit Sub")

# ---------------------------------------------------------------
# 5) Corrige espaco faltando em "status_pedido = 1" (bloco A VISTA)
# ---------------------------------------------------------------
i = find_line_exact('              "status_pedido = 1" & _', start, end)
lines[i] = '              "status_pedido = 1 " & _'

# ---------------------------------------------------------------
# 6) Bloco A PRAZO: pre-busca lNovoCodBase + BEGIN TRANSACTION
# ---------------------------------------------------------------
i = find_line("'ATUALIZAR A TABELA OS", start, end)
first_atualizar = i
lines[i] = (
    "        lNovoCodBase = Autonumeracao_Parcelas\r\n"
    "        dbData.Execute \"BEGIN TRANSACTION\"\r\n"
    "        bTrans = True\r\n"
    + lines[i]
)
end = find_line_exact("End Sub", start)

# ---------------------------------------------------------------
# 7) Bloco A PRAZO: substitui as chamadas de Autonumeracao_Parcelas
#    (jah nao ha mais leitura dentro da transacao). 3 ocorrencias em
#    sequencia: entrada (COM ENTRADA), loop continuacao (COM ENTRADA),
#    loop (SEM ENTRADA).
# ---------------------------------------------------------------
i = find_line("lNovoCod = Autonumeracao_Parcelas", start, end)
indent = lines[i][: len(lines[i]) - len(lines[i].lstrip())]
lines[i] = indent + "lNovoCod = lNovoCodBase"

i = find_line("lNovoCod = Autonumeracao_Parcelas", i + 1, end)
indent = lines[i][: len(lines[i]) - len(lines[i].lstrip())]
lines[i] = indent + "lNovoCod = lNovoCodBase + i"

i = find_line("lNovoCod = Autonumeracao_Parcelas", i + 1, end)
indent = lines[i][: len(lines[i]) - len(lines[i].lstrip())]
lines[i] = indent + "lNovoCod = lNovoCodBase + i - 1"

# ---------------------------------------------------------------
# 8) Bloco A PRAZO: COMMIT TRANSACTION antes da impressao
# ---------------------------------------------------------------
i = find_line("If iCopiasAP <> 0 Then", start, end)
indent = lines[i][: len(lines[i]) - len(lines[i].lstrip())]
lines[i] = (
    indent + 'dbData.Execute "COMMIT TRANSACTION"\r\n'
    + indent + "bTrans = False\r\n\r\n"
    + lines[i]
)
end = find_line_exact("End Sub", start)

# ---------------------------------------------------------------
# 9) Bloco A VISTA: pre-busca lNovoCodBase + BEGIN TRANSACTION
# ---------------------------------------------------------------
i = find_line("'ATUALIZAR A TABELA OS", start, end)  # 2a ocorrencia (a 1a foi consumida - ja virou parte de outra linha)
lines[i] = (
    "            lNovoCodBase = Autonumeracao_Parcelas\r\n"
    "            dbData.Execute \"BEGIN TRANSACTION\"\r\n"
    "            bTrans = True\r\n"
    + lines[i]
)
end = find_line_exact("End Sub", start)

# ---------------------------------------------------------------
# 10) Bloco A VISTA: substitui as 3 chamadas de Autonumeracao_Parcelas
#     (1 - FORMA; 2 - FORMAS parcela 1; 2 - FORMAS parcela 2)
# ---------------------------------------------------------------
i = find_line("lNovoCod = Autonumeracao_Parcelas", start, end)
indent = lines[i][: len(lines[i]) - len(lines[i].lstrip())]
lines[i] = indent + "lNovoCod = lNovoCodBase"

i = find_line("lNovoCod = Autonumeracao_Parcelas", i + 1, end)
indent = lines[i][: len(lines[i]) - len(lines[i].lstrip())]
lines[i] = indent + "lNovoCod = lNovoCodBase"

i = find_line("lNovoCod = Autonumeracao_Parcelas", i + 1, end)
indent = lines[i][: len(lines[i]) - len(lines[i].lstrip())]
lines[i] = indent + "lNovoCod = lNovoCodBase + 1"

# ---------------------------------------------------------------
# 11) Bloco A VISTA: COMMIT TRANSACTION antes da impressao
# ---------------------------------------------------------------
i = find_line("If iCopiasAP <> 0 Then", start, end)
indent = lines[i][: len(lines[i]) - len(lines[i].lstrip())]
lines[i] = (
    indent + 'dbData.Execute "COMMIT TRANSACTION"\r\n'
    + indent + "bTrans = False\r\n\r\n"
    + lines[i]
)
end = find_line_exact("End Sub", start)

# ---------------------------------------------------------------
# 12) Handler de erro no fim da sub (antes do End Sub)
# ---------------------------------------------------------------
i = find_line_exact("cmdFinalizar.Enabled = True", start, end)
lines[i] = (
    "cmdFinalizar.Enabled = True\r\n"
    "\r\n"
    "Exit Sub\r\n"
    "\r\n"
    "TrataErroFinalizar:\r\n"
    "Dim sErroFinalizar As String\r\n"
    "sErroFinalizar = Err.Description\r\n"
    "If bTrans Then dbData.Execute \"ROLLBACK TRANSACTION\"\r\n"
    "cmdFinalizar.Enabled = True\r\n"
    'MsgBox "Não foi possível finalizar a venda:" & vbCrLf & sErroFinalizar, vbCritical, "Erro ao Finalizar"'
)

out_text = "\r\n".join(lines)
out = out_text.encode("cp1252")
with open(PATH, "wb") as f:
    f.write(out)

print("OK - patch aplicado")
