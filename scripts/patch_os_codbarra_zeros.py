# -*- coding: utf-8 -*-
"""OS_Recapadora.frm: busca de produto por Cod.Barra (optCodBarra) tolera zeros a
   esquerda e tambem casa contra EAN - alinha com a folga do F2/PDV/Cadastro.
   .frm cp1252 -> edicao binaria + CRLF."""
import sys

p = r"C:\projeto\OrdemServico\Forms\OS_Recapadora.frm"
d = open(p, "rb").read().replace(b"\r\n", b"\n")

Q = b"'"   # aspas simples, pra montar os literais SQL sem dor de cabeca de escape

old = b"".join([
    b"        If optCodBarra.Value = True Then\n",
    b"            ' Formata com zeros apenas se for busca por C\xf3digo de Barras\n",
    b'            txtCodBarra.Text = Format(txtCodBarra.Text, "00000")\n',
    b"            \n",
    b'            sSQL = "SELECT codigo AS var_codprod, descricao AS var_desc, tamanho, REF, fabricante FROM produtos " & _\n',
    b'                   "WHERE (COD_BARRA = ' + Q + b'" & txtCodBarra.Text & "' + Q + b') AND (ativo = 1);"\n',
])

new = b"".join([
    b"        If optCodBarra.Value = True Then\n",
    b"            ' Busca por COD_BARRA ou EAN tolerando zeros a esquerda\n",
    b'            ' + b'\' (cadastrado "074130173026", digitado "74130173026") - mesma folga do F2/PDV.\n',
    b"            Dim sCB As String\n",
    b'            sCB = Replace(Trim(txtCodBarra.Text), "' + Q + b'", "' + Q + Q + b'")\n',
    b'            sSQL = "SELECT codigo AS var_codprod, descricao AS var_desc, tamanho, REF, fabricante FROM produtos " & _\n',
    b'                   "WHERE ativo = 1 AND (" & _\n',
    b'                   "COD_BARRA = ' + Q + b'" & sCB & "' + Q + b' OR EAN = ' + Q + b'" & sCB & "' + Q + b' " & _\n',
    b'                   "OR TRY_CONVERT(bigint, COD_BARRA) = TRY_CONVERT(bigint, ' + Q + b'" & sCB & "' + Q + b') " & _\n',
    b'                   "OR TRY_CONVERT(bigint, EAN) = TRY_CONVERT(bigint, ' + Q + b'" & sCB & "' + Q + b') " & _\n',
    b'                   "OR COD_BARRA = ' + Q + b'" & Format(sCB, "00000") & "' + Q + b');"\n',
])

if new in d:
    print("ja aplicado"); sys.exit(0)
if d.count(old) != 1:
    print("nao achou (count=%d)" % d.count(old)); sys.exit(1)
d = d.replace(old, new, 1)
open(p, "wb").write(d.replace(b"\n", b"\r\n"))
print("OK")
