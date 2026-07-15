# -*- coding: utf-8 -*-
"""
OS_Consulta_Pecas.frm: quando chkCompartibilidade = True, filtra o grid
para mostrar somente pecas compativeis com o veiculo da OS atual
(lido de OS_Recapadora.cboModelo/txtAno). Regras de parsing de
produtos_comp.modelo/ano conforme Arquivos\\mod25.txt:
 - modelo pode ter varios carros separados por "/" (compara por "contem")
 - ano vazio = qualquer ano
 - ano "NN>" ou "NNNN>" = a partir daquele ano (aberto)
 - ano "NN/NN", "NN-NN" (2 ou 4 digitos, mistura) = intervalo
 - ano com um numero isolado = ano exato
 - ano nao numerico/nao interpretavel = nao bloqueia (trata como compativel)
 - ano de 2 digitos: <=30 -> 20XX, >30 -> 19XX (convencao padrao)
"""

PATH = r"C:\projeto\OrdemServico\Forms\OS_Consulta_Pecas.frm"

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


# ---------------------------------------------------------------
# 1) Adiciona Dim's novos logo apos "Dim sModelo As String"
# ---------------------------------------------------------------
i_dim = find_line_exact("   Dim sModelo As String")
lines[i_dim + 1 : i_dim + 1] = [
    "   Dim bCompativel As Boolean",
    "   Dim sModeloOS As String",
    "   Dim iAnoOS As Long",
    "",
    "   sModeloOS = Trim(OS_Recapadora.cboModelo.Text)",
    "   iAnoOS = Val(OS_Recapadora.txtAno.Text)",
]

# ---------------------------------------------------------------
# 2) Reescreve o loop principal do Formatar_Grid
# ---------------------------------------------------------------
old_block = [
    "      If Not rTabela Is Nothing Then",
    "         Do While Not rTabela.EOF",
    "            'ALINHAMENTO",
    "            .ColAlignment(2) = 1",
    "            ",
    '            .TextMatrix(.Rows - 1, 1) = rTabela("var_cod")',
    '            .TextMatrix(.Rows - 1, 2) = rTabela("var_codbarra")',
    '            .TextMatrix(.Rows - 1, 3) = rTabela("var_desc")',
    '            .TextMatrix(.Rows - 1, 4) = ValidateNull(rTabela("var_fab"))',
    "            ",
    '            .TextMatrix(.Rows - 1, 6) = rTabela("var_med")',
    '            .TextMatrix(.Rows - 1, 7) = ValidateNull(rTabela("var_prat"))',
    '            .TextMatrix(.Rows - 1, 8) = rTabela("var_quant")',
    '            .TextMatrix(.Rows - 1, 9) = Format$(rTabela("venda"), ocMONEY)',
    '            .TextMatrix(.Rows - 1, 10) = Format$(rTabela("custo"), ocMONEY)',
    "            ",
    "            If chkCompartibilidade.Value = 1 Then",
    '                sSQL = "SELECT modelo, ano FROM produtos_comp WHERE (cod_produto = " & rTabela("var_cod") & ");"',
    "                Set r2 = dbData.OpenRecordset(sSQL)",
    '                var_Comp = ""',
    "                Do While Not r2.EOF",
    '                   sModelo = Trim(r2("modelo"))',
    '                   If Left(sModelo, 1) = "/" Then sModelo = Trim(Mid(sModelo, 2))',
    '                   var_Comp = var_Comp & sModelo & "(" & r2("ano") & "), "',
    "                   r2.MoveNext",
    "                Loop",
    "                If Len(var_Comp) > 0 Then var_Comp = Left(var_Comp, Len(var_Comp) - 2) ' Limpa a \xfaltima v\xedrgula",
    "                .TextMatrix(.Rows - 1, 5) = var_Comp",
    "                If r2.State <> 0 Then r2.Close",
    "                Set r2 = Nothing",
    "            End If",
    "            ",
    '            var_Comp = ""',
    "            ",
    "            rTabela.MoveNext",
    "            .Rows = .Rows + 1",
    "         Loop",
    "      End If",
]

new_block = [
    "      If Not rTabela Is Nothing Then",
    "         Do While Not rTabela.EOF",
    '            var_Comp = ""',
    "            bCompativel = True",
    "            ",
    "            If chkCompartibilidade.Value = 1 Then",
    '                sSQL = "SELECT modelo, ano FROM produtos_comp WHERE (cod_produto = " & rTabela("var_cod") & ");"',
    "                Set r2 = dbData.OpenRecordset(sSQL)",
    "                bCompativel = False",
    "                Do While Not r2.EOF",
    '                   sModelo = Trim(r2("modelo"))',
    '                   If Left(sModelo, 1) = "/" Then sModelo = Trim(Mid(sModelo, 2))',
    '                   var_Comp = var_Comp & sModelo & "(" & r2("ano") & "), "',
    "                   If Not bCompativel Then",
    '                      If VerificaModeloCompativel(sModelo, sModeloOS) And VerificaAnoCompativel(Trim(ValidateNull(r2("ano"))), iAnoOS) Then',
    "                         bCompativel = True",
    "                      End If",
    "                   End If",
    "                   r2.MoveNext",
    "                Loop",
    "                If Len(var_Comp) > 0 Then var_Comp = Left(var_Comp, Len(var_Comp) - 2) ' Limpa a \xfaltima v\xedrgula",
    "                If r2.State <> 0 Then r2.Close",
    "                Set r2 = Nothing",
    "            End If",
    "            ",
    "            If bCompativel Then",
    "               'ALINHAMENTO",
    "               .ColAlignment(2) = 1",
    "               ",
    '               .TextMatrix(.Rows - 1, 1) = rTabela("var_cod")',
    '               .TextMatrix(.Rows - 1, 2) = rTabela("var_codbarra")',
    '               .TextMatrix(.Rows - 1, 3) = rTabela("var_desc")',
    '               .TextMatrix(.Rows - 1, 4) = ValidateNull(rTabela("var_fab"))',
    "               .TextMatrix(.Rows - 1, 5) = var_Comp",
    '               .TextMatrix(.Rows - 1, 6) = rTabela("var_med")',
    '               .TextMatrix(.Rows - 1, 7) = ValidateNull(rTabela("var_prat"))',
    '               .TextMatrix(.Rows - 1, 8) = rTabela("var_quant")',
    '               .TextMatrix(.Rows - 1, 9) = Format$(rTabela("venda"), ocMONEY)',
    '               .TextMatrix(.Rows - 1, 10) = Format$(rTabela("custo"), ocMONEY)',
    "               .Rows = .Rows + 1",
    "            End If",
    "            ",
    "            rTabela.MoveNext",
    "         Loop",
    "      End If",
]

i_start = find_line_exact(old_block[0])
for k, l in enumerate(old_block):
    assert lines[i_start + k] == l, (i_start + k, repr(lines[i_start + k]), repr(l))

lines[i_start : i_start + len(old_block)] = new_block

# ---------------------------------------------------------------
# 3) Adiciona as 3 funcoes auxiliares no final do modulo de codigo,
#    logo apos o "End Sub" do Formatar_Grid
# ---------------------------------------------------------------
i_formatar_end = find_line_exact("Private Sub Formatar_Grid(rTabela As ADODB.Recordset)")
i_end_sub = find_line_exact("End Sub", i_formatar_end)

funcoes = [
    "",
    "Private Function VerificaModeloCompativel(sModeloCampo As String, sModeloOS As String) As Boolean",
    "    Dim arr() As String",
    "    Dim i As Integer",
    "    Dim sToken As String",
    "",
    '    If Trim(sModeloOS) = "" Then',
    "        VerificaModeloCompativel = True",
    "        Exit Function",
    "    End If",
    "",
    '    arr = Split(sModeloCampo, "/")',
    "    For i = 0 To UBound(arr)",
    "        sToken = Trim(arr(i))",
    '        If sToken <> "" Then',
    "            If InStr(1, sToken, sModeloOS, vbTextCompare) > 0 Then",
    "                VerificaModeloCompativel = True",
    "                Exit Function",
    "            End If",
    "        End If",
    "    Next i",
    "    VerificaModeloCompativel = False",
    "End Function",
    "",
    "Private Function VerificaAnoCompativel(sAnoCampo As String, iAnoOS As Long) As Boolean",
    "    Dim sA As String",
    "    Dim sSep As String",
    "    Dim partes() As String",
    "    Dim iA1 As Long, iA2 As Long",
    "",
    "    sA = Trim(sAnoCampo)",
    "",
    '    If sA = "" Or iAnoOS = 0 Then',
    "        VerificaAnoCompativel = True",
    "        Exit Function",
    "    End If",
    "",
    '    If Right(sA, 1) = ">" Then',
    "        iA1 = NormalizaAno(Left(sA, Len(sA) - 1))",
    "        If iA1 = 0 Then",
    "            VerificaAnoCompativel = True",
    "        Else",
    "            VerificaAnoCompativel = (iAnoOS >= iA1)",
    "        End If",
    "        Exit Function",
    "    End If",
    "",
    '    If InStr(sA, "/") > 0 Then',
    '        sSep = "/"',
    '    ElseIf InStr(sA, "-") > 0 Then',
    '        sSep = "-"',
    "    Else",
    '        sSep = ""',
    "    End If",
    "",
    '    If sSep <> "" Then',
    "        partes = Split(sA, sSep)",
    "        If UBound(partes) = 1 Then",
    "            iA1 = NormalizaAno(partes(0))",
    "            iA2 = NormalizaAno(partes(1))",
    "            If iA1 > 0 And iA2 > 0 Then",
    "                VerificaAnoCompativel = (iAnoOS >= iA1 And iAnoOS <= iA2)",
    "                Exit Function",
    "            End If",
    "        End If",
    "        VerificaAnoCompativel = True",
    "        Exit Function",
    "    End If",
    "",
    "    If IsNumeric(sA) Then",
    "        iA1 = NormalizaAno(sA)",
    "        If iA1 > 0 Then",
    "            VerificaAnoCompativel = (iAnoOS = iA1)",
    "        Else",
    "            VerificaAnoCompativel = True",
    "        End If",
    "        Exit Function",
    "    End If",
    "",
    "    VerificaAnoCompativel = True",
    "End Function",
    "",
    "Private Function NormalizaAno(sVal As String) As Long",
    "    Dim v As String",
    "    Dim n As Long",
    "",
    "    v = Trim(sVal)",
    "    If Not IsNumeric(v) Then",
    "        NormalizaAno = 0",
    "        Exit Function",
    "    End If",
    "",
    "    n = CLng(v)",
    "    If Len(v) <= 2 Then",
    "        If n <= 30 Then",
    "            n = 2000 + n",
    "        Else",
    "            n = 1900 + n",
    "        End If",
    "    End If",
    "    NormalizaAno = n",
    "End Function",
]

lines[i_end_sub + 1 : i_end_sub + 1] = funcoes

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK - filtro de compatibilidade adicionado")
