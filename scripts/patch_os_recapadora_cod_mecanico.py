# -*- coding: utf-8 -*-
"""
Patch OS_Recapadora.frm:
1) Declara vCodMecanicoServ (variavel de modulo - sem novo controle,
   pois o usuario esta reduzindo objetos no form) para guardar o
   codigo do funcionario selecionado em cboMecanicoServ.
2) Adiciona cboMecanicoServ_GotFocus/_KeyPress/_LostFocus, espelhando
   exatamente o padrao ja usado em cboMecanico.
3) cmdAdicionarServicosAuto_Click: exige mecanico selecionado antes de
   adicionar, e grava cod_mecanico no INSERT INTO OS_Servicos_Auto
   (branches Automoveis/Motocicletas e Informatica/Celular).
4) MostrarGrid_Servicos: ao reabrir a OS, busca o cod_mecanico do
   ultimo servico gravado em OS_Servicos_Auto e preenche
   cboMecanicoServ com o nome correspondente.

Usa extracao por linha (lines[i]) para qualquer trecho que ja exista
no arquivo (evita retyping de acentos); so o texto NOVO (MsgBox, SQL,
nomes de variaveis) e digitado diretamente neste script UTF-8.
"""

PATH = r"C:\projeto\OrdemServico\Forms\OS_Recapadora.frm"

with open(PATH, "rb") as f:
    raw = f.read()

text = raw.decode("cp1252")
lines = text.split("\r\n")


def find_line(substr, start=0):
    for i in range(start, len(lines)):
        if substr in lines[i]:
            return i
    raise SystemExit(f"ERRO: ancora nao encontrada: {substr!r}")


def find_line_exact(s, start=0):
    for i in range(start, len(lines)):
        if lines[i] == s:
            return i
    raise SystemExit(f"ERRO: linha exata nao encontrada: {s!r}")


# ---------------------------------------------------------------
# 1) Declara a variavel de modulo
# ---------------------------------------------------------------
i = find_line("Dim printSQL As String")
lines[i] = lines[i] + "\r\nDim vCodMecanicoServ As String"

# ---------------------------------------------------------------
# 2) Eventos de cboMecanicoServ - insere logo depois de cboMecanico_LostFocus
# ---------------------------------------------------------------
start_lf = find_line_exact("Private Sub cboMecanico_LostFocus()")
end_lf = find_line_exact("End Sub", start_lf)

new_subs = []
new_subs.append("")
new_subs.append("Private Sub cboMecanicoServ_GotFocus()")
new_subs.append("Dim varNomeAntes As String")
new_subs.append("Dim varCodAntes As String")
new_subs.append("")
new_subs.append("varNomeAntes = cboMecanicoServ.Text")
new_subs.append("varCodAntes = vCodMecanicoServ")
new_subs.append("")
new_subs.append("cboMecanicoServ.Clear")
new_subs.append("")
new_subs.append('sSQL = "SELECT DISTINCT nome, codigo FROM funcionario order by nome;"')
new_subs.append("Set r = dbData.OpenRecordset(sSQL)")
new_subs.append("")
new_subs.append("Do While Not r.EOF")
new_subs.append('   cboMecanicoServ.AddItem r("nome")')
new_subs.append('   cboMecanicoServ.ItemData(cboMecanicoServ.NewIndex) = r("codigo")')
new_subs.append("   r.MoveNext")
new_subs.append("Loop")
new_subs.append("")
new_subs.append("If r.State <> 0 Then r.Close")
new_subs.append("Set r = Nothing")
new_subs.append("")
new_subs.append("cboMecanicoServ.Text = varNomeAntes")
new_subs.append("vCodMecanicoServ = varCodAntes")
new_subs.append("")
new_subs.append("moCombo.AttachTo cboMecanicoServ")
new_subs.append("End Sub")
new_subs.append("")
new_subs.append("Private Sub cboMecanicoServ_KeyPress(KeyAscii As Integer)")
new_subs.append("   KeyAscii = Asc(UCase(Chr(KeyAscii)))")
new_subs.append("End Sub")
new_subs.append("")
new_subs.append("Private Sub cboMecanicoServ_LostFocus()")
new_subs.append("   On Error GoTo TrataErro")
new_subs.append("   ")
new_subs.append('   If cboMecanicoServ.Text = "" Then vCodMecanicoServ = "": Exit Sub')
new_subs.append("   ")
new_subs.append("   vCodMecanicoServ = cboMecanicoServ.ItemData(cboMecanicoServ.ListIndex)")
new_subs.append("   Exit Sub")
new_subs.append("   ")
new_subs.append("TrataErro:")
new_subs.append("   If Err.Number = 381 Then Exit Sub")
new_subs.append("End Sub")

lines[end_lf + 1 : end_lf + 1] = new_subs

# ---------------------------------------------------------------
# 3a) cmdAdicionarServicosAuto_Click - validacao do mecanico
#     (ancora pela linha ASCII-safe seguinte, que ja existe no arquivo)
# ---------------------------------------------------------------
i = find_line('If mskValorServicoAuto.Text = "" Or mskValorServicoAuto.Text = "0,00" Then Exit Sub')
lines[i] = (
    lines[i]
    + "\r\n"
    + 'If vCodMecanicoServ = "" Then MsgBox "Selecione o mecânico que executou o serviço!", vbExclamation, "Aviso do Sistema": Exit Sub'
)

# ---------------------------------------------------------------
# 3b) INSERT INTO OS_Servicos_Auto - branch Automoveis/Motocicletas
#     (a primeira ocorrencia do INSERT, antes do branch Recapadora)
# ---------------------------------------------------------------
insert_marker = "dbData.Execute \"INSERT INTO OS_Servicos_Auto (codigo, cod_os, descricao, preco, quantidade, subtotal, desconto, total, data) VALUES (\" & _"

first_insert = find_line(insert_marker)
lines[first_insert] = lines[first_insert].replace(
    "(codigo, cod_os, descricao, preco, quantidade, subtotal, desconto, total, data) VALUES (",
    "(codigo, cod_os, descricao, preco, quantidade, subtotal, desconto, total, data, cod_mecanico) VALUES (",
)
# a linha seguinte (2 linhas abaixo do marker) fecha o INSERT com "', 103))"" -> adicionar cod_mecanico
close_line_1 = first_insert + 2
OLD_SUFFIX = "', 103))\""
NEW_SUFFIX = '\', 103), " & vCodMecanicoServ & ")"'
assert lines[close_line_1].rstrip().endswith(OLD_SUFFIX), lines[close_line_1]
assert lines[close_line_1].count(OLD_SUFFIX) == 1
lines[close_line_1] = lines[close_line_1].replace(OLD_SUFFIX, NEW_SUFFIX)

# ---------------------------------------------------------------
# 3c) INSERT INTO OS_Servicos_Auto - branch Informatica/Celular
#     (segunda ocorrencia do mesmo INSERT)
# ---------------------------------------------------------------
second_insert = find_line(insert_marker, first_insert + 1)
lines[second_insert] = lines[second_insert].replace(
    "(codigo, cod_os, descricao, preco, quantidade, subtotal, desconto, total, data) VALUES (",
    "(codigo, cod_os, descricao, preco, quantidade, subtotal, desconto, total, data, cod_mecanico) VALUES (",
)
close_line_2 = second_insert + 2
assert lines[close_line_2].rstrip().endswith(OLD_SUFFIX), lines[close_line_2]
assert lines[close_line_2].count(OLD_SUFFIX) == 1
lines[close_line_2] = lines[close_line_2].replace(OLD_SUFFIX, NEW_SUFFIX)

# ---------------------------------------------------------------
# 4) MostrarGrid_Servicos - preencher cboMecanicoServ ao reabrir a OS
# ---------------------------------------------------------------
start_mgs = find_line_exact("Private Sub MostrarGrid_Servicos()")
end_mgs = find_line_exact("End Sub", start_mgs)
assert lines[end_mgs - 1] == "Set r = Nothing", lines[end_mgs - 1]

# reaproveita a condicao de vTipoOS ja usada nas linhas seguintes desta mesma sub
# (pula a variante comentada com ' que aparece antes da ativa)
vtipo_cond_line = None
for i in range(start_mgs, end_mgs):
    stripped = lines[i].strip()
    if stripped.startswith("If vTipoOS") and "Celular" in stripped:
        vtipo_cond_line = i
        break
assert vtipo_cond_line is not None, "condicao vTipoOS ativa nao encontrada"
vtipo_cond = lines[vtipo_cond_line]

new_block = []
new_block.append("")
new_block.append("Dim vCodMecServAtual As String")
new_block.append('vCodMecServAtual = ""')
new_block.append(vtipo_cond.strip())
new_block.append(
    '    sSQL = "SELECT TOP 1 cod_mecanico FROM OS_Servicos_Auto WHERE (cod_os = " & txtCodOS.Text & ") AND (cod_mecanico IS NOT NULL) ORDER BY codigo DESC;"'
)
new_block.append("    Set r = dbData.OpenRecordset(sSQL)")
new_block.append('    If Not r.EOF Then vCodMecServAtual = ValidateNull(r("cod_mecanico"))')
new_block.append("    If r.State <> 0 Then r.Close")
new_block.append("    Set r = Nothing")
new_block.append("")
new_block.append('    If vCodMecServAtual <> "" Then')
new_block.append('        sSQL = "SELECT nome FROM funcionario WHERE (codigo = " & vCodMecServAtual & ");"')
new_block.append("        Set r = dbData.OpenRecordset(sSQL)")
new_block.append("        If Not r.EOF Then")
new_block.append('            cboMecanicoServ.Text = r("nome")')
new_block.append("            vCodMecanicoServ = vCodMecServAtual")
new_block.append("        Else")
new_block.append('            cboMecanicoServ.Text = ""')
new_block.append('            vCodMecanicoServ = ""')
new_block.append("        End If")
new_block.append("        If r.State <> 0 Then r.Close")
new_block.append("        Set r = Nothing")
new_block.append("    Else")
new_block.append('        cboMecanicoServ.Text = ""')
new_block.append('        vCodMecanicoServ = ""')
new_block.append("    End If")
new_block.append("End If")

lines[end_mgs:end_mgs] = new_block

# ---------------------------------------------------------------
# Grava
# ---------------------------------------------------------------
out_text = "\r\n".join(lines)
out = out_text.encode("cp1252")
out = out.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

with open(PATH, "wb") as f:
    f.write(out)

print("OK - patch aplicado")
print("bytes originais:", len(raw), "bytes finais:", len(out))
