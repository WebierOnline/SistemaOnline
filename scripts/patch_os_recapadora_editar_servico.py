# -*- coding: utf-8 -*-
"""
Patch OS_Recapadora.frm: implementa edicao de servico (cmdEditarServicosAuto).

1) Declara vCodItemServicoEditando (variavel de modulo - guarda o codigo
   do item de OS_Servicos_Auto sendo editado; "" quando nao esta editando).
2) Grid_Servicos_DblClick: ao dar duplo-clique numa linha de SERVICO,
   carrega os campos de edicao (descricao/valor/qtde/subtotal/desc/total/
   mecanico), habilita cmdEditarServicosAuto e desabilita Adicionar/Remover.
3) MostrarGrid_Servicos / FormatarGrid_Servicos: adiciona coluna oculta
   cod_mecanico (col 11) ao grid, para o DblClick conseguir ler o mecanico
   do servico clicado.
4) cmdEditarServicosAuto_Click: valida, faz UPDATE em OS_Servicos_Auto,
   recalcula totais (reaproveita o mesmo bloco de recalculo de
   cmdAdicionarServicosAuto_Click, extraido do arquivo), limpa os campos
   e devolve o estado dos botoes (Adicionar/Remover habilitados, Editar
   desabilitado).
5) cmdEditarServicosAuto comeca desabilitado (ENAB = 0) no design do form.
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


def find_sub(name, start=0):
    s = find_line_exact(f"Private Sub {name}()", start)
    e = find_line_exact("End Sub", s)
    return s, e


# ---------------------------------------------------------------
# 0) cmdEditarServicosAuto comeca desabilitado
# ---------------------------------------------------------------
i = find_line("Begin ChamaleonBtn.chameleonButton cmdEditarServicosAuto")
j = find_line("ENAB", i)
assert lines[j].strip() == "ENAB            =   -1  'True", lines[j]
lines[j] = lines[j].replace("-1  'True", "0   'False")

# ---------------------------------------------------------------
# 1) Declara vCodItemServicoEditando
# ---------------------------------------------------------------
i = find_line_exact("Dim vCodMecanicoServ As String")
lines[i] = lines[i] + '\r\nDim vCodItemServicoEditando As String'

# ---------------------------------------------------------------
# 2) Extrai o bloco de recalculo de totais de cmdAdicionarServicosAuto_Click
#    (o trecho entre "MostrarGrid_Servicos" e "LimparObjetos_ServicosAuto")
# ---------------------------------------------------------------
start_add, end_add = find_sub("cmdAdicionarServicosAuto_Click")
mgs_call = find_line_exact("MostrarGrid_Servicos", start_add)
recalculo_start = mgs_call + 1
limpar_call = find_line_exact("LimparObjetos_ServicosAuto", start_add)
recalc_block = lines[recalculo_start:limpar_call]
# tira linhas em branco nas pontas
while recalc_block and recalc_block[0].strip() == "":
    recalc_block.pop(0)
while recalc_block and recalc_block[-1].strip() == "":
    recalc_block.pop()

# ---------------------------------------------------------------
# 3) MostrarGrid_Servicos: adicionar cod_mecanico na SELECT (branch Auto/Moto/Info/Celular)
# ---------------------------------------------------------------
i = find_line(
    "SELECT 'SERVIÇO' as var_Tipo, COD_OS as var_COD, DESCRICAO, PRECO, QUANTIDADE, SUBTOTAL as var_SUBTOTAL, DESCONTO as var_DESCONTO, TOTAL as var_TOTAL, CODIGO AS var_CODITEM, '' as var_CODPROD FROM OS_Servicos_Auto"
)
lines[i] = lines[i].replace(
    "CODIGO AS var_CODITEM, '' as var_CODPROD FROM OS_Servicos_Auto",
    "CODIGO AS var_CODITEM, '' as var_CODPROD, cod_mecanico FROM OS_Servicos_Auto",
)

i = find_line(
    "pedidos_itens.CODIGO AS var_CODITEM, pedidos_itens.COD_PRODUTO as var_CODPROD \" & _",
    i,
)
lines[i] = lines[i].replace(
    'pedidos_itens.CODIGO AS var_CODITEM, pedidos_itens.COD_PRODUTO as var_CODPROD " & _',
    'pedidos_itens.CODIGO AS var_CODITEM, pedidos_itens.COD_PRODUTO as var_CODPROD, NULL as cod_mecanico " & _',
)

# ---------------------------------------------------------------
# 4) FormatarGrid_Servicos: coluna oculta 11 = cod_mecanico (branch Auto/Moto/Info/Celular)
# ---------------------------------------------------------------
start_fmt = find_line_exact("Private Sub FormatarGrid_Servicos(rTabela As ADODB.Recordset)")

i = find_line(".Cols = 11", start_fmt)
lines[i] = lines[i].replace(".Cols = 11", ".Cols = 12")

i = find_line(".ColWidth(10) = 0", start_fmt)
lines[i] = lines[i] + "\r\n       .ColWidth(11) = 0"

i = find_line('.TextMatrix(0, 10) = "PROD"', start_fmt)
lines[i] = lines[i] + '\r\n       .TextMatrix(0, 11) = "MEC"'

i = find_line('.TextMatrix(.Rows - 1, 10) = rTabela("var_CODPROD")', start_fmt)
lines[i] = lines[i] + '\r\n             .TextMatrix(.Rows - 1, 11) = ValidateNull(rTabela("cod_mecanico"))'

# ---------------------------------------------------------------
# 5) Grid_Servicos_DblClick - novo evento (inserido antes de MostrarGrid_Servicos)
# ---------------------------------------------------------------
start_mgs = find_line_exact("Private Sub MostrarGrid_Servicos()")

dblclick = []
dblclick.append("Private Sub Grid_Servicos_DblClick()")
dblclick.append(
    'If vTipoOS <> "Automóveis" And vTipoOS <> "Motocicletas" And vTipoOS <> "Informática" And vTipoOS <> "Celular" Then Exit Sub'
)
dblclick.append("If Grid_Servicos.Row = 0 Then Exit Sub")
dblclick.append('If Grid_Servicos.TextMatrix(Grid_Servicos.Row, 1) = "" Then Exit Sub')
dblclick.append('If Grid_Servicos.TextMatrix(Grid_Servicos.Row, 2) <> "SERVIÇO" Then Exit Sub')
dblclick.append("")
dblclick.append("vCodItemServicoEditando = Grid_Servicos.TextMatrix(Grid_Servicos.Row, 9)")
dblclick.append("vServico = Grid_Servicos.TextMatrix(Grid_Servicos.Row, 3)")
dblclick.append("mskValorServicoAuto.Text = Grid_Servicos.TextMatrix(Grid_Servicos.Row, 4)")
dblclick.append("txtQuantServicoAuto.Text = Grid_Servicos.TextMatrix(Grid_Servicos.Row, 5)")
dblclick.append("txtSubTotalServicoAuto.Text = Grid_Servicos.TextMatrix(Grid_Servicos.Row, 6)")
dblclick.append("txtDescServicoAuto.Text = Grid_Servicos.TextMatrix(Grid_Servicos.Row, 7)")
dblclick.append("txtTotalServicoAuto.Text = Grid_Servicos.TextMatrix(Grid_Servicos.Row, 8)")
dblclick.append("")
dblclick.append("Dim vCodMecEdit As String")
dblclick.append("vCodMecEdit = Grid_Servicos.TextMatrix(Grid_Servicos.Row, 11)")
dblclick.append('If vCodMecEdit <> "" Then')
dblclick.append('    sSQL = "SELECT nome FROM funcionario WHERE (codigo = " & vCodMecEdit & ");"')
dblclick.append("    Set r = dbData.OpenRecordset(sSQL)")
dblclick.append("    If Not r.EOF Then")
dblclick.append('        cboMecanicoServ.Text = r("nome")')
dblclick.append("        vCodMecanicoServ = vCodMecEdit")
dblclick.append("    End If")
dblclick.append("    If r.State <> 0 Then r.Close")
dblclick.append("    Set r = Nothing")
dblclick.append("Else")
dblclick.append('    cboMecanicoServ.Text = ""')
dblclick.append('    vCodMecanicoServ = ""')
dblclick.append("End If")
dblclick.append("")
dblclick.append("cmdEditarServicosAuto.Enabled = True")
dblclick.append("cmdAdicionarServicosAuto.Enabled = False")
dblclick.append("cmdRemoverServicosAuto.Enabled = False")
dblclick.append("End Sub")
dblclick.append("")

lines[start_mgs:start_mgs] = dblclick

# ---------------------------------------------------------------
# 6) cmdEditarServicosAuto_Click - novo evento (inserido logo apos cmdAdicionarServicosAuto_Click)
# ---------------------------------------------------------------
start_add2, end_add2 = find_sub("cmdAdicionarServicosAuto_Click")

editar = []
editar.append("")
editar.append("Private Sub cmdEditarServicosAuto_Click()")
editar.append('If vCodItemServicoEditando = "" Or txtCodOS.Text = "" Then Exit Sub')
editar.append('If txtQuantServicoAuto.Text = "" Then txtQuantServicoAuto.Text = 1')
editar.append('If mskValorServicoAuto.Text = "" Or mskValorServicoAuto.Text = "0,00" Then Exit Sub')
editar.append(
    'If vCodMecanicoServ = "" Then MsgBox "Selecione o mecânico que executou o serviço!", vbExclamation, "Aviso do Sistema": Exit Sub'
)
editar.append("")
editar.append("'CHECAR SE A OS ESTÁ FECHADA")
editar.append("Verificar_OS_Fechada")
editar.append("If OS_FECHADA = True Then Exit Sub")
editar.append("")
editar.append('If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Informática" Or vTipoOS = "Celular" Then')
editar.append('    dbData.Execute "UPDATE OS_Servicos_Auto SET descricao = \'" & vServico & "\', preco = " & Replace(CCur(mskValorServicoAuto.Text), ",", ".") & ", quantidade = " & txtQuantServicoAuto.Text & ", subtotal = " & Replace(CCur(txtSubTotalServicoAuto.Text), ",", ".") & ", desconto = " & Replace(CCur(txtDescServicoAuto.Text), ",", ".") & ", total = " & Replace(CCur(txtTotalServicoAuto.Text), ",", ".") & ", cod_mecanico = " & vCodMecanicoServ & " WHERE (codigo = " & vCodItemServicoEditando & ") AND (cod_os = " & txtCodOS.Text & ");"')
editar.append("End If")
editar.append("")
editar.append("MostrarGrid_Servicos")
editar.append("")
editar.extend(recalc_block)
editar.append("")
editar.append("vCodItemServicoEditando = \"\"")
editar.append("LimparObjetos_ServicosAuto")
editar.append("If cboTipo.Visible = True Then cboTipo.SetFocus Else cboServicosAuto.SetFocus")
editar.append("Somar_Totais")
editar.append("")
editar.append("cmdEditarServicosAuto.Enabled = False")
editar.append("cmdAdicionarServicosAuto.Enabled = True")
editar.append("cmdRemoverServicosAuto.Enabled = True")
editar.append("End Sub")

lines[end_add2 + 1 : end_add2 + 1] = editar

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
print("linhas do bloco de recalculo reaproveitado:", len(recalc_block))
