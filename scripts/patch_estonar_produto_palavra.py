# -*- coding: utf-8 -*-
"""Estonar.frm - criterio PRODUTO ganha busca por Palavra/Palavras Duplas via cboProduto,
replicando o padrao ja usado em Produtos_Cadastro (vDescricao com optCompleto/optPorIniciais/
optPorPalavra/optPalavrasDuplas) e a versao refinada aplicada em Etiquetas_Impressao
(cboDesc virou texto livre, sem popular lista inteira nem autocomplete).

- cboProduto_GotFocus: parava de popular TODOS os produtos (SELECT * FROM produtos) +
  moCombo.AttachTo -> agora so Clear (texto livre), igual Etiquetas_Impressao.cboDesc.
- cboProduto_LostFocus: resolvia ItemData pra txtCodProduto (exato); sem popular, ListIndex
  fica sempre -1 -> essa resolucao nunca mais dispara. Removida (o filtro passa a ser via
  descricao, nao codigo exato).
- Mostrar_Pedido, ramo PRODUTO: `pedidos_itens.cod_produto = txtCodProduto` (exato) vira
  EXISTS (pedidos_itens JOIN produtos pr) com pr.descricao LIKE, no padrao optPorPalavra
  (1 LIKE) / optPalavrasDuplas (AND de varios LIKE, um por palavra, Split por espaco) -
  mesma logica de Produtos_Cadastro ~linha 7295. Usa alias `pr` (nao `p2`, que ja e o alias
  da derived table de pagamento no mesmo Sub). COLLATE Latin1_General_CI_AI (acento/case
  insensitive, mesmo padrao de Etiquetas_Impressao).
- cboCriterios_LostFocus: optPorPalavra/optPalavrasDuplas ficam visiveis SO no criterio
  PRODUTO (escondidos nos outros 6 ramos); ao entrar em PRODUTO, optPorPalavra vira o
  default (Value = True), igual optDesc_Click faz em Etiquetas_Impressao.

.frm cp1252 -> edicao binaria + normaliza CRLF.
"""
import sys

p = r"C:\projeto\PDV\Forms\Estonar.frm"
d = open(p, "rb").read().replace(b"\r\n", b"\n").replace(b"\r", b"\n")

reps = []
# cada item: (nome, old, new, quantas ocorrencias de old esperamos - 1 normalmente, 5 pro bloco comum)

# ---------------------------------------------------- cboProduto_GotFocus: vira texto livre
o = (b'Private Sub cboProduto_GotFocus()\n'
     b'Dim sSQL As String\n'
     b'Dim r As ADODB.Recordset\n'
     b'\n'
     b'cboProduto.Clear\n'
     b'\n'
     b'sSQL = "SELECT * FROM produtos ORDER BY descricao;"\n'
     b'Set r = dbData.OpenRecordset(sSQL)\n'
     b'\n'
     b'Do While Not r.EOF\n'
     b'   cboProduto.AddItem ValidateNull(r("descricao"))\n'
     b'   cboProduto.ItemData(cboProduto.NewIndex) = r("codigo")\n'
     b'   r.MoveNext\n'
     b'Loop\n'
     b'\n'
     b'If r.State <> 0 Then r.Close\n'
     b'Set r = Nothing\n'
     b'\n'
     b'moCombo.AttachTo cboProduto\n'
     b'End Sub\n')
n = (b'Private Sub cboProduto_GotFocus()\n'
     b"'busca por descricao (Palavra/Palavras Duplas) - nao popula mais a lista toda nem autocomplete\n"
     b'cboProduto.Clear\n'
     b'End Sub\n')
reps.append(("cboProduto_GotFocus -> texto livre (nao popula mais)", o, n, 1))

# ---------------------------------------------------- cboProduto_LostFocus: nao resolve mais ItemData
o = (b'Private Sub cboProduto_LostFocus()\n'
     b'On Error GoTo TrataErro\n'
     b'\n'
     b'If cboProduto.Text = "" Then txtCodProduto.Text = "": Exit Sub\n'
     b'If cboProduto.ListIndex = -1 Then txtCodProduto.Text = "": Exit Sub\n'
     b'txtCodProduto = cboProduto.ItemData(cboProduto.ListIndex)\n'
     b'   \n'
     b'TrataErro:\n'
     b'   If Err.Number = 381 Then Exit Sub\n'
     b'End Sub\n')
n = (b'Private Sub cboProduto_LostFocus()\n'
     b"'texto livre agora - filtro e por descricao (optPorPalavra/optPalavrasDuplas), nao por codigo exato\n"
     b'End Sub\n')
reps.append(("cboProduto_LostFocus -> no-op (filtro nao usa mais txtCodProduto)", o, n, 1))

# ---------------------------------------------------- Mostrar_Pedido: ramo PRODUTO
o = (b'ElseIf cboCriterios.Text = "PRODUTO" Then\n'
     b'    If txtCodProduto.Text = "" Then Exit Sub\n'
     b'    varCriterio = " and EXISTS (SELECT 1 FROM pedidos_itens WHERE pedidos_itens.cod_pedido = pedidos.COD_PEDIDO AND pedidos_itens.cod_produto = " & txtCodProduto.Text & ")"\n'
     b'End If\n')
n = (b'ElseIf cboCriterios.Text = "PRODUTO" Then\n'
     b'    If Trim(cboProduto.Text) = "" Then Exit Sub\n'
     b'    Dim vDescProduto As String\n'
     b'    vDescProduto = ""\n'
     b'    If optPorPalavra.Value = True Then\n'
     b"        vDescProduto = \"(pr.descricao COLLATE Latin1_General_CI_AI LIKE '%\" & cboProduto.Text & \"%')\"\n"
     b'    ElseIf optPalavrasDuplas.Value = True Then\n'
     b'        Dim aPalavras() As String\n'
     b'        Dim iPal As Integer\n'
     b'        Dim sPartes As String\n'
     b'        aPalavras = Split(Trim(cboProduto.Text), " ")\n'
     b'        sPartes = ""\n'
     b'        For iPal = 0 To UBound(aPalavras)\n'
     b'            If Trim(aPalavras(iPal)) <> "" Then\n'
     b'                If sPartes <> "" Then sPartes = sPartes & " AND "\n'
     b"                sPartes = sPartes & \"(pr.descricao COLLATE Latin1_General_CI_AI LIKE '%\" & Trim(aPalavras(iPal)) & \"%')\"\n"
     b'            End If\n'
     b'        Next iPal\n'
     b'        If sPartes <> "" Then vDescProduto = "(" & sPartes & ")"\n'
     b'    End If\n'
     b'    If vDescProduto = "" Then Exit Sub\n'
     b'    varCriterio = " and EXISTS (SELECT 1 FROM pedidos_itens INNER JOIN produtos pr ON pr.codigo = pedidos_itens.cod_produto WHERE pedidos_itens.cod_pedido = pedidos.COD_PEDIDO AND " & vDescProduto & ")"\n'
     b'End If\n')
reps.append(("Mostrar_Pedido PRODUTO -> busca por descricao (EXISTS + JOIN produtos)", o, n, 1))

# ---------------------------------------------------- cboCriterios_LostFocus: esconde nos 5 ramos comuns
o = (b'    lblProduto.Visible = False\n'
     b'    cboProduto.Visible = False\n'
     b'    lblCodBarra.Visible = False\n'
     b'    txtCodBarra.Visible = False\n')
n = (b'    lblProduto.Visible = False\n'
     b'    cboProduto.Visible = False\n'
     b'    lblCodBarra.Visible = False\n'
     b'    txtCodBarra.Visible = False\n'
     b'    optPorPalavra.Visible = False\n'
     b'    optPalavrasDuplas.Visible = False\n')
reps.append(("cboCriterios_LostFocus (5 ramos comuns) esconde optPorPalavra/optPalavrasDuplas", o, n, 5))

# ---------------------------------------------------- cboCriterios_LostFocus: mostra + default no ramo PRODUTO
o = (b'    lblProduto.Visible = True\n'
     b'    cboProduto.Visible = True\n'
     b'    lblCodBarra.Visible = False\n'
     b'    txtCodBarra.Visible = False\n'
     b'    cboProduto.SetFocus\n')
n = (b'    lblProduto.Visible = True\n'
     b'    cboProduto.Visible = True\n'
     b'    optPorPalavra.Visible = True\n'
     b'    optPalavrasDuplas.Visible = True\n'
     b'    optPorPalavra.Value = True\n'
     b'    lblCodBarra.Visible = False\n'
     b'    txtCodBarra.Visible = False\n'
     b'    cboProduto.SetFocus\n')
reps.append(("cboCriterios_LostFocus / PRODUTO mostra optPorPalavra/optPalavrasDuplas (default Palavra)", o, n, 1))

# ---------------------------------------------------- cboCriterios_LostFocus: esconde no ramo C\xd3D. BARRA
o = (b'    lblProduto.Visible = False\n'
     b'    cboProduto.Visible = False\n'
     b"    'cboStatus.ListIndex = 0\n")
n = (b'    lblProduto.Visible = False\n'
     b'    cboProduto.Visible = False\n'
     b'    optPorPalavra.Visible = False\n'
     b'    optPalavrasDuplas.Visible = False\n'
     b"    'cboStatus.ListIndex = 0\n")
reps.append(("cboCriterios_LostFocus / C\xd3D. BARRA esconde optPorPalavra/optPalavrasDuplas", o, n, 1))

for nome, old, new, esperado in reps:
    if new in d and old not in d:
        print("[ja] " + nome); continue
    c = d.count(old)
    if c != esperado:
        print("[ABORTA] %s -- %d ocorrencias de old (esperava %d)" % (nome, c, esperado)); sys.exit(1)
    if esperado > 1:
        d = d.replace(old, new)   # troca todas as ocorrencias de uma vez
    else:
        d = d.replace(old, new, 1)
    print("[ok]  " + nome + (" (%dx)" % c if esperado > 1 else ""))

open(p, "wb").write(d.replace(b"\n", b"\r\n"))
print("gravado " + p)
