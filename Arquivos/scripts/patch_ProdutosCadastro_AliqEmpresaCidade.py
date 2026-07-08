# -*- coding: utf-8 -*-
# Patch em Produtos_Cadastro.frm:
#   1. Adiciona vars de modulo gCBSpAliq, gIBSUFpAliq, gIBSMunpAliq
#   2. Carrega essas vars no Form_Load a partir de Empresa e Cidade
#   3. Em CarregarProduto: usa vars (nao mais o campo salvo no produto)
#   4. Em LimparCampos: usa vars como default para novo produto
#   5. Em AtualizarReforma e cboIBSCBSCST_Click: substitui hardcoded 8,80/17,70/3,50/7,10

FRM = 'C:/Projeto/Compartilhado/Forms/Produtos_Cadastro.frm'

data = open(FRM, 'rb').read()
data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n')

subs = []

# ─────────────────────────────────────────────────────────────
# 1. Adicionar vars de modulo apos "Dim var_UF_Empresa As String"
# ─────────────────────────────────────────────────────────────
old = b'Dim var_UF_Empresa As String'
new = (b'Dim var_UF_Empresa As String\r\n'
       b'Dim gCBSpAliq As Double\r\n'
       b'Dim gIBSUFpAliq As Double\r\n'
       b'Dim gIBSMunpAliq As Double')
subs.append(('vars modulo', old, new))

# ─────────────────────────────────────────────────────────────
# 2. Form_Load: logo apos "r.Close / Set r = Nothing" do bloco Empresa,
#    antes de "Call PopularModalidades", inserir leitura de Empresa+Cidade
# ─────────────────────────────────────────────────────────────
old = b"r.Close\r\nSet r = Nothing\r\n\r\n'AGORA CHAMA O PREENCHIMENTO DO COMBO\r\nCall PopularModalidades"
new = (b"r.Close\r\nSet r = Nothing\r\n\r\n"
       b"' Carrega aliquotas IBS/CBS da Empresa e da Cidade da empresa\r\n"
       b"Dim rEmp As ADODB.Recordset\r\n"
       b"RsOpen rEmp, \"SELECT E.CBSpAliq, C.IBSUFpAliq, C.IBSMunpAliq \" & _\r\n"
       b"             \"FROM Empresa E \" & _\r\n"
       b"             \"LEFT JOIN Cidade C ON CAST(C.CodigoMunicipio AS NVARCHAR(7)) = \" & _\r\n"
       b"             \"                     CAST(E.CodigoIBGE AS NVARCHAR(7))\"\r\n"
       b"If Not rEmp.EOF Then\r\n"
       b"    gCBSpAliq  = IIf(IsNull(rEmp(\"CBSpAliq\")),  0, rEmp(\"CBSpAliq\"))\r\n"
       b"    gIBSUFpAliq = IIf(IsNull(rEmp(\"IBSUFpAliq\")), 0, rEmp(\"IBSUFpAliq\"))\r\n"
       b"    gIBSMunpAliq = IIf(IsNull(rEmp(\"IBSMunpAliq\")), 0, rEmp(\"IBSMunpAliq\"))\r\n"
       b"Else\r\n"
       b"    gCBSpAliq = 0: gIBSUFpAliq = 0: gIBSMunpAliq = 0\r\n"
       b"End If\r\n"
       b"If rEmp.State <> 0 Then rEmp.Close\r\n"
       b"Set rEmp = Nothing\r\n\r\n"
       b"'AGORA CHAMA O PREENCHIMENTO DO COMBO\r\nCall PopularModalidades")
subs.append(('form_load empresa+cidade', old, new))

# ─────────────────────────────────────────────────────────────
# 3. CarregarProduto: substituir leitura de r("CBSpAliq") etc. pelas vars
# ─────────────────────────────────────────────────────────────
old = (b'    txtCBSpAliq.Text = FormatNumber(ValidateNull(r("CBSpAliq")), 2)\r\n'
       b'    txtIBSUFpAliq.Text = FormatNumber(ValidateNull(r("IBSUFpAliq")), 2)\r\n'
       b'    txtIBSMunpAliq.Text = FormatNumber(ValidateNull(r("IBSMunpAliq")), 2)')
new  = (b'    txtCBSpAliq.Text   = FormatNumber(gCBSpAliq, 2)\r\n'
        b'    txtIBSUFpAliq.Text  = FormatNumber(gIBSUFpAliq, 2)\r\n'
        b'    txtIBSMunpAliq.Text = FormatNumber(gIBSMunpAliq, 2)')
subs.append(('carregarproduto', old, new))

# ─────────────────────────────────────────────────────────────
# 4. LimparCampos: default para novo produto usa vars
# ─────────────────────────────────────────────────────────────
old = (b'txtCBSpAliq.Text = "0,00"\r\n'
       b'txtIBSUFpAliq.Text = "0,00"\r\n'
       b'txtIBSMunpAliq.Text = "0,00"\r\n'
       b'cboISCST.ListIndex = 0')
new  = (b'txtCBSpAliq.Text   = FormatNumber(gCBSpAliq, 2)\r\n'
        b'txtIBSUFpAliq.Text  = FormatNumber(gIBSUFpAliq, 2)\r\n'
        b'txtIBSMunpAliq.Text = FormatNumber(gIBSMunpAliq, 2)\r\n'
        b'cboISCST.ListIndex = 0')
subs.append(('limparcampos', old, new))

# ─────────────────────────────────────────────────────────────
# 5. AtualizarReforma e cboIBSCBSCST_Click:
#    substituir aliquotas hardcoded pelos valores das vars
#    Nao mexer nos casos de isencao (que ficam em 0,00)
# ─────────────────────────────────────────────────────────────

# 5a. "8,80" (CBS integral/monofasico) → FormatNumber(gCBSpAliq, 2)
old = b'txtCBSpAliq.Text = "8,80"'
new = b'txtCBSpAliq.Text = FormatNumber(gCBSpAliq, 2)'
subs.append(('cbsaliq 8,80', old, new, True))   # replace_all

# 5b. "17,70" (IBS UF integral/monofasico) → FormatNumber(gIBSUFpAliq, 2)
old = b'txtIBSUFpAliq.Text = "17,70"'
new = b'txtIBSUFpAliq.Text = FormatNumber(gIBSUFpAliq, 2)'
subs.append(('ibsuf 17,70', old, new, True))

# 5c. "3,50" (CBS reduzido 60%) → FormatNumber(gCBSpAliq, 2)  [reducao fica em txtCBSpRedAliq]
old = b'txtCBSpAliq.Text = "3,50"'
new = b'txtCBSpAliq.Text = FormatNumber(gCBSpAliq, 2)'
subs.append(('cbsaliq 3,50', old, new, True))

# 5d. "7,10" (IBS UF reduzido 60%) → FormatNumber(gIBSUFpAliq, 2)
old = b'txtIBSUFpAliq.Text = "7,10"'
new = b'txtIBSUFpAliq.Text = FormatNumber(gIBSUFpAliq, 2)'
subs.append(('ibsuf 7,10', old, new, True))

# ─────────────────────────────────────────────────────────────
# Aplicar substituicoes
# ─────────────────────────────────────────────────────────────
ok = 0
for item in subs:
    name    = item[0]
    o       = item[1]
    n       = item[2]
    all_occ = item[3] if len(item) > 3 else False

    # normaliza CRLF no padrao de busca
    o_norm = o.replace(b'\r\n', b'\n')
    data_norm = data.replace(b'\r\n', b'\n')

    count = data_norm.count(o_norm)
    if count == 0:
        print(f'NENHUMA ocorrencia: {name}')
        continue
    if not all_occ and count > 1:
        print(f'AMBIGUO ({count} ocorrencias): {name}')
        continue

    # aplica substituicao normalizada e depois re-normaliza CRLF no resultado
    n_norm = n.replace(b'\r\n', b'\n')
    data = data_norm.replace(o_norm, n_norm)
    print(f'OK ({count}x): {name}')
    ok += 1

# Normalizar CRLF final
data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(FRM, 'wb').write(data)
print(f'\nTotal: {ok}/{len(subs)} substituicoes aplicadas.')
