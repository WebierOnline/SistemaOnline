# -*- coding: utf-8 -*-
# Patch: Reescreve AtualizarReforma() conforme novo mapeamento de categorias
# Novo mapeamento IBS/CBS:
#   CST 400 (isencao 0%):     ALIMENTOS (CESTA BASICA), HORTIFR�TI, CARNES  -> cClassTrib 400001
#   CST 200 (reduzida 60%):   LIMPEZA, HIGIENE E PERFUMARIA, BEBIDAS          -> cClassTrib 200001
#   CST 000 + IS:             BEBIDAS (ALCOOLICAS), BEBIDAS (ACUCARADAS)       -> cClassTrib 000001
#   CST 000 (padrao 1%):      demais (GERAL, FRIOS, PADARIA, PET SHOP, BAZAR) -> cClassTrib 000001
# Tambem adiciona SelecionarNoCombo cboIBSCBSClasse para auto-selecionar cClassTrib

FRM = 'C:/Projeto/Compartilhado/Forms/Produtos_Cadastro.frm'

data = open(FRM, 'rb').read()
data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n')

# ── Substitui o corpo completo de AtualizarReforma ──────────────────────────
old = (
    b"Private Sub AtualizarReforma()\r\n"
    b"    Dim sCat As String\r\n"
    b"    Dim sPISCST As String\r\n"
    b"    Dim sCOFINSCST As String\r\n"
    b"    Dim dPISAliq As String\r\n"
    b"    Dim dCOFINSAliq As String\r\n"
    b"    Dim sIBSCST As String\r\n"
    b"    sCat = UCase(cboCategoria.Text)\r\n"
    b"\r\n"
    b"    ' Reset inicial\r\n"
    b"    txtCBSpAliq.Text = \"0,00\"\r\n"
    b"    txtIBSUFpAliq.Text = \"0,00\"\r\n"
    b"    txtIBSMunpAliq.Text = \"0,00\"\r\n"
    b"    txtISpIS.Text = \"0,00\"\r\n"
    b"    SelecionarNoCombo cboISCST, \"00\", True\r\n"
    b"\r\n"
    b"    Select Case sCat\r\n"
    b"\r\n"
    b"        ' --- IBS/CBS 400 + PIS/COFINS 06: isen\xe7\xe3o (essenciais) ---\r\n"
    b"        Case \"ALIMENTOS (CESTA B\xc1SICA)\", \"HORTIFR\xdaTI\", \"CARNES\", \"LATIC\xcdNIOS\", \"CONGELADOS\"\r\n"
    b"            sIBSCST = \"400\"\r\n"
    b"            txtCBSpAliq.Text = \"0,00\": txtIBSUFpAliq.Text = \"0,00\"\r\n"
    b"            sPISCST = \"06\": sCOFINSCST = \"06\"\r\n"
    b"            dPISAliq = \"0,00\": dCOFINSAliq = \"0,00\"\r\n"
    b"\r\n"
    b"        ' --- IBS/CBS 620 + PIS/COFINS 04: monof\xe1sico ---\r\n"
    b"        Case \"HIGIENE E PERFUMARIA\", \"BEBIDAS\", \"BEBIDAS (A\xc7UCARADAS)\", \"BEBIDAS (ALCO\xd3LICAS)\", \"TABACARIA\"\r\n"
    b"            sIBSCST = \"600\"\r\n"
    b"            txtCBSpAliq.Text = FormatNumber(gCBSpAliq, 2): txtIBSUFpAliq.Text = FormatNumber(gIBSUFpAliq, 2)\r\n"
    b"            sPISCST = \"04\": sCOFINSCST = \"04\"\r\n"
    b"            dPISAliq = \"0,00\": dCOFINSAliq = \"0,00\"\r\n"
    b"            If sCat = \"BEBIDAS (A\xc7UCARADAS)\" Or sCat = \"BEBIDAS (ALCO\xd3LICAS)\" Or sCat = \"TABACARIA\" Then\r\n"
    b"                txtISpIS.Text = \"10,00\"\r\n"
    b"                SelecionarNoCombo cboISCST, \"01\", True\r\n"
    b"            End If\r\n"
    b"\r\n"
    b"        ' --- IBS/CBS 200 + PIS/COFINS 01: al\xedquota reduzida (alimentos processados) ---\r\n"
    b"        Case \"ALIMENTOS\", \"FRIOS E LATIC\xcdNIOS\", \"PADARIA E CONFEITARIA\"\r\n"
    b"            sIBSCST = \"200\"\r\n"
    b"            txtCBSpAliq.Text = FormatNumber(gCBSpAliq, 2): txtIBSUFpAliq.Text = FormatNumber(gIBSUFpAliq, 2)\r\n"
    b"            sPISCST = \"01\": sCOFINSCST = \"01\"\r\n"
    b"            If var_CRT = 3 Then\r\n"
    b"                dPISAliq = \"0,65\": dCOFINSAliq = \"3,00\"\r\n"
    b"            Else\r\n"
    b"                dPISAliq = \"0,00\": dCOFINSAliq = \"0,00\"\r\n"
    b"            End If\r\n"
    b"\r\n"
    b"        ' --- IBS/CBS 000 + PIS/COFINS 01: tributa\xe7\xe3o integral (limpeza, bazar, pet shop, etc.) ---\r\n"
    b"        Case Else\r\n"
    b"            sIBSCST = \"000\"\r\n"
    b"            txtCBSpAliq.Text = FormatNumber(gCBSpAliq, 2): txtIBSUFpAliq.Text = FormatNumber(gIBSUFpAliq, 2)\r\n"
    b"            sPISCST = \"01\": sCOFINSCST = \"01\"\r\n"
    b"            If var_CRT = 3 Then\r\n"
    b"                dPISAliq = \"0,65\": dCOFINSAliq = \"3,00\"\r\n"
    b"            Else\r\n"
    b"                dPISAliq = \"0,00\": dCOFINSAliq = \"0,00\"\r\n"
    b"            End If\r\n"
    b"\r\n"
    b"    End Select\r\n"
    b"\r\n"
    b"    SelecionarNoCombo cboIBSCBSCST, sIBSCST, True\r\n"
    b"    txtPISCST.Text = sPISCST\r\n"
    b"    txtCOFINSCST.Text = sCOFINSCST\r\n"
    b"    txtPisAliquota.Text = dPISAliq\r\n"
    b"    txtCofinsAliquota.Text = dCOFINSAliq\r\n"
    b"End Sub"
)

new = (
    b"Private Sub AtualizarReforma()\r\n"
    b"    Dim sCat As String\r\n"
    b"    Dim sPISCST As String\r\n"
    b"    Dim sCOFINSCST As String\r\n"
    b"    Dim dPISAliq As String\r\n"
    b"    Dim dCOFINSAliq As String\r\n"
    b"    Dim sIBSCST As String\r\n"
    b"    Dim sCClassTrib As String\r\n"
    b"    sCat = UCase(cboCategoria.Text)\r\n"
    b"\r\n"
    b"    ' Reset inicial\r\n"
    b"    txtCBSpAliq.Text = \"0,00\"\r\n"
    b"    txtIBSUFpAliq.Text = \"0,00\"\r\n"
    b"    txtIBSMunpAliq.Text = \"0,00\"\r\n"
    b"    txtISpIS.Text = \"0,00\"\r\n"
    b"    SelecionarNoCombo cboISCST, \"00\", True\r\n"
    b"\r\n"
    b"    Select Case sCat\r\n"
    b"\r\n"
    b"        ' --- IBS/CBS 400 (Isen\xe7\xe3o 0%): Cesta B\xe1sica, Hortifr\xfati, Carnes ---\r\n"
    b"        Case \"ALIMENTOS (CESTA B\xc1SICA)\", \"HORTIFR\xdaTI\", \"CARNES\"\r\n"
    b"            sIBSCST = \"400\": sCClassTrib = \"400001\"\r\n"
    b"            sPISCST = \"06\": sCOFINSCST = \"06\"\r\n"
    b"            dPISAliq = \"0,00\": dCOFINSAliq = \"0,00\"\r\n"
    b"\r\n"
    b"        ' --- IBS/CBS 200 (Reduzida 60%): Limpeza, Higiene, Bebidas n\xe3o-alcool\xedcas ---\r\n"
    b"        Case \"LIMPEZA\", \"HIGIENE E PERFUMARIA\", \"BEBIDAS\"\r\n"
    b"            sIBSCST = \"200\": sCClassTrib = \"200001\"\r\n"
    b"            sPISCST = \"04\": sCOFINSCST = \"04\"\r\n"
    b"            dPISAliq = \"0,00\": dCOFINSAliq = \"0,00\"\r\n"
    b"\r\n"
    b"        ' --- IBS/CBS 000 + IS: Bebidas alco\xf3licas e a\xe7ucaradas ---\r\n"
    b"        Case \"BEBIDAS (ALCO\xd3LICAS)\", \"BEBIDAS (A\xc7UCARADAS)\"\r\n"
    b"            sIBSCST = \"000\": sCClassTrib = \"000001\"\r\n"
    b"            sPISCST = \"04\": sCOFINSCST = \"04\"\r\n"
    b"            dPISAliq = \"0,00\": dCOFINSAliq = \"0,00\"\r\n"
    b"            txtISpIS.Text = \"10,00\"\r\n"
    b"            SelecionarNoCombo cboISCST, \"01\", True\r\n"
    b"\r\n"
    b"        ' --- IBS/CBS 000 (Padr\xe3o 1%): demais categorias ---\r\n"
    b"        Case \"ALIMENTOS\", \"ALIMENTOS (GERAL)\", \"FRIOS E LATIC\xcdNIOS\", _\r\n"
    b"             \"PADARIA E CONFEITARIA\", \"PET SHOP\", \"BAZAR E UTILIDADES\", _\r\n"
    b"             \"LATIC\xcdNIOS\", \"CONGELADOS\", \"TABACARIA\"\r\n"
    b"            sIBSCST = \"000\": sCClassTrib = \"000001\"\r\n"
    b"            sPISCST = \"01\": sCOFINSCST = \"01\"\r\n"
    b"            If var_CRT = 3 Then\r\n"
    b"                dPISAliq = \"0,65\": dCOFINSAliq = \"3,00\"\r\n"
    b"            Else\r\n"
    b"                dPISAliq = \"0,00\": dCOFINSAliq = \"0,00\"\r\n"
    b"            End If\r\n"
    b"\r\n"
    b"        ' --- Case Else: CST 000 gen\xe9rico ---\r\n"
    b"        Case Else\r\n"
    b"            sIBSCST = \"000\": sCClassTrib = \"000001\"\r\n"
    b"            sPISCST = \"01\": sCOFINSCST = \"01\"\r\n"
    b"            If var_CRT = 3 Then\r\n"
    b"                dPISAliq = \"0,65\": dCOFINSAliq = \"3,00\"\r\n"
    b"            Else\r\n"
    b"                dPISAliq = \"0,00\": dCOFINSAliq = \"0,00\"\r\n"
    b"            End If\r\n"
    b"\r\n"
    b"    End Select\r\n"
    b"\r\n"
    b"    ' Seleciona CST (dispara cboIBSCBSCST_Click que preenche cboIBSCBSClasse)\r\n"
    b"    SelecionarNoCombo cboIBSCBSCST, sIBSCST, True\r\n"
    b"    ' Seleciona a classificacao especifica (dispara cboIBSCBSClasse_Click com reducao)\r\n"
    b"    SelecionarNoCombo cboIBSCBSClasse, sCClassTrib\r\n"
    b"    txtPISCST.Text = sPISCST\r\n"
    b"    txtCOFINSCST.Text = sCOFINSCST\r\n"
    b"    txtPisAliquota.Text = dPISAliq\r\n"
    b"    txtCofinsAliquota.Text = dCOFINSAliq\r\n"
    b"End Sub"
)

# Normalizar para comparacao
o_n = old.replace(b'\r\n', b'\n')
d_n = data.replace(b'\r\n', b'\n')
count = d_n.count(o_n)

if count == 0:
    print('NENHUMA ocorrencia encontrada - verificando trecho...')
    # Tenta encontrar parte do inicio para debug
    trecho = old[:200].replace(b'\r\n', b'\n')
    if d_n.count(trecho) > 0:
        print('  Inicio do sub encontrado, mas corpo diferente do esperado')
    else:
        print('  Inicio do sub NAO encontrado')
elif count > 1:
    print(f'AMBIGUO ({count}x)')
else:
    n_n = new.replace(b'\r\n', b'\n')
    data = d_n.replace(o_n, n_n)
    data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(FRM, 'wb').write(data)
    print('OK: AtualizarReforma atualizado com novo mapeamento de categorias')
