import sys

# ============================================================
# 1. Produtos_Cadastro.frm — adicionar Public Sub EditarProduto
# ============================================================
data = open('Compartilhado/Forms/Produtos_Cadastro.frm', 'rb').read()

old_criar = (
    b"Public Sub CriarNovoProduto()\r\n"
    b"HabilitarFrames\r\n"
    b"cmdNovo.Enabled = False\r\n"
    b"cmdSalvar.Enabled = True\r\n"
    b"cmdCancelar.Enabled = True\r\n"
    b"vTipoEdicao = \"Novo\" 'desativei para teste\r\n"
    b"cmdExcluir.Enabled = False\r\n"
    b"\r\n"
    b"LimparObjetos_Produtos\r\n"
    b"\r\n"
    b"If frmComp.Visible = True Then LimparGrid_Comp\r\n"
    b"\r\n"
    b"AutoNumeracao\r\n"
    b"\r\n"
    b"cboUnidMedida.Text = \"UN\"\r\n"
    b"txtQuant.Text = \"0\"\r\n"
    b"'txtCodBarra.SetFocus\r\n"
    b"End Sub\r\n"
)

new_criar = (
    b"Public Sub CriarNovoProduto()\r\n"
    b"HabilitarFrames\r\n"
    b"cmdNovo.Enabled = False\r\n"
    b"cmdSalvar.Enabled = True\r\n"
    b"cmdCancelar.Enabled = True\r\n"
    b"vTipoEdicao = \"Novo\" 'desativei para teste\r\n"
    b"cmdExcluir.Enabled = False\r\n"
    b"\r\n"
    b"LimparObjetos_Produtos\r\n"
    b"\r\n"
    b"If frmComp.Visible = True Then LimparGrid_Comp\r\n"
    b"\r\n"
    b"AutoNumeracao\r\n"
    b"\r\n"
    b"cboUnidMedida.Text = \"UN\"\r\n"
    b"txtQuant.Text = \"0\"\r\n"
    b"'txtCodBarra.SetFocus\r\n"
    b"End Sub\r\n"
    b"\r\n"
    b"' Abre o form ja em modo edicao para o produto informado.\r\n"
    b"' O evento txtCodigo_Change dispara MostrarDados_Produto automaticamente.\r\n"
    b"Public Sub EditarProduto(lCodigo As Long)\r\n"
    b"   HabilitarFrames\r\n"
    b"   cmdNovo.Enabled = False\r\n"
    b"   cmdSalvar.Enabled = True\r\n"
    b"   cmdCancelar.Enabled = True\r\n"
    b"   cmdExcluir.Enabled = False\r\n"
    b"   vTipoEdicao = \"Edicao\"\r\n"
    b"   SSTab1.Tab = 0\r\n"
    b"   SSTab2.Tab = 0\r\n"
    b"   txtCodigo.Text = CStr(lCodigo)\r\n"
    b"End Sub\r\n"
)

c = data.count(old_criar)
print(f'1. EditarProduto em Produtos_Cadastro: count={c}')
if c == 1:
    data = data.replace(old_criar, new_criar)

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('Compartilhado/Forms/Produtos_Cadastro.frm', 'wb').write(data)
print('   Produtos_Cadastro salvo.')

# ============================================================
# 2. frmVinculoProdutoXML.frm — todas as mudancas
# ============================================================
data = open('OnlineCommerce/Forms/frmVinculoProdutoXML.frm', 'rb').read()

# 2a. Designer: reduzir cmdDesvincular, adicionar cmdAlterar, ajustar cmdEncerrar
old_botoes = (
    b"   Begin VB.CommandButton cmdDesvincular \r\n"
    b"      Caption         =   \"&Desvincular\"\r\n"
    b"      Height          =   555\r\n"
    b"      Left            =   5280\r\n"
    b"      TabIndex        =   19\r\n"
    b"      Top             =   8580\r\n"
    b"      Width           =   1905\r\n"
    b"   End\r\n"
    b"   Begin VB.CommandButton cmdEncerrar \r\n"
    b"      Caption         =   \"&Encerrar\"\r\n"
    b"      Height          =   555\r\n"
    b"      Left            =   7380\r\n"
    b"      TabIndex        =   14\r\n"
    b"      Top             =   8580\r\n"
    b"      Width           =   2640\r\n"
    b"   End\r\n"
)

new_botoes = (
    b"   Begin VB.CommandButton cmdDesvincular \r\n"
    b"      Caption         =   \"&Desvincular\"\r\n"
    b"      Height          =   555\r\n"
    b"      Left            =   5280\r\n"
    b"      TabIndex        =   19\r\n"
    b"      Top             =   8580\r\n"
    b"      Width           =   1440\r\n"
    b"   End\r\n"
    b"   Begin VB.CommandButton cmdAlterar \r\n"
    b"      Caption         =   \"&Alterar\"\r\n"
    b"      Enabled         =   0   'False\r\n"
    b"      Height          =   555\r\n"
    b"      Left            =   6840\r\n"
    b"      TabIndex        =   34\r\n"
    b"      Top             =   8580\r\n"
    b"      Width           =   1440\r\n"
    b"   End\r\n"
    b"   Begin VB.CommandButton cmdEncerrar \r\n"
    b"      Caption         =   \"&Encerrar\"\r\n"
    b"      Height          =   555\r\n"
    b"      Left            =   8400\r\n"
    b"      TabIndex        =   14\r\n"
    b"      Top             =   8580\r\n"
    b"      Width           =   1695\r\n"
    b"   End\r\n"
)

c = data.count(old_botoes)
print(f'2a. Designer cmdAlterar: count={c}')
if c == 1:
    data = data.replace(old_botoes, new_botoes)

# 2b. Form_Load: adicionar cmdAlterar.Enabled = False
old_load = (
    b"   cmdDesvincular.Enabled = False\r\n"
    b"   cmdCadastrar.Enabled = False\r\n"
)

new_load = (
    b"   cmdDesvincular.Enabled = False\r\n"
    b"   cmdCadastrar.Enabled = False\r\n"
    b"   cmdAlterar.Enabled = False\r\n"
)

c = data.count(old_load)
print(f'2b. Form_Load cmdAlterar: count={c}')
if c == 1:
    data = data.replace(old_load, new_load)

# 2c. AtualizarBotoes: adicionar cmdAlterar.Enabled = bTemProd
old_botoes_sub = (
    b"   cmdVincular.Enabled = bTemItem And (Not bVinculado) And bTemProd\r\n"
    b"   cmdDesvincular.Enabled = bTemItem And bVinculado And bTemProd\r\n"
    b"   cmdCadastrar.Enabled = bTemItem And (Not bVinculado)\r\n"
    b"End Sub\r\n"
)

new_botoes_sub = (
    b"   cmdVincular.Enabled = bTemItem And (Not bVinculado) And bTemProd\r\n"
    b"   cmdDesvincular.Enabled = bTemItem And bVinculado And bTemProd\r\n"
    b"   cmdCadastrar.Enabled = bTemItem And (Not bVinculado)\r\n"
    b"   cmdAlterar.Enabled = bTemProd\r\n"
    b"End Sub\r\n"
)

c = data.count(old_botoes_sub)
print(f'2c. AtualizarBotoes cmdAlterar: count={c}')
if c == 1:
    data = data.replace(old_botoes_sub, new_botoes_sub)

# 2d. Adicionar cmdAlterar_Click apos cmdDesvincular_Click End Sub
# Encontrar o final de cmdDesvincular_Click para inserir o novo sub depois
old_desvinc_end = (
    b"   MsgBox \"Vinculo desfeito com sucesso.\", vbInformation\r\n"
    b"   AtualizarBotoes\r\n"
    b"End Sub\r\n"
    b"\r\n"
    b"'==============================================================\r\n"
    b"Private Sub cmdEncerrar_Click()\r\n"
)

new_desvinc_end = (
    b"   MsgBox \"Vinculo desfeito com sucesso.\", vbInformation\r\n"
    b"   AtualizarBotoes\r\n"
    b"End Sub\r\n"
    b"\r\n"
    b"Private Sub cmdAlterar_Click()\r\n"
    b"   Dim idxP As Integer\r\n"
    b"   idxP = lstProdutos.Row + 1\r\n"
    b"   If UBound(arrIDProduto) < idxP Then Exit Sub\r\n"
    b"   Dim lCodProd As Long\r\n"
    b"   lCodProd = arrIDProduto(idxP)\r\n"
    b"   If lCodProd <= 0 Then Exit Sub\r\n"
    b"\r\n"
    b"   Load Produtos_Cadastro\r\n"
    b"   Produtos_Cadastro.SSTab1.Tab = 0\r\n"
    b"   Produtos_Cadastro.EditarProduto lCodProd\r\n"
    b"   Produtos_Cadastro.Show vbModal, Me\r\n"
    b"   Unload Produtos_Cadastro\r\n"
    b"\r\n"
    b"   ' Recarrega lstProdutos para exibir dados atualizados\r\n"
    b"   cmdBuscar_Click\r\n"
    b"End Sub\r\n"
    b"\r\n"
    b"'==============================================================\r\n"
    b"Private Sub cmdEncerrar_Click()\r\n"
)

c = data.count(old_desvinc_end)
print(f'2d. cmdAlterar_Click: count={c}')
if c == 1:
    data = data.replace(old_desvinc_end, new_desvinc_end)

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/frmVinculoProdutoXML.frm', 'wb').write(data)
print('   frmVinculoProdutoXML salvo. Tamanho:', len(data))
