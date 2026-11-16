VERSION 5.00
Begin VB.Form frmNFSe_Editar
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NFS-e - Conferir antes de transmitir"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9480
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraServ
      Caption         =   "Serviço desta nota (tabelas oficiais da NFS-e)"
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   780
      Width           =   9240
      Begin VB.ComboBox cboCodTribNac
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   270
         Width           =   5655
      End
      Begin VB.TextBox txtCodTribMun
         Height          =   315
         Left            =   7920
         MaxLength       =   3
         TabIndex        =   7
         Top             =   270
         Width           =   555
      End
      Begin VB.ComboBox cboNBS
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   660
         Width           =   5655
      End
      Begin VB.TextBox txtCompetencia
         Height          =   315
         Left            =   7920
         MaxLength       =   10
         TabIndex        =   11
         Top             =   660
         Width           =   1155
      End
      Begin VB.OptionButton optEmpresa
         Caption         =   "Na empresa"
         Height          =   255
         Left            =   1080
         TabIndex        =   28
         Top             =   1080
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optCliente
         Caption         =   "No local do cliente"
         Height          =   255
         Left            =   2340
         TabIndex        =   29
         Top             =   1080
         Width           =   1755
      End
      Begin VB.ComboBox cboUFPrest
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1050
         Width           =   735
      End
      Begin VB.ComboBox cboMunPrest
         Height          =   315
         Left            =   5940
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1050
         Width           =   3135
      End
      Begin VB.ComboBox cboIndOp
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   1470
         Width           =   7995
      End
      Begin VB.ComboBox cboClassTrib
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   1860
         Width           =   7995
      End
      Begin VB.Label lblCodNac
         Caption         =   "Serviço:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   315
         Width           =   915
      End
      Begin VB.Label lblCodMun
         Caption         =   "Municipal:"
         Height          =   255
         Left            =   6840
         TabIndex        =   6
         Top             =   315
         Width           =   1035
      End
      Begin VB.Label lblNBS
         Caption         =   "NBS:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   705
         Width           =   915
      End
      Begin VB.Label lblComp
         Caption         =   "Competência:"
         Height          =   255
         Left            =   6840
         TabIndex        =   10
         Top             =   705
         Width           =   1035
      End
      Begin VB.Label lblLocal
         Caption         =   "Feito:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1095
         Width           =   915
      End
      Begin VB.Label lblMunPrest
         Caption         =   "Município:"
         Height          =   255
         Left            =   4200
         TabIndex        =   30
         Top             =   1095
         Width           =   915
      End
      Begin VB.Label lblIndOp
         Caption         =   "Indicador:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1515
         Width           =   915
      End
      Begin VB.Label lblClassTrib
         Caption         =   "cClassTrib:"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1905
         Width           =   915
      End
   End
   Begin VB.TextBox txtDescricao
      Height          =   1800
      Left            =   120
      MaxLength       =   2000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   3405
      Width           =   9240
   End
   Begin VB.TextBox txtPedido
      Height          =   315
      Left            =   6600
      MaxLength       =   60
      TabIndex        =   16
      Top             =   5250
      Width           =   2760
   End
   Begin VB.TextBox txtInfo
      Height          =   975
      Left            =   120
      MaxLength       =   2000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   5610
      Width           =   9240
   End
   Begin VB.Frame fraISS
      Caption         =   "ISS"
      Height          =   735
      Left            =   120
      TabIndex        =   19
      Top             =   6690
      Width           =   9240
      Begin VB.CheckBox chkRetido
         Caption         =   "ISS retido pelo tomador (só quando o cliente reter o imposto)"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   315
         Width           =   4935
      End
      Begin VB.TextBox txtAliquota
         Height          =   315
         Left            =   6720
         MaxLength       =   6
         TabIndex        =   22
         Top             =   270
         Width           =   795
      End
      Begin VB.Label lblAliq
         Caption         =   "Alíquota do ISS (%):"
         Height          =   255
         Left            =   5160
         TabIndex        =   21
         Top             =   315
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdSalvar
      Caption         =   "&Salvar"
      Height          =   375
      Left            =   3960
      TabIndex        =   24
      Top             =   8250
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalvarTransmitir
      Caption         =   "Salvar e &Transmitir"
      Height          =   375
      Left            =   5760
      TabIndex        =   25
      Top             =   8250
      Width           =   1815
   End
   Begin VB.CommandButton cmdVoltar
      Cancel          =   -1  'True
      Caption         =   "&Voltar"
      Height          =   375
      Left            =   7665
      TabIndex        =   26
      Top             =   8250
      Width           =   1695
   End
   Begin VB.Label lblValores
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   7530
      Width           =   9240
   End
   Begin VB.Label lblInfoT
      Caption         =   "Informações complementares (veículo, observações):"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   5370
      Width           =   4455
   End
   Begin VB.Label lblPedido
      Alignment       =   1  'Right Justify
      Caption         =   "Pedido/OC do cliente:"
      Height          =   255
      Left            =   4680
      TabIndex        =   15
      Top             =   5295
      Width           =   1845
   End
   Begin VB.Label lblContDesc
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6360
      TabIndex        =   14
      Top             =   3165
      Width           =   3000
   End
   Begin VB.Label lblDesc
      Caption         =   "Descrição dos serviços (vai na nota):"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3165
      Width           =   4335
   End
   Begin VB.Label lblTomador
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9240
   End
   Begin VB.Label lblTomadorEnd
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   390
      Width           =   9240
   End
   Begin VB.Label lblMensagem
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   8010
      Width           =   9240
   End
End
Attribute VB_Name = "frmNFSe_Editar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Conferência/edição do rascunho da NFS-e (Situacao 0 ou 2) antes de transmitir. O NFSeEmissor manda ao Sefin
'exatamente o que estiver gravado na tabela NFSe.
'Serviço, NBS, indicador da operação e cClassTrib vêm das tabelas oficiais (modNFSeTabelas), começando pelos do
'Configurar NFS-e. Local da prestação: "Na empresa" (município da empresa, indicador final 01, ex.: 050101) ou
'"No local do cliente" (serviço feito fora do estabelecimento: município escolhido, indicador final 02, ex.: 050102).
'O indicador depende de ONDE o serviço foi feito, não da cidade do cliente.
Private mIdNFSe As Long
Private mTransmitir As Boolean
Private mCarregando As Boolean
Private mTem182 As Boolean        'script 182: IndOp, ClassTrib e PrestacaoNoCliente na nota
Private mIbsCbs As Boolean        'grupo IBS/CBS ligado no Configurar
Private mMunEmpresa As Long
Private mMunCliente As Long
Private mSemTomador As String     'NFSe_SemTomador: "" / "PF" / "PJ" (aviso em vermelho; a confirmação é no TransmitirId)
Private mMei As Boolean          'empresa MEI (CRT 4): ISS fixo no DAS, nunca retido

'Mostra a nota e devolve True quando o usuário escolheu "Salvar e Transmitir"
Public Function Editar(ByVal lId As Long) As Boolean
mIdNFSe = lId
mTransmitir = False
Load Me
If Not Carregar() Then
   Unload Me
   Exit Function
End If
Me.Show vbModal
Editar = mTransmitir
Unload Me
End Function

Private Function TemCampo(r As ADODB.Recordset, ByVal sCampo As String) As Boolean
Dim f As ADODB.Field
For Each f In r.Fields
   If StrComp(f.Name, sCampo, vbTextCompare) = 0 Then TemCampo = True: Exit Function
Next
End Function

Private Function Carregar() As Boolean
Dim r As ADODB.Recordset
Dim sDoc As String, sEnd As String
Dim sIndOpEmp As String, sClassEmp As String, sIndOp As String, sClass As String
Dim bNoCliente As Boolean
Dim lMunPrest As Long

mCarregando = True
NFSeTab_Carregar
NFSeTab_LarguraLista cboCodTribNac, 900
NFSeTab_LarguraLista cboNBS, 760
NFSeTab_LarguraLista cboIndOp, 820
NFSeTab_LarguraLista cboClassTrib, 760

Set r = dbData.OpenRecordset("SELECT CodigoIBGE, CRT, NFSeIBSCBS, NFSeIndOp, NFSeClassTrib FROM empresa;")
If Not r.EOF Then
   mMunEmpresa = Val(ValidateNull(r("CodigoIBGE")))
   mMei = (Val(ValidateNull(r("CRT"))) = 4)
   mIbsCbs = CBool(ValidateNull(r("NFSeIBSCBS")))
   sIndOpEmp = SoDigitos(ValidateNull(r("NFSeIndOp")))
   sClassEmp = SoDigitos(ValidateNull(r("NFSeClassTrib")))
End If
r.Close

Set r = dbData.OpenRecordset("SELECT * FROM NFSe WHERE (IdNFSe = " & mIdNFSe & ");")
If r.EOF Then
   r.Close
   MsgBox "NFS-e não encontrada.", vbExclamation, "NFS-e"
   Exit Function
End If
If Val(ValidateNull(r("Situacao"))) <> NFSE_DIGITACAO And Val(ValidateNull(r("Situacao"))) <> NFSE_REJEITADA Then
   r.Close
   MsgBox "Só é possível alterar uma NFS-e em digitação ou rejeitada.", vbExclamation, "NFS-e"
   Exit Function
End If
mTem182 = TemCampo(r, "IndOp") And TemCampo(r, "ClassTrib") And TemCampo(r, "PrestacaoNoCliente")

sDoc = ValidateNull(r("TomadorCpfCnpj"))
If Len(sDoc) = 14 Then
   sDoc = Format$(sDoc, "@@.@@@.@@@/@@@@-@@")
ElseIf Len(sDoc) = 11 Then
   sDoc = Format$(sDoc, "@@@.@@@.@@@-@@")
Else
   sDoc = "SEM CPF/CNPJ (a nota sai sem tomador identificado)"
End If
mSemTomador = NFSe_SemTomador(mIdNFSe)
If mSemTomador = "PJ" Then sDoc = "EMPRESA SEM CNPJ NO CADASTRO (a nota sai sem tomador)"
lblTomador.ForeColor = IIf(mSemTomador = "", vbButtonText, &HC0&)
lblTomador.Caption = "Tomador: " & ValidateNull(r("TomadorNome")) & "  -  " & sDoc
sEnd = Trim$(ValidateNull(r("TomadorLogradouro")))
If sEnd <> "" Then sEnd = sEnd & ", " & ValidateNull(r("TomadorNumero")) & " - " & ValidateNull(r("TomadorBairro")) & " - IBGE " & ValidateNull(r("TomadorCodMunicipio")) & " - CEP " & ValidateNull(r("TomadorCEP"))
If sEnd = "" Then sEnd = "Endereço do cliente incompleto (sem logradouro ou código IBGE): a nota vai sem o endereço do tomador."
lblTomadorEnd.Caption = sEnd & "   (dados copiados do cadastro do cliente quando a nota foi gerada)"
mMunCliente = Val(ValidateNull(r("TomadorCodMunicipio")))

'serviço -> NBS -> indicador -> cClassTrib: os da nota; sem eles (nota antiga / sem script 182), os do Configurar
sIndOp = sIndOpEmp
sClass = sClassEmp
If mTem182 Then
   If SoDigitos(ValidateNull(r("IndOp"))) <> "" Then sIndOp = SoDigitos(ValidateNull(r("IndOp")))
   If SoDigitos(ValidateNull(r("ClassTrib"))) <> "" Then sClass = SoDigitos(ValidateNull(r("ClassTrib")))
End If
NFSeTab_PreencherServicos cboCodTribNac, SoDigitos(ValidateNull(r("CodTribNac")))
NFSeTab_PreencherNBS cboCodTribNac, cboNBS, SoDigitos(ValidateNull(r("CodNBS")))
NFSeTab_PreencherIndOp cboCodTribNac, cboNBS, cboIndOp, sIndOp
NFSeTab_PreencherClassTrib cboCodTribNac, cboNBS, cboClassTrib, sClass

lMunPrest = Val(ValidateNull(r("CodMunicipioPrestacao")))
If mTem182 Then
   bNoCliente = CBool(ValidateNull(r("PrestacaoNoCliente")))
Else
   bNoCliente = (lMunPrest <> 0 And lMunPrest <> mMunEmpresa)
End If

txtCodTribMun.Text = ValidateNull(r("CodTribMun"))
txtCompetencia.Text = Format$(r("DataCompetencia"), "dd/mm/yyyy")
txtDescricao.Text = ValidateNull(r("Descricao"))
txtInfo.Text = ValidateNull(r("InfoComplementar"))
txtPedido.Text = ValidateNull(r("NumeroPedido"))
chkRetido.Value = IIf(Val(ValidateNull(r("RetencaoISS"))) = 2, vbChecked, vbUnchecked)
If mMei Then
   chkRetido.Value = vbUnchecked
   chkRetido.Enabled = False
   chkRetido.Caption = "MEI: o ISS é recolhido no DAS e nunca é retido pelo tomador"
End If
txtAliquota.Text = IIf(IsNull(r("AliquotaISS")), "", Format$(ValidateNull(r("AliquotaISS")), "0.00"))
lblValores.Caption = "Serviços: R$ " & Format$(ValidateNull(r("ValorServicos")), ocMONEY) & _
                     "     Desconto: R$ " & Format$(ValidateNull(r("DescontoIncondicionado")), ocMONEY) & _
                     "     Líquido: R$ " & Format$(CDbl(ValidateNull(r("ValorServicos"))) - CDbl(ValidateNull(r("DescontoIncondicionado"))), ocMONEY) & vbCrLf & _
                     "(valores vindos das OS; para mudar, ajuste a OS, exclua este rascunho e gere de novo)"
lblMensagem.Caption = IIf(Val(ValidateNull(r("Situacao"))) = NFSE_REJEITADA, "Rejeição: " & ValidateNull(r("Mensagem")), "")
If lblMensagem.Caption = "" Then
   If mSemTomador = "PJ" Then
      lblMensagem.Caption = "Cliente EMPRESA sem CNPJ: corrija o cadastro, exclua este rascunho e gere a nota de novo."
   ElseIf mSemTomador = "PF" Then
      lblMensagem.Caption = "Sem CPF: só vale para pessoa física que não pediu o CPF na nota."
   End If
End If
r.Close
Set r = Nothing

CarregarUFs
If bNoCliente Then optCliente.Value = True Else optEmpresa.Value = True
SelecionarMunicipio IIf(bNoCliente, lMunPrest, mMunEmpresa)
AjustarLocal False

If txtAliquota.Text = "" Then txtAliquota.Text = Format$(SQLExecutaRetorno("SELECT NFSeAliquotaISS FROM empresa", "NFSeAliquotaISS", ""), "0.00")
chkRetido_Click
txtDescricao_Change

'indicador e classificação só vão na nota com o IBS/CBS ligado no Configurar
cboIndOp.Enabled = mIbsCbs And mTem182
cboClassTrib.Enabled = cboIndOp.Enabled
lblIndOp.Enabled = cboIndOp.Enabled
lblClassTrib.Enabled = cboIndOp.Enabled
If Not mTem182 Then
   fraServ.Caption = "Serviço desta nota  (para escolher local e indicador por nota, atualize o banco: script 182)"
ElseIf Not mIbsCbs Then
   fraServ.Caption = "Serviço desta nota  (indicador e cClassTrib só vão na nota com IBS/CBS ligado no Configurar NFS-e)"
End If
If Not NFSeTab_Ok() Then
   lblMensagem.Caption = "Tabelas oficiais da NFS-e não encontradas na pasta NFSe: atualize o sistema."
End If
mCarregando = False
Carregar = True
End Function

'---------------------------------------------------------------------------------------------------- serviço

Private Sub cboCodTribNac_Click()
If mCarregando Then Exit Sub
NFSeTab_PreencherNBS cboCodTribNac, cboNBS, NFSeTab_CodigoDoCombo(cboNBS)
cboNBS_Click
End Sub

Private Sub cboNBS_Click()
If mCarregando Then Exit Sub
NFSeTab_PreencherIndOp cboCodTribNac, cboNBS, cboIndOp, NFSeTab_CodigoDoCombo(cboIndOp)
NFSeTab_PreencherClassTrib cboCodTribNac, cboNBS, cboClassTrib, NFSeTab_CodigoDoCombo(cboClassTrib)
NFSeTab_IndOpPorLocal cboIndOp, optCliente.Value
End Sub

'---------------------------------------------------------------------------------------------------- local da prestação

Private Sub CarregarUFs()
Dim r As ADODB.Recordset
cboUFPrest.Clear
On Error GoTo semTabela
Set r = dbData.OpenRecordset("SELECT DISTINCT UF FROM Cidade WHERE (ISNULL(UF, '') <> '') ORDER BY UF;")
Do While Not r.EOF
   cboUFPrest.AddItem Trim$(r("UF"))
   r.MoveNext
Loop
r.Close
Set r = Nothing
Exit Sub
semTabela:
End Sub

Private Sub CarregarMunicipios(ByVal sUF As String)
Dim r As ADODB.Recordset
cboMunPrest.Clear
On Error GoTo semTabela
Set r = dbData.OpenRecordset("SELECT Nome, CodigoMunicipio FROM Cidade WHERE (UF = " & NFSe_Q(sUF) & ") ORDER BY Nome;")
Do While Not r.EOF
   If Val(ValidateNull(r("CodigoMunicipio"))) > 0 Then
      cboMunPrest.AddItem Trim$(r("Nome"))
      cboMunPrest.ItemData(cboMunPrest.NewIndex) = Val(ValidateNull(r("CodigoMunicipio")))
   End If
   r.MoveNext
Loop
r.Close
Set r = Nothing
Exit Sub
semTabela:
End Sub

'UF e município pelo código IBGE (tabela Cidade). Fora da tabela: mostra só o código.
Private Sub SelecionarMunicipio(ByVal lCodigo As Long)
Dim sUF As String
Dim i As Integer
If lCodigo = 0 Then Exit Sub
On Error Resume Next
sUF = Trim$(SQLExecutaRetorno("SELECT TOP 1 UF FROM Cidade WHERE (CodigoMunicipio = '" & lCodigo & "')", "UF", ""))
On Error GoTo 0
If sUF <> "" Then
   For i = 0 To cboUFPrest.ListCount - 1
      If cboUFPrest.List(i) = sUF Then cboUFPrest.ListIndex = i: Exit For
   Next
   If cboUFPrest.Text <> sUF Then CarregarMunicipios sUF
End If
For i = 0 To cboMunPrest.ListCount - 1
   If cboMunPrest.ItemData(i) = lCodigo Then cboMunPrest.ListIndex = i: Exit Sub
Next
cboMunPrest.AddItem "IBGE " & lCodigo
cboMunPrest.ItemData(cboMunPrest.NewIndex) = lCodigo
cboMunPrest.ListIndex = cboMunPrest.NewIndex
End Sub

Private Sub cboUFPrest_Click()
CarregarMunicipios cboUFPrest.Text
End Sub

Private Function MunicipioEscolhido() As Long
If cboMunPrest.ListIndex < 0 Then Exit Function
MunicipioEscolhido = cboMunPrest.ItemData(cboMunPrest.ListIndex)
End Function

'Na empresa: município da empresa e indicador final 01. No local do cliente: município (começa pelo do cliente) e final 02.
Private Sub AjustarLocal(ByVal bTrocou As Boolean)
cboUFPrest.Enabled = optCliente.Value
cboMunPrest.Enabled = optCliente.Value
lblMunPrest.Enabled = optCliente.Value
optEmpresa.Enabled = mTem182
optCliente.Enabled = mTem182
If Not bTrocou Then Exit Sub
If optEmpresa.Value Then
   SelecionarMunicipio mMunEmpresa
ElseIf mMunCliente <> 0 Then
   SelecionarMunicipio mMunCliente
End If
NFSeTab_IndOpPorLocal cboIndOp, optCliente.Value
End Sub

Private Sub optEmpresa_Click()
If mCarregando Then Exit Sub
AjustarLocal True
End Sub

Private Sub optCliente_Click()
If mCarregando Then Exit Sub
AjustarLocal True
End Sub

'---------------------------------------------------------------------------------------------------- tela

Private Sub chkRetido_Click()
txtAliquota.Enabled = (chkRetido.Value = vbChecked)
lblAliq.Enabled = txtAliquota.Enabled
End Sub

Private Sub txtDescricao_Change()
lblContDesc.Caption = Len(txtDescricao.Text) & " de 2000 caracteres"
End Sub

'Pedido/OC: a nota é gerada com "OS 10" (número da OS da oficina). Quando a empresa cliente manda a ordem de
'pedido/compra dela, o número vai no xPed da nota e também nas informações complementares, logo depois dos
'dados do veículo: "Ref. Ordem de Serviços Nº 10, Veículo: ..., Placa AAA1234 / Ordem de Pedido/Compra: 4587".
Private Sub txtPedido_Change()
If mCarregando Then Exit Sub
txtInfo.Text = InfoComPedido(txtInfo.Text, txtPedido.Text)
End Sub

Private Function PedidoDoCliente(ByVal sPed As String) As String
sPed = Trim$(sPed)
If UCase$(Left$(sPed, 3)) = "OS " Then sPed = ""    'o padrão "OS 10" não é pedido do cliente
PedidoDoCliente = sPed
End Function

Private Function InfoComPedido(ByVal sInfo As String, ByVal sPed As String) As String
Const ROTULO As String = "Ordem de Pedido/Compra:"
Dim p As Long, pIni As Long, pFim As Long

'tira o pedido que já estiver no texto (até o próximo " | " ou o fim)
p = InStr(1, sInfo, ROTULO, vbTextCompare)
If p > 0 Then
   pIni = p
   If p > 3 Then
      If Mid$(sInfo, p - 3, 3) = " / " Then pIni = p - 3
   End If
   pFim = InStr(p, sInfo, " | ")
   If pFim = 0 Then pFim = Len(sInfo) + 1
   sInfo = Left$(sInfo, pIni - 1) & Mid$(sInfo, pFim)
End If

sPed = PedidoDoCliente(sPed)
If sPed <> "" Then
   'depois dos dados do veículo da 1ª OS (antes do " | " da próxima OS ou da mensagem da empresa)
   p = InStr(1, sInfo, " | ")
   If Trim$(sInfo) = "" Then
      sInfo = ROTULO & " " & sPed
   ElseIf p = 0 Then
      sInfo = sInfo & " / " & ROTULO & " " & sPed
   Else
      sInfo = Left$(sInfo, p - 1) & " / " & ROTULO & " " & sPed & Mid$(sInfo, p)
   End If
End If
If Len(sInfo) > 2000 Then sInfo = Left$(sInfo, 2000)
InfoComPedido = sInfo
End Function

Private Function SoDigitos(ByVal s As String) As String
Dim i As Integer
For i = 1 To Len(s)
   If Mid$(s, i, 1) Like "[0-9]" Then SoDigitos = SoDigitos & Mid$(s, i, 1)
Next
End Function

Private Function Salvar() As Boolean
Dim sNac As String, sMun As String, sNBS As String, sAliq As String, sIndOp As String, sClass As String
Dim lMunPrest As Long
Dim nAfetados As Long

sNac = NFSeTab_CodigoDoCombo(cboCodTribNac)
sMun = SoDigitos(txtCodTribMun.Text)
sNBS = NFSeTab_CodigoDoCombo(cboNBS)
sIndOp = NFSeTab_CodigoDoCombo(cboIndOp)
sClass = NFSeTab_CodigoDoCombo(cboClassTrib)
If Len(sNac) <> 6 Then
   MsgBox "Escolha o serviço (código de tributação nacional) na lista.", vbExclamation, "NFS-e"
   cboCodTribNac.SetFocus
   Exit Function
End If
If sMun <> "" And Len(sMun) <> 3 Then
   MsgBox "O código de tributação municipal tem 3 dígitos (deixe em branco se o município não usa).", vbExclamation, "NFS-e"
   txtCodTribMun.SetFocus
   Exit Function
End If
If mIbsCbs And mTem182 Then
   If sNBS = "" Or Len(sIndOp) <> 6 Or Len(sClass) <> 6 Then
      MsgBox "Com IBS/CBS ligado no Configurar NFS-e, escolha o NBS, o indicador da operação e a classificação (cClassTrib)." & vbCrLf & _
             "Se o serviço não tem NBS na lista, ele não pode ir com IBS/CBS.", vbExclamation, "NFS-e"
      cboNBS.SetFocus
      Exit Function
   End If
End If
If optCliente.Value Then
   lMunPrest = MunicipioEscolhido()
   If lMunPrest = 0 Then
      MsgBox "Escolha a UF e o município onde o serviço foi feito.", vbExclamation, "NFS-e"
      cboMunPrest.SetFocus
      Exit Function
   End If
Else
   lMunPrest = mMunEmpresa
End If
If Not IsDate(txtCompetencia.Text) Then
   MsgBox "Data de competência inválida.", vbExclamation, "NFS-e"
   txtCompetencia.SetFocus
   Exit Function
End If
If CDate(txtCompetencia.Text) > Date Then
   MsgBox "A competência não pode ser depois de hoje.", vbExclamation, "NFS-e"
   txtCompetencia.SetFocus
   Exit Function
End If
If Trim$(txtDescricao.Text) = "" Then
   MsgBox "Informe a descrição dos serviços.", vbExclamation, "NFS-e"
   txtDescricao.SetFocus
   Exit Function
End If
sAliq = "NULL"
If chkRetido.Value = vbChecked Then
   If Not IsNumeric(txtAliquota.Text) Then
      MsgBox "Com ISS retido pelo tomador, informe a alíquota do ISS (%).", vbExclamation, "NFS-e"
      txtAliquota.SetFocus
      Exit Function
   End If
   If CDbl(txtAliquota.Text) <= 0 Or CDbl(txtAliquota.Text) > 5 Then
      MsgBox "A alíquota do ISS deve ficar entre 0,01% e 5%.", vbExclamation, "NFS-e"
      txtAliquota.SetFocus
      Exit Function
   End If
   sAliq = NFSe_N(txtAliquota.Text)
End If

dbData.Execute "UPDATE NFSe SET " & _
   "CodTribNac = " & NFSe_Q(sNac) & ", " & _
   "CodTribMun = " & IIf(sMun = "", "NULL", NFSe_Q(sMun)) & ", " & _
   "CodNBS = " & IIf(sNBS = "", "NULL", NFSe_Q(sNBS)) & ", " & _
   "CodMunicipioPrestacao = " & lMunPrest & ", " & _
   IIf(mTem182, "IndOp = " & IIf(sIndOp = "", "NULL", NFSe_Q(sIndOp)) & ", " & _
                "ClassTrib = " & IIf(sClass = "", "NULL", NFSe_Q(sClass)) & ", " & _
                "PrestacaoNoCliente = " & IIf(optCliente.Value, "1", "0") & ", ", "") & _
   "DataCompetencia = '" & Format$(CDate(txtCompetencia.Text), "yyyymmdd") & "', " & _
   "Descricao = " & NFSe_Q(Trim$(txtDescricao.Text)) & ", " & _
   "InfoComplementar = " & IIf(Trim$(txtInfo.Text) = "", "NULL", NFSe_Q(Trim$(txtInfo.Text))) & ", " & _
   "NumeroPedido = " & IIf(Trim$(txtPedido.Text) = "", "NULL", NFSe_Q(Trim$(txtPedido.Text))) & ", " & _
   "RetencaoISS = " & IIf(chkRetido.Value = vbChecked, "2", "1") & ", " & _
   "AliquotaISS = " & sAliq & " " & _
   "WHERE (IdNFSe = " & mIdNFSe & ") AND (Situacao IN (0, 2));", nAfetados
If nAfetados = 0 Then
   MsgBox "A NFS-e não pôde ser alterada (já foi transmitida em outro computador?).", vbExclamation, "NFS-e"
   Exit Function
End If
Salvar = True
End Function

Private Sub cmdSalvar_Click()
If Salvar() Then Me.Hide
End Sub

Private Sub cmdSalvarTransmitir_Click()
If Not Salvar() Then Exit Sub
mTransmitir = True
Me.Hide
End Sub

Private Sub cmdVoltar_Click()
mTransmitir = False
Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
   Cancel = True
   cmdVoltar_Click
End If
End Sub
