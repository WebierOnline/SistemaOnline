VERSION 5.00
Begin VB.Form frmNFSe_Editar
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NFS-e - Conferir antes de transmitir"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9480
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraServ
      Caption         =   "Código do serviço (lista nacional da NFS-e)"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   780
      Width           =   9240
      Begin VB.TextBox txtCodTribNac
         Height          =   315
         Left            =   1440
         MaxLength       =   8
         TabIndex        =   5
         Top             =   270
         Width           =   1035
      End
      Begin VB.TextBox txtCodTribMun
         Height          =   315
         Left            =   3480
         MaxLength       =   3
         TabIndex        =   7
         Top             =   270
         Width           =   555
      End
      Begin VB.TextBox txtNBS
         Height          =   315
         Left            =   4560
         MaxLength       =   12
         TabIndex        =   9
         Top             =   270
         Width           =   1335
      End
      Begin VB.TextBox txtCompetencia
         Height          =   315
         Left            =   7920
         MaxLength       =   10
         TabIndex        =   11
         Top             =   270
         Width           =   1155
      End
      Begin VB.Label lblCodNac
         Caption         =   "Nacional (6 díg.):"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   315
         Width           =   1335
      End
      Begin VB.Label lblCodMun
         Caption         =   "Municipal:"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   315
         Width           =   855
      End
      Begin VB.Label lblNBS
         Caption         =   "NBS:"
         Height          =   255
         Left            =   4140
         TabIndex        =   8
         Top             =   315
         Width           =   435
      End
      Begin VB.Label lblComp
         Caption         =   "Competência:"
         Height          =   255
         Left            =   6840
         TabIndex        =   10
         Top             =   315
         Width           =   1035
      End
   End
   Begin VB.TextBox txtDescricao
      Height          =   2295
      Left            =   120
      MaxLength       =   2000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   1860
      Width           =   9240
   End
   Begin VB.TextBox txtInfo
      Height          =   975
      Left            =   120
      MaxLength       =   2000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   4560
      Width           =   9240
   End
   Begin VB.Frame fraISS
      Caption         =   "ISS"
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   5640
      Width           =   9240
      Begin VB.CheckBox chkRetido
         Caption         =   "ISS retido pelo tomador (só quando o cliente reter o imposto)"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   315
         Width           =   4935
      End
      Begin VB.TextBox txtAliquota
         Height          =   315
         Left            =   6720
         MaxLength       =   6
         TabIndex        =   20
         Top             =   270
         Width           =   795
      End
      Begin VB.Label lblAliq
         Caption         =   "Alíquota do ISS (%):"
         Height          =   255
         Left            =   5160
         TabIndex        =   19
         Top             =   315
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdSalvar
      Caption         =   "&Salvar"
      Height          =   375
      Left            =   3960
      TabIndex        =   22
      Top             =   7200
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalvarTransmitir
      Caption         =   "Salvar e &Transmitir"
      Height          =   375
      Left            =   5760
      TabIndex        =   23
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton cmdVoltar
      Cancel          =   -1  'True
      Caption         =   "&Voltar"
      Height          =   375
      Left            =   7665
      TabIndex        =   24
      Top             =   7200
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
      TabIndex        =   21
      Top             =   6480
      Width           =   9240
   End
   Begin VB.Label lblInfoT
      Caption         =   "Informações complementares (veículo/equipamento, observações):"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4320
      Width           =   6135
   End
   Begin VB.Label lblContDesc
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6360
      TabIndex        =   14
      Top             =   1620
      Width           =   3000
   End
   Begin VB.Label lblDesc
      Caption         =   "Descrição dos serviços (vai na nota):"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1620
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
      Top             =   6960
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
Private mId As Long
Private mTransmitir As Boolean

'Mostra a nota e devolve True quando o usuário escolheu "Salvar e Transmitir"
Public Function Editar(ByVal lId As Long) As Boolean
mId = lId
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

Private Function Carregar() As Boolean
Dim r As ADODB.Recordset
Dim sDoc As String, sEnd As String

Set r = dbData.OpenRecordset("SELECT * FROM NFSe WHERE (IdNFSe = " & mId & ");")
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

sDoc = ValidateNull(r("TomadorCpfCnpj"))
If Len(sDoc) = 14 Then
   sDoc = Format$(sDoc, "@@.@@@.@@@/@@@@-@@")
ElseIf Len(sDoc) = 11 Then
   sDoc = Format$(sDoc, "@@@.@@@.@@@-@@")
Else
   sDoc = "SEM CPF/CNPJ (a nota sai sem tomador identificado)"
End If
lblTomador.Caption = "Tomador: " & ValidateNull(r("TomadorNome")) & "  -  " & sDoc & "      " & ValidateNull(r("NumeroPedido"))
sEnd = Trim$(ValidateNull(r("TomadorLogradouro")))
If sEnd <> "" Then sEnd = sEnd & ", " & ValidateNull(r("TomadorNumero")) & " - " & ValidateNull(r("TomadorBairro")) & " - IBGE " & ValidateNull(r("TomadorCodMunicipio")) & " - CEP " & ValidateNull(r("TomadorCEP"))
If sEnd = "" Then sEnd = "Endereço do cliente incompleto (sem logradouro ou código IBGE): a nota vai sem o endereço do tomador."
lblTomadorEnd.Caption = sEnd & "   (dados copiados do cadastro do cliente quando a nota foi gerada)"

txtCodTribNac.Text = ValidateNull(r("CodTribNac"))
txtCodTribMun.Text = ValidateNull(r("CodTribMun"))
txtNBS.Text = ValidateNull(r("CodNBS"))
txtCompetencia.Text = Format$(r("DataCompetencia"), "dd/mm/yyyy")
txtDescricao.Text = ValidateNull(r("Descricao"))
txtInfo.Text = ValidateNull(r("InfoComplementar"))
chkRetido.Value = IIf(Val(ValidateNull(r("RetencaoISS"))) = 2, vbChecked, vbUnchecked)
txtAliquota.Text = IIf(IsNull(r("AliquotaISS")), "", Format$(ValidateNull(r("AliquotaISS")), "0.00"))
lblValores.Caption = "Serviços: R$ " & Format$(ValidateNull(r("ValorServicos")), ocMONEY) & _
                     "     Desconto: R$ " & Format$(ValidateNull(r("DescontoIncondicionado")), ocMONEY) & _
                     "     Líquido: R$ " & Format$(CDbl(ValidateNull(r("ValorServicos"))) - CDbl(ValidateNull(r("DescontoIncondicionado"))), ocMONEY) & vbCrLf & _
                     "(valores vindos das OS; para mudar, ajuste a OS, exclua este rascunho e gere de novo)"
lblMensagem.Caption = IIf(Val(ValidateNull(r("Situacao"))) = NFSE_REJEITADA, "Rejeição: " & ValidateNull(r("Mensagem")), "")
r.Close
Set r = Nothing
If txtAliquota.Text = "" Then txtAliquota.Text = Format$(SQLExecutaRetorno("SELECT NFSeAliquotaISS FROM empresa", "NFSeAliquotaISS", ""), "0.00")
chkRetido_Click
txtDescricao_Change
Carregar = True
End Function

Private Sub chkRetido_Click()
txtAliquota.Enabled = (chkRetido.Value = vbChecked)
lblAliq.Enabled = txtAliquota.Enabled
End Sub

Private Sub txtDescricao_Change()
lblContDesc.Caption = Len(txtDescricao.Text) & " de 2000 caracteres"
End Sub

Private Function SoDigitos(ByVal s As String) As String
Dim i As Integer
For i = 1 To Len(s)
   If Mid$(s, i, 1) Like "[0-9]" Then SoDigitos = SoDigitos & Mid$(s, i, 1)
Next
End Function

Private Function Salvar() As Boolean
Dim sNac As String, sMun As String, sNBS As String, sAliq As String
Dim nAfetados As Long

sNac = SoDigitos(txtCodTribNac.Text)
sMun = SoDigitos(txtCodTribMun.Text)
sNBS = SoDigitos(txtNBS.Text)
If Len(sNac) <> 6 Then
   MsgBox "Informe o código de tributação nacional com 6 dígitos (item, subitem e desdobro da lista nacional, ex.: 140101).", vbExclamation, "NFS-e"
   txtCodTribNac.SetFocus
   Exit Function
End If
If sMun <> "" And Len(sMun) <> 3 Then
   MsgBox "O código de tributação municipal tem 3 dígitos (deixe em branco se o município não usa).", vbExclamation, "NFS-e"
   txtCodTribMun.SetFocus
   Exit Function
End If
If sNBS <> "" And Len(sNBS) <> 9 Then
   MsgBox "O código NBS tem 9 dígitos (deixe em branco se não usar).", vbExclamation, "NFS-e"
   txtNBS.SetFocus
   Exit Function
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
   "DataCompetencia = '" & Format$(CDate(txtCompetencia.Text), "yyyymmdd") & "', " & _
   "Descricao = " & NFSe_Q(Trim$(txtDescricao.Text)) & ", " & _
   "InfoComplementar = " & IIf(Trim$(txtInfo.Text) = "", "NULL", NFSe_Q(Trim$(txtInfo.Text))) & ", " & _
   "RetencaoISS = " & IIf(chkRetido.Value = vbChecked, "2", "1") & ", " & _
   "AliquotaISS = " & sAliq & " " & _
   "WHERE (IdNFSe = " & mId & ") AND (Situacao IN (0, 2));", nAfetados
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
