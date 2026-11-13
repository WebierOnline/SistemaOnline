VERSION 5.00
Begin VB.Form frmNFSe_Config
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configurar NFS-e"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraEmpresa
      Caption         =   "Empresa (vem do Cadastro da Empresa do Online Commerce)"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   8880
      Begin VB.Label lblEmpresa
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8640
      End
   End
   Begin VB.Frame fraAmbiente
      Caption         =   "Ambiente"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1020
      Width           =   4335
      Begin VB.OptionButton optHomologacao
         Caption         =   "Homologação (testes)"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   330
         Width           =   2055
      End
      Begin VB.OptionButton optProducao
         Caption         =   "Produção"
         Height          =   255
         Left            =   2400
         TabIndex        =   4
         Top             =   330
         Width           =   1335
      End
   End
   Begin VB.Frame fraNumeracao
      Caption         =   "Numeração da DPS"
      Height          =   735
      Left            =   4560
      TabIndex        =   5
      Top             =   1020
      Width           =   4440
      Begin VB.TextBox txtSerie
         Height          =   315
         Left            =   600
         MaxLength       =   5
         TabIndex        =   7
         Top             =   285
         Width           =   735
      End
      Begin VB.TextBox txtUltimo
         Height          =   315
         Left            =   2880
         MaxLength       =   15
         TabIndex        =   9
         Top             =   285
         Width           =   1395
      End
      Begin VB.Label lblSerie
         Caption         =   "Série:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   330
         Width           =   495
      End
      Begin VB.Label lblUltimo
         Caption         =   "Último número:"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   330
         Width           =   1215
      End
   End
   Begin VB.Frame fraServico
      Caption         =   "Serviço padrão (lista nacional da NFS-e)"
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   1860
      Width           =   8880
      Begin VB.TextBox txtCodTribNac
         Height          =   315
         Left            =   2040
         MaxLength       =   8
         TabIndex        =   12
         Top             =   285
         Width           =   1035
      End
      Begin VB.TextBox txtCodTribMun
         Height          =   315
         Left            =   4200
         MaxLength       =   3
         TabIndex        =   14
         Top             =   285
         Width           =   555
      End
      Begin VB.TextBox txtNBS
         Height          =   315
         Left            =   5400
         MaxLength       =   12
         TabIndex        =   16
         Top             =   285
         Width           =   1335
      End
      Begin VB.Label lblCodNac
         Caption         =   "Cód. tributação nacional:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   330
         Width           =   1935
      End
      Begin VB.Label lblCodMun
         Caption         =   "Municipal:"
         Height          =   255
         Left            =   3360
         TabIndex        =   13
         Top             =   330
         Width           =   855
      End
      Begin VB.Label lblNBS
         Caption         =   "NBS:"
         Height          =   255
         Left            =   4920
         TabIndex        =   15
         Top             =   330
         Width           =   435
      End
      Begin VB.Label lblDicaCod
         Caption         =   "Ex.: 140101 = conserto/manutenção"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   6840
         TabIndex        =   17
         Top             =   330
         Width           =   1935
      End
   End
   Begin VB.Frame fraTributacao
      Caption         =   "Tributação"
      Height          =   1815
      Left            =   120
      TabIndex        =   18
      Top             =   2700
      Width           =   8880
      Begin VB.ComboBox cboRegimeEspecial
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   270
         Width           =   3255
      End
      Begin VB.ComboBox cboRegimeApuracao
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   660
         Width           =   6255
      End
      Begin VB.TextBox txtAliquotaISS
         Height          =   315
         Left            =   2520
         MaxLength       =   6
         TabIndex        =   24
         Top             =   1050
         Width           =   735
      End
      Begin VB.TextBox txtTotTribSN
         Height          =   315
         Left            =   5160
         MaxLength       =   6
         TabIndex        =   26
         Top             =   1050
         Width           =   735
      End
      Begin VB.TextBox txtTotTribFed
         Height          =   315
         Left            =   7080
         MaxLength       =   6
         TabIndex        =   28
         Top             =   1050
         Width           =   615
      End
      Begin VB.TextBox txtTotTribMun
         Height          =   315
         Left            =   8160
         MaxLength       =   6
         TabIndex        =   30
         Top             =   1050
         Width           =   615
      End
      Begin VB.Label lblRegEsp
         Caption         =   "Regime especial de tributação:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   315
         Width           =   2295
      End
      Begin VB.Label lblRegApur
         Caption         =   "Apuração no Simples (ME/EPP):"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   705
         Width           =   2295
      End
      Begin VB.Label lblAliq
         Caption         =   "Alíquota ISS (%, só ISS retido):"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1095
         Width           =   2400
      End
      Begin VB.Label lblTotSN
         Caption         =   "Tributos SN (%):"
         Height          =   255
         Left            =   3840
         TabIndex        =   25
         Top             =   1095
         Width           =   1335
      End
      Begin VB.Label lblTotFed
         Caption         =   "Fed. (%):"
         Height          =   255
         Left            =   6360
         TabIndex        =   27
         Top             =   1095
         Width           =   735
      End
      Begin VB.Label lblTotMun
         Caption         =   "Mun.:"
         Height          =   255
         Left            =   7740
         TabIndex        =   29
         Top             =   1095
         Width           =   435
      End
      Begin VB.Label lblDicaTrib
         Caption         =   "Tributos aproximados (Lei 12.741): Simples = % do Simples Nacional; regime normal = % federal e municipal."
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1470
         Width           =   8640
      End
   End
   Begin VB.TextBox txtMensagem
      Height          =   555
      Left            =   120
      MaxLength       =   500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   33
      Top             =   4860
      Width           =   8880
   End
   Begin VB.Frame fraIBSCBS
      Caption         =   "IBS / CBS (reforma tributária - opcional em 2026)"
      Height          =   1095
      Left            =   120
      TabIndex        =   34
      Top             =   5520
      Width           =   8880
      Begin VB.CheckBox chkIBSCBS
         Caption         =   "Informar o grupo IBS/CBS na NFS-e (exige o código NBS)"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   300
         Width           =   5175
      End
      Begin VB.TextBox txtIndOp
         Height          =   315
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   37
         Top             =   645
         Width           =   855
      End
      Begin VB.TextBox txtCST
         Height          =   315
         Left            =   3720
         MaxLength       =   3
         TabIndex        =   39
         Top             =   645
         Width           =   555
      End
      Begin VB.TextBox txtClassTrib
         Height          =   315
         Left            =   5400
         MaxLength       =   6
         TabIndex        =   41
         Top             =   645
         Width           =   855
      End
      Begin VB.Label lblIndOp
         Caption         =   "Indicador da operação:"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   690
         Width           =   2055
      End
      Begin VB.Label lblCST
         Caption         =   "CST:"
         Height          =   255
         Left            =   3240
         TabIndex        =   38
         Top             =   690
         Width           =   435
      End
      Begin VB.Label lblClassTrib
         Caption         =   "cClassTrib:"
         Height          =   255
         Left            =   4440
         TabIndex        =   40
         Top             =   690
         Width           =   915
      End
   End
   Begin VB.CheckBox chkSHA256
      Caption         =   "Assinatura SHA-256 (só altere se o suporte pedir)"
      Height          =   255
      Left            =   120
      TabIndex        =   42
      Top             =   6720
      Width           =   4935
   End
   Begin VB.CommandButton cmdSalvar
      Caption         =   "&Salvar"
      Height          =   375
      Left            =   5520
      TabIndex        =   43
      Top             =   7380
      Width           =   1695
   End
   Begin VB.CommandButton cmdVoltar
      Cancel          =   -1  'True
      Caption         =   "&Voltar"
      Height          =   375
      Left            =   7305
      TabIndex        =   44
      Top             =   7380
      Width           =   1695
   End
   Begin VB.Label lblMensagem
      Caption         =   "Mensagem fixa nas informações complementares (ex.: garantia dos serviços):"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   4620
      Width           =   8880
   End
End
Attribute VB_Name = "frmNFSe_Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Parâmetros da NFS-e na tabela empresa (campos NFSe*, script 178). CNPJ, inscrição municipal, município,
'regime (CRT) e certificado vêm do Cadastro da Empresa do Online Commerce (os mesmos da NF-e).
Private mProducaoAntes As Boolean

Private Sub Form_Load()
Dim r As ADODB.Recordset
Dim sCrt As String

cboRegimeEspecial.AddItem "0 - Nenhum"
cboRegimeEspecial.AddItem "1 - Ato Cooperado (Cooperativa)"
cboRegimeEspecial.AddItem "2 - Estimativa"
cboRegimeEspecial.AddItem "3 - Microempresa Municipal"
cboRegimeEspecial.AddItem "4 - Notário ou Registrador"
cboRegimeEspecial.AddItem "5 - Profissional Autônomo"
cboRegimeEspecial.AddItem "6 - Sociedade de Profissionais"
cboRegimeEspecial.AddItem "9 - Outros"
cboRegimeApuracao.AddItem "1 - Tributos federais e municipal pelo Simples Nacional"
cboRegimeApuracao.AddItem "2 - Federais pelo Simples e ISS por fora (legislação municipal)"
cboRegimeApuracao.AddItem "3 - Federais e municipal por fora do Simples"

Set r = dbData.OpenRecordset("SELECT TOP 1 * FROM empresa;")
If r.EOF Then
   r.Close
   MsgBox "Cadastro da empresa não encontrado.", vbExclamation, "NFS-e"
   Exit Sub
End If
Select Case Val(ValidateNull(r("CRT")))
   Case 3: sCrt = "Regime normal (não optante do Simples)"
   Case 4: sCrt = "MEI"
   Case Else: sCrt = "Simples Nacional (ME/EPP)"
End Select
lblEmpresa.Caption = "CNPJ: " & ValidateNull(r("CNPJ")) & "    Inscrição municipal: " & IIf(Trim$(ValidateNull(r("IEMunicipal"))) = "", "(não informada)", ValidateNull(r("IEMunicipal"))) & _
                     "    Município IBGE: " & ValidateNull(r("CodigoIBGE")) & vbCrLf & _
                     "Regime: " & sCrt & "    Certificado digital (série): " & IIf(Trim$(ValidateNull(r("CertificadoDigital"))) = "", "(não informado)", ValidateNull(r("CertificadoDigital")))

mProducaoAntes = (Val(ValidateNull(r("NFSeAmbiente"))) = 1)
optProducao.Value = mProducaoAntes
optHomologacao.Value = Not mProducaoAntes
txtSerie.Text = ValidateNull(r("NFSeSerieDPS"))
txtUltimo.Text = ValidateNull(r("NFSeUltimoDPS"))
txtCodTribNac.Text = ValidateNull(r("NFSeCodTribNac"))
txtCodTribMun.Text = ValidateNull(r("NFSeCodTribMun"))
txtNBS.Text = ValidateNull(r("NFSeCodNBS"))
SelecionarCodigo cboRegimeEspecial, CStr(Val(ValidateNull(r("NFSeRegimeEspecial"))))
SelecionarCodigo cboRegimeApuracao, CStr(Val(ValidateNull(r("NFSeRegimeApuracaoSN"))))
txtAliquotaISS.Text = IIf(IsNull(r("NFSeAliquotaISS")), "", Format$(ValidateNull(r("NFSeAliquotaISS")), "0.00"))
txtTotTribSN.Text = Format$(ValidateNull(r("NFSeTotTribSN")), "0.00")
txtTotTribFed.Text = Format$(ValidateNull(r("NFSeTotTribFed")), "0.00")
txtTotTribMun.Text = Format$(ValidateNull(r("NFSeTotTribMun")), "0.00")
txtMensagem.Text = ValidateNull(r("NFSeMensagem"))
chkIBSCBS.Value = IIf(CBool(ValidateNull(r("NFSeIBSCBS"))), vbChecked, vbUnchecked)
txtIndOp.Text = ValidateNull(r("NFSeIndOp"))
txtCST.Text = ValidateNull(r("NFSeCST"))
txtClassTrib.Text = ValidateNull(r("NFSeClassTrib"))
chkSHA256.Value = IIf(CBool(ValidateNull(r("NFSeAssinaturaSHA256"))), vbChecked, vbUnchecked)
cboRegimeApuracao.Enabled = (Val(ValidateNull(r("CRT"))) <> 3 And Val(ValidateNull(r("CRT"))) <> 4)
lblRegApur.Enabled = cboRegimeApuracao.Enabled
r.Close
Set r = Nothing
chkIBSCBS_Click
End Sub

Private Sub SelecionarCodigo(cbo As ComboBox, ByVal sCodigo As String)
Dim i As Integer
cbo.ListIndex = 0
For i = 0 To cbo.ListCount - 1
   If Left$(cbo.List(i), Len(sCodigo) + 2) = sCodigo & " -" Then cbo.ListIndex = i: Exit For
Next
End Sub

Private Sub chkIBSCBS_Click()
txtIndOp.Enabled = (chkIBSCBS.Value = vbChecked)
txtCST.Enabled = txtIndOp.Enabled
txtClassTrib.Enabled = txtIndOp.Enabled
End Sub

Private Function SoDigitos(ByVal s As String) As String
Dim i As Integer
For i = 1 To Len(s)
   If Mid$(s, i, 1) Like "[0-9]" Then SoDigitos = SoDigitos & Mid$(s, i, 1)
Next
End Function

Private Function Pct(t As TextBox, ByVal sNome As String, ByVal dMax As Double, ByVal bVazioPode As Boolean, ByRef sSql As String) As Boolean
If Trim$(t.Text) = "" And bVazioPode Then sSql = "NULL": Pct = True: Exit Function
If Trim$(t.Text) = "" Then t.Text = "0"
If Not IsNumeric(t.Text) Then
   MsgBox sNome & ": informe um número.", vbExclamation, "NFS-e"
   t.SetFocus
   Exit Function
End If
If CDbl(t.Text) < 0 Or CDbl(t.Text) > dMax Then
   MsgBox sNome & ": valor fora do permitido (0 a " & dMax & ").", vbExclamation, "NFS-e"
   t.SetFocus
   Exit Function
End If
sSql = NFSe_N(t.Text)
Pct = True
End Function

Private Sub cmdSalvar_Click()
Dim sNac As String, sMun As String, sNBS As String
Dim sAliq As String, sSN As String, sFed As String, sMunTrib As String

sNac = SoDigitos(txtCodTribNac.Text)
sMun = SoDigitos(txtCodTribMun.Text)
sNBS = SoDigitos(txtNBS.Text)
If Len(sNac) <> 6 Then
   MsgBox "Informe o código de tributação nacional padrão com 6 dígitos (ex.: 140101).", vbExclamation, "NFS-e"
   txtCodTribNac.SetFocus
   Exit Sub
End If
If sMun <> "" And Len(sMun) <> 3 Then
   MsgBox "O código de tributação municipal tem 3 dígitos (deixe em branco se o município não usa).", vbExclamation, "NFS-e"
   txtCodTribMun.SetFocus
   Exit Sub
End If
If sNBS <> "" And Len(sNBS) <> 9 Then
   MsgBox "O código NBS tem 9 dígitos.", vbExclamation, "NFS-e"
   txtNBS.SetFocus
   Exit Sub
End If
If Val(txtSerie.Text) < 1 Or Val(txtSerie.Text) > 49999 Then
   MsgBox "A série da DPS deve ficar entre 1 e 49999 (faixa de aplicativo próprio).", vbExclamation, "NFS-e"
   txtSerie.SetFocus
   Exit Sub
End If
If Not IsNumeric(txtUltimo.Text) Then txtUltimo.Text = "0"
'diminuir o numero repetiria DPS ja enviadas (o Sefin devolveria a nota antiga no lugar da nova)
If CDbl(txtUltimo.Text) < CDbl(SQLExecutaRetorno("SELECT NFSeUltimoDPS FROM empresa", "NFSeUltimoDPS", "0")) And _
   Val(txtSerie.Text) = Val(SQLExecutaRetorno("SELECT NFSeSerieDPS FROM empresa", "NFSeSerieDPS", "1")) Then
   MsgBox "O último número da DPS não pode diminuir na mesma série (repetiria números já usados)." & vbCrLf & _
          "Para recomeçar a numeração, troque a série.", vbExclamation, "NFS-e"
   txtUltimo.SetFocus
   Exit Sub
End If
If Not Pct(txtAliquotaISS, "Alíquota do ISS", 5, True, sAliq) Then Exit Sub
If Not Pct(txtTotTribSN, "Tributos do Simples", 99, False, sSN) Then Exit Sub
If Not Pct(txtTotTribFed, "Tributos federais", 99, False, sFed) Then Exit Sub
If Not Pct(txtTotTribMun, "Tributos municipais", 99, False, sMunTrib) Then Exit Sub
If chkIBSCBS.Value = vbChecked Then
   If sNBS = "" Or Len(SoDigitos(txtIndOp.Text)) <> 6 Or Len(SoDigitos(txtCST.Text)) <> 3 Or Len(SoDigitos(txtClassTrib.Text)) <> 6 Then
      MsgBox "Para informar IBS/CBS preencha o NBS (9 dígitos), o indicador da operação (6), o CST (3) e o cClassTrib (6).", vbExclamation, "NFS-e"
      Exit Sub
   End If
End If
If optProducao.Value And Not mProducaoAntes Then
   If MsgBox("Mudar para PRODUÇÃO?" & vbCrLf & vbCrLf & "A partir daí as NFS-e transmitidas têm valor fiscal.", vbQuestion + vbYesNo + vbDefaultButton2, "NFS-e") <> vbYes Then Exit Sub
End If

dbData.Execute "UPDATE empresa SET " & _
   "NFSeAmbiente = " & IIf(optProducao.Value, "1", "2") & ", " & _
   "NFSeSerieDPS = " & Val(txtSerie.Text) & ", " & _
   "NFSeUltimoDPS = " & Format$(CDbl(txtUltimo.Text), "0") & ", " & _
   "NFSeCodTribNac = " & NFSe_Q(sNac) & ", " & _
   "NFSeCodTribMun = " & IIf(sMun = "", "NULL", NFSe_Q(sMun)) & ", " & _
   "NFSeCodNBS = " & IIf(sNBS = "", "NULL", NFSe_Q(sNBS)) & ", " & _
   "NFSeRegimeEspecial = " & Val(cboRegimeEspecial.Text) & ", " & _
   "NFSeRegimeApuracaoSN = " & Val(cboRegimeApuracao.Text) & ", " & _
   "NFSeAliquotaISS = " & sAliq & ", " & _
   "NFSeTotTribSN = " & sSN & ", " & _
   "NFSeTotTribFed = " & sFed & ", " & _
   "NFSeTotTribMun = " & sMunTrib & ", " & _
   "NFSeMensagem = " & IIf(Trim$(txtMensagem.Text) = "", "NULL", NFSe_Q(Trim$(txtMensagem.Text))) & ", " & _
   "NFSeIBSCBS = " & IIf(chkIBSCBS.Value = vbChecked, "1", "0") & ", " & _
   "NFSeIndOp = " & IIf(SoDigitos(txtIndOp.Text) = "", "NULL", NFSe_Q(SoDigitos(txtIndOp.Text))) & ", " & _
   "NFSeCST = " & IIf(SoDigitos(txtCST.Text) = "", "NULL", NFSe_Q(SoDigitos(txtCST.Text))) & ", " & _
   "NFSeClassTrib = " & IIf(SoDigitos(txtClassTrib.Text) = "", "NULL", NFSe_Q(SoDigitos(txtClassTrib.Text))) & ", " & _
   "NFSeAssinaturaSHA256 = " & IIf(chkSHA256.Value = vbChecked, "1", "0") & ";"
Unload Me
End Sub

Private Sub cmdVoltar_Click()
Unload Me
End Sub
