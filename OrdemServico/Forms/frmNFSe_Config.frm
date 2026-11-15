VERSION 5.00
Begin VB.Form frmNFSe_Config
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configurar NFS-e"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
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
      Height          =   1860
      Left            =   120
      TabIndex        =   10
      Top             =   1860
      Width           =   8880
      Begin VB.ComboBox cboCodTribNac
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   510
         Width           =   7440
      End
      Begin VB.TextBox txtCodTribMun
         Height          =   315
         Left            =   7680
         MaxLength       =   3
         TabIndex        =   14
         Top             =   510
         Width           =   1080
      End
      Begin VB.ComboBox cboNBS
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1140
         Width           =   8640
      End
      Begin VB.Label lblCodNac
         Caption         =   "Serviço (código de tributação nacional):"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   270
         Width           =   6000
      End
      Begin VB.Label lblCodMun
         Caption         =   "Cód. municipal:"
         Height          =   255
         Left            =   7680
         TabIndex        =   13
         Top             =   270
         Width           =   1080
      End
      Begin VB.Label lblNBS
         Caption         =   "NBS (Nomenclatura Brasileira de Serviços) - só os aceitos para o serviço escolhido:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   900
         Width           =   8640
      End
      Begin VB.Label lblDicaCod
         Caption         =   "As listas seguem as tabelas oficiais (Anexos B, VII e VIII): ao trocar o serviço, o NBS, o indicador e a classificação se ajustam."
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1530
         Width           =   8640
      End
   End
   Begin VB.Frame fraTributacao
      Caption         =   "Tributação"
      Height          =   1815
      Left            =   120
      TabIndex        =   18
      Top             =   3825
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
      Begin VB.ComboBox cboFormaTotTrib
         Height          =   315
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   46
         ToolTipText     =   "Valor aproximado dos tributos (Lei 12.741) - as mesmas opções do portal nacional (Perfil > Configuração)"
         Top             =   1050
         Width           =   3255
      End
      Begin VB.TextBox txtTotTribSN
         Height          =   315
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   26
         Top             =   1440
         Width           =   675
      End
      Begin VB.TextBox txtTotTribFed
         Height          =   315
         Left            =   3120
         MaxLength       =   6
         TabIndex        =   28
         Top             =   1440
         Width           =   675
      End
      Begin VB.TextBox txtTotTribEst
         Height          =   315
         Left            =   5160
         MaxLength       =   6
         TabIndex        =   48
         Top             =   1440
         Width           =   675
      End
      Begin VB.TextBox txtTotTribMun
         Height          =   315
         Left            =   7200
         MaxLength       =   6
         TabIndex        =   30
         Top             =   1440
         Width           =   675
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
      Begin VB.Label lblFormaTot
         Caption         =   "Tributos aproximados:"
         Height          =   255
         Left            =   3840
         TabIndex        =   31
         ToolTipText     =   "Lei 12.741 (valor aproximado dos tributos)"
         Top             =   1095
         Width           =   1635
      End
      Begin VB.Label lblTotSN
         Caption         =   "Simples (%):"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1485
         Width           =   1035
      End
      Begin VB.Label lblTotFed
         Caption         =   "Federal (%):"
         Height          =   255
         Left            =   2100
         TabIndex        =   27
         Top             =   1485
         Width           =   975
      End
      Begin VB.Label lblTotEst
         Caption         =   "Estadual (%):"
         Height          =   255
         Left            =   4020
         TabIndex        =   47
         Top             =   1485
         Width           =   1095
      End
      Begin VB.Label lblTotMun
         Caption         =   "Municipal (%):"
         Height          =   255
         Left            =   6060
         TabIndex        =   29
         Top             =   1485
         Width           =   1095
      End
   End
   Begin VB.TextBox txtMensagem
      Height          =   555
      Left            =   120
      MaxLength       =   500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   33
      Top             =   5985
      Width           =   8880
   End
   Begin VB.Frame fraIBSCBS
      Caption         =   "IBS / CBS (reforma tributária - opcional em 2026)"
      Height          =   1455
      Left            =   120
      TabIndex        =   34
      Top             =   6645
      Width           =   8880
      Begin VB.CheckBox chkIBSCBS
         Caption         =   "Informar o grupo IBS/CBS na NFS-e (exige o NBS do serviço)"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   300
         Width           =   8640
      End
      Begin VB.ComboBox cboIndOp
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   645
         Width           =   6600
      End
      Begin VB.TextBox txtCST
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   8040
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1035
         Width           =   720
      End
      Begin VB.ComboBox cboClassTrib
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   1035
         Width           =   5280
      End
      Begin VB.Label lblIndOp
         Caption         =   "Indicador da operação:"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   690
         Width           =   2000
      End
      Begin VB.Label lblCST
         Caption         =   "CST:"
         Height          =   255
         Left            =   7560
         TabIndex        =   40
         Top             =   1080
         Width           =   435
      End
      Begin VB.Label lblClassTrib
         Caption         =   "Classificação (cClassTrib):"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1080
         Width           =   2000
      End
   End
   Begin VB.CheckBox chkSHA256
      Caption         =   "Assinatura SHA-256 (só altere se o suporte pedir)"
      Height          =   255
      Left            =   120
      TabIndex        =   42
      Top             =   8205
      Width           =   4935
   End
   Begin VB.CommandButton cmdConvenio
      Caption         =   "Verificar &município no Sefin"
      Height          =   375
      Left            =   120
      TabIndex        =   45
      ToolTipText     =   "Consulta se o município da empresa tem convênio com a NFS-e Nacional no ambiente escolhido (erro E0037)"
      Top             =   8580
      Width           =   2535
   End
   Begin VB.CommandButton cmdSalvar
      Caption         =   "&Salvar"
      Height          =   375
      Left            =   5520
      TabIndex        =   43
      Top             =   8580
      Width           =   1695
   End
   Begin VB.CommandButton cmdVoltar
      Cancel          =   -1  'True
      Caption         =   "&Voltar"
      Height          =   375
      Left            =   7305
      TabIndex        =   44
      Top             =   8580
      Width           =   1695
   End
   Begin VB.Label lblMensagem
      Caption         =   "Mensagem fixa nas informações complementares (ex.: garantia dos serviços):"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   5745
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
'Serviço, NBS, indicador da operação e cClassTrib saem das tabelas oficiais (arquivos NFSe_*.txt da pasta NFSe,
'os mesmos que o NFSeEmissor usa para conferir antes de transmitir): cada lista só oferece o que combina com a
'escolha anterior (Anexo VIII), então não dá para escolher, por exemplo, indicador de imóvel para oficina.
Private mProducaoAntes As Boolean
Private mCarregando As Boolean
Private mCrt As Integer               'CRT da empresa: 3 regime normal, 4 MEI, demais Simples (ME/EPP)
Private mTemFormaTotTrib As Boolean   'script 181 aplicado (NFSeFormaTotTrib / NFSeTotTribEst)

'Forma dos tributos aproximados (Lei 12.741) = opcoes do portal nacional. O NFSeEmissor confere de novo pelo regime.
Private Const TOT_AUTO = 0, TOT_VALORES = 1, TOT_PERCENTUAIS = 2, TOT_SIMPLES = 3

'Anexo B: serviços (código de tributação nacional)
Private mSrvCod() As String, mSrvDesc() As String, mSrvQtd As Long
'Anexo VII: indicadores da operação usados na NFS-e
Private mIndCod() As String, mIndTxt() As String, mIndImovel() As Boolean, mIndQtd As Long
'Anexo VIII: item da LC 116 x NBS x indOp x cClassTrib
Private mCorItem() As String, mCorNBS() As String, mCorNBSDesc() As String, mCorInd() As String
Private mCorClass() As String, mCorClassDesc() As String, mCorQtd As Long
Private mTabelasOk As Boolean

Private Const CB_SETDROPPEDWIDTH = &H160
Private Declare Function NFSeCfg_SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub Form_Load()
Dim r As ADODB.Recordset
Dim sCrt As String

mCarregando = True
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

'lista aberta mais larga que a caixa (descrições longas), em pixels
NFSeCfg_SendMessage cboCodTribNac.hwnd, CB_SETDROPPEDWIDTH, 900, ByVal 0&
NFSeCfg_SendMessage cboNBS.hwnd, CB_SETDROPPEDWIDTH, 760, ByVal 0&
NFSeCfg_SendMessage cboIndOp.hwnd, CB_SETDROPPEDWIDTH, 820, ByVal 0&
NFSeCfg_SendMessage cboClassTrib.hwnd, CB_SETDROPPEDWIDTH, 760, ByVal 0&
CarregarTabelas

Set r = dbData.OpenRecordset("SELECT TOP 1 * FROM empresa;")
If r.EOF Then
   r.Close
   mCarregando = False
   MsgBox "Cadastro da empresa não encontrado.", vbExclamation, "NFS-e"
   Exit Sub
End If
mCrt = Val(ValidateNull(r("CRT")))
Select Case mCrt
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
txtCodTribMun.Text = ValidateNull(r("NFSeCodTribMun"))
SelecionarCodigo cboRegimeEspecial, CStr(Val(ValidateNull(r("NFSeRegimeEspecial"))))
SelecionarCodigo cboRegimeApuracao, CStr(Val(ValidateNull(r("NFSeRegimeApuracaoSN"))))
txtAliquotaISS.Text = IIf(IsNull(r("NFSeAliquotaISS")), "", Format$(ValidateNull(r("NFSeAliquotaISS")), "0.00"))
txtTotTribSN.Text = Format$(ValidateNull(r("NFSeTotTribSN")), "0.00")
txtTotTribFed.Text = Format$(ValidateNull(r("NFSeTotTribFed")), "0.00")
txtTotTribMun.Text = Format$(ValidateNull(r("NFSeTotTribMun")), "0.00")
mTemFormaTotTrib = TemCampo(r, "NFSeFormaTotTrib") And TemCampo(r, "NFSeTotTribEst")
If mTemFormaTotTrib Then
   txtTotTribEst.Text = Format$(ValidateNull(r("NFSeTotTribEst")), "0.00")
   PreencherFormaTotTrib Val(ValidateNull(r("NFSeFormaTotTrib")))
Else
   txtTotTribEst.Text = "0.00"
   PreencherFormaTotTrib TOT_AUTO
End If
txtMensagem.Text = ValidateNull(r("NFSeMensagem"))
chkIBSCBS.Value = IIf(CBool(ValidateNull(r("NFSeIBSCBS"))), vbChecked, vbUnchecked)
chkSHA256.Value = IIf(CBool(ValidateNull(r("NFSeAssinaturaSHA256"))), vbChecked, vbUnchecked)
cboRegimeApuracao.Enabled = (Val(ValidateNull(r("CRT"))) <> 3 And Val(ValidateNull(r("CRT"))) <> 4)
lblRegApur.Enabled = cboRegimeApuracao.Enabled

'serviço -> NBS -> indicador -> cClassTrib, cada um preferindo o valor gravado
PreencherServicos SoDigitos(ValidateNull(r("NFSeCodTribNac")))
PreencherNBS SoDigitos(ValidateNull(r("NFSeCodNBS")))
PreencherIndOp SoDigitos(ValidateNull(r("NFSeIndOp")))
PreencherClassTrib SoDigitos(ValidateNull(r("NFSeClassTrib")))
r.Close
Set r = Nothing
mCarregando = False
chkIBSCBS_Click

If Not mTabelasOk Then
   MsgBox "As tabelas oficiais da NFS-e não foram encontradas em " & appPathApp & "NFSe\ (NFSe_Servicos.txt, NFSe_IndOp.txt e NFSe_Correlacao.txt)." & vbCrLf & vbCrLf & _
          "Atualize o sistema. Até lá só os códigos já gravados aparecem nas listas.", vbExclamation, "NFS-e"
End If
End Sub

'---------------------------------------------------------------------------------------------------- tabelas

Private Function LerTabela(ByVal sArquivo As String, ByVal nCampos As Integer, ByRef c As Collection) As Boolean
Dim nArq As Integer
Dim sLinha As String, sCaminho As String
sCaminho = appPathApp & "NFSe\" & sArquivo
Set c = New Collection
If Dir$(sCaminho) = "" Then Exit Function
nArq = FreeFile
Open sCaminho For Input As #nArq
Do While Not EOF(nArq)
   Line Input #nArq, sLinha
   If UBound(Split(sLinha, ";")) >= nCampos - 1 Then c.Add sLinha
Loop
Close #nArq
LerTabela = (c.Count > 0)
End Function

Private Sub CarregarTabelas()
Dim c As Collection
Dim v As Variant, p() As String
Dim i As Long

On Error GoTo deuErro
mTabelasOk = False
If Not LerTabela("NFSe_Servicos.txt", 2, c) Then Exit Sub
mSrvQtd = c.Count
ReDim mSrvCod(1 To mSrvQtd): ReDim mSrvDesc(1 To mSrvQtd)
i = 0
For Each v In c
   i = i + 1
   p = Split(v, ";")
   mSrvCod(i) = p(0)
   mSrvDesc(i) = p(1)
Next

If Not LerTabela("NFSe_IndOp.txt", 4, c) Then Exit Sub
mIndQtd = c.Count
ReDim mIndCod(1 To mIndQtd): ReDim mIndTxt(1 To mIndQtd): ReDim mIndImovel(1 To mIndQtd)
i = 0
For Each v In c
   i = i + 1
   p = Split(v, ";")
   mIndCod(i) = p(0)
   mIndTxt(i) = p(1) & " - " & p(2)
   mIndImovel(i) = (p(3) = "1")
Next

If Not LerTabela("NFSe_Correlacao.txt", 6, c) Then Exit Sub
mCorQtd = c.Count
ReDim mCorItem(1 To mCorQtd): ReDim mCorNBS(1 To mCorQtd): ReDim mCorNBSDesc(1 To mCorQtd)
ReDim mCorInd(1 To mCorQtd): ReDim mCorClass(1 To mCorQtd): ReDim mCorClassDesc(1 To mCorQtd)
i = 0
For Each v In c
   i = i + 1
   p = Split(v, ";")
   mCorItem(i) = p(0)
   mCorNBS(i) = p(1)
   mCorNBSDesc(i) = p(2)
   mCorInd(i) = p(3)
   mCorClass(i) = p(4)
   mCorClassDesc(i) = p(5)
Next
mTabelasOk = True
Exit Sub

deuErro:
mTabelasOk = False
mSrvQtd = 0: mIndQtd = 0: mCorQtd = 0
End Sub

'Código no começo do texto do combo ("140101 - ...", "1.2001.31.10 - ..."); "(...)" = nada escolhido
Private Function CodigoDoCombo(cbo As ComboBox) As String
Dim t As String
If cbo.ListIndex < 0 Then Exit Function
t = cbo.List(cbo.ListIndex)
If Left$(t, 1) = "(" Then Exit Function
If InStr(t, " - ") > 0 Then t = Left$(t, InStr(t, " - ") - 1)
CodigoDoCombo = SoDigitos(t)
End Function

Private Sub SelecionarPorCodigo(cbo As ComboBox, ByVal sCodigo As String)
Dim i As Integer
For i = 0 To cbo.ListCount - 1
   If SoDigitos(Split(cbo.List(i) & " - ", " - ")(0)) = sCodigo And sCodigo <> "" Then
      cbo.ListIndex = i
      Exit Sub
   End If
Next
End Sub

Private Function Curto(ByVal s As String, ByVal n As Integer) As String
Curto = IIf(Len(s) > n, Left$(s, n - 3) & "...", s)
End Function

Private Function NBSFormatado(ByVal s As String) As String
If Len(s) = 9 Then
   NBSFormatado = Left$(s, 1) & "." & Mid$(s, 2, 4) & "." & Mid$(s, 6, 2) & "." & Mid$(s, 8, 2)
Else
   NBSFormatado = s
End If
End Function

'Item do Anexo VIII do serviço: desdobro ("99.03.02") se a tabela tiver, senão subitem ("14.01")
Private Function ItemDoServico(ByVal sCod As String) As String
Dim i As Long, sDesd As String
If Len(sCod) <> 6 Then Exit Function
sDesd = Left$(sCod, 2) & "." & Mid$(sCod, 3, 2) & "." & Mid$(sCod, 5, 2)
For i = 1 To mCorQtd
   If mCorItem(i) = sDesd Then ItemDoServico = sDesd: Exit Function
Next
ItemDoServico = Left$(sCod, 2) & "." & Mid$(sCod, 3, 2)
End Function

Private Function JaNaLista(cbo As ComboBox, ByVal sCodigo As String) As Boolean
Dim i As Integer
For i = 0 To cbo.ListCount - 1
   If SoDigitos(Split(cbo.List(i) & " - ", " - ")(0)) = sCodigo Then JaNaLista = True: Exit Function
Next
End Function

Private Function IndiceIndOp(ByVal sCod As String) As Long
Dim i As Long
For i = 1 To mIndQtd
   If mIndCod(i) = sCod Then IndiceIndOp = i: Exit Function
Next
End Function

'---------------------------------------------------------------------------------------------------- listas

Private Sub PreencherServicos(ByVal sPreferido As String)
Dim i As Long
cboCodTribNac.Clear
For i = 1 To mSrvQtd
   cboCodTribNac.AddItem mSrvCod(i) & " - " & Curto(mSrvDesc(i), 160)
Next
If sPreferido <> "" And Not JaNaLista(cboCodTribNac, sPreferido) Then
   cboCodTribNac.AddItem sPreferido & " - (código gravado, fora da lista nacional: escolha outro)", 0
End If
SelecionarPorCodigo cboCodTribNac, sPreferido
End Sub

Private Sub PreencherNBS(ByVal sPreferido As String)
Dim i As Long, sItem As String, n As Long
cboNBS.Clear
cboNBS.AddItem "(sem NBS - não informar)"
sItem = ItemDoServico(CodigoDoCombo(cboCodTribNac))
For i = 1 To mCorQtd
   If mCorItem(i) = sItem Then
      If Not JaNaLista(cboNBS, mCorNBS(i)) Then
         cboNBS.AddItem NBSFormatado(mCorNBS(i)) & " - " & Curto(mCorNBSDesc(i), 150)
         n = n + 1
      End If
   End If
Next
'sem tabelas (ou NBS gravado fora dela): mantém o gravado para não perder
If sPreferido <> "" And Not JaNaLista(cboNBS, sPreferido) And (Not mTabelasOk Or n = 0) Then
   cboNBS.AddItem NBSFormatado(sPreferido) & " - (NBS gravado)"
End If
cboNBS.ListIndex = 0
SelecionarPorCodigo cboNBS, sPreferido
If cboNBS.ListIndex = 0 And n = 1 Then cboNBS.ListIndex = 1
If mTabelasOk And sItem <> "" And n = 0 Then
   lblDicaCod.Caption = "Este serviço não tem NBS na tabela oficial (Anexo VIII): o grupo IBS/CBS não pode ser informado."
Else
   lblDicaCod.Caption = "As listas seguem as tabelas oficiais (Anexos B, VII e VIII): ao trocar o serviço, o NBS, o indicador e a classificação se ajustam."
End If
End Sub

'indOp do Anexo VIII para serviço+NBS, os do mesmo grupo (050101 -> 050102) e os da regra de operação não onerosa/
'exterior (100101/100301/100501); os de imóvel ficam de fora (exigem os dados do imóvel, que a OS não tem)
Private Sub PreencherIndOp(ByVal sPreferido As String)
Dim i As Long, j As Long, sItem As String, sNBS As String, sPrimeiro As String
cboIndOp.Clear
sItem = ItemDoServico(CodigoDoCombo(cboCodTribNac))
sNBS = CodigoDoCombo(cboNBS)
If sNBS <> "" Then
   For i = 1 To mCorQtd
      If mCorItem(i) = sItem And mCorNBS(i) = sNBS Then
         If sPrimeiro = "" Then sPrimeiro = mCorInd(i)
         For j = 1 To mIndQtd
            If Left$(mIndCod(j), 4) = Left$(mCorInd(i), 4) Then AdicionarIndOp mIndCod(j)
         Next
         Select Case mCorInd(i)
            Case "100101": AdicionarIndOp "100102": AdicionarIndOp "100201"
            Case "100301": AdicionarIndOp "100302": AdicionarIndOp "100401"
            Case "100501": AdicionarIndOp "100502": AdicionarIndOp "100601"
         End Select
      End If
   Next
End If
If sPreferido <> "" And Not JaNaLista(cboIndOp, sPreferido) And Not mTabelasOk Then
   cboIndOp.AddItem sPreferido & " - (indicador gravado)"
End If
SelecionarPorCodigo cboIndOp, sPreferido
If cboIndOp.ListIndex < 0 Then SelecionarPorCodigo cboIndOp, sPrimeiro
If cboIndOp.ListIndex < 0 And cboIndOp.ListCount > 0 Then cboIndOp.ListIndex = 0
End Sub

Private Sub AdicionarIndOp(ByVal sCod As String)
Dim k As Long
k = IndiceIndOp(sCod)
If k = 0 Then Exit Sub
If mIndImovel(k) Then Exit Sub
If JaNaLista(cboIndOp, sCod) Then Exit Sub
cboIndOp.AddItem sCod & " - " & Curto(mIndTxt(k), 200)
End Sub

Private Sub PreencherClassTrib(ByVal sPreferido As String)
Dim i As Long, sItem As String, sNBS As String
cboClassTrib.Clear
sItem = ItemDoServico(CodigoDoCombo(cboCodTribNac))
sNBS = CodigoDoCombo(cboNBS)
If sNBS <> "" Then
   For i = 1 To mCorQtd
      If mCorItem(i) = sItem And mCorNBS(i) = sNBS Then
         If Not JaNaLista(cboClassTrib, mCorClass(i)) Then cboClassTrib.AddItem mCorClass(i) & " - " & Curto(mCorClassDesc(i), 150)
      End If
   Next
End If
If sPreferido <> "" And Not JaNaLista(cboClassTrib, sPreferido) And Not mTabelasOk Then
   cboClassTrib.AddItem sPreferido & " - (classificação gravada)"
End If
SelecionarPorCodigo cboClassTrib, sPreferido
If cboClassTrib.ListIndex < 0 And cboClassTrib.ListCount > 0 Then cboClassTrib.ListIndex = 0
cboClassTrib_Click
End Sub

Private Sub cboCodTribNac_Click()
If mCarregando Then Exit Sub
PreencherNBS CodigoDoCombo(cboNBS)
cboNBS_Click
End Sub

Private Sub cboNBS_Click()
If mCarregando Then Exit Sub
PreencherIndOp CodigoDoCombo(cboIndOp)
PreencherClassTrib CodigoDoCombo(cboClassTrib)
If CodigoDoCombo(cboNBS) = "" And chkIBSCBS.Value = vbChecked Then chkIBSCBS.Value = vbUnchecked
End Sub

'CST = 3 primeiros dígitos do cClassTrib
Private Sub cboClassTrib_Click()
txtCST.Text = Left$(CodigoDoCombo(cboClassTrib), 3)
End Sub

'---------------------------------------------------------------------------------------------------- tela

Private Sub SelecionarCodigo(cbo As ComboBox, ByVal sCodigo As String)
Dim i As Integer
cbo.ListIndex = 0
For i = 0 To cbo.ListCount - 1
   If Left$(cbo.List(i), Len(sCodigo) + 2) = sCodigo & " -" Then cbo.ListIndex = i: Exit For
Next
End Sub

'---------------------------------------------------------------------------------------------------- tributos aproximados
Private Function TemCampo(r As ADODB.Recordset, ByVal sCampo As String) As Boolean
Dim f As ADODB.Field
For Each f In r.Fields
   If StrComp(f.Name, sCampo, vbTextCompare) = 0 Then TemCampo = True: Exit Function
Next
End Function

'So as formas que o Sefin aceita para o regime: ME/EPP as 3 do portal; regime normal R$ ou %; MEI nao informa
'(pTotTribSN fora do ME/EPP = E0710/E0713; "nao informar" fora do MEI = E0712/E0713)
Private Sub PreencherFormaTotTrib(ByVal nForma As Integer)
cboFormaTotTrib.Clear
If mCrt = 4 Then
   cboFormaTotTrib.AddItem "0 - Não informar (MEI)"
   cboFormaTotTrib.ListIndex = 0
   cboFormaTotTrib.Enabled = False
   AjustarTotTrib
   Exit Sub
End If
cboFormaTotTrib.Enabled = True
cboFormaTotTrib.AddItem TOT_VALORES & " - Valores em R$ (calculados pelos %)"
cboFormaTotTrib.AddItem TOT_PERCENTUAIS & " - Percentuais federal/estadual/municipal"
If mCrt <> 3 Then cboFormaTotTrib.AddItem TOT_SIMPLES & " - Alíquota do Simples Nacional"
If nForma = TOT_AUTO Or (nForma = TOT_SIMPLES And mCrt = 3) Then nForma = IIf(mCrt = 3, TOT_PERCENTUAIS, TOT_SIMPLES)
SelecionarCodigo cboFormaTotTrib, CStr(nForma)
AjustarTotTrib
End Sub

Private Sub AjustarTotTrib()
Dim nForma As Integer
nForma = Val(cboFormaTotTrib.Text)
txtTotTribSN.Enabled = (mCrt <> 4 And nForma = TOT_SIMPLES)
lblTotSN.Enabled = txtTotTribSN.Enabled
txtTotTribFed.Enabled = (mCrt <> 4 And (nForma = TOT_VALORES Or nForma = TOT_PERCENTUAIS))
txtTotTribEst.Enabled = txtTotTribFed.Enabled
txtTotTribMun.Enabled = txtTotTribFed.Enabled
lblTotFed.Enabled = txtTotTribFed.Enabled
lblTotEst.Enabled = txtTotTribFed.Enabled
lblTotMun.Enabled = txtTotTribFed.Enabled
End Sub

Private Sub cboFormaTotTrib_Click()
AjustarTotTrib
End Sub

Private Sub chkIBSCBS_Click()
If mCarregando Then Exit Sub
If chkIBSCBS.Value = vbChecked And CodigoDoCombo(cboNBS) = "" Then
   MsgBox "Para informar o IBS/CBS escolha primeiro o NBS do serviço.", vbExclamation, "NFS-e"
   chkIBSCBS.Value = vbUnchecked
   Exit Sub
End If
cboIndOp.Enabled = (chkIBSCBS.Value = vbChecked)
cboClassTrib.Enabled = cboIndOp.Enabled
lblIndOp.Enabled = cboIndOp.Enabled
lblClassTrib.Enabled = cboIndOp.Enabled
lblCST.Enabled = cboIndOp.Enabled
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
Dim sNac As String, sMun As String, sNBS As String, sIndOp As String, sClass As String, sCST As String
Dim sAliq As String, sSN As String, sFed As String, sEst As String, sMunTrib As String
Dim nForma As Integer

sNac = CodigoDoCombo(cboCodTribNac)
sMun = SoDigitos(txtCodTribMun.Text)
sNBS = CodigoDoCombo(cboNBS)
sIndOp = CodigoDoCombo(cboIndOp)
sClass = CodigoDoCombo(cboClassTrib)
sCST = Left$(sClass, 3)
If Len(sNac) <> 6 Then
   MsgBox "Escolha o serviço padrão (código de tributação nacional) na lista.", vbExclamation, "NFS-e"
   cboCodTribNac.SetFocus
   Exit Sub
End If
If sMun <> "" And Len(sMun) <> 3 Then
   MsgBox "O código de tributação municipal tem 3 dígitos (deixe em branco se o município não usa).", vbExclamation, "NFS-e"
   txtCodTribMun.SetFocus
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
If Not Pct(txtTotTribEst, "Tributos estaduais", 99, False, sEst) Then Exit Sub
If Not Pct(txtTotTribMun, "Tributos municipais", 99, False, sMunTrib) Then Exit Sub
nForma = Val(cboFormaTotTrib.Text)
If Not mTemFormaTotTrib And mCrt <> 4 And (nForma <> IIf(mCrt = 3, TOT_PERCENTUAIS, TOT_SIMPLES) Or CDbl(txtTotTribEst.Text) <> 0) Then
   MsgBox "Para escolher a forma dos tributos aproximados e o % estadual, atualize antes o banco de dados " & _
          "(script 181 - AlterTable_Empresa_NFSeTotTrib.sql).", vbExclamation, "NFS-e"
   Exit Sub
End If
If chkIBSCBS.Value = vbChecked Then
   If sNBS = "" Or Len(sIndOp) <> 6 Or Len(sClass) <> 6 Then
      MsgBox "Para informar IBS/CBS escolha o NBS, o indicador da operação e a classificação (cClassTrib).", vbExclamation, "NFS-e"
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
   IIf(mTemFormaTotTrib, "NFSeTotTribEst = " & sEst & ", NFSeFormaTotTrib = " & nForma & ", ", "") & _
   "NFSeMensagem = " & IIf(Trim$(txtMensagem.Text) = "", "NULL", NFSe_Q(Trim$(txtMensagem.Text))) & ", " & _
   "NFSeIBSCBS = " & IIf(chkIBSCBS.Value = vbChecked, "1", "0") & ", " & _
   "NFSeIndOp = " & IIf(sIndOp = "", "NULL", NFSe_Q(sIndOp)) & ", " & _
   "NFSeCST = " & IIf(sCST = "", "NULL", NFSe_Q(sCST)) & ", " & _
   "NFSeClassTrib = " & IIf(sClass = "", "NULL", NFSe_Q(sClass)) & ", " & _
   "NFSeAssinaturaSHA256 = " & IIf(chkSHA256.Value = vbChecked, "1", "0") & ";"
Unload Me
End Sub

'Consulta o convenio do municipio da empresa no Sefin (ambiente gravado; salve antes se mudou o ambiente)
Private Sub cmdConvenio_Click()
Dim oRes As Object
Dim sStatus As String
sStatus = NFSe_Executar("acao=convenio", oRes)
If Left$(sStatus, 2) <> "OK" Then
   MsgBox "Não foi possível consultar o município:" & vbCrLf & vbCrLf & Mid$(sStatus, 7), vbExclamation, "NFS-e"
   Exit Sub
End If
MsgBox Replace(NFSe_Valor(oRes, "mensagem"), " | ", vbCrLf), _
       IIf(Val(NFSe_Valor(oRes, "situacao")) = 200, vbInformation, vbExclamation), "Convênio do município - NFS-e"
End Sub

Private Sub cmdVoltar_Click()
Unload Me
End Sub
