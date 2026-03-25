VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Recibos_Avulso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RECIBO"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8220
   Icon            =   "Recibos_Avulso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      ScaleHeight     =   645
      ScaleWidth      =   8025
      TabIndex        =   11
      Top             =   60
      Width           =   8055
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "RECIBO AVULSO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   1260
         TabIndex        =   12
         Top             =   180
         Width           =   2580
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   300
         Picture         =   "Recibos_Avulso.frx":23D2
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2235
      Left            =   60
      ScaleHeight     =   2205
      ScaleWidth      =   8025
      TabIndex        =   9
      Top             =   840
      Width           =   8055
      Begin ChamaleonBtn.chameleonButton cmdCal1 
         Height          =   315
         Left            =   960
         TabIndex        =   22
         Tag             =   "Calendario"
         Top             =   1620
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         BTYPE           =   8
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Recibos_Avulso.frx":8742
         PICN            =   "Recibos_Avulso.frx":875E
         PICH            =   "Recibos_Avulso.frx":AAB1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ComboBox cboForma 
         Height          =   315
         Left            =   5700
         TabIndex        =   2
         Top             =   960
         Width           =   2235
      End
      Begin VB.TextBox txtSobrenome 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4860
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1680
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox cboFuncionario 
         Height          =   315
         Left            =   60
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtCodFunc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   60
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.ComboBox cboReferente 
         Height          =   315
         Left            =   60
         TabIndex        =   1
         Top             =   960
         Width           =   5595
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   1620
         Width           =   1395
      End
      Begin MSMask.MaskEdBox mskData 
         Height          =   315
         Left            =   60
         TabIndex        =   3
         Top             =   1620
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtCliente 
         Height          =   315
         Left            =   2400
         MaxLength       =   80
         TabIndex        =   0
         Top             =   360
         Width           =   5535
      End
      Begin VB.Label lblFormaPgto 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de Pgto"
         Height          =   195
         Left            =   5745
         TabIndex        =   21
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Funcionário"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   19
         Top             =   120
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1380
         TabIndex        =   15
         Top             =   1380
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   14
         Top             =   1380
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Referente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   13
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Recebemos de:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2460
         TabIndex        =   10
         Top             =   120
         Width           =   1335
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   16
      Top             =   3525
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10160
            Text            =   "Online.Info Sistemas"
            TextSave        =   "Online.Info Sistemas"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "23:23"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ChamaleonBtn.chameleonButton cmdImprimir 
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Imprimir"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Recibos_Avulso.frx":CE04
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdPDF 
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Criar PDF"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Recibos_Avulso.frx":CE20
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdEmail 
      Height          =   315
      Left            =   2580
      TabIndex        =   7
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Envio E-Mail"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Recibos_Avulso.frx":CE3C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdWhatsapp 
      Height          =   315
      Left            =   3840
      TabIndex        =   8
      Top             =   3120
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Envio Whatsapp"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Recibos_Avulso.frx":CE58
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image imLogoCupom 
      Height          =   1125
      Left            =   5460
      Picture         =   "Recibos_Avulso.frx":CE74
      Top             =   3120
      Visible         =   0   'False
      Width           =   2850
   End
End
Attribute VB_Name = "Recibos_Avulso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moCombo As cComboHelper
Dim sSQL As String
Dim r As ADODB.Recordset
Dim IMPRIMIR As Boolean
Dim var_ImpTermica As String
Dim vImpressoraNormal As String
Dim varTipoRecPgto As String
Dim varTipoRecHaver As String
Dim vCidadeUF As String
Private Sub Imprimir_Folha()
varImpPDF = False
'abrindo arquivo .ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"

'nome da maquina
'vImpressoraNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")

'nome da maquina
If varImpPDF = True Then
    vImpressoraNormal = "Impressora PDF"
Else
    vImpressoraNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")
End If

Dim Prt As Printer
Dim oldPrinter As String

'Armazena o nome da impressora atual
oldPrinter = Printer.DeviceName

' Find and use the printer just selected in the ListBox
For Each Prt In Printers
   If Prt.DeviceName = vImpressoraNormal Then
      Set Printer = Prt
      Exit For
   End If
Next
   
   Me.Hide
   'Principal_Impressao.Hide
   
   With REL_Recibo_Avulso
      .txtCliente.Caption = UCase(txtCliente.Text)
      .txtValor.Caption = UCase(NumeroExtenso(txtValor.Text, True))
      .txthead.Caption = "R$ " & Format(txtValor.Text, "##,##0.00")
      .txtProveniente.Caption = cboReferente
      .txtData.Caption = "" & vCidadeUF & ", " & Day(mskData) & " de " & MonthName(Month(mskData)) & " de " & Year(mskData)
      .txtUsuario.Caption = Format(txtCodFunc, "00")
      .txtAssinatura.Caption = UCase(cboFuncionario.Text) & " " & UCase(txtSobrenome.Text)
      .txtFormaPgto.Caption = UCase(cboForma.Text)
      .txtDataPgto.Caption = Format(Date, "dd/mm/yy") & " às " & Format(Now, "hh:mm") & " hs."
      .Relatorio.NumeroRegistros = 1
      .Relatorio.NomeImpressora = vImpressoraNormal
      .Relatorio.Ativar
   End With
   Unload REL_Recibo_Avulso
   
   Me.Show 1
End Sub

Private Sub cboReferente_GotFocus()
   cboReferente.Clear
   cboReferente.AddItem "COMPRA DE MATERIAL"
   cboReferente.AddItem "PAGAMENTO DE CONTA"
   moCombo.AttachTo cboReferente
End Sub

Private Sub cboReferente_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub cmdCal1_Click()
Dim varData As Variant
Dim fCal As Calendario

varData = Empty                    'Inicializa a variável

Set fCal = New Calendario      'Cria o form de calendário
fCal.Show vbModal

varData = fCal.DateSelected    'Recupera a data selecionada

Unload fCal                           'Fecha o form
Set fCal = Nothing                   'Destrói a variável

If Not IsDate(varData) Then Exit Sub   'Valida a data
If varData = 0 Then Exit Sub

mskData = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdPDF_Click()
varImpPDF = True
'abrindo arquivo .ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"

'nome da maquina
'vImpressoraNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")

'nome da maquina
If varImpPDF = True Then
    vImpressoraNormal = "Impressora PDF"
Else
    vImpressoraNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")
End If

Dim Prt As Printer
Dim oldPrinter As String

'Armazena o nome da impressora atual
oldPrinter = Printer.DeviceName

' Find and use the printer just selected in the ListBox
For Each Prt In Printers
   If Prt.DeviceName = vImpressoraNormal Then
      Set Printer = Prt
      Exit For
   End If
Next
   
   Me.Hide
   'Principal_Impressao.Hide
   
   With REL_Recibo_Avulso
      .txtCliente.Caption = UCase(txtCliente.Text)
      .txtValor.Caption = UCase(NumeroExtenso(txtValor.Text, True))
      .txthead.Caption = "R$ " & Format(txtValor.Text, "##,##0.00")
      .txtProveniente.Caption = cboReferente
      .txtData.Caption = "" & vCidadeUF & ", " & Day(mskData) & " de " & MonthName(Month(mskData)) & " de " & Year(mskData)
      .txtUsuario.Caption = Format(txtCodFunc, "00")
      .txtAssinatura.Caption = UCase(cboFuncionario.Text) & " " & UCase(txtSobrenome.Text)
      .txtFormaPgto.Caption = UCase(cboForma.Text)
      .txtDataPgto.Caption = Format(Date, "dd/mm/yy") & " às " & Format(Now, "hh:mm") & " hs."
      .Relatorio.NumeroRegistros = 1
      .Relatorio.NomeImpressora = vImpressoraNormal
      .Relatorio.Ativar
   End With
   Unload REL_Recibo_Avulso
   
   Me.Show 1
End Sub

Private Sub txtCodFunc_Change()
If txtCodFunc.Text = "" Then Exit Sub

'If cmdAlterar.Enabled = True Then
sSQL = "SELECT codigo, nome, sobrenome FROM funcionario WHERE (codigo = " & txtCodFunc.Text & ");"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then
    cboFuncionario.Text = r("nome")
    txtSobrenome.Text = r("sobrenome")
Else
    cboFuncionario.Text = ""
    txtSobrenome.Text = ""
End If
If r.State <> 0 Then r.Close
Set r = Nothing
'End If
End Sub
Private Sub cboFuncionario_LostFocus()
On Error GoTo TrataErro

If cboFuncionario.Text = "" Then txtCodFunc.Text = "": Exit Sub

'If cmdAlterar.Enabled = False Then
   If cboFuncionario.ListIndex = -1 Then
      'txtCodFunc.Text = ""
      'Exit Sub
   End If
'End If

txtCodFunc = cboFuncionario.ItemData(cboFuncionario.ListIndex)

sSQL = "SELECT sobrenome FROM funcionario WHERE codigo = " & txtCodFunc.Text & ""
Set r = dbData.OpenRecordset(sSQL)

If r.BOF Then
    txtSobrenome.Text = ""
Else
    txtSobrenome.Text = ValidateNull(r("sobrenome"))
End If

Exit Sub

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub
Private Sub Fonte(Tamanho As Byte, Negrito As Boolean, Italico As Boolean) 'Altera a fonte
   Printer.FontSize = Tamanho
   Printer.FontBold = Negrito
   Printer.FontItalic = Italico
End Sub

Private Sub cmdImprimir_Click()
If ShowMsg("Deseja imprimir o recibo ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
   IMPRIMIR = True
Else
   IMPRIMIR = False
End If

If IMPRIMIR = True Then
    If varTipoRecPgto = "CUPOM" Then
        Imprimir_ReciboCupom
    Else
        Imprimir_Folha
    End If
End If
End Sub

Private Sub Form_Load()
sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
Set r = dbData.OpenRecordset(sSQL)

vCidadeUF = r("cidade") & "-" & r("estado")

If r.State <> 0 Then r.Close
Set r = Nothing

'logomarca impressa do cupom
Dim sLogo As String
Set oCfg = sysConfig("LOGO_CUPOM")
sLogo = oCfg.Value
Set oCfg = Nothing
If Dir$(sLogo) <> "" Then Set imLogoCupom.Picture = LoadPicture(sLogo)

'tipo de recibo de pagamento
Set oCfg = sysConfig("TIPORECPGTO")
varTipoRecPgto = oCfg.Value
Set oCfg = Nothing

Set moCombo = New cComboHelper
mskData.Text = Format(Date, "dd/mm/yy")
StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
End Sub

Private Sub cboForma_GotFocus()
cboForma.Clear
cboForma.AddItem UCase("DINHEIRO")
cboForma.AddItem UCase("CHEQUE")
cboForma.AddItem UCase("CARTÃO DE CRÉDITO")
cboForma.AddItem UCase("CARTÃO DE DÉBITO")
cboForma.AddItem UCase("DEPÓSITO")
cboForma.AddItem UCase("TRANSFERÊNCIA")
cboForma.AddItem UCase("PIX")

If cboForma.ListCount <> 0 Then cboForma.ListIndex = 0

moCombo.AttachTo cboForma
End Sub
Private Sub cboFuncionario_GotFocus()
Dim varNomeAntes As String
Dim varCodAntes As String

varNomeAntes = cboFuncionario.Text
varCodAntes = txtCodFunc.Text

cboFuncionario.Clear

sSQL = "SELECT DISTINCT nome, codigo, sobrenome FROM funcionario ORDER BY nome;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboFuncionario.AddItem r("nome")
   cboFuncionario.ItemData(cboFuncionario.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

cboFuncionario.Text = varNomeAntes
txtCodFunc.Text = varCodAntes

cboFuncionario.SelStart = 0
cboFuncionario.SelLength = Len(cboFuncionario)

moCombo.AttachTo cboFuncionario
End Sub

Private Sub cboFuncionario_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Imprimir_ReciboCupom()
'pegar o nome da impressora no ini
 Dim oIni As Ini
 
 Set oIni = New Ini
 oIni.Arquivo = appPathApp & "config.ini"
 var_ImpTermica = oIni.LerTexto("IMPRESSORA_TERMICA", "impressora")
 Set oIni = Nothing
 
 Dim Prt As Printer
 Dim oldPrinter As String
 
 'Armazena o nome da impressora atual
 oldPrinter = Printer.DeviceName
 
 ' Find and use the printer just selected in the ListBox
 For Each Prt In Printers
    If Prt.DeviceName = var_ImpTermica Then
       Set Printer = Prt
       Exit For
    End If
 Next
 
Dim sSQL As String
Dim rEmpresa As ADODB.Recordset
Dim i As Integer
Dim f As Integer
  
'tabela empresa
sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
Set rEmpresa = dbData.OpenRecordset(sSQL)
vCidadeUF = rEmpresa("cidade") & "-" & rEmpresa("estado")

 f = FreeFile()
    
With Printer
      .ScaleMode = vbPixels
      .PaintPicture imLogoCupom.Picture, 100, 0, 372, 150
      
      For i = 1 To 6
         Printer.Print " "
      Next
      
      .ScaleMode = vbCentimeters
      .FontName = "courier new"

      Fonte 10, True, False
      Printer.Print Tab((35 - Len(rEmpresa("fantasia"))) / 2); rEmpresa("fantasia")   'Esse /2 é p/ centralizar
      Fonte 10, False, False
      Printer.Print Tab((35 - Len(rEmpresa("razao"))) / 2); rEmpresa("razao")
      Fonte 8, False, False
      Printer.Print rEmpresa("endereco") & ", " & rEmpresa("cidade") & "-" & rEmpresa("estado")
      Printer.Print "FONE: "; rEmpresa("telefone")                                        '& " - (89) 9986-3739"
      Fonte 8, False, False
      Printer.Print "CNPJ:"; rEmpresa("cnpj") & "  IE:" & rEmpresa("ie")
      Fonte 8, False, False
      Printer.Print String(40, "-")
      
       For i = 1 To 2
         Printer.Print " "
      Next
      
      Fonte 10, True, False
      Printer.Print Tab(3); "R E C I B O   A V U L S O"
      
      
      For i = 1 To 1
         Printer.Print " "
      Next
               
      Fonte 8, True, False
      Printer.Print Tab(28); "R$ " & Format(txtValor.Text, "##,##0.00")
      
      For i = 1 To 1
         Printer.Print " "
      Next
      
      Dim Line1 As String
      Dim Line2 As String
      
      Dim Texto As String
      Texto = UCase(NumeroExtenso(txtValor.Text, True))

      Line1 = Mid(Texto, 1, 40)
      Line2 = Mid(Texto, 41, 80)
     
      Fonte 8, False, False
      Printer.Print Tab(2); "Recebi(emos) de: "
      Fonte 8, True, False
      Printer.Print Tab(2); txtCliente.Text
      
      For i = 1 To 1
         Printer.Print " "
      Next
      
      Fonte 8, False, False
      Printer.Print Tab(2); "A importância supra de: "
      Fonte 8, False, False
      Printer.Print Tab(2); Line1
      Printer.Print Tab(2); Line2
      Fonte 8, False, False
     
      For i = 1 To 1
         Printer.Print " "
      Next
      
      Printer.Print Tab(2); "Referente à: " & UCase(cboReferente.Text)
      
      For i = 1 To 1
         Printer.Print " "
      Next
      
      For i = 1 To 2
         Printer.Print " "
      Next
     
      Fonte 8, False, False
      Printer.Print Tab(10); "" & vCidadeUF & ", " & Day(mskData) & " de " & MonthName(Month(mskData)) & " de " & Year(mskData)
      
      For i = 1 To 3
            Printer.Print " "
      Next
      
      Printer.Print Tab((40 - Len("______________________________________")) / 2); "______________________________________"
      Printer.Print Tab((40 - Len("Assinatura")) / 2); "Assinatura"
      

     
   Close #f
   .EndDoc
End With
 
Tratar_Erro:
 ' Atribui a impressora inicial
 'For Each Prt In Printers
 '   If Prt.DeviceName = oldPrinter Then
 '      Set Printer = Prt
 '      Exit For
 '   End If
 'Next
 
 If Not rEmpresa Is Nothing Then If rEmpresa.State <> 0 Then rEmpresa.Close
 'If Err.Number = 52 Then
  '  ShowMsg "Impressora não esta pronta ou está com problemas, Verifique !!!", vbInformation
  '  Printer.KillDoc
  '  Exit Sub
 'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub mskData_GotFocus()
   SelectControl mskData
End Sub

Private Sub mskData_KeyPress(KeyAscii As Integer)
   mskData.Mask = "##/##/##"
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtValor_LostFocus()
txtValor.Text = Format(txtValor.Text, ocMONEY)
End Sub
