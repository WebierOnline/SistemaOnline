VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Recibo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RECIBO"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   Icon            =   "Recibo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      ScaleHeight     =   645
      ScaleWidth      =   7425
      TabIndex        =   12
      Top             =   60
      Width           =   7455
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "RECIBO"
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
         Left            =   30
         TabIndex        =   13
         Top             =   120
         Width           =   4800
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   900
         Picture         =   "Recibo.frx":23D2
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1995
      Left            =   60
      ScaleHeight     =   1965
      ScaleWidth      =   7425
      TabIndex        =   8
      Top             =   840
      Width           =   7455
      Begin VB.TextBox txtDataPgto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4380
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1500
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtSobrenome 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4380
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1200
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txtValor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4380
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   900
         Width           =   1095
      End
      Begin VB.TextBox txtFormaPgto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   900
         Width           =   1215
      End
      Begin VB.ComboBox cboFuncionario 
         Height          =   315
         Left            =   60
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   300
         Width           =   2475
      End
      Begin VB.TextBox txtCodFunc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.ComboBox cboParcela 
         Height          =   315
         Left            =   2220
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   900
         Width           =   2115
      End
      Begin VB.ComboBox cboPedido 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   900
         Width           =   2115
      End
      Begin VB.ComboBox cboCliente 
         Height          =   315
         Left            =   2580
         TabIndex        =   1
         Top             =   300
         Width           =   4755
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
         TabIndex        =   16
         Top             =   60
         Width           =   990
      End
      Begin VB.Label lblParcela 
         AutoSize        =   -1  'True
         Caption         =   "Parcela"
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
         Left            =   2220
         TabIndex        =   11
         Top             =   660
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pedido | Data"
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
         TabIndex        =   10
         Top             =   660
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Left            =   2580
         TabIndex        =   9
         Top             =   60
         Width           =   600
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdImprimir 
      Height          =   315
      Left            =   60
      TabIndex        =   4
      Top             =   2880
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
      MICON           =   "Recibo.frx":8742
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   14
      Top             =   3225
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9075
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
   Begin ChamaleonBtn.chameleonButton cmdPDF 
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   2880
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
      MICON           =   "Recibo.frx":875E
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
      TabIndex        =   6
      Top             =   2880
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
      MICON           =   "Recibo.frx":877A
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
      TabIndex        =   7
      Top             =   2880
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
      MICON           =   "Recibo.frx":8796
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "Recibo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moCombo As cComboHelper
Dim sSQL As String
Dim r As ADODB.Recordset
Dim vCidadeUF As String
Private Sub cboCliente_Click()
If cboCliente.Text = "" Then Exit Sub

sSQL = "SELECT codigo FROM cliente WHERE (nome = '" & cboCliente.Text & "');"
Set r = dbData.OpenRecordset(sSQL)

sSQL = "SELECT cod_pedido, data_compra FROM pedidos WHERE (cod_cliente = " & r("codigo") & ") ORDER BY data_compra DESC;"
If r.State <> 0 Then r.Close
Set r = Nothing

Set r = dbData.OpenRecordset(sSQL)

cboPedido.Clear

If r.BOF Then
   cboPedido.AddItem "NENHUM PEDIDO"
Else
   Do While Not r.EOF
      cboPedido.AddItem Format(r("cod_pedido"), "00000") & " -> " & Format(r("data_compra"), "dd/mm/yy")
      r.MoveNext
   Loop
End If

If r.State <> 0 Then r.Close
Set r = Nothing

cboPedido.ListIndex = 0
End Sub

Private Sub CboCliente_GotFocus()
sSQL = "SELECT codigo, nome FROM cliente ORDER BY nome;"
Set r = dbData.OpenRecordset(sSQL)

cboCliente.Clear
Do While Not r.EOF
   cboCliente.AddItem r("nome")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing
moCombo.AttachTo cboCliente
End Sub

Private Sub CboCliente_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then cboCliente_Click
End Sub

Private Sub CboCliente_LostFocus()
   cboCliente_Click
End Sub

Private Sub cboParcela_Click()
   'On Error Resume Next
   'cboValor.ListIndex = cboParcela.ListIndex
End Sub

Private Sub cboPedido_Click()
cboParcela.Clear

Dim Pedido As Long
'Dim sSQL As String
'Dim r As ADODB.Recordset

If cboPedido.Text = "NENHUM PEDIDO" Then
   Pedido = 0
Else
   Pedido = Mid(cboPedido.Text, 1, InStr(1, cboPedido.Text, "->") - 1)
End If

sSQL = "SELECT *, parcelas.PAGAMENTO as var_DataPgto FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
   "WHERE (status = 1) AND (pedidos.cod_pedido = " & Pedido & ");"

Set r = dbData.OpenRecordset(sSQL)

If r.BOF Then
   cboParcela.AddItem "NENHUMA PARCELA ENCONTRADA"
   txtValor.Text = ""
   txtFormaPgto.Text = ""
   txtDataPgto.Text = ""
Else
   Do Until r.EOF
      cboParcela.AddItem r("numero")
      txtValor.Text = r("VALOR_FINAL")
      txtFormaPgto.Text = r("FORMA_PGTO")
      txtDataPgto.Text = Format(r("var_DataPgto"), "dd/mm/yy")
      r.MoveNext
   Loop
End If

If r.State <> 0 Then r.Close
Set r = Nothing

cboParcela.ListIndex = 0
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

Private Sub cboValor_Change()

End Sub

Private Sub cmdImprimir_Click()
If cboParcela.Text = "" Or cboParcela.Text = "NENHUMA PARCELA ENCONTRADA" Then
   ShowMsg "Nenhuma parcela para ser impressa !!", vbInformation
   Exit Sub
End If

If cboPedido.Text = "NENHUM PEDIDO" Then
   ShowMsg "ESSE PEDIDO NÃO PODE SER IMPRESSO", vbExclamation
   Exit Sub
End If

varImpPDF = False

'abrindo arquivo .ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"

'nome da maquina
'var_ImpNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")

'nome da maquina
If varImpPDF = True Then
    var_ImpNormal = "Impressora PDF"
Else
    var_ImpNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")
End If

Dim Prt As Printer
Dim oldPrinter As String

'Armazena o nome da impressora atual
oldPrinter = Printer.DeviceName

' Find and use the printer just selected in the ListBox
For Each Prt In Printers
   If Prt.DeviceName = var_ImpNormal Then
      Set Printer = Prt
      Exit For
   End If
Next
   
Me.Hide
'Principal_Impressao.Hide

With REL_Recibo
   .txtCliente.Caption = UCase(cboCliente.Text)
   .txtValor.Caption = UCase(NumeroExtenso(txtValor.Text, True))
   .txthead.Caption = "R$ " & Format(txtValor.Text, "##,##0.00")
   .txtProveniente.Caption = "Pagamento da " & cboParcela.Text & "ª parcela do pedido Nº " & Format(Mid(cboPedido.Text, 1, InStr(1, cboPedido.Text, "->") - 1), "00000")
   .txtData.Caption = " " & vCidadeUF & ", " & Day(txtDataPgto) & " de " & MonthName(Month(txtDataPgto)) & " de " & Year(txtDataPgto)
   .txtUsuario.Caption = Format(txtCodFunc, "00")
   .txtAssinatura.Caption = UCase(cboFuncionario.Text) & " " & UCase(txtSobrenome.Text)
   .txtFormaPgto.Caption = UCase(txtFormaPgto.Text)
   .txtDataPgto.Caption = Format(Date, "dd/mm/yy") & " às " & Format(Now, "hh:mm") & " hs."
   .Relatorio.NumeroRegistros = 1
   .Relatorio.NomeImpressora = var_ImpNormal
   .Relatorio.Ativar
End With
Unload REL_Recibo

Me.Show 1
End Sub

Private Sub cmdPDF_Click()
If cboParcela.Text = "" Or cboParcela.Text = "NENHUMA PARCELA ENCONTRADA" Then
   ShowMsg "Nenhuma parcela para ser impressa !!", vbInformation
   Exit Sub
End If

If cboPedido.Text = "NENHUM PEDIDO" Then
   ShowMsg "ESSE PEDIDO NÃO PODE SER IMPRESSO", vbExclamation
   Exit Sub
End If

varImpPDF = True

'abrindo arquivo .ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"

'nome da maquina
'var_ImpNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")

'nome da maquina
If varImpPDF = True Then
    var_ImpNormal = "Impressora PDF"
Else
    var_ImpNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")
End If

Dim Prt As Printer
Dim oldPrinter As String

'Armazena o nome da impressora atual
oldPrinter = Printer.DeviceName

' Find and use the printer just selected in the ListBox
For Each Prt In Printers
   If Prt.DeviceName = var_ImpNormal Then
      Set Printer = Prt
      Exit For
   End If
Next
   
Me.Hide
'Principal_Impressao.Hide

With REL_Recibo
   .txtCliente.Caption = UCase(cboCliente.Text)
   .txtValor.Caption = UCase(NumeroExtenso(txtValor.Text, True))
   .txthead.Caption = "R$ " & Format(txtValor.Text, "##,##0.00")
   .txtProveniente.Caption = "Pagamento da " & cboParcela.Text & "ª parcela do pedido Nº " & Format(Mid(cboPedido.Text, 1, InStr(1, cboPedido.Text, "->") - 1), "00000")
   .txtData.Caption = " " & vCidadeUF & ", " & Day(txtDataPgto) & " de " & MonthName(Month(txtDataPgto)) & " de " & Year(txtDataPgto)
   .txtUsuario.Caption = Format(txtCodFunc, "00")
   .txtAssinatura.Caption = UCase(cboFuncionario.Text) & " " & UCase(txtSobrenome.Text)
   .txtFormaPgto.Caption = UCase(txtFormaPgto.Text)
   .txtDataPgto.Caption = Format(Date, "dd/mm/yy") & " às " & Format(Now, "hh:mm") & " hs."
   .Relatorio.NumeroRegistros = 1
   .Relatorio.NomeImpressora = var_ImpNormal
   .Relatorio.Ativar
End With
Unload REL_Recibo

Me.Show 1
End Sub


Private Sub Form_Load()
sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
Set r = dbData.OpenRecordset(sSQL)

vCidadeUF = r("cidade") & "-" & r("estado")

If r.State <> 0 Then r.Close
Set r = Nothing

Set moCombo = New cComboHelper
StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
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
Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub
