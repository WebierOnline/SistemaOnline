VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Recados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Recados"
   ClientHeight    =   7380
   ClientLeft      =   360
   ClientTop       =   1410
   ClientWidth     =   11835
   Icon            =   "Recados.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Novo"
      Height          =   750
      Left            =   2400
      Picture         =   "Recados.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Novo Cadastro"
      Top             =   6585
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "S&alvar"
      Height          =   735
      Left            =   60
      Picture         =   "Recados.frx":1594
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Salvar"
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "A&pagar"
      Height          =   750
      Left            =   1620
      Picture         =   "Recados.frx":26CE
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Apagar Cadastro"
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Alterar"
      Height          =   750
      Left            =   840
      Picture         =   "Recados.frx":29D8
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Alterar Cadastro"
      Top             =   6600
      Width           =   735
   End
   Begin VB.Frame Frame3 
      Caption         =   "Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4140
      TabIndex        =   22
      Top             =   6600
      Width           =   5895
      Begin VB.OptionButton optResponder 
         Caption         =   "À R&esponder"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   1275
      End
      Begin VB.OptionButton optRespondido 
         Caption         =   "Resp&ondido"
         Height          =   195
         Left            =   1440
         TabIndex        =   29
         Top             =   360
         Width           =   1155
      End
      Begin VB.CommandButton Command7 
         Caption         =   "<<"
         Height          =   285
         Left            =   4560
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   300
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Caption         =   ">>"
         Height          =   285
         Left            =   5400
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   300
         Width           =   375
      End
      Begin VB.CommandButton Command12 
         Caption         =   "ok"
         Height          =   285
         Left            =   3720
         TabIndex        =   23
         Top             =   300
         Width           =   375
      End
      Begin MSMask.MaskEdBox MskDataNova 
         Height          =   285
         Left            =   2760
         TabIndex        =   24
         Top             =   300
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Line Line1 
         X1              =   4680
         X2              =   5640
         Y1              =   420
         Y2              =   420
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cadastro de Recados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   60
      TabIndex        =   13
      Top             =   600
      Width           =   11715
      Begin VB.ComboBox cboTipoTelefone 
         Height          =   315
         Left            =   5640
         TabIndex        =   2
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox cboStatus 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   9960
         TabIndex        =   7
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox cboPara 
         Height          =   315
         Left            =   3240
         TabIndex        =   1
         Top             =   600
         Width           =   2355
      End
      Begin VB.TextBox txtDe 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   3075
      End
      Begin VB.TextBox txtAssunto 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   9765
      End
      Begin MSMask.MaskEdBox mskHora 
         Height          =   315
         Left            =   10500
         TabIndex        =   5
         Top             =   600
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskData 
         Height          =   315
         Left            =   9000
         TabIndex        =   4
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFone 
         Height          =   315
         Left            =   7380
         TabIndex        =   3
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label35 
         Caption         =   "Tipo do Telefone"
         Height          =   255
         Left            =   5640
         TabIndex        =   28
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Height          =   195
         Left            =   9960
         TabIndex        =   27
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone:"
         Height          =   195
         Left            =   7380
         TabIndex        =   20
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Para:"
         Height          =   195
         Left            =   3240
         TabIndex        =   19
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         Height          =   195
         Left            =   9000
         TabIndex        =   17
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora:"
         Height          =   195
         Left            =   10500
         TabIndex        =   16
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assunto:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblCodigo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
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
         Left            =   11520
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   120
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Sair"
      Height          =   750
      Left            =   11040
      Picture         =   "Recados.frx":32A2
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Sair"
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "DATA  HORA  DE   TELEFONE   PARA   ASSUNTO"
      Height          =   195
      Left            =   120
      TabIndex        =   31
      Top             =   2700
      Width           =   3765
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RECADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   45
      TabIndex        =   21
      Top             =   60
      Width           =   11730
   End
End
Attribute VB_Name = "Recados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moCombo As cComboHelper

Private Sub Atualiza_Dados()
   'If lblCodigo.Caption <> "" Then dYn("CODIGO") = lblCodigo.Caption
   'If txtDe.Text <> "" Then dYn("DE") = txtDe.Text
   'If cboPara.Text <> "" Then dYn("PARA") = cboPara.Text
   'If mskFone.Text <> "" Then dYn("TELEFONE") = mskFone.Text
   'If mskData.Text <> "" Then dYn("DATA") = mskData.Text
   'If mskHora.Text <> "" Then dYn("Hora") = mskHora.Text
   'If txtAssunto.Text <> "" Then dYn("ASSUNTO") = txtAssunto.Text
   'If cboStatus.Text <> "" Then dYn("Status") = cboStatus.Text
   'If cboTipoTelefone.Text <> "" Then dYn("TIPO_TELEFONE") = cboTipoTelefone.Text
End Sub

Private Function Atualizar_Dados2() As Boolean
   'If lblCodigo.Caption <> "" Then TBRecados("CODIGO") = lblCodigo.Caption
   'If txtDe.Text <> "" Then TBRecados("DE") = txtDe.Text
   'If cboPara.Text <> "" Then TBRecados("PARA") = cboPara.Text
   'If mskFone.Text <> "" Then TBRecados("TELEFONE") = mskFone.Text
   'If mskData.Text <> "" Then TBRecados("DATA") = mskData.Text
   'If mskHora.Text <> "" Then TBRecados("Hora") = mskHora.Text
   'If txtAssunto.Text <> "" Then TBRecados("ASSUNTO") = txtAssunto.Text
   'If cboStatus.Text <> "" Then TBRecados("Status") = cboStatus.Text
   'If cboTipoTelefone.Text <> "" Then TBRecados("TIPO_TELEFONE") = cboTipoTelefone.Text
End Function

Private Sub AutoNumeracao()
   Dim lNum As Long
   
   'If TBRecados.RecordCount = 0 Then
   '   lblCodigo.Caption = "1"
   'Else
   '   TBRecados.MoveLast
   '   lNum = TBRecados!Codigo + 1
   '   lblCodigo.Caption = lNum
   'End If
End Sub

Private Sub Campos_Brancos2()
   lblCodigo.Caption = 0
   txtDe.Text = ""
   cboPara.Text = ""
   mskFone.Mask = ""
   mskFone.Text = ""
   mskData.Mask = ""
   mskData.Text = ""
   mskHora.Mask = ""
   mskHora.Text = ""
   txtAssunto.Text = ""
   cboStatus.Text = ""
   cboTipoTelefone.Text = ""
End Sub

Private Sub Mostrar_Dados2()
   'lblCodigo.Caption = TBRecados!Codigo
   'txtDe.Text = TBRecados!DE
   'cboPara.Text = TBRecados!PARA
   'mskFone.Text = TBRecados!TELEFONE
   'mskData.Text = TBRecados!Data
   'mskHora.Text = TBRecados!Hora
   'txtAssunto.Text = TBRecados!ASSUNTO
   'cboStatus.Text = TBRecados!Status
   'cboTipoTelefone.Text = TBRecados!tipo_telefone
End Sub

Private Sub cboPara_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboPara.Clear
     
   sSQL = "SELECT nome FROM funcionario;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboPara.AddItem r("nome")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboPara
   SendKeys "{F4}"
End Sub

Private Sub cboStatus_GotFocus()
   cboStatus.Clear
   cboStatus.AddItem "À Responder"
   cboStatus.AddItem "Respondido"
   
   moCombo.AttachTo cboStatus
   
   If cboStatus.Text = "" Then
      cboStatus.ListIndex = 0
   End If
   
   SendKeys "{F4}"
End Sub

Private Sub cboTipoTelefone_GotFocus()
   cboTipoTelefone.Clear
   cboTipoTelefone.AddItem "Nenhum"
   cboTipoTelefone.AddItem "Casa"
   cboTipoTelefone.AddItem "Trabalho"
   cboTipoTelefone.AddItem "Vizinho"
   cboTipoTelefone.AddItem "Orelhão"
   cboTipoTelefone.AddItem "Celular"
   
   moCombo.AttachTo cboTipoTelefone
   SendKeys "{F4}"
End Sub

Private Sub cboTipoTelefone_LostFocus()
   If cboTipoTelefone.Text = "" Then cboTipoTelefone.Text = "Nenhum"
End Sub

Private Sub Command1_Click()
   Campos_Brancos2
   Form_Load
   AutoNumeracao
   txtDe.SetFocus
End Sub

Private Sub Command12_Click()
   If MskDataNova.Text = "" Then Exit Sub
   If MskDataNova.Text = "__/__/__" Then Exit Sub
   
   If optResponder.Value = True Then
      'ABRIR_BD_com_Data Me.Data2
      'Data2.RecordSource = "SELECT * FROM RECADOS WHERE DATA = #" & Format(MskDataNova, "mm/dd/yy") & "# and STATUS = 'À Responder' order by DATA, HORA"
      'Data2.Refresh
   
   ElseIf optRespondido.Value = True Then
      'ABRIR_BD_com_Data Me.Data2
      'Data2.RecordSource = "SELECT * FROM RECADOS WHERE DATA = #" & Format(MskDataNova, "mm/dd/yy") & "# and STATUS = 'Respondido' order by DATA, HORA"
      'Data2.Refresh
   End If
End Sub

Private Sub Command2_Click()
   'On Error GoTo TrataErro
   
   If txtDe.Text = "" Or cboPara.Text = "" Or txtAssunto.Text = "" Or cboStatus.Text = "" Then
      ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Verifique se todos os dados estão cadastrados.", vbInformation
      Exit Sub
   Else
      If Not Atualizar_Dados2 Then
         
         Exit Sub
      End If
      
      Campos_Brancos2
      Form_Load
   End If
   
   txtDe.SetFocus
   Exit Sub
   
'TrataErro:
   'If Err.Number = 3022 Then
   '   MsgBox "DADOS DUPLICADO!" & vbCrLf & "Verifique se este recado já está cadastrado.", vbInformation, "Aviso do Sistema"
   '   Exit Sub
   'End If
   '
   'If Err.Number = 3421 Then
   '   MsgBox "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Verifique se todos os dados estão nos campos.", vbInformation, "Aviso do Sistema"
   '   Exit Sub
   '   txtDe.SetFocus
   'End If
End Sub

Private Sub Command3_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtDe.Text = "" Or cboPara.Text = "" Then
      ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Verifique se todos os dados estão nos campos.", vbInformation
      Exit Sub
   End If
   
   If MsgBox("Tem certeza que deseja excluir este registro?", vbInformation + vbYesNo, "Aviso do Sistema") = vbNo Then
      txtDe.SetFocus
      Exit Sub
   End If
   
   sSQL = "SELECT * FROM recados WHERE (codigo = " & lblCodigo.Caption & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then
      'dYn.Delete
      ShowMsg "Operação realizada com sucesso.", vbExclamation
   End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   Campos_Brancos2
   Form_Load
   txtDe.SetFocus
End Sub

Private Sub Command4_Click()
   'If TBRecados.EOF Then Exit Sub
   'TBRecados.MoveNext
   'If TBRecados.EOF Then Exit Sub
   'Mostrar_Dados2
End Sub

Private Sub Command5_Click()
   On Error GoTo TrataErro
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtDe.Text = "" Or cboPara.Text = "" Then
      ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Verifique se todos os dados estão nos campos.", vbInformation
      Exit Sub
   End If
   
   sSQL = "SELECT * FROM recados WHERE (codigo = " & lblCodigo.Caption & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not Atualizar_Dados2 Then
      
      Exit Sub
   End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   ShowMsg "Operação concretizada com sucesso", vbInformation
   Campos_Brancos2
   Form_Load
   txtDe.SetFocus
   Exit Sub
   
TrataErro:
   If Err.Number = 3421 Then
      ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Verifique se todos os dados estão nos campos.", vbInformation
      Exit Sub
   End If
End Sub

Private Sub Command6_Click()
   Unload Me
End Sub

Private Sub Command7_Click()
   'If TBRecados.BOF Then Exit Sub
   'TBRecados.MovePrevious
   'If TBRecados.BOF Then Exit Sub
   'Mostrar_Dados2
End Sub

Private Sub DBGrid2_DblClick()
'   DBGrid2.Col = 0
'   mskData.Text = DBGrid2.Text
'   DBGrid2.Col = 1
'   mskHora.Text = DBGrid2.Text
'   DBGrid2.Col = 2
'   txtDe.Text = DBGrid2.Text
'   DBGrid2.Col = 3
'   mskFone.Text = DBGrid2.Text
'   DBGrid2.Col = 4
'   cboPara.Text = DBGrid2.Text
'   DBGrid2.Col = 5
'   cboStatus.Text = DBGrid2.Text
'   DBGrid2.Col = 6
'   lblCodigo.Caption = DBGrid2.Text
'   DBGrid2.Col = 7
'   cboTipoTelefone.Text = DBGrid2.Text
'   DBGrid2.Col = 8
'   txtAssunto.Text = DBGrid2.Text
'   txtDe.SetFocus
End Sub

Private Sub DBGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   'If Data2.Recordset.RecordCount Then DBGrid2.SelBookmarks.Add Data2.Recordset.Bookmark
End Sub

Private Sub Form_Load()
     
   'Set TBRecados = BD.OpenRecordset("RECADOS", dbOpenTable)
   'TBRecados.Index = ("index_codigo")
   
   mskData.Text = Format(Date, "dd/mm/yy")
   
   'ABRIR_BD_com_Data Me.Data2
   'Data2.RecordSource = "SELECT * FROM RECADOS WHERE STATUS = 'À Responder' order by DATA, HORA"
   'Data2.Refresh
   
   Set moCombo = New cComboHelper
End Sub

Function Maiuscula(KeyAscii As Integer)
   If KeyAscii > 96 And KeyAscii < 123 Then
      KeyAscii = KeyAscii - 32
   End If
   Maiuscula = KeyAscii
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub mskData_KeyPress(KeyAscii As Integer)
   mskData.Mask = "##/##/##"
End Sub

Private Sub mskData_LostFocus()
   If Not IsDate(mskData) Then
      ShowMsg "DATA INVÁLIDA" & vbCrLf & "Digite a data novamente!", vbInformation
      mskData.SetFocus
      mskData.SelStart = 0
      mskData.SelLength = Len(mskData)
      Exit Sub
      End If
End Sub

Private Sub MskDataNova_GotFocus()
   If optResponder.Value = False And optRespondido.Value = False Then
      ShowMsg "Escolha uma das opções: À Responder ou Respondido", vbExclamation
      optResponder.SetFocus
   End If
End Sub

Private Sub MskDataNova_LostFocus()
   If MskDataNova.Text = "__/__/__" Then
      txtDe.SetFocus
      Exit Sub
   End If
   
   If Not IsDate(MskDataNova) Then
      ShowMsg "DATA INVÁLIDA" & vbCrLf & "Digite a data novamente!", vbInformation
      MskDataNova.SetFocus
      MskDataNova.SelStart = 0
      MskDataNova.SelLength = Len(MskDataNova)
      Exit Sub
   End If
End Sub

Private Sub mskFone_GotFocus()
   If cboTipoTelefone.Text = "Nenhum" Then
      mskFone.Mask = ""
      mskFone.Text = ""
   End If
   
   mskFone.SelStart = 4
End Sub

Private Sub mskFone_KeyPress(KeyAscii As Integer)
   If cboTipoTelefone.Text = "Nenhum" Then
      mskFone.Mask = ""
      mskFone.Text = ""
   ElseIf cboTipoTelefone.Text = "Casa" Then
      mskFone.Mask = "(0xx##) ###-####"
   ElseIf cboTipoTelefone.Text = "Trabalho" Then
      mskFone.Mask = "(0xx##) ###-####"
   ElseIf cboTipoTelefone.Text = "Vizinho" Then
      mskFone.Mask = "(0xx##) ###-####"
   ElseIf cboTipoTelefone.Text = "Orelhão" Then
      mskFone.Mask = "(0xx##) ###-####"
   ElseIf cboTipoTelefone.Text = "Celular" Then
      mskFone.Mask = "(0xx##) ####-####"
   Else
      ShowMsg "Escolha uma das opções", vbExclamation
      cboTipoTelefone.SetFocus
   End If
End Sub

Private Sub mskHora_GotFocus()
   SelectControl mskHora
End Sub

Private Sub mskHora_KeyPress(KeyAscii As Integer)
   mskHora.Mask = "##:##"
End Sub

Private Sub optResponder_Click()
   MskDataNova.SetFocus
End Sub

Private Sub optRespondido_Click()
   MskDataNova.SetFocus
End Sub

Private Sub txtAssunto_KeyPress(KeyAscii As Integer)
   KeyAscii = Maiuscula(KeyAscii)
End Sub

Private Sub txtDe_KeyPress(KeyAscii As Integer)
   If lblCodigo.Caption = 0 Then
      ShowMsg "Antes de inserir um novo registro, clique no botão NOVO", vbExclamation
      txtDe.Text = ""
   End If
   
   KeyAscii = Maiuscula(KeyAscii)
End Sub
