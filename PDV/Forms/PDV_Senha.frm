VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form PDV_Senha 
   BorderStyle     =   0  'None
   Caption         =   "Seja Bem Vindo!"
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   Icon            =   "PDV_Senha.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "PDV_Senha.frx":030A
   ScaleHeight     =   3750
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin VB.TextBox txtSenha 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2460
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2520
      Width           =   1635
   End
   Begin VB.ComboBox cboUsuario 
      Height          =   315
      Left            =   2460
      TabIndex        =   0
      Top             =   2040
      Width           =   1635
   End
   Begin MSMask.MaskEdBox mskCPF 
      Height          =   315
      Left            =   2460
      TabIndex        =   1
      Top             =   2040
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      MaxLength       =   14
      Mask            =   "###.###.###-##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblVersao 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   4110
      TabIndex        =   6
      Top             =   3495
      Width           =   270
   End
   Begin VB.Label lblCopyright 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   60
      TabIndex        =   5
      Top             =   3495
      Width           =   270
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário:"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1575
      TabIndex        =   4
      Top             =   2100
      Width           =   825
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1695
      TabIndex        =   3
      Top             =   2580
      Width           =   675
   End
End
Attribute VB_Name = "PDV_Senha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moCombo As cComboHelper
Dim oCfg As ConfigItem


Private Sub Entrar()
Dim sSQL As String
Dim r As ADODB.Recordset
Static Tag As Byte
Set oCfg = sysConfig("TIPOLOGIN")

If oCfg.Value = "NOME" Then
   If cboUsuario.Text = "" Then Exit Sub
   
   If cboUsuario.List(cboUsuario.ListIndex) <> Empty Then
      sSQL = "SELECT * FROM usuario WHERE (codigo = " & cboUsuario.ItemData(cboUsuario.ListIndex) & ");"
      Set r = dbData.OpenRecordset(sSQL)
      
      If r.BOF Then
         ShowMsg "USUÁRIO INEXISTENTE!" & vbCrLf & "O usuário escolhido não existe." & vbCrLf & "Verifique e escolha novamente.", vbInformation
         cboUsuario.Text = ""
         Exit Sub
      End If
      
      If txtSenha = r("password") Then
         Tag = 0
         PDV_Senha.Hide
         vCodFunc = PDV_Senha.cboUsuario.ItemData(cboUsuario.ListIndex)
         PDV.txtCodFuncAP.Text = PDV_Senha.cboUsuario.ItemData(cboUsuario.ListIndex)
         PDV.txtCodFunc.Text = PDV_Senha.cboUsuario.ItemData(cboUsuario.ListIndex)
         PDV.StatusBar1.Panels(3).Text = PDV_Senha.cboUsuario
         PDV.txtFuncAP.Text = PDV_Senha.cboUsuario
         PDV.txtNivel.Text = ValidateNull(r("nivel"))
         PDV.Show
         PDV_Senha.cboUsuario.Text = ""
         PDV_Senha.txtSenha.Text = ""
      Else
         ShowMsg "SENHA ERRADA!" & vbCrLf & "Verifique sua senha e tente novamente", vbInformation
         txtSenha = ""
         Tag = Tag + 1
         If Tag = 3 Then
            ShowMsg "Nº DE TENTATIVAS ESGOTADO!" & vbCrLf & "Entre em  contato como o administrador" & vbCrLf & "O sistema será fechado.", vbInformation
            End
         End If
         txtSenha.Text = ""
         txtSenha.SetFocus
      End If
    End If
Else
      sSQL = "SELECT * FROM usuario WHERE (CPF = '" & mskCPF.Text & "');"
      Set r = dbData.OpenRecordset(sSQL)
      
      If r.BOF Then
         ShowMsg "USUÁRIO INEXISTENTE!" & vbCrLf & "O usuário escolhido não existe." & vbCrLf & "Verifique e escolha novamente.", vbInformation
         mskCPF.Mask = ""
         mskCPF.Text = ""
         mskCPF.Mask = "###.###.###-##"
         Exit Sub
      End If
      
      If txtSenha = r("password") Then
         Tag = 0
         PDV_Senha.Hide
         vCodFunc = r("codigo")
         PDV.txtCodFuncAP.Text = r("codigo")
         PDV.StatusBar1.Panels(3).Text = r("login")
         PDV.txtFuncAP.Text = r("login")
         PDV.txtNivel.Text = ValidateNull(r("nivel"))
         PDV.Show
         PDV_Senha.mskCPF.Mask = ""
         PDV_Senha.mskCPF.Text = ""
         PDV_Senha.mskCPF.Mask = "###.###.###-##"
         PDV_Senha.txtSenha.Text = ""
      Else
         ShowMsg "SENHA ERRADA!" & vbCrLf & "Verifique sua senha e tente novamente", vbInformation
         txtSenha = ""
         Tag = Tag + 1
         If Tag = 3 Then
            ShowMsg "Nº DE TENTATIVAS ESGOTADO!" & vbCrLf & "Entre em  contato como o administrador" & vbCrLf & "O sistema será fechado.", vbInformation
            End
         End If
         txtSenha.Text = ""
         txtSenha.SetFocus
      End If
End If
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub cboUsuario_GotFocus()
   moCombo.AttachTo cboUsuario
End Sub

Private Sub cboUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 27 Then
      End
   ElseIf KeyCode = 13 Then
      Call Entrar
   End If
End Sub

Private Sub cboUsuario_LostFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If cboUsuario.Text = "" Then Exit Sub
   
   If cboUsuario.List(Me.cboUsuario.ListIndex) <> Empty Then
      sSQL = "SELECT * FROM funcionario WHERE (codigo = " & cboUsuario.ItemData(cboUsuario.ListIndex) & ");"
      Set r = dbData.OpenRecordset(sSQL)
      
      'If IsNull(r("password")) Then
      '   Senha_Alterar.txtCodFuncionario.Text = cboUsuario.ItemData(cboUsuario.ListIndex)
      '   Senha_Alterar.Show 1
      'End If
      
      If r.State <> 0 Then r.Close
      Set r = Nothing
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set moCombo = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   End
End If
End Sub

Private Sub Form_Load()
Dim sSQL As String
Dim r As ADODB.Recordset
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

lblVersao.Caption = "Versão: " & Trim(Str(App.Major)) & "." & Trim(Str(App.Minor)) & "." & Trim(Str(App.Revision))
lblCopyright.Caption = "Online.Info Sistemas (1998 - " & Year(Now) & ")"

Set oCfg = sysConfig("TIPOLOGIN")

If App.PrevInstance = True Then
   ShowMsg "O sistema já encontra-se em execução nesta máquina!", vbInformation
   End
Else
    If oCfg.Value = "NOME" Then
        cboUsuario.Visible = True
        mskCPF.Visible = False
        Label5.Caption = "Usuário:"
        sSQL = "SELECT codigo, login FROM usuario WHERE (visivel = 1) ORDER BY login;"
        Set r = dbData.OpenRecordset(sSQL)
        
        Do While Not r.EOF
           cboUsuario.AddItem r("login")
           cboUsuario.ItemData(cboUsuario.NewIndex) = r("codigo")
           r.MoveNext
        Loop
        
        If r.State <> 0 Then r.Close
        Set r = Nothing
    Else
        cboUsuario.Visible = False
        mskCPF.Visible = True
        Label5.Caption = "CPF:"
        If mskCPF.Visible = True Then mskCPF.SetFocus
    End If
End If


Set oCfg = Nothing
Set moCombo = New cComboHelper
End Sub

Private Sub mskCPF_GotFocus()
mskCPF.SelStart = 0
mskCPF.SelLength = Len(mskCPF.Text)
End Sub


Private Sub txtSenha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   KillProcess "PDV"
ElseIf KeyCode = 13 Then
   Call Entrar
End If
End Sub
