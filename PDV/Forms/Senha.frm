VERSION 5.00
Begin VB.Form Senha 
   BorderStyle     =   0  'None
   Caption         =   "Seja Bem Vindo!"
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   Icon            =   "Senha.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Senha.frx":030A
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
      TabIndex        =   1
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário:"
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
      Left            =   1680
      TabIndex        =   3
      Top             =   2100
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
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
      Left            =   1800
      TabIndex        =   2
      Top             =   2580
      Width           =   615
   End
End
Attribute VB_Name = "Senha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moCombo As cComboHelper

'Lista os usuarios permitidos para execução
Sub pListarUsuarios()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT codigo, login FROM usuario WHERE (VISIVEL = 1) ORDER BY login"
   Set r = dbData.OpenRecordset(sSQL)
   
   cboUsuario.Clear
   
   Do While Not r.EOF
      cboUsuario.AddItem r("login")
      cboUsuario.ItemData(cboUsuario.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub Entrar()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim codUsuario As Long, nomeUsuario As String
   Dim nivelAcesso As Integer
   
   Static Tag As Byte
   
   If cboUsuario.ListIndex = -1 Then Exit Sub
   If cboUsuario.Text = "" Then Exit Sub
   
   'Recupera o código e o login do usuário
   codUsuario = cboUsuario.ItemData(cboUsuario.ListIndex)
   nomeUsuario = cboUsuario.Text
   
   'Consulta os dados do usuário
   sSQL = "SELECT login, senha, nivel FROM usuario WHERE (codigo = " & codUsuario & ")"
   Set r = dbData.OpenRecordset(sSQL)
   
   If r.BOF And r.RecordCount = 0 Then
      ShowMsg "USUÁRIO INEXISTENTE!" & vbCrLf & "O usuário escolhido não existe." & vbCrLf & "Verifique e escolha novamente.", vbInformation
      cboUsuario.ListIndex = -1
      Exit Sub
   End If
   
   If r("senha") <> txtSenha Then
      ShowMsg "SENHA ERRADA!" & vbCrLf & "Verifique sua senha e tente novamente", vbInformation
      txtSenha = ""
      Tag = Tag + 1
      
      If Tag = 3 Then
         ShowMsg "Nº DE TENTATIVAS ESGOTADO!" & vbCrLf & "Entre em  contato como o administrador" & vbCrLf & "O sistema será fechado.", vbInformation
         EncerrarPrograma
         Exit Sub
      End If
      
      txtSenha.Text = ""
      txtSenha.SetFocus
      Exit Sub
   End If
   
   'Recupera o nível de acesso do usuário
   nivelAcesso = ValidateNull(r("nivel"))
   
   'Fecha a tabela
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   'Reseta as variáveis
   Tag = 0
   cboUsuario.Text = ""
   txtSenha.Text = ""
   
   'Carrega o form principal
   Load Tela_Principal
   Tela_Principal.txtCodFuncionario.Text = codUsuario
   Tela_Principal.StatusBar1.Panels(2).Text = nomeUsuario
   Tela_Principal.txtNivel.Text = nivelAcesso
   Tela_Principal.Show
   
   'Descarrega o form de login
   Unload Me
End Sub

Private Sub cboUsuario_GotFocus()
   moCombo.AttachTo cboUsuario
End Sub

Private Sub cboUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 27 Then
      EncerrarPrograma
   ElseIf KeyCode = 13 Then
      Entrar
   End If
End Sub

Private Sub cboUsuario_Validate(Cancel As Boolean)
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim codUsuario As Long
   
   If cboUsuario.ListIndex = -1 Then Exit Sub
   If cboUsuario.Text = "" Then Exit Sub
   
   sSQL = "SELECT senha FROM usuario WHERE (codigo = " & codUsuario & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   If IsNull(r("senha")) Then
      Dim fMod As Senha_Alterar
      Set fMod = New Senha_Alterar
      
      Load fMod
      fMod.CodigoFuncionario = cboUsuario.ItemData(cboUsuario.ListIndex)
      fMod.Show vbModal
      
      Unload fMod
      Set fMod = Nothing
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 27 Then
      EncerrarPrograma
   End If
End Sub

Private Sub Form_Load()
   CenterForm Me, Width, Height
   Set moCombo = New cComboHelper
   pListarUsuarios
End Sub

Private Sub txtSenha_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 27 Then
      EncerrarPrograma
   ElseIf KeyCode = 13 Then
      Call Entrar
   End If
End Sub
