VERSION 5.00
Begin VB.Form LiberarVenda 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Autorização"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   3795
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   60
      TabIndex        =   3
      Top             =   660
      Width           =   3675
      Begin VB.ComboBox cboGerente 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   420
         Width           =   3375
      End
      Begin VB.TextBox txtSenha 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gerente:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   180
         Width           =   615
      End
      Begin VB.Label lblTexto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Senha:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   900
         Width           =   510
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   2340
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2340
      Width           =   1455
   End
   Begin VB.Label lblTexto 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Você precisará da permissão do gerente para continuar essa operação."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   2
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   3645
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "LiberarVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moCombo As New cComboHelper
Private mCancelled As Boolean
Private mGerente As Long
Dim sSQL As String
Dim r As ADODB.Recordset

Public Property Get Cancelled() As Boolean
   Cancelled = mCancelled
End Property

Public Property Get Gerente() As Long
   Gerente = mGerente
End Property

Sub pListarGerentes()
'Dim sSQL As String
'Dim r As ADODB.Recordset

sSQL = "SELECT codigo, login FROM usuario WHERE (nivel = 1) ORDER BY login;"
Set r = dbData.OpenRecordset(sSQL)

cboGerente.Clear
Do While Not r.EOF
   cboGerente.AddItem r("login")
   cboGerente.ItemData(cboGerente.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub cboGerente_GotFocus()
moCombo.AttachTo cboGerente
End Sub

Private Sub cboGerente_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If (KeyAscii = 13) Then SendKey ocKEYTAB
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If Not ExistInList(cboGerente) Or Trim(cboGerente) = "" Then
   ShowMsg "Selecione o gerente da lista.", vbExclamation
   Exit Sub
End If

mGerente = cboGerente.ItemData(cboGerente.ListIndex)
'MsgBox cboGerente.ItemData(cboGerente.ListIndex)

sSQL = "SELECT * FROM usuario WHERE (codigo = " & mGerente & ") AND (password = '" & txtSenha & "');"
Set r = dbData.OpenRecordset(sSQL)

If r.BOF Then
   ShowMsg "Gerente não encontrato ou senha inválida.", vbExclamation
   Exit Sub
End If

mCancelled = False
Unload Me
End Sub

Private Sub Form_Load()
Set Icon = Nothing
CenterForm Me, Width, Height
mCancelled = True
mGerente = -1
pListarGerentes
txtSenha.PasswordChar = Chr$(42)
End Sub

Private Sub txtSenha_GotFocus()
   SelectControl txtSenha
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then SendKey ocKEYTAB
   If (KeyAscii = 39) Then KeyAscii = 0
End Sub
