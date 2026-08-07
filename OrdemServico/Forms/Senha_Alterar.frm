VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Begin VB.Form Senha_Alterar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SENHA"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   Icon            =   "Senha_Alterar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   120
      TabIndex        =   3
      Top             =   540
      Width           =   3435
      Begin VB.TextBox txtSenha2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1140
         Width           =   1455
      End
      Begin VB.TextBox txtSenha1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   540
         Width           =   1455
      End
      Begin ChamaleonBtn.chameleonButton cmdSalvarSenha 
         Height          =   315
         Left            =   1860
         TabIndex        =   2
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Salvar"
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
         MICON           =   "Senha_Alterar.frx":23D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblSenha2 
         Caption         =   "Confirmar"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   900
         Width           =   975
      End
      Begin VB.Label lblSenha1 
         Caption         =   "Senha"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   975
      End
      Begin VB.Line Line1 
         X1              =   1740
         X2              =   1740
         Y1              =   240
         Y2              =   1500
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ALTERAÇÃO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3435
   End
End
Attribute VB_Name = "Senha_Alterar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CodigoFuncionario As Long    'Variável publica para evitar o uso de controles escondidos

Private Sub cmdSalvarSenha_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim bRet As Boolean
   
   If txtSenha1.Text = "" Or txtSenha2.Text = "" Or CodigoFuncionario = 0 Then Exit Sub
   
   If txtSenha1.Text <> txtSenha2.Text Then
      ShowMsg "As senhas digitas são diferentes!", vbInformation
      txtSenha1.Text = ""
      txtSenha2.Text = ""
      txtSenha1.SetFocus
      Exit Sub
   End If
   
   'Atualiza a senha do funcionário
   sSQL = "UPDATE funcionario SET senha = '" & txtSenha1.Text & "' WHERE (codigo = " & CodigoFuncionario & ");"
   'Executa a atualização e verifica o resultado
   bRet = dbData.Execute(sSQL)
   
   If (Not bRet) Then
      ShowMsg "Não foi possível alterar a senha.", vbCritical
      Exit Sub
   End If
   
   'Verifica se a senha foi cadastrada e exibe os dados
   'sSQL = "SELECT * FROM funcionario WHERE (codigo = " & CodigoFuncionario & ");"
   'Set r = dbData.OpenRecordset(sSQL)
   
   'If Not IsNull(r("senha")) Then
      Unload Me
   'End If
End Sub

Private Sub Form_Load()
   CenterForm Me, Width, Height
   CodigoFuncionario = 0
End Sub
