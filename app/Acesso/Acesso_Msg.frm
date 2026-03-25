VERSION 5.00
Begin VB.Form Acesso_msg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aguarde o Processamento"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4260
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   4260
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3720
      Top             =   720
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "Acesso_Msg.frx":0000
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Carregando permissões de acesso..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   3150
   End
End
Attribute VB_Name = "Acesso_msg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Call Conectar
Unload Me
Call Acesso_Definir.ListarAcesso
Call Acesso_Definir.ListarUsuario
End Sub
