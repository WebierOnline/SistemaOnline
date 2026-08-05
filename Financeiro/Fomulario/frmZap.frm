VERSION 5.00
Begin VB.Form frmZap 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BUSCA WHATSAPP WEB"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5565
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmZap.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdZap 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      Height          =   435
      Left            =   4440
      MaskColor       =   &H00000000&
      Picture         =   "frmZap.frx":1084A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1485
      Width           =   450
   End
   Begin VB.TextBox txtCompara 
      Appearance      =   0  'Flat
      Height          =   420
      Left            =   765
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2070
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.TextBox txtZap 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   420
      Left            =   765
      TabIndex        =   0
      Text            =   "Digite o número do whatsapp"
      Top             =   1470
      Width           =   3630
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      X1              =   780
      X2              =   4395
      Y1              =   1890
      Y2              =   1890
   End
   Begin VB.Label lblZap 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Digite o número do whatsapp"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   765
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   2325
   End
End
Attribute VB_Name = "frmZap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Chama Whatsapp Web
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long
Private Const conSwNormal = 1

Private Sub cmdZap_Click()
'Se campo estiver vázio, aparece uma mensagem
If txtZap = "" Or txtZap = "Digite o número do whatsapp" Then
   MsgBox "Nenhum número de whatsapp foi localizado!", vbExclamation, "Busca whatsapp web"
   Exit Sub
End If
'Se quantidade de caracteres for inferior a 15, ele bloqueia a pesquisa
If Len(txtZap.TEXT) < 15 Then
   MsgBox "Número de telefone inválido!", vbCritical, "Busca whatsapp web"
   txtZap = ""
   Exit Sub
End If
'Apaga algum valor na textbox txtCompara
txtCompara = ""
'Trás o mesmo valor da textbox txtZap
txtCompara = txtZap
'Utiliza o Replace para retirar os pontos na textbox txtCompara.
txtCompara = Replace(Replace(Replace(txtCompara, "(", ""), ")", ""), "-", "")
'Chama a função ShellExecute = url da api do whatsapp web
ShellExecute hwnd, "open", ("https://api.whatsapp.com/send?phone=55" & txtCompara), _
vbNullString, vbNullString, conSwNormal
End Sub

Private Sub txtZap_Change()
If txtZap = "" Then txtCompara = ""
If txtZap = "Digite o número do whatsapp" Then
   txtZap.ForeColor = &H404040
Else
   txtZap.ForeColor = &HFFFFFF
End If
End Sub

Private Sub txtZap_GotFocus()
'If txtZap = "Digite o número do whatsapp" Then
'   With txtZap
'   .SelStart = 0
'   .SelLength = Len(txtZap.TEXT)
'   End With
'End If
End Sub

Private Sub txtZap_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox KeyCode
'Exit Sub
'Se apertar a tecla enter formata o campo celular
If KeyCode = 13 Then
   'Formata o campo celular
   If Len(txtZap.TEXT) = 14 Or Len(txtZap.TEXT) = 12 Then
      txtZap.TEXT = FormataTelefone(txtZap.TEXT)
      If Len(txtZap.TEXT) = 15 Then
         cmdZap_Click
      End If
   Else
      If Len(txtZap.TEXT) = 15 Then Exit Sub
         txtZap.TEXT = ""
   End If
End If
End Sub

Private Sub txtZap_KeyPress(KeyAscii As Integer)
If txtZap = "Digite o número do whatsapp" Then
   txtZap = ""
   lblZap.Visible = True
   txtZap.ForeColor = &HFFFFFF
End If
'Não permite espaço
If KeyAscii = 32 Then KeyAscii = 0
'Apenas Números
ApenasNrs KeyAscii
'Organiza modelo para celular
CampoCelular txtZap, KeyAscii
'Limita a Qde de caracteres
txtZap.MaxLength = 15
End Sub

Private Sub txtZap_LostFocus()
'Formata o campo celular
If Len(txtZap.TEXT) = 14 Or Len(txtZap.TEXT) = 12 Then
   txtZap.TEXT = FormataTelefone(txtZap.TEXT)
End If
End Sub
