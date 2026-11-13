VERSION 5.00
Begin VB.Form frmNFSe_Cancelar
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelar NFS-e"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboMotivo
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   540
      Width           =   3015
   End
   Begin VB.TextBox txtJustificativa
      Height          =   1335
      Left            =   120
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1260
      Width           =   6960
   End
   Begin VB.CommandButton cmdConfirmar
      Caption         =   "&Cancelar a NFS-e"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   2820
      Width           =   1815
   End
   Begin VB.CommandButton cmdVoltar
      Cancel          =   -1  'True
      Caption         =   "&Voltar"
      Height          =   375
      Left            =   5385
      TabIndex        =   7
      Top             =   2820
      Width           =   1695
   End
   Begin VB.Label lblNota
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   6960
   End
   Begin VB.Label lblMotivo
      Caption         =   "Motivo:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   585
      Width           =   855
   End
   Begin VB.Label lblJust
      Caption         =   "Justificativa (de 15 a 255 caracteres):"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1020
      Width           =   4335
   End
   Begin VB.Label lblCont
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   3255
   End
End
Attribute VB_Name = "frmNFSe_Cancelar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Motivo (1, 2 ou 9) e justificativa do evento de cancelamento (e101101). Quem chama lê Confirmado/Motivo/Justificativa.
Public Confirmado As Boolean
Public Motivo As String
Public Justificativa As String

Private Sub Form_Load()
cboMotivo.AddItem "1 - Erro na emissão"
cboMotivo.AddItem "2 - Serviço não prestado"
cboMotivo.AddItem "9 - Outros"
cboMotivo.ListIndex = 0
txtJustificativa_Change
End Sub

Private Sub txtJustificativa_Change()
lblCont.Caption = Len(Trim$(txtJustificativa.Text)) & " caracteres (mínimo 15)"
cmdConfirmar.Enabled = (Len(Trim$(txtJustificativa.Text)) >= 15)
End Sub

Private Sub cmdConfirmar_Click()
If Len(Trim$(txtJustificativa.Text)) < 15 Then Exit Sub
If MsgBox("Confirma o CANCELAMENTO desta NFS-e no Sefin Nacional?" & vbCrLf & vbCrLf & "Essa operação não pode ser desfeita.", _
          vbQuestion + vbYesNo + vbDefaultButton2, "Cancelar NFS-e") <> vbYes Then Exit Sub
Motivo = Left$(cboMotivo.Text, 1)
Justificativa = Trim$(txtJustificativa.Text)
Confirmado = True
Me.Hide
End Sub

Private Sub cmdVoltar_Click()
Confirmado = False
Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
   Cancel = True
   cmdVoltar_Click
End If
End Sub
