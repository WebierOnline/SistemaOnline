VERSION 5.00
Begin VB.Form frmIBPTProgresso
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Atualizando Tabela IBPT"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5745
   ControlBox      =   0   'False
   BeginProperty Font
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrPulse
      Enabled         =   0   'False
      Interval        =   90
      Left            =   5280
      Top             =   840
   End
   Begin VB.PictureBox picBarraFundo
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   180
      ScaleHeight     =   270
      ScaleWidth      =   5205
      TabIndex        =   1
      Top             =   840
      Width           =   5265
      Begin VB.PictureBox picBarra
         Appearance      =   0  'Flat
         BackColor       =   &H00C08000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   0
         ScaleHeight     =   270
         ScaleWidth      =   120
         TabIndex        =   2
         Top             =   0
         Width           =   120
      End
   End
   Begin VB.Label lblMsg
      Caption         =   "Aguarde..."
      Height          =   615
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   5385
   End
End
Attribute VB_Name = "frmIBPTProgresso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mPulseX   As Single
Private mPulseDir As Integer

' Mostra a forminha e comeca a animacao "indeterminada" (usada durante o download,
' cujo GET e bloqueante - a barra volta a andar nos DoEvents em volta).
Public Sub Iniciar(ByVal sMsg As String)
    lblMsg.Caption = sMsg
    picBarra.Left = 0
    picBarra.Width = 0
    mPulseX = 0
    mPulseDir = 1
    tmrPulse.Enabled = True
    Me.Show
    Me.Refresh
    DoEvents
End Sub

Public Sub DefinirMensagem(ByVal sMsg As String)
    lblMsg.Caption = sMsg
    Me.Refresh
    DoEvents
End Sub

' Troca a animacao indeterminada por barra proporcional (usada na importacao,
' que tem DoEvents a cada 100 registros).
Public Sub DefinirProgresso(ByVal nAtual As Long, ByVal nTotal As Long)
    tmrPulse.Enabled = False
    Dim p As Single
    p = 0
    If nTotal > 0 Then p = nAtual / nTotal
    If p < 0 Then p = 0
    If p > 1 Then p = 1
    picBarra.Left = 0
    picBarra.Width = picBarraFundo.ScaleWidth * p
    Me.Refresh
    DoEvents
End Sub

Private Sub Form_Load()
    mPulseDir = 1
End Sub

Private Sub tmrPulse_Timer()
    Dim bloco As Single
    bloco = picBarraFundo.ScaleWidth * 0.28
    picBarra.Width = bloco
    mPulseX = mPulseX + mPulseDir * (picBarraFundo.ScaleWidth * 0.07)
    If mPulseX + bloco >= picBarraFundo.ScaleWidth Then
        mPulseX = picBarraFundo.ScaleWidth - bloco
        mPulseDir = -1
    End If
    If mPulseX <= 0 Then
        mPulseX = 0
        mPulseDir = 1
    End If
    picBarra.Left = mPulseX
End Sub
