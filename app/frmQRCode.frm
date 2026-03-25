VERSION 5.00
Begin VB.Form frmQRCode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "API WhatsApp EkklesiaSoft"
   ClientHeight    =   6360
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   9870
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4389.785
   ScaleMode       =   0  'User
   ScaleWidth      =   9268.44
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picQRCode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5250
      Left            =   120
      Picture         =   "frmQRCode.frx":0000
      ScaleHeight     =   118.266
      ScaleMode       =   0  'User
      ScaleWidth      =   122.004
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   540
      Width           =   5415
   End
   Begin VB.TextBox txtText1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5640
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmQRCode.frx":A9A4
      Top             =   4440
      Width           =   4095
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   60
      Top             =   60
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONECTADO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3660
      TabIndex        =   3
      Top             =   5820
      Width           =   2565
   End
   Begin VB.Image Image1 
      Height          =   3840
      Left            =   5640
      Picture         =   "frmQRCode.frx":A9C9
      Stretch         =   -1  'True
      Top             =   540
      Width           =   4080
   End
   Begin VB.Label lblLabel1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LEIA O QRCODE PARA CONECTAR AO WHASTAPP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1380
      TabIndex        =   0
      Top             =   120
      Width           =   7125
   End
End
Attribute VB_Name = "frmQRCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* Sistema...: frmQRCode
'* Empresa...: EkklesiaSoft Tecnologia em Sistemas
'* Módulo....: frmQRCode
'* Função....: Formulário que carrega QRCode para conexão há API WhatsApp
'* CopyRight.: (C)2025 EkklesiaSoft Tecnologia em Sistemas
'* Criação...: EkklesiaSoft Tecnologia em Sistemas
'* Data......: 07/12/2025 02:22:46
'* * * * * * *

Option Explicit
DefInt A-Z

Public Sub CarregarImagemBase64(ByVal base64String As String)
    Dim stream As Object
    Dim arquivoTemp As String, processando As String
    Dim img As Object
    
    On Error GoTo deuErro
    
    ' Caminho temporário para salvar a imagem (mesmo que PNG ou JPG)
    arquivoTemp = App.path & "\temp_img.jpg" ' ou .png
    
    ' Decodifica base64 e salva como arquivo
    'MsgBox "Set stream = CreateObject('ADODB.Stream')", vbCritical + vbOKOnly, "CarregarImagemBase64"
    processando = "CreateObject('ADODB.Stream')"
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' Binary
    stream.Open
    processando = "DecodeBase64(base64String)"
    stream.Write DecodeBase64(base64String)
    stream.SaveToFile arquivoTemp, 2 ' Overwrite
    stream.Close
    Set stream = Nothing

    ' Usa WIA para carregar a imagem moderna
    'MsgBox "Set img = CreateObject('WIA.ImageFile')", vbCritical + vbOKOnly, "CarregarImagemBase64"
    processando = "CreateObject('WIA.ImageFile')"
    Set img = CreateObject("WIA.ImageFile")
    img.LoadFile arquivoTemp

    ' Exibe a imagem em um controle Image (não PictureBox)
    'MsgBox "picQRCode.Picture", vbCritical + vbOKOnly, "CarregarImagemBase64"
    processando = "img.FileData.Picture"
    picQRCode.Picture = img.FileData.Picture

    Set img = Nothing
    
    Exit Sub
    
deuErro:
    MsgBox Err.Description + vbNewLine + processando, vbCritical + vbOKOnly, "ERRO: CarregarImagemBase64"
    Set stream = Nothing
    Set img = Nothing
    Err.Clear
End Sub

Private Sub Form_Load()
   DoEvents
   Timer1.Enabled = True
   If WhatsAppConectado Then
      Timer1.Enabled = False
      lblStatus.Caption = "CONECTADO"
      lblStatus.BackColor = vbBlack
   Else
      Timer1.Enabled = True
      lblStatus.Caption = "DESCONECTADO"
      lblStatus.BackColor = vbRed
   End If
   DoEvents
End Sub







Private Sub Timer1_Timer()
    
    If WhatsAppConectado Then
       Timer1.Enabled = False
       
'       imgOnLine.Visible = True
'       imgOffLine.Visible = False
'       cpStatusAPIWhats.Visible = True
       lblStatus.Caption = "CONECTADO"
       
       DoEvents
       Sleep 10000
       Unload Me
    Else
'       imgOnLine.Visible = False
'       imgOffLine.Visible = True
'       cpStatusAPIWhats.Visible = True
'       cpStatusAPIWhats = "API WhatsApp Desconectada"
       DoEvents
    End If
End Sub
