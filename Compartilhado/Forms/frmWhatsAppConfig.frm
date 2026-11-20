VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Begin VB.Form frmWhatsAppConfig
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configurar WhatsApp"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrQR
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   120
      Top             =   6600
   End
   Begin ChamaleonBtn.chameleonButton cmdFechar 
      Height          =   495
      Left            =   7080
      TabIndex        =   6
      Top             =   6480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Fechar"
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
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame fraTeste
      Caption         =   " Enviar mensagem de teste "
      Height          =   1095
      Left            =   240
      TabIndex        =   9
      Top             =   5160
      Width           =   8295
      Begin ChamaleonBtn.chameleonButton cmdTeste 
         Height          =   435
         Left            =   6360
         TabIndex        =   5
         Top             =   420
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   767
         BTYPE           =   3
         TX              =   "Enviar &teste"
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
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtNumeroTeste
         Height          =   375
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   4
         Top             =   450
         Width           =   2535
      End
      Begin VB.Label lblNumeroTeste
         Caption         =   "Celular com DDD:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   510
         Width           =   1575
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdDesconectar 
      Height          =   495
      Left            =   6480
      TabIndex        =   3
      Top             =   1560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Desconectar"
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
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdConectar 
      Height          =   495
      Left            =   6480
      TabIndex        =   2
      Top             =   960
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Conectar"
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
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image imgQR
      Height          =   3855
      Left            =   240
      Top             =   1200
      Width           =   3855
   End
   Begin VB.Label lblInstrucoes
      Height          =   2895
      Left            =   4560
      TabIndex        =   8
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Label lblNumero
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   6015
   End
   Begin VB.Label lblSituacao
      Caption         =   "Consultando..."
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   8295
   End
   Begin VB.Label lblAviso
      Caption         =   "Use para enviar orçamentos, documentos e avisos aos seus clientes. Evite mensagens em massa para quem nunca falou com a empresa: o WhatsApp pode bloquear o número."
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   6420
      Width           =   6615
   End
End
Attribute VB_Name = "frmWhatsAppConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Conecta o WhatsApp da empresa (celular) pelo QR code. A conexão fica no servidor de licenças (uma por CNPJ)
'e vale para o Online Commerce e a Ordem de Serviço em todos os computadores. Funções em LicencaOnline.bas (ZapOnline*).

Private Const TEMPO_QR As Long = 180     'segundos mostrando QR codes antes de desistir
Private mInicioQR As Date
Private mOcupado As Boolean
Private mJaConsultou As Boolean

Private Sub Form_Load()
mJaConsultou = False
mOcupado = False
lblNumero.Caption = ""
lblInstrucoes.Caption = "Como conectar:" & vbCrLf & vbCrLf & _
    "1. Clique em Conectar." & vbCrLf & _
    "2. No celular da empresa, abra o WhatsApp." & vbCrLf & _
    "3. Toque nos 3 pontinhos (Android) ou em Configurações (iPhone) e depois em Aparelhos conectados." & vbCrLf & _
    "4. Toque em Conectar um aparelho e aponte a câmera para o QR code ao lado." & vbCrLf & vbCrLf & _
    "Depois de conectado, o celular não precisa ficar perto do computador, mas abra o WhatsApp nele pelo menos uma vez a cada 14 dias."
cmdDesconectar.Enabled = False
fraTeste.Enabled = False
End Sub

Private Sub Form_Activate()
If mJaConsultou Then Exit Sub
mJaConsultou = True
AtualizarSituacao
End Sub

Private Sub AtualizarSituacao()
Dim sit As String, numero As String, nome As String, msg As String

If mOcupado Then Exit Sub
mOcupado = True
Me.MousePointer = vbHourglass
lblSituacao.Caption = "Consultando..."
DoEvents
sit = ZapOnlineSituacao(numero, nome, msg)
Me.MousePointer = vbDefault
mOcupado = False
MostrarSituacao sit, numero, nome, msg
If ZapOnlineForaDoPlano() Then MsgBox msg, vbExclamation, "WhatsApp"
End Sub

Private Sub MostrarSituacao(ByVal Sit As String, ByVal Numero As String, ByVal Nome As String, ByVal Msg As String)
Select Case Sit
    Case "conectado"
        lblSituacao.Caption = "WhatsApp conectado"
        lblSituacao.ForeColor = &H8000&
        lblNumero.Caption = "Número: " & FormatarCelular(Numero) & IIf(Len(Nome) > 0, "   (" & Nome & ")", "")
        Set imgQR.Picture = Nothing
        cmdConectar.Enabled = False
        cmdDesconectar.Enabled = True
        fraTeste.Enabled = True
    Case "conectando"
        lblSituacao.Caption = "Aguardando a leitura do QR code..."
        lblSituacao.ForeColor = &H800000
        lblNumero.Caption = "Leia o QR code com o celular da empresa."
        cmdConectar.Enabled = Not tmrQR.Enabled
        cmdDesconectar.Enabled = True
        fraTeste.Enabled = False
    Case "desconectado"
        lblSituacao.Caption = "WhatsApp desconectado"
        lblSituacao.ForeColor = &HC0&
        lblNumero.Caption = "Clique em Conectar e leia o QR code com o celular da empresa."
        Set imgQR.Picture = Nothing
        cmdConectar.Enabled = True
        cmdDesconectar.Enabled = False
        fraTeste.Enabled = False
    Case Else
        If ZapOnlineForaDoPlano() Then
            'Recurso fora do plano (servidor recusou): nada de QR code
            tmrQR.Enabled = False
            Set imgQR.Picture = Nothing
            lblSituacao.Caption = "WhatsApp não incluído no plano"
            lblSituacao.ForeColor = &HC0&
            lblNumero.Caption = Replace(Msg, vbCrLf & vbCrLf, " ")
            cmdConectar.Enabled = False
            cmdDesconectar.Enabled = False
            fraTeste.Enabled = False
            Exit Sub
        End If
        lblSituacao.Caption = "Não foi possível consultar o WhatsApp"
        lblSituacao.ForeColor = &HC0&
        lblNumero.Caption = Msg
        cmdConectar.Enabled = True
        cmdDesconectar.Enabled = False
        fraTeste.Enabled = False
End Select
End Sub

Private Sub cmdConectar_Click()
mInicioQR = Now
cmdConectar.Enabled = False
tmrQR.Enabled = True
PedirQR
End Sub

'Pede o QR atual ao servidor (o QR muda sozinho a cada ~30 s) até o celular conectar.
Private Sub PedirQR()
Dim sit As String, arq As String, numero As String, nome As String, msg As String

If mOcupado Then Exit Sub
mOcupado = True
Me.MousePointer = vbHourglass
sit = ZapOnlineConectar(arq, numero, nome, msg)
Me.MousePointer = vbDefault
mOcupado = False

If sit = "conectando" Then
    If Len(arq) > 0 Then
        On Error Resume Next
        Set imgQR.Picture = LoadPicture(arq)
        On Error GoTo 0
    End If
    MostrarSituacao sit, numero, nome, msg
    Exit Sub
End If

tmrQR.Enabled = False
MostrarSituacao sit, numero, nome, msg
If sit = "conectado" Then
    MsgBox "WhatsApp conectado com sucesso!" & vbCrLf & vbCrLf & lblNumero.Caption, vbInformation, "WhatsApp"
ElseIf sit = "" Then
    MsgBox msg, vbExclamation, "WhatsApp"
End If
End Sub

Private Sub tmrQR_Timer()
If DateDiff("s", mInicioQR, Now) > TEMPO_QR Then
    tmrQR.Enabled = False
    Set imgQR.Picture = Nothing
    lblSituacao.Caption = "O QR code expirou"
    lblSituacao.ForeColor = &HC0&
    lblNumero.Caption = "Clique em Conectar para gerar um novo QR code."
    cmdConectar.Enabled = True
    Exit Sub
End If
PedirQR
End Sub

Private Sub cmdDesconectar_Click()
Dim msg As String

If MsgBox("Desconectar o WhatsApp da empresa?" & vbCrLf & vbCrLf & _
          "O Online Commerce e a Ordem de Serviço deixam de enviar mensagens até conectar de novo.", _
          vbQuestion + vbYesNo + vbDefaultButton2, "WhatsApp") <> vbYes Then Exit Sub
tmrQR.Enabled = False
Me.MousePointer = vbHourglass
If Not ZapOnlineDesconectar(msg) Then
    Me.MousePointer = vbDefault
    MsgBox msg, vbExclamation, "WhatsApp"
    Exit Sub
End If
Me.MousePointer = vbDefault
AtualizarSituacao
End Sub

Private Sub cmdTeste_Click()
Dim msg As String

If Len(SoNumeros(txtNumeroTeste.Text)) < 10 Then
    MsgBox "Informe o celular com DDD.", vbExclamation, "WhatsApp"
    txtNumeroTeste.SetFocus
    Exit Sub
End If
Me.MousePointer = vbHourglass
cmdTeste.Enabled = False
If ZapOnlineEnviar(txtNumeroTeste.Text, "Mensagem de teste do Online Commerce: o WhatsApp da empresa está conectado.", "", msg) Then
    Me.MousePointer = vbDefault
    MsgBox "Mensagem enviada.", vbInformation, "WhatsApp"
Else
    Me.MousePointer = vbDefault
    MsgBox msg, vbExclamation, "WhatsApp"
End If
cmdTeste.Enabled = True
End Sub

Private Sub cmdFechar_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
tmrQR.Enabled = False
If mOcupado Then Cancel = 1     'esperando o servidor: fecha quando a resposta chegar
End Sub

'5589994275280 -> (89) 9 9427-5280
Private Function FormatarCelular(ByVal Numero As String) As String
Dim d As String
d = SoNumeros(Numero)
If Len(d) >= 12 And Left$(d, 2) = "55" Then d = Mid$(d, 3)
If Len(d) = 11 Then
    FormatarCelular = "(" & Left$(d, 2) & ") " & Mid$(d, 3, 1) & " " & Mid$(d, 4, 4) & "-" & Mid$(d, 8)
ElseIf Len(d) = 10 Then
    FormatarCelular = "(" & Left$(d, 2) & ") " & Mid$(d, 3, 4) & "-" & Mid$(d, 7)
Else
    FormatarCelular = Numero
End If
End Function

Private Function SoNumeros(ByVal Texto As String) As String
Dim i As Long, ch As String
For i = 1 To Len(Texto)
    ch = Mid$(Texto, i, 1)
    If ch Like "#" Then SoNumeros = SoNumeros & ch
Next
End Function
