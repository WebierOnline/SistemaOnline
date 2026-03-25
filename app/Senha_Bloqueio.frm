VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Begin VB.Form Senha_Bloqueio 
   BorderStyle     =   0  'None
   Caption         =   "LICENÇA EXPIRADA"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   5475
   ForeColor       =   &H00000000&
   Icon            =   "Senha_Bloqueio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5475
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
      Height          =   1275
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5295
      Begin VB.TextBox txtMesRef 
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   60
         TabIndex        =   6
         Top             =   540
         Width           =   1815
      End
      Begin VB.TextBox txtSenhaDesbloq 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   0
         Top             =   540
         Width           =   1455
      End
      Begin ChamaleonBtn.chameleonButton cmdSalvarSenha 
         Height          =   315
         Left            =   3420
         TabIndex        =   5
         Top             =   540
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Desbloquear"
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
         MICON           =   "Senha_Bloqueio.frx":23D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblDesbTemp 
         AutoSize        =   -1  'True
         Caption         =   "DESBLOQUEIO TEMPORÁRIO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   60
         TabIndex        =   14
         Top             =   960
         Width           =   2115
      End
      Begin VB.Label lblCodMens 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   1320
         TabIndex        =   8
         Top             =   180
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label Label3 
         Caption         =   "Referente"
         Height          =   195
         Left            =   60
         TabIndex        =   7
         Top             =   300
         Width           =   975
      End
      Begin VB.Label lblSenha1 
         Caption         =   "Desbloqueio"
         Height          =   195
         Left            =   1920
         TabIndex        =   2
         Top             =   300
         Width           =   975
      End
   End
   Begin ChamaleonBtn.chameleonButton chameleonButton1 
      Height          =   255
      Left            =   5100
      TabIndex        =   4
      Top             =   60
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "X"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      MICON           =   "Senha_Bloqueio.frx":23EE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caso não tenha pago, CLIQUE AQUI!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   720
      TabIndex        =   13
      Top             =   2850
      Width           =   2880
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3. Verifique se o código digitado é referente ão mês do bloqueio."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   720
      TabIndex        =   12
      Top             =   2640
      Width           =   3960
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2. Verifique no whatsapp se não chegou o código de desbloqueio."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   720
      TabIndex        =   11
      Top             =   2460
      Width           =   4080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1. Verifique se o pagamento de seu boleto mensal foi efetuado."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   720
      TabIndex        =   10
      Top             =   2280
      Width           =   3870
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Se sua licença expirou, siga os seguintes passos:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   720
      TabIndex        =   9
      Top             =   2040
      Width           =   4125
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   100
      Picture         =   "Senha_Bloqueio.frx":240A
      Top             =   2100
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LICENÇA EXPIRADA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Left            =   60
      TabIndex        =   3
      Top             =   120
      Width           =   5355
   End
   Begin VB.Shape Shape1 
      Height          =   3195
      Left            =   0
      Top             =   0
      Width           =   5475
   End
End
Attribute VB_Name = "Senha_Bloqueio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String
Dim r As ADODB.Recordset
Dim vSenhaTemp As String

Private Sub chameleonButton1_Click()
End
End Sub

Private Sub cmdSalvarSenha_Click()
sSQL = "SELECT codigo, bloqueio, mes_ref, COD_DESBLOQUEIO, COD_TEMP, Debloqueio_Temp, data_bloqueio FROM licenca_pagamentos WHERE (codigo = " & lblCodMens.Caption & ");"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    'If txtSenhaDesbloq.Text <> "" And txtSenhaDesbloqTemp = "" Then
        If txtSenhaDesbloq.Text = r("COD_DESBLOQUEIO") Then
            dbData.Execute "UPDATE licenca_pagamentos SET bloqueio = 0, pago = 1, data_liberacao = '" & Format$(Date, "yyyy-dd-MM") & "' WHERE (codigo = " & lblCodMens.Caption & ");"
            MsgBox "MÊS REFERENTE FOI DESBLOQUEADO" & vbCrLf & "Tente novamente fazer o login no sistema", vbInformation
            Unload Me
            Senha.Show 1
        Else
             MsgBox "Código de desbloqueio errado" & vbCrLf & "Entre em contato com o administador do sistema", vbInformation
             Exit Sub
        End If
    'ElseIf txtSenhaDesbloq.Text = "" And txtSenhaDesbloqTemp <> "" Then
    '    If txtSenhaDesbloqTemp.Text = r("COD_TEMP") Then
    '        Dim vDataBloq As Date
    '        vDataBloq = r("data_bloqueio")
    '        If r("Debloqueio_Temp") = 0 Then
    '            dbData.Execute "UPDATE licenca_pagamentos SET bloqueio = 0, Debloqueio_Temp = 1, data_bloqueio = '" & Format$(vDataBloq + 3, "yyyy-dd-MM") & "' WHERE (codigo = " & lblCodMens.Caption & ");"
    '            MsgBox "VOCÊ USOU UM CÓD. TEMPORÁRIO" & vbCrLf & "Você ganhou mais 3 dias de desbloqueio!", vbInformation
    '            Unload Me
    '            Senha.Show 1
    '        Else
    '            MsgBox "Você já usou esse cód. de desbloqueio." & vbCrLf & "Entre em contato com o administador do sistema", vbInformation
    '            Exit Sub
    '        End If
    '    Else
    '         MsgBox "Código de desbloqueio temporário errado" & vbCrLf & "Entre em contato com o administador do sistema", vbInformation
    '         Exit Sub
    '    End If
    'Else
    '    MsgBox "Digite somente um código de desbloqueio!" & vbCrLf & "Os dois códigos não podem ser preenchidos ao mesmo tempo.", vbInformation
    '    Exit Sub
    'End If
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
lblDesbTemp.Font.Bold = False
lblDesbTemp.ForeColor = &H0&
End Sub


Private Sub Label8_Click()
MsgBox "Olá, tudo bem?" & vbCrLf & " -Caso o pagamento não tenha sido efetuado, aconselhamos pagar usando o QR-CODE(PIX) que fica no proprio boleto em aberto." & vbCrLf & "-CUIDADO! para não pagar o mês errado ao bloqueio." & vbCrLf & "-Após isso, favor nos enviar comprovante de pagamento para o whatsapp do nosso financeiro, no número (89) 9 9427-5280", vbExclamation, "Aviso do Sistema"
End Sub

Private Sub SSTab1_DblClick()

End Sub


Private Sub Label9_Click()

End Sub

Private Sub lblDesbTemp_Click()
vSenhaTemp = InputBox("Informe a senha temporária:", "DESBLOQUEIO TEMPORÁRIO", "")

If Not Vazio(vSenhaTemp) Then
sSQL = "SELECT codigo, bloqueio, mes_ref, COD_DESBLOQUEIO, COD_TEMP, Debloqueio_Temp, data_bloqueio FROM licenca_pagamentos WHERE (codigo = " & lblCodMens.Caption & ");"
Set r = dbData.OpenRecordset(sSQL)
    If Not r.BOF Then
        If vSenhaTemp = r("COD_TEMP") Then
            Dim vDataBloq As Date
            vDataBloq = r("data_bloqueio")
            If r("Debloqueio_Temp") = 0 Then
                dbData.Execute "UPDATE licenca_pagamentos SET bloqueio = 0, Debloqueio_Temp = 1, data_bloqueio = '" & Format$(vDataBloq + 3, "yyyy-dd-MM") & "' WHERE (codigo = " & lblCodMens.Caption & ");"
                MsgBox "VOCÊ USOU UM CÓD. TEMPORÁRIO" & vbCrLf & "Você ganhou mais 3 dias de desbloqueio!", vbInformation
                Unload Me
                Senha.Show 1
            Else
                MsgBox "Você já usou esse cód. de desbloqueio." & vbCrLf & "Entre em contato com o administador do sistema", vbInformation
                Exit Sub
            End If
        Else
             MsgBox "Código de desbloqueio temporário errado" & vbCrLf & "Entre em contato com o administador do sistema", vbInformation
             Exit Sub
        End If
    End If
End If
End Sub


Private Sub lblDesbTemp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
lblDesbTemp.Font.Bold = True
lblDesbTemp.ForeColor = &H80&
End Sub


Private Sub txtSenhaDesbloq_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

