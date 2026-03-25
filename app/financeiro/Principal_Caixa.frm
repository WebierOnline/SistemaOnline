VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Principal_Caixa 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FINANCEIRO"
   ClientHeight    =   5775
   ClientLeft      =   -15
   ClientTop       =   210
   ClientWidth     =   3105
   Icon            =   "Principal_Caixa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   60
      ScaleHeight     =   945
      ScaleWidth      =   2925
      TabIndex        =   5
      Top             =   60
      Width           =   2955
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FINANCEIRO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   900
         TabIndex        =   6
         Top             =   300
         Width           =   1890
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   60
         Picture         =   "Principal_Caixa.frx":23D2
         Top             =   60
         Width           =   900
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   4335
      Left            =   60
      ScaleHeight     =   4275
      ScaleWidth      =   2895
      TabIndex        =   4
      Top             =   1080
      Width           =   2955
      Begin ChamaleonBtn.chameleonButton cmdCaixaParcelas 
         Height          =   615
         Left            =   300
         TabIndex        =   0
         Top             =   180
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   ""
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
         MICON           =   "Principal_Caixa.frx":9258
         PICN            =   "Principal_Caixa.frx":9274
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCaixaSuprimento 
         Height          =   615
         Left            =   300
         TabIndex        =   1
         Top             =   2820
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   ""
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
         BCOL            =   -2147483639
         BCOLO           =   -2147483639
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "Principal_Caixa.frx":9BAC
         PICN            =   "Principal_Caixa.frx":9BC8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCaixaLivro 
         Height          =   615
         Left            =   300
         TabIndex        =   3
         Top             =   3480
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   ""
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Principal_Caixa.frx":A70F
         PICN            =   "Principal_Caixa.frx":A72B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCaixaSangria 
         Height          =   615
         Left            =   300
         TabIndex        =   2
         Top             =   1500
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   ""
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
         MICON           =   "Principal_Caixa.frx":AEA7
         PICN            =   "Principal_Caixa.frx":AEC3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCaixaRetroativa 
         Height          =   615
         Left            =   300
         TabIndex        =   8
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   ""
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Principal_Caixa.frx":B90E
         PICN            =   "Principal_Caixa.frx":B92A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton ccmdCaixaRetirada 
         Height          =   615
         Left            =   300
         TabIndex        =   9
         Top             =   2160
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   ""
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
         MICON           =   "Principal_Caixa.frx":C433
         PICN            =   "Principal_Caixa.frx":C44F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   7
      Top             =   5505
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   2699
            MinWidth        =   2117
            TextSave        =   "18:19"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   2699
            MinWidth        =   2117
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Principal_Caixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCfg As ConfigItem
Public oIni As Ini
Public var_Caixa As String
Public var_Maquina As String
Dim vCodUsuario As Long
Dim sSQL As String
Dim r As ADODB.Recordset
Dim vEntradaJanelas As Boolean



Private Sub FecharForms()
Dim f As Byte

f = Forms.Count

Do While f > 0
    Unload Forms(f - 1)
    If f = Forms.Count Then Exit Do
    f = f - 1
Loop
End Sub

Private Sub ccmdCaixaRetirada_Click()
vEntradaJanelas = True
Unload Me
Caixa_Retirada.Hide
Caixa_Retirada.txtCodFunc.Text = vCodFunc
Caixa_Retirada.Show 1
'Unload Me
vEntradaJanelas = False
End Sub

Private Sub cmdCaixaRetroativa_Click()
vEntradaJanelas = True
Unload Me
Receber_Cadastro.Show
vEntradaJanelas = False
End Sub


Private Sub cmdCaixaLivro_Click()
vEntradaJanelas = True
'abrindo arquivo .ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"

'nome da caixa
var_Caixa = oIni.LerTexto("DADOS_CAIXA", "caixa")
'StatusBar1.Panels(2).Text = var_Caixa

'nome da caixa
var_Maquina = oIni.LerTexto("DADOS_MAQUINA", "maquina")
'StatusBar1.Panels(4).Text = var_Maquina

Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT * " & _
       "FROM caixa_dia " & _
       "WHERE (caixa = '" & var_Caixa & "') and caixa_dia.status = 0;"
Set r = dbData.OpenRecordset(sSQL)

If r.BOF Then
    'FecharForms
    Unload Me
    'Caixa_Controle_semOS.Hide
    'Unload Principal_Caixa
    varFluxoCaixa = False
    Caixa_Controle_semOS.Show
    'Unload PDV
Else
    Unload Me
    'Caixa_Controle_semOS.Hide
    'Unload Principal_Caixa
    varFluxoCaixa = False
    Caixa_Controle_semOS.Show
End If
vEntradaJanelas = False
End Sub

Private Sub cmdCaixaParcelas_Click()
vEntradaJanelas = True
Unload Me
Parcelas.Hide
Parcelas.txtCodFuncionario.Text = vCodFunc
Parcelas.Show
vEntradaJanelas = False
'Unload Me
End Sub

Private Sub cmdCaixaSangria_Click()
vEntradaJanelas = True
'If vCodFunc <> "" Then
    Unload Me
    Caixa_Saida.Hide
    Caixa_Saida.txtCodFunc.Text = vCodFunc
    Caixa_Saida.Show 1
'Else
'    Unload Me
'    Caixa_Saida.Show 1
'End If
vEntradaJanelas = False
End Sub


Private Sub cmdCaixaSuprimento_Click()
vEntradaJanelas = True
Unload Me
Caixa_Suprimento.Show
vEntradaJanelas = False
End Sub

Private Sub cmdContaReceber_Click()
   'Tela_Principal.Menu_Fin_AReceber_Click
End Sub

Private Sub Form_Load()
StatusBar1.Panels(2).Text = Format(Date, "dd/mm/yy")
vCodUsuario = vCodFunc

'permissões
If LerPermissoesUsuario(vCodUsuario, 13) = True Then
    cmdCaixaParcelas.Enabled = True
Else
    cmdCaixaParcelas.Enabled = False
End If

If LerPermissoesUsuario(vCodUsuario, 14) = True Then
    cmdCaixaRetroativa.Enabled = True
Else
    cmdCaixaRetroativa.Enabled = False
End If

If LerPermissoesUsuario(vCodUsuario, 15) = True Then
    cmdCaixaSangria.Enabled = True
Else
    cmdCaixaSangria.Enabled = False
End If

If LerPermissoesUsuario(vCodUsuario, 16) = True Then
    ccmdCaixaRetirada.Enabled = True
Else
    ccmdCaixaRetirada.Enabled = False
End If

If LerPermissoesUsuario(vCodUsuario, 17) = True Then
    cmdCaixaSuprimento.Enabled = True
Else
    cmdCaixaSuprimento.Enabled = False
End If

If LerPermissoesUsuario(vCodUsuario, 18) = True Then
    cmdCaixaLivro.Enabled = True
Else
    cmdCaixaLivro.Enabled = False
End If




End Sub

Public Function LerPermissoesUsuario(vCodUser As Long, permissao As Long) As Boolean
sSQL = "SELECT Usuario_Acessos.Cod_Permissao FROM Usuario INNER JOIN Usuario_Acessos ON Usuario.Codigo = Usuario_Acessos.Cod_Usuario WHERE (Usuario_Acessos.Cod_Usuario = " & vCodUser & ") AND Usuario_Acessos.Cod_Permissao = " & permissao & ";"
Set r = dbData.OpenRecordset(sSQL)

If r.EOF And r.BOF Then
   LerPermissoesUsuario = False ' não achou a permissao
Else
   LerPermissoesUsuario = True 'aqui achou
End If
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'EncerrarPrograma
If vChamouCaixa = "PDV" Then
    If vEntradaJanelas = False Then
        Principal_Caixa.Hide
        'PDV.Show 'desativei somente para geerar o online comerce
    End If
Else
End If
End Sub

