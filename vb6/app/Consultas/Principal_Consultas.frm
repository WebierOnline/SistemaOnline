VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Principal_Consultas 
   Caption         =   "CONSULTAS"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5190
   Icon            =   "Principal_Consultas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   60
      ScaleHeight     =   945
      ScaleWidth      =   5025
      TabIndex        =   2
      Top             =   60
      Width           =   5055
      Begin VB.Image Image1 
         Height          =   720
         Left            =   780
         Picture         =   "Principal_Consultas.frx":23D2
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONSULTAS"
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
         Left            =   1740
         TabIndex        =   3
         Top             =   300
         Width           =   1920
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   4755
      Left            =   60
      ScaleHeight     =   4695
      ScaleWidth      =   2415
      TabIndex        =   1
      Top             =   1140
      Width           =   2475
      Begin ChamaleonBtn.chameleonButton cmdConsVendas 
         Height          =   615
         Left            =   60
         TabIndex        =   4
         Top             =   60
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
         MICON           =   "Principal_Consultas.frx":2F70
         PICN            =   "Principal_Consultas.frx":2F8C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdConsOS 
         Height          =   615
         Left            =   60
         TabIndex        =   5
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   ""
         ENAB            =   0   'False
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
         MICON           =   "Principal_Consultas.frx":8AFD
         PICN            =   "Principal_Consultas.frx":8B19
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdConsServicosPecas 
         Height          =   615
         Left            =   60
         TabIndex        =   6
         Top             =   1380
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
         MICON           =   "Principal_Consultas.frx":ED7A
         PICN            =   "Principal_Consultas.frx":ED96
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
   Begin VB.PictureBox Picture1 
      Height          =   4755
      Left            =   2580
      ScaleHeight     =   4695
      ScaleWidth      =   2475
      TabIndex        =   0
      Top             =   1140
      Width           =   2535
      Begin ChamaleonBtn.chameleonButton cmdConsParcelas 
         Height          =   615
         Left            =   60
         TabIndex        =   7
         Top             =   60
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
         MICON           =   "Principal_Consultas.frx":14EE7
         PICN            =   "Principal_Consultas.frx":14F03
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdConsEstoqueM 
         Height          =   615
         Left            =   60
         TabIndex        =   8
         Top             =   720
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
         MICON           =   "Principal_Consultas.frx":1ADDA
         PICN            =   "Principal_Consultas.frx":1ADF6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCaixa_Cons_Setor 
         Height          =   615
         Left            =   60
         TabIndex        =   9
         Top             =   1380
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
         MICON           =   "Principal_Consultas.frx":2119E
         PICN            =   "Principal_Consultas.frx":211BA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCaixa_Cons_Mes 
         Height          =   615
         Left            =   60
         TabIndex        =   10
         Top             =   2040
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
         MICON           =   "Principal_Consultas.frx":26F66
         PICN            =   "Principal_Consultas.frx":26F82
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
      TabIndex        =   11
      Top             =   5985
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4339
            Text            =   "Online.Info Sistemas"
            TextSave        =   "Online.Info Sistemas"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "08:56"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2117
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
Attribute VB_Name = "Principal_Consultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCaixa_Cons_Mes_Click()
   'If Tela_Principal.txtNivel.Text <> "1" Then MsgBox "Seu nível de acesso não permite a essa operação!", vbInformation, "Aviso do Sistema": Exit Sub
   Lanc_Caixa.Show 1
End Sub

Private Sub cmdCaixa_Cons_Setor_Click()
   'If Tela_Principal.txtNivel.Text <> "1" Then MsgBox "Seu nível de acesso não permite a essa operação!", vbInformation, "Aviso do Sistema": Exit Sub
   'Unload Me
   Fluxo_Caixa.Show 1
End Sub

'Private Sub cmdConsClientes_Click()
'   Clientes_Consulta.Show 1
'End Sub

Private Sub cmdConsEstoqueM_Click()
   Consulta_Estoque_Minimo.Show 1
End Sub

Private Sub cmdCONSos_Click()
   Ordem_Servicos_Consulta_Geral.Show 1
   Unload Me
End Sub

Private Sub cmdConsParcelas_Click()
   Parcelas_Consulta.Show
   Unload Me
End Sub

'Private Sub cmdConsServicos_Click()
'   Unload Me
'   Ordem_Servicos_Consulta_Servicos.Show 1
'End Sub

Private Sub cmdConsServicosPecas_Click()
   OS_Consulta_Servicos_Pecas.Show 1
End Sub

Private Sub cmdConsVendas_Click()
Vendas_Consulta_PorPedidos.Show 1
Unload Me
End Sub

Private Sub Form_Load()
   StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
End Sub
