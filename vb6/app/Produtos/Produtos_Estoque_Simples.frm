VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Produtos_Estoque_Simples 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ESTOQUE SIMPLES"
   ClientHeight    =   10035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16275
   Icon            =   "Produtos_Estoque_Simples.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Produtos_Estoque_Simples.frx":1D82
   ScaleHeight     =   10035
   ScaleWidth      =   16275
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optMostrarFiscal 
      Caption         =   "Fiscal"
      Height          =   195
      Left            =   7140
      TabIndex        =   39
      Top             =   7680
      Width           =   795
   End
   Begin VB.Frame frmTotalFiscal 
      Caption         =   "Totais"
      Height          =   675
      Left            =   13320
      TabIndex        =   36
      Top             =   8040
      Visible         =   0   'False
      Width           =   2895
      Begin VB.Label lblValorTotalFiscal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   1260
         TabIndex        =   38
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Total Fiscal:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame frmSenha 
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
      Height          =   795
      Left            =   9900
      TabIndex        =   30
      Top             =   8700
      Visible         =   0   'False
      Width           =   1935
      Begin VB.TextBox txtSenha 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   31
         Top             =   360
         Width           =   1335
      End
      Begin ChamaleonBtn.chameleonButton cmdSenha 
         Height          =   315
         Left            =   1500
         TabIndex        =   32
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "OK"
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
         MICON           =   "Produtos_Estoque_Simples.frx":264C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.OptionButton optMostrarTodos 
      Caption         =   "Todos"
      Height          =   195
      Left            =   6240
      TabIndex        =   29
      Top             =   7680
      Width           =   795
   End
   Begin VB.OptionButton optMostrarZerados 
      Caption         =   "Somente zerados"
      Height          =   195
      Left            =   4500
      TabIndex        =   28
      Top             =   7680
      Width           =   1635
   End
   Begin VB.OptionButton optMostrarNegativos 
      Caption         =   "Somente Negativos"
      Height          =   195
      Left            =   2460
      TabIndex        =   27
      Top             =   7680
      Width           =   1935
   End
   Begin VB.OptionButton optMostrarQuant 
      Caption         =   "Somente com quantidade"
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   7680
      Value           =   -1  'True
      Width           =   2235
   End
   Begin VB.Frame Frame1 
      Caption         =   "Critérios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1635
      Left            =   120
      TabIndex        =   12
      Top             =   8040
      Width           =   1395
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.OptionButton optCategoria 
         Caption         =   "Categoria"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   1035
      End
      Begin VB.OptionButton optDesc 
         Caption         =   "Descrição"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   900
         Width           =   1035
      End
      Begin VB.OptionButton optCodBarra 
         Caption         =   "Cód. Barra"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Busca Avançada"
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
      Left            =   1560
      TabIndex        =   6
      Top             =   8040
      Width           =   5085
      Begin VB.TextBox txtCodBarra 
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.CheckBox chkDescPorIniciais 
         Caption         =   "Por Iniciais"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1500
         TabIndex        =   11
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkDescPorProduto 
         Caption         =   "Por Produto"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   1080
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.ComboBox cboConsLinha 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.ComboBox cboDesc 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   4815
      End
      Begin ChamaleonBtn.chameleonButton cmdLocalizar 
         Height          =   495
         Left            =   3480
         TabIndex        =   21
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Exibir"
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
         MICON           =   "Produtos_Estoque_Simples.frx":2668
         PICN            =   "Produtos_Estoque_Simples.frx":2684
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblCodBarra 
         Caption         =   "Cód. de Barra"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblCategoria 
         Caption         =   "Categoria"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblDesc 
         Caption         =   "Descrição"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Ordem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1635
      Left            =   6660
      TabIndex        =   5
      Top             =   8040
      Width           =   2535
      Begin VB.OptionButton optORDTFiscal 
         Caption         =   "Total Fiscal"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   1380
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Frame Frame4 
         Caption         =   "Direção"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1455
         Left            =   1440
         TabIndex        =   46
         Top             =   120
         Width           =   975
         Begin VB.OptionButton optORDDescrescente 
            Caption         =   "Desc"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   420
            Width           =   675
         End
         Begin VB.OptionButton optORDASC 
            Caption         =   "Asc"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Value           =   -1  'True
            Width           =   675
         End
      End
      Begin VB.OptionButton optORDValorCusto 
         Caption         =   "Valor Custo"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   1020
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.OptionButton optORDLinha 
         Caption         =   "Categoria"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   840
         Width           =   1035
      End
      Begin VB.OptionButton optORDValor 
         Caption         =   "Valor Venda"
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   660
         Width           =   1275
      End
      Begin VB.OptionButton ORDQuantFiscal 
         Caption         =   "Quant. Fiscal"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   1200
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.OptionButton optORDQuant 
         Caption         =   "Quant."
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   480
         Width           =   1035
      End
      Begin VB.OptionButton optORDDesc 
         Caption         =   "Descrição"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.PictureBox picAguarde 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   6300
      Picture         =   "Produtos_Estoque_Simples.frx":2F5E
      ScaleHeight     =   1095
      ScaleWidth      =   2895
      TabIndex        =   4
      Top             =   3540
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   2520
      TabIndex        =   3
      Top             =   2640
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   16125
      TabIndex        =   0
      Top             =   60
      Width           =   16155
      Begin VB.Image Image1 
         Height          =   645
         Left            =   480
         Picture         =   "Produtos_Estoque_Simples.frx":3F96
         Top             =   120
         Width           =   645
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "AJUSTE DE ESTOQUE"
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
         Left            =   1365
         TabIndex        =   1
         Top             =   240
         Width           =   3360
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdSair 
      Height          =   315
      Left            =   13380
      TabIndex        =   9
      Top             =   7680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
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
      MICON           =   "Produtos_Estoque_Simples.frx":9969
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   22
      Top             =   9765
      Width           =   16275
      _ExtentX        =   28707
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   24368
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "17:02"
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
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   6615
      Left            =   60
      TabIndex        =   2
      Top             =   1020
      Width           =   16155
      _ExtentX        =   28496
      _ExtentY        =   11668
      _Version        =   393216
      Cols            =   5
      AllowBigSelection=   0   'False
      HighLight       =   0
      ScrollBars      =   2
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdAtualizar 
      Height          =   315
      Left            =   11640
      TabIndex        =   23
      Top             =   7680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Salvar"
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
      MICON           =   "Produtos_Estoque_Simples.frx":9985
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdAtualizarPreco 
      Height          =   315
      Left            =   8160
      TabIndex        =   24
      Top             =   7680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Alterar Preço"
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
      MICON           =   "Produtos_Estoque_Simples.frx":99A1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdAtualizarQuant 
      Height          =   315
      Left            =   9900
      TabIndex        =   25
      Top             =   7680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Alterar Quantidade"
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
      MICON           =   "Produtos_Estoque_Simples.frx":99BD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblCodUsuario 
      AutoSize        =   -1  'True
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10440
      TabIndex        =   35
      Top             =   8160
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblUsuario 
      AutoSize        =   -1  'True
      Caption         =   "Nenhum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10440
      TabIndex        =   34
      Top             =   8340
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Usuário:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9660
      TabIndex        =   33
      Top             =   8340
      Visible         =   0   'False
      Width           =   705
   End
End
Attribute VB_Name = "Produtos_Estoque_Simples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String
Dim r As ADODB.Recordset
Private moCombo As cComboHelper
Private iRow As Long, iCol As Long
Dim varTipoValorVenda As String
Dim var_Indice As String
Dim var_Direcao As String

'arquivo .ini
Public cCfg As ConfigItem
'Public oIni As Ini


Private Sub LimparGrid2()
 
sSQL = "SELECT  produtos.NCM AS var_NCM, produtos.CFOP AS var_CFOP, produtos.ICMSCST AS var_ICMS, produtos.categoria AS var_cat, produtos.fabricante AS var_fab, produtos.PRATELEIRA AS var_Local, " & _
   "produtos.codigo AS var_cod, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, produtos.quant_estoque AS var_quant, produtos.ESTOQUE_FISCAL AS var_EstoqueFiscal, produtos.UNID_MEDIDA AS var_UnidMed, " & _
   "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
   "FROM produtos " & _
   "WHERE 1 = 0"

Set r = dbData.OpenRecordset(sSQL)

Formatar_Grid r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub LimparGrid()
   Dim i As Integer
   
   txtEdit.Text = ""
   
   With Grid
      .Clear
      .Cols = 9
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 0
      .ColWidth(3) = 1500
      .ColWidth(4) = 4200
      .ColWidth(5) = 800
      .ColWidth(6) = 800
      .ColWidth(7) = 1000
      .ColWidth(8) = 2000
      
      '.RowHeight(-1) = (315 * 1)    'definir a altura da linha
      
      .TextMatrix(0, 1) = "CÓD.ENT"
      .TextMatrix(0, 2) = "CÓD.PROD"
      .TextMatrix(0, 3) = "CÓD.BARRA"
      .TextMatrix(0, 4) = "DESCRIÇÃO"
      .TextMatrix(0, 5) = "QUANT."
      .TextMatrix(0, 6) = "MIN."
      .TextMatrix(0, 7) = "VENDA"
      .TextMatrix(0, 8) = "CATEGORIA"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .Redraw = False
      
      'ALINHAMENTO
      .ColAlignment(2) = 1
      
      .rows = .rows + 1
      
      .rows = .rows - 1
      .Redraw = True
   End With
End Sub

Private Sub chkDescPorIniciais_Click()
If optDesc.Value = Unchecked Then Exit Sub

If chkDescPorIniciais.Value = Checked Then
   cboDesc.Clear
   chkDescPorProduto.Value = Unchecked
   cboDesc.SetFocus
End If
End Sub

Private Sub chkDescPorProduto_Click()
If optDesc.Value = Unchecked Then Exit Sub

If chkDescPorProduto.Value = Checked Then
   chkDescPorIniciais.Value = Unchecked
   cboDesc.SetFocus
End If
End Sub

Private Sub MostrarCriterios()
   Dim var_Criterio As String
   
   
   var_Criterio = ""
   
   If chkDescPorProduto.Value = Checked Then
      var_Criterio = var_Criterio & IIf(optDesc.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.descricao = '" & cboDesc.Text & "'", "")
   ElseIf chkDescPorIniciais.Value = Checked Then
      var_Criterio = Chr$(39) & cboDesc.Text & "%" & Chr(39)
      var_Criterio = IIf(optDesc.Value, IIf(var_Criterio <> "", "", " AND ") & "produtos.descricao  LIKE " & var_Criterio & "", "")
   End If
   
   var_Criterio = var_Criterio & IIf(optCategoria.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.categoria = '" & cboConsLinha.Text & "'", "")
   var_Criterio = var_Criterio & IIf(optCodBarra.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.cod_barra = '" & txtCodBarra.Text & "'", "")
   
   If var_Criterio <> "" Then var_Criterio = " WHERE (produtos.ativo = 1) and " & var_Criterio
   
   var_Indice = ""
   var_Indice = var_Indice & IIf(optORDQuant.Value, IIf(var_Indice <> "", ", ", "") & "quant_estoque", "")
   var_Indice = var_Indice & IIf(optORDDesc.Value, IIf(var_Indice <> "", ", ", "") & "produtos.descricao", "")
   var_Indice = var_Indice & IIf(ORDQuantFiscal.Value, IIf(var_Indice <> "", ", ", "") & "quant_min", "")
   var_Indice = var_Indice & IIf(optORDValor.Value, IIf(var_Indice <> "", ", ", "") & "produtos_entrada_itens.venda", "")
   var_Indice = var_Indice & IIf(optORDLinha.Value, IIf(var_Indice <> "", ", ", "") & "produtos.categoria", "")
   
   If var_Indice <> "" Then var_Indice = " ORDER BY " & var_Indice
   
   sSQL = "SELECT produtos.NCM AS var_NCM, produtos.ICMSCST AS var_ICMS, produtos.CFOP AS var_CFOP, produtos.categoria AS var_cat, produtos.fabricante AS var_fab, produtos.PRATELEIRA AS var_Local, " & _
      "produtos.codigo AS var_cod, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, produtos.quant_estoque AS var_quant, produtos.ESTOQUE_FISCAL AS var_EstoqueFiscal, produtos.UNID_MEDIDA AS var_UnidMed, " & _
      "(SELECT TOP 1 Produtos_Precos.CUSTO FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS custo, " & _
      "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
      "FROM produtos " & var_Criterio & " " & var_Indice
   
   Set r = dbData.OpenRecordset(sSQL)
   'Debug.Print sSQL
   
       If r.RecordCount > 32000 Then
        MsgBox "A Consulta retornou um valor maior de registros que é permitido na grade!", vbInformation, "Aviso do sistema"
        LimparGrid2
        Exit Sub
    Else
        Formatar_Grid r
    End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
End Sub

Private Sub optORDASC_Click()
cmdLocalizar_Click
End Sub

Private Sub optORDDesc_Click()
cmdLocalizar_Click
End Sub

Private Sub optORDDescrescente_Click()
cmdLocalizar_Click
End Sub

Private Sub optORDLinha_Click()
cmdLocalizar_Click
End Sub

Private Sub optORDQuant_Click()
cmdLocalizar_Click
End Sub

Private Sub optORDTFiscal_Click()
cmdLocalizar_Click
End Sub

Private Sub optORDValorCusto_Click()
cmdLocalizar_Click
End Sub

Private Sub ORDQuantFiscal_Click()
cmdLocalizar_Click
End Sub

Private Sub optORDValor_Click()
cmdLocalizar_Click
End Sub

Private Sub cmdAtualizar_Click()
Dim i As Integer

picAguarde.Visible = True
DoEvents
    'txtDescricao.Text = TirarEspaco(txtDescricao.Text)
For i = 1 To Grid.rows - 1
   'Atualiza a tabela de produtos
   sSQL = "UPDATE produtos SET " & _
      "cod_barra = '" & Grid.TextMatrix(i, 3) & "', " & _
      "descricao = '" & TirarEspaco(Grid.TextMatrix(i, 4)) & "', " & _
      "UNID_MEDIDA = '" & Grid.TextMatrix(i, 6) & "', " & _
      "categoria = '" & Grid.TextMatrix(i, 7) & "', " & _
      "fabricante = '" & Grid.TextMatrix(i, 5) & "', " & _
      "PRATELEIRA = '" & Grid.TextMatrix(i, 8) & "', " & _
      "ESTOQUE_FISCAL = " & Replace(CDbl(Grid.TextMatrix(i, 9)), ",", ".") & " " & _
      "WHERE (codigo = " & Grid.TextMatrix(i, 2) & ");"
            

      '" & Replace(CDbl(Grid.TextMatrix(i, 15)), ",", ".") & "
      '"ESTOQUE_FISCAL = '" & Grid.TextMatrix(i, 9) & "' " & _
    'sSQL = "UPDATE produtos SET " & _
      "cod_barra = '" & Grid.TextMatrix(i, 3) & "', " & _
      "descricao = '" & Grid.TextMatrix(i, 4) & "', " & _
      "UNID_MEDIDA = '" & Grid.TextMatrix(i, 6) & "', " & _
      "categoria = '" & Grid.TextMatrix(i, 7) & "', " & _
      "fabricante = '" & Grid.TextMatrix(i, 5) & "', " & _
      "PRATELEIRA = '" & Grid.TextMatrix(i, 8) & "', " & _
      "ESTOQUE_FISCAL = '" & Grid.TextMatrix(i, 9) & "' " & _
      "WHERE (codigo = 2057);"
      'Debug.Print sSQL
   dbData.Execute sSQL
Next

picAguarde.Visible = False
cmdLocalizar_Click
End Sub
Public Function TirarEspaco(ByVal Value As String) As String
Dim bRepete As Boolean
Value = Replace$(Value, "'", vbNullString)
Do
  Value = Replace$(Value, "  ", " ")
  bRepete = InStr(1, Value, "  ", vbTextCompare)
  Value = Trim(Value)
Loop Until Not bRepete

TirarEspaco = Value
End Function
Private Sub cmdAtualizarPreco_Click()
Me.Hide
'Load Produtos_AjustoPreco
Produtos_AjustoPreco.Show
Produtos_AjustoPreco.txtCodProduto.Text = (Grid.TextMatrix(Grid.Row, 2))
End Sub

Private Sub cmdAtualizarQuant_Click()
Me.Hide
Dim i As Integer
i = Grid.Row

If Grid.TextMatrix(i, 2) = "" Then Exit Sub

Produtos_AdicionarQuant.Show
Produtos_AdicionarQuant.txtCodProduto.Text = (Grid.TextMatrix(i, 2))
Produtos_AdicionarQuant.txtCodUsuario.Text = lblCodUsuario.Caption
Produtos_AdicionarQuant.txtQuantNova.SetFocus
End Sub


Private Sub cmdLocalizar_Click()
Dim varTipoMostrar As String


'criado pela IA
If optMostrarQuant.Value = True Then
    varTipoMostrar = " AND p.quant_estoque > 0" ' Mudei de produtos para p
ElseIf optMostrarNegativos.Value = True Then
    varTipoMostrar = " AND p.quant_estoque < 0"
ElseIf optMostrarZerados.Value = True Then
    varTipoMostrar = " AND p.quant_estoque = 0"
ElseIf optMostrarFiscal.Value = True Then
    varTipoMostrar = " AND p.ESTOQUE_FISCAL > 0"
ElseIf optMostrarTodos.Value = True Then
    varTipoMostrar = " "
End If

'meu código
'If optMostrarQuant.Value = True Then
'    varTipoMostrar = " AND produtos.quant_estoque > 0"
'ElseIf optMostrarNegativos.Value = True Then
'    varTipoMostrar = " AND produtos.quant_estoque < 0"
'ElseIf optMostrarZerados.Value = True Then
'    varTipoMostrar = " AND produtos.quant_estoque = 0"
'ElseIf optMostrarFiscal.Value = True Then
'    varTipoMostrar = " AND produtos.ESTOQUE_FISCAL > 0"
'ElseIf optMostrarTodos.Value = True Then
'    varTipoMostrar = " "
'End If

'criado pela IA
If optORDDesc.Value = True Then
   var_Indice = "p.descricao"
ElseIf optORDQuant.Value = True Then
   var_Indice = "p.quant_estoque"
ElseIf ORDQuantFiscal.Value = True Then
   var_Indice = "p.ESTOQUE_FISCAL"
ElseIf optORDValor.Value = True Then
   var_Indice = "precos.VALOR_VV" ' <--- Aqui a mágica do maior para o menor
ElseIf optORDValorCusto.Value = True Then
   var_Indice = "precos.CUSTO" ' Ordena pelo custo do maior para o menor
ElseIf optORDLinha.Value = True Then
   var_Indice = "p.categoria"
ElseIf optORDTFiscal.Value = True Then
   ' Realiza o cálculo (Custo * Estoque Fiscal) para ordenar
   var_Indice = "(precos.CUSTO * p.ESTOQUE_FISCAL)"
End If

If optORDASC.Value = True Then
   var_Direcao = " ASC"
Else
   var_Direcao = " DESC"
End If

'meu código
'If optORDDesc.Value = True Then
'   var_Indice = "produtos.descricao"
'ElseIf optORDQuant.Value = True Then
'   var_Indice = "produtos.quant_estoque"
'ElseIf ORDQuantFiscal.Value = True Then
'   var_Indice = "produtos.ESTOQUE_FISCAL"
'ElseIf optORDValor.Value = True Then
'    var_Indice = "(SELECT TOP 1 VALOR_VV FROM Produtos_Precos Where COD_PRODUTO = produtos.codigo order by CODIGO desc) DESC"
'ElseIf optORDLinha.Value = True Then
'   var_Indice = "produtos.categoria"
'End If


If optTodos.Value = True Then
    sSQL = "SELECT p.NCM AS var_NCM, p.CFOP AS var_CFOP, p.ICMSCST AS var_ICMS, " & _
           "p.categoria AS var_cat, p.fabricante AS var_fab, p.PRATELEIRA AS var_Local, " & _
           "p.codigo AS var_cod, p.cod_barra AS var_codbarra, p.descricao AS var_desc, " & _
           "p.quant_estoque AS var_quant, p.ESTOQUE_FISCAL AS var_EstoqueFiscal, p.UNID_MEDIDA AS var_UnidMed, " & _
           "precos.CUSTO AS custo, precos.VALOR_VV AS venda " & _
           "FROM produtos p " & _
           "LEFT JOIN (" & _
           "   SELECT COD_PRODUTO, CUSTO, VALOR_VV, " & _
           "   ROW_NUMBER() OVER (PARTITION BY COD_PRODUTO ORDER BY CODIGO DESC) as RN " & _
           "   FROM Produtos_Precos" & _
           ") precos ON p.codigo = precos.COD_PRODUTO AND precos.RN = 1 " & _
           "WHERE (p.ativo = 1) " & varTipoMostrar & _
           " ORDER BY " & var_Indice & var_Direcao


    'meu código
    'sSQL = "SELECT  produtos.NCM AS var_NCM, produtos.CFOP AS var_CFOP, produtos.ICMSCST AS var_ICMS, produtos.categoria AS var_cat, produtos.fabricante AS var_fab, produtos.PRATELEIRA AS var_Local, " & _
      "produtos.codigo AS var_cod, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, produtos.quant_estoque AS var_quant, produtos.ESTOQUE_FISCAL AS var_EstoqueFiscal, produtos.UNID_MEDIDA AS var_UnidMed, " & _
      "(SELECT TOP 1 Produtos_Precos.CUSTO FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS custo, " & _
      "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
      "FROM produtos " & _
      "WHERE (produtos.ativo = 1) " & varTipoMostrar & " ORDER BY " & var_Indice
   Set r = dbData.OpenRecordset(sSQL)
   
    If r.RecordCount > 32000 Then
        MsgBox "A Consulta retornou um valor maior de registros que é permitido na grade!", vbInformation, "Aviso do sistema"
        LimparGrid2
        Exit Sub
    Else
        If optMostrarFiscal.Value = True Then
            Formatar_Grid_Fiscal r
        Else
            Formatar_Grid r
        End If
        
    End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
Else
   MostrarCriterios
End If
'Debug.Print sSQL
If optCodBarra.Value = True Then txtCodBarra_GotFocus
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub cmdSenha_Click()
sSQL = "SELECT * FROM usuario WHERE (password = '" & txtSenha.Text & "');"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    lblCodUsuario.Caption = ValidateNull(r("codigo"))
    lblUsuario.Caption = ValidateNull(r("login"))
    
    If lblCodUsuario.Caption = "" Then Exit Sub
    sSQL = "SELECT Usuario_permissoes.Codigo, Usuario_permissoes.permissao " & _
           "FROM Usuario_permissoes INNER JOIN Usuario_Acessos ON Usuario_permissoes.Codigo = Usuario_Acessos.Cod_Permissao " & _
           "WHERE (Usuario_permissoes.permissao = 'AJUSTE') AND (Usuario_Acessos.Cod_Usuario = " & lblCodUsuario.Caption & ")"
    Set r = dbData.OpenRecordset(sSQL)
    
    If Not r.BOF Then
        cmdAtualizarPreco.Enabled = True
        cmdAtualizarQuant.Enabled = True
        txtSenha.Text = ""
        frmSenha.Visible = False
    Else
        cmdAtualizarPreco.Enabled = False
        cmdAtualizarQuant.Enabled = False
        ShowMsg "ACESSO NEGADO!" & vbCrLf & "Você não tem nivel de acesso a esse recurso", vbInformation
        lblCodUsuario.Caption = ""
        lblUsuario.Caption = ""
    End If
Else
    ShowMsg "ACESSO NEGADO!" & vbCrLf & "Senha Inválida!", vbInformation
    lblCodUsuario.Caption = ""
    lblUsuario.Caption = ""
End If
End Sub

Private Sub Form_Activate()
cmdLocalizar_Click
End Sub

Private Sub Form_Load()
Set moCombo = New cComboHelper

'tipo de venda = 1 simples e 2 multiplus preços
Set cCfg = sysConfig("TIPOVALORVENDA")
varTipoValorVenda = cCfg.Value
Set cCfg = Nothing
End Sub

Private Sub cboDesc_Change()
   'cboDesc_Click
End Sub

Private Sub cboDesc_Click()
   'If chkDescPorProduto.Value = Checked Then
   '   If cboDesc.Text = "" Then Exit Sub
   '   MostrarCriterios
   'ElseIf chkDescPorIniciais.Value = Checked Then
   '   MostrarCriterios
   'End If
End Sub

Private Sub cboDesc_GotFocus()
If chkDescPorProduto.Value = Checked Then
   cboDesc.Clear
   
   sSQL = "SELECT DISTINCT descricao FROM produtos WHERE (produtos.ativo = 1) ORDER BY descricao;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboDesc.AddItem ValidateNull(r("descricao"))
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboDesc
End If
End Sub

Private Sub cboDesc_LostFocus()
   'cboDesc_Click
End Sub

Private Sub cboConsLinha_Click()
   'If cboConsLinha.Text <> "" Then MostrarCriterios
End Sub

Private Sub cboConsLinha_GotFocus()
cboConsLinha.Clear

sSQL = "SELECT DISTINCT categoria FROM produtos where (produtos.ativo = 1) ORDER BY categoria;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboConsLinha.AddItem ValidateNull(r("categoria"))
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

moCombo.AttachTo cboConsLinha
End Sub

Private Sub cboConsLinha_LostFocus()
   'cboConsLinha_Click
End Sub

Private Sub Formatar_Grid_Fiscal(rTabela As ADODB.Recordset)
   Dim i As Integer
   Dim VarTotalGrid As Currency
   
    LimparGrid
    picAguarde.Visible = True
    DoEvents

    VarTotalGrid = 0

'If varTipoValorVenda = 2 Then
   With Grid
      .Clear
      .Cols = 14
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 0
      .ColWidth(3) = 1500
      .ColWidth(4) = 4200
      .ColWidth(5) = 1600
      .ColWidth(6) = 800
      .ColWidth(7) = 1750
      .ColWidth(8) = 800
      .ColWidth(9) = 1000
      .ColWidth(10) = 1000
      .ColWidth(11) = 1000
      .ColWidth(12) = 1100
      .ColWidth(13) = 1100
      
      '.RowHeight(-1) = (315 * 1)    'definir a altura da linha
      
      .TextMatrix(0, 1) = "CÓD.ENT"
      .TextMatrix(0, 2) = "CÓD.PROD"
      .TextMatrix(0, 3) = "CÓD.BARRA"
      .TextMatrix(0, 4) = "DESCRIÇÃO"
      .TextMatrix(0, 5) = "FABRICANTE"
      .TextMatrix(0, 6) = "MED."
      .TextMatrix(0, 7) = "CATEGORIA"
      .TextMatrix(0, 8) = "LOCAL"
      .TextMatrix(0, 9) = "FISCAL"
      .TextMatrix(0, 10) = "ESTOQUE"
      .TextMatrix(0, 11) = "VENDA"
      .TextMatrix(0, 12) = "CUSTO"
      .TextMatrix(0, 13) = "T.FISCAL"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            'ALINHAMENTO
            .ColAlignment(2) = 1
            VarTotalGrid = 0
            '.TextMatrix(.Rows - 1, 1) = ValidateNull(rTabela("var_codent"))
            .TextMatrix(.rows - 1, 2) = ValidateNull(rTabela("var_cod"))
            .TextMatrix(.rows - 1, 3) = ValidateNull(rTabela("var_codbarra"))
            .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("var_desc"))
            .TextMatrix(.rows - 1, 5) = ValidateNull(rTabela("var_fab"))
            .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("var_UnidMed"))
            .TextMatrix(.rows - 1, 7) = Format$(ValidateNull(rTabela("var_cat")), ocMONEY)
            .TextMatrix(.rows - 1, 8) = ValidateNull(rTabela("var_Local"))
            .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("var_EstoqueFiscal"))
            .TextMatrix(.rows - 1, 10) = ValidateNull(rTabela("var_quant"))
            .TextMatrix(.rows - 1, 11) = Format$(ValidateNull(rTabela("venda")), ocMONEY)
            .TextMatrix(.rows - 1, 12) = Format$(ValidateNull(rTabela("custo")), ocMONEY)
            
            VarTotalGrid = .TextMatrix(.rows - 1, 12) * .TextMatrix(.rows - 1, 9)
            .TextMatrix(.rows - 1, 13) = Format(VarTotalGrid, ocMONEY)
            '.TextMatrix(.rows - 1, 12) = Format$(ValidateNull(rTabela("venda")), ocMONEY)

            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If
      lblValorTotalFiscal.Caption = Format(SomaGrid(Grid, 13), ocMONEY)
      .rows = .rows - 1
      .Redraw = True
      picAguarde.Visible = False
   End With
'Else
'End If
End Sub
Private Sub Formatar_Grid(rTabela As ADODB.Recordset)
   Dim i As Integer
   Dim VarTotalGrid As Currency
   
    LimparGrid
    picAguarde.Visible = True
    DoEvents

    VarTotalGrid = 0

'If varTipoValorVenda = 2 Then
   With Grid
      .Clear
      .Cols = 11
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 0
      .ColWidth(3) = 1500
      .ColWidth(4) = 5200
      .ColWidth(5) = 1600
      .ColWidth(6) = 800
      .ColWidth(7) = 1750
      .ColWidth(8) = 800
      .ColWidth(9) = 1000
      .ColWidth(10) = 1000
      
      '.RowHeight(-1) = (315 * 1)    'definir a altura da linha
      
      .TextMatrix(0, 1) = "CÓD.ENT"
      .TextMatrix(0, 2) = "CÓD.PROD"
      .TextMatrix(0, 3) = "CÓD.BARRA"
      .TextMatrix(0, 4) = "DESCRIÇÃO"
      .TextMatrix(0, 5) = "FABRICANTE"
      .TextMatrix(0, 6) = "MED."
      .TextMatrix(0, 7) = "CATEGORIA"
      .TextMatrix(0, 8) = "LOCAL"
      .TextMatrix(0, 9) = "ESTOQUE"
      .TextMatrix(0, 10) = "VENDA"

      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            'ALINHAMENTO
            .ColAlignment(2) = 1
            VarTotalGrid = 0
            '.TextMatrix(.Rows - 1, 1) = ValidateNull(rTabela("var_codent"))
            .TextMatrix(.rows - 1, 2) = ValidateNull(rTabela("var_cod"))
            .TextMatrix(.rows - 1, 3) = ValidateNull(rTabela("var_codbarra"))
            .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("var_desc"))
            .TextMatrix(.rows - 1, 5) = ValidateNull(rTabela("var_fab"))
            .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("var_UnidMed"))
            .TextMatrix(.rows - 1, 7) = Format$(ValidateNull(rTabela("var_cat")), ocMONEY)
            .TextMatrix(.rows - 1, 8) = ValidateNull(rTabela("var_Local"))
            .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("var_quant"))
            .TextMatrix(.rows - 1, 10) = Format$(ValidateNull(rTabela("venda")), ocMONEY)
            '.TextMatrix(.rows - 1, 12) = Format$(ValidateNull(rTabela("custo")), ocMONEY)
            
            'VarTotalGrid = .TextMatrix(.rows - 1, 12) * .TextMatrix(.rows - 1, 9)
            '.TextMatrix(.rows - 1, 13) = Format(VarTotalGrid, ocMONEY)
            ''.TextMatrix(.rows - 1, 12) = Format$(ValidateNull(rTabela("venda")), ocMONEY)

            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If
      'lblValorTotalFiscal.Caption = Format(SomaGrid(Grid, 13), ocMONEY)
      .rows = .rows - 1
      .Redraw = True
      picAguarde.Visible = False
   End With
'Else
'End If
End Sub

Public Function SomaGrid(Grid As MSFlexGrid, Col As Integer) As Currency
   Dim i As Integer, Valor As Currency
   
   Valor = 0
   For i = 0 To Grid.rows - 1
      If IsNumeric(Grid.TextMatrix(i, Col)) Then
         Valor = Valor + CCur(Grid.TextMatrix(i, Col))
      End If
   Next
   
   SomaGrid = Valor
End Function
Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Grid_Click()
'Criado por mim
Dim i As Integer
Dim ColLimite As Integer

If optMostrarFiscal.Value = True Then
    ColLimite = 9
Else
    ColLimite = 8
End If

For i = 3 To ColLimite
   If Grid.ColSel = i Then
      txtEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth, Grid.CellHeight
      txtEdit.Text = Grid.TextMatrix(Grid.Row, Grid.Col)
      txtEdit.Visible = True
      txtEdit.SetFocus
      txtEdit.SelStart = 0
      txtEdit.SelLength = Len(txtEdit.Text)
      iRow = Grid.Row
      iCol = Grid.Col
   End If
Next
End Sub

Private Sub optCategoria_Click()
   lblCategoria.Visible = True
   cboConsLinha.Visible = True
   lblDesc.Visible = False
   cboDesc.Visible = False
   cboDesc.Visible = False
   chkDescPorProduto.Visible = False
   chkDescPorIniciais.Visible = False
   lblCodBarra.Visible = False
   txtCodBarra.Visible = False
   cmdLocalizar.Visible = True
   cboConsLinha.SetFocus
End Sub

Private Sub optCodBarra_Click()
   lblCategoria.Visible = False
   cboConsLinha.Visible = False
   lblDesc.Visible = False
   cboDesc.Visible = False
   cboDesc.Visible = False
   chkDescPorProduto.Visible = False
   chkDescPorIniciais.Visible = False
   lblCodBarra.Visible = True
   txtCodBarra.Visible = True
   cmdLocalizar.Visible = True
   txtCodBarra.SetFocus
End Sub

Private Sub optDesc_Click()
   lblCategoria.Visible = False
   cboConsLinha.Visible = False
   lblDesc.Visible = True
   cboDesc.Visible = True
   chkDescPorProduto.Visible = True
   chkDescPorIniciais.Visible = True
   lblCodBarra.Visible = False
   txtCodBarra.Visible = False
   cmdLocalizar.Visible = True
   cboDesc.SetFocus
End Sub

Private Sub optMostrarFiscal_Click()
cmdLocalizar_Click
frmTotalFiscal.Visible = True
optORDValorCusto.Visible = True
ORDQuantFiscal.Visible = True
optORDTFiscal.Visible = True
End Sub

Private Sub optMostrarNegativos_Click()
cmdLocalizar_Click
frmTotalFiscal.Visible = False
optORDValorCusto.Visible = False
ORDQuantFiscal.Visible = False
optORDTFiscal.Visible = False
End Sub

Private Sub optMostrarQuant_Click()
cmdLocalizar_Click
frmTotalFiscal.Visible = False
optORDValorCusto.Visible = False
ORDQuantFiscal.Visible = False
optORDTFiscal.Visible = False
End Sub

Private Sub optMostrarTodos_Click()
cmdLocalizar_Click
frmTotalFiscal.Visible = False
optORDValorCusto.Visible = False
ORDQuantFiscal.Visible = False
optORDTFiscal.Visible = False
End Sub

Private Sub optMostrarZerados_Click()
cmdLocalizar_Click
frmTotalFiscal.Visible = False
optORDValorCusto.Visible = False
ORDQuantFiscal.Visible = False
optORDTFiscal.Visible = False
End Sub

Private Sub optTodos_Click()
   lblCategoria.Visible = False
   cboConsLinha.Visible = False
   lblDesc.Visible = False
   cboDesc.Visible = False
   cboDesc.Visible = False
   chkDescPorProduto.Visible = False
   chkDescPorIniciais.Visible = False
   lblCodBarra.Visible = False
   txtCodBarra.Visible = False
   cmdLocalizar.Visible = False
   cmdLocalizar_Click
End Sub

Private Sub txtCodBarra_Change()
   If Len(txtCodBarra.Text) = 13 Then cmdLocalizar_Click
End Sub

Private Sub txtCodBarra_GotFocus()
   SelectControl txtCodBarra
End Sub

Private Sub txtEdit_GotFocus()
'criado por IA
' Registra as coordenadas ONDE o editor nasceu.
' O LostFocus usará essas variáveis para saber onde salvar o valor.
iRow = Grid.Row
iCol = Grid.Col
End Sub

Private Sub txtEdit_KeyUp(KeyCode As Integer, Shift As Integer)
'criado pela IA
' Captura a linha e coluna ATUAIS onde o txtEdit está posicionado
  Dim r As Long, c As Long
  r = Grid.Row
  c = Grid.Col

  If KeyCode = 38 Then ' Seta para CIMA
     If r > 1 Then ' Evita subir além do cabeçalho
        ' 1. Salva o valor atual do texto na célula antes de sair dela
        Grid.TextMatrix(r, c) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
        
        ' 2. Move para a linha de cima e reativa o editor
        Grid.Row = r - 1
        Grid_Click
     Else
        MsgBox "VOCÊ JÁ ESTÁ NA PRIMEIRA LINHA !!!", vbExclamation
     End If
  
  ElseIf KeyCode = 40 Then ' Seta para BAIXO
     If r < Grid.rows - 1 Then
        ' 1. Salva o valor atual do texto na célula
        Grid.TextMatrix(r, c) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
        
        ' 2. Move para a linha de baixo e reativa o editor
        Grid.Row = r + 1
        Grid_Click
     Else
        MsgBox "VOCÊ JÁ ESTÁ NA ULTIMA LINHA !!!", vbExclamation
     End If
  End If
'criado por mim
 'If KeyCode = 38 Then
 '  If Grid.Row - 1 = 0 Then ShowMsg "VOCÊ JÁ ESTÁ NA PRIMEIRA LINHA !!!", vbExclamation: Exit Sub
 '  Grid.Row = iRow - 1
 '  Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
 '  Grid_Click

'ElseIf KeyCode = 40 Then
'   If Grid.rows = Grid.Row + 1 Then ShowMsg "VOCÊ JÁ ESTÁ NA ULTIMA LINHA !!!", vbExclamation: Exit Sub
'   Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
'   Grid.Row = iRow + 1
'   Grid_Click
'End If
End Sub

Private Sub txtEdit_LostFocus()
'criado por mim
If iCol = 6 Then
    txtEdit.Text = Replace(txtEdit.Text, ".", "")
    txtEdit.Text = Trim(txtEdit.Text)
End If

Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
txtEdit.Visible = False
End Sub



Private Sub txtSenha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdSenha_Click
End Sub


