VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Etiquetas_Impressao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IMPRESSÃO DE ETIQUETAS"
   ClientHeight    =   10035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15180
   Icon            =   "Etiquetas_Impressao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Etiquetas_Impressao.frx":1D82
   ScaleHeight     =   10035
   ScaleWidth      =   15180
   StartUpPosition =   2  'CenterScreen
   Begin ChamaleonBtn.chameleonButton cmdDesmarcar 
      Height          =   315
      Left            =   11640
      TabIndex        =   44
      Top             =   7680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Desmarcar Todos"
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
      MICON           =   "Etiquetas_Impressao.frx":264C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tipos de Códigos"
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
      Left            =   3300
      TabIndex        =   38
      Top             =   8040
      Width           =   1815
      Begin VB.OptionButton optTCCriados 
         Caption         =   "Criados"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.OptionButton optTCProprios 
         Caption         =   "Próprios"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   900
         Width           =   1035
      End
      Begin VB.OptionButton optTCTodos 
         Caption         =   "Todos"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Quantidades"
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
      Left            =   1560
      TabIndex        =   33
      Top             =   8040
      Width           =   1695
      Begin VB.OptionButton optMostrarQuant 
         Caption         =   "Com quantidade"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Width           =   1515
      End
      Begin VB.OptionButton optMostrarNegativos 
         Caption         =   "Negativos"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   900
         Width           =   1035
      End
      Begin VB.OptionButton optMostrarZerados 
         Caption         =   "Zerados"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   1200
         Width           =   915
      End
      Begin VB.OptionButton optMostrarTodos 
         Caption         =   "Todos"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   300
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.PictureBox picUnchecked 
      Height          =   180
      Left            =   8760
      Picture         =   "Etiquetas_Impressao.frx":2668
      ScaleHeight     =   120
      ScaleWidth      =   120
      TabIndex        =   32
      Top             =   1920
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picChecked 
      Height          =   180
      Left            =   8760
      Picture         =   "Etiquetas_Impressao.frx":274A
      ScaleHeight     =   120
      ScaleWidth      =   120
      TabIndex        =   31
      Top             =   1560
      Visible         =   0   'False
      Width           =   180
   End
   Begin ChamaleonBtn.chameleonButton cmdCancelar 
      Height          =   315
      Left            =   11640
      TabIndex        =   30
      Top             =   7680
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Cancelar"
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
      MICON           =   "Etiquetas_Impressao.frx":282C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdImprimirEtiqueta 
      Height          =   315
      Left            =   9900
      TabIndex        =   29
      Top             =   7680
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Imprimir Etiqueta"
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
      MICON           =   "Etiquetas_Impressao.frx":2848
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdGerarEtiquetas 
      Height          =   315
      Left            =   13380
      TabIndex        =   28
      Top             =   8040
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Gerar Etiquetas"
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
      MICON           =   "Etiquetas_Impressao.frx":2864
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      TabIndex        =   17
      Top             =   8040
      Width           =   1395
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   300
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.OptionButton optCategoria 
         Caption         =   "Categoria"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   1200
         Width           =   1035
      End
      Begin VB.OptionButton optDesc 
         Caption         =   "Descrição"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   900
         Width           =   1035
      End
      Begin VB.OptionButton optCodBarra 
         Caption         =   "Cód. Barra"
         Height          =   195
         Left            =   120
         TabIndex        =   18
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
      Left            =   6720
      TabIndex        =   11
      Top             =   8040
      Width           =   5085
      Begin VB.OptionButton optPalavrasDuplas 
         Caption         =   "Palavras Duplas"
         Height          =   195
         Left            =   1080
         TabIndex        =   15
         Top             =   1080
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.OptionButton optPorPalavra 
         Caption         =   "Palavra"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox txtCodBarra 
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.ComboBox cboConsLinha 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.ComboBox cboDesc 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Visible         =   0   'False
         Width           =   4815
      End
      Begin ChamaleonBtn.chameleonButton cmdLocalizar 
         Height          =   495
         Left            =   3480
         TabIndex        =   26
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
         MICON           =   "Etiquetas_Impressao.frx":2880
         PICN            =   "Etiquetas_Impressao.frx":289C
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
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblCategoria 
         Caption         =   "Categoria"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblDesc 
         Caption         =   "Descrição"
         Height          =   195
         Left            =   180
         TabIndex        =   22
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
      Left            =   5160
      TabIndex        =   5
      Top             =   8040
      Width           =   1515
      Begin VB.CheckBox ckkORDDesc 
         Caption         =   "Descrição"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   300
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox ckkORDLinha 
         Caption         =   "Categoria"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1260
         Width           =   1215
      End
      Begin VB.CheckBox ckkORDQuantMin 
         Caption         =   "Quant. Min."
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   780
         Width           =   1215
      End
      Begin VB.CheckBox ckkORDQuant 
         Caption         =   "Quant."
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   540
         Width           =   1335
      End
      Begin VB.CheckBox ckkORDValor 
         Caption         =   "Valor"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1020
         Width           =   975
      End
   End
   Begin VB.PictureBox picAguarde 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   6300
      Picture         =   "Etiquetas_Impressao.frx":3176
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
      ScaleWidth      =   14985
      TabIndex        =   0
      Top             =   60
      Width           =   15015
      Begin VB.Image Image1 
         Height          =   750
         Left            =   480
         Picture         =   "Etiquetas_Impressao.frx":41AE
         Top             =   60
         Width           =   690
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "IMPRESSÃO DE ETIQUETAS"
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
         Width           =   4245
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdSair 
      Height          =   315
      Left            =   13380
      TabIndex        =   14
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
      MICON           =   "Etiquetas_Impressao.frx":5617
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
      TabIndex        =   27
      Top             =   9765
      Width           =   15180
      _ExtentX        =   26776
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22437
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "09:30"
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
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   11668
      _Version        =   393216
      Cols            =   5
      AllowBigSelection=   0   'False
      HighLight       =   0
      ScrollBars      =   2
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Label lblQuantItens 
      AutoSize        =   -1  'True
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   660
      TabIndex        =   46
      Top             =   7680
      Width           =   330
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Itens:"
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
      Left            =   60
      TabIndex        =   45
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Etiquetas:"
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
      Left            =   1140
      TabIndex        =   43
      Top             =   7680
      Width           =   870
   End
   Begin VB.Label lblQuantSelecionada 
      AutoSize        =   -1  'True
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   2100
      TabIndex        =   42
      Top             =   7680
      Width           =   330
   End
End
Attribute VB_Name = "Etiquetas_Impressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moCombo As cComboHelper
Private iRow As Long, iCol As Long
Private editandoEtiqueta As Boolean
Private chamandoImpressao As Boolean
Private ExibirGidEtiquetas As Boolean
Private bSuprimirBuscaEtq As Boolean   'evita cascata de cmdLocalizar_Click ao setar optTC*/optMostrar* por codigo
Dim sSQL As String
Dim r As ADODB.Recordset
Dim xProdutosSelecionados As String
Dim xQtdImpressao As String   ' qtde de etiquetas por produto marcado, formato "|cod=qtd|cod=qtd"

Private Sub LimparGrid2()
Dim sSQL As String
Dim r As ADODB.Recordset
   
sSQL = "SELECT  produtos.NCM AS var_NCM, produtos.CFOP AS var_CFOP, produtos.ICMSCST AS var_ICMS, produtos.categoria AS var_cat, produtos.fabricante AS var_fab, " & _
   "produtos.codigo AS var_cod, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, produtos.quant_estoque AS var_quant, produtos.UNID_MEDIDA AS var_UnidMed, " & _
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



Private Sub MostrarCriterios()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim var_Criterio As String
   Dim var_Indice As String
   
   var_Criterio = ""
   
   If optDesc.Value = True Then
      Dim vDescFiltro As String
      Dim aPalDesc() As String
      Dim iPalDesc As Integer
      vDescFiltro = ""
      If optPorPalavra.Value = True Then
         vDescFiltro = "produtos.descricao COLLATE Latin1_General_CI_AI LIKE '%" & cboDesc.Text & "%'"
      ElseIf optPalavrasDuplas.Value = True Then
         aPalDesc = Split(Trim(cboDesc.Text), " ")
         For iPalDesc = 0 To UBound(aPalDesc)
            If Trim(aPalDesc(iPalDesc)) <> "" Then
               If vDescFiltro <> "" Then vDescFiltro = vDescFiltro & " AND "
               vDescFiltro = vDescFiltro & "produtos.descricao COLLATE Latin1_General_CI_AI LIKE '%" & Trim(aPalDesc(iPalDesc)) & "%'"
            End If
         Next iPalDesc
      End If
      If vDescFiltro <> "" Then var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "(" & vDescFiltro & ")"
   End If
   
   var_Criterio = var_Criterio & IIf(optCategoria.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.categoria = '" & cboConsLinha.Text & "'", "")
   var_Criterio = var_Criterio & IIf(optCodBarra.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.cod_barra = '" & txtCodBarra.Text & "'", "")
   
   If var_Criterio <> "" Then var_Criterio = " WHERE " & var_Criterio
   
   var_Indice = ""
   var_Indice = var_Indice & IIf(ckkORDQuant.Value, IIf(var_Indice <> "", ", ", "") & "quant_estoque", "")
   var_Indice = var_Indice & IIf(ckkORDDesc.Value, IIf(var_Indice <> "", ", ", "") & "produtos.descricao", "")
   var_Indice = var_Indice & IIf(ckkORDQuantMin.Value, IIf(var_Indice <> "", ", ", "") & "quant_min", "")
   var_Indice = var_Indice & IIf(ckkORDValor.Value, IIf(var_Indice <> "", ", ", "") & "produtos_entrada_itens.venda", "")
   var_Indice = var_Indice & IIf(ckkORDLinha.Value, IIf(var_Indice <> "", ", ", "") & "produtos.categoria", "")
   
   If var_Indice <> "" Then var_Indice = " ORDER BY " & var_Indice

    Dim varTipoCodigo As String
    
    If optTCCriados.Value = True Then
        varTipoCodigo = " AND len(produtos.cod_barra) <= 6"   '6 chars = codigo criado (interno passou de 99999); > 6 = codigo proprio. sem fresta.
    ElseIf optTCProprios.Value = True Then
        varTipoCodigo = " AND len(produtos.cod_barra) > 6"
    ElseIf optTCTodos.Value = True Then
        varTipoCodigo = " "
    End If
    If optDesc.Value = True Then varTipoCodigo = " "   'busca por descricao ignora Tipos de Codigos
   
   sSQL = "SELECT produtos.NCM AS var_NCM, produtos.ICMSCST AS var_ICMS, produtos.CFOP AS var_CFOP, produtos.categoria AS var_cat, produtos.fabricante AS var_fab, " & _
      "produtos.codigo AS var_cod, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, produtos.ref AS var_referencia, produtos.quant_estoque AS var_quant, produtos.UNID_MEDIDA AS var_UnidMed, " & _
      "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
      "FROM produtos " & var_Criterio & " " & varTipoCodigo & " " & var_Indice
   
   Set r = dbData.OpenRecordset(sSQL)
   
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

Private Sub ckkORDDesc_Click()
   MostrarCriterios
End Sub

Private Sub ckkORDLinha_Click()
   MostrarCriterios
End Sub

Private Sub ckkORDQuant_Click()
   MostrarCriterios
End Sub

Private Sub ckkORDQuantMin_Click()
   MostrarCriterios
End Sub

Private Sub ckkORDValor_Click()
   MostrarCriterios
End Sub


Private Sub cmdDesmarcar_Click()
    'unica forma (alem de fechar a janela) de zerar a selecao e as quantidades
    xProdutosSelecionados = ""
    xQtdImpressao = ""
    If editandoEtiqueta Then
        Dim k As Integer
        For k = 1 To Grid.rows - 1
            Grid.Col = 3: Grid.Row = k
            Set Grid.CellPicture = picUnchecked.Picture
            Grid.TextMatrix(k, 4) = ""
        Next k
    End If
    AtualizarLabelsSelecao
End Sub

Private Sub cmdCancelar_Click()
ExibirGidEtiquetas = False

'volta pro estado inicial (botao Gerar Etiquetas visivel) - antes so trocava o grid e
'deixava Imprimir/Cancelar na tela sem caminho de volta pro modo etiqueta
cmdGerarEtiquetas.Visible = True
cmdImprimirEtiqueta.Visible = False
cmdCancelar.Visible = False
cmdImprimirEtiqueta.Top = 8040
cmdCancelar.Top = 8040
cmdSair.Top = 8040

cmdLocalizar_Click
End Sub

Private Sub cmdGerarEtiquetas_Click()
    ExibirGidEtiquetas = True

    cmdGerarEtiquetas.Visible = False
    cmdImprimirEtiqueta.Visible = True
    cmdCancelar.Visible = True
    cmdImprimirEtiqueta.Top = 7680
    cmdCancelar.Top = 7680
    cmdSair.Top = 7680
    
    'Carregar GRID
    Dim varTipoMostrar As String
    
    If optMostrarQuant.Value = True Then
        varTipoMostrar = " AND produtos.quant_estoque > 0"
    ElseIf optMostrarNegativos.Value = True Then
        varTipoMostrar = " AND produtos.quant_estoque < 0"
    ElseIf optMostrarZerados.Value = True Then
        varTipoMostrar = " AND produtos.quant_estoque = 0"
    ElseIf optMostrarTodos.Value = True Then
        varTipoMostrar = " "
    End If

    Dim varTipoCodigo As String
    
    If optTCCriados.Value = True Then
        varTipoCodigo = " AND len(produtos.cod_barra) <= 6"   '6 chars = codigo criado (interno passou de 99999); > 6 = codigo proprio. sem fresta.
    ElseIf optTCProprios.Value = True Then
        varTipoCodigo = " AND len(produtos.cod_barra) > 6"
    ElseIf optTCTodos.Value = True Then
        varTipoCodigo = " "
    End If

    
    If optTodos.Value = True Then
       sSQL = "SELECT  produtos.NCM AS var_NCM, produtos.CFOP AS var_CFOP, produtos.ICMSCST AS var_ICMS, produtos.categoria AS var_cat, produtos.fabricante AS var_fab, " & _
          "produtos.codigo AS var_cod, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, produtos.REF AS var_referencia, produtos.quant_estoque AS var_quant, produtos.UNID_MEDIDA AS var_UnidMed, " & _
          "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
          "FROM produtos " & _
          "WHERE (produtos.ativo = 1) " & varTipoMostrar & " " & varTipoCodigo & " ORDER BY produtos.descricao;"
    
       Set r = dbData.OpenRecordset(sSQL)
       
        If r.RecordCount > 32000 Then
            MsgBox "A Consulta retornou um valor maior de registros que é permitido na grade!", vbInformation, "Aviso do sistema"
            LimparGrid2
            Exit Sub
        Else
            Formatar_Grid_Etiquetas r
        End If
       
       If r.State <> 0 Then r.Close
       Set r = Nothing
       
    Else
       MostrarCriterios
    End If
    
    If optCodBarra.Value = True Then txtCodBarra_GotFocus
    
End Sub
Private Sub cmdImprimirEtiqueta_Click()
    'monta a impressao a partir de xProdutosSelecionados (tudo que foi marcado), consultando
    'o cadastro pelos codigos - independe do filtro que esta no grid nesse momento.
    Dim objB As Object
    Set objB = CreateObject("ImpressaoDeEtiquetas.GeradorDeEtiquetas")

    Dim sLista As String
    sLista = SelValor(xProdutosSelecionados)   ' devolve " 123, 456" (sem a virgula final)
    If Trim(sLista) = "" Then
        Set objB = Nothing
        MsgBox "Marque ao menos um produto para imprimir a etiqueta.", vbExclamation, "Etiquetas"
        Exit Sub
    End If

    chamandoImpressao = True

    Dim aCods() As String
    aCods = Split(Replace(sLista, " ", ""), ",")

    Dim arrayDeDados()
    ReDim arrayDeDados(0 To UBound(aCods), 0 To 4)

    Dim rP As ADODB.Recordset
    Dim sSQLImp As String
    sSQLImp = "SELECT p.codigo, p.cod_barra, p.descricao, p.fabricante, p.quant_estoque, " & _
              "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos WHERE Produtos_Precos.COD_PRODUTO = p.codigo ORDER BY CODIGO DESC) AS venda " & _
              "FROM produtos p WHERE p.codigo IN (" & Replace(sLista, " ", "") & ")"
    Set rP = dbData.OpenRecordset(sSQLImp)

    Dim idx As Long
    Dim dVenda As Double
    idx = 0
    Do While Not rP.EOF
        arrayDeDados(idx, 0) = GetQtdImpr(CStr(rP("codigo")))
        arrayDeDados(idx, 1) = ValidateNull(rP("descricao")) & " / " & ValidateNull(rP("fabricante"))
        arrayDeDados(idx, 2) = ValidateNull(rP("cod_barra"))
        arrayDeDados(idx, 3) = ValidateNull(rP("quant_estoque"))
        If IsNull(rP("venda")) Then dVenda = 0 Else dVenda = rP("venda")
        arrayDeDados(idx, 4) = CDec(dVenda)
        idx = idx + 1
        rP.MoveNext
    Loop
    If rP.State <> 0 Then rP.Close
    Set rP = Nothing

    If idx = 0 Then
        chamandoImpressao = False
        Set objB = Nothing
        MsgBox "Nenhum dos produtos marcados foi encontrado no cadastro.", vbExclamation, "Etiquetas"
        Exit Sub
    End If

    ReDimPreserve arrayDeDados, idx - 1, UBound(arrayDeDados, 2)

    objB.ExibirModalConfiguracaoImpressao (arrayDeDados)
    chamandoImpressao = False
End Sub

Private Function GetQtdImpr(ByVal sCod As String) As Long
    'qtde de etiquetas do produto, persistida no mapa xQtdImpressao; default 1
    Dim p1 As Long, p2 As Long, sVal As String
    GetQtdImpr = 1
    p1 = InStr(1, xQtdImpressao, "|" & sCod & "=")
    If p1 = 0 Then Exit Function
    p1 = p1 + Len(sCod) + 2
    p2 = InStr(p1, xQtdImpressao, "|")
    If p2 = 0 Then p2 = Len(xQtdImpressao) + 1
    sVal = Mid$(xQtdImpressao, p1, p2 - p1)
    If IsNumeric(sVal) Then GetQtdImpr = CLng(sVal)
    If GetQtdImpr < 1 Then GetQtdImpr = 1
End Function

Private Sub RemoveQtdImpr(ByVal sCod As String)
    Dim p1 As Long, p2 As Long
    p1 = InStr(1, xQtdImpressao, "|" & sCod & "=")
    If p1 = 0 Then Exit Sub
    p2 = InStr(p1 + 1, xQtdImpressao, "|")
    If p2 = 0 Then p2 = Len(xQtdImpressao) + 1
    xQtdImpressao = Left$(xQtdImpressao, p1 - 1) & Mid$(xQtdImpressao, p2)
End Sub

Private Sub SetQtdImpr(ByVal sCod As String, ByVal nQtd As Long)
    RemoveQtdImpr sCod
    If nQtd < 1 Then nQtd = 1
    xQtdImpressao = xQtdImpressao & "|" & sCod & "=" & nQtd
End Sub

Private Sub AtualizarQtdDaLinha(ByVal nRow As Long)
    'chamado quando o usuario edita a coluna IMPRESSOES - so grava se a linha estiver marcada
    If nRow < 1 Or nRow > Grid.rows - 1 Then Exit Sub
    Dim sCod As String
    sCod = Grid.TextMatrix(nRow, 2)
    If Not SelProcuraValor(xProdutosSelecionados, sCod) Then Exit Sub
    If IsNumeric(Grid.TextMatrix(nRow, 4)) Then SetQtdImpr sCod, CLng(Grid.TextMatrix(nRow, 4))
    AtualizarLabelsSelecao
End Sub

Private Function ContarItensMarcados() As Long
    'qtos itens da lista estao marcados (independe do grid) = qtde de entradas em xQtdImpressao
    ContarItensMarcados = Len(xQtdImpressao) - Len(Replace(xQtdImpressao, "|", ""))
End Function

Private Function SomarQtdImpressoes() As Long
    'soma das qtdes de impressao de TODOS os marcados (independe do grid)
    Dim aEnt() As String, k As Long, p As Long, sQ As String
    SomarQtdImpressoes = 0
    If xQtdImpressao = "" Then Exit Function
    aEnt = Split(xQtdImpressao, "|")
    For k = 0 To UBound(aEnt)
        If aEnt(k) <> "" Then
            p = InStr(aEnt(k), "=")
            If p > 0 Then
                sQ = Mid$(aEnt(k), p + 1)
                If IsNumeric(sQ) Then SomarQtdImpressoes = SomarQtdImpressoes + CLng(sQ)
            End If
        End If
    Next k
End Function

Private Sub AtualizarLabelsSelecao()
    lblQuantItens.Caption = ContarItensMarcados()
    lblQuantSelecionada.Caption = SomarQtdImpressoes()
End Sub


Public Sub ReDimPreserve(ByRef arr, ByVal size1 As Long, ByVal size2 As Long)
Dim arr2 As Variant
Dim x As Long, y As Long
Dim maxX As Long, maxY As Long

'Check if it's an array first
If Not IsArray(arr) Then Exit Sub

'create new array with initial start
ReDim arr2(LBound(arr, 1) To size1, LBound(arr, 2) To size2)

'copia so a regiao que existe nos DOIS arrays - funciona pra crescer E pra encolher.
'antes o loop ia ate UBound(arr) e estourava (erro 9) quando arr2 era menor que arr.
maxX = UBound(arr, 1): If size1 < maxX Then maxX = size1
maxY = UBound(arr, 2): If size2 < maxY Then maxY = size2

For x = LBound(arr, 1) To maxX
    For y = LBound(arr, 2) To maxY
        arr2(x, y) = arr(x, y)
    Next
Next
'return byref
arr = arr2
End Sub



Private Sub cmdLocalizar_Click()
If bSuprimirBuscaEtq Then Exit Sub
Dim varTipoMostrar As String

If optMostrarQuant.Value = True Then
    varTipoMostrar = " AND produtos.quant_estoque > 0"
ElseIf optMostrarNegativos.Value = True Then
    varTipoMostrar = " AND produtos.quant_estoque < 0"
ElseIf optMostrarZerados.Value = True Then
    varTipoMostrar = " AND produtos.quant_estoque = 0"
ElseIf optMostrarTodos.Value = True Then
    varTipoMostrar = " "
End If

Dim varTipoCodigo As String

If optTCCriados.Value = True Then
    varTipoCodigo = " AND len(produtos.cod_barra) <= 6"   '6 chars = codigo criado (interno passou de 99999); > 6 = codigo proprio. sem fresta.
ElseIf optTCProprios.Value = True Then
    varTipoCodigo = " AND len(produtos.cod_barra) > 6"
ElseIf optTCTodos.Value = True Then
    varTipoCodigo = " "
End If


If optTodos.Value = True Then
   sSQL = "SELECT  produtos.NCM AS var_NCM, produtos.CFOP AS var_CFOP, produtos.ICMSCST AS var_ICMS, produtos.categoria AS var_cat, produtos.fabricante AS var_fab, " & _
      "produtos.codigo AS var_cod, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc,produtos.ref AS var_referencia, produtos.quant_estoque AS var_quant, produtos.UNID_MEDIDA AS var_UnidMed, " & _
      "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
      "FROM produtos " & _
      "WHERE (produtos.ativo = 1) " & varTipoMostrar & " " & varTipoCodigo & " ORDER BY produtos.descricao;"

   Set r = dbData.OpenRecordset(sSQL)
   
    If r.RecordCount > 32000 Then
        MsgBox "A Consulta retornou um valor maior de registros que é permitido na grade!", vbInformation, "Aviso do sistema"
        LimparGrid2
        Exit Sub
    Else
        Formatar_Grid r
    End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
Else
   MostrarCriterios
End If

If optCodBarra.Value = True Then txtCodBarra_GotFocus
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Public Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Currency
   Dim i As Integer, Valor As Currency
   
   Valor = 0
   For i = 0 To var_Grid.rows - 1
      If IsNumeric(var_Grid.TextMatrix(i, Col)) Then
         Valor = Valor + CDbl(var_Grid.TextMatrix(i, Col))
      End If
   Next
   
   SomaGrid = Valor
End Function

Private Sub Form_Load()
Set moCombo = New cComboHelper
editandoEtiqueta = False
cmdGerarEtiquetas_Click
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
   'cboDesc e texto livre: nao preenche lista de produtos nem faz autocomplete.
   cboDesc.Clear
End Sub

Private Sub cboDesc_LostFocus()
   'cboDesc_Click
End Sub

Private Sub cboConsLinha_Click()
   'If cboConsLinha.Text <> "" Then MostrarCriterios
End Sub

Private Sub cboConsLinha_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboConsLinha.Clear
   
   sSQL = "SELECT DISTINCT categoria FROM produtos ORDER BY categoria;"
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

Private Sub Formatar_Grid(rTabela As ADODB.Recordset)
   If ExibirGidEtiquetas Then
       Formatar_Grid_Etiquetas rTabela
       Exit Sub
   End If
   
   Dim i As Integer
   
   
   LimparGrid
   picAguarde.Visible = True
   DoEvents
   
   With Grid
      .Clear
      .Cols = 13
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 0
      .ColWidth(3) = 1500
      .ColWidth(4) = 4200
      .ColWidth(5) = 1600
      .ColWidth(6) = 1200
      .ColWidth(7) = 800
      .ColWidth(8) = 800
      .ColWidth(9) = 800
      .ColWidth(10) = 1850
      .ColWidth(11) = 800
      .ColWidth(12) = 900
      
      '.RowHeight(-1) = (315 * 1)    'definir a altura da linha
      
      .TextMatrix(0, 1) = "CÓD.ENT"
      .TextMatrix(0, 2) = "CÓD.PROD"
      .TextMatrix(0, 3) = "CÓD.BARRA"
      .TextMatrix(0, 4) = "DESCRIÇÃO"
      .TextMatrix(0, 5) = "FABRICANTE"
      .TextMatrix(0, 6) = "NCM."
      .TextMatrix(0, 7) = "CFOP."
      .TextMatrix(0, 8) = "ICMS."
      .TextMatrix(0, 9) = "MED."
      .TextMatrix(0, 10) = "CATEGORIA"
      .TextMatrix(0, 11) = "QUANT."
      .TextMatrix(0, 12) = "VENDA"
      
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
            
            '.TextMatrix(.Rows - 1, 1) = ValidateNull(rTabela("var_codent"))
             If SelProcuraValor(xProdutosSelecionados, rTabela("var_cod")) Then
                Set Grid.CellPicture = picChecked
             End If
            .TextMatrix(.rows - 1, 2) = ValidateNull(rTabela("var_cod"))
            .TextMatrix(.rows - 1, 3) = Format(ValidateNull(rTabela("var_codbarra")), "0000000000000")
            .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("var_desc"))
            .TextMatrix(.rows - 1, 5) = ValidateNull(rTabela("var_fab"))
            .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("var_NCM"))
            .TextMatrix(.rows - 1, 7) = ValidateNull(rTabela("var_CFOP"))
            .TextMatrix(.rows - 1, 8) = ValidateNull(rTabela("var_ICMS"))
            .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("var_UnidMed"))
            .TextMatrix(.rows - 1, 10) = ValidateNull(rTabela("var_cat"))
            .TextMatrix(.rows - 1, 11) = ValidateNull(rTabela("var_quant"))
            .TextMatrix(.rows - 1, 12) = Format$(ValidateNull(rTabela("venda")), ocMONEY)
            
            '.TextMatrix(.Rows - 1, 9) = ValidateNull(rTabela("var_quant"))
            '.TextMatrix(.Rows - 1, 11) = Format$(ValidateNull(rTabela("venda")), ocMONEY)
            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If
      
      .rows = .rows - 1
      .Redraw = True
      picAguarde.Visible = False
      AtualizarLabelsSelecao
   End With
   
   editandoEtiqueta = False
End Sub

Private Sub Formatar_Grid_Etiquetas(rTabela As ADODB.Recordset)
Dim i As Integer

LimparGrid
picAguarde.Visible = True
DoEvents

With Grid
   .Clear
   .Cols = 16
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 0
   .ColWidth(2) = 0
   
   'Colunas de checkbox
   .ColWidth(3) = 800
   .ColWidth(4) = 1500
   
   .ColWidth(5) = 1500
   .ColWidth(6) = 6000
   .ColWidth(7) = 0
   .ColWidth(8) = 0
   .ColWidth(9) = 0
   .ColWidth(10) = 0
   .ColWidth(11) = 0
   .ColWidth(12) = 0
   .ColWidth(13) = 2000
   .ColWidth(14) = 800
   .ColWidth(15) = 2220
   
   '.RowHeight(-1) = (315 * 1)    'definir a altura da linha
   
   .TextMatrix(0, 1) = "CÓD.ENT"
   .TextMatrix(0, 2) = "CÓD.PROD"
   
   'Colunas de checkbox
   .TextMatrix(0, 3) = " "
   .TextMatrix(0, 4) = "IMPRESSÕES"
   
   
   'Avançar duas colunas para incluir coluna de checkbox e coluna de textField
   .TextMatrix(0, 5) = "CÓD.BARRA"
   .TextMatrix(0, 6) = "DESCRIÇÃO"
   .TextMatrix(0, 7) = "FABRICANTE"
   .TextMatrix(0, 8) = "NCM."
   .TextMatrix(0, 9) = "CFOP."
   .TextMatrix(0, 10) = "ICMS."
   .TextMatrix(0, 11) = "MED."
   .TextMatrix(0, 12) = "CATEGORIA"
   .TextMatrix(0, 13) = "REFERÊNCIA"
   .TextMatrix(0, 14) = "QUANT."
   .TextMatrix(0, 15) = "VENDA"
   
   
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

         '.TextMatrix(.Rows - 1, 1) = ValidateNull(rTabela("var_codent"))
         .TextMatrix(.rows - 1, 2) = ValidateNull(rTabela("var_cod"))
         .TextMatrix(.rows - 1, 5) = ValidateNull(rTabela("var_codbarra"))
         .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("var_desc")) & " / " & ValidateNull(rTabela("var_fab"))
         .TextMatrix(.rows - 1, 7) = ValidateNull(rTabela("var_fab"))
         .TextMatrix(.rows - 1, 8) = ValidateNull(rTabela("var_NCM"))
         .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("var_CFOP"))
         .TextMatrix(.rows - 1, 10) = ValidateNull(rTabela("var_ICMS"))
         .TextMatrix(.rows - 1, 11) = ValidateNull(rTabela("var_UnidMed"))
         .TextMatrix(.rows - 1, 12) = ValidateNull(rTabela("var_cat"))
         .TextMatrix(.rows - 1, 13) = ValidateNull(rTabela("var_referencia"))
         .TextMatrix(.rows - 1, 14) = ValidateNull(rTabela("var_quant"))
         .TextMatrix(.rows - 1, 15) = Format$(ValidateNull(rTabela("venda")), ocMONEY)
         
         '.TextMatrix(.Rows - 1, 9) = ValidateNull(rTabela("var_quant"))
         '.TextMatrix(.Rows - 1, 11) = Format$(ValidateNull(rTabela("venda")), ocMONEY)
         rTabela.MoveNext
         .rows = .rows + 1
      Loop
   End If
         
   For i = 1 To .rows - 1
      .Col = 3
      .Row = i
      If SelProcuraValor(xProdutosSelecionados, .TextMatrix(i, 2)) Then
         Set .CellPicture = picChecked.Picture
         .TextMatrix(i, 4) = GetQtdImpr(.TextMatrix(i, 2))
      Else
         Set .CellPicture = picUnchecked.Picture
      End If
      .Row = i: .Col = 3: .CellPictureAlignment = 4 ' Align the checkbox
   Next
   
   
   .rows = .rows - 1
   .Redraw = True
   picAguarde.Visible = False
   AtualizarLabelsSelecao
End With

editandoEtiqueta = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Grid_Click()
Dim i As Integer

'so a coluna 4 (IMPRESSOES, so no modo etiqueta) e editavel no grid.
'antes abria o editor pras colunas 3..10 (cod.barra, descricao, fabricante...) e dava
'pra sobrescrever no grid o que ia pra etiqueta.
If editandoEtiqueta And Grid.Col = 4 And SelProcuraValor(xProdutosSelecionados, Grid.TextMatrix(Grid.Row, 2)) Then
    txtEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth, Grid.CellHeight
    txtEdit.Text = Grid.TextMatrix(Grid.Row, Grid.Col)
    txtEdit.Visible = True
    txtEdit.SetFocus
    txtEdit.SelStart = 0
    txtEdit.SelLength = Len(txtEdit.Text)
    iRow = Grid.Row
    iCol = Grid.Col
End If

'If Grid.Col = 4 Then
 '   Grid.Col = 3

'    For i = 3 To 10
'       If Grid.ColSel = i Then
'          txtEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth, Grid.CellHeight
'          txtEdit.Text = Grid.TextMatrix(Grid.Row, Grid.Col)
'          txtEdit.Visible = True
'          txtEdit.SetFocus
'          txtEdit.SelStart = 0
'          txtEdit.SelLength = Len(txtEdit.Text)
'          iRow = Grid.Row
'          iCol = Grid.Col
'       End If
'    Next
'End If

    If Grid.Col = 3 And editandoEtiqueta Then
        SelAdicionaValor xProdutosSelecionados, Grid.TextMatrix(Grid.Row, 2)
        If Grid.CellPicture = picChecked Then
            Set Grid.CellPicture = picUnchecked
            Grid.TextMatrix(Grid.Row, 4) = ""   'senao SomaGrid ainda conta esta linha
            RemoveQtdImpr Grid.TextMatrix(Grid.Row, 2)
            AtualizarLabelsSelecao
        Else
            Set Grid.CellPicture = picChecked
            Grid.TextMatrix(Grid.Row, 4) = 1
            SetQtdImpr Grid.TextMatrix(Grid.Row, 2), 1
            AtualizarLabelsSelecao
        End If
    End If
End Sub

Private Sub Grid_LeaveCell()
    If Not editandoEtiqueta Or chamandoImpressao Then Exit Sub
    If Grid.Col = 4 And txtEdit.Text <> "" And Not IsNumeric(txtEdit.Text) Then
        txtEdit.Text = 1
    End If
End Sub

Private Sub optCategoria_Click()
   bSuprimirBuscaEtq = True
   optTCTodos.Value = True
   optMostrarTodos.Value = True
   bSuprimirBuscaEtq = False
   lblCategoria.Visible = True
   cboConsLinha.Visible = True
   lblDesc.Visible = False
   cboDesc.Visible = False
   cboDesc.Visible = False
   optPorPalavra.Visible = False
   optPalavrasDuplas.Visible = False
   lblCodBarra.Visible = False
   txtCodBarra.Visible = False
   cmdLocalizar.Visible = True
   cboConsLinha.SetFocus
End Sub

Private Sub optCodBarra_Click()
   bSuprimirBuscaEtq = True
   optTCTodos.Value = True
   optMostrarTodos.Value = True
   bSuprimirBuscaEtq = False
   lblCategoria.Visible = False
   cboConsLinha.Visible = False
   lblDesc.Visible = False
   cboDesc.Visible = False
   cboDesc.Visible = False
   optPorPalavra.Visible = False
   optPalavrasDuplas.Visible = False
   lblCodBarra.Visible = True
   txtCodBarra.Visible = True
   cmdLocalizar.Visible = True
   txtCodBarra.SetFocus
End Sub

Private Sub optDesc_Click()
   bSuprimirBuscaEtq = True
   optTCTodos.Value = True
   optMostrarTodos.Value = True
   bSuprimirBuscaEtq = False
   lblCategoria.Visible = False
   cboConsLinha.Visible = False
   lblDesc.Visible = True
   cboDesc.Visible = True
   optPorPalavra.Visible = True
   optPalavrasDuplas.Visible = True
   optPorPalavra.Value = True
   lblCodBarra.Visible = False
   txtCodBarra.Visible = False
   cmdLocalizar.Visible = True
   cboDesc.SetFocus
End Sub

Private Sub optMostrarNegativos_Click()
cmdLocalizar_Click
End Sub

Private Sub optMostrarQuant_Click()
cmdLocalizar_Click
End Sub

Private Sub optMostrarTodos_Click()
cmdLocalizar_Click
End Sub

Private Sub optMostrarZerados_Click()
cmdLocalizar_Click
End Sub

Private Sub optTCCriados_Click()
cmdLocalizar_Click
End Sub

Private Sub optTCProprios_Click()
cmdLocalizar_Click
End Sub


Private Sub optTCTodos_Click()
cmdLocalizar_Click
End Sub

Private Sub optTodos_Click()
   lblCategoria.Visible = False
   cboConsLinha.Visible = False
   lblDesc.Visible = False
   cboDesc.Visible = False
   cboDesc.Visible = False
   optPorPalavra.Visible = False
   optPalavrasDuplas.Visible = False
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

Private Sub txtEdit_KeyUp(KeyCode As Integer, Shift As Integer)
   'Exit Sub
   If KeyCode = 38 Then
      If Grid.Row - 1 = 0 Then ShowMsg "VOCÊ JÁ ESTÁ NA PRIMEIRA LINHA !!!", vbExclamation: Exit Sub
      Grid.Row = iRow - 1
      Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
      AtualizarQtdDaLinha iRow
      Grid_Click
   
   ElseIf KeyCode = 40 Then
      If Grid.rows = Grid.Row + 1 Then ShowMsg "VOCÊ JÁ ESTÁ NA ULTIMA LINHA !!!", vbExclamation: Exit Sub
      Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
      AtualizarQtdDaLinha iRow
      Grid.Row = iRow + 1
      Grid_Click
   End If
End Sub

Private Sub txtEdit_LostFocus()
Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
AtualizarQtdDaLinha iRow
AtualizarLabelsSelecao
txtEdit.Visible = False
End Sub



