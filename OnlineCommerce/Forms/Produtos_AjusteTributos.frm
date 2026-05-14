VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Produtos_AjusteTributos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AJUSTE DE TRIBUTOS"
   ClientHeight    =   9825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16755
   Icon            =   "Produtos_AjusteTributos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Produtos_AjusteTributos.frx":1D82
   ScaleHeight     =   9825
   ScaleWidth      =   16755
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmAlterarGrupos 
      Caption         =   "Alterar em grupos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   6300
      TabIndex        =   48
      Top             =   7260
      Width           =   6255
      Begin VB.Frame frmEdicao 
         Caption         =   "Alterar em todos"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1560
         TabIndex        =   58
         Top             =   240
         Width           =   4635
         Begin VB.TextBox txtEdicaoColetiva 
            Height          =   315
            Left            =   60
            TabIndex        =   61
            Top             =   480
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.ComboBox cboEdicaoColetiva 
            Height          =   315
            Left            =   60
            TabIndex        =   60
            Top             =   480
            Visible         =   0   'False
            Width           =   3015
         End
         Begin ChamaleonBtn.chameleonButton cmdEdicaoColetiva 
            Height          =   315
            Left            =   3180
            TabIndex        =   59
            Top             =   480
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Atualizar Todos"
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
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Produtos_AjusteTributos.frx":264C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblEdicaoColetiva 
            AutoSize        =   -1  'True
            Caption         =   "Titulo"
            Height          =   195
            Left            =   60
            TabIndex        =   62
            Top             =   240
            Visible         =   0   'False
            Width           =   390
         End
      End
      Begin VB.Frame frmEdicaoFiltros 
         Caption         =   "Alterar"
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
         Height          =   1935
         Left            =   60
         TabIndex        =   49
         Top             =   240
         Width           =   1395
         Begin VB.OptionButton optEIS 
            Caption         =   "Classif. IS"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   1500
            Width           =   1035
         End
         Begin VB.OptionButton optECBS 
            Caption         =   "Classif. CBS"
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   1320
            Width           =   1215
         End
         Begin VB.OptionButton optECest 
            Caption         =   "CEST"
            Height          =   195
            Left            =   120
            TabIndex        =   55
            Top             =   1140
            Width           =   1035
         End
         Begin VB.OptionButton optETags 
            Caption         =   "Tags"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   960
            Width           =   1035
         End
         Begin VB.OptionButton optECategoria 
            Caption         =   "Categoria"
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   780
            Width           =   1035
         End
         Begin VB.OptionButton optECFOP 
            Caption         =   "CFOP"
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   420
            Width           =   1035
         End
         Begin VB.OptionButton optENCM 
            Caption         =   "NCM"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   1155
         End
         Begin VB.OptionButton optEICMS 
            Caption         =   "ICMS CST"
            Height          =   195
            Left            =   120
            TabIndex        =   50
            Top             =   600
            Width           =   1155
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Consultar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   60
      TabIndex        =   15
      Top             =   7260
      Width           =   6195
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
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   4740
         TabIndex        =   44
         Top             =   1200
         Width           =   1395
         Begin VB.OptionButton optORDNCM 
            Caption         =   "NCM"
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   720
            Width           =   795
         End
         Begin VB.OptionButton optORDTags 
            Caption         =   "Tags"
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   540
            Width           =   915
         End
         Begin VB.OptionButton optORDCategoria 
            Caption         =   "Categoria"
            Height          =   195
            Left            =   120
            TabIndex        =   64
            Top             =   360
            Width           =   1035
         End
         Begin VB.OptionButton optORDDesc 
            Caption         =   "Descrição"
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   180
            Value           =   -1  'True
            Width           =   1035
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Busca Avançada"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   1560
         TabIndex        =   34
         Top             =   240
         Width           =   4575
         Begin VB.OptionButton PorPalavraDupla 
            Caption         =   "Palavras Duplas"
            Height          =   195
            Left            =   2700
            TabIndex        =   36
            Top             =   240
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.OptionButton optPorPalavra 
            Caption         =   "Palavra"
            Height          =   195
            Left            =   1800
            TabIndex        =   37
            Top             =   240
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.ComboBox cboDesc 
            Height          =   315
            Left            =   120
            TabIndex        =   39
            Top             =   480
            Visible         =   0   'False
            Width           =   3615
         End
         Begin VB.ComboBox cboConsLinha 
            Height          =   315
            Left            =   120
            TabIndex        =   38
            Top             =   480
            Visible         =   0   'False
            Width           =   3615
         End
         Begin VB.TextBox txtCodBarra 
            Height          =   315
            Left            =   120
            TabIndex        =   35
            Top             =   480
            Visible         =   0   'False
            Width           =   3435
         End
         Begin ChamaleonBtn.chameleonButton cmdLocalizar 
            Height          =   315
            Left            =   3780
            TabIndex        =   40
            Top             =   480
            Visible         =   0   'False
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   556
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
            MICON           =   "Produtos_AjusteTributos.frx":2668
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblDesc 
            Caption         =   "Descrição"
            Height          =   195
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lblCategoria 
            AutoSize        =   -1  'True
            Caption         =   "Categoria"
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label lblCodBarra 
            AutoSize        =   -1  'True
            Caption         =   "Cód. de Barra"
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
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
         ForeColor       =   &H00000000&
         Height          =   1935
         Left            =   60
         TabIndex        =   25
         Top             =   240
         Width           =   1395
         Begin VB.OptionButton optCodBarra 
            Caption         =   "Cód. Barra"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   420
            Width           =   1155
         End
         Begin VB.OptionButton optDesc 
            Caption         =   "Descrição"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   600
            Width           =   1035
         End
         Begin VB.OptionButton optCategoria 
            Caption         =   "Categoria"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   780
            Width           =   1035
         End
         Begin VB.OptionButton optTodos 
            Caption         =   "Todos"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Value           =   -1  'True
            Width           =   1155
         End
         Begin VB.OptionButton optTags 
            Caption         =   "Tags"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   960
            Width           =   1035
         End
         Begin VB.OptionButton optNCM 
            Caption         =   "NCM"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   1140
            Width           =   1035
         End
         Begin VB.OptionButton optClassTribCBS 
            Caption         =   "Classif. CBS"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   1320
            Width           =   1215
         End
         Begin VB.OptionButton optClassTribIS 
            Caption         =   "Classif. IS"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   1500
            Width           =   1035
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Preço"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3360
         TabIndex        =   21
         Top             =   1200
         Width           =   1335
         Begin VB.OptionButton optTodosPreco 
            Caption         =   "Todos"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   540
            Width           =   975
         End
         Begin VB.OptionButton optComPreco 
            Caption         =   "Com Preço"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   180
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optSemPreco 
            Caption         =   "Sem Preço"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame frmFiltro 
         Caption         =   "Quantidade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1560
         TabIndex        =   16
         Top             =   1200
         Width           =   1695
         Begin VB.OptionButton optMostrarTodos 
            Caption         =   "Todos"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   720
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optMostrarZerados 
            Caption         =   "Zerados"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton optMostrarNegativos 
            Caption         =   "Negativos"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   540
            Width           =   1155
         End
         Begin VB.OptionButton optMostrarQuant 
            Caption         =   "Com quantidade"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   180
            Width           =   1455
         End
      End
   End
   Begin VB.ComboBox cboEdit 
      Height          =   315
      Left            =   4020
      TabIndex        =   10
      Top             =   2820
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox picAguarde 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   6300
      Picture         =   "Produtos_AjusteTributos.frx":2684
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
      Height          =   640
      Left            =   60
      ScaleHeight     =   615
      ScaleWidth      =   16605
      TabIndex        =   0
      Top             =   60
      Width           =   16635
      Begin VB.CheckBox chkPISCOFINS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PIS e COFINS"
         Height          =   195
         Left            =   12960
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox chkCest 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cest"
         Height          =   255
         Left            =   14340
         TabIndex        =   13
         Top             =   360
         Width           =   675
      End
      Begin VB.CheckBox chkIPI 
         BackColor       =   &H00FFFFFF&
         Caption         =   "IPI"
         Height          =   255
         Left            =   15060
         TabIndex        =   12
         Top             =   360
         Width           =   555
      End
      Begin VB.CheckBox chkCBSIS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CBS e IS"
         Height          =   255
         Left            =   15660
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   600
         Picture         =   "Produtos_AjusteTributos.frx":36BC
         Top             =   60
         Width           =   480
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "AJUSTE DE TRIBUTOS"
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
         Top             =   140
         Width           =   3420
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   5
      Top             =   9555
      Width           =   16755
      _ExtentX        =   29554
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   25215
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "17:37"
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
      Height          =   6255
      Left            =   60
      TabIndex        =   2
      Top             =   720
      Width           =   16635
      _ExtentX        =   29342
      _ExtentY        =   11033
      _Version        =   393216
      Cols            =   5
      AllowBigSelection=   0   'False
      HighLight       =   0
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdAtualizar 
      Height          =   315
      Left            =   15000
      TabIndex        =   6
      Top             =   7260
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
      MICON           =   "Produtos_AjusteTributos.frx":3FE4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdConsultarNCM 
      Height          =   255
      Left            =   14280
      TabIndex        =   8
      Top             =   8820
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Consultar NCM pela Descrição"
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
      MICON           =   "Produtos_AjusteTributos.frx":4000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdConsultaNCMean 
      Height          =   375
      Left            =   14760
      TabIndex        =   9
      Top             =   9120
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Consultar NCM pelo EAN"
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
      MICON           =   "Produtos_AjusteTributos.frx":401C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblMarcarTodas 
      AutoSize        =   -1  'True
      Caption         =   "Marcar Todas"
      Height          =   195
      Left            =   300
      TabIndex        =   47
      Top             =   7020
      Width           =   990
   End
   Begin VB.Image ImgMarcadaTODAS 
      Height          =   195
      Left            =   60
      Picture         =   "Produtos_AjusteTributos.frx":4038
      Top             =   7020
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgDesmarcadaTODAS 
      Height          =   195
      Left            =   60
      Picture         =   "Produtos_AjusteTributos.frx":6437
      Top             =   7020
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgDesmarcada 
      Height          =   195
      Left            =   4860
      Picture         =   "Produtos_AjusteTributos.frx":87B3
      Top             =   7020
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image ImgMarcada 
      Height          =   195
      Left            =   4560
      Picture         =   "Produtos_AjusteTributos.frx":AB2F
      Top             =   7020
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblQuant 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Left            =   16500
      TabIndex        =   7
      Top             =   7020
      Width           =   225
   End
End
Attribute VB_Name = "Produtos_AjusteTributos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private moCombo As cComboHelper
Private iRow As Long, iCol As Long
Private tipoEmpresa As Long
Private bSetSemX As Boolean
'abrir site para consultar ncm
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long
Private Const conSwNormal = 1
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
   
   If optDesc.Value Then
      Dim sW1 As String
      Dim sW2 As String
      Dim sPosEsp As Integer
      If cboDesc.Text = "[SEM DESCRIÇÃO]" Then
         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "(produtos.descricao = '' OR produtos.descricao IS NULL)"
      ElseIf optPorPalavra.Value Then
         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "produtos.descricao LIKE '%" & Replace(cboDesc.Text, "'", "''") & "%'"
      ElseIf PorPalavraDupla.Value Then
         sPosEsp = InStr(Trim(cboDesc.Text), " ")
         If sPosEsp > 0 Then
            sW1 = Left(Trim(cboDesc.Text), sPosEsp - 1)
            sW2 = Trim(Mid(cboDesc.Text, sPosEsp + 1))
         Else
            sW1 = Trim(cboDesc.Text)
            sW2 = ""
         End If
         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "produtos.descricao LIKE '%" & Replace(sW1, "'", "''") & "%'"
         If Len(sW2) > 0 Then
            var_Criterio = var_Criterio & " AND produtos.descricao LIKE '%" & Replace(sW2, "'", "''") & "%'"
         End If
      End If
   End If
   
   If optCategoria.Value Then
      If cboConsLinha.Text = "[SEM CATEGORIA]" Then
         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "(produtos.categoria = '' OR produtos.categoria IS NULL)"
      Else
         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "produtos.categoria = '" & cboConsLinha.Text & "'"
      End If
   End If
   If optCodBarra.Value Then
      If txtCodBarra.Text = "[SEM CÓD. BARRA]" Then
         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "(produtos.cod_barra = '' OR produtos.cod_barra IS NULL)"
      Else
         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "produtos.cod_barra = '" & txtCodBarra.Text & "'"
      End If
   End If
   If optTags.Value Then
      If cboConsLinha.Text = "[SEM TAGS]" Then
         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "(produtos.TAGS = '' OR produtos.TAGS IS NULL)"
      Else
         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "produtos.TAGS = '" & cboConsLinha.Text & "'"
      End If
   End If
   If optNCM.Value Then
      If cboConsLinha.Text = "[SEM NCM]" Then
         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "(produtos.NCM = '' OR produtos.NCM IS NULL OR produtos.NCM = '00000000')"
      Else
         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "produtos.NCM = '" & cboConsLinha.Text & "'"
      End If
   End If
   If optClassTribCBS.Value Then
      If cboConsLinha.Text = "[SEM CBS CLASS.]" Then
         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "(produtos.cClassTrib = '' OR produtos.cClassTrib IS NULL)"
      Else
         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "produtos.cClassTrib = '" & Left(cboConsLinha.Text, 6) & "'"
      End If
   End If
   If optClassTribIS.Value Then
      If cboConsLinha.Text = "[SEM IS CLASS.]" Then
         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "(produtos.cClassTrib_IS = '' OR produtos.cClassTrib_IS IS NULL)"
      Else
         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "produtos.cClassTrib_IS = '" & Left(cboConsLinha.Text, 6) & "'"
      End If
   End If
   
   If var_Criterio <> "" Then var_Criterio = " WHERE " & var_Criterio
   
   If optORDDesc.Value Then
      var_Indice = " ORDER BY produtos.descricao"
   ElseIf optORDCategoria.Value Then
      var_Indice = " ORDER BY produtos.categoria"
   ElseIf optORDTags.Value Then
      var_Indice = " ORDER BY produtos.TAGS"
   ElseIf optORDNCM.Value Then
      var_Indice = " ORDER BY produtos.NCM"
   Else
      var_Indice = ""
   End If
   
   sSQL = "SELECT  produtos.NCM AS var_NCM, produtos.CFOP AS var_CFOP, produtos.ICMSCST AS var_ICMS, produtos.categoria AS var_cat, produtos.fabricante AS var_fab, " & _
      "produtos.codigo AS var_cod, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc,  produtos.ICMSaliq AS var_ICMSAliq, produtos.PISCST AS var_PIS, produtos.pisaliq AS var_PisAliq, produtos.COFINSCST AS var_COFINS, produtos.COFINSALIQ AS var_COFINSALIQ, produtos.IPICST AS var_IPI, produtos.IPIALIQ AS var_IPIALIQ, produtos.CEST AS var_CEST, produtos.cClassTrib AS var_CBS, produtos.cClassTrib_IS AS var_IS, produtos.categoria AS var_cat, produtos.TAGS AS var_tags " & _
      "FROM produtos " & var_Criterio & " " & var_Indice
   
   Set r = dbData.OpenRecordset(sSQL)
   lblQuant.Caption = "Quant.: " & r.RecordCount
   
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

Private Sub optORDDesc_Click()
   MostrarCriterios
End Sub

Private Sub optORDCategoria_Click()
   MostrarCriterios
End Sub

Private Sub optORDTags_Click()
   MostrarCriterios
End Sub

Private Sub optORDNCM_Click()
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

Private Sub cmdAtualizar_Click()
Dim i As Integer
Dim sSQL As String

picAguarde.Visible = True
DoEvents

For i = 1 To Grid.rows - 1
   'Atualiza a tabela de produtos
   sSQL = "UPDATE produtos SET " & _
      "cod_barra = '" & Grid.TextMatrix(i, 3) & "', " & _
      "descricao = '" & Grid.TextMatrix(i, 4) & "', " & _
      "categoria = '" & Grid.TextMatrix(i, 6) & "', " & _
      "TAGS = '" & Grid.TextMatrix(i, 7) & "', " & _
      "NCM = '" & Grid.TextMatrix(i, 8) & "', " & _
      "CFOP = " & Grid.TextMatrix(i, 9) & ", " & _
      "ICMSCST = '" & Grid.TextMatrix(i, 10) & "', " & _
      "ICMSaliq = " & Replace(CDbl(Grid.TextMatrix(i, 11)), ",", ".") & ", " & _
      "pisCST = '" & Grid.TextMatrix(i, 12) & "', " & _
      "pisaliq = " & Replace(CDbl(Grid.TextMatrix(i, 13)), ",", ".") & ", " & _
      "cofinsCST = '" & Grid.TextMatrix(i, 14) & "', " & _
      "cofinsaliq = " & Replace(CDbl(Grid.TextMatrix(i, 15)), ",", ".") & ", " & _
      "ipiCST = '" & Grid.TextMatrix(i, 16) & "', " & _
      "ipialiq = " & Replace(CDbl(Grid.TextMatrix(i, 17)), ",", ".") & ", " & _
      "cest = '" & Grid.TextMatrix(i, 18) & "', " & _
      "cClassTrib = '" & Grid.TextMatrix(i, 19) & "', " & _
      "cClassTrib_IS = '" & Grid.TextMatrix(i, 20) & "' " & _
      "WHERE (codigo = " & Grid.TextMatrix(i, 2) & ");"
      'Debug.Print sSQL
   dbData.Execute sSQL
Next

Dim rTag As ADODB.Recordset
Dim sCatNomeTmp As String
Dim lCatIDTmp As Long
Dim sTagTmp As String
For i = 1 To Grid.rows - 1
   sTagTmp = Trim(Grid.TextMatrix(i, 7))
   sCatNomeTmp = Trim(Grid.TextMatrix(i, 6))
   If Len(sTagTmp) > 0 And Len(sCatNomeTmp) > 0 Then
      lCatIDTmp = 0
      Set rTag = dbData.OpenRecordset("SELECT ID_Categoria FROM Categorias WHERE Categoria = '" & Replace(sCatNomeTmp, "'", "''") & "'")
      If Not rTag.EOF Then lCatIDTmp = CLng(rTag("ID_Categoria"))
      If rTag.State <> 0 Then rTag.Close
      If lCatIDTmp > 0 Then
         Set rTag = dbData.OpenRecordset("SELECT COUNT(*) AS qtd FROM Categorias_Tags WHERE Tags = '" & Replace(sTagTmp, "'", "''") & "' AND ID_Categoria = " & lCatIDTmp)
         If Not rTag.EOF Then
            If CLng(rTag("qtd")) = 0 Then
               dbData.Execute "INSERT INTO Categorias_Tags (Tags, ID_Categoria) VALUES ('" & Replace(sTagTmp, "'", "''") & "', " & lCatIDTmp & ");"
            End If
         End If
         If rTag.State <> 0 Then rTag.Close
      End If
   End If
Next i

picAguarde.Visible = False
ResetarMarcas
End Sub

Private Sub cmdConsultaNCMean_Click()
Dim varNomeProduto As String
varNomeProduto = Grid.TextMatrix(Grid.Row, 3)
ShellExecute hwnd, "open", "https://cosmos.bluesoft.com.br/pesquisar?utf8=" + Chr(95) + "&q=" & varNomeProduto & "", vbNullString, vbNullString, conSwNo
End Sub

Private Sub cmdConsultarNCM_Click()
Dim varNomeProduto As String
varNomeProduto = Replace(Grid.TextMatrix(Grid.Row, 4), " ", "+")
ShellExecute hwnd, "open", "https://cosmos.bluesoft.com.br/pesquisar?utf8=" + Chr(95) + "&q=" & varNomeProduto & "", vbNullString, vbNullString, conSwNo
End Sub

Private Sub cmdLocalizar_Click()
cboEdit.Visible = False
txtEdit.Visible = False
Dim sSQL As String
Dim r As ADODB.Recordset
Dim varTipoMostrar As String
Dim vUltimoValorVenda As String

If optMostrarQuant.Value = True Then
    varTipoMostrar = " AND produtos.quant_estoque > 0"
ElseIf optMostrarNegativos.Value = True Then
    varTipoMostrar = " AND produtos.quant_estoque < 0"
ElseIf optMostrarZerados.Value = True Then
    varTipoMostrar = " AND produtos.quant_estoque = 0"
ElseIf optMostrarTodos.Value = True Then
    varTipoMostrar = " "
End If

If optComPreco.Value = True Then
    vUltimoValorVenda = " and (SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) > 0 "
ElseIf optSemPreco.Value = True Then
    vUltimoValorVenda = " and (SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) = 0"
ElseIf optTodosPreco.Value = True Then
    vUltimoValorVenda = " and 1=1 "
End If

If optTodos.Value = True Then
   sSQL = "SELECT  produtos.NCM AS var_NCM, produtos.CFOP AS var_CFOP, produtos.ICMSCST AS var_ICMS, produtos.categoria AS var_cat, produtos.fabricante AS var_fab, " & _
      "produtos.codigo AS var_cod, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc,  produtos.ICMSaliq AS var_ICMSAliq, produtos.PISCST AS var_PIS, produtos.pisaliq AS var_PisAliq, produtos.COFINSCST AS var_COFINS, produtos.COFINSALIQ AS var_COFINSALIQ, produtos.IPICST AS var_IPI, produtos.IPIALIQ AS var_IPIALIQ, produtos.CEST AS var_CEST, produtos.cClassTrib AS var_CBS, produtos.cClassTrib_IS AS var_IS, produtos.categoria AS var_cat, produtos.TAGS AS var_tags " & _
      "FROM produtos " & _
      "WHERE (produtos.ativo = 1) " & varTipoMostrar & " " & vUltimoValorVenda & " ORDER BY produtos.descricao;"

   Set r = dbData.OpenRecordset(sSQL)
   lblQuant.Caption = "Quant.: " & r.RecordCount
   
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
If optTodos.Value Or optCodBarra.Value Or optDesc.Value Then
   frmEdicao.Enabled = False
   frmEdicaoFiltros.Enabled = False
Else
   frmEdicao.Enabled = True
   frmEdicaoFiltros.Enabled = True
End If
End Sub

Private Sub Form_Activate()
cmdLocalizar_Click
End Sub

Private Sub Form_Load()
Set moCombo = New cComboHelper
tipoEmpresa = CLng(sysConfig("TIPO_EMPRESA").Value)
End Sub

Private Sub cboDesc_Change()
   Dim pos As Integer
   pos = cboDesc.SelStart
   cboDesc.Text = UCase(cboDesc.Text)
   cboDesc.SelStart = pos
End Sub

Private Sub cboDesc_KeyPress(KeyAscii As Integer)
   If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
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
   Dim sField As String
   
   cboConsLinha.Clear
   
   If optCategoria.Value Then
      sField = "Categoria"
      sSQL = "SELECT DISTINCT Categoria FROM Categorias WHERE Tipo_Empresa = " & tipoEmpresa & " ORDER BY Categoria;"
   ElseIf optTags.Value Then
      sField = "Tags"
      sSQL = "SELECT ct.Tags FROM Categorias_Tags ct INNER JOIN Categorias c ON ct.ID_Categoria = c.ID_Categoria WHERE c.Tipo_Empresa = " & tipoEmpresa & " ORDER BY c.Categoria, ct.Tags;"
   ElseIf optNCM.Value Then
      sField = "NCM"
      sSQL = "SELECT DISTINCT NCM FROM produtos WHERE NCM IS NOT NULL AND NCM <> '' ORDER BY NCM;"
   ElseIf optClassTribCBS.Value Then
      sField = "val"
      sSQL = "SELECT cClassTrib + ' - ' + NomecClassTrib AS val FROM TbIBSCBSClassTrib GROUP BY cClassTrib, NomecClassTrib ORDER BY cClassTrib;"
   ElseIf optClassTribIS.Value Then
      sField = "val"
      sSQL = "SELECT cClassTrib_IS + ' - ' + Descricao AS val FROM tbISClassTrib GROUP BY cClassTrib_IS, Descricao ORDER BY cClassTrib_IS;"
   Else
      Exit Sub
   End If
   
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboConsLinha.AddItem ValidateNull(r(sField))
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   If optCategoria.Value Then
      cboConsLinha.AddItem "[SEM CATEGORIA]", 0
   ElseIf optTags.Value Then
      cboConsLinha.AddItem "[SEM TAGS]", 0
   ElseIf optNCM.Value Then
      cboConsLinha.AddItem "[SEM NCM]", 0
   ElseIf optClassTribCBS.Value Then
      cboConsLinha.AddItem "[SEM CBS CLASS.]", 0
   ElseIf optClassTribIS.Value Then
      cboConsLinha.AddItem "[SEM IS CLASS.]", 0
   End If
   
   moCombo.AttachTo cboConsLinha
   If bSetSemX Then
      bSetSemX = False
      If optCategoria.Value Then
         cboConsLinha.Text = "[SEM CATEGORIA]"
      ElseIf optTags.Value Then
         cboConsLinha.Text = "[SEM TAGS]"
      ElseIf optNCM.Value Then
         cboConsLinha.Text = "[SEM NCM]"
      ElseIf optClassTribCBS.Value Then
         cboConsLinha.Text = "[SEM CBS CLASS.]"
      ElseIf optClassTribIS.Value Then
         cboConsLinha.Text = "[SEM IS CLASS.]"
      End If
   End If
End Sub

Private Sub cboConsLinha_LostFocus()
   'cboConsLinha_Click
End Sub

Private Sub Formatar_Grid(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   LimparGrid
   picAguarde.Visible = True
   DoEvents
   
   With Grid
      .Clear
      .Cols = 21
      .rows = 2
      .FixedRows = 1
      .FixedCols = 0
      
      .ColWidth(0) = 360
      .ColWidth(1) = 0
      .ColWidth(2) = 0
      .ColWidth(3) = 1300
      .ColWidth(4) = 4000
      .ColWidth(5) = 1400
      .ColWidth(6) = 1400
      .ColWidth(7) = 1200
      .ColWidth(8) = 850
      .ColWidth(9) = 700
      .ColWidth(10) = 700
      .ColWidth(11) = 700
      .ColWidth(12) = 650
      .ColWidth(13) = 700
      .ColWidth(14) = 800
      .ColWidth(15) = 700
      .ColWidth(16) = 650
      .ColWidth(17) = 700
      .ColWidth(18) = 900
      .ColWidth(19) = 900
      .ColWidth(20) = 900
      
      '.RowHeight(-1) = (315 * 1)    'definir a altura da linha
      
      .TextMatrix(0, 1) = "CÓD.ENT"
      .TextMatrix(0, 2) = "CÓD.PROD"
      .TextMatrix(0, 3) = "CÓD.BARRA"
      .TextMatrix(0, 4) = "DESCRIÇÃO"
      .TextMatrix(0, 5) = "FABRICANTE"
      .TextMatrix(0, 6) = "CATEGORIA"
      .TextMatrix(0, 7) = "TAGS"
      .TextMatrix(0, 8) = "NCM."
      .TextMatrix(0, 9) = "CFOP."
      .TextMatrix(0, 10) = "ICMS."
      .TextMatrix(0, 11) = "ALIQ."
      .TextMatrix(0, 12) = "PIS"
      .TextMatrix(0, 13) = "ALIQ."
      .TextMatrix(0, 14) = "COFINS"
      .TextMatrix(0, 15) = "ALIQ."
      .TextMatrix(0, 16) = "IPI"
      .TextMatrix(0, 17) = "ALIQ."
      .TextMatrix(0, 18) = "CEST"
      .TextMatrix(0, 19) = "CL.CBS"
      .TextMatrix(0, 20) = "CL.IS"
      
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
            .TextMatrix(.rows - 1, 3) = ValidateNull(rTabela("var_codbarra"))
            .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("var_desc"))
            .TextMatrix(.rows - 1, 5) = ValidateNull(rTabela("var_fab"))
            .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("var_cat"))
            .TextMatrix(.rows - 1, 7) = ValidateNull(rTabela("var_tags"))
            .TextMatrix(.rows - 1, 8) = ValidateNull(rTabela("var_NCM"))
            .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("var_CFOP"))
            .TextMatrix(.rows - 1, 10) = ValidateNull(rTabela("var_ICMS"))
            .TextMatrix(.rows - 1, 11) = Format$(ValidateNull(rTabela("var_ICMSALIQ")), ocMONEY)
            .TextMatrix(.rows - 1, 12) = ValidateNull(rTabela("var_PIS"))
            .TextMatrix(.rows - 1, 13) = Format$(ValidateNull(rTabela("var_PISALIQ")), ocMONEY)
            .TextMatrix(.rows - 1, 14) = ValidateNull(rTabela("var_COFINS"))
            .TextMatrix(.rows - 1, 15) = Format$(ValidateNull(rTabela("var_COFINSALIQ")), ocMONEY)
            .TextMatrix(.rows - 1, 16) = ValidateNull(rTabela("var_IPI"))
            .TextMatrix(.rows - 1, 17) = Format$(ValidateNull(rTabela("var_IPIALIQ")), ocMONEY)
            .TextMatrix(.rows - 1, 18) = ValidateNull(rTabela("var_CEST"))
            .TextMatrix(.rows - 1, 19) = ValidateNull(rTabela("var_CBS"))
            .TextMatrix(.rows - 1, 20) = ValidateNull(rTabela("var_IS"))
            
            '.TextMatrix(.Rows - 1, 9) = ValidateNull(rTabela("var_quant"))
            '.TextMatrix(.Rows - 1, 11) = Format$(ValidateNull(rTabela("venda")), ocMONEY)
            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If
      
      .rows = .rows - 1
      Dim lRow As Long
      .FillStyle = 1
      For lRow = 1 To .rows - 1
         .Row = lRow
         .Col = 0
         .ColSel = .Cols - 1
         If lRow Mod 2 = 0 Then
            .CellBackColor = &HE0E0E0
         Else
            .CellBackColor = vbWhite
         End If
      Next lRow
      .FillStyle = 0
      .Redraw = True
      picAguarde.Visible = False
   Dim lChk As Long
   For lChk = 1 To Grid.rows - 1
      Grid.Row = lChk
      Grid.Col = 0
      Set Grid.CellPicture = imgDesmarcada.Picture
      Grid.CellPictureAlignment = 4
   Next lChk
   End With
   chkPISCOFINS_Click
   chkIPI_Click
   chkCest_Click
   chkCBSIS_Click
   imgDesmarcadaTODAS.Visible = True
   ImgMarcadaTODAS.Visible = False
   lblMarcarTodas.Caption = "Marcar todos"
   AvaliarFrmAlterarGrupos
End Sub

Private Sub chkPISCOFINS_Click()
If chkPISCOFINS.Value = Checked Then
   Grid.ColWidth(12) = 650
   Grid.ColWidth(13) = 700
   Grid.ColWidth(14) = 800
   Grid.ColWidth(15) = 700
Else
   Grid.ColWidth(12) = 0
   Grid.ColWidth(13) = 0
   Grid.ColWidth(14) = 0
   Grid.ColWidth(15) = 0
End If
End Sub

Private Sub chkIPI_Click()
If chkIPI.Value = Checked Then
   Grid.ColWidth(16) = 650
   Grid.ColWidth(17) = 700
Else
   Grid.ColWidth(16) = 0
   Grid.ColWidth(17) = 0
End If
End Sub

Private Sub chkCest_Click()
If chkCest.Value = Checked Then
   Grid.ColWidth(18) = 900
Else
   Grid.ColWidth(18) = 0
End If
End Sub

Private Sub chkCBSIS_Click()
If chkCBSIS.Value = Checked Then
   Grid.ColWidth(19) = 900
   Grid.ColWidth(20) = 900
Else
   Grid.ColWidth(19) = 0
   Grid.ColWidth(20) = 0
End If
End Sub

Private Sub optEICMS_Click()
   lblEdicaoColetiva.Caption = "ICMS CST"
   txtEdicaoColetiva.Visible = True
   cmdEdicaoColetiva.Visible = True
   cmdAtualizar.Visible = False
   lblEdicaoColetiva.Visible = True
   cboEdicaoColetiva.Visible = False
End Sub

Private Sub optENCM_Click()
   lblEdicaoColetiva.Caption = "NCM"
   txtEdicaoColetiva.Visible = True
   cmdEdicaoColetiva.Visible = True
   cmdAtualizar.Visible = False
   lblEdicaoColetiva.Visible = True
   cboEdicaoColetiva.Visible = False
End Sub

Private Sub optECFOP_Click()
   lblEdicaoColetiva.Caption = "CFOP"
   txtEdicaoColetiva.Visible = True
   cmdEdicaoColetiva.Visible = True
   cmdAtualizar.Visible = False
   lblEdicaoColetiva.Visible = True
   cboEdicaoColetiva.Visible = False
End Sub

Private Sub optECest_Click()
   lblEdicaoColetiva.Caption = "CEST"
   txtEdicaoColetiva.Visible = True
   cmdEdicaoColetiva.Visible = True
   cmdAtualizar.Visible = False
   lblEdicaoColetiva.Visible = True
   chkCest.Value = 1
   cboEdicaoColetiva.Visible = False
End Sub

Private Sub optECategoria_Click()
Dim r As ADODB.Recordset
   lblEdicaoColetiva.Caption = "Categoria"
   txtEdicaoColetiva.Visible = False
   cboEdicaoColetiva.Clear
   Set r = dbData.OpenRecordset("SELECT DISTINCT Categoria FROM Categorias WHERE Tipo_Empresa = " & tipoEmpresa & " ORDER BY Categoria;")
   Do While Not r.EOF: cboEdicaoColetiva.AddItem ValidateNull(r("Categoria")): r.MoveNext: Loop
   r.Close: Set r = Nothing
   cboEdicaoColetiva.Visible = True
   cmdEdicaoColetiva.Visible = True
   cmdAtualizar.Visible = False
   lblEdicaoColetiva.Visible = True
End Sub

Private Sub optETags_Click()
Dim r As ADODB.Recordset
   lblEdicaoColetiva.Caption = "Tags"
   txtEdicaoColetiva.Visible = False
   cboEdicaoColetiva.Clear
   Set r = dbData.OpenRecordset("SELECT ct.Tags FROM Categorias_Tags ct INNER JOIN Categorias c ON ct.ID_Categoria = c.ID_Categoria WHERE c.Tipo_Empresa = " & tipoEmpresa & " ORDER BY c.Categoria, ct.Tags;")
   Do While Not r.EOF: cboEdicaoColetiva.AddItem ValidateNull(r("Tags")): r.MoveNext: Loop
   r.Close: Set r = Nothing
   cboEdicaoColetiva.Visible = True
   cmdEdicaoColetiva.Visible = True
   cmdAtualizar.Visible = False
   lblEdicaoColetiva.Visible = True
End Sub

Private Sub optECBS_Click()
Dim r As ADODB.Recordset
   lblEdicaoColetiva.Caption = "CBS Class."
   txtEdicaoColetiva.Visible = False
   cboEdicaoColetiva.Clear
   Set r = dbData.OpenRecordset("SELECT cClassTrib + ' - ' + NomecClassTrib AS val FROM TbIBSCBSClassTrib GROUP BY cClassTrib, NomecClassTrib ORDER BY cClassTrib;")
   Do While Not r.EOF: cboEdicaoColetiva.AddItem ValidateNull(r("val")): r.MoveNext: Loop
   r.Close: Set r = Nothing
   cboEdicaoColetiva.Visible = True
   cmdEdicaoColetiva.Visible = True
   cmdAtualizar.Visible = False
   chkCBSIS.Value = 1
   lblEdicaoColetiva.Visible = True
End Sub

Private Sub optEIS_Click()
Dim r As ADODB.Recordset
   lblEdicaoColetiva.Caption = "IS Class."
   txtEdicaoColetiva.Visible = False
   cboEdicaoColetiva.Clear
   Set r = dbData.OpenRecordset("SELECT cClassTrib_IS + ' - ' + Descricao AS val FROM tbISClassTrib GROUP BY cClassTrib_IS, Descricao ORDER BY cClassTrib_IS;")
   Do While Not r.EOF: cboEdicaoColetiva.AddItem ValidateNull(r("val")): r.MoveNext: Loop
   r.Close: Set r = Nothing
   chkCBSIS.Value = 1
   cboEdicaoColetiva.Visible = True
   cmdEdicaoColetiva.Visible = True
   cmdAtualizar.Visible = False
   lblEdicaoColetiva.Visible = True
End Sub

Private Sub cmdEdicaoColetiva_Click()
Dim i As Integer
Dim sVal As String
Dim iCol As Integer
Dim sSet As String
Dim p As Integer

If txtEdicaoColetiva.Visible Then
   sVal = txtEdicaoColetiva.Text
Else
   sVal = cboEdicaoColetiva.Text
End If
If optETags.Value Or optECategoria.Value Then sVal = UCase(sVal)

If optEICMS.Value Then
   iCol = 10: sSet = "ICMSCST = '" & sVal & "'"
ElseIf optENCM.Value Then
   iCol = 8: sSet = "NCM = '" & sVal & "'"
ElseIf optECFOP.Value Then
   iCol = 9: sSet = "CFOP = " & sVal
ElseIf optECest.Value Then
   iCol = 18: sSet = "cest = '" & sVal & "'"
ElseIf optECategoria.Value Then
   iCol = 6: sSet = "categoria = '" & sVal & "'"
ElseIf optETags.Value Then
   iCol = 7: sSet = "TAGS = '" & sVal & "'"
ElseIf optECBS.Value Then
   p = InStr(sVal, " - "): If p > 0 Then sVal = Left(sVal, p - 1)
   iCol = 19: sSet = "cClassTrib = '" & sVal & "'"
ElseIf optEIS.Value Then
   p = InStr(sVal, " - "): If p > 0 Then sVal = Left(sVal, p - 1)
   iCol = 20: sSet = "cClassTrib_IS = '" & sVal & "'"
Else
   Exit Sub
End If

If Len(Trim(sSet)) = 0 Then Exit Sub

Dim iExact As Integer
iExact = 0
If optENCM.Value Then iExact = 8
If optECFOP.Value Then iExact = 4
If optEICMS.Value Then iExact = 3
If optECest.Value Then iExact = 8
If optECBS.Value Then iExact = 6
If optEIS.Value Then iExact = 6
If iExact > 0 Then
   Dim k As Integer
   Dim bNumOk As Boolean
   bNumOk = (Len(sVal) = iExact)
   If bNumOk Then
      For k = 1 To Len(sVal)
         If Mid(sVal, k, 1) < "0" Or Mid(sVal, k, 1) > "9" Then bNumOk = False: Exit For
      Next k
   End If
   If Not bNumOk Then
      MsgBox "Digite exatamente " & iExact & " dígitos numéricos.", vbExclamation, "Edição Coletiva"
      If optECBS.Value Or optEIS.Value Then
         cboEdicaoColetiva.SetFocus
      Else
         txtEdicaoColetiva.SetFocus
      End If
      Exit Sub
   End If
End If

Dim bTemMarcadas As Boolean
For i = 1 To Grid.rows - 1
   If Grid.TextMatrix(i, 0) = "1" Then bTemMarcadas = True: Exit For
Next i

If optETags.Value Then
   Dim iLinSemCat As Integer
   For i = 1 To Grid.rows - 1
      If bTemMarcadas And Grid.TextMatrix(i, 0) <> "1" Then GoTo ProxValCat
      If Len(Trim(Grid.TextMatrix(i, 6))) = 0 Then iLinSemCat = iLinSemCat + 1
ProxValCat:
   Next i
   If iLinSemCat > 0 Then
      MsgBox iLinSemCat & " produto(s) sem categoria. Defina a categoria antes de aplicar a tag.", vbExclamation, "Tag sem categoria"
      Exit Sub
   End If
End If

picAguarde.Visible = True
DoEvents

For i = 1 To Grid.rows - 1
   If bTemMarcadas And Grid.TextMatrix(i, 0) <> "1" Then GoTo ProximaLinha
   Grid.TextMatrix(i, iCol) = sVal
   dbData.Execute "UPDATE produtos SET " & sSet & " WHERE codigo = " & Grid.TextMatrix(i, 2) & ";"
ProximaLinha:
Next i

txtEdicaoColetiva.Text = ""
cboEdicaoColetiva.Text = ""
If optETags.Value And Len(Trim(sVal)) > 0 Then
   Dim rCat As ADODB.Recordset
   Dim sCatNome As String
   Dim lCatID As Long
   Dim j As Integer
   For j = 1 To Grid.rows - 1
      If bTemMarcadas And Grid.TextMatrix(j, 0) <> "1" Then GoTo ProxCat
      sCatNome = Grid.TextMatrix(j, 6)
      If Len(Trim(sCatNome)) = 0 Then GoTo ProxCat
      lCatID = 0
      Set rCat = dbData.OpenRecordset("SELECT ID_Categoria FROM Categorias WHERE Categoria = '" & Replace(sCatNome, "'", "''") & "'")
      If Not rCat.EOF Then lCatID = CLng(rCat("ID_Categoria"))
      If rCat.State <> 0 Then rCat.Close
      If lCatID > 0 Then
         Set rCat = dbData.OpenRecordset("SELECT COUNT(*) AS qtd FROM Categorias_Tags WHERE Tags = '" & Replace(sVal, "'", "''") & "' AND ID_Categoria = " & lCatID)
         If Not rCat.EOF Then
            If CLng(rCat("qtd")) = 0 Then
               dbData.Execute "INSERT INTO Categorias_Tags (Tags, ID_Categoria) VALUES ('" & Replace(sVal, "'", "''") & "', " & lCatID & ");"
            End If
         End If
         If rCat.State <> 0 Then rCat.Close
      End If
ProxCat:
   Next j
End If

picAguarde.Visible = False
ResetarMarcas
optENCM.Value = False
optECFOP.Value = False
optEICMS.Value = False
optECest.Value = False
optECategoria.Value = False
optETags.Value = False
optECBS.Value = False
optEIS.Value = False
txtEdicaoColetiva.Visible = False
cboEdicaoColetiva.Visible = False
cmdEdicaoColetiva.Visible = False
lblEdicaoColetiva.Visible = False
cmdAtualizar.Visible = True
End Sub

Private Sub txtEdicaoColetiva_LostFocus()
Dim s As String
s = Trim(txtEdicaoColetiva.Text)
If optENCM.Value Or optECest.Value Then
   s = Replace(s, ".", "")
   s = Replace(s, " ", "")
End If
txtEdicaoColetiva.Text = s
End Sub

Private Sub cboEdicaoColetiva_KeyPress(KeyAscii As Integer)
   If Not (optETags.Value Or optECategoria.Value) Then Exit Sub
   If KeyAscii >= 97 And KeyAscii <= 122 Then
      KeyAscii = KeyAscii - 32
   ElseIf optETags.Value Then
      If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8 Or KeyAscii = 32) Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub cboEdicaoColetiva_Change()
   If Not (optETags.Value Or optECategoria.Value) Then Exit Sub
   Dim pos As Integer
   pos = cboEdicaoColetiva.SelStart
   cboEdicaoColetiva.Text = UCase(cboEdicaoColetiva.Text)
   cboEdicaoColetiva.SelStart = pos
End Sub

Private Sub cboEdicaoColetiva_LostFocus()
   cboEdicaoColetiva.Text = Trim(cboEdicaoColetiva.Text)
End Sub

Private Sub ResetarMarcas()
Dim i As Integer
For i = 1 To Grid.rows - 1
   Grid.TextMatrix(i, 0) = ""
   Grid.Row = i: Grid.Col = 0
   Set Grid.CellPicture = imgDesmarcada.Picture
   Grid.CellPictureAlignment = 4
Next i
imgDesmarcadaTODAS.Visible = True
ImgMarcadaTODAS.Visible = False
lblMarcarTodas.Caption = "Marcar todos"
AvaliarFrmAlterarGrupos
End Sub

Private Sub AvaliarFrmAlterarGrupos()
Dim j As Integer
Dim bEnabled As Boolean
For j = 1 To Grid.rows - 1
   If Grid.TextMatrix(j, 0) = "1" Then bEnabled = True: Exit For
Next j
If Not bEnabled Then
   If Not (optTodos.Value Or optCodBarra.Value Or optDesc.Value) Then
      If Grid.rows > 1 Then bEnabled = True
   End If
End If
frmAlterarGrupos.Enabled = bEnabled
frmEdicao.Enabled = bEnabled
frmEdicaoFiltros.Enabled = bEnabled
End Sub

Private Sub imgDesmarcadaTODAS_Click()
Dim i As Integer
imgDesmarcadaTODAS.Visible = False
ImgMarcadaTODAS.Visible = True
lblMarcarTodas.Caption = "Desmarcar todos"
For i = 1 To Grid.rows - 1
   Grid.TextMatrix(i, 0) = "1"
   Grid.Row = i: Grid.Col = 0
   Set Grid.CellPicture = ImgMarcada.Picture
   Grid.CellPictureAlignment = 4
Next i
AvaliarFrmAlterarGrupos
End Sub

Private Sub ImgMarcadaTODAS_Click()
Dim i As Integer
ImgMarcadaTODAS.Visible = False
imgDesmarcadaTODAS.Visible = True
lblMarcarTodas.Caption = "Marcar todos"
For i = 1 To Grid.rows - 1
   Grid.TextMatrix(i, 0) = ""
   Grid.Row = i: Grid.Col = 0
   Set Grid.CellPicture = imgDesmarcada.Picture
   Grid.CellPictureAlignment = 4
Next i
AvaliarFrmAlterarGrupos
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Grid_Click()
Dim r As ADODB.Recordset

txtEdit.Visible = False
cboEdit.Visible = False
iRow = Grid.Row
iCol = Grid.Col

Select Case iCol
Case 0
   If iRow < 1 Then Exit Sub
   If Grid.TextMatrix(iRow, 0) = "1" Then
      Grid.TextMatrix(iRow, 0) = ""
      Grid.Row = iRow: Grid.Col = 0
      Set Grid.CellPicture = imgDesmarcada.Picture
      Grid.CellPictureAlignment = 4
   Else
      Grid.TextMatrix(iRow, 0) = "1"
      Grid.Row = iRow: Grid.Col = 0
      Set Grid.CellPicture = ImgMarcada.Picture
      Grid.CellPictureAlignment = 4
   End If
   AvaliarFrmAlterarGrupos
Case 6
   cboEdit.Clear
   Set r = dbData.OpenRecordset("SELECT DISTINCT Categoria FROM Categorias WHERE Tipo_Empresa = " & tipoEmpresa & " ORDER BY Categoria;")
   Do While Not r.EOF: cboEdit.AddItem ValidateNull(r("Categoria")): r.MoveNext: Loop
   r.Close: Set r = Nothing
   cboEdit.Text = Grid.TextMatrix(iRow, iCol)
   cboEdit.ZOrder 0
   cboEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth
   cboEdit.Visible = True
Case 7
   If Len(Trim(Grid.TextMatrix(iRow, 6))) = 0 Then
      MsgBox "Defina a categoria do produto antes de editar a tag.", vbExclamation, "Tag sem categoria"
      Exit Sub
   End If
   cboEdit.Clear
   Set r = dbData.OpenRecordset("SELECT ct.Tags FROM Categorias_Tags ct INNER JOIN Categorias c ON ct.ID_Categoria = c.ID_Categoria WHERE c.Categoria = '" & Replace(Grid.TextMatrix(iRow, 6), "'", "''") & "' ORDER BY ct.Tags;")
   Do While Not r.EOF: cboEdit.AddItem ValidateNull(r("Tags")): r.MoveNext: Loop
   r.Close: Set r = Nothing
   cboEdit.Text = Grid.TextMatrix(iRow, iCol)
   cboEdit.ZOrder 0
   cboEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth
   cboEdit.Visible = True
Case 19
   cboEdit.Clear
   Set r = dbData.OpenRecordset("SELECT cClassTrib + ' - ' + NomecClassTrib AS val FROM TbIBSCBSClassTrib GROUP BY cClassTrib, NomecClassTrib ORDER BY cClassTrib;")
   Do While Not r.EOF: cboEdit.AddItem ValidateNull(r("val")): r.MoveNext: Loop
   r.Close: Set r = Nothing
   cboEdit.Text = Grid.TextMatrix(iRow, iCol)
   cboEdit.ZOrder 0
   cboEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth
   cboEdit.Visible = True
Case 20
   cboEdit.Clear
   Set r = dbData.OpenRecordset("SELECT cClassTrib_IS + ' - ' + Descricao AS val FROM tbISClassTrib GROUP BY cClassTrib_IS, Descricao ORDER BY cClassTrib_IS;")
   Do While Not r.EOF: cboEdit.AddItem ValidateNull(r("val")): r.MoveNext: Loop
   r.Close: Set r = Nothing
   cboEdit.Text = Grid.TextMatrix(iRow, iCol)
   cboEdit.ZOrder 0
   cboEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth
   cboEdit.Visible = True
Case 8 To 18
   txtEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth, Grid.CellHeight
   txtEdit.Text = Grid.TextMatrix(iRow, iCol)
   txtEdit.Visible = True
   txtEdit.SetFocus
   txtEdit.SelStart = 0
   txtEdit.SelLength = Len(txtEdit.Text)
End Select
End Sub

Private Sub cboEdit_KeyPress(KeyAscii As Integer)
   If iCol <> 7 Then Exit Sub
   If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
End Sub

Private Sub cboEdit_Change()
   If iCol <> 7 Then Exit Sub
   Dim pos As Integer
   pos = cboEdit.SelStart
   cboEdit.Text = UCase(cboEdit.Text)
   cboEdit.SelStart = pos
End Sub

Private Sub cboEdit_Click()
   SaveCboEdit
End Sub

Private Sub cboEdit_LostFocus()
   SaveCboEdit
   cboEdit.Visible = False
End Sub

Private Sub SaveCboEdit()
Dim sVal As String
Dim p As Integer
sVal = cboEdit.Text
If iCol = 19 Or iCol = 20 Then
   p = InStr(sVal, " - ")
   If p > 0 Then sVal = Left(sVal, p - 1)
End If
Grid.TextMatrix(iRow, iCol) = sVal
End Sub

Private Sub optCategoria_Click()
cboEdit.Visible = False
txtEdit.Visible = False
lblCategoria.Visible = True
lblCategoria.Caption = "Categoria"
cboConsLinha.Visible = True
lblDesc.Visible = False
cboDesc.Visible = False
cboDesc.Visible = False
optPorPalavra.Visible = False
PorPalavraDupla.Visible = False
lblCodBarra.Visible = False
txtCodBarra.Visible = False
cmdLocalizar.Visible = True
Frame6.Enabled = False
optTodosPreco.Value = True
bSetSemX = True
cboConsLinha.SetFocus
End Sub

Private Sub optCodBarra_Click()
cboEdit.Visible = False
txtEdit.Visible = False
lblCategoria.Visible = False
cboConsLinha.Visible = False
lblDesc.Visible = False
cboDesc.Visible = False
cboDesc.Visible = False
optPorPalavra.Visible = False
PorPalavraDupla.Visible = False
lblCodBarra.Visible = True
txtCodBarra.Visible = True
cmdLocalizar.Visible = True
Frame6.Enabled = False
optTodosPreco.Value = True
txtCodBarra.SetFocus
txtCodBarra.Text = "[SEM CÓD. BARRA]"
frmEdicao.Enabled = False
frmEdicaoFiltros.Enabled = False
End Sub

Private Sub optComPreco_Click()
cmdLocalizar_Click
End Sub

Private Sub optDesc_Click()
cboEdit.Visible = False
txtEdit.Visible = False
lblCategoria.Visible = False
cboConsLinha.Visible = False
lblDesc.Visible = True
cboDesc.Visible = True
optPorPalavra.Visible = True
PorPalavraDupla.Visible = True
lblCodBarra.Visible = False
txtCodBarra.Visible = False
cmdLocalizar.Visible = True
Frame6.Enabled = False
optTodosPreco.Value = True
cboDesc.SetFocus
cboDesc.Text = "[SEM DESCRIÇÃO]"
frmEdicao.Enabled = False
frmEdicaoFiltros.Enabled = False
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

Private Sub optSemPreco_Click()
cmdLocalizar_Click
End Sub

Private Sub optTodos_Click()
cboEdit.Visible = False
txtEdit.Visible = False
lblCategoria.Visible = False
cboConsLinha.Visible = False
lblDesc.Visible = False
cboDesc.Visible = False
cboDesc.Visible = False
optPorPalavra.Visible = False
PorPalavraDupla.Visible = False
lblCodBarra.Visible = False
txtCodBarra.Visible = False
cmdLocalizar.Visible = False
Frame6.Enabled = True
optComPreco.Value = True
cmdAtualizar.Visible = True
optECategoria.Value = False
optETags.Value = False
optENCM.Value = False
optECBS.Value = False
optEIS.Value = False
optECFOP.Value = False
optEICMS.Value = False
optECest.Value = False
cmdEdicaoColetiva.Visible = False
txtEdicaoColetiva.Visible = False
cboEdicaoColetiva.Visible = False
lblEdicaoColetiva.Visible = False
cmdLocalizar_Click
frmEdicao.Enabled = False
frmEdicaoFiltros.Enabled = False
End Sub

Private Sub optTodosPreco_Click()
cmdLocalizar_Click
End Sub

Private Sub optTags_Click()
cboEdit.Visible = False
txtEdit.Visible = False
lblCategoria.Visible = True
lblCategoria.Caption = "Tags"
cboConsLinha.Visible = True
lblDesc.Visible = False
cboDesc.Visible = False
optPorPalavra.Visible = False
PorPalavraDupla.Visible = False
lblCodBarra.Visible = False
txtCodBarra.Visible = False
cmdLocalizar.Visible = True
Frame6.Enabled = False
optTodosPreco.Value = True
bSetSemX = True
cboConsLinha.SetFocus
End Sub

Private Sub optNCM_Click()
cboEdit.Visible = False
txtEdit.Visible = False
lblCategoria.Visible = True
lblCategoria.Caption = "NCM"
cboConsLinha.Visible = True
lblDesc.Visible = False
cboDesc.Visible = False
optPorPalavra.Visible = False
PorPalavraDupla.Visible = False
lblCodBarra.Visible = False
txtCodBarra.Visible = False
cmdLocalizar.Visible = True
Frame6.Enabled = False
optTodosPreco.Value = True
bSetSemX = True
cboConsLinha.SetFocus
End Sub

Private Sub optClassTribCBS_Click()
cboEdit.Visible = False
txtEdit.Visible = False
lblCategoria.Visible = True
lblCategoria.Caption = "Classif. CBS"
cboConsLinha.Visible = True
lblDesc.Visible = False
cboDesc.Visible = False
optPorPalavra.Visible = False
PorPalavraDupla.Visible = False
lblCodBarra.Visible = False
txtCodBarra.Visible = False
cmdLocalizar.Visible = True
Frame6.Enabled = False
optTodosPreco.Value = True
chkCBSIS.Value = 1
bSetSemX = True
cboConsLinha.SetFocus
End Sub

Private Sub optClassTribIS_Click()
cboEdit.Visible = False
txtEdit.Visible = False
lblCategoria.Visible = True
lblCategoria.Caption = "Classif. IS"
cboConsLinha.Visible = True
lblDesc.Visible = False
cboDesc.Visible = False
optPorPalavra.Visible = False
PorPalavraDupla.Visible = False
lblCodBarra.Visible = False
txtCodBarra.Visible = False
cmdLocalizar.Visible = True
Frame6.Enabled = False
optTodosPreco.Value = True
chkCBSIS.Value = 1
bSetSemX = True
cboConsLinha.SetFocus
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
      Grid_Click
   
   ElseIf KeyCode = 40 Then
      If Grid.rows = Grid.Row + 1 Then ShowMsg "VOCÊ JÁ ESTÁ NA ULTIMA LINHA !!!", vbExclamation: Exit Sub
      Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
      Grid.Row = iRow + 1
      Grid_Click
   End If
End Sub

Private Sub txtEdit_LostFocus()

If iCol = 6 Then
    txtEdit.Text = Trim(txtEdit.Text)
ElseIf iCol = 7 Then
    txtEdit.Text = Trim(txtEdit.Text)
ElseIf iCol = 8 Then
    txtEdit.Text = Replace(txtEdit.Text, ".", "")
    txtEdit.Text = Trim(txtEdit.Text)
ElseIf iCol = 11 Then
    txtEdit.Text = FormatNumber(txtEdit, 2)
ElseIf iCol = 13 Then
    txtEdit.Text = FormatNumber(txtEdit, 2)
ElseIf iCol = 15 Then
    txtEdit.Text = FormatNumber(txtEdit, 2)
ElseIf iCol = 17 Then
    txtEdit.Text = FormatNumber(txtEdit, 2)
ElseIf iCol = 18 Then
    txtEdit.Text = Replace(txtEdit.Text, ".", "")
    txtEdit.Text = Trim(txtEdit.Text)
ElseIf iCol = 19 Then
    txtEdit.Text = Trim(txtEdit.Text)
ElseIf iCol = 20 Then
    txtEdit.Text = Trim(txtEdit.Text)
End If
Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)

txtEdit.Visible = False
End Sub



