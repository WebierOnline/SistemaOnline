VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "ChamaleonBtn.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Configuracao_Geral 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CONFIGURAÇÕES"
   ClientHeight    =   10200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "Configuracao_Geral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10200
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   8775
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   15478
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabMaxWidth     =   2999
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "GERAL"
      TabPicture(0)   =   "Configuracao_Geral.frx":23D2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmOS"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmProdutos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame28"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame31"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame41"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame42"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame37"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame5"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   " ORÇAMENTO"
      TabPicture(1)   =   "Configuracao_Geral.frx":23EE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame29"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "PDV - À VISTA"
      TabPicture(2)   =   "Configuracao_Geral.frx":240A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "PDV - À PRAZO"
      TabPicture(3)   =   "Configuracao_Geral.frx":2426
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame15"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame5 
         Caption         =   "Incluir Preço e Quantidade no cadastro"
         Height          =   915
         Left            =   2820
         TabIndex        =   128
         Top             =   7020
         Width           =   3075
         Begin VB.OptionButton optIncluirPrecoNao 
            Caption         =   "Nã&o"
            Height          =   195
            Left            =   840
            TabIndex        =   130
            Top             =   360
            Width           =   675
         End
         Begin VB.OptionButton optIncluirPrecoSim 
            Caption         =   "Sim"
            Height          =   195
            Left            =   180
            TabIndex        =   129
            Top             =   360
            Width           =   675
         End
         Begin ChamaleonBtn.chameleonButton cmdIncluirPreco 
            Height          =   315
            Left            =   1620
            TabIndex        =   131
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "&Atualizar"
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
            MICON           =   "Configuracao_Geral.frx":2442
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
      Begin VB.Frame Frame37 
         Caption         =   "Confirmar fechamento de venda ?"
         Height          =   915
         Left            =   120
         TabIndex        =   123
         Top             =   7020
         Width           =   2655
         Begin VB.OptionButton optConfFechaSim 
            Caption         =   "Sim"
            Height          =   195
            Left            =   120
            TabIndex        =   126
            Top             =   420
            Width           =   675
         End
         Begin VB.OptionButton optConfFechaNao 
            Caption         =   "Nã&o"
            Height          =   195
            Left            =   780
            TabIndex        =   124
            Top             =   420
            Width           =   675
         End
         Begin ChamaleonBtn.chameleonButton chameleonButton32 
            Height          =   315
            Left            =   1500
            TabIndex        =   125
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "&Atualizar"
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
            MICON           =   "Configuracao_Geral.frx":245E
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
      Begin VB.Frame Frame42 
         Caption         =   "Identificar Maquina do PDV"
         Height          =   915
         Left            =   2640
         TabIndex        =   113
         Top             =   6060
         Width           =   2955
         Begin VB.OptionButton optIDEMaqNao 
            Caption         =   "Nã&o"
            Height          =   195
            Left            =   720
            TabIndex        =   114
            Top             =   420
            Width           =   675
         End
         Begin VB.OptionButton optIDEMaqSim 
            Caption         =   "Si&m"
            Height          =   195
            Left            =   120
            TabIndex        =   115
            Top             =   420
            Width           =   675
         End
         Begin ChamaleonBtn.chameleonButton chameleonButton38 
            Height          =   315
            Left            =   2040
            TabIndex        =   116
            Top             =   360
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "&Atualizar"
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
            MICON           =   "Configuracao_Geral.frx":247A
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
      Begin VB.Frame Frame41 
         Caption         =   "Identificação (PDV)"
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
         Left            =   120
         TabIndex        =   108
         Top             =   5100
         Width           =   3795
         Begin VB.TextBox txtTipoIndPDV 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2820
            TabIndex        =   112
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.OptionButton optIDELogin 
            Caption         =   "Login"
            Height          =   195
            Left            =   120
            TabIndex        =   110
            Top             =   360
            Width           =   915
         End
         Begin VB.OptionButton optIDEFunc 
            Caption         =   "Cód. Funcionário"
            Height          =   195
            Left            =   1140
            TabIndex        =   109
            Top             =   360
            Width           =   1575
         End
         Begin ChamaleonBtn.chameleonButton chameleonButton37 
            Height          =   315
            Left            =   2760
            TabIndex        =   111
            Top             =   300
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "&Atualizar"
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
            MICON           =   "Configuracao_Geral.frx":2496
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
      Begin VB.Frame Frame31 
         Caption         =   "Vender c/ Estoque Negativo"
         Height          =   915
         Left            =   120
         TabIndex        =   104
         Top             =   6060
         Width           =   2475
         Begin VB.OptionButton optVendNegNao 
            Caption         =   "Nã&o"
            Height          =   195
            Left            =   720
            TabIndex        =   105
            Top             =   420
            Width           =   675
         End
         Begin VB.OptionButton optVendNegSim 
            Caption         =   "Si&m"
            Height          =   195
            Left            =   120
            TabIndex        =   106
            Top             =   420
            Width           =   675
         End
         Begin ChamaleonBtn.chameleonButton chameleonButton36 
            Height          =   315
            Left            =   1440
            TabIndex        =   107
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "&Atualizar"
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
            MICON           =   "Configuracao_Geral.frx":24B2
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
      Begin VB.Frame Frame29 
         Caption         =   "PDV"
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
         Height          =   8175
         Left            =   -74880
         TabIndex        =   78
         Top             =   420
         Width           =   5595
         Begin VB.Frame Frame30 
            Caption         =   "Orçamento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6675
            Left            =   60
            TabIndex        =   79
            Top             =   240
            Width           =   5475
            Begin VB.Frame Frame40 
               Caption         =   "Desconto automático ?"
               Height          =   915
               Left            =   120
               TabIndex        =   96
               Top             =   240
               Width           =   2595
               Begin VB.TextBox txtValorDescORC 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   97
                  Top             =   480
                  Width           =   1335
               End
               Begin ChamaleonBtn.chameleonButton chameleonButton35 
                  Height          =   315
                  Left            =   1500
                  TabIndex        =   98
                  Top             =   480
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  BTYPE           =   3
                  TX              =   "&Atualizar"
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
                  MICON           =   "Configuracao_Geral.frx":24CE
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VB.Label Label24 
                  Caption         =   "Valor (%)"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   99
                  Top             =   240
                  Width           =   1095
               End
            End
            Begin VB.Frame Frame36 
               Caption         =   "Confirmar Impressão ?"
               Height          =   675
               Left            =   2760
               TabIndex        =   92
               Top             =   1200
               Width           =   2655
               Begin VB.OptionButton optConfImpSimORC 
                  Caption         =   "Si&m"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   94
                  Top             =   300
                  Width           =   675
               End
               Begin VB.OptionButton optConfImpNaoORC 
                  Caption         =   "Nã&o"
                  Height          =   195
                  Left            =   780
                  TabIndex        =   93
                  Top             =   300
                  Width           =   615
               End
               Begin ChamaleonBtn.chameleonButton chameleonButton31 
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   95
                  Top             =   240
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  BTYPE           =   3
                  TX              =   "&Atualizar"
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
                  MICON           =   "Configuracao_Geral.frx":24EA
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
            Begin VB.Frame Frame35 
               Caption         =   "Imprimir:"
               Height          =   675
               Left            =   120
               TabIndex        =   87
               Top             =   1980
               Width           =   2835
               Begin VB.TextBox txtTipoImpressaoORC 
                  Appearance      =   0  'Flat
                  Height          =   285
                  Left            =   840
                  TabIndex        =   90
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.OptionButton optImpPedidoORC 
                  Caption         =   "Folha"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   89
                  Top             =   300
                  Width           =   735
               End
               Begin VB.OptionButton optImpCupomGuiORC 
                  Caption         =   "Cupom"
                  Height          =   195
                  Left            =   780
                  TabIndex        =   88
                  Top             =   300
                  Width           =   795
               End
               Begin ChamaleonBtn.chameleonButton chameleonButton30 
                  Height          =   315
                  Left            =   1680
                  TabIndex        =   91
                  Top             =   240
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  BTYPE           =   3
                  TX              =   "&Atualizar"
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
                  MICON           =   "Configuracao_Geral.frx":2506
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
            Begin VB.Frame Frame33 
               Caption         =   "Número de Cópias"
               Height          =   675
               Left            =   3000
               TabIndex        =   84
               Top             =   1980
               Width           =   2415
               Begin VB.TextBox txtNumCopiaORC 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   85
                  Top             =   300
                  Width           =   735
               End
               Begin ChamaleonBtn.chameleonButton chameleonButton28 
                  Height          =   315
                  Left            =   1200
                  TabIndex        =   86
                  Top             =   300
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  BTYPE           =   3
                  TX              =   "&Atualizar"
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
                  MICON           =   "Configuracao_Geral.frx":2522
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
            Begin VB.Frame Frame32 
               Caption         =   "Imprimir ?"
               Height          =   675
               Left            =   120
               TabIndex        =   80
               Top             =   1200
               Width           =   2595
               Begin VB.OptionButton optImpNaoORC 
                  Caption         =   "Nã&o"
                  Height          =   195
                  Left            =   780
                  TabIndex        =   82
                  Top             =   300
                  Width           =   615
               End
               Begin VB.OptionButton optImpSimORC 
                  Caption         =   "Si&m"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   81
                  Top             =   300
                  Width           =   675
               End
               Begin ChamaleonBtn.chameleonButton chameleonButton27 
                  Height          =   315
                  Left            =   1500
                  TabIndex        =   83
                  Top             =   240
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  BTYPE           =   3
                  TX              =   "&Atualizar"
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
                  MICON           =   "Configuracao_Geral.frx":253E
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
         End
      End
      Begin VB.Frame Frame28 
         Caption         =   "Bloquear Cliente em Débito"
         Height          =   915
         Left            =   3960
         TabIndex        =   73
         Top             =   5100
         Width           =   2955
         Begin VB.TextBox txtQuantDiasBloqueiar 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1380
            TabIndex        =   77
            Top             =   360
            Width           =   555
         End
         Begin VB.OptionButton optBloqueiarClienteSim 
            Caption         =   "Si&m"
            Height          =   195
            Left            =   720
            TabIndex        =   75
            Top             =   420
            Width           =   615
         End
         Begin VB.OptionButton optBloqueiarClienteNao 
            Caption         =   "Nã&o"
            Height          =   195
            Left            =   60
            TabIndex        =   74
            Top             =   420
            Width           =   675
         End
         Begin ChamaleonBtn.chameleonButton chameleonButton25 
            Height          =   315
            Left            =   2040
            TabIndex        =   76
            Top             =   360
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "&Atualizar"
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
            MICON           =   "Configuracao_Geral.frx":255A
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
      Begin VB.Frame Frame4 
         Caption         =   "Principal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   68
         Top             =   3120
         Width           =   6855
         Begin VB.TextBox txtCaminhoCupom 
            Height          =   315
            Left            =   120
            TabIndex        =   100
            Top             =   1260
            Width           =   4155
         End
         Begin VB.TextBox txtCaminho 
            Height          =   315
            Left            =   120
            TabIndex        =   69
            Top             =   480
            Width           =   4155
         End
         Begin ChamaleonBtn.chameleonButton chameleonButton10 
            Height          =   315
            Left            =   4380
            TabIndex        =   70
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "&Atualizar"
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
            MICON           =   "Configuracao_Geral.frx":2576
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   4680
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin ChamaleonBtn.chameleonButton chameleonButton26 
            Height          =   315
            Left            =   4380
            TabIndex        =   101
            Top             =   1260
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "&Atualizar"
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
            MICON           =   "Configuracao_Geral.frx":2592
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblProcurarCupom 
            AutoSize        =   -1  'True
            Caption         =   "[ &PROCURAR ]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   210
            Left            =   3210
            TabIndex        =   103
            Top             =   1620
            Width           =   1095
         End
         Begin VB.Label Label25 
            Caption         =   "Logomarca do Cupom"
            Height          =   195
            Left            =   120
            TabIndex        =   102
            Top             =   1020
            Width           =   1695
         End
         Begin VB.Label Label9 
            Caption         =   "Fundo (PDV)"
            Height          =   195
            Left            =   120
            TabIndex        =   72
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblProcurar 
            AutoSize        =   -1  'True
            Caption         =   "[ &PROCURAR ]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   210
            Left            =   3180
            TabIndex        =   71
            Top             =   840
            Width           =   1095
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "PDV"
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
         Height          =   8175
         Left            =   -74880
         TabIndex        =   42
         Top             =   420
         Width           =   5595
         Begin VB.Frame Frame17 
            Caption         =   "A Prazo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6675
            Left            =   60
            TabIndex        =   43
            Top             =   240
            Width           =   5475
            Begin VB.Frame Frame27 
               Caption         =   "Desconto automático ?"
               Height          =   915
               Left            =   120
               TabIndex        =   64
               Top             =   240
               Width           =   2595
               Begin VB.TextBox txtValorDescAP 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   65
                  Top             =   480
                  Width           =   1335
               End
               Begin ChamaleonBtn.chameleonButton chameleonButton24 
                  Height          =   315
                  Left            =   1500
                  TabIndex        =   66
                  Top             =   480
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  BTYPE           =   3
                  TX              =   "&Atualizar"
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
                  MICON           =   "Configuracao_Geral.frx":25AE
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VB.Label Label19 
                  Caption         =   "Valor (%)"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   67
                  Top             =   240
                  Width           =   1095
               End
            End
            Begin VB.Frame Frame23 
               Caption         =   "Confirmar Impressão ?"
               Height          =   675
               Left            =   2760
               TabIndex        =   60
               Top             =   1200
               Width           =   2655
               Begin VB.OptionButton optConfImpSimAP 
                  Caption         =   "Si&m"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   62
                  Top             =   300
                  Width           =   675
               End
               Begin VB.OptionButton optConfImpNaoAP 
                  Caption         =   "Nã&o"
                  Height          =   195
                  Left            =   780
                  TabIndex        =   61
                  Top             =   300
                  Width           =   615
               End
               Begin ChamaleonBtn.chameleonButton chameleonButton20 
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   63
                  Top             =   240
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  BTYPE           =   3
                  TX              =   "&Atualizar"
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
                  MICON           =   "Configuracao_Geral.frx":25CA
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
            Begin VB.Frame Frame22 
               Caption         =   "Imprimir:"
               Height          =   675
               Left            =   120
               TabIndex        =   55
               Top             =   2820
               Width           =   2835
               Begin VB.TextBox txtTipoImpressaoAP 
                  Appearance      =   0  'Flat
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   58
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.OptionButton optImpPedidoAP 
                  Caption         =   "folha"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   57
                  Top             =   300
                  Width           =   735
               End
               Begin VB.OptionButton optImpCupomGuiAP 
                  Caption         =   "Cupom"
                  Height          =   195
                  Left            =   900
                  TabIndex        =   56
                  Top             =   300
                  Width           =   855
               End
               Begin ChamaleonBtn.chameleonButton chameleonButton19 
                  Height          =   315
                  Left            =   1800
                  TabIndex        =   59
                  Top             =   240
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  BTYPE           =   3
                  TX              =   "&Atualizar"
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
                  MICON           =   "Configuracao_Geral.frx":25E6
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
            Begin VB.Frame Frame20 
               Caption         =   "Número de Cópias"
               Height          =   675
               Left            =   2820
               TabIndex        =   52
               Top             =   1980
               Width           =   2535
               Begin VB.TextBox txtNumCopiaAP 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   53
                  Top             =   300
                  Width           =   735
               End
               Begin ChamaleonBtn.chameleonButton chameleonButton17 
                  Height          =   315
                  Left            =   1440
                  TabIndex        =   54
                  Top             =   300
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  BTYPE           =   3
                  TX              =   "&Atualizar"
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
                  MICON           =   "Configuracao_Geral.frx":2602
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
            Begin VB.Frame Frame19 
               Caption         =   "Imprimir ?"
               Height          =   675
               Left            =   120
               TabIndex        =   48
               Top             =   1200
               Width           =   2595
               Begin VB.OptionButton optImpNaoAP 
                  Caption         =   "Nã&o"
                  Height          =   195
                  Left            =   780
                  TabIndex        =   50
                  Top             =   300
                  Width           =   615
               End
               Begin VB.OptionButton optImpSimAP 
                  Caption         =   "Si&m"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   49
                  Top             =   300
                  Width           =   675
               End
               Begin ChamaleonBtn.chameleonButton chameleonButton16 
                  Height          =   315
                  Left            =   1500
                  TabIndex        =   51
                  Top             =   240
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  BTYPE           =   3
                  TX              =   "&Atualizar"
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
                  MICON           =   "Configuracao_Geral.frx":261E
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
            Begin VB.Frame Frame18 
               Caption         =   "Nota Adicional de Entrega  ?"
               Height          =   675
               Left            =   120
               TabIndex        =   44
               Top             =   1980
               Width           =   2655
               Begin VB.OptionButton optEntregaNaoAP 
                  Caption         =   "Nã&o"
                  Height          =   195
                  Left            =   780
                  TabIndex        =   46
                  Top             =   300
                  Width           =   615
               End
               Begin VB.OptionButton optEntregaSimAP 
                  Caption         =   "Si&m"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   45
                  Top             =   300
                  Width           =   675
               End
               Begin ChamaleonBtn.chameleonButton chameleonButton15 
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   47
                  Top             =   240
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  BTYPE           =   3
                  TX              =   "&Atualizar"
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
                  MICON           =   "Configuracao_Geral.frx":263A
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
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Financeiro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   120
         TabIndex        =   25
         Top             =   420
         Width           =   3675
         Begin VB.TextBox txtJurosMes 
            Height          =   315
            Left            =   120
            TabIndex        =   27
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtJuroDia 
            Height          =   315
            Left            =   1260
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   480
            Width           =   1095
         End
         Begin ChamaleonBtn.chameleonButton cmdAlterar 
            Height          =   315
            Left            =   2460
            TabIndex        =   28
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "&Atualizar"
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
            MICON           =   "Configuracao_Geral.frx":2656
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label1 
            Caption         =   "Juros/Mês(%)"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Juros/Dia(%)"
            Height          =   195
            Left            =   1260
            TabIndex        =   29
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame frmProdutos 
         Caption         =   "Tipo de Empresa"
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
         Left            =   120
         TabIndex        =   22
         Top             =   1380
         Width           =   4635
         Begin VB.ComboBox cboTipoEmpresa 
            Height          =   315
            Left            =   120
            TabIndex        =   127
            Top             =   300
            Width           =   3315
         End
         Begin VB.TextBox txtTipoCadastroProduto 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3420
            TabIndex        =   23
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
         Begin ChamaleonBtn.chameleonButton chameleonButton1 
            Height          =   315
            Left            =   3480
            TabIndex        =   24
            Top             =   300
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "&Atualizar"
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
            MICON           =   "Configuracao_Geral.frx":2672
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
      Begin VB.Frame frmOS 
         Caption         =   "Habilitar Ordem de Serviços"
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
         Left            =   120
         TabIndex        =   18
         Top             =   2280
         Width           =   6855
         Begin VB.TextBox txtTipoOS 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4620
            TabIndex        =   122
            Top             =   60
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1800
            ScaleHeight     =   345
            ScaleWidth      =   3885
            TabIndex        =   118
            Top             =   300
            Width           =   3915
            Begin VB.OptionButton optOSinformatica 
               Caption         =   "Informática"
               Height          =   195
               Left            =   2280
               TabIndex        =   121
               Top             =   60
               Width           =   1215
            End
            Begin VB.OptionButton optOSmotos 
               Caption         =   "Motos"
               Height          =   195
               Left            =   1380
               TabIndex        =   120
               Top             =   60
               Width           =   855
            End
            Begin VB.OptionButton OptOScarros 
               Caption         =   "Automoveis"
               Height          =   195
               Left            =   120
               TabIndex        =   119
               Top             =   60
               Value           =   -1  'True
               Width           =   1215
            End
         End
         Begin VB.OptionButton optNaoOS 
            Caption         =   "Nã&o"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optSimOS 
            Caption         =   "Si&m"
            Height          =   195
            Left            =   960
            TabIndex        =   19
            Top             =   360
            Width           =   915
         End
         Begin ChamaleonBtn.chameleonButton chameleonButton2 
            Height          =   315
            Left            =   5760
            TabIndex        =   21
            Top             =   300
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "&Atualizar"
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
            MICON           =   "Configuracao_Geral.frx":268E
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
      Begin VB.Frame Frame2 
         Caption         =   "PDV"
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
         Height          =   8175
         Left            =   -74880
         TabIndex        =   3
         Top             =   420
         Width           =   5595
         Begin VB.Frame Frame3 
            Caption         =   "A Vista"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6675
            Left            =   60
            TabIndex        =   4
            Top             =   240
            Width           =   5475
            Begin VB.Frame Frame14 
               Caption         =   "Nota Adicional de Entrega  ?"
               Height          =   675
               Left            =   120
               TabIndex        =   38
               Top             =   1980
               Width           =   2655
               Begin VB.OptionButton optEntregaSim 
                  Caption         =   "Si&m"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   40
                  Top             =   300
                  Width           =   675
               End
               Begin VB.OptionButton optEntregaNao 
                  Caption         =   "Nã&o"
                  Height          =   195
                  Left            =   780
                  TabIndex        =   39
                  Top             =   300
                  Width           =   615
               End
               Begin ChamaleonBtn.chameleonButton chameleonButton13 
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   41
                  Top             =   240
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  BTYPE           =   3
                  TX              =   "&Atualizar"
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
                  MICON           =   "Configuracao_Geral.frx":26AA
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
            Begin VB.Frame Frame13 
               Caption         =   "Imprimir ?"
               Height          =   675
               Left            =   120
               TabIndex        =   34
               Top             =   1200
               Width           =   2595
               Begin VB.OptionButton optImpSim 
                  Caption         =   "Si&m"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   36
                  Top             =   300
                  Width           =   675
               End
               Begin VB.OptionButton optImpNao 
                  Caption         =   "Nã&o"
                  Height          =   195
                  Left            =   780
                  TabIndex        =   35
                  Top             =   300
                  Width           =   615
               End
               Begin ChamaleonBtn.chameleonButton chameleonButton12 
                  Height          =   315
                  Left            =   1500
                  TabIndex        =   37
                  Top             =   240
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  BTYPE           =   3
                  TX              =   "&Atualizar"
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
                  MICON           =   "Configuracao_Geral.frx":26C6
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
            Begin VB.Frame Frame12 
               Caption         =   "Número de Cópias"
               Height          =   675
               Left            =   2820
               TabIndex        =   31
               Top             =   1980
               Width           =   2535
               Begin VB.TextBox txtNumCopia 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   32
                  Top             =   300
                  Width           =   735
               End
               Begin ChamaleonBtn.chameleonButton chameleonButton11 
                  Height          =   315
                  Left            =   1440
                  TabIndex        =   33
                  Top             =   300
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  BTYPE           =   3
                  TX              =   "&Atualizar"
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
                  MICON           =   "Configuracao_Geral.frx":26E2
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
            Begin VB.Frame Frame7 
               Caption         =   "Imprimir:"
               Height          =   675
               Left            =   120
               TabIndex        =   13
               Top             =   2820
               Width           =   2835
               Begin VB.OptionButton optImpCupomGui 
                  Caption         =   "Cupom"
                  Height          =   195
                  Left            =   840
                  TabIndex        =   16
                  Top             =   300
                  Width           =   795
               End
               Begin VB.OptionButton optImpPedido 
                  Caption         =   "Folha"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   15
                  Top             =   300
                  Width           =   735
               End
               Begin VB.TextBox txtTipoImpressaoAV 
                  Appearance      =   0  'Flat
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   14
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin ChamaleonBtn.chameleonButton chameleonButton6 
                  Height          =   315
                  Left            =   1680
                  TabIndex        =   17
                  Top             =   240
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  BTYPE           =   3
                  TX              =   "&Atualizar"
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
                  MICON           =   "Configuracao_Geral.frx":26FE
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
            Begin VB.Frame Frame6 
               Caption         =   "Confirmar Impressão ?"
               Height          =   675
               Left            =   2760
               TabIndex        =   9
               Top             =   1200
               Width           =   2655
               Begin VB.OptionButton optConfImpNao 
                  Caption         =   "Nã&o"
                  Height          =   195
                  Left            =   780
                  TabIndex        =   11
                  Top             =   300
                  Width           =   615
               End
               Begin VB.OptionButton optConfImpSim 
                  Caption         =   "Si&m"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   10
                  Top             =   300
                  Width           =   675
               End
               Begin ChamaleonBtn.chameleonButton chameleonButton5 
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   12
                  Top             =   240
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  BTYPE           =   3
                  TX              =   "&Atualizar"
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
                  MICON           =   "Configuracao_Geral.frx":271A
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
            Begin VB.Frame Frame11 
               Caption         =   "Desconto automático ?"
               Height          =   915
               Left            =   120
               TabIndex        =   5
               Top             =   240
               Width           =   2595
               Begin VB.TextBox txtValorDescAV 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   6
                  Top             =   480
                  Width           =   1335
               End
               Begin ChamaleonBtn.chameleonButton chameleonButton9 
                  Height          =   315
                  Left            =   1500
                  TabIndex        =   7
                  Top             =   480
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  BTYPE           =   3
                  TX              =   "&Atualizar"
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
                  MICON           =   "Configuracao_Geral.frx":2736
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VB.Label Label10 
                  Caption         =   "Valor (%)"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   8
                  Top             =   240
                  Width           =   1095
               End
            End
         End
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   60
      ScaleHeight     =   945
      ScaleWidth      =   7065
      TabIndex        =   0
      Top             =   60
      Width           =   7095
      Begin VB.Image Image1 
         Height          =   900
         Left            =   300
         Picture         =   "Configuracao_Geral.frx":2752
         Top             =   0
         Width           =   900
      End
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CONFIGURAÇÕES GERAIS"
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
         Left            =   1380
         TabIndex        =   1
         Top             =   300
         Width           =   4005
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   117
      Top             =   9930
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8414
            Text            =   "Online.Info - Informática"
            TextSave        =   "Online.Info - Informática"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "00:02"
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
Attribute VB_Name = "Configuracao_Geral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moCombo As cComboHelper
Private Caminho As String

Dim oCfg As ConfigItem

Private Sub AP_Mostrar_Copia()
   Set oCfg = sysConfig("COPIAS_AP")
   txtNumCopiaAP.Text = oCfg.Value
   Set oCfg = Nothing
End Sub

Private Sub ORC_Mostrar_Copia()
   Set oCfg = sysConfig("COPIAS_ORC")
   txtNumCopiaORC.Text = oCfg.Value
   Set oCfg = Nothing
End Sub

Private Sub AV_Mostrar_Copia()
   Set oCfg = sysConfig("COPIAS_AV")
   txtNumCopia.Text = oCfg.Value
   Set oCfg = Nothing
End Sub

Private Sub Mostrar_LogoCupom()
   Set oCfg = sysConfig("LOGO_CUPOM")
   txtCaminhoCupom.Text = oCfg.Value
   Set oCfg = Nothing
End Sub

Private Sub Mostrar_Fundo()
   Set oCfg = sysConfig("FUNDO_PDV")
   txtCaminho.Text = oCfg.Value
   Set oCfg = Nothing
End Sub

Private Sub Mostrar_OS()
   Set oCfg = sysConfig("OS")
   
   If CBool(oCfg.Value) = True Then
      optSimOS.Value = True
   Else
      optNaoOS.Value = True
   End If
   
   Set oCfg = sysConfig("TIPO_OS")
   
   Select Case oCfg.Value
      Case "CARROS": OptOScarros.Value = True
      Case "MOTOS": optOSmotos.Value = True
      Case "INFOR": optOSinformatica.Value = True
   End Select
   
   Set oCfg = Nothing
End Sub

Private Sub Mostrar_Tipo_Identificacao()
   Set oCfg = sysConfig("IDENT_PDV")
   
   txtTipoIndPDV.Text = oCfg.Value
   If txtTipoIndPDV.Text = 1 Then
      optIDELogin.Value = True
   ElseIf txtTipoIndPDV.Text = 2 Then
      optIDEFunc.Value = True
   End If
   
   Set oCfg = Nothing
End Sub

Private Sub Mostrar_Tipo_Empresa()
   Set oCfg = sysConfig("TIPO_EMPRESA")
   
   txtTipoCadastroProduto.Text = oCfg.Value
   If txtTipoCadastroProduto.Text = 1 Then
      cboTipoEmpresa.Text = "Varejo"
   ElseIf txtTipoCadastroProduto.Text = 2 Then
      cboTipoEmpresa.Text = "Farmacia"
   ElseIf txtTipoCadastroProduto.Text = 3 Then
      cboTipoEmpresa.Text = "Restaurante/Lannchonete"
   ElseIf txtTipoCadastroProduto.Text = 4 Then
      cboTipoEmpresa.Text = "Sapataria/Vestuário"
   ElseIf txtTipoCadastroProduto.Text = 5 Then
      cboTipoEmpresa.Text = "Autopeça/Motopeça"
   Else
      Exit Sub
   End If
   
   Set oCfg = Nothing
End Sub

Private Sub MostrarIdentMaquina()
   Set oCfg = sysConfig("IDENT_MAQ")
   
   If CBool(oCfg.Value) = True Then
      optIDEMaqSim.Value = True
   Else
      optIDEMaqNao.Value = True
   End If
   
   Set oCfg = Nothing
End Sub

Private Sub MostrarIncluirPreco()
   Set oCfg = sysConfig("INCLUIR_PRECO")
   
   If CBool(oCfg.Value) = True Then
      optIncluirPrecoSim.Value = True
   Else
      optIncluirPrecoNao.Value = True
   End If
   
   Set oCfg = Nothing
End Sub

Private Sub MostrarEstoqueNegativo()
   Set oCfg = sysConfig("ESTOQUE_NEGATIVO")
   
   If CBool(oCfg.Value) = True Then
      optVendNegSim.Value = True
   Else
      optVendNegNao.Value = True
   End If
   
   Set oCfg = Nothing
End Sub
Private Sub AP_MostrarEntrega()
   Set oCfg = sysConfig("ENTREGA_AP")
    
   If CBool(oCfg.Value) = True Then
      optEntregaSimAP.Value = True
   Else
      optEntregaNaoAP.Value = True
   End If
   
   Set oCfg = Nothing
End Sub

Private Sub AV_MostrarEntrega()
   Set oCfg = sysConfig("ENTREGA_AV")
   
   If CBool(oCfg.Value) = True Then
      optEntregaSim.Value = True
   Else
      optEntregaNao.Value = True
   End If
   
   Set oCfg = Nothing
End Sub

Private Sub AP_MostrarImp()
   Set oCfg = sysConfig("IMP_AP")
   
   If CBool(oCfg.Value) = True Then
      optImpSimAP.Value = True
   Else
      optImpNaoAP.Value = True
   End If
   
   Set oCfg = Nothing
End Sub

Private Sub ORC_MostrarImp()
   Set oCfg = sysConfig("IMP_ORC")
      
   If CBool(oCfg.Value) = True Then
      optImpSimORC.Value = True
   Else
      optImpNaoORC.Value = True
   End If
   
   Set oCfg = Nothing
End Sub

Private Sub AV_MostrarImp()
   Set oCfg = sysConfig("IMP_AV")
   
   If CBool(oCfg.Value) = True Then
      optImpSim.Value = True
   Else
      optImpNao.Value = True
   End If
   
   Set oCfg = Nothing
End Sub

Private Sub AP_MostrarConfImpressao()
   Set oCfg = sysConfig("CONF_IMPRESSAO_AP")
   
   If CBool(oCfg.Value) = True Then
      optConfImpSimAP.Value = True
   Else
      optConfImpNaoAP.Value = True
   End If
   
   Set oCfg = Nothing
End Sub

Private Sub MostrarConfBloqueioCliente()
   Set oCfg = sysConfig("BLOQUEIAR_CLIENTE")
   
   If CBool(oCfg.Value) = True Then
      Set oCfg = sysConfig("DIAS_BLOQUEIO")
      txtQuantDiasBloqueiar.Text = oCfg.Value
      optBloqueiarClienteSim.Value = True
   Else
      optBloqueiarClienteNao.Value = True
   End If
   
   Set oCfg = Nothing
End Sub

Private Sub ORC_MostrarConfImpressao()
   Set oCfg = sysConfig("CONF_IMPRESSAO_ORC")
   
   If CBool(oCfg.Value) = True Then
      optConfImpSimORC.Value = True
   Else
      optConfImpNaoORC.Value = True
   End If
   
   Set oCfg = Nothing
End Sub

Private Sub AV_MostrarConfImpressao()
   Set oCfg = sysConfig("CONF_IMPRESSAO_AV")
   
   If CBool(oCfg.Value) = True Then
      optConfImpSim.Value = True
   Else
      optConfImpNao.Value = True
   End If
   
   Set oCfg = Nothing
End Sub

Private Sub AP_MostrarFecharImpressao()
   Set oCfg = sysConfig("CONF_FECHAMENTO_AP")
   
   If CBool(oCfg.Value) = True Then
      optConfFechaSim.Value = True
   Else
      optConfFechaNao.Value = True
   End If
   
   Set oCfg = Nothing
End Sub

Private Sub ORC_MostrarImp2()
   'Set oCfg = sysConfig("IMP2_MAQ_ORC")
   'cboAVMaqCupGuiORC.Text = oCfg.Value
   'Set oCfg = sysConfig("IMP2_COMPART_ORC")
   'cboAVImpCupGuiORC.Text = oCfg.Value
   'Set oCfg = Nothing
End Sub

Private Sub AV_MostrarImp2()
   'Set oCfg = sysConfig("IMP2_MAQ_AV")
   'cboAVMaqCupGui.Text = oCfg.Value
   'Set oCfg = sysConfig("IMP2_COMPART_AV")
   'cboAVImpCupGui.Text = oCfg.Value
   'Set oCfg = Nothing
End Sub

Private Sub AP_MostrarImp2()
   'Set oCfg = sysConfig("IMP2_MAQ_AP")
   'cboAVMaqCupGuiAP.Text = oCfg.Value
   'Set oCfg = sysConfig("IMP2_COMPART_AP")
   'cboAVImpCupGuiAP.Text = oCfg.Value
End Sub

Private Sub AP_MostrarImp3()
   'Set oCfg = sysConfig("IMP3_MAQ_AP")
   'cboAVMaqCupSerAP.Text = oCfg.Value
   'Set oCfg = sysConfig("IMP3_COMPART_AP")
   'cboAVImpCupSerAP.Text = oCfg.Value
   'Set oCfg = Nothing
End Sub

Private Sub ORC_MostrarImp3()
   'Set oCfg = sysConfig("IMP3_MAQ_ORC")
   'cboAVMaqCupSerORC.Text = oCfg.Value
   'Set oCfg = sysConfig("IMP3_COMPART_ORC")
   'cboAVImpCupSerORC.Text = oCfg.Value
   'Set oCfg = Nothing
End Sub

Private Sub AV_MostrarImp3()
   'Set oCfg = sysConfig("IMP3_MAQ_AV")
   'cboAVMaqCupSer.Text = oCfg.Value
   'Set oCfg = sysConfig("IMP3_COMPART_AV")
   'cboAVImpCupSer.Text = oCfg.Value
   'Set oCfg = Nothing
End Sub

Private Sub AP_MostrarTipoImpressao()
   Set oCfg = sysConfig("IMPRIMIR_AP")
   
   txtTipoImpressaoAP.Text = oCfg.Value
   If txtTipoImpressaoAP.Text = 1 Then
      optImpPedidoAP.Value = True
   ElseIf txtTipoImpressaoAP.Text = 2 Then
      optImpCupomGui.Value = True
   ElseIf txtTipoImpressaoAP.Text = 3 Then
      optImpCupomGuiAP.Value = True
   Else
      Exit Sub
   End If
   
   Set oCfg = Nothing
End Sub

Private Sub ORC_MostrarTipoImpressao()
   Set oCfg = sysConfig("IMPRIMIR_ORC")
   
   txtTipoImpressaoORC.Text = oCfg.Value
   If txtTipoImpressaoORC.Text = 1 Then
      optImpPedidoORC.Value = True
   ElseIf txtTipoImpressaoORC.Text = 2 Then
      optImpCupomGui.Value = True
   ElseIf txtTipoImpressaoORC.Text = 3 Then
      optImpCupomGuiORC.Value = True
   Else
      Exit Sub
   End If
End Sub

Private Sub AV_MostrarTipoImpressao()
   Set oCfg = sysConfig("IMPRIMIR_AV")
   
   txtTipoImpressaoAV.Text = oCfg.Value
   If txtTipoImpressaoAV.Text = 1 Then
      optImpPedido.Value = True
   ElseIf txtTipoImpressaoAV.Text = 2 Then
      optImpCupomGui.Value = True
   ElseIf txtTipoImpressaoAV.Text = 3 Then
      optImpCupomGui.Value = True
   Else
      Exit Sub
   End If
End Sub

Private Sub cboTipoEmpresa_GotFocus()
Dim var_Texto As String
var_Texto = cboTipoEmpresa.Text

   cboTipoEmpresa.Clear
   cboTipoEmpresa.AddItem "Varejo"
   cboTipoEmpresa.AddItem "Farmacia"
   cboTipoEmpresa.AddItem "Restaurante/Lannchonete"
   cboTipoEmpresa.AddItem "Sapataria/Vestuário"
   cboTipoEmpresa.AddItem "Autopeça/Motopeça"

  
cboTipoEmpresa.Text = var_Texto

End Sub


Private Sub cboTipoEmpresa_Validate(Cancel As Boolean)
If cboTipoEmpresa.Text = "Varejo" Then
   txtTipoCadastroProduto.Text = "1"
ElseIf cboTipoEmpresa.Text = "Farmacia" Then
   txtTipoCadastroProduto.Text = "2"
ElseIf cboTipoEmpresa.Text = "Restaurante/Lannchonete" Then
   txtTipoCadastroProduto.Text = "3"
ElseIf cboTipoEmpresa.Text = "Sapataria/Vestuário" Then
   txtTipoCadastroProduto.Text = "4"
ElseIf cboTipoEmpresa.Text = "Autopeça/Motopeça" Then
   txtTipoCadastroProduto.Text = "5"
Else
   txtTipoCadastroProduto.Text = "1"
End If
End Sub


Private Sub chameleonButton1_Click()
   Dim sSQL As String
   
   If txtTipoCadastroProduto.Text = "" Then Exit Sub
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtTipoCadastroProduto.Text & "' WHERE (config_nome = 'TIPO_EMPRESA');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("TIPO_EMPRESA").Value = txtTipoCadastroProduto.Text
End Sub

Private Sub chameleonButton10_Click()
   Dim sSQL As String
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtCaminho.Text & "' WHERE (config_nome = 'FUNDO_PDV');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("FUNDO_PDV").Value = txtCaminho.Text
End Sub

Private Sub chameleonButton11_Click()
   Dim sSQL As String
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtNumCopia.Text & "' WHERE (config_nome = 'COPIAS_AV');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("COPIAS_AV").Value = txtNumCopia.Text
End Sub

Private Sub chameleonButton12_Click()
   Dim sSQL As String, bOpt As Boolean
   
   If optImpSim.Value = True Then
       bOpt = True
   ElseIf optImpNao.Value = True Then
       bOpt = False
   End If
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'IMP_AV');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("IMP_AV").Value = Abs(bOpt)
End Sub

Private Sub chameleonButton13_Click()
   Dim sSQL As String, bOpt As Boolean
   
   If optEntregaSim.Value = True Then
       bOpt = True
   ElseIf optEntregaNao.Value = True Then
       bOpt = False
   End If
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'ENTREGA_AV');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("ENTREGA_AV").Value = Abs(bOpt)
End Sub

Private Sub chameleonButton15_Click()
   Dim sSQL As String, bOpt As Boolean
   
   If optEntregaSimAP.Value = True Then
       bOpt = True
   ElseIf optEntregaNaoAP.Value = True Then
       bOpt = False
   End If
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'ENTREGA_AP');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("ENTREGA_AP").Value = Abs(bOpt)
End Sub

Private Sub chameleonButton16_Click()
   Dim sSQL As String, bOpt As Boolean
   
   If optImpSimAP.Value = True Then
       bOpt = True
   ElseIf optImpNaoAP.Value = True Then
       bOpt = False
   End If
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'IMP_AP');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("IMP_AP").Value = Abs(bOpt)
End Sub

Private Sub chameleonButton17_Click()
   Dim sSQL As String
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtNumCopiaAP.Text & "' WHERE (config_nome = 'COPIAS_AP');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("COPIAS_AP").Value = txtNumCopiaAP.Text
End Sub

Private Sub chameleonButton19_Click()
   Dim sSQL As String
   
   If txtTipoCadastroProduto.Text = "" Then Exit Sub
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtTipoImpressaoAP.Text & "' WHERE (config_nome = 'IMPRIMIR_AP');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("IMPRIMIR_AP").Value = txtTipoImpressaoAP.Text
End Sub

Private Sub chameleonButton2_Click()
   Dim sSQL As String, bOpt As Boolean
   
   If optSimOS.Value = True Then
       bOpt = True
   ElseIf optNaoOS.Value = True Then
       bOpt = False
   End If
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'OS');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("OS").Value = Abs(bOpt)
   
   If txtTipoOS.Text = "" Then Exit Sub
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtTipoOS.Text & "' WHERE (config_nome = 'TIPO_OS');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("TIPO_OS").Value = txtTipoOS.Text
End Sub

Private Sub chameleonButton20_Click()
   Dim sSQL As String, bOpt As Boolean
   
   If optConfImpSimAP.Value = True Then
       bOpt = True
   ElseIf optConfImpNaoAP.Value = True Then
       bOpt = False
   End If
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'CONF_IMPRESSAO_AP');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("CONF_IMPRESSAO_AP").Value = Abs(bOpt)
End Sub

Private Sub chameleonButton24_Click()
   Dim sSQL As String
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtValorDescAP.Text & "' WHERE (config_nome = 'DESC_AP');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("DESC_AP").Value = txtValorDescAP.Text
End Sub

Private Sub chameleonButton25_Click()
   Dim sSQL As String, bOpt As Boolean
   
   If optBloqueiarClienteSim.Value = True Then
       bOpt = True
   ElseIf optBloqueiarClienteNao.Value = True Then
       bOpt = False
   End If
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'BLOQUEIAR_CLIENTE');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("BLOQUEIAR_CLIENTE").Value = Abs(bOpt)
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtQuantDiasBloqueiar.Text & "' WHERE (config_nome = 'DIAS_BLOQUEIO');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("DIAS_BLOQUEIO").Value = txtQuantDiasBloqueiar.Text
   
   txtQuantDiasBloqueiar.Enabled = False
End Sub

Private Sub chameleonButton26_Click()
   Dim sSQL As String
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtCaminhoCupom.Text & "' WHERE (config_nome = 'LOGO_CUPOM');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("LOGO_CUPOM").Value = txtCaminhoCupom.Text
End Sub

Private Sub chameleonButton27_Click()
   Dim sSQL As String, bOpt As Boolean
   
   If optImpSimORC.Value = True Then
       bOpt = True
   ElseIf optImpNaoORC.Value = True Then
       bOpt = False
   End If
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'IMP_ORC');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("IMP_ORC").Value = Abs(bOpt)
End Sub

Private Sub chameleonButton28_Click()
   Dim sSQL As String
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtNumCopiaORC.Text & "' WHERE (config_nome = 'COPIAS_ORC');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("COPIAS_ORC").Value = txtNumCopiaORC.Text
End Sub

Private Sub chameleonButton3_Click()

End Sub

Private Sub chameleonButton30_Click()
   Dim sSQL As String
   
   If txtTipoCadastroProduto.Text = "" Then Exit Sub
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtTipoImpressaoORC.Text & "' WHERE (config_nome = 'IMPRIMIR_ORC');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("IMPRIMIR_ORC").Value = txtTipoImpressaoORC.Text
End Sub

Private Sub chameleonButton31_Click()
   Dim sSQL As String, bOpt As Boolean
   
   If optConfImpSimORC.Value = True Then
       bOpt = True
   ElseIf optConfImpNaoORC.Value = True Then
       bOpt = False
   End If
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'CONF_IMPRESSAO_ORC');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("CONF_IMPRESSAO_ORC").Value = Abs(bOpt)
End Sub

Private Sub chameleonButton32_Click()
   Dim sSQL As String, bOpt As Boolean
   
   If optConfFechaSim.Value = True Then
       bOpt = True
   ElseIf optConfFechaNao.Value = True Then
       bOpt = False
   End If
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'CONF_FECHAMENTO_ORC');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("CONF_FECHAMENTO_ORC").Value = Abs(bOpt)
End Sub

Private Sub chameleonButton35_Click()
   Dim sSQL As String
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtValorDescORC.Text & "' WHERE (config_nome = 'DESC_ORC');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("DESC_ORC").Value = txtValorDescORC.Text
End Sub

Private Sub chameleonButton36_Click()
   Dim sSQL As String, bOpt As Boolean
   
   If optVendNegSim.Value = True Then
       bOpt = True
   ElseIf optVendNegNao.Value = True Then
       bOpt = False
   End If
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'ESTOQUE_NEGATIVO');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("ESTOQUE_NEGATIVO").Value = Abs(bOpt)
End Sub

Private Sub chameleonButton37_Click()
   Dim sSQL As String, bOpt As Boolean
   
   If txtTipoIndPDV.Text = "" Then Exit Sub
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtTipoIndPDV.Text & "' WHERE (config_nome = 'IDENT_PDV');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("IDENT_PDV").Value = txtTipoIndPDV.Text
End Sub

Private Sub chameleonButton38_Click()
   Dim sSQL As String, bOpt As Boolean
   
   If optIDEMaqSim.Value = True Then
       bOpt = True
   ElseIf optIDEMaqNao.Value = True Then
       bOpt = False
   End If
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'IDENT_MAQ');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("IDENT_MAQ").Value = Abs(bOpt)
End Sub

Private Sub chameleonButton5_Click()
   Dim sSQL As String, bOpt As Boolean
   
   If optConfImpSim.Value = True Then
       bOpt = True
   ElseIf optConfImpNao.Value = True Then
       bOpt = False
   End If
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'CONF_IMPRESSAO_AV');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("CONF_IMPRESSAO_AV").Value = Abs(bOpt)
End Sub

Private Sub chameleonButton6_Click()
   Dim sSQL As String
   
   If txtTipoCadastroProduto.Text = "" Then Exit Sub
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtTipoImpressaoAV.Text & "' WHERE (config_nome = 'IMPRIMIR_AV');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("IMPRIMIR_AV").Value = txtTipoImpressaoAV.Text
End Sub

Private Sub chameleonButton9_Click()
   Dim sSQL As String
    
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtValorDescAV.Text & "' WHERE (config_nome = 'DESC_AV');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("DESC_AV").Value = txtValorDescAV.Text
End Sub

Private Sub cmdAlterar_Click()
   Dim sSQL As String
    
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtJurosMes.Text & "' WHERE (config_nome = 'JUROS_MES');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("JUROS_MES").Value = txtJurosMes.Text
    
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & txtJuroDia.Text & "' WHERE (config_nome = 'JUROS_DIA');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("JUROS_DIA").Value = txtJuroDia.Text
End Sub

Private Sub cmdIncluirPreco_Click()
   Dim sSQL As String, bOpt As Boolean
   
   If optIncluirPrecoSim.Value = True Then
       bOpt = True
   ElseIf optIncluirPrecoNao.Value = True Then
       bOpt = False
   End If
   
   'Atualiza a base de dados
   sSQL = "UPDATE configuracao SET config_valor = '" & Abs(bOpt) & "' WHERE (config_nome = 'INCLUIR_PRECO');"
   dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("INCLUIR_PRECO").Value = Abs(bOpt)

End Sub

Private Sub Form_Load()
   Set moCombo = New cComboHelper
   
   Mostrar_Fundo
   Mostrar_LogoCupom
   Mostrar_Tipo_Empresa
   Mostrar_OS
   MostrarConfBloqueioCliente
   Mostrar_Dados_Juros
   MostrarEstoqueNegativo
   MostrarIncluirPreco
   Mostrar_Tipo_Identificacao
   MostrarIdentMaquina
   
   AV_Mostrar_Desc
   AV_MostrarImp
   AV_MostrarConfImpressao
   AV_MostrarEntrega
   AV_Mostrar_Copia
   AV_MostrarTipoImpressao
   
   AP_Mostrar_Desc
   AP_MostrarImp
   AP_MostrarConfImpressao
   AP_MostrarFecharImpressao
   AP_MostrarEntrega
   AP_Mostrar_Copia
   AP_MostrarTipoImpressao
   
   ORC_Mostrar_Desc
   ORC_MostrarImp
   ORC_MostrarConfImpressao
   ORC_Mostrar_Copia
   
   ORC_MostrarTipoImpressao
   
   SSTab1.Tab = 0
   StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
   Caminho = appPathApp
End Sub

Private Sub ORC_Mostrar_Desc()
   Set oCfg = sysConfig("DESC_ORC")
   txtValorDescORC.Text = Format(oCfg.Value, "##,##0.00")
   Set oCfg = Nothing
End Sub

Private Sub AV_Mostrar_Desc()
   Set oCfg = sysConfig("DESC_AV")
   txtValorDescAV.Text = Format(oCfg.Value, "##,##0.00")
   Set oCfg = Nothing
End Sub

Private Sub AP_Mostrar_Desc()
   Set oCfg = sysConfig("DESC_AP")
   txtValorDescAP.Text = Format(oCfg.Value, "##,##0.00")
   Set oCfg = Nothing
End Sub

Private Sub Mostrar_Dados_Juros()
   Set oCfg = sysConfig("JUROS_MES")
   txtJurosMes.Text = Format(oCfg.Value, "##,##0.00")
   Set oCfg = sysConfig("JUROS_DIA")
   txtJuroDia.Text = Format(oCfg.Value, "##,##0.00")
   Set oCfg = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub lblProcurar_Click()
   CommonDialog1.Filter = "Imagens JPG(*.jpg)|*.jpg"
   CommonDialog1.ShowOpen
   txtCaminho.Text = CommonDialog1.FileName
End Sub

Private Sub lblProcurarCupom_Click()
   CommonDialog1.Filter = "Imagens JPG(*.jpg)|*.jpg"
   CommonDialog1.ShowOpen
   txtCaminhoCupom.Text = CommonDialog1.FileName
End Sub

Private Sub optBloqueiarClienteNao_Click()
   txtQuantDiasBloqueiar.Enabled = False
End Sub

Private Sub optBloqueiarClienteSim_Click()
   txtQuantDiasBloqueiar.Enabled = True
   'txtQuantDiasBloqueiar.SetFocus
End Sub

Private Sub optIDEFunc_Click()
   txtTipoIndPDV.Text = "2"
End Sub

Private Sub optIDELogin_Click()
   txtTipoIndPDV.Text = "1"
End Sub

Private Sub optImpCupomGui_Click()
   txtTipoImpressaoAV.Text = "3"
End Sub

Private Sub optImpCupomGuiAP_Click()
   txtTipoImpressaoAP.Text = "3"
End Sub

Private Sub optImpCupomGuiORC_Click()
   txtTipoImpressaoORC.Text = "3"
End Sub

Private Sub optImpPedido_Click()
   txtTipoImpressaoAV.Text = "1"
End Sub

Private Sub optImpPedidoAP_Click()
   txtTipoImpressaoAP.Text = "1"
End Sub

Private Sub optImpPedidoORC_Click()
   txtTipoImpressaoORC.Text = "1"
End Sub

Private Sub OptOScarros_Click()
   txtTipoOS.Text = "CARROS"
End Sub

Private Sub optOSinformatica_Click()
   txtTipoOS.Text = "INFOR"
End Sub

Private Sub optOSmotos_Click()
   txtTipoOS.Text = "MOTOS"
End Sub

Private Sub optSemEntrada_Click()

End Sub

Private Sub SSTab1_DblClick()
   Dim sSQL As String
   
   'Atualiza a base de dados
   'sSQL = "UPDATE configuracao SET config_valor = '" & cboAVMaqCupSer.Text & "' WHERE (config_nome = 'IMP3_MAQ_AV');"
   'dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   sysConfig("IMP3_MAQ_AV").Value = txtJurosMes.Text
   
   'Atualiza a base de dados
   'sSQL = "UPDATE configuracao SET config_valor = '" & cboAVImpCupSer.Text & "' WHERE (config_nome = 'IMP3_COMPART_AV');"
   'dbData.Execute sSQL
   
   'Atualiza a configuração carregada na memória
   'sysConfig("IMP3_COMPART_AV").Value = cboAVImpCupSer.Text
End Sub

Private Sub txtJuroDia_GotFocus()
   SelectControl txtJuroDia
End Sub

Private Sub txtJurosMes_GotFocus()
   SelectControl txtJurosMes
End Sub

Private Sub txtJurosMes_LostFocus()
   If txtJurosMes.Text = "" Then Exit Sub
   txtJurosMes.Text = Format(txtJurosMes, "##,##0.00")
   txtJuroDia.Text = Format((txtJurosMes / 30), "##,##0.00")
End Sub

Private Sub txtNumCopia_GotFocus()
   SelectControl txtNumCopia
End Sub

Private Sub txtNumCopiaAP_GotFocus()
   SelectControl txtNumCopiaAP
End Sub

Private Sub txtNumCopiaORC_GotFocus()
   SelectControl txtNumCopiaORC
End Sub

Private Sub txtValorDescAP_GotFocus()
   SelectControl txtValorDescAP
End Sub

Private Sub txtValorDescAP_LostFocus()
   If txtValorDescAP.Text = "" Then Exit Sub
   txtValorDescAP.Text = Format(txtValorDescAP, "##,##0.00")
End Sub

Private Sub txtValorDescAV_GotFocus()
   SelectControl txtValorDescAV
End Sub

Private Sub txtValorDescAV_LostFocus()
   If txtValorDescAV.Text = "" Then Exit Sub
   txtValorDescAV.Text = Format(txtValorDescAV, "##,##0.00")
End Sub

Private Sub txtValorDescORC_GotFocus()
   SelectControl txtValorDescORC
End Sub

Private Sub txtValorDescORC_LostFocus()
   If txtValorDescORC.Text = "" Then Exit Sub
   txtValorDescORC.Text = Format(txtValorDescORC, "##,##0.00")
End Sub
