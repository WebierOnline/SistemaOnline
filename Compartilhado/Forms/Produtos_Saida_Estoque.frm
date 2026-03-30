VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Produtos_Saida_Estoque 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RETIRADA DO ESTOQUE JUSTIFICADA"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   11775
   Icon            =   "Produtos_Saida_Estoque.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   60
      TabIndex        =   15
      Top             =   840
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   11456
      _Version        =   393216
      TabHeight       =   520
      TabMaxWidth     =   3528
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "SAÍDA"
      TabPicture(0)   =   "Produtos_Saida_Estoque.frx":23D2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdNovo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdExcluir"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCancelar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdAlterar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdSalvar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdSair"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "frmCadastro"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "HISTÓRICO"
      TabPicture(1)   =   "Produtos_Saida_Estoque.frx":23EE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Grid_Historico"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "CONSULTA"
      TabPicture(2)   =   "Produtos_Saida_Estoque.frx":240A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdImprimir"
      Tab(2).Control(1)=   "cmdExibir"
      Tab(2).Control(2)=   "Grid_Consulta"
      Tab(2).Control(3)=   "Frame2"
      Tab(2).Control(4)=   "Frame8"
      Tab(2).Control(5)=   "Frame9"
      Tab(2).ControlCount=   6
      Begin VB.Frame frmCadastro 
         Caption         =   "Cadastro"
         Enabled         =   0   'False
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
         Height          =   5955
         Left            =   120
         TabIndex        =   37
         Top             =   420
         Width           =   9555
         Begin VB.ComboBox cboProduto 
            Height          =   315
            Left            =   4860
            TabIndex        =   3
            Top             =   600
            Width           =   4635
         End
         Begin VB.TextBox txtCodBarra 
            Height          =   315
            Left            =   3120
            TabIndex        =   2
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtQuantAtual 
            Alignment       =   2  'Center
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox txtSaida 
            Height          =   315
            Left            =   900
            TabIndex        =   5
            Top             =   1320
            Width           =   735
         End
         Begin VB.ComboBox cboMotivo 
            Height          =   315
            Left            =   2880
            TabIndex        =   7
            Top             =   1320
            Width           =   3495
         End
         Begin VB.TextBox txtCodProduto 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8640
            TabIndex        =   40
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtCodFunc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2160
            TabIndex        =   38
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.ComboBox cboFunc 
            Height          =   315
            Left            =   120
            TabIndex        =   1
            Top             =   600
            Width           =   2955
         End
         Begin ChamaleonBtn.chameleonButton cmdConsData 
            Height          =   315
            Left            =   2580
            TabIndex        =   39
            Tag             =   "Calendario"
            Top             =   1320
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            BTYPE           =   8
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
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Produtos_Saida_Estoque.frx":2426
            PICN            =   "Produtos_Saida_Estoque.frx":2442
            PICH            =   "Produtos_Saida_Estoque.frx":4795
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSMask.MaskEdBox mskData 
            Height          =   315
            Left            =   1680
            TabIndex        =   6
            Top             =   1320
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Funcionário"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   360
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. de Barra"
            Height          =   195
            Left            =   3120
            TabIndex        =   46
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Produto"
            Height          =   195
            Left            =   4860
            TabIndex        =   45
            Top             =   360
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Atual"
            Height          =   195
            Left            =   120
            TabIndex        =   44
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saída"
            Height          =   195
            Left            =   900
            TabIndex        =   43
            Top             =   1080
            Width           =   435
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Motivo"
            Height          =   195
            Left            =   2880
            TabIndex        =   42
            Top             =   1080
            Width           =   480
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data"
            Height          =   195
            Left            =   1680
            TabIndex        =   41
            Top             =   1080
            Width           =   345
         End
      End
      Begin VB.Frame Frame9 
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
         Height          =   1095
         Left            =   -74880
         TabIndex        =   25
         Top             =   420
         Width           =   1695
         Begin VB.OptionButton optTodos 
            Caption         =   "Todos"
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
            TabIndex        =   29
            Top             =   300
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton optCodBarra 
            Caption         =   "Cód. de Barra"
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
            TabIndex        =   28
            Top             =   480
            Width           =   1515
         End
         Begin VB.OptionButton optProduto 
            Caption         =   "Produto"
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
            TabIndex        =   27
            Top             =   660
            Width           =   1095
         End
         Begin VB.OptionButton optMensal 
            Caption         =   "Mensal"
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
            TabIndex        =   26
            Top             =   840
            Width           =   1395
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
         Height          =   1095
         Left            =   -73140
         TabIndex        =   21
         Top             =   420
         Width           =   1515
         Begin VB.OptionButton optOrdData 
            Caption         =   "Data"
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
            TabIndex        =   24
            Top             =   300
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optOrdProduto 
            Caption         =   "Produto"
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
            TabIndex        =   23
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton optVOrdFunc 
            Caption         =   "Funcionário"
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
            TabIndex        =   22
            Top             =   660
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Filtro"
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
         Height          =   1095
         Left            =   -71580
         TabIndex        =   18
         Top             =   420
         Width           =   4695
         Begin VB.TextBox txtCodBarraConsulta 
            Height          =   315
            Left            =   120
            TabIndex        =   48
            Top             =   540
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VB.ComboBox cboAno 
            Height          =   315
            Left            =   2100
            Sorted          =   -1  'True
            TabIndex        =   34
            Top             =   540
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ComboBox cboMES 
            Height          =   315
            ItemData        =   "Produtos_Saida_Estoque.frx":6AE8
            Left            =   120
            List            =   "Produtos_Saida_Estoque.frx":6AEA
            TabIndex        =   33
            Top             =   540
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.TextBox txtCodProdConsulta 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3900
            TabIndex        =   32
            Top             =   180
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.ComboBox cboConsulta 
            Height          =   315
            Left            =   120
            TabIndex        =   19
            Top             =   540
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.Label lblConsulta2 
            AutoSize        =   -1  'True
            Caption         =   "Ano"
            Height          =   195
            Left            =   2160
            TabIndex        =   35
            Top             =   300
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblConsulta 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   300
            Visible         =   0   'False
            Width           =   420
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Consulta 
         Height          =   4575
         Left            =   -74940
         TabIndex        =   16
         Top             =   1560
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   8070
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid Grid_Historico 
         Height          =   4395
         Left            =   -74940
         TabIndex        =   17
         Top             =   480
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   7752
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin ChamaleonBtn.chameleonButton cmdSair 
         Height          =   615
         Left            =   9720
         TabIndex        =   12
         Top             =   3720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
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
         MICON           =   "Produtos_Saida_Estoque.frx":6AEC
         PICN            =   "Produtos_Saida_Estoque.frx":6B08
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdSalvar 
         Height          =   615
         Left            =   9720
         TabIndex        =   8
         Top             =   1080
         Width           =   1815
         _ExtentX        =   3201
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
         MICON           =   "Produtos_Saida_Estoque.frx":6E22
         PICN            =   "Produtos_Saida_Estoque.frx":6E3E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAlterar 
         Height          =   615
         Left            =   9720
         TabIndex        =   10
         Top             =   2400
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Alterar"
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
         MICON           =   "Produtos_Saida_Estoque.frx":D708
         PICN            =   "Produtos_Saida_Estoque.frx":D724
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelar 
         Height          =   615
         Left            =   9720
         TabIndex        =   9
         Top             =   1740
         Width           =   1815
         _ExtentX        =   3201
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
         MICON           =   "Produtos_Saida_Estoque.frx":DFFE
         PICN            =   "Produtos_Saida_Estoque.frx":E01A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdExcluir 
         Height          =   615
         Left            =   9720
         TabIndex        =   11
         Top             =   3060
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Excluir"
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
         MICON           =   "Produtos_Saida_Estoque.frx":14ABE
         PICN            =   "Produtos_Saida_Estoque.frx":14ADA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdNovo 
         Height          =   615
         Left            =   9720
         TabIndex        =   0
         Top             =   420
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Novo"
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
         MICON           =   "Produtos_Saida_Estoque.frx":14DF4
         PICN            =   "Produtos_Saida_Estoque.frx":14E10
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdExibir 
         Height          =   675
         Left            =   -66780
         TabIndex        =   30
         Top             =   540
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1191
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
         MICON           =   "Produtos_Saida_Estoque.frx":15AEA
         PICN            =   "Produtos_Saida_Estoque.frx":15B06
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdImprimir 
         Height          =   675
         Left            =   -65220
         TabIndex        =   31
         Top             =   540
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1191
         BTYPE           =   3
         TX              =   "Imprimir"
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
         MICON           =   "Produtos_Saida_Estoque.frx":163E0
         PICN            =   "Produtos_Saida_Estoque.frx":163FC
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
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      ScaleHeight     =   645
      ScaleWidth      =   11625
      TabIndex        =   13
      Top             =   60
      Width           =   11655
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   10740
         TabIndex        =   49
         Top             =   180
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   300
         Picture         =   "Produtos_Saida_Estoque.frx":16716
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "RETIRADA DO ESTOQUE JUSTIFICADA"
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
         Left            =   1200
         TabIndex        =   14
         Top             =   180
         Width           =   5910
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   36
      Top             =   7320
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16431
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "16:42"
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
Attribute VB_Name = "Produtos_Saida_Estoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String
Dim r As ADODB.Recordset
Private moCombo As cComboHelper
Dim i As Integer, j As Integer

Private Function Inserir_Dados() As Boolean
Dim sSQL As String

'Comando de inclusão
sSQL = "INSERT INTO produtos_saida (" & _
   "codigo, cod_produto, data, saida, motivo, cod_funcionario, excluido) VALUES (" & _
   txtCodigo.Text & ", " & txtCodProduto.Text & ", CONVERT(DATETIME, '" & Format$(mskData.Text, ocDATA) & "', 103), " & _
   Replace(CDbl(txtSaida.Text), ",", ".") & ", '" & cboMotivo.Text & "', " & vCodFunc & ", 0);"

'Retorna o resultado da inclusão
Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Function Atualizar_Dados() As Boolean
 'Comando de atualização
sSQL = "UPDATE produtos_saida SET " & _
   "cod_produto = " & txtCodProduto.Text & ", " & _
   "data = CONVERT(DATETIME, '" & Format$(mskData.Text, ocDATA) & "', 103), " & _
   "motivo = '" & cboMotivo.Text & "', " & _
   "cod_funcionario = " & vCodFunc
'   "saida = " & Replace(CDbl(txtSaida.Text), ",", ".") & ", " & _

'Condição para atualização
sSQL = sSQL & " WHERE (codigo = " & txtCodigo.Text & ");"

'Retorna o resultado da atualização
Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Sub Auto_Numeracao()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim lRet As Long
   
   lRet = 1
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod FROM produtos_saida;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then lRet = r("cod") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
   txtCodigo.Text = lRet
End Sub

Private Sub FormatarGridConsulta(rTabela As ADODB.Recordset)
   Dim i As Integer, x As Integer, j As Integer
   
   With Grid_Consulta
      .Clear
      .Cols = 10
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 0
      .ColWidth(3) = 800
      .ColWidth(4) = 5300
      .ColWidth(5) = 700
      .ColWidth(6) = 2100
      .ColWidth(7) = 0
      .ColWidth(8) = 1400
      .ColWidth(9) = 1000
      
      For x = 0 To .Cols - 1
         .Col = x
         .Row = 0
         .CellFontBold = True
      Next
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "COD.PROD"
      .TextMatrix(0, 3) = "DATA"
      .TextMatrix(0, 4) = "PRODUTO"
      .TextMatrix(0, 5) = "SAIDA"
      .TextMatrix(0, 6) = "MOTIVO"
      .TextMatrix(0, 7) = "COD.FUNC"
      .TextMatrix(0, 8) = "FUNCIONÁRIO"
      .TextMatrix(0, 9) = "EXCLUIDO"
      
      .Redraw = False
      i = 1
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = rTabela("var_cod")
            .TextMatrix(.Rows - 1, 2) = rTabela("var_codpro")
            .TextMatrix(.Rows - 1, 3) = Format$(rTabela("var_data"), "dd/mm/yy")
            .TextMatrix(.Rows - 1, 4) = rTabela("var_desc")
            .TextMatrix(.Rows - 1, 5) = rTabela("var_saida")
            .TextMatrix(.Rows - 1, 6) = rTabela("var_mot")
            .TextMatrix(.Rows - 1, 7) = rTabela("var_codfunc")
            .TextMatrix(.Rows - 1, 8) = rTabela("var_nome")
            .TextMatrix(.Rows - 1, 9) = rTabela("varExcluido")
            
            rTabela.MoveNext
            .Rows = .Rows + 1
            i = i + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 5
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
    
'         For j = 1 To .Rows - 1
'         .Row = j
'         '.Col = 9
'          If .TextMatrix(i, 9) = "SIM" Then
'             .CellForeColor = vbRed
'          Else
'             .CellForeColor = vbBlack
'          End If
'      Next
    
    
    For i = 1 To .Rows - 1
       For j = 0 To .Cols - 1
          .Col = j
          .Row = i
          
          If .TextMatrix(i, 9) = "SIM" Then
             .CellForeColor = vbRed
          Else
             .CellForeColor = vbBlack
          End If
       Next
    Next
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
End Sub

Private Sub Limpar_Objetos()
txtCodigo.Text = ""
txtCodBarra.Text = ""
cboProduto.Text = ""
txtCodProduto.Text = ""
txtQuantAtual.Text = ""
txtSaida.Text = ""
mskData.Text = Format(Date, "dd/mm/yy")
cboMotivo.Text = ""
txtCodFunc.Text = ""
cboFunc.Text = ""
End Sub

Private Function TotalizarEstoque() As Double
Dim sSQL As String
Dim rE As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim totalEntr As Double
Dim totalSaid As Double

'Inicializa as vari+aveis
totalEntr = 0
totalSaid = 0

'Calcula as entradas
sSQL = "SELECT ISNULL(SUM(quant), 0) AS total_entradas FROM produtos_entrada_itens WHERE (codigo_produto = " & txtCodProduto & ");"
Set rE = dbData.OpenRecordset(sSQL)
If Not rE.BOF Then totalEntr = rE("total_entradas")
If rE.State <> 0 Then rE.Close
Set rE = Nothing

'Calcula as saídas
sSQL = "SELECT ISNULL(SUM(saida), 0) AS total_saidas FROM produtos_saida WHERE (cod_produto = " & txtCodProduto & ");"
Set rs = dbData.OpenRecordset(sSQL)
If Not rs.BOF Then totalSaid = rs("total_saidas")
If Not rs.State <> 0 Then rs.Close
Set rs = Nothing

'Retorna o saldo atual em estoque
TotalizarEstoque = totalEntr - totalSaid
End Function

Private Sub cboAno_GotFocus()
   Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
   Dim i As Integer
   
   cboAno.Clear
   
   iAno = Year(Date)
   FirstYear = iAno - 2
   LastYear = iAno + 2
   
   For i = FirstYear To LastYear
      cboAno.AddItem i
   Next
   
   'For x = iAno To FirstYear Step -1
   '   cboAno.AddItem x
   'Next
   '
   'iAno = iAno + 1
   'For x = iAno To LastYear
   '   cboAno.AddItem x
   'Next
End Sub

Private Sub cboConsulta_GotFocus()
Dim var_cboTexto As String
   
   If optProduto.Value = True Then
      sSQL = "SELECT DISTINCT descricao, codigo FROM produtos ORDER BY descricao;"
      Set r = dbData.OpenRecordset(sSQL)
      
      If cboConsulta.Text <> "" Then var_cboTexto = cboConsulta.Text
      cboConsulta.Clear
      
      Do While Not r.EOF
         cboConsulta.AddItem r("descricao")
         cboConsulta.ItemData(cboConsulta.NewIndex) = r("codigo")
         r.MoveNext
      Loop
      
      cboConsulta.Text = var_cboTexto
      moCombo.AttachTo cboConsulta
   
   Else
      cboConsulta.Clear
   End If
End Sub

Private Sub cboConsulta_LostFocus()
   On Error GoTo TrataErro
   
   If cboConsulta.Text = "" Then
      txtCodProdConsulta.Text = ""
      Exit Sub
   End If
   
   If cboConsulta.ListIndex = -1 Then
      txtCodProdConsulta.Text = ""
      ShowMsg "Produto não cadastrado!", vbInformation
      cboConsulta.SetFocus
      Exit Sub
   End If
   
   txtCodProdConsulta = cboConsulta.ItemData(cboConsulta.ListIndex)
   
TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cboFunc_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset

Dim varNomeAntes As String
Dim varCodAntes As String

varNomeAntes = cboFunc.Text
varCodAntes = txtCodFunc.Text

cboFunc.Clear

sSQL = "SELECT DISTINCT nome, codigo FROM funcionario ORDER BY nome;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboFunc.AddItem r("nome")
   cboFunc.ItemData(cboFunc.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

cboFunc.Text = varNomeAntes
txtCodFunc.Text = varCodAntes

moCombo.AttachTo cboFunc
End Sub


Private Sub cboFunc_LostFocus()
'On Error GoTo TrataErro

If cboFunc.Text = "" Then txtCodFunc.Text = "": Exit Sub

If cmdAlterar.Visible = False Then
   If cboFunc.ListIndex = -1 Then
      txtCodFunc.Text = ""
      vCodFunc = 0
      Exit Sub
   End If
End If

If cboFunc.ListIndex <> -1 Then
    txtCodFunc = cboFunc.ItemData(cboFunc.ListIndex)
    vCodFunc = cboFunc.ItemData(cboFunc.ListIndex)
End If
Exit Sub

'TrataErro:
'   If Err.Number = 381 Then Exit Sub
End Sub


Private Sub cboMes_GotFocus()
   Dim vMes As Integer
   
   cboMes.Clear
   
   For vMes = 1 To 12
      cboMes.AddItem StrConv(monthName(vMes), vbProperCase)
   Next
   
   moCombo.AttachTo cboMes
End Sub

Private Sub cboMotivo_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset

cboMotivo.Clear

sSQL = "SELECT motivo FROM produtos_saida GROUP BY motivo;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboMotivo.AddItem r("motivo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

moCombo.AttachTo cboMotivo
End Sub


Private Sub cboMotivo_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cboMotivo_LostFocus()
'cboMotivo.Text = TirarEspaco(cboMotivo.Text)
End Sub

Private Sub cboProduto_Click()
On Error GoTo TrataErro
Dim sSQL As String
Dim r As ADODB.Recordset

If cboProduto.Text = "" Then
   txtCodProduto.Text = ""
   Exit Sub
End If

'If cboProduto.ListIndex = -1 Then
'   txtCodProduto.Text = ""
'   ShowMsg "Produto não cadastrado!", vbInformation
'   cboProduto.SetFocus
'   Exit Sub
'End If

txtCodProduto = cboProduto.ItemData(cboProduto.ListIndex)

'mostrar o codigo de BARRA
sSQL = "SELECT codigo, cod_barra, quant_estoque FROM produtos WHERE (codigo = " & txtCodProduto.Text & ");"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
   txtCodBarra.Text = r("cod_barra")
   txtQuantAtual.Text = r("quant_estoque")
End If

If r.State <> 0 Then r.Close
Set r = Nothing
   
TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cboProduto_GotFocus()
   moCombo.AttachTo cboProduto
End Sub

Private Sub cboProduto_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboProduto_LostFocus()
Dim sSQL As String
Dim r As ADODB.Recordset

cboProduto_Click

If txtCodProduto.Text = "" Then Exit Sub

sSQL = "SELECT codigo FROM produtos WHERE (descricao = '" & cboProduto.Text & "');"
Set r = dbData.OpenRecordset(sSQL)

If r.BOF Then
   ShowMsg "Produto não encontrado", vbInformation
   txtCodProduto.Text = ""
   cboProduto.Text = ""
   txtCodBarra.Text = ""
End If

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub cmdAlterar_Click()
Dim r As ADODB.Recordset
Dim saldoEsto As Double

If txtCodigo.Text = "" Or cboProduto.Text = "" Or cboMotivo.Text = "" Then Exit Sub

If Not Atualizar_Dados Then
   ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

'saldoEsto = TotalizarEstoque

'Atualiza o saldo em estoque
'sSQL = "UPDATE produtos SET quant_estoque = " & Replace(saldoEsto, ",", ".") & " WHERE (codigo = " & txtCodProduto & ");"
'dbData.Execute sSQL

Limpar_Objetos
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
frmCadastro.Enabled = False
cmdExibir_Click
End Sub

Private Sub cmdCancelar_Click()
   Limpar_Objetos
   frmCadastro.Enabled = False
   cmdSalvar.Enabled = False
   cmdCancelar.Enabled = False
   cmdAlterar.Enabled = False
   cmdExcluir.Enabled = False
End Sub

Private Sub cmdConsData_Click()
Dim varData As Variant
Dim fCal As Calendario

varData = Empty                    'Inicializa a variável

Set fCal = New Calendario      'Cria o form de calendário
fCal.Show vbModal

varData = fCal.DateSelected    'Recupera a data selecionada

Unload fCal                           'Fecha o form
Set fCal = Nothing                   'Destrói a variável

If Not IsDate(varData) Then Exit Sub   'Valida a data
If varData = 0 Then Exit Sub

mskData = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdExcluir_Click()
Dim sSQL As String
Dim bRet As Boolean
Dim saldoEsto As Double

If txtCodigo.Text = "" Or cboProduto.Text = "" Or cboMotivo.Text = "" Then Exit Sub

'Solicita a confirmação do usuário
If ShowMsg("Excluir essa saída do estoque?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub

'Atualiza o saldo em estoque
sSQL = "UPDATE produtos_saida SET excluido = 1 WHERE (codigo = " & txtCodigo.Text & ");"
dbData.Execute sSQL

'Atualizar o estoque
dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque + " & Replace(CDbl(txtSaida.Text), ",", ".") & " WHERE (codigo = " & txtCodProduto.Text & ");"

Dim AutoNumeracao As Long

'AUTONUMERAÇÃO
sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_itens FROM Produtos_Quant;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then AutoNumeracao = r("cod_itens") + 1
If r.State <> 0 Then r.Close
Set r = Nothing

sSQL = "INSERT INTO Produtos_Quant (Codigo, COD_PRODUTO, Data, COD_ENTRADA, FORMA, QUANT, TIPO, HORA, COD_USUARIO, ESTOQUE) VALUES (" & _
   AutoNumeracao & ", " & txtCodProduto.Text & ", CONVERT(DATETIME, '" & Format(mskData.Text, ocDATA) & "', 103), 0, 'SAÍDA', " & Replace(CDbl(txtSaida.Text), ",", ".") & ", 'ADIÇÃO', '" & Format(Now, ocHRMN) & "', " & txtCodFunc.Text & ", " & Replace(CDbl(txtQuantAtual.Text), ",", ".") & ");"
dbData.Execute sSQL

Limpar_Objetos
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
frmCadastro.Enabled = False
cmdExibir_Click
End Sub

Private Sub cmdExibir_Click()
Dim sSQL As String
Dim r As ADODB.Recordset

Dim INDICE As String

If optOrdData.Value = True Then
   INDICE = "produtos_saida.data;"
ElseIf optOrdProduto.Value = True Then
   INDICE = "produtos_saida.cod_produto;"
ElseIf optVOrdFunc.Value = True Then
   INDICE = "produtos_saida.cod_funcionario;"
End If
 
If optTodos.Value = True Then
   sSQL = "SELECT produtos_saida.codigo AS var_cod, produtos_saida.cod_produto AS var_codpro, produtos_saida.data AS var_data, " & _
      "produtos.descricao AS var_desc, produtos_saida.saida AS var_saida, produtos_saida.motivo AS var_mot, " & _
      "produtos_saida.cod_funcionario AS var_codfunc, funcionario.nome AS var_nome, " & _
      "(CASE WHEN produtos_saida.excluido = 1 THEN 'SIM' ELSE 'NÃO' END) AS varExcluido " & _
      "FROM produtos_saida INNER JOIN produtos ON produtos_saida.cod_produto = produtos.codigo " & _
      "INNER JOIN funcionario ON funcionario.codigo = produtos_saida.cod_funcionario " & _
      "ORDER BY " & INDICE
   
ElseIf optCodBarra.Value = True Then
   sSQL = "SELECT produtos_saida.codigo AS var_cod, produtos_saida.cod_produto AS var_codpro, produtos_saida.data AS var_data, " & _
      "produtos.descricao AS var_desc, produtos_saida.saida AS var_saida, produtos_saida.motivo AS var_mot, " & _
      "produtos_saida.cod_funcionario AS var_codfunc, funcionario.nome AS var_nome, " & _
      "(CASE WHEN produtos_saida.excluido = 1 THEN 'SIM' ELSE 'NÃO' END) AS varExcluido " & _
      "FROM produtos_saida INNER JOIN produtos ON produtos_saida.cod_produto = produtos.codigo " & _
      "INNER JOIN funcionario ON funcionario.codigo = produtos_saida.cod_funcionario " & _
      "WHERE (produtos.cod_barra = '" & txtCodBarraConsulta.Text & "') ORDER BY " & INDICE
   
ElseIf optProduto.Value = True Then
    sSQL = "SELECT produtos_saida.codigo AS var_cod, produtos_saida.cod_produto AS var_codpro, produtos_saida.data AS var_data, " & _
      "produtos.descricao AS var_desc, produtos_saida.saida AS var_saida, produtos_saida.motivo AS var_mot, " & _
      "produtos_saida.cod_funcionario AS var_codfunc, funcionario.nome AS var_nome, " & _
      "(CASE WHEN produtos_saida.excluido = 1 THEN 'SIM' ELSE 'NÃO' END) AS varExcluido " & _
      "FROM produtos_saida INNER JOIN produtos ON produtos_saida.cod_produto = produtos.codigo " & _
      "INNER JOIN funcionario ON funcionario.codigo = produtos_saida.cod_funcionario " & _
      "WHERE (produtos_saida.cod_produto = " & txtCodProdConsulta.Text & ") ORDER BY " & INDICE
   
ElseIf optMensal.Value = True Then
   sSQL = "SELECT produtos_saida.codigo AS var_cod, produtos_saida.cod_produto AS var_codpro, produtos_saida.data AS var_data, " & _
      "produtos.descricao AS var_desc, produtos_saida.saida AS var_saida, produtos_saida.motivo AS var_mot, " & _
      "produtos_saida.cod_funcionario AS var_codfunc, funcionario.nome AS var_nome, " & _
      "(CASE WHEN produtos_saida.excluido = 1 THEN 'SIM' ELSE 'NÃO' END) AS varExcluido " & _
      "FROM produtos_saida INNER JOIN produtos ON produtos_saida.cod_produto = produtos.codigo " & _
      "INNER JOIN funcionario ON funcionario.codigo = produtos_saida.cod_funcionario " & _
      "WHERE (MONTH(produtos_saida.data) = " & cboMes.ListIndex + 1 & ") AND (YEAR(produtos_saida.data) = " & cboAno & ")  ORDER BY " & INDICE
End If
Debug.Print sSQL

Set r = dbData.OpenRecordset(sSQL)

FormatarGridConsulta r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub cmdNovo_Click()
frmCadastro.Enabled = True
cmdSalvar.Enabled = True
cmdCancelar.Enabled = True
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
Limpar_Objetos
Auto_Numeracao
If vCodFunc <> Empty Then
    txtCodFunc.Text = vCodFunc
End If
cboFunc.SetFocus
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub cmdSalvar_Click()
'On Error GoTo TrataErro

If vCodFunc = 0 Or vCodFunc = Null Then
   ShowMsg "Escolha um funcionário para dar essa saída", vbInformation
   cboFunc.SetFocus
   Exit Sub
End If

If txtQuantAtual.Text = "" Then Exit Sub

If txtQuantAtual.Text <= 0 Then
   ShowMsg "Quantidade atual do estoque é insuficiente para saída!", vbInformation
   Exit Sub
End If

If txtSaida.Text = "" Then Exit Sub

If CDbl(txtQuantAtual.Text) < CDbl(txtSaida.Text) Then
   ShowMsg "Quantidade atual é menor que saída de estoque!", vbInformation
   Exit Sub
End If

'Não é necessário consultar todos os registros antes de inserir um novo
'sSQL = "SELECT * FROM produtos_saida;"
'Set r = dbData.OpenRecordset(sSQL)

'A auto numeração do código deve ser utilizada no momento de salvar o registro
'para evitar duplicidade de código para quando houver mais de um terminal operando
'ao mesmo tempo
'AutoNumeracao

'Faz a inserção de forma direta e verifica se houve algum erro
If Not Inserir_Dados Then
   ShowMsg "Não foi possível cadastrar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
   Exit Sub
End If

'Atualizar o estoque
dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque - " & Replace(CDbl(txtSaida.Text), ",", ".") & " WHERE (codigo = " & txtCodProduto.Text & ");"

Dim AutoNumeracao As Long

'AUTONUMERAÇÃO
sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_itens FROM Produtos_Quant;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then AutoNumeracao = r("cod_itens") + 1
If r.State <> 0 Then r.Close
Set r = Nothing

sSQL = "INSERT INTO Produtos_Quant (Codigo, COD_PRODUTO, Data, COD_ENTRADA, FORMA, QUANT, TIPO, HORA, COD_USUARIO, ESTOQUE) VALUES (" & _
   AutoNumeracao & ", " & txtCodProduto.Text & ", CONVERT(DATETIME, '" & Format(mskData.Text, ocDATA) & "', 103), 0, 'SAÍDA', " & Replace(CDbl(txtSaida.Text), ",", ".") & ", 'REMOÇÃO', '" & Format(Now, ocHRMN) & "', " & txtCodFunc.Text & ", " & Replace(CDbl(txtQuantAtual.Text), ",", ".") & ");"
dbData.Execute sSQL

Limpar_Objetos
frmCadastro.Enabled = False
cmdNovo.Enabled = True
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdExibir_Click
Exit Sub

'TrataErro:
'   If Err.Number = 3022 Then
'      ShowMsg "DADOS DUPLICADO!" & vbCrLf & "Verifique se já está cadastrado.", vbInformation
'      Exit Sub
'   End If
End Sub

Private Sub Form_Load()
Set moCombo = New cComboHelper

cmdExibir_Click
Call PreencheProdutos

StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
mskData.Text = Format(Date, "dd/mm/yy")

frmCadastro.Enabled = False
cmdNovo.Enabled = True
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
SSTab1.Tab = 0
End Sub

Sub PreencheProdutos()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim var_cboTexto As String

sSQL = "SELECT DISTINCT descricao, codigo FROM produtos where ativo = 1 ORDER BY descricao;"
Set r = dbData.OpenRecordset(sSQL)

If cboProduto.Text <> "" Then var_cboTexto = cboProduto.Text
cboProduto.Clear

Do While Not r.EOF
   cboProduto.AddItem ValidateNull(r("descricao"))
   cboProduto.ItemData(cboProduto.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

cboProduto.Text = var_cboTexto
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Grid_Consulta_DblClick()
i = Grid_Consulta.Row
If Grid_Consulta.TextMatrix(i, 9) = "SIM" Then
    MsgBox "Essa saída já foi excluída", vbInformation, "Aviso do Sistema"
    cmdAlterar.Enabled = False
    cmdExcluir.Enabled = False
    cmdSalvar.Enabled = False
    cmdCancelar.Enabled = False
    cmdNovo.Enabled = True
    frmCadastro.Enabled = False
    Limpar_Objetos
    txtCodigo.Text = ""
    txtCodigo.Text = (Grid_Consulta.TextMatrix(Grid_Consulta.Row, 1))
    SSTab1.Tab = 0
Else
    cmdAlterar.Enabled = True
    cmdExcluir.Enabled = True
    cmdSalvar.Enabled = False
    cmdCancelar.Enabled = False
    frmCadastro.Enabled = True
    Limpar_Objetos
    txtCodigo.Text = ""
    txtCodigo.Text = (Grid_Consulta.TextMatrix(Grid_Consulta.Row, 1))
    SSTab1.Tab = 0
    If cmdAlterar.Enabled = True Or cmdExcluir.Enabled = True Then
        txtSaida.Enabled = False
        Label4.Enabled = False
    End If
End If
End Sub

Private Sub mskData_GotFocus()
   SelectControl mskData
End Sub

Private Sub mskData_KeyPress(KeyAscii As Integer)
   mskData.Mask = "##/##/##"
End Sub

Private Sub mskData_LostFocus()
   If mskData.Text = "" Or mskData.Text = "__/__/__" Then
      mskData.Mask = ""
      mskData.Text = ""
   Else
      If Not IsDate(mskData.Text) Then
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskData.SetFocus
      End If
   End If
End Sub

Private Sub optCodBarra_Click()
   cboMes.Text = ""
   cboAno.Text = ""
   cboConsulta.Text = ""
   lblConsulta.Caption = "Cód. de Barra"
   lblConsulta.Visible = True
   txtCodBarraConsulta.Visible = True
   cboConsulta.Visible = False
   lblConsulta2.Visible = False
   cboMes.Visible = False
   cboAno.Visible = False
   txtCodBarraConsulta.SetFocus
End Sub

Private Sub optMensal_Click()
   cboMes.Text = ""
   cboAno.Text = ""
   cboConsulta.Text = ""
   txtCodBarraConsulta.Visible = False
   lblConsulta.Caption = "Mês"
   lblConsulta.Visible = True
   cboConsulta.Visible = False
   lblConsulta2.Visible = True
   cboMes.Visible = True
   cboAno.Visible = True
   cboMes.SetFocus
End Sub

Private Sub optOrdData_Click()
   cmdExibir_Click
End Sub

Private Sub optORDProduto_Click()
   cmdExibir_Click
End Sub

Private Sub optProduto_Click()
   cboMes.Text = ""
   cboAno.Text = ""
   cboConsulta.Text = ""
   txtCodBarraConsulta.Visible = False
   lblConsulta.Caption = "Produto"
   lblConsulta.Visible = True
   cboConsulta.Visible = True
   lblConsulta2.Visible = False
   cboMes.Visible = False
   cboAno.Visible = False
   cboConsulta.SetFocus
End Sub

Private Sub optTodos_Click()
   lblConsulta.Visible = False
   cboConsulta.Visible = False
   lblConsulta2.Visible = False
   txtCodBarraConsulta.Visible = False
   cboMes.Visible = False
   cboAno.Visible = False
   cboMes.Text = ""
   cboAno.Text = ""
   cboConsulta.Text = ""
   cmdExibir_Click
End Sub

Private Sub optVOrdFunc_Click()
   cmdExibir_Click
End Sub

Private Sub txtCodBarra_Change()
   If Len(txtCodBarra.Text) = 13 Then txtCodBarra_LostFocus
End Sub

Private Sub txtCodBarra_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCodBarra_LostFocus()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodBarra.Text = "" Then Exit Sub

sSQL = "SELECT codigo AS var_codprod, descricao AS var_desc, quant_estoque FROM produtos WHERE (cod_barra = '" & txtCodBarra.Text & "') AND (ativo = 1);"
Set r = dbData.OpenRecordset(sSQL)

Debug.Print sSQL

If Not r.BOF Then
   'r.MoveLast
   txtCodProduto.Text = r("var_codprod")
   cboProduto.Text = r("var_desc")
   txtQuantAtual.Text = r("quant_estoque")
Else
   ShowMsg "Produto Inexistente!", vbCritical
   txtCodBarra.Text = ""
   txtCodBarra.SetFocus
   Exit Sub
End If

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub txtCodFunc_Change()
If txtCodFunc.Text = "" Then Exit Sub

'If cmdExcluir.Enabled = True Then
    If cboFunc.ListIndex = -1 Then
        cboFunc.Text = ""
        sSQL = "SELECT nome, codigo FROM funcionario WHERE (codigo= " & txtCodFunc.Text & ");"
        Set r = dbData.OpenRecordset(sSQL)
        If Not r.BOF Then cboFunc.Text = r("nome")
        If r.State <> 0 Then r.Close
        Set r = Nothing
    End If
'End If
End Sub

Private Sub txtCodigo_Change()
If txtCodigo.Text = "" Then Exit Sub
Dim r As ADODB.Recordset

'If cmdExcluir.Enabled = True Then
sSQL = "SELECT * FROM produtos_saida WHERE (codigo = " & txtCodigo.Text & ");"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then Mostrar_Saida r
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub Mostrar_Saida(rTabela As ADODB.Recordset)
If Not rTabela Is Nothing Then
    mskData.Text = Format(rTabela("data"), "dd/mm/yy")
    txtSaida.Text = rTabela("saida")
    cboMotivo.Text = rTabela("motivo")
    txtCodProduto.Text = rTabela("cod_produto")
    txtCodFunc.Text = rTabela("cod_funcionario")
    'vCodFunc = rTabela("cod_funcionario")
End If
End Sub
Private Sub txtCodProduto_Change()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodProduto.Text = "" Then txtQuantAtual.Text = "": Exit Sub

If txtCodProduto.Text = "" Then Exit Sub

'If cmdExcluir.Enabled = True Then
   sSQL = "SELECT codigo, cod_barra, quant_estoque, descricao FROM produtos WHERE (codigo = " & txtCodProduto.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   txtQuantAtual.Text = r("quant_estoque")
   cboProduto.Text = r("descricao")
   txtCodBarra.Text = r("cod_barra")
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
'End If
End Sub
