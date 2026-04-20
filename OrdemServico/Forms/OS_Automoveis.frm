VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form OS_Automoveis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ORDEM DE SERVIÇO"
   ClientHeight    =   10275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14700
   Icon            =   "OS_Automoveis.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10275
   ScaleWidth      =   14700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab2 
      Height          =   8835
      Left            =   8160
      TabIndex        =   89
      Top             =   1080
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   15584
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabMaxWidth     =   3528
      TabCaption(0)   =   "CONSULTA"
      TabPicture(0)   =   "OS_Automoveis.frx":23D2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Grid"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmLocalizar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.Frame Frame7 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1680
         TabIndex        =   104
         Top             =   480
         Width           =   4695
         Begin VB.OptionButton optFechados 
            Caption         =   "Fechados"
            Height          =   195
            Left            =   1920
            TabIndex        =   107
            Top             =   300
            Width           =   1275
         End
         Begin VB.OptionButton OptAbertos 
            Caption         =   "Abertos"
            Height          =   195
            Left            =   960
            TabIndex        =   106
            Top             =   300
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OptTodos 
            Caption         =   "Todos"
            Height          =   195
            Left            =   120
            TabIndex        =   105
            Top             =   300
            Width           =   855
         End
      End
      Begin VB.Frame frmLocalizar 
         Caption         =   "Situação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   60
         TabIndex        =   96
         Top             =   480
         Width           =   1575
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   3420
            ScaleHeight     =   375
            ScaleWidth      =   3075
            TabIndex        =   102
            Top             =   180
            Width           =   3075
         End
         Begin VB.OptionButton optStatusTerminado 
            Caption         =   "&Terminado"
            Height          =   195
            Left            =   120
            TabIndex        =   101
            Top             =   1260
            Width           =   1395
         End
         Begin VB.OptionButton optStatusAguardando 
            Caption         =   "A&guardando"
            Height          =   195
            Left            =   120
            TabIndex        =   100
            Top             =   1020
            Width           =   1395
         End
         Begin VB.OptionButton optStatusExecucao 
            Caption         =   "Em E&xecução"
            Height          =   195
            Left            =   120
            TabIndex        =   99
            Top             =   780
            Width           =   1395
         End
         Begin VB.OptionButton optStatusComecar 
            Caption         =   "À C&omeçar"
            Height          =   195
            Left            =   120
            TabIndex        =   98
            Top             =   540
            Width           =   1395
         End
         Begin VB.OptionButton optStatusTodos 
            Caption         =   "&Todos"
            Height          =   195
            Left            =   120
            TabIndex        =   97
            Top             =   300
            Value           =   -1  'True
            Width           =   1395
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   180
            TabIndex        =   103
            Top             =   240
            Width           =   45
         End
      End
      Begin VB.Frame Frame6 
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
         Height          =   915
         Left            =   1680
         TabIndex        =   90
         Top             =   1140
         Width           =   4695
         Begin VB.CheckBox optLocalizarCliente 
            Caption         =   "Cliente"
            Height          =   195
            Left            =   1080
            TabIndex        =   109
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox optLocalizarCodOS 
            Caption         =   "Cód.:"
            Height          =   195
            Left            =   60
            TabIndex        =   108
            Top             =   240
            Width           =   795
         End
         Begin VB.TextBox txtLocalizarCodOS 
            Height          =   315
            Left            =   60
            TabIndex        =   93
            Top             =   480
            Width           =   975
         End
         Begin VB.ComboBox cboLocalizar 
            Height          =   315
            Left            =   1080
            TabIndex        =   92
            Top             =   480
            Width           =   2925
         End
         Begin VB.TextBox txtCodClienteLocalizar 
            Appearance      =   0  'Flat
            Height          =   225
            Left            =   3420
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   555
         End
         Begin ChamaleonBtn.chameleonButton cmdExibir2 
            Height          =   315
            Left            =   4020
            TabIndex        =   94
            Top             =   480
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Exibir"
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
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "OS_Automoveis.frx":23EE
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
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   6675
         Left            =   60
         TabIndex        =   95
         Top             =   2100
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   11774
         _Version        =   393216
         SelectionMode   =   1
         Appearance      =   0
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   60
      ScaleHeight     =   945
      ScaleWidth      =   14565
      TabIndex        =   85
      Top             =   60
      Width           =   14595
      Begin VB.TextBox txtCodOS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   13380
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   12900
         TabIndex        =   88
         Top             =   120
         Width           =   420
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   120
         Picture         =   "OS_Automoveis.frx":240A
         Top             =   0
         Width           =   900
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ORDEM DE SERVIÇOS"
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
         Left            =   1140
         TabIndex        =   86
         Top             =   300
         Width           =   3360
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8865
      Left            =   60
      TabIndex        =   110
      Top             =   1080
      Width           =   8010
      _ExtentX        =   14129
      _ExtentY        =   15637
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   452
      TabMaxWidth     =   2293
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "ENTRADA"
      TabPicture(0)   =   "OS_Automoveis.frx":293A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "chameleonButton1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdGarantia"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCadastrarCliente"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCadastrarAcessorios"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdCancelarEntrada"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdNovo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdApagar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdAlterar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdGerarEntrada"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdImprimirEntrada"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "frmPrincipal"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Picture14"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "frmSecundario"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Picture9"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "SERVIÇOS"
      TabPicture(1)   =   "OS_Automoveis.frx":2956
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "PRODUTOS"
      TabPicture(2)   =   "OS_Automoveis.frx":2972
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "PGTO"
      TabPicture(3)   =   "OS_Automoveis.frx":298E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "frmVendaAvista"
      Tab(3).Control(1)=   "frmVendaPrazo"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   " "
      TabPicture(4)   =   "OS_Automoveis.frx":29AA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   " "
      TabPicture(5)   =   "OS_Automoveis.frx":29C6
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lblQuantFiltro"
      Tab(5).ControlCount=   1
      Begin VB.PictureBox Picture9 
         Height          =   735
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   6135
         TabIndex        =   201
         Top             =   6540
         Width           =   6195
         Begin VB.ComboBox cboTipoOS 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            TabIndex        =   19
            Top             =   300
            Width           =   2355
         End
         Begin MSMask.MaskEdBox mskDataSaida 
            Height          =   315
            Left            =   2460
            TabIndex        =   20
            Top             =   300
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   12648384
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskHoraSaida 
            Height          =   315
            Left            =   3960
            TabIndex        =   21
            Top             =   300
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   12648384
            PromptChar      =   "_"
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de OS."
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
            TabIndex        =   204
            Top             =   60
            Width           =   1035
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hora de Saída"
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
            Left            =   3960
            TabIndex        =   203
            Top             =   60
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data de Saída"
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
            Left            =   2460
            TabIndex        =   202
            Top             =   60
            Width           =   1260
         End
      End
      Begin VB.Frame frmVendaPrazo 
         BackColor       =   &H00C0C0FF&
         Caption         =   "VENDA À PRAZO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4755
         Left            =   -74820
         TabIndex        =   166
         Top             =   480
         Visible         =   0   'False
         Width           =   7515
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0C0FF&
            Height          =   555
            Left            =   120
            TabIndex        =   183
            Top             =   780
            Width           =   7275
            Begin VB.OptionButton optAvulso 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Avulso"
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
               TabIndex        =   65
               TabStop         =   0   'False
               Top             =   240
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton optPromissoria 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Promissória"
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
               Left            =   1200
               TabIndex        =   66
               TabStop         =   0   'False
               Top             =   240
               Width           =   1335
            End
            Begin VB.OptionButton optCheque 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Cheque"
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
               Left            =   2640
               TabIndex        =   67
               TabStop         =   0   'False
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0C0FF&
            Height          =   1815
            Left            =   3960
            TabIndex        =   176
            Top             =   1380
            Width           =   3435
            Begin VB.PictureBox Picture6 
               BackColor       =   &H00C0C0FF&
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   1440
               ScaleHeight     =   210
               ScaleWidth      =   1035
               TabIndex        =   178
               Top             =   660
               Width           =   1035
               Begin VB.OptionButton optDescRS 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "R$"
                  Height          =   210
                  Left            =   60
                  TabIndex        =   69
                  TabStop         =   0   'False
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   555
               End
               Begin VB.OptionButton optDescPorc 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "%"
                  Height          =   210
                  Left            =   600
                  TabIndex        =   70
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   435
               End
            End
            Begin VB.TextBox txtTotalDesc 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   75
               TabStop         =   0   'False
               Text            =   "0,00"
               Top             =   1320
               Width           =   1455
            End
            Begin VB.TextBox txtDesc 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2520
               TabIndex        =   71
               TabStop         =   0   'False
               Top             =   600
               Width           =   855
            End
            Begin VB.TextBox txtSubTotal 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1860
               Locked          =   -1  'True
               TabIndex        =   68
               TabStop         =   0   'False
               Text            =   "0,00"
               Top             =   240
               Width           =   1515
            End
            Begin VB.TextBox txtAcresc 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2520
               TabIndex        =   74
               TabStop         =   0   'False
               Top             =   960
               Width           =   855
            End
            Begin VB.PictureBox Picture4 
               BackColor       =   &H00C0C0FF&
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   1440
               ScaleHeight     =   210
               ScaleWidth      =   1035
               TabIndex        =   177
               Top             =   1020
               Width           =   1035
               Begin VB.OptionButton optAscrescRS 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "R$"
                  Height          =   210
                  Left            =   60
                  TabIndex        =   72
                  TabStop         =   0   'False
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   555
               End
               Begin VB.OptionButton optAscrescPorc 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "%"
                  Height          =   210
                  Left            =   600
                  TabIndex        =   73
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   435
               End
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SubTotal:"
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
               Left            =   960
               TabIndex        =   182
               Top             =   300
               Width           =   840
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total:"
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
               Left            =   1380
               TabIndex        =   181
               Top             =   1380
               Width           =   510
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Desc.:"
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
               Left            =   900
               TabIndex        =   180
               Top             =   660
               Width           =   570
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Acresc.:"
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
               Left            =   720
               TabIndex        =   179
               Top             =   1020
               Width           =   780
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0C0FF&
            Height          =   975
            Left            =   120
            TabIndex        =   168
            Top             =   3240
            Width           =   7275
            Begin VB.ComboBox cboPrazo 
               Height          =   315
               Left            =   1200
               TabIndex        =   77
               Text            =   "30"
               Top             =   480
               Width           =   675
            End
            Begin VB.TextBox txtEntrada 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               TabIndex        =   76
               Top             =   480
               Width           =   1035
            End
            Begin VB.TextBox txtValorRest 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   78
               Text            =   "0"
               Top             =   480
               Width           =   1035
            End
            Begin VB.ComboBox cboQuantParc 
               Height          =   315
               Left            =   3000
               TabIndex        =   79
               Text            =   "1"
               Top             =   480
               Width           =   735
            End
            Begin VB.TextBox txtValorParc 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3780
               Locked          =   -1  'True
               TabIndex        =   80
               Text            =   "0"
               Top             =   480
               Width           =   1155
            End
            Begin MSMask.MaskEdBox mskInicio 
               Height          =   315
               Left            =   4980
               TabIndex        =   81
               Top             =   480
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskTermino 
               Height          =   315
               Left            =   6060
               TabIndex        =   82
               Top             =   480
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin VB.Label lblEntrada 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Entrada:"
               Height          =   195
               Left            =   120
               TabIndex        =   175
               Top             =   240
               Width           =   600
            End
            Begin VB.Label lblInicio 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Inicio:"
               Height          =   195
               Left            =   4980
               TabIndex        =   174
               Top             =   240
               Width           =   420
            End
            Begin VB.Label lblQuantParc 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Prazo:"
               Height          =   195
               Left            =   1200
               TabIndex        =   173
               Top             =   240
               Width           =   450
            End
            Begin VB.Label lblValorParc 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor Rest."
               Height          =   195
               Left            =   1920
               TabIndex        =   172
               Top             =   240
               Width           =   780
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Quant:"
               Height          =   195
               Left            =   3000
               TabIndex        =   171
               Top             =   240
               Width           =   480
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor Parc.:"
               Height          =   195
               Left            =   3780
               TabIndex        =   170
               Top             =   240
               Width           =   825
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Termino:"
               Height          =   195
               Left            =   6060
               TabIndex        =   169
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00C0C0FF&
            Height          =   555
            Left            =   120
            TabIndex        =   167
            Top             =   240
            Width           =   7275
            Begin VB.TextBox txtCodFuncAP 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               TabIndex        =   63
               Top             =   180
               Width           =   1035
            End
            Begin VB.TextBox txtFuncAP 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   180
               Width           =   5955
            End
         End
         Begin ChamaleonBtn.chameleonButton cmdFinalizar 
            Height          =   315
            Left            =   5640
            TabIndex        =   83
            Top             =   4260
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Finalizar"
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
            MICON           =   "OS_Automoveis.frx":29E2
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
            Height          =   315
            Left            =   6540
            TabIndex        =   84
            Top             =   4260
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Cancelar"
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
            MICON           =   "OS_Automoveis.frx":29FE
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
      Begin VB.Frame frmVendaAvista 
         BackColor       =   &H00C0FFC0&
         Caption         =   "VENDA À VISTA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4755
         Left            =   -74820
         TabIndex        =   184
         Top             =   420
         Visible         =   0   'False
         Width           =   7515
         Begin VB.Frame Frame9 
            BackColor       =   &H00C0FFC0&
            Height          =   555
            Left            =   120
            TabIndex        =   196
            Top             =   960
            Width           =   7275
            Begin VB.PictureBox frmCartao 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               FillColor       =   &H00004000&
               FillStyle       =   0  'Solid
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   4320
               ScaleHeight     =   285
               ScaleWidth      =   2745
               TabIndex        =   197
               Top             =   160
               Visible         =   0   'False
               Width           =   2775
               Begin VB.OptionButton optDebito 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Débito"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00008000&
                  Height          =   195
                  Left            =   180
                  TabIndex        =   49
                  Top             =   60
                  Value           =   -1  'True
                  Width           =   1035
               End
               Begin VB.OptionButton optCredito 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Crédito"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00008000&
                  Height          =   195
                  Left            =   1200
                  TabIndex        =   50
                  Top             =   60
                  Width           =   1035
               End
            End
            Begin VB.OptionButton optAVcheque 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Cheque"
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
               Left            =   1260
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optAVcartao 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Cartão de Crédito"
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
               Left            =   2340
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   240
               Width           =   1995
            End
            Begin VB.OptionButton optAVdinheiro 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Dinheiro"
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
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   240
               Value           =   -1  'True
               Width           =   1095
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00C0FFC0&
            Height          =   855
            Left            =   120
            TabIndex        =   193
            Top             =   3420
            Width           =   7275
            Begin VB.TextBox txtRecebido 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   360
               Left            =   120
               TabIndex        =   59
               TabStop         =   0   'False
               Top             =   420
               Width           =   1875
            End
            Begin VB.TextBox txtTroco 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   360
               Left            =   2040
               Locked          =   -1  'True
               TabIndex        =   60
               TabStop         =   0   'False
               Top             =   420
               Width           =   1335
            End
            Begin VB.Label Label45 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Troco:"
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
               Left            =   2040
               TabIndex        =   195
               Top             =   180
               Width           =   570
            End
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Recebido:"
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
               Left            =   180
               TabIndex        =   194
               Top             =   180
               Width           =   885
            End
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00C0FFC0&
            Height          =   555
            Left            =   120
            TabIndex        =   192
            Top             =   300
            Width           =   7275
            Begin VB.TextBox txtFuncAV 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   180
               Width           =   5955
            End
            Begin VB.TextBox txtCodFuncAV 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               TabIndex        =   44
               Top             =   180
               Width           =   1035
            End
         End
         Begin VB.Frame Frame13 
            BackColor       =   &H00C0FFC0&
            Height          =   1815
            Left            =   3960
            TabIndex        =   185
            Top             =   1560
            Width           =   3435
            Begin VB.PictureBox Picture8 
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   1440
               ScaleHeight     =   210
               ScaleWidth      =   1035
               TabIndex        =   187
               Top             =   1020
               Width           =   1035
               Begin VB.OptionButton optAcrescPorcAV 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "%"
                  Height          =   210
                  Left            =   600
                  TabIndex        =   56
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   435
               End
               Begin VB.OptionButton optAcrescRSAV 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "R$"
                  Height          =   210
                  Left            =   60
                  TabIndex        =   55
                  TabStop         =   0   'False
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   555
               End
            End
            Begin VB.TextBox txtAcrescAV 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2520
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   960
               Width           =   855
            End
            Begin VB.TextBox txtSubTotalAV 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1860
               Locked          =   -1  'True
               TabIndex        =   51
               TabStop         =   0   'False
               Text            =   "0,00"
               Top             =   240
               Width           =   1515
            End
            Begin VB.TextBox txtDescAV 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2520
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   600
               Width           =   855
            End
            Begin VB.TextBox txtTotalDescAV 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   58
               TabStop         =   0   'False
               Text            =   "0,00"
               Top             =   1320
               Width           =   1455
            End
            Begin VB.PictureBox Picture7 
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   1440
               ScaleHeight     =   210
               ScaleWidth      =   1035
               TabIndex        =   186
               Top             =   660
               Width           =   1035
               Begin VB.OptionButton optDescPorcAV 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "%"
                  Height          =   210
                  Left            =   600
                  TabIndex        =   53
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   435
               End
               Begin VB.OptionButton optDescRSAV 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "R$"
                  Height          =   210
                  Left            =   60
                  TabIndex        =   52
                  TabStop         =   0   'False
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   555
               End
            End
            Begin VB.Label Label43 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Acresc.:"
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
               Left            =   720
               TabIndex        =   191
               Top             =   1020
               Width           =   780
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Desc.:"
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
               Left            =   900
               TabIndex        =   190
               Top             =   660
               Width           =   570
            End
            Begin VB.Label Label40 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total:"
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
               Left            =   1380
               TabIndex        =   189
               Top             =   1380
               Width           =   510
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SubTotal:"
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
               Left            =   960
               TabIndex        =   188
               Top             =   300
               Width           =   840
            End
         End
         Begin ChamaleonBtn.chameleonButton cmdAVfinalizar 
            Height          =   315
            Left            =   5280
            TabIndex        =   61
            Top             =   4320
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Finalizar"
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
            MICON           =   "OS_Automoveis.frx":2A1A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdAVcancelar 
            Height          =   315
            Left            =   6360
            TabIndex        =   62
            Top             =   4320
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Cancelar"
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
            MICON           =   "OS_Automoveis.frx":2A36
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
      Begin VB.PictureBox Picture2 
         Height          =   8295
         Left            =   -74940
         ScaleHeight     =   8235
         ScaleWidth      =   7815
         TabIndex        =   143
         Top             =   420
         Width           =   7875
         Begin VB.PictureBox Picture19 
            Height          =   975
            Left            =   60
            ScaleHeight     =   915
            ScaleWidth      =   7635
            TabIndex        =   151
            Top             =   120
            Width           =   7695
            Begin VB.Label lblCarro2b 
               AutoSize        =   -1  'True
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   330
               Left            =   4080
               TabIndex        =   155
               Top             =   360
               Width           =   75
            End
            Begin VB.Label lblCarro2a 
               AutoSize        =   -1  'True
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   330
               Left            =   120
               TabIndex        =   154
               Top             =   360
               Width           =   75
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               Caption         =   "COR:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   4080
               TabIndex        =   153
               Top             =   120
               Width           =   390
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               Caption         =   "MOTO:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   120
               TabIndex        =   152
               Top             =   120
               Width           =   540
            End
         End
         Begin VB.PictureBox frmPecas 
            Enabled         =   0   'False
            Height          =   1275
            Left            =   60
            ScaleHeight     =   1215
            ScaleWidth      =   7635
            TabIndex        =   144
            Top             =   1140
            Width           =   7695
            Begin VB.TextBox txtCodPeca 
               Appearance      =   0  'Flat
               Height          =   255
               Left            =   720
               TabIndex        =   145
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   555
            End
            Begin VB.TextBox txtQuantPeca 
               Height          =   315
               Left            =   5940
               TabIndex        =   37
               Top             =   360
               Width           =   555
            End
            Begin VB.TextBox txtPecas 
               Height          =   315
               Left            =   60
               MaxLength       =   90
               TabIndex        =   35
               Top             =   360
               Width           =   4695
            End
            Begin VB.TextBox txtTotalPeca 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6540
               TabIndex        =   38
               Top             =   360
               Width           =   1035
            End
            Begin MSMask.MaskEdBox mskValorPeca 
               Height          =   315
               Left            =   4800
               TabIndex        =   36
               Top             =   360
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin ChamaleonBtn.chameleonButton cmdRemoverPecas 
               Height          =   315
               Left            =   6420
               TabIndex        =   41
               Top             =   780
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "&Remover"
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
               MICON           =   "OS_Automoveis.frx":2A52
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdAdicionarPecas 
               Height          =   315
               Left            =   5160
               TabIndex        =   39
               Top             =   780
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "&Adicionar"
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
               MICON           =   "OS_Automoveis.frx":2A6E
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Left            =   4800
               TabIndex        =   150
               Top             =   120
               Width           =   360
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Quant:"
               Height          =   195
               Left            =   5940
               TabIndex        =   149
               Top             =   120
               Width           =   480
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Peças:"
               Height          =   195
               Left            =   60
               TabIndex        =   148
               Top             =   120
               Width           =   495
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Pressione a tecla  [ F2 ]  para obter os produtos."
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
               Left            =   60
               TabIndex        =   147
               Top             =   720
               Visible         =   0   'False
               Width           =   4440
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total"
               Height          =   195
               Left            =   6540
               TabIndex        =   146
               Top             =   120
               Width           =   360
            End
         End
         Begin ChamaleonBtn.chameleonButton cmdPecas 
            Height          =   615
            Left            =   60
            TabIndex        =   156
            TabStop         =   0   'False
            Top             =   7560
            Width           =   1875
            _ExtentX        =   3307
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
            MICON           =   "OS_Automoveis.frx":2A8A
            PICN            =   "OS_Automoveis.frx":2AA6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSFlexGridLib.MSFlexGrid Grid_Pecas 
            Height          =   4995
            Left            =   60
            TabIndex        =   40
            Top             =   2520
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   8811
            _Version        =   393216
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin VB.Label lblTotalPeca 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
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
            Left            =   6720
            TabIndex        =   157
            Top             =   7620
            Width           =   1035
         End
      End
      Begin VB.PictureBox frmSecundario 
         Enabled         =   0   'False
         Height          =   5295
         Left            =   120
         ScaleHeight     =   5235
         ScaleWidth      =   6135
         TabIndex        =   129
         Top             =   1200
         Width           =   6195
         Begin VB.ComboBox cboTanque 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   4020
            TabIndex        =   13
            Top             =   2100
            Width           =   2055
         End
         Begin VB.TextBox txtKM 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   2100
            TabIndex        =   12
            Top             =   2100
            Width           =   1875
         End
         Begin VB.TextBox txtPlaca2 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   900
            MaxLength       =   4
            TabIndex        =   11
            Top             =   2100
            Width           =   1155
         End
         Begin VB.TextBox txtPlaca1 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   60
            MaxLength       =   3
            TabIndex        =   10
            Top             =   2100
            Width           =   795
         End
         Begin VB.ComboBox cboCor 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   4020
            TabIndex        =   9
            Top             =   1500
            Width           =   2055
         End
         Begin VB.TextBox txtCodCliente 
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   5340
            TabIndex        =   135
            TabStop         =   0   'False
            Top             =   660
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cboModelo 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   2100
            TabIndex        =   8
            Top             =   1500
            Width           =   1875
         End
         Begin VB.TextBox txtDescricao 
            Appearance      =   0  'Flat
            Height          =   915
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   18
            Top             =   4260
            Width           =   6015
         End
         Begin VB.TextBox txtCodFuncionario 
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   3240
            TabIndex        =   134
            TabStop         =   0   'False
            Top             =   60
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cboCliente 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   60
            TabIndex        =   6
            Top             =   900
            Width           =   6015
         End
         Begin VB.ComboBox cboFuncionario 
            Height          =   315
            Left            =   2100
            TabIndex        =   5
            Top             =   300
            Width           =   3975
         End
         Begin VB.Frame Frame1 
            Caption         =   "Acessorios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1515
            Left            =   60
            TabIndex        =   130
            Top             =   2460
            Width           =   6015
            Begin VB.PictureBox Picture15 
               Height          =   1215
               Left            =   60
               ScaleHeight     =   1155
               ScaleWidth      =   2655
               TabIndex        =   131
               Top             =   240
               Width           =   2715
               Begin VB.ComboBox cboAcessorios 
                  Height          =   315
                  Left            =   60
                  TabIndex        =   14
                  Top             =   360
                  Width           =   2595
               End
               Begin VB.TextBox txtCodAcessorio 
                  Appearance      =   0  'Flat
                  Height          =   285
                  Left            =   1920
                  TabIndex        =   132
                  Top             =   60
                  Visible         =   0   'False
                  Width           =   675
               End
               Begin ChamaleonBtn.chameleonButton cmdRemoverAcessorios 
                  Height          =   315
                  Left            =   1380
                  TabIndex        =   17
                  Top             =   780
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   556
                  BTYPE           =   3
                  TX              =   "R&emover"
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
                  MICON           =   "OS_Automoveis.frx":2FA2
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin ChamaleonBtn.chameleonButton cmdAdicionarAcessorios 
                  Height          =   315
                  Left            =   60
                  TabIndex        =   15
                  Top             =   780
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   556
                  BTYPE           =   3
                  TX              =   "A&dicionar"
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
                  MICON           =   "OS_Automoveis.frx":2FBE
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VB.Label Label21 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nome"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   133
                  Top             =   120
                  Width           =   420
               End
            End
            Begin MSFlexGridLib.MSFlexGrid Grid_Acessorio 
               Height          =   1155
               Left            =   2820
               TabIndex        =   16
               Top             =   240
               Width           =   3135
               _ExtentX        =   5530
               _ExtentY        =   2037
               _Version        =   393216
               ScrollBars      =   2
               SelectionMode   =   1
               Appearance      =   0
            End
         End
         Begin VB.ComboBox cboFabricante 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   60
            TabIndex        =   7
            Top             =   1500
            Width           =   1995
         End
         Begin MSMask.MaskEdBox mskDataEntrada 
            Height          =   315
            Left            =   60
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   300
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskHoraEntrada 
            Height          =   315
            Left            =   1380
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   300
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kilometragem"
            Height          =   195
            Left            =   2100
            TabIndex        =   208
            Top             =   1860
            Width           =   945
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Placa"
            Height          =   195
            Left            =   60
            TabIndex        =   207
            Top             =   1860
            Width           =   405
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanque"
            Height          =   195
            Left            =   4020
            TabIndex        =   206
            Top             =   1860
            Width           =   555
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cor"
            Height          =   195
            Left            =   4020
            TabIndex        =   205
            Top             =   1260
            Width           =   240
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente"
            Height          =   195
            Left            =   60
            TabIndex        =   142
            Top             =   660
            Width           =   480
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data"
            Height          =   195
            Left            =   60
            TabIndex        =   141
            Top             =   60
            Width           =   345
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fabricante"
            Height          =   195
            Left            =   60
            TabIndex        =   140
            Top             =   1260
            Width           =   750
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Modelo"
            Height          =   195
            Left            =   2100
            TabIndex        =   139
            Top             =   1260
            Width           =   525
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Recepcionista"
            Height          =   195
            Left            =   2100
            TabIndex        =   138
            Top             =   60
            Width           =   1020
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIÇÃO DO DEFEITO"
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
            TabIndex        =   137
            Top             =   4020
            Width           =   2265
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hora"
            Height          =   195
            Left            =   1380
            TabIndex        =   136
            Top             =   60
            Width           =   345
         End
      End
      Begin VB.PictureBox Picture14 
         Height          =   1455
         Left            =   120
         ScaleHeight     =   1395
         ScaleWidth      =   6135
         TabIndex        =   128
         Top             =   7320
         Width           =   6195
         Begin VB.Frame Frame2 
            Height          =   1335
            Left            =   3120
            TabIndex        =   159
            Top             =   0
            Width           =   2955
            Begin VB.TextBox txtTotalServicos 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   162
               Top             =   180
               Width           =   1575
            End
            Begin VB.TextBox txtTotalGeral 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1320
               TabIndex        =   161
               Top             =   900
               Width           =   1575
            End
            Begin VB.TextBox txtTotalPecas 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   160
               Top             =   540
               Width           =   1575
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Serviços:"
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
               Left            =   420
               TabIndex        =   165
               Top             =   240
               Width           =   810
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total:"
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
               Left            =   660
               TabIndex        =   164
               Top             =   960
               Width           =   510
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Peças:"
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
               Left            =   600
               TabIndex        =   163
               Top             =   600
               Width           =   600
            End
         End
         Begin ChamaleonBtn.chameleonButton cmdFinalizarAV 
            Height          =   315
            Left            =   600
            TabIndex        =   42
            Top             =   300
            Visible         =   0   'False
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Venda à Vista (F10)"
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
            MICON           =   "OS_Automoveis.frx":2FDA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdFinalizarAP 
            Height          =   315
            Left            =   600
            TabIndex        =   43
            Top             =   720
            Visible         =   0   'False
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Venda à Prazo (F12)"
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
            MICON           =   "OS_Automoveis.frx":2FF6
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
      Begin VB.PictureBox Picture1 
         Height          =   8295
         Left            =   -74940
         ScaleHeight     =   8235
         ScaleWidth      =   7815
         TabIndex        =   115
         Top             =   420
         Width           =   7875
         Begin MSFlexGridLib.MSFlexGrid Grid_Servicos 
            Height          =   4995
            Left            =   60
            TabIndex        =   33
            Top             =   2460
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   8811
            _Version        =   393216
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin VB.PictureBox Picture17 
            Height          =   975
            Left            =   60
            ScaleHeight     =   915
            ScaleWidth      =   7635
            TabIndex        =   121
            Top             =   120
            Width           =   7695
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               Caption         =   "MOTO:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   120
               TabIndex        =   125
               Top             =   120
               Width           =   540
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               Caption         =   "COR:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   4080
               TabIndex        =   124
               Top             =   120
               Width           =   390
            End
            Begin VB.Label lblCarro1a 
               AutoSize        =   -1  'True
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   330
               Left            =   120
               TabIndex        =   123
               Top             =   360
               Width           =   75
            End
            Begin VB.Label lblCarro1b 
               AutoSize        =   -1  'True
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   330
               Left            =   4080
               TabIndex        =   122
               Top             =   360
               Width           =   75
            End
         End
         Begin VB.PictureBox frmServico 
            Enabled         =   0   'False
            Height          =   1275
            Left            =   60
            ScaleHeight     =   1215
            ScaleWidth      =   7635
            TabIndex        =   116
            Top             =   1140
            Width           =   7695
            Begin VB.ComboBox cboServicos 
               Height          =   315
               Left            =   60
               Sorted          =   -1  'True
               TabIndex        =   29
               Top             =   360
               Width           =   5835
            End
            Begin VB.TextBox txtCodServico 
               Appearance      =   0  'Flat
               Height          =   225
               Left            =   900
               TabIndex        =   117
               TabStop         =   0   'False
               Top             =   60
               Visible         =   0   'False
               Width           =   555
            End
            Begin VB.TextBox txtQuantServico 
               Height          =   315
               Left            =   5940
               TabIndex        =   30
               Top             =   360
               Width           =   555
            End
            Begin ChamaleonBtn.chameleonButton cmdRemoverServicos 
               Height          =   315
               Left            =   6360
               TabIndex        =   34
               Top             =   780
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "R&emover"
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
               MICON           =   "OS_Automoveis.frx":3012
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSMask.MaskEdBox mskValorServico 
               Height          =   315
               Left            =   6540
               TabIndex        =   31
               Top             =   360
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin ChamaleonBtn.chameleonButton cmdAdicionarServicos 
               Height          =   315
               Left            =   5100
               TabIndex        =   32
               Top             =   780
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "A&dicionar"
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
               MICON           =   "OS_Automoveis.frx":302E
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Left            =   6540
               TabIndex        =   120
               Top             =   120
               Width           =   360
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Quant:"
               Height          =   195
               Left            =   5940
               TabIndex        =   119
               Top             =   120
               Width           =   480
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Serviços:"
               Height          =   195
               Left            =   60
               TabIndex        =   118
               Top             =   120
               Width           =   660
            End
         End
         Begin ChamaleonBtn.chameleonButton cmdServicos 
            Height          =   615
            Left            =   60
            TabIndex        =   126
            TabStop         =   0   'False
            Top             =   7560
            Width           =   1875
            _ExtentX        =   3307
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
            MICON           =   "OS_Automoveis.frx":304A
            PICN            =   "OS_Automoveis.frx":3066
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
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
            Left            =   6720
            TabIndex        =   127
            Top             =   7560
            Width           =   1035
         End
      End
      Begin VB.PictureBox frmPrincipal 
         Enabled         =   0   'False
         Height          =   735
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   6135
         TabIndex        =   111
         Top             =   420
         Width           =   6195
         Begin VB.ComboBox cboStatus 
            BackColor       =   &H00C0C0FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            TabIndex        =   1
            Top             =   300
            Width           =   2355
         End
         Begin VB.ComboBox cboMecanico 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2460
            TabIndex        =   2
            Top             =   300
            Width           =   3615
         End
         Begin VB.TextBox txtCodMecanico 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5280
            TabIndex        =   112
            TabStop         =   0   'False
            Top             =   60
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            Height          =   195
            Left            =   60
            TabIndex        =   114
            Top             =   60
            Width           =   450
         End
         Begin VB.Label lblMecanico 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Técnico Responsável"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2460
            TabIndex        =   113
            Top             =   60
            Width           =   1560
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdImprimirEntrada 
         Height          =   615
         Left            =   6360
         TabIndex        =   26
         Top             =   3720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Imprimir"
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
         MICON           =   "OS_Automoveis.frx":354E
         PICN            =   "OS_Automoveis.frx":356A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdGerarEntrada 
         Height          =   615
         Left            =   6360
         TabIndex        =   22
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Salvar"
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
         MICON           =   "OS_Automoveis.frx":3884
         PICN            =   "OS_Automoveis.frx":38A0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAlterar 
         Height          =   615
         Left            =   6360
         TabIndex        =   24
         Top             =   2400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Alterar"
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
         MICON           =   "OS_Automoveis.frx":3BBA
         PICN            =   "OS_Automoveis.frx":3BD6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdApagar 
         Height          =   615
         Left            =   6360
         TabIndex        =   25
         Top             =   3060
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Excluir"
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
         MICON           =   "OS_Automoveis.frx":44B0
         PICN            =   "OS_Automoveis.frx":44CC
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
         Left            =   6360
         TabIndex        =   0
         Top             =   420
         Width           =   1575
         _ExtentX        =   2778
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
         MICON           =   "OS_Automoveis.frx":47E6
         PICN            =   "OS_Automoveis.frx":4802
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelarEntrada 
         Height          =   615
         Left            =   6360
         TabIndex        =   23
         Top             =   1740
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Cancelar"
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
         MICON           =   "OS_Automoveis.frx":54DC
         PICN            =   "OS_Automoveis.frx":54F8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCadastrarAcessorios 
         Height          =   615
         Left            =   6360
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   5040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Acessorios"
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
         MICON           =   "OS_Automoveis.frx":5812
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCadastrarCliente 
         Height          =   615
         Left            =   6360
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   6360
         Width           =   1575
         _ExtentX        =   2778
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
         MICON           =   "OS_Automoveis.frx":582E
         PICN            =   "OS_Automoveis.frx":584A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdGarantia 
         Height          =   615
         Left            =   6360
         TabIndex        =   199
         Top             =   4380
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Garantia"
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
         MICON           =   "OS_Automoveis.frx":824C
         PICN            =   "OS_Automoveis.frx":8268
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton chameleonButton1 
         Height          =   615
         Left            =   6360
         TabIndex        =   200
         TabStop         =   0   'False
         Top             =   5700
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Equipamentos"
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
         MICON           =   "OS_Automoveis.frx":8582
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblQuantFiltro 
         AutoSize        =   -1  'True
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
         Left            =   -74880
         TabIndex        =   158
         Top             =   8040
         Width           =   75
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   198
      Top             =   10005
      Width           =   14700
      _ExtentX        =   25929
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19897
            Text            =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
            TextSave        =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1676
            MinWidth        =   1676
            TextSave        =   "19:47"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
Attribute VB_Name = "OS_Automoveis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moCombo As cComboHelper

Dim xParc As Long, xAcess As Long
Dim xPeca As Long, xServ As Long

Dim w As Long
Dim var_Num_Pedido As Long
Dim numCol As Integer
Dim numRow As Integer

Dim Texto As String         'usado pra preencher os combos
Dim i, Posicao As Integer   'usado pra preencher os combos
Dim Posicionar As Boolean   'usado pra preencher os combos

Dim OS_FECHADA As Boolean
Dim OS_FECHADAPAGA As Boolean
Dim VERIFICAR_QUANTIDADE As Boolean
Dim CAIXA_FECHADO As Boolean

Private Sub Calcular_Prazo()
   If cboPrazo.Text = "" Then Exit Sub
   
   If txtEntrada.Text = "0,00" Or txtEntrada.Text = "" Then
      mskInicio.Text = Format(DateAdd("m", Val(1), Date), "dd/mm/yy")
      mskTermino.Text = Format(DateAdd("m", Val(cboQuantParc.Text) - 1, mskInicio.Text), "dd/mm/yy")
   Else
      mskInicio.Text = Format(Date, "dd/mm/yy")
      mskTermino.Text = Format(DateAdd("m", Val(cboQuantParc.Text), mskInicio.Text), "dd/mm/yy")
   End If
   
   'mskInicio.Text = Format(DateAdd("d", cboPrazo, Date), "dd/mm/yy")
End Sub

Private Sub Mostrar_ValorRestante()
   Dim Valor As Currency
   Dim QUANT As Integer
   Dim Entrada As Currency
   Dim RESULTADO As Currency
   Dim VALOR_SENTRADA As Currency
   
   If txtEntrada.Text = "" Then Entrada = 0 Else Entrada = txtEntrada.Text
   If txtTotalDesc.Text = "" Then Valor = 0 Else Valor = txtTotalDesc.Text
   ' QUANT = txtQuantParc.Text
   
   VALOR_SENTRADA = Valor - Entrada
   txtValorRest.Text = Format(VALOR_SENTRADA, "##,##0.00")
End Sub

Private Sub LimparObjetos_Avista()
   txtSubTotalAV.Text = "0,00"
   optDescPorcAV.Value = True
   optAVdinheiro.Value = True
   optDebito.Value = True
   frmCartao.Visible = False
End Sub

Private Sub Autonumeracao_Parcelas()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultima_parcela FROM parcelas;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then xParc = r("ultima_parcela") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub LimparObjetos_Prazo()
   txtEntrada.Text = Format(0, "##,##0.00")
   cboPrazo.Text = "30"
   txtValorParc.Text = Format(0, "##,##0.00")
   mskInicio.Mask = ""
   mskInicio.Text = ""
   optDescRS.Value = True
   optAvulso.Value = True
   txtDesc.Text = "0,00"
   cboQuantParc.Text = "1"
End Sub

Private Function Atualizar_Dados_OS() As Boolean
   'A atualização deve ser feita utilizando o comando UPDATE do sql
   'e não mais usando o método .Update do Recordset
   
   'Não se deve comparar se o campo está vazio ou não, pois dessa forma não
   'haverá atualização quando for necessário apagar alguma informação
   
   Dim sSQL As String
   
   'Comando de atualização
   sSQL = "UPDATE OS_Carro_Cadastro SET " & _
      "data_entrada = CONVERT(DATETIME, '" & Format$(mskDataEntrada.Text, ocDATA) & "', 103), " & _
      "hora_entrada = '" & Format$(mskHoraEntrada.Text, ocHORA) & "', " & _
      "fabricante = '" & cboFabricante.Text & "', " & _
      "modelo = '" & cboModelo.Text & "', " & _
      "cor = '" & cboCor.Text & "', " & _
      "placa1 = '" & txtPlaca1.Text & "', " & _
      "placa2 = '" & txtPlaca2.Text & "', " & _
      "km = '" & txtKM.Text & "', " & _
      "tanque = '" & cboTanque.Text & "', " & _
      "cod_cliente = " & txtCodCliente.Text & ", " & _
      "cod_funcionario = " & txtCodFuncionario.Text & ", " & _
      "descricao = '" & txtDescricao.Text & "', " & _
      "status = '" & cboStatus.Text & "', " & _
      "status_os = 0, " & _
      "cod_mecanico = " & IIf(txtCodMecanico.Text = "", "Null", txtCodMecanico.Text) & ", " & _
      "data_saida = " & IIf(mskDataSaida.Text = "", "Null", "CONVERT(DATETIME, '" & Format$(mskDataSaida.Text, ocDATA) & "', 103)") & ", " & _
      "hora_saida = " & IIf(mskHoraSaida.Text = "", "Null", "'" & Format$(mskHoraSaida.Text, ocHORA) & "'") & ", " & _
      "tipo_os = '" & cboTipoOS.Text & "' "

   'Condição para atualização
   sSQL = sSQL & "WHERE (cod_os = " & txtCodOS.Text & ");"
   
   'Retorna o resultado da atualização
   Atualizar_Dados_OS = dbData.Execute(sSQL)
End Function

Private Sub AutoNumeracao_Acessorio()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultimo FROM os_acessorios;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then xAcess = Format(r("ultimo") + 1, "000000")
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub AutoNumeracao_OS()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT ISNULL(MAX(cod_pedido), 0) AS ultima_os FROM pedidos;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then txtCodOS.Text = Format(r("ultima_os") + 1, "000000")
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub AutoNumeracao_Peca()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultima_peca FROM pedidos_itens;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then xPeca = Format(r("ultima_peca") + 1, "000000")
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub AutoNumeracao_Servico()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultimo FROM os_servicos;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then xServ = Format(r("ultimo") + 1, "000000")
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub FormatarGrid_Acessorios(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid_Acessorio
      .Visible = False
      
      .Clear
      .Cols = 5
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 0
      .ColWidth(3) = 0
      .ColWidth(4) = 2900
      
      .RowHeight(0) = 0
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "OS"
      .TextMatrix(0, 3) = "COD_ACESSORIO"
      .TextMatrix(0, 4) = "ACESSÓRIO"
      
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = rTabela("codigo")
            .TextMatrix(.rows - 1, 2) = rTabela("cod_os")
            .TextMatrix(.rows - 1, 3) = rTabela("cod_acessorio")
            .TextMatrix(.rows - 1, 4) = rTabela("acessorio")
            
            rTabela.MoveNext
            .rows = .rows + 1
            i = i + 1
         Loop
      End If
      
      .rows = .rows - 1
      .Redraw = True
      .Visible = True
   End With
End Sub

Private Sub LimparGrid_Pecas()
   Dim i As Integer
   
   With Grid_Pecas
      .Clear
      .Cols = 6
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 4400
      .ColWidth(3) = 1000
      .ColWidth(4) = 1000
      .ColWidth(5) = 1000
      
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "PEÇAS"
      .TextMatrix(0, 3) = "VALOR"
      .TextMatrix(0, 4) = "QUANT."
      .TextMatrix(0, 5) = "TOTAL"
      
      .Redraw = False
      .rows = .rows + 1
      
      i = i + 1
      .rows = .rows - 1
      .Redraw = True
      
      lblTotalPeca.Caption = Format(0, "##,##0.00")
   End With
End Sub

Private Sub FormatarGrid_Pecas(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid_Pecas
      .Clear
      .Cols = 7
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 0
      .ColWidth(3) = 4400
      .ColWidth(4) = 1000
      .ColWidth(5) = 1000
      .ColWidth(6) = 1000
      
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "COD_PROD"
      .TextMatrix(0, 3) = "PEÇAS"
      .TextMatrix(0, 4) = "VALOR"
      .TextMatrix(0, 5) = "QUANT."
      .TextMatrix(0, 6) = "TOTAL"
      
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = rTabela("codigo")
            .TextMatrix(.rows - 1, 2) = rTabela("cod_produto")
            .TextMatrix(.rows - 1, 3) = rTabela("descricao")
            .TextMatrix(.rows - 1, 4) = Format(rTabela("preco"), ocMONEY)
            .TextMatrix(.rows - 1, 5) = rTabela("quantidade")
            .TextMatrix(.rows - 1, 6) = Format(rTabela("total"), ocMONEY)
            
            rTabela.MoveNext
            .rows = .rows + 1
            i = i + 1
         Loop
      End If
      
      .rows = .rows - 1
      .Redraw = True
      
      lblTotalPeca.Caption = Format(SomaGrid(Grid_Pecas, 6), ocMONEY)
   End With
End Sub

Private Sub FormatarGrid_Servicos(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid_Servicos
      .Clear
      .Cols = 6
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 4400
      .ColWidth(3) = 1000
      .ColWidth(4) = 1000
      .ColWidth(5) = 1000
      
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "SERVIÇOS"
      .TextMatrix(0, 3) = "VALOR"
      .TextMatrix(0, 4) = "QUANT."
      .TextMatrix(0, 5) = "TOTAL"
      
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = rTabela("codigo")
            .TextMatrix(.rows - 1, 2) = rTabela("descricao")
            .TextMatrix(.rows - 1, 3) = Format(rTabela("preco"), ocMONEY)
            .TextMatrix(.rows - 1, 4) = Format(rTabela("quantidade"), "00")
            .TextMatrix(.rows - 1, 5) = Format(rTabela("total"), ocMONEY)
            
            rTabela.MoveNext
            .rows = .rows + 1
            i = i + 1
         Loop
      End If
      
      .rows = .rows - 1
      .Redraw = True
      
      lblTotal.Caption = Format(SomaGrid(Grid_Servicos, 5), ocMONEY)
   End With
End Sub

Private Sub LimparGrid_Servicos()
   Dim i As Integer
   
   With Grid_Servicos
      .Clear
      .Cols = 6
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 4400
      .ColWidth(3) = 1000
      .ColWidth(4) = 1000
      .ColWidth(5) = 1000
      
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "SERVIÇOS"
      .TextMatrix(0, 3) = "VALOR"
      .TextMatrix(0, 4) = "QUANT."
      .TextMatrix(0, 5) = "TOTAL"
      
      .Redraw = False
      .rows = .rows + 1
      i = i + 1
      
      .rows = .rows - 1
      .Redraw = True
      
      lblTotal.Caption = Format(0, ocMONEY)
   End With
End Sub

Private Sub LimparGrid_Acessorios()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT * FROM os_acessorios WHERE 0 = 1;"
   Set r = dbData.OpenRecordset(sSQL)
   
   FormatarGrid_Acessorios r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub LimparObjetos_Pecas()
   txtCodPeca.Text = ""
   txtPecas.Text = ""
   txtQuantPeca.Text = ""
   mskValorPeca.Mask = ""
   mskValorPeca.Text = ""
   txtTotalPeca.Text = ""
End Sub

Private Sub LimparObjetos_Servicos()
   txtCodServico.Text = ""
   cboServicos.Text = ""
   txtQuantServico.Text = ""
   mskValorServico.Mask = ""
   mskValorServico.Text = ""
End Sub

Private Sub MostrarGrid_Acessorios()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtCodOS.Text = "" Then txtCodOS.Text = 0
   
   sSQL = "SELECT * FROM os_acessorios WHERE (cod_os = " & txtCodOS.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   FormatarGrid_Acessorios r
   
   If r.State <> 0 Then r.Close
End Sub

Private Sub MostrarGrid_Servicos()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtCodOS.Text = "" Then txtCodOS.Text = 0
   
   sSQL = "SELECT * FROM os_servicos WHERE (cod_os = " & txtCodOS.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   FormatarGrid_Servicos r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub MostrarGrid_Pecas()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtCodOS.Text = "" Then txtCodOS.Text = 0

   sSQL = "SELECT descricao, codigo, cod_produto, preco, quantidade, (preco * quantidade) as total FROM pedidos_itens WHERE (cod_pedido = " & txtCodOS.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   FormatarGrid_Pecas r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub LimparObjetos_Entrada()
   txtCodCliente.Text = ""
   txtCodFuncionario.Text = ""
   mskDataEntrada.Mask = ""
   mskDataEntrada.Text = ""
   mskHoraEntrada.Mask = ""
   mskHoraEntrada.Text = ""
   cboCliente.Text = ""
   cboFuncionario.Text = ""
   txtDescricao.Text = ""
   cboStatus.Text = ""
   txtCodMecanico.Text = ""
   cboMecanico.Text = ""
   txtCodAcessorio.Text = ""
   cboAcessorios.Text = ""
   cboTipoOS.Text = ""
   mskDataSaida.Mask = ""
   mskDataSaida.Text = ""
   mskHoraSaida.Mask = ""
   mskHoraSaida.Text = ""
   cboFabricante.Text = ""
   cboModelo.Text = ""
   cboCor.Text = ""
   txtPlaca1.Text = ""
   txtPlaca2.Text = ""
   txtKM.Text = ""
   cboTanque.Text = ""
   
   LimparGrid_Pecas
   LimparGrid_Servicos
End Sub

Private Sub Mostrar_Entrada(rTabela As ADODB.Recordset)
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   'Se o parametro passado é Nothing, sai da rotina
   If rTabela Is Nothing Then Exit Sub
   
   If Not rTabela.BOF Then
      mskDataEntrada.Text = Format(rTabela("data_entrada"), "dd/mm/yy")
      mskHoraEntrada.Text = Format(rTabela("hora_entrada"), ocHORA)
      cboFabricante.Text = ValidateNull(rTabela("fabricante"))
      cboModelo.Text = ValidateNull(rTabela("modelo"))
      cboCor.Text = ValidateNull(rTabela("cor"))
      txtPlaca1.Text = ValidateNull(rTabela("placa1"))
      txtPlaca2.Text = ValidateNull(rTabela("placa2"))
      txtKM.Text = ValidateNull(rTabela("km"))
      cboTanque.Text = ValidateNull(rTabela("tanque"))
      txtCodCliente.Text = ValidateNull(rTabela("cod_cliente"))
      txtCodFuncionario.Text = ValidateNull(rTabela("cod_funcionario"))
      txtDescricao.Text = ValidateNull(rTabela("descricao"))
      cboStatus.Text = ValidateNull(rTabela("status"))
      txtCodMecanico.Text = ValidateNull(rTabela("cod_mecanico"))
      mskDataSaida.Text = Format(rTabela("data_saida"), "dd/mm/yy")
      mskHoraSaida.Text = Format(rTabela("hora_saida"), ocHRMN)
      cboTipoOS.Text = ValidateNull(rTabela("tipo_os"))
   End If
   
   sSQL = "SELECT cod_cliente, cod_pedido FROM pedidos WHERE (cod_pedido = " & txtCodOS.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then txtCodCliente.Text = rTabela("cod_cliente")
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub MostrarGrid_OS()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   Dim SITUACAO As String
   Dim var_STATUS As String
   
   'Situação
   If optStatusTodos.Value = True Then
      SITUACAO = ""
   ElseIf optStatusComecar.Value = True Then
      SITUACAO = "AND (OS_Carro_Cadastro.status = 'À COMEÇAR') "
   ElseIf optStatusExecucao.Value = True Then
      SITUACAO = "AND (OS_Carro_Cadastro.status = 'EM EXECUÇÃO') "
   ElseIf optStatusAguardando.Value = True Then
      SITUACAO = "AND (OS_Carro_Cadastro.status = 'AGUARDANDO') "
   ElseIf optStatusTerminado.Value = True Then
      SITUACAO = "AND (OS_Carro_Cadastro.status = 'TERMINADO') "
   End If
   
   'Situação
   If OptTodos.Value = True Then
      var_STATUS = ""
   ElseIf OptAbertos.Value = True Then
      var_STATUS = "AND (status_os = 0) "
   ElseIf optFechados.Value = True Then
      var_STATUS = "AND (status_os = 1) "
   End If
   
   If optLocalizarCliente.Value = 1 Then
      If txtCodClienteLocalizar.Text = "" Then Exit Sub
      sSQL = "SELECT cliente.*, OS_Carro_Cadastro.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, OS_Carro_Cadastro.* " & _
         "FROM cliente INNER JOIN OS_Carro_Cadastro ON cliente.codigo = OS_Carro_Cadastro.cod_cliente WHERE (cod_cliente = " & txtCodClienteLocalizar.Text & ") " & SITUACAO & var_STATUS & _
         "ORDER BY data_entrada, hora_entrada, OS_Carro_Cadastro.status;"
      
   ElseIf optLocalizarCodOS.Value = 1 Then
      If txtLocalizarCodOS.Text = "" Then Exit Sub
      sSQL = "SELECT cliente.*, OS_Carro_Cadastro.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, OS_Carro_Cadastro.* " & _
         "FROM cliente INNER JOIN OS_Carro_Cadastro ON cliente.codigo = OS_Carro_Cadastro.cod_cliente WHERE (OS_Carro_Cadastro.cod_os = " & txtLocalizarCodOS.Text & ") " & SITUACAO & var_STATUS & _
         "ORDER BY data_entrada, hora_entrada, OS_Carro_Cadastro.status;"
      
   Else
      sSQL = "SELECT cliente.*, OS_Carro_Cadastro.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, OS_Carro_Cadastro.* " & _
         "FROM cliente INNER JOIN OS_Carro_Cadastro ON cliente.codigo = OS_Carro_Cadastro.cod_cliente WHERE (OS_Carro_Cadastro.cod_os <> 0 ) " & SITUACAO & var_STATUS & _
         "ORDER BY data_entrada, hora_entrada, OS_Carro_Cadastro.status;"
      
   End If
   
   Set r = dbData.OpenRecordset(sSQL)
   
   FormatarGrid_OS r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub Somar_Totais()
   Dim Servicos As Currency
   Dim Pecas As Currency
   Dim Total As Currency
   
   If lblTotal.Caption <> "" Then Servicos = lblTotal.Caption Else Servicos = 0
   If lblTotalPeca.Caption <> "" Then Pecas = lblTotalPeca.Caption Else Servicos = 0
   Total = Servicos + Pecas
   
   txtSubtotal.Text = FormatCurrency(Total)
   txtTotalGeral.Text = Format(Total, ocMONEY)
End Sub

Private Sub cboAcessorios_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboAcessorios.Clear
   
   sSQL = "SELECT DISTINCT acessorio, codigo FROM acessorios ORDER BY acessorio;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboAcessorios.AddItem r("acessorio")
      cboAcessorios.ItemData(cboAcessorios.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboAcessorios
End Sub

Private Sub cboAcessorios_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboAcessorios_LostFocus()
   On Error GoTo TrataErro
   
   If cboAcessorios.Text = "" Then txtCodAcessorio.Text = "": Exit Sub
   If cboAcessorios.ListIndex = -1 Then txtCodAcessorio.Text = "": Exit Sub
   txtCodAcessorio = cboAcessorios.ItemData(cboAcessorios.ListIndex)
   Exit Sub

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub CboCliente_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   Dim varNomeAntes As String
   Dim varCodAntes As String
   
   varNomeAntes = cboCliente.Text
   varCodAntes = txtCodCliente.Text
   
   cboCliente.Clear
   
   sSQL = "SELECT DISTINCT nome, codigo FROM cliente ORDER BY nome;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboCliente.AddItem r("nome")
      cboCliente.ItemData(cboCliente.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   cboCliente.Text = varNomeAntes
   txtCodCliente.Text = varCodAntes
   
   moCombo.AttachTo cboCliente
End Sub

Private Sub CboCliente_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub CboCliente_LostFocus()
   On Error GoTo TrataErro
   
   If cboCliente.Text = "" Then txtCodCliente.Text = "": Exit Sub
   
   If cmdAlterar.Enabled = False Then
      If cboCliente.ListIndex = -1 Then
         txtCodCliente.Text = ""
         Exit Sub
      End If
   End If
   
   txtCodCliente = cboCliente.ItemData(cboCliente.ListIndex)
   Exit Sub

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cboCor_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   Dim varNomeAntes As String
   
   varNomeAntes = cboCor.Text
   
   cboCor.Clear
   
   sSQL = "SELECT DISTINCT cor FROM OS_Carro_Cadastro ORDER BY cor;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboCor.AddItem ValidateNull(r("cor"))
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing

   cboCor.Text = varNomeAntes
   
   moCombo.AttachTo cboCor
End Sub


Private Sub cboCor_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cboFabricante_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   Dim varNomeAntes As String
   
   varNomeAntes = cboFabricante.Text
   
   cboFabricante.Clear
   
   sSQL = "SELECT DISTINCT fabricante FROM carros ORDER BY fabricante;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboFabricante.AddItem ValidateNull(r("fabricante"))
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing

   cboFabricante.Text = varNomeAntes
   
   moCombo.AttachTo cboFabricante
End Sub

Private Sub cboFabricante_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboFuncionario_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   Dim varNomeAntes As String
   Dim varCodAntes As String
   
   varNomeAntes = cboFuncionario.Text
   varCodAntes = txtCodFuncionario.Text
   
   cboFuncionario.Clear
   
   sSQL = "SELECT DISTINCT nome, codigo FROM funcionario WHERE (cargo <> 'mecânico') ORDER BY nome;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboFuncionario.AddItem r("nome")
      cboFuncionario.ItemData(cboFuncionario.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   cboFuncionario.Text = varNomeAntes
   txtCodFuncionario.Text = varCodAntes
   
   moCombo.AttachTo cboFuncionario
End Sub

Private Sub cboFuncionario_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboFuncionario_LostFocus()
   On Error GoTo TrataErro
   
   If cboFuncionario.Text = "" Then txtCodFuncionario.Text = "": Exit Sub
   
   If cmdAlterar.Enabled = False Then
      If cboFuncionario.ListIndex = -1 Then
         txtCodFuncionario.Text = ""
         Exit Sub
      End If
   End If
   
   txtCodFuncionario = cboFuncionario.ItemData(cboFuncionario.ListIndex)
   Exit Sub
   
TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cboLocalizar_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboLocalizar.Clear
   
   sSQL = "SELECT codigo, nome FROM cliente ORDER BY nome;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboLocalizar.AddItem r("nome")
      cboLocalizar.ItemData(cboLocalizar.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   SelectControl cboLocalizar
   moCombo.AttachTo cboLocalizar
End Sub

Private Sub cboLocalizar_LostFocus()
   On Error GoTo TrataErro
   
   If cboLocalizar.Text = "" Then Exit Sub
   If cboLocalizar.ListIndex = -1 Then txtCodClienteLocalizar.Text = "": Exit Sub
   txtCodClienteLocalizar = cboLocalizar.ItemData(cboLocalizar.ListIndex)
   Exit Sub
   
TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cboMecanico_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   Dim varNomeAntes As String
   Dim varCodAntes As String
   
   varNomeAntes = cboMecanico.Text
   varCodAntes = txtCodMecanico.Text
   
   cboMecanico.Clear
   
   sSQL = "SELECT DISTINCT nome, codigo FROM funcionario WHERE (cargo IN ('tecnico', 'aux. tecnico')) ORDER BY nome;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboMecanico.AddItem r("nome")
      cboMecanico.ItemData(cboMecanico.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing

   cboMecanico.Text = varNomeAntes
   txtCodMecanico.Text = varCodAntes
   
   moCombo.AttachTo cboMecanico
End Sub

Private Sub cboMecanico_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboMecanico_LostFocus()
   On Error GoTo TrataErro
   
   If cboMecanico.Text = "" Then txtCodMecanico.Text = "": Exit Sub
   
   If cmdAlterar.Enabled = False Then
      If cboMecanico.ListIndex = -1 Then
         txtCodMecanico.Text = ""
         Exit Sub
      End If
   End If
   
   txtCodMecanico = cboMecanico.ItemData(cboMecanico.ListIndex)
   Exit Sub
   
TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub


Private Sub cboModelo_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   Dim varNomeAntes As String
   
   varNomeAntes = cboModelo.Text
   
   cboModelo.Clear
   
   sSQL = "SELECT DISTINCT modelo FROM carros ORDER BY modelo;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboModelo.AddItem ValidateNull(r("modelo"))
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing

   cboModelo.Text = varNomeAntes
   
   moCombo.AttachTo cboModelo
End Sub


Private Sub cboPrazo_Change()
   Calcular_Prazo
End Sub

Private Sub cboPrazo_Click()
   Calcular_Prazo
End Sub

Private Sub cboPrazo_GotFocus()
   cboPrazo.Clear
   cboPrazo.AddItem "5"
   cboPrazo.AddItem "10"
   cboPrazo.AddItem "15"
   cboPrazo.AddItem "20"
   cboPrazo.AddItem "30"
   moCombo.AttachTo cboPrazo
End Sub

Private Sub cboQuantParc_Change()
   Calcular_Parcelas
   Calcular_Prazo
End Sub

Private Sub cboQuantParc_Click()
   Calcular_Parcelas
   Calcular_Prazo
End Sub

Private Sub cboQuantParc_GotFocus()
   Dim i As Integer
   
   cboQuantParc.Clear
   For i = 1 To 12
      cboQuantParc.AddItem i
   Next
   
   moCombo.AttachTo cboQuantParc
End Sub

Private Sub cboQuantParc_LostFocus()
   Calcular_Parcelas
   Calcular_Prazo
End Sub

Private Sub cboServicos_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboServicos.Clear
   
   sSQL = "SELECT codigo, servico, valor FROM servicos ORDER BY servico;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboServicos.AddItem r("servico")
      cboServicos.ItemData(cboServicos.NewIndex) = r("codigo")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboServicos
End Sub

Private Sub cboServicos_LostFocus()
   On Error GoTo TrataErro
   
   If cboServicos.Text = "" Then txtCodServico.Text = "": Exit Sub
   txtCodServico = cboServicos.ItemData(cboServicos.ListIndex)
   Exit Sub
   
TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cboStatus_Change()
   If cboStatus.Text = "À COMEÇAR" Then
      'cmdImprimirEntrada.Enabled = False
      lblMecanico.Enabled = False
      cboMecanico.Enabled = False
      cmdFinalizarAP.Visible = False
      cmdFinalizarAV.Visible = False
      frmServico.Enabled = False
      frmPecas.Enabled = False
   ElseIf cboStatus.Text = "EM EXECUÇÃO" Or cboStatus.Text = "AGUARDANDO" Then
      'cmdImprimirEntrada.Enabled = True
      lblMecanico.Enabled = True
      cboMecanico.Enabled = True
      cmdFinalizarAP.Visible = False
      cmdFinalizarAV.Visible = False
      frmServico.Enabled = True
      frmPecas.Enabled = True
   ElseIf cboStatus.Text = "TERMINADO" Then
      'dbData.Execute "UPDATE OS_Carro_Cadastro SET data_saida = '" & Format(Date, ocDATA_EUA) & "', hora_saida = '" & Format(Now, ocHORA) & "' WHERE (cod_os = " & txtCodOS.Text & ");"
      lblMecanico.Enabled = True
      cboMecanico.Enabled = True
      frmServico.Enabled = False
      frmPecas.Enabled = False
   End If
End Sub

Private Sub cboStatus_Click()
   cboStatus_Change
End Sub

Private Sub cboStatus_GotFocus()
   Dim itemAtual As String
   itemAtual = cboStatus.Text
   cboStatus.Clear
   cboStatus.AddItem "À COMEÇAR"
   cboStatus.AddItem "EM EXECUÇÃO"
   cboStatus.AddItem "AGUARDANDO"
   cboStatus.AddItem "TERMINADO"
   cboStatus.Text = itemAtual
   moCombo.AttachTo cboStatus
End Sub

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboStatus_LostFocus()
   cboStatus_Change
   If cboStatus.Text = "TERMINADO" Then cboMecanico.SetFocus
End Sub

Private Sub cboTanque_GotFocus()
   Dim varNomeAntes As String
   
   varNomeAntes = cboTanque.Text
   
   cboTanque.Clear
   cboTanque.AddItem "VAZIO"
   cboTanque.AddItem "1/4"
   cboTanque.AddItem "MEIO"
   cboTanque.AddItem "3/4"
   cboTanque.AddItem "CHEIO"

   cboTanque.Text = varNomeAntes
   
   moCombo.AttachTo cboTanque
End Sub

Private Sub cboTanque_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cboTipoOS_GotFocus()
   Dim varNomeAntes As String
   
   varNomeAntes = cboTipoOS.Text
   
   cboTipoOS.Clear
   cboTipoOS.AddItem "NOVA"
   cboTipoOS.AddItem "RETORNO"

   cboTipoOS.Text = varNomeAntes
   
   moCombo.AttachTo cboTipoOS
End Sub

Private Sub cboTipoOS_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cmdAdicionarAcessorios_Click()
   If txtCodAcessorio.Text = "" Or txtCodOS.Text = "" Then Exit Sub
   
   'CHECAR SE A OS ESTÁ FECHADA
   Verificar_OS_Fechada
   If OS_FECHADA = True Then Exit Sub
   
   'ADICIONAR NA TABELA OS SERVIÇOS
   AutoNumeracao_Acessorio
   dbData.Execute "INSERT INTO os_acessorios (codigo, cod_os, cod_acessorio, acessorio) VALUES(" & xAcess & ", " & txtCodOS.Text & ", " & txtCodAcessorio.Text & ", '" & cboAcessorios.Text & "')"
   
   MostrarGrid_Acessorios
   
   txtCodAcessorio.Text = ""
   cboAcessorios.Text = ""
   cboAcessorios.SetFocus
End Sub

Private Sub Verificar_OS_FechadaePaga()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT cod_os, status_os FROM OS_Carro_Cadastro WHERE (cod_os = " & txtCodOS.Text & ") AND (status_os = 0);"
   Set r = dbData.OpenRecordset(sSQL)
   
   'OS_FECHADAPAGA = False
   'If r.RecordCount <> 0 Then OS_FECHADAPAGA = True
   
   OS_FECHADAPAGA = (r.RecordCount <> 0)
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub Verificar_OS_Fechada()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT cod_os, status_os FROM OS_Carro_Cadastro WHERE (cod_os = " & txtCodOS.Text & ") AND (status_os = 1);"
   Set r = dbData.OpenRecordset(sSQL)
   
   OS_FECHADA = False
   
   If r.RecordCount <> 0 Then
      ShowMsg "ESTA O.S. JÁ ESTÁ FECHADA!", vbExclamation
      OS_FECHADA = True
   End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub Verificar_Caixa()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT * FROM caixa_dia WHERE (data_abertura = CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103)) AND (status = 1);"
   Set r = dbData.OpenRecordset(sSQL)
   
   CAIXA_FECHADO = False
    
   If r.RecordCount <> 0 Then
      ShowMsg "ESTE CAIXA JÁ ESTÁ FECHADO!", vbExclamation
      CAIXA_FECHADO = True
   End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub cmdAdicionarPecas_Click()
   Dim QUANT As Integer
   Dim Valor As Currency
   Dim Total As Currency
   
   If txtCodPeca.Text = "" Or txtCodOS.Text = "" Then Exit Sub
   
   If txtQuantPeca.Text = "" Then txtQuantPeca.Text = 1
   
   'CHECAR SE A OS ESTÁ FECHADA
   Verificar_OS_Fechada
   If OS_FECHADA = True Then Exit Sub
   
   'VERIFICAR O STATUS
   If cboStatus.Text = "À COMEÇAR" Then
      ShowMsg "Não é possivel adicionar peças em uma OS com Status = À COMEÇAR!", vbExclamation
      Exit Sub
   End If
   
   'Verifica_Quantidade do Estoque
   Verifica_QuantEstoque
   If VERIFICAR_QUANTIDADE = True Then
      LimparObjetos_Pecas
      Exit Sub
   End If
   
   'calcular o total das peças no grid
   If txtQuantPeca.Text = "" Then txtQuantPeca.Text = 1
   
   If txtQuantPeca.Text <> "" Or mskValorPeca.Text <> "" Then
      QUANT = txtQuantPeca.Text
      Valor = mskValorPeca.Text
      Total = Valor * QUANT
   End If
   
   'adicionar na tabela PEDIDOS_ITENS
   AutoNumeracao_Peca
   dbData.Execute "INSERT INTO pedidos_itens (" & _
      "codigo, cod_pedido, cod_produto, descricao, preco, quantidade, data, tipo_venda, maquina) VALUES (" & _
      xPeca & ", " & txtCodOS.Text & ", " & txtCodPeca.Text & ", '" & txtPecas.Text & "', " & Replace(CCur(mskValorPeca.Text), ",", ".") & ", " & txtQuantPeca.Text & ", '" & Format(Date, ocDATA_EUA) & "', 'OFICINA', '" & StatusBar1.Panels(2).Text & "')"
   
   MostrarGrid_Pecas
   LimparObjetos_Pecas
   txtPecas.SetFocus
End Sub

Private Sub Verifica_QuantEstoque()
   
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim oCfg As ConfigItem
   Dim bEstNeg As Boolean
   
   If txtCodPeca.Text = "" Then Exit Sub
   If txtQuantPeca.Text = "" Then txtQuantPeca.Text = 1
   
   Set oCfg = sysConfig("ESTOQUE_NEGATIVO")
   bEstNeg = CBool(oCfg.Value)
   Set oCfg = Nothing
   
   If bEstNeg = False Then
      sSQL = "SELECT codigo, quant_estoque FROM produtos WHERE (codigo = " & txtCodPeca.Text & ");"
      Set r = dbData.OpenRecordset(sSQL)
      
      VERIFICAR_QUANTIDADE = False
      
      If Not r.BOF Then
         If r("quant_estoque") < CDbl(txtQuantPeca.Text) And r("quant_estoque") <> 0 Then
            ShowMsg "ESSA QUANTIDADE É INVÁLIDA!" & vbCrLf & "SEU ESTOQUE ATUAL É DE " & r("quant_estoque") & " PRODUTO(S)", vbExclamation
            LimparObjetos_Pecas
            VERIFICAR_QUANTIDADE = True
            
         ElseIf r("quant_estoque") = 0 Then
            ShowMsg "ESSA QUANTIDADE É INVÁLIDA!" & vbCrLf & "SEU ESTOQUE ATUAL É DE 0 PRODUTO(S)", vbExclamation
            LimparObjetos_Pecas
            VERIFICAR_QUANTIDADE = True
            
         End If
      End If
   Else
      Exit Sub
   End If
End Sub

Private Sub cmdAdicionarServicos_Click()
   Dim QUANT As Integer
   Dim Valor As Currency
   Dim Total As Currency
   
   If txtCodServico.Text = "" Or txtCodOS.Text = "" Then Exit Sub
   
   If txtQuantServico.Text = "" Then txtQuantServico.Text = 1
   
   'CHECAR SE A OS ESTÁ FECHADA
   Verificar_OS_Fechada
   If OS_FECHADA = True Then Exit Sub
   
   'VERIFICAR O STATUS
   If cboStatus.Text = "À COMEÇAR" Then
      ShowMsg "Não é possivel adicionar serviços em uma OS com Status = À COMEÇAR!", vbExclamation
      Exit Sub
   End If
   
   If txtQuantServico.Text <> "" Or mskValorServico.Text <> "" Then
      QUANT = txtQuantServico.Text
      Valor = mskValorServico.Text
      Total = Valor * QUANT
   End If
   
   'ADICIONAR NA TABELA OS SERVIÇOS
   AutoNumeracao_Servico
   dbData.Execute "INSERT INTO os_servicos (codigo, cod_os, descricao, preco, quantidade, total, data) VALUES(" & _
      xServ & ", " & txtCodOS.Text & ", '" & cboServicos.Text & "', " & Replace(CCur(mskValorServico.Text), ",", ".") & ", " & _
      txtQuantServico.Text & ", " & Replace(Total, ",", ".") & ", CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103))"
   
   MostrarGrid_Servicos
   LimparObjetos_Servicos
   cboServicos.SetFocus
End Sub

Private Sub cmdAlterar_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtCodOS.Text = "" Then
      ShowMsg "OS VAZIA! Selecione uma OS na guia FILTRO!", vbInformation
      Exit Sub
   End If
   
   If txtCodCliente.Text = "" Then
      ShowMsg "Este cliente não encontra-se cadastrado!", vbInformation
      cboCliente.SetFocus
      Exit Sub
   End If
   
   If txtCodFuncionario.Text = "" Then
      ShowMsg "Este funcionário não encontra-se cadastrado!", vbInformation
      cboFuncionario.SetFocus
      Exit Sub
   End If
   
   If cboStatus.Text = "EM EXECUÇÃO" Or cboStatus.Text = "AGUARDANDO" Or cboStatus.Text = "TERMINADO" Then
      If cboMecanico.Text = "" Then
         ShowMsg "Indique o nome do mecânico!", vbInformation
         cboMecanico.SetFocus
         Exit Sub
      End If
   End If
   
   'If cboStatus.Text = "TERMINADO" Then
   '   dbData.Execute "UPDATE os SET data_saida = '" & Format(Date, ocDATA_EUA) & "', hora_Saida = '" & Format(Now, ocHORA) & "' WHERE (codigo = " & txtCodOS.Text & ");"
   'End If
   
   'EDITAR NA TABELA OS
   'sSQL = "SELECT * FROM os WHERE (codigo = " & txtCodOS.Text & ");"
   'Set r = dbData.OpenRecordset(sSQL)
    
   'Faz a atualização de forma direta e verifica se houve algum erro
   If Not Atualizar_Dados_OS Then
      ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
   'editar tabela pedidos
   dbData.Execute "UPDATE pedidos SET cod_cliente = " & txtCodCliente.Text & " WHERE (cod_pedido = " & txtCodOS.Text & ");"
   
   If cboStatus.Text <> "EM EXECUÇÃO" Then
      'LimparObjetos_Entrada
      'LimparObjetos_Servicos
      'LimparObjetos_Pecas
      'LimparObjetos_Pgto
      'txtCodOS.Text = ""
      'Form_Load
   End If
   
   MostrarGrid_OS
   
   'CHECAR SE A OS ESTÁ FECHADA & PAGA
   Verificar_OS_FechadaePaga
   
   If OS_FECHADAPAGA = True Then
      If cboStatus.Text = "TERMINADO" Then
         cmdFinalizarAV.Visible = True
         cmdFinalizarAP.Visible = True
      End If
   Else
      cmdFinalizarAV.Visible = False
      cmdFinalizarAP.Visible = False
   End If
End Sub

Private Sub cmdApagar_Click()
   
   If txtCodOS.Text = "" Or txtCodCliente.Text = "" Or txtCodFuncionario.Text = "" Then Exit Sub
   
   If ShowMsg("Excluir essa O.S. ?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
   
   Retorna_Produtos_Estoque
   
   'EXCLUIR NA TABELA OS
   dbData.Execute "DELETE FROM OS_Carro_Cadastro WHERE (cod_os = " & txtCodOS.Text & ");"
   
   'EXCLUIR NA TABELA OS_SERVICOS
   dbData.Execute "DELETE FROM os_servicos WHERE (cod_os = " & txtCodOS.Text & ");"
   
   'EXCLUIR NA TABELA PEDIDOS_ITENS
   dbData.Execute "DELETE FROM pedidos_itens WHERE (cod_pedido = " & txtCodOS.Text & ");"
   
   'EXCLUIR NA TABELA PEDIDOS
   dbData.Execute "DELETE FROM pedidos WHERE (cod_pedido = " & txtCodOS.Text & ");"
   
   LimparObjetos_Entrada
   LimparObjetos_Servicos
   LimparObjetos_Pecas
   LimparGrid_Acessorios
   'LimparObjetos_Pgto
   
   txtCodOS.Text = ""
   Form_Load
   MostrarGrid_OS
End Sub

Private Sub Retorna_Produtos_Estoque()
   Dim i As Integer
   
   For i = 1 To Grid_Pecas.rows - 1
      dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque + " & Replace(CDbl(Grid_Pecas.TextMatrix(i, 5)), ",", ".") & " WHERE (codigo = " & Grid_Pecas.TextMatrix(i, 2) & ");"
   Next
End Sub

Private Sub cmdAVfinalizar_Click()
   Dim NumCopias As Integer
   Dim ii As Integer
   
   Dim varTipoCartao As String
   Dim var_PAGAMENTO As String
   Dim oCfg As ConfigItem
   
   Dim bConfFechOS As Boolean
   Dim iCopiasAV As Integer
   Dim bEntregaAV As Boolean
   Dim bImprAV As Integer
   Dim bConfImprAV As Boolean
   
   If txtTotalDescAV.Text = "" Then Exit Sub
   If txtCodOS.Text = "" Then Exit Sub
   
   If txtFuncAV.Text = "" Then
      ShowMsg "Digite o código do funcionário!", vbInformation
      txtCodFuncAV.SetFocus
      Exit Sub
   End If
   
   If optAVdinheiro.Value = True Then
      var_PAGAMENTO = "DINHEIRO"
   ElseIf optAVcheque.Value = True Then
      var_PAGAMENTO = "CHEQUE"
   ElseIf optAVcartao.Value = True Then
      var_PAGAMENTO = "CARTAO"
   End If
     
   varTipoCartao = ""
   If optAVcartao.Value = True Then
      If optDebito.Value = True Then
         varTipoCartao = "D"
      ElseIf Me.optCredito.Value = True Then
         varTipoCartao = "C"
      End If
   End If
   
   'MOSTRAR SE O CAIXA ESTÁ FECHADO
   Verificar_Caixa
   If CAIXA_FECHADO = True Then Exit Sub
   
   Set oCfg = sysConfig("CONF_FECHAMENTO_AV")
   bConfFechOS = CBool(oCfg.Value)
   Set oCfg = Nothing
   
   If bConfFechOS = True Then
      If ShowMsg("Deseja finalizar essa OS?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
   End If
   
   'Retirar da tabela PRODUTOS as QUANTIDADES mencionadas no grid
   For i = 1 To Grid_Pecas.rows - 1
      dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque - " & Replace(CDbl(Grid_Pecas.TextMatrix(i, 5)), ",", ".") & " WHERE (codigo = " & Grid_Pecas.TextMatrix(i, 2) & ");"
   Next
   
   'colocar a data da Ultima compra de cada produro
   For i = 1 To Grid_Pecas.rows - 1
      dbData.Execute "UPDATE produtos SET ult_compra = CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103) WHERE (codigo = " & Grid_Pecas.TextMatrix(i, 2) & ");"
   Next
   
   'ATUALIZANDO A TABELA OS
   dbData.Execute "UPDATE OS_Carro_Cadastro SET status_os = 1 WHERE (cod_os = " & txtCodOS.Text & ");"
   
   'ATUALIZANDO A TABELA PEDIDOS                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       'IF optAVcartao.VALUE = FALSE THEN tipo_cartao = "" ELSE
   dbData.Execute "UPDATE pedidos SET cod_pedido = " & txtCodOS.Text & ", data_compra = CONVERT(DATETIME, '" & Format(mskDataEntrada.Text, ocDATA) & "', 103), data_entrega = CONVERT(DATETIME, '" & Format(mskDataEntrada.Text, ocDATA) & "', 103), tipo_desc = '" & IIf(optDescRSAV.Value = True, "R", "P") & "', valor_desc = " & Replace(CCur(txtDescAV.Text), ",", ".") & ", tipo_acrescimo = '" & IIf(optAcrescRSAV.Value = True, "R", "P") & "', valor_acrescimo = " & Replace(CCur(txtAcrescAV.Text), ",", ".") & ", subtotal = " & Replace(CCur(txtSubTotalAV.Text), ",", ".") & ", total = " & Replace(CCur(txtTotalDescAV.Text), ",", ".") & ", valor_parc = " & Replace(CCur(txtTotalDescAV.Text), ",", ".") & ", tipo_pagamento = 'À Vista', pagamento = '" & var_PAGAMENTO & "', tipo_cartao = '" & varTipoCartao & "', cod_funcionario = " & txtCodFuncAV.Text & ", tipo_pedido = 'OFICINA', vencimento = NULL, maquina = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "', status_pedido = 1 " & _
      "WHERE (cod_pedido = " & txtCodOS.Text & ");"
   
   'Criando as Parcelas
   Autonumeracao_Parcelas
   dbData.Execute "INSERT INTO parcelas (codigo, cod_pedido, numero, data, valor) VALUES(" & xParc & ", " & txtCodOS.Text & ", 1, CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), " & Replace(CCur(txtTotalDescAV.Text), ",", ".") & ")"
   
   'Colocando a data da ultima compra
   dbData.Execute "UPDATE cliente SET ultima_compra = CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103) WHERE (codigo = " & txtCodCliente.Text & ");"
   
   'Colocando a data em cada produto
   dbData.Execute "UPDATE pedidos_itens SET data = CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103) WHERE (cod_pedido = " & txtCodOS.Text & ");"
   
   'dar baixa na parcela de entrada ou compra à vista
   dbData.Execute "UPDATE parcelas SET status = 1 , valor_final = " & Replace(CCur(txtTotalDescAV.Text), ",", ".") & ", pagamento = CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), hora = '" & Format(Now, ocHORA) & "', forma_pgto = '" & var_PAGAMENTO & "', maquina = '" & StatusBar1.Panels(2).Text & "' WHERE (cod_pedido = " & txtCodOS.Text & ") AND (numero = 1)"
   
   Set oCfg = sysConfig("COPIAS_AV")
   iCopiasAV = CInt(oCfg.Value)
   
   Set oCfg = sysConfig("ENTREGA_AV")
   bEntregaAV = CBool(oCfg.Value)
   
   Set oCfg = sysConfig("IMP_AV")
   bImprAV = CBool(oCfg.Value)
   
   Set oCfg = sysConfig("CONF_IMPRESSAO_AV")
   bConfImprAV = CBool(oCfg.Value)
   
   'impressão
   If iCopiasAV <> 0 Then  'saber a quantidade de copias
      If bEntregaAV = True Then
         If MsgBox("Desesa Imprimir o pedido para ENTREGAR?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            NumCopias = iCopiasAV + 1
         Else
            NumCopias = iCopiasAV
         End If
      Else
         NumCopias = iCopiasAV
      End If
   Else
      NumCopias = "1"
   End If
   
   If bImprAV = True Then       'Confirma se vai ter impressão
      If bConfImprAV = True Then
         If ShowMsg("Desesa Imprimir o pedido?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            For ii = 1 To NumCopias
               Imprimir_Pedido
            Next
         End If
      Else
         For ii = 1 To NumCopias
            Imprimir_Pedido
         Next
      End If
   End If
   
   'LimparObjetos_Pedido
   LimparObjetos_Avista
   'LimparGrid_Pedido
   
   'txtTotalGeral.Text = ""
   MostrarGrid_OS
   txtCodOS.Text = ""
   frmVendaAvista.Visible = False
   SSTab1.Tab = 0
End Sub

Private Sub cmdCadastrarAcessorios_Click()
   OS_Automoveis_Acessorios.Show 1
End Sub

Private Sub cmdCadastrarCliente_Click()
   Clientes_Cadastro.Show 1
End Sub

Private Sub cmdCancelar_Click()
   LimparObjetos_Prazo
   frmVendaPrazo.Visible = False
   txtTotalGeral.Text = Format(txtSubtotal.Text, "##,##0.00")
End Sub

Private Sub cmdExibir2_Click()
   MostrarGrid_OS
End Sub

Private Sub cmdFinalizar_Click()
   Dim NumCopias As Integer
   Dim ii As Integer
   
   Dim var_Vencimento As Date
   Dim Var_NumParc As Integer
   Dim var_PAGAMENTO As String
   
   Dim oCfg As ConfigItem
   Dim bConfFechAP As Boolean
   Dim iCopiasAP As Integer
   Dim bEntregaAP As Boolean
   Dim bImprAP As Integer
   Dim bConfImprAP As Boolean
   
   If txtTotalGeral.Text = "" Then Exit Sub
   If txtCodOS.Text = "" Then Exit Sub
   
   If cboCliente.Text = "" Then
      ShowMsg "Escolha o nome do cliente!", vbExclamation
      cboCliente.SetFocus
      Exit Sub
   End If
   
   If txtFuncAP.Text = "" Then
      ShowMsg "Digite o código do funcionário!", vbInformation
      txtCodFuncAP.SetFocus
      Exit Sub
   End If
   
   If txtCodCliente.Text = "" Then CboCliente_LostFocus
   
   If optAvulso.Value = True Then
      var_PAGAMENTO = "AVULSO"
   ElseIf optPromissoria.Value = True Then
      var_PAGAMENTO = "PROMISSORIA"
   ElseIf optCheque.Value = True Then
      var_PAGAMENTO = "CHEQUE"
   End If
   
   'MOSTRAR SE O CAIXA ESTÁ FECHADO
   Verificar_Caixa
   If CAIXA_FECHADO = True Then Exit Sub
   
   Set oCfg = sysConfig("CONF_FECHAMENTO_AP")
   bConfFechAP = CBool(oCfg.Value)
   
   If bConfFechAP = True Then
      If ShowMsg("Deseja finalizar essa compra?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
   End If
   
   'Retirar da tabela PRODUTOS as QUANTIDADES mencionadas no grid
   For i = 1 To Grid_Pecas.rows - 1
      dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque - " & Replace(CDbl(Grid_Pecas.TextMatrix(i, 5)), ",", ".") & " WHERE (codigo = " & Grid_Pecas.TextMatrix(i, 2) & ");"
   Next
   
   'colocar a data da Ultima compra de cada produto
   For i = 1 To Grid_Pecas.rows - 1
      dbData.Execute "UPDATE produtos SET ult_compra = CONVERT(DATETIME, '" & Format(Date, ocDATA_EUA) & "', 103) WHERE (codigo = " & Grid_Pecas.TextMatrix(i, 2) & ");"
   Next
   
   'ATUALIZAR A TABELA OS
   dbData.Execute "UPDATE OS_Carro_Cadastro SET status_os = 1 WHERE (cod_os = " & txtCodOS.Text & ");"
   
   If txtEntrada.Text <> "0,00" And txtValorParc.Text <> "0,00" Then
      'ATUALIZANDO A TABELA PEDIDOS
      dbData.Execute "UPDATE pedidos SET tipo_desc = '" & IIf(optDescRS.Value = True, "R", "P") & "', valor_desc = " & Replace(CCur(txtDesc.Text), ",", ".") & ", tipo_acrescimo = '" & IIf(optAscrescRS.Value = True, "R", "P") & "', valor_acrescimo = " & Replace(CCur(txtAcresc.Text), ",", ".") & ", subtotal = " & Replace(CCur(txtSubtotal.Text), ",", ".") & ", total = " & Replace(CCur(txtTotalDesc.Text), ",", ".") & ", valor_parc = " & Replace(CCur(txtValorParc.Text), ",", ".") & ", tipo_pagamento = 'À Prazo', pagamento = '" & var_PAGAMENTO & "', entrada = " & Replace(CCur(txtEntrada.Text), ",", ".") & ", prazo = " & cboPrazo.Text & ", tipo_pedido = 'OFICINA', maquina = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "', status_pedido = 1, vencimento = CONVERT(DATETIME, '" & Format(mskInicio, ocDATA) & "', 103) WHERE (cod_pedido = " & txtCodOS.Text & ");"
      
      'CRIANDO AS PARCELAS
      'ENTRADA
      Autonumeracao_Parcelas
      dbData.Execute "INSERT INTO parcelas (codigo, cod_pedido, numero, data, valor) VALUES(" & xParc & ", " & txtCodOS.Text & ", 1, CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), '" & txtEntrada.Text & "');"
      
      'PARCELAS
      var_Vencimento = Format(DateAdd("m", Val(1), mskInicio.Text), "dd/mm/yy")
      Var_NumParc = 2
      For i = 1 To CInt(cboQuantParc)
         Autonumeracao_Parcelas
         dbData.Execute "INSERT INTO parcelas (codigo, cod_pedido, numero, data, valor) VALUES(" & xParc & ", " & txtCodOS.Text & ", " & Var_NumParc & ", CONVERT(DATETIME, '" & Format(var_Vencimento, ocDATA) & "', 103), '" & txtValorParc.Text & "');"
         var_Vencimento = Format(DateAdd("m", Val(1), var_Vencimento), "dd/mm/yy")
         Var_NumParc = Var_NumParc + 1
      Next
      
   ElseIf txtEntrada.Text = "0,00" And txtValorParc.Text <> "0,00" Then
      'ATUALIZANDO A TABELA PEDIDOS
      dbData.Execute "UPDATE pedidos SET tipo_desc = '" & IIf(optDescRS.Value = True, "R", "P") & "', valor_desc = " & Replace(CCur(txtDesc.Text), ",", ".") & ", tipo_acrescimo = '" & IIf(optAscrescRS.Value = True, "R", "P") & "', valor_acrescimo = " & Replace(CCur(txtAcresc.Text), ",", ".") & ", subtotal = " & Replace(CCur(txtSubtotal.Text), ",", ".") & ", total = " & Replace(CCur(txtTotalDesc.Text), ",", ".") & ", valor_parc = " & Replace(CCur(txtValorParc.Text), ",", ".") & ", tipo_pagamento = 'À Prazo', pagamento = '" & var_PAGAMENTO & "', entrada = " & Replace(CCur(txtEntrada.Text), ",", ".") & ", prazo = " & cboPrazo.Text & ", tipo_pedido = 'OFICINA', maquina = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "', status_pedido = 1, vencimento = CONVERT(DATETIME, '" & Format(mskInicio, ocDATA_EUA) & "', 103) WHERE (cod_pedido = " & txtCodOS.Text & ");"
      
      'PARCELAS
      var_Vencimento = CDate(mskInicio.Text)
      Var_NumParc = 1
      For i = 1 To CInt(cboQuantParc)
         Autonumeracao_Parcelas
         dbData.Execute "INSERT INTO parcelas (codigo, cod_pedido, numero, data, valor) VALUES(" & xParc & ", " & txtCodOS.Text & ", " & Var_NumParc & ", CONVERT(DATETIME, '" & Format(var_Vencimento, ocDATA) & "', 103), " & Replace(CCur(txtValorParc.Text), ",", ".") & ");"
         var_Vencimento = Format(DateAdd("m", Val(1), var_Vencimento), "dd/mm/yy")
         Var_NumParc = Var_NumParc + 1
      Next
      
   ElseIf txtEntrada.Text <> "0,00" And txtValorParc.Text = "0,00" Then
      dbData.Execute "UPDATE OS_Carro_Cadastro SET tipo_desc = '" & IIf(optDescRS.Value = True, "R", "P") & "', valor_desc = " & Replace(CCur(txtDesc.Text), ",", ".") & ", tipo_acrescimo = '" & IIf(optAscrescRS.Value = True, "R", "P") & "', valor_acrescimo = " & Replace(CCur(txtAcresc.Text), ",", ".") & ", subtotal = " & Replace(CCur(txtSubtotal.Text), ",", ".") & ", total = " & Replace(CCur(txtTotalDesc.Text), ",", ".") & ", valor_parc = " & Replace(CCur(txtValorParc.Text), ",", ".") & ", tipo_pagamento = 'À Prazo', pagamento = '" & var_PAGAMENTO & "', entrada = " & Replace(CCur(txtEntrada.Text), ",", ".") & ", prazo = " & cboPrazo.Text & ", tipo_pedido = 'OFICINA', maquina = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "', status_os = 1, vencimento = CONVERT(DATETIME, '" & Format(mskInicio, ocDATA) & "', 103) WHERE (cod_pedido = " & txtCodOS.Text & ");"
      
      'CRIANDO PARCELAS
      Autonumeracao_Parcelas
      dbData.Execute "INSERT INTO parcelas (codigo, cod_pedido, numero, data, valor) VALUES(" & xParc & ", " & txtCodOS.Text & ", 1, CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), '" & txtEntrada.Text & "')"
   End If
   
   'Colocando a data da ultima compra
   dbData.Execute "UPDATE cliente SET ultima_compra = CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103) WHERE (codigo = " & txtCodCliente.Text & ");"
   
   'Colocando a data em cada produto
   dbData.Execute "UPDATE pedidos_itens SET data = CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103) WHERE (cod_pedido = " & txtCodOS.Text & ");"
   
   'Dar baixa na parcela de entrada ou compra à vista
   If txtEntrada.Text <> "0,00" Then
      dbData.Execute "UPDATE parcelas SET status = 1 , valor_final = " & Replace(CCur(txtEntrada.Text), ",", ".") & ", pagamento = CONVERT(DATETIME, '" & Format(Date, ocDATA_EUA) & "', 103), hora = '" & Format(Now, ocHORA) & "', forma_pgto = 'DINHEIRO', maquina = '" & StatusBar1.Panels(2).Text & "' WHERE (cod_pedido = " & txtCodOS.Text & ") AND (numero = 1);"
   End If
   
   Set oCfg = sysConfig("COPIAS_AP")
   iCopiasAP = CInt(oCfg.Value)
   
   Set oCfg = sysConfig("ENTREGA_AP")
   bEntregaAP = CBool(oCfg.Value)
   
   Set oCfg = sysConfig("IMP_AP")
   bImprAP = CBool(oCfg.Value)
   
   Set oCfg = sysConfig("CONF_IMPRESSAO_AP")
   bConfImprAP = CBool(oCfg.Value)
   
   If iCopiasAP <> 0 Then  'saber a quantidade de copias
      If bEntregaAP = True Then
         If ShowMsg("Desesa Imprimir o pedido para ENTREGAR?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            NumCopias = iCopiasAP + 1
         Else
            NumCopias = iCopiasAP
         End If
      Else
         NumCopias = iCopiasAP
      End If
   Else
      NumCopias = "1"
   End If
   
   If bImprAP = True Then       'Confirma se vai ter impressão
      If bConfImprAP = True Then
         If ShowMsg("Desesa Imprimir o pedido?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            For ii = 1 To NumCopias
               Imprimir_Pedido
            Next
         End If
      Else
         For ii = 1 To NumCopias
            Imprimir_Pedido
         Next
      End If
   End If
   
   LimparObjetos_Entrada
   txtTotalGeral.Text = ""
   LimparObjetos_Prazo
   txtCodOS.Text = ""
   frmVendaPrazo.Visible = False
   MostrarGrid_OS
   SSTab1.Tab = 0
End Sub

Private Sub Imprimir_Pedido()
   If frmVendaPrazo.Visible = True Then
      REL_Pedido_Mod05.loadPedidos txtCodOS.Text, "OFICINA"
      'REL_Pedido_Mod05.Show vbModal
   ElseIf frmVendaAvista.Visible = True Then
      REL_Pedido_Mod06.loadPedidos txtCodOS.Text, "OFICINA"
   'ElseIf frmOrcamento.Visible = True Then
   '   REL_Pedido_Mod06.loadPedidos txtCodOS.Text
   End If
   
'   Unload REL_Pedido_Mod06
'   Me.Show
End Sub

Private Sub cmdFinalizarAP_Click()
   Dim oCfg As ConfigItem
   Dim sDescAP As String
   
   If txtTotalGeral.Text = "" Or txtTotalGeral.Text = "0,00" Then Exit Sub
   
   SSTab1.Tab = 3
   LimparObjetos_Prazo
   frmVendaPrazo.Visible = True
   frmVendaAvista.Visible = False
   'frmOrcamento.Visible = False
   
   Mostrar_ValorRestante
   Calcular_Parcelas
   Calcular_Prazo
   optDescRS.Value = True
   txtSubtotal.Text = txtTotalGeral.Text
   cmdFinalizar.Default = True
   cmdCancelar.Cancel = True
   
   Set oCfg = sysConfig("DESC_AP")
   sDescAP = oCfg.Value
   Set oCfg = Nothing
   
   'mostrar o DESCONTO
   If sDescAP <> "" Then
      txtDesc.Text = Format(CCur(sDescAP), ocMONEY)
   Else
      txtDesc.Text = Format(0, ocMONEY)
   End If
   
   txtAcresc.Text = Format(0, ocMONEY)
   
   'limpar campo funcionario
   'If Not IsNull(RS_Conf!IDENT_PDV) Then
   '   If RS_Conf!IDENT_PDV = "2" Then
   '      txtCodFuncAP.Text = ""
   '      txtFuncAP.Text = ""
   '      txtCodFuncAP.SetFocus
   '   Else
   '      cboCliente.SetFocus
   '   End If
   'End If
End Sub

Private Sub cmdFinalizarAV_Click()
   Dim oCfg As ConfigItem
   Dim sDescAV As String
     
   If txtTotalGeral.Text = "" Or txtTotalGeral.Text = "0,00" Then Exit Sub
   
   frmVendaAvista.Visible = True
   frmVendaPrazo.Visible = False
   'frmOrcamento.Visible = False
   optDescRSAV.Value = True
   'txtSubTotal.Text = txtTotalGeral.Text
   txtSubTotalAV.Text = txtTotalGeral.Text
   cmdAVfinalizar.Default = True
   cmdAVcancelar.Cancel = True
   
   Set oCfg = sysConfig("DESC_AV")
   sDescAV = oCfg.Value
   Set oCfg = Nothing
   
   'mostrar o DESCONTO
   If sDescAV <> "" Then
      txtDescAV.Text = Format(sDescAV, ocMONEY)
   Else
      txtDescAV.Text = Format(0, ocMONEY)
   End If
   
   txtAcrescAV.Text = Format(0, ocMONEY)
   SSTab1.Tab = 3

   'limpar campo funcionario
   'If Not IsNull(RS_Conf!IDENT_PDV) Then
   '   If RS_Conf!IDENT_PDV = "2" Then
   '      txtCodFuncAV.Text = ""
   '      txtFuncAV.Text = ""
   '      txtCodFuncAV.SetFocus
   '   Else
   '      txtRecebido.SetFocus
   '   End If
   'End If
End Sub

Private Sub cmdGarantia_Click()
   'colocar o nome da maquina na barra de status
   Dim var_Impressora As String
   Dim oIni As Ini
   
   Set oIni = New Ini
   oIni.Arquivo = appPathApp & "config.ini"
   var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
   Set oIni = Nothing
   
   If txtCodOS.Text = "" Or txtCodCliente.Text = "" Or cboModelo.Text = "" Or txtDescricao.Text = "" Then
      ShowMsg "Não é possível imprimir uma Ordem de Serviço em branco!", vbInformation
      Exit Sub
   End If
   
   Me.Hide
   
   With REL_Garantia
      .txtNumero.Caption = " " & Format(txtCodOS.Text, "000000")
      .rfCodCliente.Caption = " " & txtCodCliente.Text
      .rfModelo.Caption = " " & UCase(cboFabricante.Text) & "-" & cboModelo.Text
      .frCor.Caption = " " & UCase(cboCor.Text)
      .frPlaca.Caption = " " & UCase(txtPlaca1.Text) & "-" & txtPlaca2.Text
      .rfQuilometragem.Caption = " " & txtKM.Text
      
      .rfQuiloPrimeira.Caption = " " & txtKM.Text + 500
      .rfQuiloSegunda.Caption = " " & .rfQuiloPrimeira.Caption + 1000
      .rfQuiloTerceira.Caption = " " & .rfQuiloSegunda.Caption + 1000
      .rfQuiloQuarta.Caption = " " & .rfQuiloTerceira.Caption + 1000
      
      .Relatorio.NumeroRegistros = 1
      .Relatorio.NomeImpressora = var_Impressora
      .Relatorio.Ativar
   End With
   
   Unload REL_Garantia
   
   Me.Show 1
End Sub

Private Sub cmdGerarEntrada_Click()
   'On Error GoTo TrataErro
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtCodOS.Text = "" Or cboModelo.Text = "" Then
      ShowMsg "Não é possível gerar uma entrada de uma Ordem de Serviço em branco!", vbInformation
      Exit Sub
   End If
   
   If txtDescricao.Text = "" Then
      ShowMsg "Falta a descrição do cliente!", vbInformation
      txtDescricao.SetFocus
      Exit Sub
   End If
   
   'If txtKM.Text = "" Then ShowMsg "Quilometragem não especificada", vbInformation, "Aviso do Sistema": txtKM.SetFocus: Exit Sub
   
   If txtCodCliente.Text = "" Then
      ShowMsg "Cliente não encontra-se cadastrado no sistema", vbInformation
      cboCliente.SetFocus
      Exit Sub
   End If
   
   If txtCodFuncionario.Text = "" Then
      ShowMsg "Funcionário não encontra-se cadastrado no sistema", vbInformation
      cboFuncionario.SetFocus
      Exit Sub
   End If
   
   'Alterando os dados da tabela OS
   'execSQL "UPDATE OS SET DATA_ENTRADA =  #" & Format(mskDataEntrada.Text, "MM/DD/yy") & "#, HORA_ENTRADA = #" & Format(mskHoraEntrada.Text, "hh:mm") & "#, EQUIPAMENTO = '" & cboEquipamento.Text & "', COR = '" & cboCor.Text & "', fabricante = '" & cboFabricante.Text & "', PLACA1 = '" & txtPlaca1.Text & "', PLACA2 = '" & txtPlaca2.Text & "', KM = '" & txtKM.Text & "', COMBUSTIVEL = '" & cboCombustivel.Text & "', COD_FUNCIONARIO = " & txtCodFuncionario.Text & ", DESCRICAO = '" & txtDescricao.Text & "', Status = '" & cboStatus.Text & "' WHERE CODIGO = " & txtCodOS.Text & ""
   
   'sSQL = "SELECT * FROM os WHERE (codigo = " & txtCodOS.Text & ");"
   'Set r = dbData.OpenRecordset(sSQL)
   
   If Not Atualizar_Dados_OS Then
      ShowMsg "Não foi possível atualizar o registro." & vbCr & "Verifique os dados informados e tente novamente.", vbExclamation
      Exit Sub
   End If
   
   'alterar dados do pedido
   dbData.Execute "UPDATE pedidos SET cod_cliente = " & txtCodCliente.Text & ", cod_funcionario = " & txtCodFuncionario.Text & ", data_entrega = CONVERT(DATETIME, '" & Format(StatusBar1.Panels(4).Text, ocDATA) & "', 103), data_compra = CONVERT(DATETIME, '" & Format(StatusBar1.Panels(4).Text, ocDATA) & "', 103) WHERE (cod_pedido = " & txtCodOS.Text & ");"
   
   cmdImprimirEntrada.Enabled = True
   cmdGerarEntrada.Enabled = False
   cmdCancelarEntrada.Enabled = False
   cmdImprimirEntrada.Enabled = True
   cmdNovo.Enabled = True
   MostrarGrid_OS
   
   'LimparObjetos_Entrada
   'LimparObjetos_Servicos
   'LimparObjetos_Pecas
   'txtCodOS.Text = ""
   'Form_Load
   
'TrataErro:
   'If Err.Number = 3022 Then
   '   MsgBox "DADOS DUPLICADO!" & vbCrLf & "Verifique se já está cadastrado.", vbInformation, "Aviso do Sistema"
   '   Exit Sub
   'End If
End Sub

Private Sub cmdImprimirEntrada_Click()
   Dim var_Impressora As String
   Dim oIni As Ini
   
   'colocar o nome da maquina na barra de status
   Set oIni = New Ini
   oIni.Arquivo = appPathApp & "config.ini"
   var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
   Set oIni = Nothing
   
   If txtCodOS.Text = "" Or txtCodCliente.Text = "" Or cboModelo.Text = "" Or txtDescricao.Text = "" Then
      ShowMsg "Não é possível imprimir uma Ordem de Serviço em branco!", vbInformation
      Exit Sub
   End If
   
   Me.Hide
   
   With REL_OS_Entrada_Automoveis
      .txtOS.Caption = " " & Format(txtCodOS.Text, "000000")
      .txtCliente.Caption = " " & UCase(cboCliente.Text)
      .txtSaida.Caption = " " & Format(mskDataSaida.Text, "dd/mm/yy") & " - " & Format(mskHoraSaida.Text, "hh:mm")
      .txtDataEntrada.Caption = " " & Format(mskDataEntrada.Text, "dd/mm/yy") & " - " & Format(mskHoraEntrada.Text, "hh:mm")
      .txtFuncionario.Caption = " " & UCase(cboFuncionario)
      .txtFabricante.Caption = " " & UCase(cboFabricante.Text)
      .txtModelo.Caption = " " & UCase(cboModelo.Text)
      .txtCor.Caption = " " & UCase(cboCor.Text)
      .txtPlaca1.Caption = " " & UCase(txtPlaca1.Text) & "-" & txtPlaca2.Text
      .txtKM.Caption = " " & txtKM.Text
      .txtTanque.Caption = " " & cboTanque.Text
      .txtDescricao.Caption = UCase(txtDescricao.Text)
      .Preencher_Acessorios txtCodOS.Text
      .Relatorio.NumeroRegistros = 1
      .Relatorio.NomeImpressora = var_Impressora
      .Relatorio.Ativar
   End With
   Unload REL_OS_Entrada_Automoveis
   
   Me.Show 1
End Sub

Private Sub cmdNovo_Click()
   LimparObjetos_Entrada
   LimparObjetos_Servicos
   LimparObjetos_Pecas
   AutoNumeracao_OS
   mskDataEntrada.Text = Format(Date, "dd/mm/yy")
   mskDataSaida.Text = Format(Date, "dd/mm/yy")
   mskHoraEntrada.Text = Format(Time, "hh:ss")
   cboStatus_GotFocus
   cboStatus.ListIndex = 0
   
   dbData.Execute "INSERT INTO pedidos (cod_pedido, status_pedido, tipo_pedido) VALUES (" & txtCodOS.Text & ", 0, 'OFICINA')"
   dbData.Execute "INSERT INTO OS_Carro_Cadastro (cod_os) VALUES (" & txtCodOS.Text & ")"
   
   cmdGerarEntrada.Enabled = True
   cmdCancelarEntrada.Enabled = True
   frmSecundario.Enabled = True
   frmPrincipal.Enabled = False
   cmdNovo.Enabled = False
   cmdAlterar.Enabled = False
   cmdApagar.Enabled = False
   cboFuncionario.SetFocus
End Sub

Private Sub cmdPecas_Click()
Produtos_Cadastro.Show 1
End Sub

Private Sub cmdRemoverAcessorios_Click()
   
   On Error GoTo erro
   
   If Grid_Acessorio.TextMatrix(Grid_Acessorio.Row, 1) = "" Then GoSub erro
   If ShowMsg("Deseja excluir o acessório: " & Grid_Acessorio.TextMatrix(Grid_Acessorio.Row, 4) & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
   
   dbData.Execute "DELETE FROM os_acessorios WHERE (codigo = " & Grid_Acessorio.TextMatrix(Grid_Acessorio.Row, 1) & ") AND (cod_os = " & txtCodOS.Text & ");"
   
   MostrarGrid_Acessorios
   Exit Sub
   
erro:
   ShowMsg "Não existe nenhum acessório para ser excluido!", vbExclamation
   Exit Sub
End Sub

Private Sub cmdRemoverPecas_Click()
   On Error GoTo erro
   
   'CHECAR SE A OS ESTÁ FECHADA
   Verificar_OS_Fechada
   If OS_FECHADA = True Then Exit Sub
   
   'REMOVER O ITEM DA LISTA
   If Grid_Pecas.TextMatrix(Grid_Pecas.Row, 1) = "" Then GoSub erro
   If ShowMsg("Deseja remover da lista a peça: " & Grid_Pecas.TextMatrix(Grid_Pecas.Row, 2) & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
   
   dbData.Execute "DELETE FROM pedidos_itens WHERE (codigo = " & Grid_Pecas.TextMatrix(Grid_Pecas.Row, 1) & ") AND (cod_pedido = " & txtCodOS.Text & ");"
   
   MostrarGrid_Pecas
   Exit Sub
   
erro:
   ShowMsg "Não existe nenhuma peça para ser removido!", vbExclamation
   Exit Sub
End Sub

Private Sub cmdRemoverServicos_Click()
   On Error GoTo erro
   
   'CHECAR SE A OS ESTÁ FECHADA
   Verificar_OS_Fechada
   If OS_FECHADA = True Then Exit Sub
   
   'REMOVER O ITEM DA LISTA
   If Grid_Servicos.TextMatrix(Grid_Servicos.Row, 1) = "" Then GoSub erro
   If ShowMsg("Deseja remover da lista o serviço: " & Grid_Servicos.TextMatrix(Grid_Servicos.Row, 2) & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
   
   dbData.Execute "DELETE FROM os_servicos WHERE (codigo = " & Grid_Servicos.TextMatrix(Grid_Servicos.Row, 1) & ") AND (cod_os = " & txtCodOS.Text & ");"
   
   MostrarGrid_Servicos
   Exit Sub
   
erro:
   ShowMsg "Não existe nenhum serviço para ser removido!", vbExclamation
   Exit Sub
End Sub

Private Sub cmdServicos_Click()
   'Ordem_Servicos_Cadastro.Show 1
End Sub

Private Sub Form_Load()
   SSTab1.Tab = 0
   cmdNovo.Enabled = True
   cmdGerarEntrada.Enabled = False
   cmdCancelarEntrada.Enabled = False
   cmdAlterar.Enabled = False
   cmdApagar.Enabled = False
   cmdImprimirEntrada.Enabled = False
   cmdGarantia.Enabled = False
   frmPrincipal.Enabled = False
   frmSecundario.Enabled = False
   txtDesc.Text = 0
   LimparGrid_Acessorios
   LimparGrid_Pecas
   LimparGrid_Servicos
   lblTotal.Caption = FormatCurrency(0)
   lblTotalPeca.Caption = FormatCurrency(0)
   MostrarGrid_OS
   StatusBar1.Panels(4).Text = Format(Date, "dd/mm/yy")
   
   'colocar o nome da maquina na barra de status
   Dim var_Maquina As String
   Dim oIni As Ini
   Dim oCfg As ConfigItem
   
   Set oIni = New Ini
   oIni.Arquivo = appPathApp & "config.ini"
   var_Maquina = oIni.LerTexto("DADOS_MAQUINA", "maquina")
   Set oIni = Nothing
   StatusBar1.Panels(2).Text = var_Maquina
   
   'abrir o banco de dados de configurações
   'sSQL = "SELECT * FROM configuracao WHERE (codgio = 1);"
   'Set r = dbData.OpenRecordset(sSQL)
   
   Set oCfg = sysConfig("TIPO_OS")
      
   If oCfg.Value = "CARROS" Then
      'cmdCadastrarAcessorios.Enabled = True
      'cmdMotos.Enabled = False
      'cmdAutomoveis.Enabled = True
   
   ElseIf oCfg.Value = "MOTOS" Then
      'cmdCadastrarAcessorios.Enabled = True
      'cmdMotos.Enabled = True
      'cmdAutomoveis.Enabled = False
   
   ElseIf oCfg.Value = "INFOR" Then
      'cmdCadastrarAcessorios.Enabled = False
      'cmdMotos.Enabled = False
      'cmdAutomoveis.Enabled = False
   
   Else
      'cmdCadastrarAcessorios.Enabled = False
      'cmdMotos.Enabled = False
      'cmdAutomoveis.Enabled = False
   
   End If
   
   Set moCombo = New cComboHelper
End Sub

Private Sub FormatarGrid_OS(rTabela As ADODB.Recordset)
   Dim i As Integer
   Dim aCor As ColorConstants
   
   With Grid
      .rows = 1       'INICIA O GRID COM UMA LINHA
      .FixedCols = 0  'DETERMINA QUE NÃO HAJA COLUNA FIXA
      
      'Abaixo o cabeçalho é criado
      .FormatString = "^COD OS|^SITUAÇÃO|^STATUS|^CLIENTE"
      .ColWidth(0) = 0
      .ColWidth(1) = 1300
      .ColWidth(2) = 1000
      .ColWidth(3) = 3700
       
       'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            'ALINHAMENTO
            .ColAlignment(3) = 1
            
            'A linha abaixo cria mais linha no grid
            .rows = .rows + 1
            
            'Preenche com os dados, e assim sucessivamente
            .TextMatrix(.rows - 1, 0) = rTabela("cod_os")
            .TextMatrix(.rows - 1, 1) = rTabela("var_status")
            .TextMatrix(.rows - 1, 2) = rTabela("var_status_os") & ""
            .TextMatrix(.rows - 1, 3) = rTabela("nome")
            
            rTabela.MoveNext
         Loop
      End If
      
      'agora sim coloco a fução para mudar a cor da coluna e pronto
      'mudar a cor da fonte
      For i = 1 To .rows - 1
         If UCase(Trim(.TextMatrix(i, 2))) = UCase("ABERTO") Then
            aCor = vbBlue
         Else
            aCor = vbRed
         End If
         
         .Col = 2 'a coluna do aberto ou fechado
         .Row = i
         .CellForeColor = aCor
      Next
      
      'mudar a cor da fonte
      For i = 1 To .rows - 1
         If UCase(Trim(.TextMatrix(i, 1))) = UCase("À COMEÇAR") Then
            aCor = vbBlack
         ElseIf UCase(Trim(.TextMatrix(i, 1))) = UCase("EM EXECUÇÃO") Then
            aCor = vbGreen
         ElseIf UCase(Trim(.TextMatrix(i, 1))) = UCase("AGUARDANDO") Then
            aCor = vbYellow
         ElseIf UCase(Trim(.TextMatrix(i, 1))) = UCase("TERMINADO") Then
            aCor = vbRed
         End If
         
         .Col = 1 'a coluna do aberto ou fechado
         .Row = i
         .CellForeColor = aCor
      Next
      
      .Redraw = True
   End With
End Sub

Public Function SomaGrid(Grid As MSFlexGrid, Col As Integer) As Double
   Dim i As Integer, Valor As Currency
   
   Valor = 0
   For i = 0 To Grid.rows - 1
      If IsNumeric(Grid.TextMatrix(i, Col)) Then
         Valor = Valor + CDbl(Grid.TextMatrix(i, Col))
      End If
   Next
   
   SomaGrid = Valor
End Function

Private Sub Grid_DblClick()
   SSTab1.Tab = 0
   frmSecundario.Enabled = True
   frmPrincipal.Enabled = True
   cmdGerarEntrada.Enabled = False
   cmdCancelarEntrada.Enabled = False
   cmdImprimirEntrada.Enabled = True
   cmdGarantia.Enabled = True
   cmdAlterar.Enabled = True
   cmdApagar.Enabled = True
   cmdNovo.Enabled = True
   txtCodOS.Text = ""
   txtCodOS.Text = (Grid.TextMatrix(Grid.Row, 0))
End Sub

Private Sub cmdCancelarEntrada_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtCodOS.Text = "" Then Exit Sub
   
   'CHECAR SE A OS ESTÁ FECHADA
   sSQL = "SELECT * FROM OS_Carro_Cadastro WHERE (cod_os = " & txtCodOS.Text & ") AND (status_os = 1);"
   Set r = dbData.OpenRecordset(sSQL)
   
   If r.RecordCount > 0 Then
      ShowMsg "Essa OS encontra-se fechada para alterações!", vbInformation
      Exit Sub
   End If
   
   If ShowMsg("Cancelando a OS todos os produtos adicionado até agora serão perdidos!" & vbCrLf & "Deseja cancelar essa OS ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
   
   dbData.Execute "DELETE FROM os_servicos WHERE (codigo = " & txtCodOS.Text & ");"
   dbData.Execute "DELETE FROM OS_Carro_Cadastro WHERE (cod_os = " & txtCodOS.Text & ");"
   
   LimparObjetos_Entrada
   LimparObjetos_Servicos
   LimparObjetos_Pecas
   'LimparObjetos_Pgto
   txtCodOS.Text = ""
   Form_Load
End Sub

Private Sub mskDataSaida_GotFocus()
   SelectControl mskDataSaida
End Sub

Private Sub mskDataSaida_KeyPress(KeyAscii As Integer)
   mskDataSaida.Mask = "##/##/##"
End Sub

Private Sub mskDataSaida_LostFocus()
   If mskDataSaida.Text = "" Or mskDataSaida.Text = "__/__/__" Then
      mskDataSaida.Mask = ""
      mskDataSaida.Text = ""
   Else
      If Not IsDate(mskDataSaida.Text) Then
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskDataSaida.SetFocus
      End If
   End If
End Sub

Private Sub mskHoraSaida_GotFocus()
   SelectControl mskHoraSaida
End Sub

Private Sub mskHoraSaida_KeyPress(KeyAscii As Integer)
   mskHoraSaida.Mask = "##:##"
End Sub

Private Sub mskHoraSaida_LostFocus()
   If mskHoraSaida.Text = "" Or mskHoraSaida.Text = "__:__" Then
      mskHoraSaida.Mask = ""
      mskHoraSaida.Text = ""
   Else
      If Not IsDate(mskHoraSaida.Text) Then
         ShowMsg "HORA INVÁLIDA!" & vbCrLf & "A hora digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskHoraSaida.SetFocus
      End If
   End If
End Sub

Private Sub mskInicio_GotFocus()
   Calcular_Prazo
   SelectControl mskInicio
End Sub

Private Sub mskInicio_KeyPress(KeyAscii As Integer)
   If Not IsDate(mskInicio.Text) Then Exit Sub
   mskInicio.Mask = "##/##/##"
End Sub

Private Sub mskInicio_LostFocus()
   If cboPrazo.Text = "" Then Exit Sub
   
   If txtEntrada.Text = "0,00" Or txtEntrada.Text = "" Or Not IsDate(mskInicio) = True Then
      mskTermino.Text = Format(DateAdd("m", Val(cboQuantParc.Text) - 1, mskInicio.Text), "dd/mm/yy")
   Else
      mskTermino.Text = Format(DateAdd("m", Val(cboQuantParc.Text), mskInicio.Text), "dd/mm/yy")
   End If
End Sub

Private Sub mskTermino_Change()
   If Not IsDate(mskTermino.Text) Then Exit Sub
   mskTermino.Mask = "##/##/##"
End Sub

Private Sub mskTermino_LostFocus()
   SelectControl mskTermino
End Sub

Private Sub OptAbertos_Click()
   MostrarGrid_OS
End Sub

Private Sub optAcrescPorcAV_Click()
   Calcular_DescontoAV
   txtAcrescAV.SetFocus
End Sub

Private Sub optAcrescRSAV_Click()
   Calcular_DescontoAV
   txtAcrescAV.SetFocus
End Sub

Private Sub optAscrescPorc_Click()
   Calcular_DescontoAP
   txtAcresc.SetFocus
End Sub

Private Sub optAscrescRS_Click()
   Calcular_DescontoAP
   txtAcresc.SetFocus
End Sub

Private Sub optAVcartao_Click()
   frmCartao.Visible = True
End Sub

Private Sub optAVcheque_Click()
   frmCartao.Visible = False
End Sub

Private Sub optAVdinheiro_Click()
   frmCartao.Visible = False
End Sub

Private Sub optAvulso_Click()
   txtSubtotal.SetFocus
End Sub

Private Sub optCheque_Click()
   txtSubtotal.SetFocus
End Sub

Private Sub optCredito_Click()
   txtDescAV.Text = Format(0, "##,##0.00")
End Sub

Private Sub optDebito_Click()
   Dim oCfg As ConfigItem
   
   Set oCfg = sysConfig("DESC_AV")
   If oCfg.Value <> "" Then txtDescAV.Text = Format(oCfg.Value, ocMONEY)
   Set oCfg = Nothing
End Sub

Private Sub optDescPorc_Click()
   txtDesc_Change
   txtDesc.SetFocus
End Sub

Private Sub optDescPorc_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyTab Then
    txtDesc.SetFocus
    End If
End Sub

Private Sub optDescPorcAV_Click()
   Calcular_DescontoAV
   txtDescAV.SetFocus
End Sub

Private Sub optDescRS_Click()
   txtDesc_Change
   txtDesc.SetFocus
End Sub

Private Sub optDescRS_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyTab Then txtDesc.SetFocus
End Sub

Private Sub optDescRSAV_Click()
   Calcular_DescontoAV
   txtDescAV.SetFocus
End Sub

Private Sub optFechados_Click()
   MostrarGrid_OS
End Sub

Private Sub optPromissoria_Click()
   txtSubtotal.SetFocus
End Sub

Private Sub optStatusAguardando_Click()
   MostrarGrid_OS
End Sub

Private Sub optStatusComecar_Click()
   MostrarGrid_OS
End Sub

Private Sub optStatusExecucao_Click()
   MostrarGrid_OS
End Sub

Private Sub optStatusTerminado_Click()
   MostrarGrid_OS
End Sub

Private Sub optStatusTodos_Click()
   MostrarGrid_OS
End Sub

Private Sub optTodos_Click()
   MostrarGrid_OS
End Sub

Private Sub txtAcresc_Change()
   'On Error GoTo Erro
   
   If txtAcresc.Text = "" Or txtSubtotal.Text = "" Then
      txtAcresc.Text = 0
      SelectControl txtAcresc
      Exit Sub
   End If
   
   Calcular_DescontoAP
   Exit Sub
   
'Erro:
'   ShowMsg "O valor digitado é inválido!", vbExclamation
'   txtAcresc.Text = 0
End Sub

Private Sub txtAcresc_GotFocus()
   SelectControl txtAcresc
End Sub

Private Sub txtAcresc_KeyPress(KeyAscii As Integer)
   KeyAscii = aNumeros(KeyAscii, True)
End Sub

Private Sub txtAcresc_LostFocus()
   'On Error GoTo Erro
   
   If txtAcresc.Text = "" Or txtSubtotal.Text = "" Then
      txtAcresc.Text = 0
      SelectControl txtAcresc
      Exit Sub
   End If
   
   Calcular_DescontoAP
   txtAcresc.Text = Format(txtAcresc.Text, ocMONEY)
   Exit Sub
   
'Erro:
'   ShowMsg "O valor digitado é inválido!", vbExclamation, "Aviso do Sistema"
'   txtAcresc.Text = 0
End Sub

Private Sub txtAcrescAV_Change()
   On Error GoTo erro
   
   If txtAcrescAV.Text = "" Or txtSubTotalAV.Text = "" Then
      txtAcrescAV.Text = "0"
      SelectControl txtAcrescAV
      Exit Sub
   End If
   
   Calcular_DescontoAV
   Exit Sub
   
erro:
   ShowMsg "O valor digitado é inválido txtAcrescAV!", vbExclamation
   txtAcrescAV.Text = 0
End Sub

Private Sub txtAcrescAV_GotFocus()
   SelectControl txtAcrescAV
End Sub

Private Sub txtAcrescAV_KeyPress(KeyAscii As Integer)
   KeyAscii = aNumeros(KeyAscii, True)
End Sub

Private Sub txtAcrescAV_LostFocus()
   On Error GoTo erro
   
   If txtAcrescAV.Text = "" Or txtSubTotalAV.Text = "" Then
      txtAcrescAV.Text = 0
      SelectControl txtAcrescAV
      Exit Sub
   End If
   
   Calcular_DescontoAV
   txtAcrescAV.Text = Format(txtAcrescAV.Text, ocMONEY)
   Exit Sub
   
erro:
   ShowMsg "O valor digitado é inválido txtAcrescAV_L!", vbExclamation
   txtAcrescAV.Text = 0
End Sub

Private Sub txtCodFuncAP_Change()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtCodFuncAP.Text = "" Then Exit Sub
   txtFuncAP.Text = ""
   
   sSQL = "SELECT codigo, nome, sobrenome FROM funcionario WHERE (codigo = " & txtCodFuncAP.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then txtFuncAP.Text = r("nome") & " " & r("sobrenome")
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub txtCodFuncAP_KeyPress(KeyAscii As Integer)
   KeyAscii = aNumeros(KeyAscii, True)
End Sub

Private Sub txtCodFuncAV_Change()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtCodFuncAV.Text = "" Then Exit Sub
   txtFuncAV.Text = ""
   
   sSQL = "SELECT codigo, nome, sobrenome FROM funcionario WHERE (codigo = " & txtCodFuncAV.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then txtFuncAV.Text = r("nome") & " " & r("sobrenome")
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub txtCodFuncAV_KeyPress(KeyAscii As Integer)
   KeyAscii = aNumeros(KeyAscii, True)
End Sub

Private Sub txtDesc_Change()
   On Error GoTo erro
   
   If txtDesc.Text = "" Or txtSubtotal.Text = "" Then
      txtDesc.Text = 0
      SelectControl txtDesc
      Exit Sub
   End If
   
   Calcular_DescontoAP
   Exit Sub
   
erro:
   ShowMsg "O valor digitado é inválido!", vbExclamation
   txtDesc.Text = 0
End Sub

Private Sub txtDesc_GotFocus()
   SelectControl txtDesc
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
   KeyAscii = aNumeros(KeyAscii, True)
End Sub

Private Sub txtDesc_LostFocus()
   On Error GoTo erro
   
   If txtDesc.Text = "" Or txtSubtotal.Text = "" Then
      txtDesc.Text = 0
      SelectControl txtDesc
      Exit Sub
   End If
   
   Calcular_DescontoAP
   txtDesc.Text = Format(txtDesc.Text, ocMONEY)
   Exit Sub
   
erro:
   ShowMsg "O valor digitado é inválido!", vbExclamation
   txtDesc.Text = 0
End Sub

Private Sub txtDescAV_Change()
   On Error GoTo erro
   
   If txtDescAV.Text = "" Or txtSubTotalAV.Text = "" Then
      txtDescAV.Text = "0"
      SelectControl txtDescAV
      Exit Sub
   End If
   
   Calcular_DescontoAV
   Exit Sub
   
erro:
   ShowMsg "O valor digitado é inválido! txtDescAV_C", vbExclamation
   txtDescAV.Text = 0
End Sub

Private Sub txtDescAV_GotFocus()
   SelectControl txtDescAV
End Sub

Private Sub txtDescAV_KeyPress(KeyAscii As Integer)
   KeyAscii = aNumeros(KeyAscii, True)
End Sub

Private Sub txtDescAV_LostFocus()
   On Error GoTo erro
   
   If txtDescAV.Text = "" Or txtSubTotalAV.Text = "" Then
      txtDescAV.Text = 0
      SelectControl txtDescAV
      Exit Sub
   End If
   
   Calcular_DescontoAV
   txtDescAV.Text = Format(txtDescAV.Text, ocMONEY)
   Exit Sub
   
erro:
   ShowMsg "O valor digitado é inválido txtDescAV_L!", vbExclamation
   txtDescAV.Text = 0
End Sub

Private Sub txtEntrada_Change()
   txtEntrada_Click
End Sub

Private Sub txtEntrada_Click()
   If txtTotalGeral.Text = "" Then
      Exit Sub
   Else
      Mostrar_ValorRestante
      Calcular_Parcelas
      Calcular_Prazo
   End If
End Sub

Private Sub txtEntrada_GotFocus()
   SelectControl txtEntrada
End Sub

Private Sub txtEntrada_KeyPress(KeyAscii As Integer)
   KeyAscii = aNumeros(KeyAscii, True)
End Sub

Private Sub txtEntrada_LostFocus()
   txtEntrada_Click
   
   If txtEntrada = "" Then
      txtEntrada = Format(0, ocMONEY)
   Else
      txtEntrada = Format(txtEntrada, ocMONEY)
   End If
End Sub

Private Sub txtLocalizarCodOS_GotFocus()
   SelectControl txtLocalizarCodOS
End Sub

Private Sub txtPecas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKey ocKEYTAB
   End If
End Sub

Private Sub txtPecas_Validate(Cancel As Boolean)
Dim sSQL As String
Dim r As ADODB.Recordset

'lstBusca.Visible = False
   
Dim ItemLst As ListItem
Dim fGrid As Object
Dim bCancel As Boolean
Dim vProd() As String
Dim rPos As RECT
Dim lLft As Long, lTop As Long

Dim cCfg As ConfigItem
Dim tipoEmpresa As Integer

Set cCfg = sysConfig("TIPO_EMPRESA")
tipoEmpresa = cCfg.Value
Set cCfg = Nothing

If txtPecas.Text = "" Then Exit Sub
If txtPecas.Text <> "" And txtCodPeca.Text <> "" Then Exit Sub

If txtPecas.Text <> "" And txtCodPeca.Text = "" Then
   DoEvents
   'lblInfoBusca.Visible = True
   'lblInfoBusca.Refresh
   Screen.MousePointer = vbHourglass
   
   'Otimizando a conslta
   sSQL = "SELECT DISTINCT produtos.codigo AS var_cod, produtos.ref AS var_ref, produtos.tamanho AS var_tam, produtos.fabricante AS var_fab, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, " & _
      "produtos.quant_estoque AS var_quant, (SELECT  TOP 1 produtos_entrada_itens.venda FROM produtos_entrada_itens " & _
      "left JOIN produtos_entrada ON produtos_entrada_itens.codigo_entrada = produtos_entrada.codigo " & _
      "WHERE produtos_entrada_itens.CodigoProduto = produtos.codigo ORDER BY " & _
      "produtos_entrada.data_entrada DESC, produtos_entrada.hora_entrada) AS venda " & _
      "FROM produtos WHERE (descricao LIKE '%" & txtPecas.Text & "%') AND (produtos.ativo = 1) " & _
      "ORDER BY descricao;"
   
   Set r = dbData.OpenRecordset(sSQL)
End If
   GetWindowRect txtPecas.hwnd, rPos
   lLft = rPos.Left * Screen.TwipsPerPixelX - 160
   lTop = rPos.Top * Screen.TwipsPerPixelY + txtPecas.Height
   
   If tipoEmpresa = 5 Then
      Set fGrid = BuscaGrid_Automotivo
   Else
      'Set fGrid = BuscaGrid_Comum
   End If
   
   Load fGrid
   LockWindowUpdate fGrid.lstBusca.hwnd
   
If txtPecas.Text <> "" Then
   If Not r Is Nothing Then
      Do While Not r.EOF
         'primeira coluna
         Set ItemLst = fGrid.lstBusca.ListItems.Add(, , r("var_cod"))
         'segunda e terceira coluna, que são sub itens da coluna 1
         ItemLst.SubItems(1) = r("var_codbarra")
         ItemLst.SubItems(2) = r("var_desc")
         If Not IsNull(r("var_quant")) Then ItemLst.SubItems(4) = r("var_quant")
         If Not IsNull(r("venda")) Then ItemLst.SubItems(5) = Format(r("venda"), ocMONEY)
         
         'Compartibilidade
         Dim sSQL_Comp As String
         Dim var_Comp As String
         Dim rS2 As ADODB.Recordset
         
         sSQL_Comp = "Select MODELO, ANO From PRODUTOS_COMP Where COD_PRODUTO = " & r("var_cod")
         Set rS2 = dbData.OpenRecordset(sSQL_Comp)
         
         Do While Not rS2.EOF
         var_Comp = var_Comp & rS2!Modelo & "(" & rS2!ANO & "),  "
         rS2.MoveNext
         Loop
         
         If Not IsNull(var_Comp) Then ItemLst.SubItems(3) = var_Comp
         var_Comp = ""
         
         r.MoveNext
      Loop
      
      If r.State <> 0 Then r.Close
      Set r = Nothing
   End If
End If

   'lblInfoBusca.Visible = False
   Screen.MousePointer = vbDefault
   
   LockWindowUpdate 0
   fGrid.Move lLft, lTop
   fGrid.Show vbModal
   
   bCancel = fGrid.Cancelled
   vProd = fGrid.InfoProduct
   
   Unload fGrid
   Set fGrid = Nothing
   
   If Not bCancel Then
      txtCodPeca.Text = vProd(1)      'lstBusca.SelectedItem
      txtPecas.Text = vProd(3)        'lstBusca.SelectedItem.ListSubItems.Item(1).Text
      mskValorPeca.Text = vProd(5)    'lstBusca.SelectedItem.ListSubItems.Item(2).Text
      
      Cancel = True
      'GoTo ValidarBusca
   End If
End Sub

Private Sub txtPlaca1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtRecebido_Change()
   Calcular_Troco
End Sub

Private Sub txtRecebido_GotFocus()
   SelectControl txtRecebido
End Sub

Private Sub txtRecebido_KeyPress(KeyAscii As Integer)
   KeyAscii = aNumeros(KeyAscii, True)
   If KeyAscii = 13 Then txtRecebido_LostFocus
End Sub

Private Sub txtRecebido_LostFocus()
   If txtRecebido.Text = "" Then
      txtRecebido.Text = Format(0, ocMONEY)
   Else
      txtRecebido.Text = Format(txtRecebido.Text, ocMONEY)
   End If
   
   Calcular_Troco
End Sub

Private Sub Calcular_DescontoAP()

   If txtSubtotal.Text = "" Or txtSubtotal.Text = "0,00" Then Exit Sub
   If txtDesc.Text = "" Then txtDesc.Text = Format(0, ocMONEY)
   If txtAcresc.Text = "" Then txtAcresc.Text = Format(0, ocMONEY)
   
   If txtDesc.Text <> "0,00" And txtAcresc.Text = "0,00" Then
      If optDescRS.Value = True Then
         txtTotalDesc.Text = Format(CCur(txtSubtotal.Text) - CCur(txtDesc.Text), ocMONEY)
      ElseIf optDescPorc.Value = True Then
         txtTotalDesc.Text = Format(CCur(txtSubtotal.Text) - ((CCur(txtSubtotal.Text) * CCur(txtDesc.Text)) / 100), ocMONEY)
      End If
      
   ElseIf txtAcresc.Text <> "0,00" And txtDesc.Text = "0,00" Then
      If optAscrescRS.Value = True Then
         txtTotalDesc.Text = Format(CCur(txtSubtotal.Text) + CCur(txtAcresc.Text), ocMONEY)
      ElseIf optAscrescPorc.Value = True Then
         txtTotalDesc.Text = Format(CCur(txtSubtotal.Text) + ((CCur(txtSubtotal.Text) * CCur(txtAcresc.Text)) / 100), ocMONEY)
      End If
      
   Else
      txtTotalDesc.Text = Format(txtSubtotal.Text, ocMONEY)
      'optDescRS.Value = True
      'optAscrescRS.Value = True
   End If
   
   Mostrar_ValorRestante
End Sub

Private Sub Calcular_DescontoAV()
   If txtSubTotalAV.Text = "" Or txtSubTotalAV.Text = "0,00" Then Exit Sub
   If txtDescAV.Text = "" Then txtDescAV.Text = Format(0, ocMONEY)
   If txtAcrescAV.Text = "" Then txtAcrescAV.Text = Format(0, ocMONEY)
   
   If txtDescAV.Text <> "0,00" And txtAcrescAV.Text = "0,00" Then
      If optDescRSAV.Value = True Then
         txtTotalDescAV.Text = Format(CCur(txtSubTotalAV.Text) - CCur(txtDescAV.Text), ocMONEY)
      ElseIf optDescPorcAV.Value = True Then
         txtTotalDescAV.Text = Format(CCur(txtSubTotalAV.Text) - ((CCur(txtSubTotalAV.Text) * CCur(txtDescAV.Text)) / 100), ocMONEY)
      End If
      
   ElseIf txtAcrescAV.Text <> "0,00" And txtDescAV.Text = "0,00" Then
      If optAcrescRSAV.Value = True Then
         txtTotalDescAV.Text = Format(CCur(txtSubTotalAV.Text) + CCur(txtAcrescAV.Text), ocMONEY)
      ElseIf optAcrescPorcAV.Value = True Then
         txtTotalDescAV.Text = Format(CCur(txtSubTotalAV.Text) + ((CCur(txtSubTotalAV.Text) * CCur(txtAcrescAV.Text)) / 100), ocMONEY)
      End If
      
   Else
      txtTotalDescAV.Text = Format(txtSubTotalAV.Text, ocMONEY)
      'optDescRS.Value = True
      'optAscrescRS.Value = True
   End If
End Sub

Private Sub Calcular_Troco()
   Dim VAR_GERAL As Currency, VAR_RECEBIDO As Currency, var_Troco As Currency
   
   If txtTotalGeral.Text = "" Or txtRecebido.Text = "" Then Exit Sub
   
   If txtRecebido.Text = "0,00" Or txtRecebido.Text = "" Then
      txtTroco.Text = Format(var_Troco, ocMONEY)
   Else
      VAR_GERAL = txtTotalGeral.Text
      VAR_RECEBIDO = txtRecebido.Text
      var_Troco = VAR_RECEBIDO - VAR_GERAL
      txtTroco.Text = Format(var_Troco, ocMONEY)
   End If
End Sub

Private Sub Calcular_Parcelas()
   Dim var_ValorRest As Currency
   Dim QUANT As Integer
   Dim RESULTADO As Currency
   
   If txtTotalDesc.Text = "0,00" Or txtValorRest.Text = "0,00" Or cboQuantParc.Text = "" Then Exit Sub
   
   var_ValorRest = txtValorRest.Text
   QUANT = cboQuantParc.Text
   RESULTADO = CCur(var_ValorRest / QUANT)
   txtValorParc = Format(RESULTADO, ocMONEY)
End Sub

Public Function aNumeros(ByVal KeyAscii As Integer, Optional Virgula As Boolean = False, Optional Ponto As Boolean = False) As Integer
   'FUNÇÃO PARA PERMITIR NUMEROS, VIRGULAS E PONTO
   Select Case KeyAscii
      Case IIf(Virgula = True, 44, 0), IIf(Ponto = True, 46, 0), 8, 13, 48 To 57
         aNumeros = KeyAscii
      Case Else
         aNumeros = 0
   End Select
End Function

Private Sub txtSubTotal_Change()
   txtDesc_Change
End Sub

Private Sub txtSubTotal_GotFocus()
   SelectControl txtSubtotal
End Sub

Private Sub txtSubTotal_LostFocus()
   txtSubtotal = FormatCurrency(txtSubtotal)
End Sub

Function ChecarLimite() As Boolean
   Dim Limite As Currency
   Dim Total As Currency
   Dim LimiteAtual As Currency
   
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT * FROM cliente WHERE (codigo = " & txtCodCliente.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then Limite = r("limite_credito")
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   If Limite = 0 Then
      ChecarLimite = True
      Exit Function
   End If
   
   Total = 0
   sSQL = "SELECT OS_Carro_Cadastro.cod_cliente, SUM(OS_Carro_Cadastro.total) AS total FROM parcelas INNER JOIN OS_Carro_Cadastro ON parcelas.codigo = OS_Carro_Cadastro.codigo WHERE (OS_Carro_Cadastro.cod_cliente = " & txtCodCliente.Text & ") AND (parcelas.status = 0) GROUP BY OS_Carro_Cadastro.cod_cliente;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then Total = r("total")
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   LimiteAtual = Limite - Total
   
   If Left(LimiteAtual, 1) = "-" Then
      LimiteAtual = Mid(LimiteAtual, 2, Len(LimiteAtual))
   End If
   
   If LimiteAtual < (CCur(txtTotalGeral.Text) - CCur(txtEntrada.Text)) Then
      ShowMsg "O CLIENTE POSSUE UM TOTAL DE R$ " & FormatNumber(Total, 2) & " EM COMPRAS NÃO PAGAS E O VALOR DA COMPRA É DE R$ " & FormatNumber(txtTotalGeral.Text, 2) & " E O SALDO DELE É R$ " & FormatNumber(Limite - Total), vbExclamation
      ChecarLimite = False
   Else
      ChecarLimite = True
   End If
End Function

Private Sub Form_Unload(Cancel As Integer)
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtCodOS.Text = "" Then Exit Sub
   
   sSQL = "SELECT * FROM OS_Carro_Cadastro WHERE (cod_os = " & txtCodOS.Text & ") AND (status IS NULL);"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then
      If ShowMsg("Essa OS encontra-se em aberto. Deseja sair e cancelar essa OS?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         Cancel = 1
         Exit Sub
      End If
      
      dbData.Execute "DELETE FROM os_servicos WHERE (codigo = " & txtCodOS.Text & ");"
      dbData.Execute "DELETE FROM os_acessorios WHERE (codigo = " & txtCodOS.Text & ");"
      dbData.Execute "DELETE FROM OS_Carro_Cadastro WHERE (codigo = " & txtCodOS.Text & ");"
   End If
   
   Set moCombo = Nothing
End Sub

Private Sub lblTotal_Change()
   txtTotalServicos.Text = Format(lblTotal.Caption, "##,##0.00")
   Somar_Totais
End Sub

Private Sub lblTotalPeca_Change()
   txtTotalPecas.Text = Format(lblTotalPeca.Caption, "##,##0.00")
   Somar_Totais
End Sub

Private Sub mskDataEntrada_GotFocus()
   SelectControl mskDataEntrada
End Sub

Private Sub mskDataEntrada_KeyPress(KeyAscii As Integer)
   mskDataEntrada.Mask = "##/##/##"
End Sub

Private Sub mskDataEntrada_LostFocus()
   If mskDataEntrada.Text = "" Or mskDataEntrada.Text = "__/__/__" Then
      mskDataEntrada.Mask = ""
      mskDataEntrada.Text = ""
   Else
      If Not IsDate(mskDataEntrada.Text) Then
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskDataEntrada.SetFocus
      End If
   End If
End Sub

Private Sub mskHoraEntrada_GotFocus()
   SelectControl mskHoraEntrada
End Sub

Private Sub mskHoraEntrada_KeyPress(KeyAscii As Integer)
   mskHoraEntrada.Mask = "##:##"
End Sub

Private Sub mskHoraEntrada_LostFocus()
   If mskHoraEntrada.Text = "" Or mskHoraEntrada.Text = "__:__" Then
      mskHoraEntrada.Mask = ""
      mskHoraEntrada.Text = ""
Else
    If Not IsDate(mskHoraEntrada.Text) Then
        MsgBox "HORA INVÁLIDA!" & vbCrLf & "A hora digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation, "Aviso do Sistema"
        mskHoraEntrada.SetFocus
    End If
End If
End Sub

Private Sub mskValorPeca_GotFocus()
   SelectControl mskValorPeca
End Sub

Private Sub mskValorPeca_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cmdAdicionarPecas_Click
End Sub

Private Sub mskValorServico_GotFocus()
   SelectControl mskValorServico
End Sub

Private Sub mskValorServico_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmdAdicionarServicos_Click
   End If
End Sub

Private Sub mskValorServico_LostFocus()
   If mskValorServico.Text = "" Then
      mskValorServico.Text = FormatCurrency(0)
   Else
      mskValorServico.Text = FormatCurrency(mskValorServico)
   End If
End Sub

Private Sub optLocalizarCliente_Click()
   If optLocalizarCliente.Value = 1 Then optLocalizarCodOS.Value = 0
   cboLocalizar.SetFocus
End Sub

Public Sub optLocalizarCodOS_Click()
   If optLocalizarCodOS.Value = 1 Then optLocalizarCliente.Value = 0
   txtLocalizarCodOS.SetFocus
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab = 0 Then
      If cmdGerarEntrada.Enabled = True Then mskDataEntrada.SetFocus
   ElseIf SSTab1.Tab = 1 Then
      If frmServico.Enabled = True Then cboServicos.SetFocus
   ElseIf SSTab1.Tab = 2 Then
      If frmPecas.Enabled = True Then txtPecas.SetFocus
   ElseIf SSTab1.Tab = 3 Then
'      cboStatus.SetFocus
   ElseIf SSTab1.Tab = 4 Then
'      optAV.SetFocus
   ElseIf SSTab1.Tab = 5 Then
'      optStatusTodos.SetFocus
   End If
End Sub

Private Sub TxtCodCliente_Change()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtCodCliente.Text = "" Then Exit Sub
   
   If cmdAlterar.Enabled = True Then
      sSQL = "SELECT * FROM cliente WHERE (codigo = " & txtCodCliente.Text & ");"
      Set r = dbData.OpenRecordset(sSQL)
      If Not r.BOF Then cboCliente.Text = r("nome")
      If r.State <> 0 Then r.Close
      Set r = Nothing
   End If
End Sub

Private Sub txtCodFuncionario_Change()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtCodFuncionario.Text = "" Then Exit Sub
   If txtCodFuncionario.Text = 0 Then Exit Sub
   
   txtCodFuncAV.Text = txtCodFuncionario.Text
   txtCodFuncAP.Text = txtCodFuncionario.Text
   
   If cmdAlterar.Enabled = True Then
      sSQL = "SELECT * FROM funcionario WHERE (codigo = " & txtCodFuncionario.Text & ");"
      Set r = dbData.OpenRecordset(sSQL)
      If Not r.BOF Then cboFuncionario.Text = r("nome")
      If r.State <> 0 Then r.Close
      Set r = Nothing
   End If
End Sub

Private Sub txtCodMecanico_Change()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtCodMecanico.Text = "" Then Exit Sub
   
   If cmdAlterar.Enabled = True Then
      sSQL = "SELECT * FROM funcionario WHERE (codigo = " & txtCodMecanico.Text & ");"
      Set r = dbData.OpenRecordset(sSQL)
      If Not r.BOF Then cboMecanico.Text = r("nome")
      If r.State <> 0 Then r.Close
      Set r = Nothing
   End If
End Sub

Private Sub txtCodOS_Change()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   'If txtCodOS.Text <> "" Then cmdImprimirEntrada.Enabled = True Else cmdImprimirEntrada.Enabled = False
   
   If txtCodOS.Text = "" Then
      'imgCancelar.Enabled = False
      cmdGerarEntrada.Enabled = False
      LimparGrid_Acessorios
      cmdFinalizarAV.Visible = False
      cmdFinalizarAP.Visible = False
      lblCarro1a.Caption = ""
      lblCarro1b.Caption = ""
      lblCarro2a.Caption = ""
      lblCarro2b.Caption = ""
      Exit Sub
   Else
      'imgCancelar.Enabled = True
      cmdGerarEntrada.Enabled = True
   End If
   
   LimparObjetos_Entrada
   
   sSQL = "SELECT * FROM OS_Carro_Cadastro WHERE (cod_os = " & txtCodOS.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   Mostrar_Entrada r
   MostrarGrid_Acessorios
   MostrarGrid_Servicos
   MostrarGrid_Pecas
   Somar_Totais
   
   frmSecundario.Enabled = True
   cmdApagar.Enabled = True
   cmdAlterar.Enabled = True
   cmdImprimirEntrada.Enabled = True
   cmdGarantia.Enabled = True
   cmdGerarEntrada.Enabled = False
   cmdNovo.Enabled = True
   
   'CHECAR SE A OS ESTÁ FECHADA & PAGA
   Verificar_OS_FechadaePaga
   
   If OS_FECHADAPAGA = True Then
      If cboStatus.Text = "TERMINADO" Then
         cmdFinalizarAV.Visible = True
         cmdFinalizarAP.Visible = True
      End If
   Else
      cmdFinalizarAV.Visible = False
      cmdFinalizarAP.Visible = False
   End If
   
   If frmPrincipal.Enabled = True Then cboStatus.SetFocus
End Sub

Private Sub txtCodServico_Change()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtCodServico.Text = "" Then Exit Sub
   
   sSQL = "SELECT * FROM servicos WHERE (codigo = " & txtCodServico.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then mskValorServico.Text = Format(r("valor"), ocMONEY)
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPecas_GotFocus()
   SelectControl txtPecas
End Sub

Private Sub txtQuantPeca_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtQuantPeca.Text = "" Then txtQuantPeca.Text = 1
      cmdAdicionarPecas_Click
   End If
End Sub

Private Sub txtQuantPeca_LostFocus()
   If txtQuantPeca.Text = "" And txtPecas.Text <> "" Then txtQuantPeca.Text = 1
   If txtQuantPeca.Text = "" Or mskValorPeca.Text = "" Then txtTotalPeca.Text = "0,00": Exit Sub
   txtTotalPeca.Text = Format(txtQuantPeca.Text * CDbl(mskValorPeca.Text), "##,##0.00")
End Sub

Private Sub txtQuantServico_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If txtQuantServico.Text = "" Then txtQuantServico.Text = 1
      cmdAdicionarServicos_Click
   End If
End Sub

Private Sub txtQuantServico_LostFocus()
   If txtQuantServico.Text = "" And cboServicos <> "" Then txtQuantServico.Text = 1
End Sub

Private Sub txtSubTotalAV_Change()
   txtDescAV_Change
End Sub

Private Sub txtTotalDesc_Change()
   txtTotalGeral.Text = Format(txtTotalDesc.Text, "##,##0.00")
End Sub

Private Sub txtTotalDescAV_Change()
   txtTotalGeral.Text = Format(txtTotalDescAV.Text, "##,##0.00")
End Sub

Private Sub txtValorParc_GotFocus()
   If txtTotalGeral.Text = "" Then
      Exit Sub
   Else
      Mostrar_ValorRestante
   End If
   
   SelectControl txtValorParc
End Sub

Private Sub txtValorParc_LostFocus()
   If txtValorParc = "" Then
      txtValorParc = Format(0, ocMONEY)
   Else
      txtValorParc = Format(txtValorParc, ocMONEY)
   End If
End Sub

Private Sub txtValorRest_Change()
   Calcular_Parcelas
End Sub
