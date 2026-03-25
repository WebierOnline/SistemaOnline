VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "ChamaleonBtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL32.OCX"
Begin VB.Form Contas_Cadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CONTAS À PAGAR"
   ClientHeight    =   9915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13320
   Icon            =   "Contas_Cadastro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9915
   ScaleWidth      =   13320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   107
      Top             =   9420
      Width           =   195
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00800080&
      Height          =   195
      Left            =   120
      TabIndex        =   106
      Top             =   9180
      Width           =   195
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   105
      Top             =   8940
      Width           =   195
   End
   Begin VB.TextBox txtTotalQuant 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4920
      TabIndex        =   58
      Top             =   9060
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   60
      ScaleHeight     =   765
      ScaleWidth      =   13125
      TabIndex        =   70
      Top             =   60
      Width           =   13155
      Begin VB.Timer Timer1 
         Interval        =   600
         Left            =   9540
         Top             =   180
      End
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8400
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lblHora 
         AutoSize        =   -1  'True
         Caption         =   "00:00"
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
         Left            =   10200
         TabIndex        =   74
         Top             =   300
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONTAS À PAGAR"
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
         Left            =   1515
         TabIndex        =   71
         Top             =   180
         Width           =   2850
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   480
         Picture         =   "Contas_Cadastro.frx":23D2
         Top             =   60
         Width           =   645
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdFechar 
      Height          =   675
      Left            =   11460
      TabIndex        =   14
      Top             =   8940
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   1191
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
      MICON           =   "Contas_Cadastro.frx":8A68
      PICN            =   "Contas_Cadastro.frx":8A84
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
      Left            =   9660
      TabIndex        =   13
      Top             =   8940
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   1191
      BTYPE           =   3
      TX              =   "&Imprimir"
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
      MICON           =   "Contas_Cadastro.frx":8D9E
      PICN            =   "Contas_Cadastro.frx":8DBA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7995
      Left            =   60
      TabIndex        =   15
      Top             =   900
      Width           =   13155
      _ExtentX        =   23204
      _ExtentY        =   14102
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
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
      TabCaption(0)   =   "CADASTRO"
      TabPicture(0)   =   "Contas_Cadastro.frx":90D4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdNovo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdSalvar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdExcluir"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdAlterar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdCancelar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Picture4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "frmCadastro"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "CONSULTA"
      TabPicture(1)   =   "Contas_Cadastro.frx":90F0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtCONhaveres"
      Tab(1).Control(1)=   "txtCONsubtotal"
      Tab(1).Control(2)=   "txtCONtotal"
      Tab(1).Control(3)=   "txtCONquant"
      Tab(1).Control(4)=   "Picture3"
      Tab(1).Control(5)=   "Grid"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "HAVER"
      TabPicture(2)   =   "Contas_Cadastro.frx":910C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frmHaver"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "HISTÓRICO"
      TabPicture(3)   =   "Contas_Cadastro.frx":9128
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture2"
      Tab(3).ControlCount=   1
      Begin VB.TextBox txtCONhaveres 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -65940
         Locked          =   -1  'True
         TabIndex        =   76
         ToolTipText     =   "HAVERES"
         Top             =   7620
         Width           =   1995
      End
      Begin VB.TextBox txtCONsubtotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67980
         Locked          =   -1  'True
         TabIndex        =   75
         ToolTipText     =   "SUBTOTAIS"
         Top             =   7620
         Width           =   1995
      End
      Begin VB.Frame frmCadastro 
         Enabled         =   0   'False
         Height          =   6135
         Left            =   120
         TabIndex        =   69
         Top             =   420
         Width           =   9495
         Begin VB.Frame frmPrincipal 
            Caption         =   "Cadastro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   60
            TabIndex        =   85
            Top             =   180
            Width           =   9375
            Begin ChamaleonBtn.chameleonButton cmdCal1 
               Height          =   315
               Left            =   1200
               TabIndex        =   118
               Tag             =   "Calendario"
               Top             =   1140
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
               MICON           =   "Contas_Cadastro.frx":9144
               PICN            =   "Contas_Cadastro.frx":9160
               PICH            =   "Contas_Cadastro.frx":B4B3
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.ComboBox cboForma 
               Height          =   315
               Left            =   3900
               Sorted          =   -1  'True
               TabIndex        =   7
               Top             =   1140
               Width           =   1875
            End
            Begin VB.ComboBox cboStatus 
               BackColor       =   &H00C0FFFF&
               Height          =   315
               ItemData        =   "Contas_Cadastro.frx":D806
               Left            =   120
               List            =   "Contas_Cadastro.frx":D808
               TabIndex        =   0
               TabStop         =   0   'False
               Top             =   480
               Width           =   1095
            End
            Begin VB.TextBox txtDuplic 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1560
               TabIndex        =   5
               Text            =   "1"
               Top             =   1140
               Width           =   615
            End
            Begin VB.ComboBox cboFavorecido 
               Height          =   315
               Left            =   5820
               TabIndex        =   8
               Top             =   1140
               Width           =   3435
            End
            Begin VB.ComboBox cboDescricao 
               Height          =   315
               Left            =   1260
               TabIndex        =   1
               Top             =   480
               Width           =   4515
            End
            Begin VB.ComboBox cboTipo 
               Height          =   315
               Left            =   2220
               Sorted          =   -1  'True
               TabIndex        =   6
               Top             =   1140
               Width           =   1695
            End
            Begin VB.ComboBox cboSetor 
               Height          =   315
               Left            =   7380
               TabIndex        =   3
               Top             =   480
               Width           =   1875
            End
            Begin VB.TextBox txtRef 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   5820
               TabIndex        =   2
               Top             =   480
               Width           =   1515
            End
            Begin VB.TextBox txtVlrUnit 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               TabIndex        =   9
               Top             =   1800
               Width           =   1515
            End
            Begin MSMask.MaskEdBox mskVenc 
               Height          =   315
               Left            =   120
               TabIndex        =   4
               Top             =   1140
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskCadastro 
               Height          =   315
               Left            =   8160
               TabIndex        =   10
               Top             =   1800
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cadastro"
               Height          =   195
               Left            =   8160
               TabIndex        =   112
               Top             =   1560
               Width           =   630
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Forma"
               Height          =   195
               Left            =   3900
               TabIndex        =   95
               Top             =   900
               Width           =   435
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Descrição"
               Height          =   195
               Left            =   1260
               TabIndex        =   94
               Top             =   240
               Width           =   720
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Venc."
               Height          =   195
               Left            =   120
               TabIndex        =   93
               Top             =   900
               Width           =   420
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Status"
               Height          =   195
               Left            =   120
               TabIndex        =   92
               Top             =   240
               Width           =   450
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Repetir"
               Height          =   195
               Left            =   1560
               TabIndex        =   91
               Top             =   900
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Favorecido"
               Height          =   195
               Left            =   5820
               TabIndex        =   90
               Top             =   900
               Width           =   795
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo"
               Height          =   195
               Left            =   2220
               TabIndex        =   89
               Top             =   900
               Width           =   315
            End
            Begin VB.Label Setor 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Setor"
               Height          =   195
               Left            =   7380
               TabIndex        =   88
               Top             =   240
               Width           =   375
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ref.:"
               Height          =   195
               Left            =   5820
               TabIndex        =   87
               Top             =   240
               Width           =   345
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Left            =   120
               TabIndex        =   86
               Top             =   1560
               Width           =   360
            End
         End
         Begin VB.Frame frmPagamento 
            Caption         =   "Pagamento"
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
            Left            =   60
            TabIndex        =   82
            Top             =   2460
            Width           =   9375
            Begin ChamaleonBtn.chameleonButton chameleonButton1 
               Height          =   315
               Left            =   3060
               TabIndex        =   119
               Tag             =   "Calendario"
               Top             =   480
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
               MICON           =   "Contas_Cadastro.frx":D80A
               PICN            =   "Contas_Cadastro.frx":D826
               PICH            =   "Contas_Cadastro.frx":FB79
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.ComboBox cboFonte 
               BackColor       =   &H00C0FFFF&
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               Sorted          =   -1  'True
               TabIndex        =   11
               Top             =   480
               Width           =   1935
            End
            Begin MSMask.MaskEdBox mskPgto 
               Height          =   315
               Left            =   2100
               TabIndex        =   12
               Top             =   480
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   12648447
               ForeColor       =   192
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fonte"
               Enabled         =   0   'False
               Height          =   195
               Left            =   120
               TabIndex        =   84
               Top             =   240
               Width           =   405
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Pgto."
               Enabled         =   0   'False
               Height          =   195
               Left            =   2100
               TabIndex        =   83
               Top             =   240
               Width           =   375
            End
         End
         Begin ChamaleonBtn.chameleonButton cmdHabilitarQuitar 
            Height          =   435
            Left            =   1980
            TabIndex        =   102
            Top             =   4440
            Visible         =   0   'False
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   767
            BTYPE           =   3
            TX              =   "QUITAR A CONTA"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   -2147483633
            BCOLO           =   -2147483633
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Contas_Cadastro.frx":11ECC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdHabilitarHaver 
            Height          =   435
            Left            =   4740
            TabIndex        =   103
            Top             =   4440
            Visible         =   0   'False
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   767
            BTYPE           =   3
            TX              =   "HAVER NA CONTA"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   12632256
            BCOLO           =   12632256
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Contas_Cadastro.frx":11EE8
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
         Height          =   7395
         Left            =   -74880
         ScaleHeight     =   7335
         ScaleWidth      =   11235
         TabIndex        =   61
         Top             =   420
         Width           =   11295
         Begin MSFlexGridLib.MSFlexGrid Grid_Historico 
            Height          =   6915
            Left            =   60
            TabIndex        =   100
            Top             =   60
            Width           =   11115
            _ExtentX        =   19606
            _ExtentY        =   12197
            _Version        =   393216
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin VB.TextBox txtQuantHist 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   60
            TabIndex        =   62
            Top             =   7020
            Width           =   795
         End
         Begin VB.Label lblTotalGridHistorico 
            Alignment       =   1  'Right Justify
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
            Left            =   11100
            TabIndex        =   104
            Top             =   7020
            Width           =   75
         End
      End
      Begin VB.Frame frmHaver 
         Caption         =   "Haver"
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
         ForeColor       =   &H00000000&
         Height          =   7395
         Left            =   -74880
         TabIndex        =   45
         Top             =   420
         Width           =   5775
         Begin VB.TextBox txtCodHaver 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3060
            TabIndex        =   46
            Top             =   180
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.PictureBox Picture8 
            Height          =   6255
            Left            =   120
            ScaleHeight     =   6195
            ScaleWidth      =   5475
            TabIndex        =   57
            Top             =   1020
            Width           =   5535
            Begin MSFlexGridLib.MSFlexGrid Grid_Haver 
               Height          =   5775
               Left            =   120
               TabIndex        =   52
               Top             =   120
               Width           =   5295
               _ExtentX        =   9340
               _ExtentY        =   10186
               _Version        =   393216
               SelectionMode   =   1
            End
            Begin VB.Label lblTotalGridHaver 
               Alignment       =   1  'Right Justify
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
               Left            =   5280
               TabIndex        =   101
               Top             =   5880
               Width           =   75
            End
         End
         Begin VB.PictureBox Picture7 
            Height          =   735
            Left            =   120
            ScaleHeight     =   675
            ScaleWidth      =   5475
            TabIndex        =   47
            Top             =   240
            Width           =   5535
            Begin VB.TextBox txtValorHaver 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   1020
               TabIndex        =   49
               Top             =   300
               Width           =   915
            End
            Begin VB.ComboBox cboFonteHaver 
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   1980
               Sorted          =   -1  'True
               TabIndex        =   50
               Top             =   300
               Width           =   1695
            End
            Begin MSMask.MaskEdBox mskDataHaver 
               Height          =   315
               Left            =   60
               TabIndex        =   48
               Top             =   300
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   12648447
               PromptChar      =   "_"
            End
            Begin ChamaleonBtn.chameleonButton cmdAdicionarHaver 
               Height          =   315
               Left            =   3720
               TabIndex        =   51
               Top             =   300
               Width           =   795
               _ExtentX        =   1402
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
               MICON           =   "Contas_Cadastro.frx":11F04
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdRemoverHaver 
               Height          =   315
               Left            =   4560
               TabIndex        =   53
               Top             =   300
               Width           =   795
               _ExtentX        =   1402
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
               MICON           =   "Contas_Cadastro.frx":11F20
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor:"
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
               Left            =   1020
               TabIndex        =   56
               Top             =   60
               Width           =   510
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Data:"
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
               TabIndex        =   55
               Top             =   60
               Width           =   480
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fonte"
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
               Left            =   1980
               TabIndex        =   54
               Top             =   60
               Width           =   495
            End
         End
      End
      Begin VB.TextBox txtCONtotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -63900
         Locked          =   -1  'True
         TabIndex        =   44
         ToolTipText     =   "TOTAIS"
         Top             =   7620
         Width           =   1995
      End
      Begin VB.TextBox txtCONquant 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74920
         TabIndex        =   43
         Top             =   7620
         Width           =   1575
      End
      Begin VB.PictureBox Picture3 
         Height          =   1515
         Left            =   -74940
         ScaleHeight     =   1455
         ScaleWidth      =   11295
         TabIndex        =   23
         Top             =   420
         Width           =   11355
         Begin VB.Frame frmData 
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
            Height          =   1335
            Left            =   3600
            TabIndex        =   64
            Top             =   60
            Width           =   1275
            Begin VB.OptionButton optVencimento 
               Caption         =   "Vencimento"
               Height          =   195
               Left            =   60
               TabIndex        =   68
               Top             =   300
               Width           =   1155
            End
            Begin VB.OptionButton optPagamento 
               Caption         =   "Pagamento"
               Height          =   195
               Left            =   60
               TabIndex        =   67
               Top             =   540
               Value           =   -1  'True
               Width           =   1155
            End
            Begin VB.CheckBox chkPgtoOutros 
               Caption         =   "Outros..."
               Height          =   195
               Left            =   240
               TabIndex        =   66
               Top             =   1020
               Width           =   975
            End
            Begin VB.CheckBox chkPgtoMes 
               Caption         =   "Mês"
               Height          =   195
               Left            =   240
               TabIndex        =   65
               Top             =   780
               Width           =   855
            End
         End
         Begin VB.Frame Frame3 
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
            Height          =   1335
            Left            =   4920
            TabIndex        =   42
            Top             =   60
            Width           =   1335
            Begin VB.OptionButton optOrdCadastro 
               Caption         =   "&Cadastro"
               Height          =   195
               Left            =   60
               TabIndex        =   63
               Top             =   450
               Width           =   1155
            End
            Begin VB.OptionButton optOrdVencimento 
               Caption         =   "&Vencimento"
               Height          =   195
               Left            =   60
               TabIndex        =   99
               Top             =   240
               Value           =   -1  'True
               Width           =   1155
            End
            Begin VB.OptionButton optOrdFavorecido 
               Caption         =   "Favor&ecido"
               Height          =   195
               Left            =   60
               TabIndex        =   98
               Top             =   660
               Width           =   1095
            End
            Begin VB.OptionButton optOrdReferente 
               Caption         =   "&Descrição"
               Height          =   195
               Left            =   60
               TabIndex        =   97
               Top             =   870
               Width           =   1095
            End
            Begin VB.OptionButton optOrdForma 
               Caption         =   "F&orma"
               Height          =   195
               Left            =   60
               TabIndex        =   96
               Top             =   1080
               Width           =   915
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Tipo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   60
            TabIndex        =   41
            Top             =   60
            Width           =   1275
            Begin VB.CheckBox optNome 
               Caption         =   "Favorecido"
               Height          =   195
               Left            =   120
               TabIndex        =   80
               Top             =   1020
               Width           =   1095
            End
            Begin VB.CheckBox optIntervalo 
               Caption         =   "Intervalo"
               Height          =   195
               Left            =   120
               TabIndex        =   79
               Top             =   780
               Width           =   1035
            End
            Begin VB.CheckBox optMensal 
               Caption         =   "Mês"
               Height          =   195
               Left            =   120
               TabIndex        =   78
               Top             =   540
               Width           =   1035
            End
            Begin VB.CheckBox optTodos 
               Caption         =   "Todos"
               Height          =   195
               Left            =   120
               TabIndex        =   77
               Top             =   300
               Width           =   1035
            End
         End
         Begin VB.Frame Frame4 
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
            Height          =   1335
            Left            =   1380
            TabIndex        =   34
            Top             =   60
            Width           =   2175
            Begin VB.ComboBox cboCONStatus 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
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
               Left            =   660
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   240
               Width           =   1395
            End
            Begin VB.ComboBox cboCONsetor 
               Height          =   315
               Left            =   660
               Sorted          =   -1  'True
               TabIndex        =   36
               Top             =   600
               Width           =   1395
            End
            Begin VB.ComboBox cboCONForma 
               Height          =   315
               Left            =   660
               TabIndex        =   35
               Top             =   960
               Width           =   1395
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Setor:"
               Height          =   195
               Left            =   120
               TabIndex        =   40
               Top             =   645
               Width           =   420
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Status:"
               Height          =   195
               Left            =   120
               TabIndex        =   39
               Top             =   300
               Width           =   495
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Forma:"
               Height          =   195
               Left            =   120
               TabIndex        =   38
               Top             =   1005
               Width           =   480
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Sub - Critérios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   6300
            TabIndex        =   24
            Top             =   60
            Width           =   4935
            Begin VB.ComboBox cboMES 
               Height          =   315
               ItemData        =   "Contas_Cadastro.frx":11F3C
               Left            =   1260
               List            =   "Contas_Cadastro.frx":11F3E
               TabIndex        =   27
               Top             =   240
               Visible         =   0   'False
               Width           =   2115
            End
            Begin VB.ComboBox cboNome 
               Height          =   315
               Left            =   1260
               TabIndex        =   26
               Top             =   600
               Visible         =   0   'False
               Width           =   3555
            End
            Begin VB.ComboBox cboAno 
               Height          =   315
               Left            =   3480
               Sorted          =   -1  'True
               TabIndex        =   25
               Top             =   240
               Visible         =   0   'False
               Width           =   1335
            End
            Begin MSMask.MaskEdBox Mask2 
               Height          =   315
               Left            =   3300
               TabIndex        =   28
               Top             =   240
               Visible         =   0   'False
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Mask1 
               Height          =   315
               Left            =   1260
               TabIndex        =   29
               Top             =   240
               Visible         =   0   'False
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin ChamaleonBtn.chameleonButton cmdExibir 
               Height          =   315
               Left            =   1920
               TabIndex        =   81
               Top             =   960
               Visible         =   0   'False
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "&EXIBIR"
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
               MICON           =   "Contas_Cadastro.frx":11F40
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label lblCONmes 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "E&scolha o mês:"
               Height          =   195
               Left            =   120
               TabIndex        =   33
               Top             =   300
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.Label lblCONint1 
               AutoSize        =   -1  'True
               Caption         =   "Da&ta Inicial:"
               Height          =   195
               Left            =   360
               TabIndex        =   32
               Top             =   285
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.Label lblCONint2 
               AutoSize        =   -1  'True
               Caption         =   "Data &Final:"
               Height          =   195
               Left            =   2460
               TabIndex        =   31
               Top             =   285
               Visible         =   0   'False
               Width           =   765
            End
            Begin VB.Label lblCONnome 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Favorecido:"
               Height          =   195
               Left            =   300
               TabIndex        =   30
               Top             =   660
               Visible         =   0   'False
               Width           =   840
            End
         End
      End
      Begin VB.PictureBox Picture4 
         Height          =   1215
         Left            =   6840
         ScaleHeight     =   1155
         ScaleWidth      =   2655
         TabIndex        =   16
         Top             =   6600
         Width           =   2715
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL:"
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
            Top             =   780
            Width           =   675
         End
         Begin VB.Label lblTotalGeral 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
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
            Height          =   315
            Left            =   1380
            TabIndex        =   21
            Top             =   780
            Width           =   1215
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HAVER:"
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
            TabIndex        =   20
            Top             =   420
            Width           =   705
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SUB-TOTAL:"
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
            TabIndex        =   19
            Top             =   60
            Width           =   1110
         End
         Begin VB.Label lblTotalHaver 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1380
            TabIndex        =   18
            Top             =   420
            Width           =   1215
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1380
            TabIndex        =   17
            Top             =   60
            Width           =   1215
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   5595
         Left            =   -74940
         TabIndex        =   72
         Top             =   1980
         Width           =   13035
         _ExtentX        =   22992
         _ExtentY        =   9869
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelar 
         Height          =   615
         Left            =   9720
         TabIndex        =   113
         Top             =   1800
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
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
         MICON           =   "Contas_Cadastro.frx":11F5C
         PICN            =   "Contas_Cadastro.frx":11F78
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
         Left            =   9720
         TabIndex        =   114
         Top             =   2460
         Width           =   2175
         _ExtentX        =   3836
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
         MICON           =   "Contas_Cadastro.frx":13D0A
         PICN            =   "Contas_Cadastro.frx":13D26
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdExcluir 
         Height          =   615
         Left            =   9720
         TabIndex        =   115
         Top             =   3120
         Width           =   2175
         _ExtentX        =   3836
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
         MICON           =   "Contas_Cadastro.frx":15AB8
         PICN            =   "Contas_Cadastro.frx":15AD4
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
         TabIndex        =   116
         Top             =   1140
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
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
         MICON           =   "Contas_Cadastro.frx":17866
         PICN            =   "Contas_Cadastro.frx":17882
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
         TabIndex        =   117
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Novo"
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
         MICON           =   "Contas_Cadastro.frx":19614
         PICN            =   "Contas_Cadastro.frx":19630
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label29 
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
         Left            =   -68520
         TabIndex        =   60
         Top             =   7560
         Width           =   510
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quant.:"
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
         TabIndex        =   59
         Top             =   7560
         Width           =   645
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   111
      Top             =   9645
      Width           =   13320
      _ExtentX        =   23495
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   16589
            Text            =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
            TextSave        =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "12:51"
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
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "À vencer"
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
      Left            =   360
      TabIndex        =   110
      Top             =   9420
      Width           =   780
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencendo hoje"
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
      Left            =   360
      TabIndex        =   109
      Top             =   9180
      Width           =   1290
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencidas"
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
      Left            =   360
      TabIndex        =   108
      Top             =   8940
      Width           =   795
   End
End
Attribute VB_Name = "Contas_Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mes As Integer
Dim Ano As Integer
Dim Data As String
Dim i As Integer
Dim x As Long
Dim DIA1 As Integer
Dim DIA As Integer
Dim PARCELA As Currency
Dim Y As Long

Private moCombo As cComboHelper
Private printSQL As String

Private Function Atualizar_Dados_Caixa_Haver(ByVal Cod As Long, ByVal Acao As Integer) As Boolean
   'A atualização deve ser feita utilizando o comando UPDATE do sql
   'e não mais usando o método .Update do Recordset
   
   'Não se deve comparar se o campo está vazio ou não, pois dessa forma não
   'haverá atualização quando for necessário apagar alguma informação
   
   Dim sSQL As String
   
   If Acao = 1 Then        'Insere novo
      sSQL = "INSERT INTO caixa_saida (codigo, descricao, subdescricao, data, valor, setor, cod_haver, hora) VALUES (" & _
         Cod & ", 'CONTA/HAVER DE: " & cboFavorecido.Text & "', 'PGTO DE CONTA', CONVERT(DATETIME, '" & Format$(CDate(mskDataHaver), ocDATA) & "', 103), " & _
         Replace(CCur(txtValorHaver.Text), ",", ".") & ", '" & cboSetor.Text & "', " & txtCodHaver.Text & ", '" & lblHora.Caption & "');"
      
   ElseIf Acao = 2 Then    'Atualiza
      'Comando de atualização
      sSQL = "UPDATE caixa_saida SET " & _
         "descricao = 'CONTA/HAVER DE: " & cboFavorecido.Text & "', " & _
         "subdescricao = 'PGTO DE CONTA', " & _
         "data = CONVERT(DATETIME, '" & Format$(CDate(mskDataHaver), ocDATA) & "', 103), " & _
         "valor = " & Replace(CCur(txtValorHaver.Text), ",", ".") & ", " & _
         "setor = '" & cboSetor.Text & "', " & _
         "cod_haver = " & IIf(txtCodHaver.Text = "", Null, txtCodHaver.Text) & ", " & _
         "hora = '" & lblHora.Caption & "' "
      
      'Condição para atualização
      sSQL = sSQL & "WHERE (codigo = " & Cod & ");"
   End If
   
   'Retorna o resultado da atualização
   Atualizar_Dados_Caixa_Haver = dbData.Execute(sSQL)
End Function

Private Function Atualizar_Dados() As Boolean
   'A atualização deve ser feita utilizando o comando UPDATE do sql
   'e não mais usando o método .Update do Recordset
   
   'Não se deve comparar se o campo está vazio ou não, pois dessa forma não
   'haverá atualização quando for necessário apagar alguma informação
   
   Dim sSQL As String
   
   'Comando de atualização
   sSQL = "UPDATE a_pagar SET " & _
      "cadastro = CONVERT(DATETIME, '" & Format(CDate(mskCadastro), ocDATA) & "', 103), " & _
      "vencimento = CONVERT(DATETIME, '" & Format(CDate(mskVenc), ocDATA) & "', 103), " & _
      "total = " & Replace(CCur(txtVlrUnit), ",", ".") & ", " & _
      "status = '" & cboStatus.Text & "', " & _
      "descricao = '" & cboDescricao.Text & "', " & _
      "favorecido = '" & cboFavorecido.Text & "', " & _
      "setor = '" & cboSetor.Text & "', " & _
      "tipo = '" & cboTipo.Text & "', " & _
      "forma = '" & cboForma.Text & "', " & _
      "fonte = '" & cboFonte.Text & "', " & _
      "ref = '" & txtRef.Text & "'"
   
   If IsDate(mskPgto) Then
      sSQL = sSQL & ", data_pgto = CONVERT(DATETIME, '" & Format(CDate(mskPgto), ocDATA) & "', 103)"
   End If
   
   'Condição para atualização
   sSQL = sSQL & " WHERE (codigo = " & txtCodigo.Text & ");"
   
   'Retorna o resultado da atualização
   Atualizar_Dados = dbData.Execute(sSQL)
End Function

Private Function Auto_Numeracao_APagar() As Long
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim lRet As Long
   
   lRet = 1
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultima_conta FROM a_pagar;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then lRet = r("ultima_conta") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   Auto_Numeracao_APagar = lRet
End Function

Private Sub FormatarGrid_Contas(rTabela As ADODB.Recordset)
   Dim i As Integer, j As Integer
   
   If cboCONStatus.Text = "PAGA" Then
      With Grid
         .Clear
         .Cols = 11
         .Rows = 2
         
         .ColWidth(0) = 0
         .ColWidth(1) = 0
         .ColWidth(2) = 850
         .ColWidth(3) = 2150
         .ColWidth(4) = 2200
         .ColWidth(5) = 700
         .ColWidth(6) = 1000
         .ColWidth(7) = 850
         .ColWidth(8) = 870
         .ColWidth(9) = 870
         .ColWidth(10) = 870
         
         .TextMatrix(0, 1) = "COD"
         .TextMatrix(0, 2) = "VENC."
         .TextMatrix(0, 3) = "FAVORECIDO"
         .TextMatrix(0, 4) = "DESCRIÇÃO"
         .TextMatrix(0, 5) = "REF."
         .TextMatrix(0, 6) = "FONTE"
         .TextMatrix(0, 7) = "PGTO"
         .TextMatrix(0, 8) = "VALOR"
         .TextMatrix(0, 9) = "HAVER"
         .TextMatrix(0, 10) = "TOTAL"
         .Redraw = False
         
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
         
         If Not rTabela Is Nothing Then
            Do While Not rTabela.EOF
               'mudar a cor da coluna
               'For i = 1 To .Rows - 1
               '.Row = i
               '.Col = 6:   .CellBackColor = &HC0FFFF
               ' Next
               
               'ALINHAMENTO
               '.ColAlignment(2) = 1
               
               .TextMatrix(.Rows - 1, 1) = ValidateNull(rTabela("codigo"))
               .TextMatrix(.Rows - 1, 2) = Format(ValidateNull(rTabela("vencimento")), "dd/mm/yy")
               .TextMatrix(.Rows - 1, 3) = ValidateNull(rTabela("favorecido"))
               .TextMatrix(.Rows - 1, 4) = ValidateNull(rTabela("descricao"))
               .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("ref"))
               .TextMatrix(.Rows - 1, 6) = ValidateNull(rTabela("fonte"))
               .TextMatrix(.Rows - 1, 7) = Format(ValidateNull(rTabela("data_pgto")), "dd/mm/yy")
               .TextMatrix(.Rows - 1, 8) = Format(ValidateNull(rTabela("total")), ocMONEY)
               .TextMatrix(.Rows - 1, 9) = Format(ValidateNull(rTabela("hav")), ocMONEY)
               .TextMatrix(.Rows - 1, 10) = Format(ValidateNull(rTabela("total")), ocMONEY)
               
               rTabela.MoveNext
               .Rows = .Rows + 1
            Loop
         End If
         
         .Rows = .Rows - 1
         
         'MUDAR COR DE FONTE DA COLUNA
         For i = 1 To .Rows - 1
            .Row = i
            .Col = 10
            .CellForeColor = &HC0&
            .CellFontBold = True
         Next
         
         .Redraw = True
      End With
   
   Else
      With Grid
         .Clear
         .Cols = 11
         .Rows = 2
         
         .ColWidth(0) = 0
         .ColWidth(1) = 0
         .ColWidth(2) = 0
         .ColWidth(3) = 850
         .ColWidth(4) = 2650
         .ColWidth(5) = 2700
         .ColWidth(6) = 1000
         .ColWidth(7) = 850
         .ColWidth(8) = 950
         .ColWidth(9) = 900
         .ColWidth(10) = 1050
         
         .TextMatrix(0, 1) = "COD"
         '.TextMatrix(0, 2) = "CAD."
         .TextMatrix(0, 3) = "VENC."
         .TextMatrix(0, 4) = "FAVORECIDO"
         .TextMatrix(0, 5) = "DESCRIÇÃO"
         .TextMatrix(0, 6) = "FORMA"
         .TextMatrix(0, 7) = "REF."
         .TextMatrix(0, 8) = "VALOR"
         .TextMatrix(0, 9) = "HAVER"
         .TextMatrix(0, 10) = "LIQUIDO"
         
         .Redraw = False
         
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
         
         If Not rTabela Is Nothing Then
            Do While Not rTabela.EOF
               'ALINHAMENTO
               '.ColAlignment(2) = 1
               
               .TextMatrix(.Rows - 1, 1) = ValidateNull(rTabela("codigo"))
               .TextMatrix(.Rows - 1, 3) = Format(ValidateNull(rTabela("vencimento")), "dd/mm/yy")
               '.TextMatrix(.Rows - 1, 2) = Format(RS!CADASTRO, "dd/mm/yy")
               .TextMatrix(.Rows - 1, 4) = ValidateNull(rTabela("favorecido"))
               .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("descricao"))
               .TextMatrix(.Rows - 1, 6) = ValidateNull(rTabela("forma"))
               .TextMatrix(.Rows - 1, 7) = ValidateNull(rTabela("ref"))
               .TextMatrix(.Rows - 1, 8) = Format(ValidateNull(rTabela("VALOR_UND")), ocMONEY)
               .TextMatrix(.Rows - 1, 9) = Format(ValidateNull(rTabela("hav")), ocMONEY)
               .TextMatrix(.Rows - 1, 10) = Format(ValidateNull(rTabela("VALOR_UND")) - ValidateNull(rTabela("hav")), ocMONEY)
               
               rTabela.MoveNext
               .Rows = .Rows + 1
            Loop
         End If
         
         .Rows = .Rows - 1
         
         'Deixar negrito quando vencido
         For i = 1 To .Rows - 1
            For j = 0 To .Cols - 1
               .Col = j
               .Row = i
               
               If CDate(.TextMatrix(i, 3)) < Date Then
                  .CellForeColor = vbRed
               ElseIf CDate(.TextMatrix(i, 3)) = Date Then
                  .CellForeColor = &H800080
               ElseIf CDate(.TextMatrix(i, 3)) > Date Then
                  .CellForeColor = vbBlack
               End If
            Next
         Next
         
         'MUDAR COR DE FONTE DA COLUNA
         For i = 1 To .Rows - 1
            .Row = i
            .Col = 10
            .CellForeColor = &HC0&
            .CellFontBold = True
         Next
         
         'MUDAR COR DE FONTE DA COLUNA
         For i = 1 To .Rows - 1
            .Row = i
            .Col = 6
            .CellForeColor = vbBlue
            .CellFontBold = True
         Next
         
         .Redraw = True
      End With
   End If
   
   'SOMAR REGISTROS
   If cboCONStatus.Text = "À PAGAR" Then
      txtCONsubtotal.Text = Format(SomaGrid(Grid, 8), ocMONEY)
      txtCONhaveres.Text = Format(SomaGrid(Grid, 9), ocMONEY)
      txtCONtotal.Text = Format(SomaGrid(Grid, 10), ocMONEY)
   ElseIf cboCONStatus.Text = "PAGA" Then
      txtCONsubtotal.Text = Format(SomaGrid(Grid, 8), ocMONEY)
      txtCONhaveres.Text = Format(SomaGrid(Grid, 9), ocMONEY)
      txtCONtotal.Text = Format(SomaGrid(Grid, 10), ocMONEY)
   End If
End Sub

Private Sub FormatarGrid_Haver(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid_Haver
      .Clear
      .Cols = 5
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 1200
      .ColWidth(3) = 1200
      .ColWidth(4) = 2000
      
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "DATA"
      .TextMatrix(0, 3) = "VALOR"
      .TextMatrix(0, 4) = "FONTE"
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = rTabela("codigo")
            .TextMatrix(.Rows - 1, 2) = Format(rTabela("data"), "dd/mm/yy")
            .TextMatrix(.Rows - 1, 3) = Format(rTabela("valor"), ocMONEY)
            .TextMatrix(.Rows - 1, 4) = rTabela("fonte")
            
            rTabela.MoveNext
            .Rows = .Rows + 1
            i = i + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 3
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
   
   lblTotalGridHaver.Caption = Format(SomaGrid(Grid_Haver, 3), ocMONEY)
End Sub

Private Sub FormatarGrid_Historico(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid_Historico
      .Clear
      .Cols = 12
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 850
      .ColWidth(3) = 2200
      .ColWidth(4) = 2400
      .ColWidth(5) = 700
      .ColWidth(6) = 1000
      .ColWidth(7) = 850
      .ColWidth(8) = 1000
      .ColWidth(9) = 1100
      .ColWidth(10) = 1000
      .ColWidth(11) = 1000
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "VENC."
      .TextMatrix(0, 3) = "FAVORECIDO"
      .TextMatrix(0, 4) = "DESCRIÇÃO"
      .TextMatrix(0, 5) = "REF."
      .TextMatrix(0, 6) = "FONTE"
      .TextMatrix(0, 7) = "PGTO"
      .TextMatrix(0, 8) = "VALOR"
      .TextMatrix(0, 9) = "HAVER(ES)"
      .TextMatrix(0, 10) = "TOTAL"
      .TextMatrix(0, 11) = "TOTAL"
      .Redraw = False
      
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
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            'mudar a cor da coluna
            'For i = 1 To .Rows - 1
            '.Row = i
            '.Col = 6:   .CellBackColor = &HC0FFFF
            ' Next
            
            'ALINHAMENTO
            '.ColAlignment(2) = 1
            
            .TextMatrix(.Rows - 1, 1) = rTabela("codigo")
            .TextMatrix(.Rows - 1, 2) = Format(rTabela("vencimento"), "dd/mm/yy")
            .TextMatrix(.Rows - 1, 3) = rTabela("favorecido")
            .TextMatrix(.Rows - 1, 4) = rTabela("descricao")
            .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("ref"))
            .TextMatrix(.Rows - 1, 6) = ValidateNull(rTabela("fonte"))
            .TextMatrix(.Rows - 1, 7) = Format(rTabela("data_pgto"), "dd/mm/yy")
            .TextMatrix(.Rows - 1, 8) = Format(rTabela("total"), ocMONEY)
            '.TextMatrix(.Rows - 1, 9) = Format(RS!HAV, "##,##0.00")
            '.TextMatrix(.Rows - 1, 10) = Format(RS!Total, "##,##0.00")
            
            rTabela.MoveNext
            .Rows = .Rows + 1
         Loop
      End If
      
      .Rows = .Rows - 1
      .Redraw = True
      
      'SOMAR REGUSTROS
      lblTotalGridHistorico.Caption = Format(SomaGrid(Grid, 9), ocMONEY)
    End With
End Sub

Private Sub LimparGrid_Consulta()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT * FROM a_pagar WHERE false;"
   Set r = dbData.OpenRecordset(sSQL)
   MostrarGrid_Contas
   If Not r.State <> 0 Then r.Close
   Set r = Nothing
   
   txtCONquant.Text = ""
   txtCONtotal.Text = ""
End Sub

Private Sub LimparGrid_Haver()
   Dim i As Integer
   
   With Grid_Haver
      .Clear
      .Cols = 4
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 1200
      .ColWidth(2) = 1200
      .ColWidth(3) = 1200
      
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "DATA"
      .TextMatrix(0, 3) = "VALOR"
      
      .Redraw = False
      .Rows = .Rows + 1
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 3
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
   
   lblTotalGridHaver.Caption = Format(SomaGrid(Grid_Haver, 3), ocMONEY)
End Sub

Private Sub LimparGrid_Historico()
   Dim i As Integer
   
   With Grid_Historico
      .Clear
      .Cols = 12
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 850
      .ColWidth(3) = 2200
      .ColWidth(4) = 2400
      .ColWidth(5) = 700
      .ColWidth(6) = 1000
      .ColWidth(7) = 850
      .ColWidth(8) = 1000
      .ColWidth(9) = 1100
      .ColWidth(10) = 1000
      .ColWidth(11) = 1000
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "VENC."
      .TextMatrix(0, 3) = "FAVORECIDO"
      .TextMatrix(0, 4) = "DESCRIÇÃO"
      .TextMatrix(0, 5) = "REF."
      .TextMatrix(0, 6) = "FONTE"
      .TextMatrix(0, 7) = "PGTO"
      .TextMatrix(0, 8) = "VALOR"
      .TextMatrix(0, 9) = "HAVER(ES)"
      .TextMatrix(0, 10) = "TOTAL"
      .TextMatrix(0, 11) = "TOTAL"
      
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
      .Rows = .Rows + 1
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
   
   'SOMAR REGUSTROS
   lblTotalGridHistorico.Caption = Format(SomaGrid(Grid, 9), ocMONEY)
End Sub

Private Sub Limpar_Objetos_Conta()
   If cmdAlterar.Visible = False Then txtCodigo.Text = ""
   mskVenc.Mask = ""
   mskVenc.Text = ""
   mskCadastro.Mask = ""
   mskCadastro.Text = ""
   mskPgto.Mask = ""
   mskPgto.Text = ""
   txtVlrUnit.Text = ""
   cboDescricao.Text = ""
   cboFavorecido.Text = ""
   cboTipo.Text = ""
   cboForma.Text = ""
   cboFonte.Text = ""
   cboSetor.Text = ""
   txtDuplic.Text = 1
   lblTotal.Caption = ""
   lblTotalHaver.Caption = ""
   lblTotalGeral.Caption = ""
   cboStatus.Text = "À PAGAR"
   txtRef.Text = ""
End Sub

Private Sub LimparGrid_Conta()
   Dim i As Integer
   
   With Grid
      .Clear
      .Cols = 10
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 850
      .ColWidth(3) = 2650
      .ColWidth(4) = 2770
      .ColWidth(5) = 700
      .ColWidth(6) = 1000
      .ColWidth(7) = 1000
      .ColWidth(8) = 1100
      .ColWidth(9) = 1000
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "VENC."
      .TextMatrix(0, 3) = "FAVORECIDO"
      .TextMatrix(0, 4) = "DESCRIÇÃO"
      .TextMatrix(0, 5) = "REF."
      .TextMatrix(0, 6) = "FORMA"
      .TextMatrix(0, 7) = "VALOR"
      .TextMatrix(0, 8) = "HAVER(ES)"
      .TextMatrix(0, 9) = "TOTAL"
      
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
      .Rows = .Rows + 1
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
   
   txtCONsubtotal.Text = Format(0, ocMONEY)
   txtCONhaveres.Text = Format(0, ocMONEY)
   txtCONtotal.Text = Format(0, ocMONEY)
   
   optTodos.Value = Unchecked
   optMensal.Value = Unchecked
   optIntervalo.Value = Unchecked
   optNome.Value = Unchecked
End Sub

Private Sub MostrarGrid_Contas()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim totalRegistros As Long
   
   Dim INDICE As String
   Dim Tipo_Data As String
   Dim var_Setor As String
   Dim var_Forma As String
   
   If optORDvencimento.Value = True Then
      INDICE = "vencimento"
   ElseIf optOrdCadastro.Value = True Then
      INDICE = "cadastro"
   ElseIf optOrdFavorecido.Value = True Then
      INDICE = "favorecido"
   ElseIf optOrdReferente.Value = True Then
      INDICE = "descricao"
   ElseIf optOrdForma.Value = True Then
      INDICE = "forma"
   Else
      optORDvencimento.Value = True
      INDICE = "vencimento"
   End If

   If cboCONsetor.Text <> "TODOS" Then
      var_Setor = "AND (setor = '" & cboCONsetor.Text & "') "
   Else
      var_Setor = ""
   End If

   If cboCONForma.Text <> "TODOS" Then
      var_Forma = "AND (forma = '" & cboCONForma.Text & "') "
   Else
      var_Forma = ""
   End If
    
   If optTodos.Value = Checked Then
      sSQL = "SELECT *, valor_und, ISNULL(haveres, 0) AS hav, (valor_und - ISNULL(haveres, 0)) as var_Liquido FROM a_pagar " & _
         "WHERE (status = '" & cboCONStatus.Text & "') " & var_Forma & var_Setor & " ORDER BY " & INDICE
      
   ElseIf optMensal.Value = Checked And optNome.Value = Unchecked Then
      If cboMES.Text = "" Or cboAno.Text = "" Then Exit Sub
      
      'INDICE SOMENTE PARA CONSULTA DE DATAS
      'If cboCONStatus.Text = "PAGA" Then
      '    If optVencimento.Value = True Then
      '        INDICE = "VENCIMENTO"
      '   ElseIf optPagamento.Value = True Then
      '        INDICE = "DATA_PGTO"
      '    End If
      'Else
      '    INDICE = "VENCIMENTO"
      'End If
      
      'DATA DE VENCIMENTO OU PAGAMENTO
      If cboCONStatus.Text = "PAGA" Then
         If optVencimento.Value = True Then
            Tipo_Data = "(MONTH(vencimento) = " & cboMES.ListIndex + 1 & ") AND (YEAR(vencimento) = " & cboAno & ") "
         ElseIf optPagamento.Value = True And chkPgtoOutros.Value = 0 And chkPgtoMes.Value = 0 Then
            Tipo_Data = "(MONTH(data_pgto) = " & cboMES.ListIndex + 1 & ") AND (YEAR(data_pgto) = " & cboAno & ") "
         ElseIf optPagamento.Value = True And chkPgtoOutros.Value = 1 And chkPgtoMes.Value = 0 Then
            Tipo_Data = "(MONTH(data_pgto) = " & cboMES.ListIndex + 1 & ") AND (YEAR(data_pgto) = " & cboAno & ") AND (MONTH(vencimento) <> " & cboMES.ListIndex + 1 & ") "
         ElseIf optPagamento.Value = True And chkPgtoOutros.Value = 0 And chkPgtoMes.Value = 1 Then
            Tipo_Data = "(MONTH(data_pgto) = " & cboMES.ListIndex + 1 & ") AND (YEAR(data_pgto) = " & cboAno & ") AND (MONTH(vencimento) = " & cboMES.ListIndex + 1 & " AND (YEAR(vencimento) = " & cboAno & ") "
         Else
            Tipo_Data = "(MONTH(vencimento) = " & cboMES.ListIndex + 1 & ") AND (YEAR(vencimento) = " & cboAno & ") "
         End If
      Else
        Tipo_Data = "(MONTH(vencimento) = " & cboMES.ListIndex + 1 & ") AND (YEAR(vencimento) = " & cboAno & ") "
      End If
      'AND (MONTH(" & DATAS & ") = " & cboMES.ListIndex + 1 & ") AND (YEAR(" & DATAS & ") = " & cboAno & ")
    
    'MOSTRAR OS DADOS
      sSQL = "SELECT *, total, ISNULL(haveres, 0) AS hav, (total - ISNULL(haveres, 0)) as var_Liquido FROM a_pagar " & _
         "WHERE " & Tipo_Data & " AND (STATUS = '" & cboCONStatus.Text & "') " & var_Forma & var_Setor & " ORDER BY " & INDICE
      
   ElseIf optIntervalo.Value = Checked And optNome.Value = Unchecked Then
      If Mask1.Text = "" Or Mask2.Text = "" Then Exit Sub
      If Not IsDate(Mask1.Text) Or Not IsDate(Mask2.Text) Then Exit Sub
      
      If Mask1.Text = "" Or Mask2.Text = "" Then
         ShowMsg "Digite a DATA INICIAL e DATA FINAL!", vbExclamation
         Mask1.SetFocus
         Exit Sub
      End If
      
      'MOSTRAR OS DADOS
      sSQL = "SELECT *, total, ISNULL(haveres, 0) AS hav, (total - ISNULL(haveres, 0)) as var_Liquido FROM a_pagar " & _
         "WHERE (STATUS = '" & cboCONStatus.Text & "') AND (vencimento >= CONVERT(DATETIME, '" & Format(Mask1, ocDATA) & "', 103)) AND (vencimento <= CONVERT(DATETIME, '" & Format(Mask2, ocDATA) & "', 103)) " & _
         var_Forma & var_Setor & " ORDER BY " & INDICE
      
   ElseIf optNome.Value = Checked And optIntervalo.Value = Unchecked And optMensal.Value = Unchecked Then
      If cboNome.Text = "" Then Exit Sub
      
      'MOSTRAR OS DADOS
      sSQL = "SELECT *, total, ISNULL(haveres, 0) AS hav, (total - ISNULL(haveres, 0)) as var_Liquido FROM a_pagar " & _
         "WHERE (status = '" & cboCONStatus.Text & "') AND (favorecido = '" & cboNome.Text & "') " & var_Forma & var_Setor & " ORDER BY " & INDICE
      
   ElseIf optNome.Value = Checked And optIntervalo.Value = Checked And optMensal.Value = Unchecked Then
      If cboNome.Text = "" Then Exit Sub
      If Mask1.Text = "" Or Mask2.Text = "" Then Exit Sub
      
      'MOSTRAR OS DADOS
      sSQL = "SELECT *, total, ISNULL(haveres, 0) AS hav, (total - ISNULL(haveres, 0)) as var_Liquido FROM a_pagar " & _
         "WHERE (status = '" & cboCONStatus.Text & "') AND (favorecido = '" & cboNome.Text & "') " & _
         "AND (vencimento >= CONVERT(DATETIME, '" & Format(Mask1, ocDATA) & "', 103)) " & _
         "AND (vencimento <= CONVERT(DATETIME, '" & Format(Mask2, ocDATA) & "', 103)) " & _
         var_Forma & var_Setor & " ORDER BY " & INDICE
      
   ElseIf optNome.Value = Checked And optIntervalo.Value = Unchecked And optMensal.Value = Checked Then
      If cboNome.Text = "" Then Exit Sub
      If cboMES.Text = "" Or cboAno.Text = "" Then Exit Sub
      
      'INDICE SOMENTE PARA CONSULTA DE DATAS
      If optVencimento.Value = True Then
         INDICE = "vencimento"
      ElseIf optPagamento.Value = True Then
         INDICE = "data_pgto"
      End If
      
      'DATA DE VENCIMENTO OU PAGAMENTO
      If optVencimento.Value = True Then
         Tipo_Data = "(MONTH(vencimento) = " & cboMES.ListIndex + 1 & ") AND (YEAR(vencimento) = " & cboAno & ") "
      ElseIf optPagamento.Value = True And chkPgtoOutros.Value = 0 And chkPgtoMes.Value = 0 Then
         Tipo_Data = "(MONTH(data_pgto) = " & cboMES.ListIndex + 1 & ") AND (YEAR(data_pgto) = " & cboAno & ") "
      ElseIf optPagamento.Value = True And chkPgtoOutros.Value = 1 And chkPgtoMes.Value = 0 Then
         Tipo_Data = "(MONTH(data_pgto) = " & cboMES.ListIndex + 1 & ") AND (YEAR(data_pgto) = " & cboAno & ") AND (MONTH(vencimento) <> " & cboMES.ListIndex + 1 & ") "
      ElseIf optPagamento.Value = True And chkPgtoOutros.Value = 0 And chkPgtoMes.Value = 1 Then
         Tipo_Data = "(MONTH(data_pgto) = " & cboMES.ListIndex + 1 & ") AND (YEAR(data_pgto) = " & cboAno & ") AND (MONTH(vencimento) = " & cboMES.ListIndex + 1 & " AND (YEAR(vencimento) = " & cboAno & ") "
      Else
         Tipo_Data = "(MONTH(vencimento) = " & cboMES.ListIndex + 1 & " AND (YEAR(vencimento) = " & cboAno & ") "
      End If
      
      'MOSTRAR OS DADOS
      sSQL = "SELECT *, total, ISNULL(haveres, 0) AS hav, (total - ISNULL(haveres, 0)) as var_Liquido FROM a_pagar " & _
         "WHERE " & Tipo_Data & " AND (status = '" & cboCONStatus.Text & "') AND (favorecido = '" & cboNome.Text & "') " & _
         var_Forma & var_Setor & " ORDER BY " & INDICE
      
   Else
      sSQL = "SELECT *, total, ISNULL(haveres, 0) AS hav, (total - ISNULL(haveres, 0)) as var_Liquido FROM a_pagar WHERE 0 = 1;"
   
   End If
   
   Set r = dbData.OpenRecordset(sSQL, totalRegistros)
   FormatarGrid_Contas r
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   printSQL = sSQL
   txtCONquant.Text = Format(totalRegistros, "00")
End Sub

Private Sub MostrarGrid_Historico()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim totalRegistros As Long
   
   If cboFavorecido.Text = "" Then Exit Sub
   
   'GRID - HISTÓRICO
   sSQL = "SELECT * FROM a_pagar WHERE (favorecido = '" & cboFavorecido.Text & "') ORDER BY vencimento;"
   Set r = dbData.OpenRecordset(sSQL, totalRegistros)
   
   'QUANTIDADE - HISTÓRICO
   txtQuantHist.Text = totalRegistros
   FormatarGrid_Historico r
   
   If Not r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub Ocultar_SubCriterios()
   If optTodos.Value = Checked Then
      lblCONmes.Visible = False
      cboMES.Visible = False
      cboAno.Visible = False
      
      lblCONint1.Visible = False
      Mask1.Visible = False
      lblCONint2.Visible = False
      Mask2.Visible = False
      
      lblCONnome.Visible = False
      cboNome.Visible = False
      
      optVencimento.Enabled = False
      optPagamento.Enabled = False
      chkPgtoMes.Enabled = False
      chkPgtoOutros.Enabled = False
   End If
   
   If optMensal.Value = Checked Then
      If optNome.Value = Checked Then
         lblCONnome.Visible = True
         cboNome.Visible = True
      Else
         lblCONnome.Visible = False
         cboNome.Visible = False
      End If
      
      lblCONint1.Visible = False
      Mask1.Visible = False
      lblCONint2.Visible = False
      Mask2.Visible = False
      
      lblCONmes.Visible = True
      cboMES.Visible = True
      cboAno.Visible = True
      
      'If cboCONStatus.Text = "PAGA" Then
         optVencimento.Enabled = True
         optPagamento.Enabled = True
         chkPgtoMes.Enabled = True
         chkPgtoOutros.Enabled = True
      'Else
      '   optVencimento.Enabled = False
      '   optPagamento.Enabled = False
      '   chkPgtoMes.Enabled = False
      '   chkPgtoOutros.Enabled = False
      'End If
   
   Else
      lblCONmes.Visible = False
      cboMES.Visible = False
      cboAno.Visible = False
   End If
   
   If optIntervalo.Value = Checked Then
      lblCONmes.Visible = False
      cboMES.Visible = False
      cboAno.Visible = False
      
      lblCONint1.Visible = True
      Mask1.Visible = True
      lblCONint2.Visible = True
      Mask2.Visible = True
      
      If optNome.Value = Checked Then
         lblCONnome.Visible = True
         cboNome.Visible = True
      Else
         lblCONnome.Visible = False
         cboNome.Visible = False
      End If
      
      optVencimento.Enabled = False
      optPagamento.Enabled = False
      chkPgtoMes.Enabled = False
      chkPgtoOutros.Enabled = False
   Else
      lblCONint1.Visible = False
      Mask1.Visible = False
      lblCONint2.Visible = False
      Mask2.Visible = False
   End If
   
   If optNome.Value = Checked Then
      If optMensal.Value = Checked Then
         lblCONmes.Visible = True
         cboMES.Visible = True
         cboAno.Visible = True
      Else
         lblCONmes.Visible = False
         cboMES.Visible = False
         cboAno.Visible = False
      End If
      
      If optIntervalo.Value = Checked Then
         lblCONint1.Visible = True
         Mask1.Visible = True
         lblCONint2.Visible = True
         Mask2.Visible = True
      Else
         lblCONint1.Visible = False
         Mask1.Visible = False
         lblCONint2.Visible = False
         Mask2.Visible = False
      End If
      
      lblCONnome.Visible = True
      cboNome.Visible = True
      
      'optVencimento.Enabled = False
      'optPagamento.Enabled = False
      'chkPgtoMes.Enabled = False
      'chkPgtoOutros.Enabled = False
   Else
      lblCONnome.Visible = False
      cboNome.Visible = False
   End If
   
   If optMensal.Value = Checked Or optIntervalo.Value = Checked Or optNome.Value = Checked Then
      cmdExibir.Visible = True
   Else
      cmdExibir.Visible = False
   End If
End Sub

Public Function SomaGrid(Grid As MSFlexGrid, Col As Integer) As Currency
   Dim i As Integer, Valor As Currency
   
   Valor = 0
   For i = 0 To Grid.Rows - 1
      If IsNumeric(Grid.TextMatrix(i, Col)) Then
         Valor = Valor + CCur(Grid.TextMatrix(i, Col))
      End If
   Next
   
   SomaGrid = Valor
End Function

Private Sub cboAno_GotFocus()
   Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
   Dim i As Integer
   
   cboAno.Clear
   
   iAno = Year(Date)
   FirstYear = iAno - 2
   LastYear = iAno + 2
   
   For i = LastYear To FirstYear Step -1
      cboAno.AddItem i
   Next
   
   'For i = ANO To FirstYear Step -1
   '   cboAno.AddItem i
   'Next
   
   'iAno = iAno + 1
   'For i = iAno To LastYear
   '   cboAno.AddItem i
   'Next
End Sub

Private Sub cboAno_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then cmdExibir_Click
End Sub

Private Sub cboAno_LostFocus()
   'If cboAno.Text = "" Then Exit Sub Else cmdExibir.SetFocus
End Sub

Private Sub cboCONForma_Change()
   cboCONForma_Click
End Sub

Private Sub cboCONForma_Click()
   If optTodos.Value = Checked Then
      optTodos_Click
   ElseIf optMensal.Value = Checked Then
      cmdExibir_Click
   ElseIf optIntervalo.Value = Checked Then
      cmdExibir_Click
   ElseIf optNome.Value = Checked Then
      cmdExibir_Click
   End If
End Sub

Private Sub cboCONForma_GotFocus()
Dim varTexto As String
varTexto = cboCONForma.Text
   
   cboCONForma.Clear
   cboCONForma.AddItem "TODOS"
   cboCONForma.AddItem "CHEQUE"
   cboCONForma.AddItem "BOLETO"
   cboCONForma.AddItem "CARNÊ"
   cboCONForma.AddItem "PROMISSORIA"
   cboCONForma.AddItem "DUPLICATA"
   cboCONForma.AddItem "DÉBITO EM CONTA"
   cboCONForma.AddItem "AVULSA"
   moCombo.AttachTo cboCONForma
   
cboCONForma.Text = varTexto
End Sub

Private Sub cboCONsetor_Change()
   cboCONsetor_Click
End Sub

Private Sub cboCONsetor_Click()
   If optTodos.Value = Checked Then
      optTodos_Click
   ElseIf optMensal.Value = Checked Then
      cmdExibir_Click
   ElseIf optIntervalo.Value = Checked Then
      cmdExibir_Click
   ElseIf optNome.Value = Checked Then
      cmdExibir_Click
   End If
End Sub

Private Sub cboCONsetor_GotFocus()
Dim varTexto As String
varTexto = cboCONsetor.Text

   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboCONsetor.Clear
  
   sSQL = "SELECT DISTINCT setor FROM setor ORDER BY setor;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboCONsetor.AddItem r("setor")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboCONsetor
   
cboCONsetor.Text = varTexto
End Sub

Private Sub cboCONStatus_Change()
   Ocultar_SubCriterios
   cboCONStatus_Click
End Sub

Private Sub cboCONStatus_Click()
   If optTodos.Value = Checked Then
      optTodos_Click
   ElseIf optMensal.Value = Checked Then
      cmdExibir_Click
   ElseIf optIntervalo.Value = Checked Then
      cmdExibir_Click
   ElseIf optNome.Value = Checked Then
      cmdExibir_Click
   End If
   
   Ocultar_SubCriterios
End Sub

Private Sub cboCONStatus_GotFocus()
Dim varTexto As String
varTexto = cboCONStatus.Text

   cboCONStatus.Clear
   cboCONStatus.AddItem "À PAGAR"
   cboCONStatus.AddItem "PAGA"
   cboCONStatus.AddItem "VENCIDAS"

cboCONStatus.Text = varTexto
End Sub

Private Sub cboFavorecido_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset

   Dim VarText As String
   VarText = cboFavorecido.Text
   
   cboFavorecido.Clear
   
   sSQL = "SELECT favorecido FROM a_pagar GROUP BY favorecido ORDER BY favorecido;"
   Set r = dbData.OpenRecordset(sSQL)
   
   If r.BOF Then
      cboFavorecido.AddItem "Nada"
   Else
      Do While Not r.EOF
         cboFavorecido.AddItem r("favorecido")
         r.MoveNext
      Loop
   End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing

   cboFavorecido.Text = VarText
   
   moCombo.AttachTo cboFavorecido
End Sub

Private Sub cboFavorecido_KeyPress(KeyAscii As Integer)
   If Len(cboFavorecido) = 40 Then txtVlrUnit.SetFocus
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboFonte_GotFocus()
   cboFonte.Clear
   cboFonte.AddItem "CAIXA ATUAL"
   cboFonte.AddItem "SALDOS"
   cboFonte.AddItem "NENHUMA"
   moCombo.AttachTo cboFonte
End Sub

Private Sub cboFonteHaver_GotFocus()
   cboFonteHaver.Clear
   cboFonteHaver.AddItem "CAIXA ATUAL"
   cboFonteHaver.AddItem "SALDOS"
   cboFonteHaver.AddItem "NENHUMA"
   moCombo.AttachTo cboFonteHaver
End Sub

Private Sub cboFonteHaver_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      cmdAdicionarHaver_Click
   End If
End Sub

Private Sub cboFonteHaver_LostFocus()
   'cmdAdicionarHaver.SetFocus
End Sub

Private Sub cboForma_GotFocus()
   Dim VarText As String
   VarText = cboForma.Text
   
   cboForma.Clear
   cboForma.AddItem "CHEQUE"
   cboForma.AddItem "BOLETO"
   cboForma.AddItem "CARNÊ"
   cboForma.AddItem "PROMISSORIA"
   cboForma.AddItem "DUPLICATA"
   cboForma.AddItem "DÉBITO EM CONTA"
   cboForma.AddItem "AVULSA"

   cboForma.Text = VarText
   
   moCombo.AttachTo cboForma
End Sub

Private Sub cboMes_GotFocus()
   Dim vMes As Integer
   
   cboMES.Clear
   
   For vMes = 1 To 12
      cboMES.AddItem StrConv(MonthName(vMes), vbProperCase)
   Next
   
   moCombo.AttachTo cboMES
End Sub

Private Sub cboMes_LostFocus()
   If cboMES.Text = "" Then Exit Sub Else cboAno.SetFocus
End Sub

Private Sub cboNome_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   cboNome.Clear
   
   sSQL = "SELECT DISTINCT favorecido FROM a_pagar ORDER BY favorecido;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboNome.AddItem r("favorecido")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   moCombo.AttachTo cboNome
End Sub

Private Sub cboNome_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cmdExibir_Click
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboNome_LostFocus()
   If cmdExibir.Value = True Then cmdExibir.SetFocus
End Sub

Private Sub cboSetor_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   Dim VarText As String
   VarText = cboSetor.Text
   
   cboSetor.Clear
   sSQL = "SELECT DISTINCT setor FROM setor ORDER BY setor;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboSetor.AddItem r("setor")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   cboSetor.Text = VarText
   moCombo.AttachTo cboSetor
End Sub

Private Sub cboStatus_Change()
   cboStatus_LostFocus
End Sub

Private Sub cboStatus_Click()
   cboStatus_LostFocus
End Sub

Private Sub cboStatus_GotFocus()
   cboStatus.Clear
   cboStatus.AddItem "À PAGAR"
   cboStatus.AddItem "PAGA"
   'cboStatus.ListIndex = 0
   
   SelectControl cboStatus
End Sub

Private Sub cboStatus_LostFocus()
   'If cboStatus.Text = "PAGA" Then
   '   Label10.Enabled = True
   '   Label11.Enabled = True
   '   mskPgto.Enabled = True
   '   mskPgto.Text = Format(Date, "DD/MM/YY")
   '   cboFonte.Enabled = True
   '   cboFonte.SetFocus
   'Else
   '   Label10.Enabled = False
   '   Label11.Enabled = False
   '   mskPgto.Enabled = False
   '   cboFonte.Enabled = False
   '   cboDescricao.SetFocus
   'End If
End Sub

Private Sub cboTipo_GotFocus()
   Dim VarText As String
   VarText = cboTipo.Text
   
   cboTipo.Clear
   cboTipo.AddItem "FIXA"
   cboTipo.AddItem "TEMPORÁRIA"
   moCombo.AttachTo cboTipo

   cboTipo.Text = VarText
End Sub

Private Sub chameleonButton1_Click()
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

mskPgto = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub chkPgtoMes_Click()
   If chkPgtoMes.Value = 1 Then chkPgtoOutros.Value = 0
   cmdExibir_Click
End Sub

Private Sub chkPgtoOutros_Click()
   If chkPgtoOutros.Value = 1 Then chkPgtoMes.Value = 0
   cmdExibir_Click
End Sub

Private Sub cmdAdicionarHaver_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim novoCodHaver As Long
   Dim novoCodCaixa As Long
   Dim novoCodSaldoRet  As Long
   
   Dim S_SAIDAS As Currency
   Dim S_ENTRADAS As Currency
   Dim S_PARCELAS As Currency
   Dim S_TOTAIS As Currency
   
   If txtCodigo.Text = "" Then Exit Sub
   
   Auto_Numeracao_Haver
   
   If txtCodHaver.Text = "" Or txtValorHaver.Text = "" Or mskDataHaver.Text = "" Then Exit Sub
   
   If cboFonteHaver.Text = "CAIXA ATUAL" Then
      'MOSTRAR SE O CAIXA ESTÁ FECHADO
      sSQL = "SELECT * FROM caixa_dia WHERE (data_abertura = CONVERT(DATETIME, '" & Format(mskDataHaver, ocDATA) & "', 103)) AND (status = 1);"
      Set r = dbData.OpenRecordset(sSQL)
      
      
      If Not r.BOF Then
         ShowMsg "ESTE CAIXA JÁ ESTÁ FECHADO!", vbExclamation
         Exit Sub
      End If
      
      'VER SE A SALDO NO CAIXA ATUAL
      sSQL = "SELECT ISNULL(SUM(valor), 0) AS soma_saida FROM caixa_saida WHERE (data = CONVERT(DATETIME, '" & Format(mskDataHaver, ocDATA) & "', 103));"
      Set r = dbData.OpenRecordset(sSQL)
      If Not r.BOF Then S_SAIDAS = r("soma_saida")
      If r.State <> 0 Then r.Close
      Set r = Nothing
      
      sSQL = "SELECT ISNULL(SUM(valor), 0) AS soma_entrada FROM caixa_entrada WHERE (data = CONVERT(DATETIME, '" & Format(mskDataHaver, ocDATA) & "', 103));"
      Set r = dbData.OpenRecordset(sSQL)
      If Not r.BOF Then S_ENTRADAS = r("soma_entrada")
      If r.State <> 0 Then r.Close
      Set r = Nothing
      
      sSQL = "SELECT ISNULL(SUM(valor_final), 0) AS soma_parcelas FROM parcelas WHERE (pagamento = CONVERT(DATETIME, '" & Format(mskDataHaver, ocDATA) & "', 103));"
      Set r = dbData.OpenRecordset(sSQL)
      If Not r.BOF Then S_PARCELAS = r("soma_parcelas")
      If r.State <> 0 Then r.Close
      Set r = Nothing
      
      S_TOTAIS = S_PARCELAS + S_ENTRADAS
      S_TOTAIS = S_TOTAIS - S_SAIDAS
      
      If S_TOTAIS < CCur(txtValorHaver) Then
         ShowMsg "O Saldo do CAIXA é menor que o valor da conta!", vbInformation
         Exit Sub
      End If
      
      '=====================================
      'ADICIONAR NA TABELA CAIXA_SAIDA
      novoCodCaixa = AutoNumeracao_Caixa_Saida
      
      'sSQL = "SELECT * FROM caixa_saida;"
      'Set r = dbData.OpenRecordset(sSQL)
      
      'RS.AddNew
      Atualizar_Dados_Caixa_Haver novoCodCaixa, 1
      'RS.Update
      '========================================
      
   ElseIf cboFonteHaver.Text = "SALDOS" Then
      'VER SE O SE O CAIXA 2 TEM SALDO SUFICIENTE
      Dim VALOR_CONTA As Currency
      Dim RESULTADO As Currency
      Dim VALOR_ATUAL As Currency
      Dim var_Saldo As Currency
      
      Dim RETIRADA_ATUAL As Currency
      
      VALOR_CONTA = CCur(txtValorHaver.Text)
      var_Saldo = 0
      RETIRADA_ATUAL = VALOR_CONTA
      
      'ver o saldo atual
      sSQL = "SELECT TOP 1 saldo_atual AS saldo FROM caixa_saldo ORDER BY data DESC;"
      Set r = dbData.OpenRecordset(sSQL)
      If Not r.BOF Then var_Saldo = r("saldo")
      If r.State <> 0 Then r.Close
      Set r = Nothing
      
      'ver o valor da retirada
      sSQL = "SELECT TOP 1 retirada AS retirada_caixa FROM caixa_saldo ORDER BY data DESC;"
      Set r = dbData.OpenRecordset(sSQL)
      If Not r.BOF Then RETIRADA_ATUAL = r("retirada_Caixa") + VALOR_CONTA
      If r.State <> 0 Then r.Close
      Set r = Nothing
      
      If var_Saldo < VALOR_CONTA Then
         ShowMsg "O Saldo do CAIXA 2 é menor que o valor da conta!", vbInformation
         Exit Sub
      Else
         VALOR_ATUAL = var_Saldo - VALOR_CONTA
      End If
      
      'ADICIONAR A RETIRADA NA TABELA CAIXA
      novoCodCaixa = AutoNumeracao_Caixa_Saldo
      
      dbData.Execute "UPDATE caixa_saldo SET retirada = " & Replace(CCur(RETIRADA_ATUAL), ",", ".") & ", " & _
         "saldo_atual = " & Replace(CCur(VALOR_ATUAL), ",", ".") & " WHERE (codigo = " & novoCodHaver & ");"
      
      'ADICIONAR NA TABELA CAIXA_SALDO_RETIRADA
      novoCodSaldoRet = AutoNumeracao_Saldo_Retirada
      
      sSQL = "SELECT * FROM caixa_saldo_retirada"
      Set r = dbData.OpenRecordset(sSQL)
      
      'RS.AddNew
      Atualizar_Dados_Saldo_Retirada_Haver novoCodSaldoRet, novoCodCaixa, 1
      'RS.Update
      '============================================
   End If
   
   Dim LAST_HAVER As Currency
   Dim ATUAL_HAVER As Currency
   Dim SOMA_HAVERES As Currency
   
   'ADICIONAR O HAVER NA TABELA HAVER
   sSQL = "SELECT * FROM a_pagar_haver;"
   Set r = dbData.OpenRecordset(sSQL)
   
   'RS.AddNew
   Atualizar_Dados_Haver 1
   'RS.Update
   
   LAST_HAVER = 0
   
   'ATUALIZAR CAMPO HAVERES NA TABELA A_PAGAR
   sSQL = "SELECT haveres FROM a_pagar WHERE (codigo = " & txtCodigo.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then LAST_HAVER = ValidateNull(r("haveres"))
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   ATUAL_HAVER = CCur(txtValorHaver.Text)
   SOMA_HAVERES = LAST_HAVER + ATUAL_HAVER
   
   dbData.Execute "UPDATE a_pagar SET haveres = " & Replace(CCur(SOMA_HAVERES), ",", ".") & " WHERE (codigo = " & txtCodigo.Text & ");"
   
   Calcular_Valores
   MostrarGrid_Contas
   LimparObjetos_Haver
   MostrarGrid_Haver
   txtValorHaver.SetFocus
End Sub

Private Sub cmdAlterar_Click()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If cboFonte.Text = "" And cboStatus.Text = "PAGA" Then MsgBox "Selecione uma fonte!", vbInformation, "Aviso do sistema": cboFonte.SetFocus: Exit Sub
   
   If Not IsDate(mskVenc.Text) Then
      ShowMsg "Data Inválida!", vbInformation
      mskVenc.SetFocus
      Exit Sub
   End If
   
   If cboStatus.Text = "" Or cboDescricao.Text = "" Or cboFavorecido.Text = "" Or txtVlrUnit.Text = "" Or cboTipo.Text = "" Then
      ShowMsg "Falta preencher alguns campos!", vbInformation
      cboStatus.SetFocus
      Exit Sub
   End If
   
   If cboStatus.Text = "Paga" And mskPgto.Text = "" Then
      ShowMsg "Falta a data de pagamento", vbInformation
      mskPgto.SetFocus
      Exit Sub
   End If
   
   If cboStatus.Text = "À Pagar" And IsDate(mskPgto) Then
      ShowMsg "Falta escolher a opção no campos Posição da Conta", vbInformation
      cboStatus.SetFocus
      Exit Sub
   End If
   
   Dim S_SAIDAS As Currency
   Dim S_ENTRADAS As Currency
   Dim S_PARCELAS As Currency
   Dim S_TOTAIS As Currency
   
   Dim lNovoCodCaixaSaida As Long
   Dim lNovoCodRetirada As Long
   
   'ADICIONAR NA TABELA CAIXA_SAIDA
   If cboFonte.Text = "CAIXA ATUAL" And cboStatus.Text = "PAGA" Then
      
      'VERIFICAR O STATUS DO CAIXA
      Dim cStatus As Integer
      cStatus = Verificar_Caixa
      Select Case cStatus
         Case -1
            ShowMsg "Este caixa ainda não foi aberto.", vbExclamation
            Exit Sub
         Case 1
            ShowMsg "O caixa está fechado!", vbExclamation
            Exit Sub
      End Select
      
      S_SAIDAS = 0
      S_ENTRADAS = 0
      S_PARCELAS = 0
      
      'mostrar a soma das saídas
      sSQL = "SELECT ISNULL(SUM(valor), 0) AS soma_saida FROM caixa_saida WHERE (data = CONVERT(DATETIME, '" & Format(mskPgto, ocDATA) & "', 103));"
      Set r = dbData.OpenRecordset(sSQL)
      If Not r.BOF Then S_SAIDAS = r("soma_saida")
      If Not r.BOF Then r.Close
      Set r = Nothing
      
      'mostrar a soma das entradas
      sSQL = "SELECT ISNULL(SUM(valor), 0) AS soma_entrada FROM caixa_entrada WHERE (data = CONVERT(DATETIME, '" & Format(mskPgto, ocDATA) & "', 103));"
      Set r = dbData.OpenRecordset(sSQL)
      If Not r.BOF Then S_ENTRADAS = r("soma_entrada")
      If Not r.BOF Then r.Close
      Set r = Nothing
      
      'mostrar a soma das parcelas
      sSQL = "SELECT ISNULL(SUM(valor_final), 0) AS soma_parcelas FROM parcelas WHERE (pagamento = CONVERT(DATETIME, '" & Format(mskPgto, ocDATA) & "', 103));"
      Set r = dbData.OpenRecordset(sSQL)
      If Not r.BOF Then S_PARCELAS = r("soma_parcelas")
      If r.State <> 0 Then r.Close
      Set r = Nothing
      
      'calculo de todas as somas
      S_TOTAIS = S_PARCELAS + S_ENTRADAS
      S_TOTAIS = S_TOTAIS - S_SAIDAS
      
      If S_TOTAIS < CCur(txtVlrUnit) Then
         MsgBox "O Saldo do CAIXA é menor que o valor da conta!", vbInformation
         Exit Sub
      End If
      
      lNovoCodCaixaSaida = AutoNumeracao_Caixa_Saida
      
      sSQL = "SELECT * FROM caixa_saida;"
      Set r = dbData.OpenRecordset(sSQL)
      
      'RS.AddNew
      Atualizar_Dados_Caixa_Saida lNovoCodCaixaSaida, 1
      'RS.Update
      
   ElseIf cboFonte.Text = "SALDOS" And cboStatus.Text = "PAGA" Then
      Dim VALOR_ATUAL As Currency
      Dim var_Saldo As Currency
      Dim RETIRADA_ATUAL As Currency
      Dim VALOR_CONTA As Currency
      Dim RESULTADO As Currency
      
      var_Saldo = 0
      
      'pegar o ultimo saldo
      sSQL = "SELECT TOP 1 saldo_atual AS saldo FROM caixa_saldo ORDER BY data DESC;"
      Set r = dbData.OpenRecordset(sSQL)
      If Not r.BOF Then var_Saldo = r("saldo")
      If r.State <> 0 Then r.Close
      Set r = Nothing
      
      'calculo
      VALOR_CONTA = CCur(lblTotalGeral.Caption)
      RETIRADA_ATUAL = VALOR_CONTA
      
      'pegar a ultima retirada
      sSQL = "SELECT TOP 1 retirada AS retirada_caixa FROM caixa_saldo ORDER BY data DESC;"
      Set r = dbData.OpenRecordset(sSQL)
      If Not r.BOF Then RETIRADA_ATUAL = r("retirada_caixa") + VALOR_CONTA
      If r.State <> 0 Then r.Close
      Set r = Nothing
      
      If var_Saldo < VALOR_CONTA Then
         ShowMsg "O Saldo do CAIXA 2 é menor que o valor da conta!", vbInformation
         Exit Sub
      Else
         VALOR_ATUAL = var_Saldo - VALOR_CONTA
      End If
      
      lNovoCodCaixaSaida = AutoNumeracao_Caixa_Saldo
      
      dbData.Execute "UPDATE caixa_saldo SET retirada = " & Replace(CCur(RETIRADA_ATUAL), ",", ".") & ", " & _
         "saldo_atual = " & Replace(CCur(VALOR_ATUAL), ",", ".") & " WHERE (codigo = " & x & ");"
      
      'ADICIONAR NA TABELA CAIXA_SALDO_RETIRADA
      lNovoCodRetirada = AutoNumeracao_Saldo_Retirada
      
      sSQL = "SELECT * FROM caixa_saldo_retirada;"
      Set r = dbData.OpenRecordset(sSQL)
      
      'RS.AddNew
      Atualizar_Dados_Saldo_Retirada lNovoCodRetirada, lNovoCodCaixaSaida, 1
      'RS.Update
      
      If r.State <> 0 Then r.Close
      Set r = Nothing
      
   End If
   
   'ADICIONAR NA TABELA A_PAGAR
   sSQL = "SELECT * FROM a_pagar WHERE (codigo = " & txtCodigo.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not Atualizar_Dados Then
      ShowMsg "Erro ao atualizar os dados.", vbExclamation
      Exit Sub
   End If
   
   cmdSalvar.Enabled = False
   cmdCancelar.Enabled = False
   cmdAlterar.Enabled = False
   cmdExcluir.Enabled = False
   cmdHabilitarQuitar.Visible = False
   cmdHabilitarHaver.Visible = False
   Limpar_Objetos_Conta
   Form_Load
   'LimparGrid_Conta
   'MostrarGrid_Contas
End Sub

Private Function Verificar_Caixa() As Integer
 Dim sSQL As String
 Dim r As ADODB.Recordset
 Dim cxaStatus As Integer

 cxaStatus = -1   'Não foi aberto
 
 If cmdAlterar.Enabled = True Then
    sSQL = "SELECT status FROM caixa_dia WHERE (data_abertura = CONVERT(DATETIME, '" & Format(mskPgto.FormattedText, ocDATA) & "', 103)) AND (maquina = '" & StatusBar1.Panels(2).Text & "');"
 Else
    sSQL = "SELECT status FROM caixa_dia WHERE (data_abertura = CONVERT(DATETIME, '" & Format(StatusBar1.Panels(4), ocDATA) & "', 103)) AND (maquina = '" & StatusBar1.Panels(2).Text & "');"
 End If
 
 Set r = dbData.OpenRecordset(sSQL)
 If Not r.BOF Then cxaStatus = CInt(ValidateNull(r("status")))   '0 = aberto, 1 = fechado
 If r.State <> 0 Then r.Close
 Set r = Nothing
 Verificar_Caixa = cxaStatus
End Function
Private Function AutoNumeracao_Caixa_Saida() As Long
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim lRet As Long
   
   lRet = 0
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod FROM caixa_saida;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then lRet = r("cod") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   AutoNumeracao_Caixa_Saida = lRet
End Function

Private Function AutoNumeracao_Caixa_Saldo() As Long
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim lRet As Long
      
   lRet = 0
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod FROM caixa_saldo;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then lRet = r("cod")
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   AutoNumeracao_Caixa_Saldo = lRet
End Function

Private Function AutoNumeracao_Saldo_Retirada() As Long
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim lRet As Long
   
   lRet = 0
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod FROM caixa_saldo_retirada;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then lRet = r("cod") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   AutoNumeracao_Saldo_Retirada = lRet
End Function

Private Function Atualizar_Dados_Saldo_Retirada_Haver(ByVal Cod As Long, ByVal CodSaldo As Long, ByVal Acao As Integer) As Boolean
   'A atualização deve ser feita utilizando o comando UPDATE do sql
   'e não mais usando o método .Update do Recordset
   
   'Não se deve comparar se o campo está vazio ou não, pois dessa forma não
   'haverá atualização quando for necessário apagar alguma informação
   
   Dim sSQL As String
      
   If Acao = 1 Then        'Insere novo
      sSQL = "INSERT INTO caixa_saldo_retirada (codigo, cod_saldo, data, descricao, valor) VALUES (" & _
         Cod & ", " & CodSaldo & ", CONVERT(DATETIME, '" & Format$(CDate(mskDataHaver), ocDATA) & "', 103), '" & _
         cboDescricao.Text & "', " & Replace(CCur(txtValorHaver.Text), ",", ".") & ");"
   
   ElseIf Acao = 2 Then    'Atualiza
      'Comando de atualização
      sSQL = "UPDATE caixa_saida_retirada SET " & _
         "data = CONVERT(DATETIME, '" & Format(CDate(mskDataHaver), ocDATA) & "', 103), " & _
         "descricao = '" & cboDescricao.Text & "', " & _
         "valor = " & Replace(CCur(txtValorHaver.Text), ",", ".")
      
      'Condição para atualização
      sSQL = sSQL & "WHERE (codigo = " & Cod & ") AND (cod_saldo = " & CodSaldo & ");"
   End If

   'Retorna o resultado da atualização
   Atualizar_Dados_Saldo_Retirada_Haver = dbData.Execute(sSQL)
End Function

Private Function Atualizar_Dados_Saldo_Retirada(ByVal Cod As Long, ByVal CodSaldo As Long, ByVal Acao As Integer) As Boolean
   'A atualização deve ser feita utilizando o comando UPDATE do sql
   'e não mais usando o método .Update do Recordset
   
   'Não se deve comparar se o campo está vazio ou não, pois dessa forma não
   'haverá atualização quando for necessário apagar alguma informação
   
   Dim sSQL As String
   
   If Acao = 1 Then        'Insere novo
      sSQL = "INSERT INTO caixa_saldo_retirada (codigo, data, cod_saldo, tipo, cod_descricao, descricao, valor) VALUES (" & _
         Cod & ", CONVERT(DATETIME, '" & Format(mskPgto.Text, ocDATA) & "', 103), " & CodSaldo & ", 'CONTA A PAGAR', " & txtCodigo.Text & ", '" & _
         cboDescricao.Text & "', " & Replace(CCur(lblTotalGeral.Caption), ",", ".") & ");"
   
   ElseIf Acao = 2 Then
      sSQL = "UPDATE caixa_saldo_retirada SET " & _
         "data = CONVERT(DATETIME, '" & Format(CDate(mskPgto.Text), ocDATA) & "', 103), " & _
         "cod_saldo = " & CodSaldo & ", " & _
         "tipo = 'CONTA A PAGAR', " & _
         "cod_descricao = " & txtCodigo.Text & ", " & _
         "descricao = '" & cboDescricao.Text & "', " & _
         "valor = " & Replace(CCur(lblTotalGeral.Caption), ",", ".")
      
      'Condição para atualização
      sSQL = sSQL & " WHERE (codigo = " & Cod & ");"
   End If
   
   'Retorna o resultado da atualização
   Atualizar_Dados_Saldo_Retirada = dbData.Execute(sSQL)
End Function

Private Function Atualizar_Dados_Caixa_Saida(ByVal Cod As Long, ByVal Acao As Integer) As Boolean
   
   Dim sSQL As String
   
   If Acao = 1 Then
      sSQL = "INSERT INTO caixa_saida (codigo, subdescricao, descricao, data, valor, setor, hora) VALUES (" & _
         Cod & ", 'PGTO DE CONTA', 'CONTA: " & cboDescricao.Text & "', CONVERT(DATETIME, '" & Format$(CDate(mskPgto), ocDATA) & "', 103), " & _
         Replace(CCur(lblTotalGeral.Caption), ",", ".") & ", '" & cboSetor.Text & "', '" & lblHora.Caption & "');"
   
   ElseIf Acao = 2 Then
      sSQL = "UPDATE caixa_saida SET " & _
         "subdescricao = 'PGTO DE CONTA', " & _
         "descricao = 'CONTA: " & cboDescricao.Text & "', " & _
         "data = CONVERT(DATETIME, '" & Format$(CDate(mskPgto), ocDATA) & "', 103), " & _
         "valor = " & Replace(CCur(lblTotalGeral.Caption), ",", ".") & ", " & _
         "setor = '" & cboSetor.Text & "', " & _
         "hora = '" & lblHora.Caption & "' "
      
      'Condição para atualização
      sSQL = sSQL & "WHERE (codigo = " & Cod & ");"
   End If
   
   'Retorna o resultado da atualização
   Atualizar_Dados_Caixa_Saida = dbData.Execute(sSQL)
End Function

Private Sub cmdCal1_Click()
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

mskVenc = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdCancelar_Click()
Limpar_Objetos_Conta
Form_Load
End Sub

Private Sub cmdExcluir_Click()
   Dim sSQL As String
   Dim bRet As Boolean
   
   On Error GoTo TrataErro
   
   If cboStatus.Text = "" Or cboDescricao.Text = "" Or cboFavorecido.Text = "" Or txtVlrUnit.Text = "" Or cboTipo.Text = "" Then
      ShowMsg "Formulário incompleto!", vbInformation
      cboStatus.SetFocus
      Exit Sub
   End If
   
   If ShowMsg("Excluir essa conta?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
   
   'sSQL = "SELECT * FROM a_pagar WHERE (codigo = " & txtCodigo.Text & ");"
   'Set r = dbData.OpenRecordset(sSQL)
   
   sSQL = "DELETE FROM a_pagar WHERE (codigo = " & txtCodigo.Text & ");"
   bRet = dbData.Execute(sSQL)
   
   If Not bRet Then
      Exit Sub
   End If
   
   If cboFonte.Text = "SALDOS" Then
      sSQL = "DELETE FROM caixa_saldo_retirada WHERE (COD_DESCRICAO = " & txtCodigo.Text & ");"
      bRet = dbData.Execute(sSQL)
   
      If Not bRet Then
         Exit Sub
      End If
   End If
   
   MostrarGrid_Contas   'ATUALIZAR O GRID DA CONSULTA
   Limpar_Objetos_Conta
   Form_Load
   
   optTodos_Click
   Exit Sub
   
TrataErro:
   If Err.Number = 3421 Then
      ShowMsg "FORMULÁRIO INCOMPLETO!" & vbCrLf & "Verifique se todos os dados estão nos campos.", vbInformation
      Exit Sub
   End If
End Sub

Private Sub cmdExibir_Click()
   MostrarGrid_Contas
End Sub

Private Sub cmdHabilitarHaver_Click()
   SSTab1.Tab = 2
   frmHaver.Enabled = True
   mskDataHaver.Text = Format(Date, "dd/mm/yy")
   cmdHabilitarQuitar.Visible = False
   cmdHabilitarHaver.Visible = False
   txtValorHaver.SetFocus
End Sub

Private Sub cmdHabilitarQuitar_Click()
   Label10.Enabled = True
   Label11.Enabled = True
   mskPgto.Enabled = True
   mskPgto.Text = Format(Date, "DD/MM/YY")
   cboFonte.Enabled = True
   cboStatus.Text = "PAGA"
   cmdHabilitarQuitar.Visible = False
   cmdHabilitarHaver.Visible = False
   cboFonte.SetFocus
End Sub

Private Sub cmdImprimir_Click()
   Dim r As ADODB.Recordset
   
   Me.Hide
   
   Set r = dbData.OpenRecordset(printSQL)
   
   Set REL_ContaApagar.Relatorio.Recordset = r
   REL_ContaApagar.dfQuant.Caption = "QUANTIDADE: " & txtCONquant.Text
   REL_ContaApagar.lblTitulo.Caption = "RELATÓRIO DE CONTAS " & cboCONStatus.Text & ""
   REL_ContaApagar.rfsubtotal.Caption = txtCONsubtotal.Text
   REL_ContaApagar.rfhaveres.Caption = txtCONhaveres.Text
   REL_ContaApagar.rftotal.Caption = txtCONtotal.Text

   If optTodos.Value = Checked Then
      REL_ContaApagar.dfTipo.Caption = "Tipo: Todos os registros"
   ElseIf optIntervalo.Value = Checked Then
      REL_ContaApagar.dfTipo.Caption = "Tipo: Intervalo de " & Mask1.Text & " à " & Mask2.Text
   ElseIf optMensal.Value = Checked Then
      REL_ContaApagar.dfTipo.Caption = "Tipo: Mês = " & cboMES.Text & "/" & cboAno.Text
   ElseIf optNome.Value = Checked Then
      REL_ContaApagar.dfTipo.Caption = "Tipo: Favorecido"
   Else
      REL_ContaApagar.dfTipo.Caption = "Tipo:"
   End If

   REL_ContaApagar.Relatorio.Ativar
   Unload REL_ContaApagar
   
   Me.Show 1
End Sub

Private Sub cmdNovo_Click()
   frmCadastro.Enabled = True
   Limpar_Objetos_Conta
   LimparObjetos_Haver
   'LimparGrid_Consulta
   LimparGrid_Haver
   LimparGrid_Historico
   cmdSalvar.Enabled = True
   cmdCancelar.Enabled = True
   cmdAlterar.Enabled = False
   cmdExcluir.Enabled = False
   cboStatus.Text = "À PAGAR"
   cboStatus.Locked = True
   mskCadastro.Text = Format(CDate(Date), "dd/mm/yy")
   SSTab1.Tab = 0
   cboStatus.Text = "À PAGAR"
   cboDescricao.SetFocus
   cmdNovo.Enabled = False
End Sub

Private Sub cmdRemoverHaver_Click()
   On Error GoTo erro
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   Dim var_Linha As Integer
   Dim LAST_HAVER As Currency
   Dim ATUAL_HAVER As Currency
   Dim SOMA_HAVERES As Currency
   
   'Row retorna a linha selecionada
   var_Linha = Grid_Haver.Row
   
   If Grid_Haver.TextMatrix(var_Linha, 2) = "" Then GoSub erro
   
   If ShowMsg("Deseja excluir o haver: " & Grid_Haver.TextMatrix(var_Linha, 3) & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
   
   'APAGA DA TABELA APAGAR_HAVER
   dbData.Execute "DELETE FROM a_pagar_haver WHERE (codigo = " & Grid_Haver.TextMatrix(var_Linha, 1) & ");"
   
   'APAGA DA TABELA CAIXA_SAIDA
   dbData.Execute "DELETE FROM caixa_saida WHERE (cod_haver = " & Grid_Haver.TextMatrix(var_Linha, 1) & ");"
   
   'ATUALIZAR CAMPO HAVERES NA TABELA A_PAGAR
   sSQL = "SELECT haveres FROM a_pagar WHERE (codigo = " & txtCodigo.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then LAST_HAVER = r("haveres")
   If r.State <> 0 Then r.Clone
   Set r = Nothing
   
   ATUAL_HAVER = CCur(Grid_Haver.TextMatrix(var_Linha, 3))
   SOMA_HAVERES = LAST_HAVER - ATUAL_HAVER
   
   dbData.Execute "UPDATE a_pagar SET haveres = " & Replace(SOMA_HAVERES, ",", ".") & " WHERE (codigo = " & txtCodigo.Text & ");"
   
   If Grid_Haver.TextMatrix(var_Linha, 4) = "SALDOS" Then
      Dim VALOR_ATUAL As Currency
      Dim VALOR_CONTA As Currency
      Dim RESULTADO As Currency
      Dim RETIRADA_ATUAL As Currency
      Dim var_Saldo As Currency
      
      var_Saldo = 0
      VALOR_CONTA = CCur(Grid_Haver.TextMatrix(var_Linha, 3))
      RETIRADA_ATUAL = 0
      
      sSQL = "SELECT TOP 1 saldo_atura AS saldo FROM caixa_saldo ORDER BY data DESC;"
      Set r = dbData.OpenRecordset(sSQL)
      If Not r.BOF Then var_Saldo = r("saldo")
      If r.State <> 0 Then r.Close
      Set r = Nothing
      
      sSQL = "SELECT TOP 1 retirada AS retirada_caixa FROM caixa_saldo ORDER BY data DESC;"
      Set r = dbData.OpenRecordset(sSQL)
      If Not r.BOF Then RETIRADA_ATUAL = r("retirada_caixa") - VALOR_CONTA
      If r.State <> 0 Then r.Close
      Set r = Nothing
      
      VALOR_ATUAL = var_Saldo + VALOR_CONTA
      
      AutoNumeracao_Caixa_Saldo
      
      'ver isso depois
      'dbData.Execute "UPDATE caixa_saldo SET retirada = " & Replace(cRETIRADA_ATUAL, ",", ".") & ", " & saldo_atual = " & Replace(VALOR_ATUAL, ",", ".") & " WHERE (codigo = " & X & ");"
      
      'APAGA DA TABELA CAIXA_SALDO_RETIRADA
      dbData.Execute "DELETE FROM caixa_saldo_retirada WHERE (cod_haver = " & Grid_Haver.TextMatrix(var_Linha, 1) & ");"
   End If
   
   txtCodigo_Change
   MostrarGrid_Contas    'ATUALIZAR O GRID DA CONSULTA
   Calcular_Valores
   Exit Sub
   
erro:
   MsgBox "Não existe nenhum haver selecionado para ser excluído!", vbExclamation
   Exit Sub
End Sub

Private Sub cmdSalvar_Click()
   Dim lNovoCod As Long
   
   If Not IsDate(mskVenc.Text) Then
      ShowMsg "Data Inválida", vbExclamation
      mskVenc.SetFocus
      Exit Sub
   End If
   
   If cboStatus.Text = "" Or cboDescricao.Text = "" Or cboFavorecido.Text = "" Or txtVlrUnit.Text = "" Or cboTipo.Text = "" Then
      ShowMsg "Falta preencher alguns campos!", vbExclamation
      cboStatus.SetFocus
      Exit Sub
   End If
   
   'ADICIONAR NA TABELA A_PAGAR
   'sSQL = "SELECT * FROM a_pagar;"
   'Set r = dbData.OpenRecordset(sSQL)
   
   For i = 1 To Val(txtDuplic.Text)
      'Enconta a última chave primária
      lNovoCod = Auto_Numeracao_APagar
      
      Data = DateAdd("m", i - 1, mskVenc.FormattedText)
      
      'PARCELA = CCur(txtVlrUnit.Text)
      
      'RS.AddNew
      Adicionar_Dados_Conta lNovoCod, 1
      'RS.Update
   Next
   
   Limpar_Objetos_Conta
   Form_Load
   'optTodos_Click
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub Form_Load()
Set moCombo = New cComboHelper
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdNovo.Enabled = True
Label10.Enabled = False
Label11.Enabled = False
mskPgto.Enabled = False
cboFonte.Enabled = False
cboCONStatus_GotFocus
cboCONsetor_GotFocus
cboCONForma_GotFocus
cboCONStatus.ListIndex = 0
cboCONForma.ListIndex = 0
cboStatus.Locked = True
LimparGrid_Conta
LimparGrid_Historico
LimparGrid_Haver
frmCadastro.Enabled = False
SSTab1.Tab = 0
If cboCONsetor.ListCount <> 0 Then cboCONsetor.ListIndex = cboCONsetor.ListCount - 1
optTodos.Value = Checked

'colocar o nome da maquina na barra de status
Dim var_Maquina As String
Dim oIni As Ini

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Maquina = oIni.LerTexto("DADOS_MAQUINA", "maquina")
Set oIni = Nothing

StatusBar1.Panels(2).Text = var_Maquina
   
StatusBar1.Panels(3).Text = Format(Date, "dd/mm/yy")
End Sub

Private Sub Calcular_Valores()
   Dim Total As Currency, HAVER As Currency, TOTAL_GERAL As Currency
   
   Total = 0
   HAVER = 0
   
   If lblTotal.Caption <> "" Then Total = CCur(lblTotal.Caption)
   If lblTotalHaver.Caption <> "" Then HAVER = CCur(lblTotalHaver.Caption)
   
   TOTAL_GERAL = Total - HAVER
   lblTotalGeral = Format(TOTAL_GERAL, ocMONEY)
End Sub

Public Function Verifica_Dia(DIA, mes As Integer)
   Dim diasDoMes As Variant
   
   DIA = Val(DIA)
   diasDoMes = Array(31, 28, 30, 30, 31, 30, 31, 30, 30, 31, 30, 31)
   
   If DIA = 31 Then
      Verifica_Dia = diasDoMes(mes - 1)
   Else
      Verifica_Dia = DIA
   End If
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set moCombo = Nothing
End Sub

Private Sub Grid_DblClick()
   cmdSalvar.Enabled = False
   cmdCancelar.Enabled = False
   cmdAlterar.Enabled = True
   cmdExcluir.Enabled = True
   txtCodigo.Text = ""
   txtCodigo.Text = (Grid.TextMatrix(Grid.Row, 1))
End Sub

Private Function Atualizar_Dados_Haver(ByVal Acao As Integer) As Boolean
   'A atualização deve ser feita utilizando o comando UPDATE do sql
   'e não mais usando o método .Update do Recordset
   
   'Não se deve comparar se o campo está vazio ou não, pois dessa forma não
   'haverá atualização quando for necessário apagar alguma informação
   
   Dim sSQL As String
   
   If Acao = 1 Then
      sSQL = "INSERT INTO a_pagar_haver (codigo, cod_conta, data, fonte, valor, hora) VALUES (" & _
         txtCodHaver.Text & ", " & txtCodigo.Text & ", CONVERT(DATETIME, '" & Format(CDate(mskDataHaver), ocDATA) & "', 103), '" & _
         cboFonteHaver.Text & "', " & Replace(CCur(txtValorHaver.Text), ",", ".") & ", '" & lblHora.Caption & "');"
      
   Else
      sSQL = "UPDATE a_pagar_haver SET " & _
         "cod_conta = " & txtCodigo.Text & ", " & _
         "data = CONVERT(DATETIME, '" & Format$(CDate(mskDataHaver.Text), ocDATA) & "', 103), " & _
         "fonte = '" & cboFonteHaver.Text & "', " & _
         "valor = " & Replace(CCur(txtValorHaver.Text), ",", ".") & ", " & _
         "hora = '" & lblHora.Caption & "' "
      
      'Condição para atualização
      sSQL = sSQL & "WHERE (codigo = " & txtCodHaver.Text & ");"
   End If
   
   'Retorna o resultado da atualização
   Atualizar_Dados_Haver = dbData.Execute(sSQL)
End Function

Private Sub Auto_Numeracao_Haver()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS codigo_haver FROM a_pagar_haver;"
   Set r = dbData.OpenRecordset(sSQL)
   txtCodHaver.Text = r("codigo_haver") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub LimparObjetos_Haver()
   txtCodHaver.Text = ""
   mskDataHaver.Mask = ""
   mskDataHaver.Text = ""
   txtValorHaver.Text = ""
   cboFonteHaver.Text = ""
   mskDataHaver.Text = Format(Date, "dd/mm/yy")
End Sub

Private Sub MostrarGrid_Haver()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT * FROM a_pagar_haver WHERE (cod_conta = " & txtCodigo.Text & ") ORDER BY data, codigo;"
   Set r = dbData.OpenRecordset(sSQL)
   'AND (status = 0)
   FormatarGrid_Haver r
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub lblTotal_Change()
   Calcular_Valores
End Sub

Private Sub lblTotalGridHaver_Change()
   If lblTotalGridHaver.Caption = "" Then
      lblTotalHaver.Caption = Format(0, ocMONEY)
   Else
      lblTotalHaver.Caption = Format(lblTotalGridHaver.Caption, ocMONEY)
   End If
End Sub

Private Sub lblTotalHaver_Change()
   Calcular_Valores
End Sub

Private Sub MASK1_KeyPress(KeyAscii As Integer)
   Mask1.Mask = "##/##/##"
End Sub

Private Sub Mask1_LostFocus()
   If Mask1.Text = "" Then Exit Sub Else Mask2.SetFocus
End Sub

Private Sub Mask2_KeyPress(KeyAscii As Integer)
   Mask2.Mask = "##/##/##"
End Sub

Private Sub Mask2_LostFocus()
   If Mask2.Text = "" Then Exit Sub Else cmdExibir.SetFocus
End Sub

Private Sub mskCadastro_GotFocus()
   SelectControl mskCadastro
End Sub

Private Sub mskCadastro_KeyPress(KeyAscii As Integer)
   mskCadastro.Mask = "##/##/##"
End Sub

Private Sub mskCadastro_LostFocus()
   If mskCadastro.Text = "" Or mskCadastro.Text = "__/__/__" Then
      mskCadastro.Mask = ""
      mskCadastro.Text = ""
      Exit Sub
   Else
      If IsDate(mskCadastro.Text) Then
         Exit Sub
      Else
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskCadastro.SetFocus
         SelectControl mskCadastro
      End If
   End If
End Sub

Private Sub mskDataHaver_GotFocus()
   SelectControl mskDataHaver
End Sub

Private Sub mskDataHaver_KeyPress(KeyAscii As Integer)
   mskDataHaver.Mask = "##/##/##"
End Sub

Private Sub mskDataHaver_LostFocus()
   txtValorHaver.SetFocus
End Sub

Private Sub mskPgto_GotFocus()
   SelectControl mskPgto
End Sub

Private Sub mskPgto_KeyPress(KeyAscii As Integer)
   mskPgto.Mask = "##/##/##"
End Sub

Private Sub mskPgto_LostFocus()
   If mskPgto.Text = "" Or mskPgto.Text = "__/__/__" Then
      mskPgto.Mask = ""
      mskPgto.Text = ""
      Exit Sub
   Else
      If IsDate(mskPgto.Text) Then
         Exit Sub
      Else
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskPgto.SetFocus
         SelectControl mskPgto
      End If
   End If
End Sub

Private Sub cboDescricao_GotFocus()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   Dim VarText As String
   VarText = cboDescricao.Text
   
   cboDescricao.Clear
   
   sSQL = "SELECT descricao FROM a_pagar ORDER BY descricao;"
   Set r = dbData.OpenRecordset(sSQL)
   
   Do While Not r.EOF
      cboDescricao.AddItem r("descricao")
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing

   cboDescricao.Text = VarText
   
   moCombo.AttachTo cboDescricao
End Sub

Private Sub cboDescricao_KeyPress(KeyAscii As Integer)
   If Len(cboDescricao) = 40 Then cboSetor.SetFocus
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub mskVenc_GotFocus()
   SelectControl mskVenc
End Sub

Private Sub optIntervalo_Click()
   optTodos.Value = 0
   optMensal.Value = 0
   Ocultar_SubCriterios
   If Mask1.Visible = True Then Mask1.SetFocus
End Sub

Private Sub optMensal_Click()
   optTodos.Value = 0
   optIntervalo.Value = 0
   Ocultar_SubCriterios
   If cboMES.Visible = True Then cboMES.SetFocus
End Sub

Private Sub optNome_Click()
   optTodos.Value = 0
   Ocultar_SubCriterios
   If cboNome.Visible = True Then cboNome.SetFocus
End Sub

Private Sub optOrdCadastro_Click()
   If optTodos.Value = Checked Then
      optTodos_Click
   ElseIf optMensal.Value = Checked Then
      cmdExibir_Click
   ElseIf optIntervalo.Value = Checked Then
      cmdExibir_Click
   ElseIf optNome.Value = Checked Then
      cmdExibir_Click
   End If
End Sub

Private Sub optOrdFavorecido_Click()
   If optTodos.Value = Checked Then
      optTodos_Click
   ElseIf optMensal.Value = Checked Then
      cmdExibir_Click
   ElseIf optIntervalo.Value = Checked Then
      cmdExibir_Click
   ElseIf optNome.Value = Checked Then
      cmdExibir_Click
   End If
End Sub

Private Sub optOrdForma_Click()
   If optTodos.Value = Checked Then
      optTodos_Click
   ElseIf optMensal.Value = Checked Then
      cmdExibir_Click
   ElseIf optIntervalo.Value = Checked Then
      cmdExibir_Click
   ElseIf optNome.Value = Checked Then
      cmdExibir_Click
   End If
End Sub

Private Sub optOrdReferente_Click()
   If optTodos.Value = Checked Then
      optTodos_Click
   ElseIf optMensal.Value = Checked Then
      cmdExibir_Click
   ElseIf optIntervalo.Value = Checked Then
      cmdExibir_Click
   ElseIf optNome.Value = Checked Then
      cmdExibir_Click
   End If
End Sub

Private Sub optOrdVencimento_Click()
   If optTodos.Value = Checked Then
      optTodos_Click
   ElseIf optMensal.Value = Checked Then
      cmdExibir_Click
   ElseIf optIntervalo.Value = Checked Then
      cmdExibir_Click
   ElseIf optNome.Value = Checked Then
      cmdExibir_Click
   End If
End Sub

Private Sub optPagamento_Click()
   If optPagamento.Value = True Then
      chkPgtoOutros.Enabled = True
      chkPgtoMes.Enabled = True
   Else
      chkPgtoOutros.Enabled = False
      chkPgtoMes.Enabled = False
   End If
   
   If cboCONStatus.Text = "À PAGAR" Then cboCONStatus.Text = "PAGA"
   cmdExibir_Click
End Sub

Private Sub optTodos_Click()
   If optTodos.Value = 1 Then
      optMensal.Value = 0
      optIntervalo.Value = 0
      optNome.Value = 0
      cmdExibir.Visible = False
   End If
   
   Ocultar_SubCriterios
   MostrarGrid_Contas
End Sub

Private Sub optVencimento_Click()
   chkPgtoOutros.Enabled = False
   chkPgtoMes.Enabled = False
   chkPgtoOutros.Value = 0
   chkPgtoMes.Value = 0
   cmdExibir_Click
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab = 0 Then
      cmdImprimir.Visible = False
      frmHaver.Enabled = False
   ElseIf SSTab1.Tab = 1 Then
      cmdImprimir.Visible = True
      frmHaver.Enabled = False
   ElseIf SSTab1.Tab = 2 Then
      cmdImprimir.Visible = False
   ElseIf SSTab1.Tab = 3 Then
      frmHaver.Enabled = False
   End If
End Sub

Private Sub Timer1_Timer()
   lblHora.Caption = Format(Time, "hh:mm")
End Sub

Private Sub txtCodigo_Change()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   If txtCodigo.Text = "" Then Exit Sub
   
   sSQL = "SELECT * FROM a_pagar WHERE (codigo = " & txtCodigo.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then
      frmCadastro.Enabled = True
      Limpar_Objetos_Conta
      LimparObjetos_Haver
      cboStatus.Locked = False
      Mostrar_Conta r
      MostrarGrid_Haver
   End If
   
   Calcular_Valores
   MostrarGrid_Historico
End Sub

Private Sub Mostrar_Conta(rTabela As ADODB.Recordset)
   If Not rTabela Is Nothing Then
      mskVenc.Text = Format(rTabela("vencimento"), "dd/mm/yy")
      mskCadastro.Text = Format(rTabela("cadastro"), "dd/mm/yy")
      txtVlrUnit.Text = Format(rTabela("VALOR_UND"), ocMONEY)
      lblTotal.Caption = Format(rTabela("total"), ocMONEY)
      cboStatus.Text = ValidateNull(rTabela("status"))
      cboDescricao.Text = ValidateNull(rTabela("descricao"))
      cboFavorecido.Text = ValidateNull(rTabela("favorecido"))
      cboSetor.Text = ValidateNull(rTabela("setor"))
      cboTipo.Text = ValidateNull(rTabela("tipo"))
      cboForma.Text = ValidateNull(rTabela("forma"))
      mskPgto.Text = Format(rTabela("data_pgto"), "dd/mm/yy")
      cboFonte.Text = ValidateNull(rTabela("fonte"))
      txtRef.Text = ValidateNull(rTabela("ref"))
      
      If rTabela("status") <> "PAGO" Then
         cmdHabilitarQuitar.Visible = True
         cmdHabilitarHaver.Visible = True
      Else
         cmdHabilitarQuitar.Visible = False
         cmdHabilitarHaver.Visible = False
      End If
   End If
   
   SSTab1.Tab = 0
End Sub

Private Function Adicionar_Dados_Conta(ByVal Cod As Long, ByVal Acao As Integer) As Boolean
   
   Dim sSQL As String
   
   If Acao = 1 Then
      sSQL = "INSERT INTO a_pagar (codigo, cadastro, vencimento, valor_und, status, descricao, favorecido, setor, " & _
         "tipo, forma, ref) VALUES (" & Cod & ", CONVERT(DATETIME, '" & Format$(CDate(mskCadastro), ocDATA) & "', 103), CONVERT(DATETIME, '" & Format$(Data, ocDATA) & "', 103), " & _
         Replace(CCur(txtVlrUnit), ",", ".") & ", '" & cboStatus & "', '" & cboDescricao & "', '" & cboFavorecido & "', '" & _
         cboSetor & "', '" & cboTipo & "', '" & cboForma & "', '" & txtRef & "');"
   
   ElseIf Acao = 2 Then
      'Comando de atualizacao
      sSQL = "UPDATE a_pagar SET " & _
         "cadastro = CONVERT(DATETIME, '" & Format(CDate(mskCadastro), ocDATA) & "', 103), " & _
         "vencimento = CONVERT(DATETIME, '" & Format(Data, ocDATA) & "', 103), " & _
         "valor_und = " & Replace(CCur(txtVlrUnit), ",", ".") & ", " & _
         "status = '" & cboStatus & "', " & _
         "descricao = '" & cboDescricao & "', " & _
         "favorecido = '" & cboFavorecido & "', " & _
         "setor = '" & cboSetor & "', " & _
         "tipo = '" & cboTipo & "', " & _
         "forma = '" & cboForma & "', " & _
         "ref = '" & txtRef & "' "
         
      'Condição para atualização
      sSQL = sSQL & "WHERE (codigo = " & Cod & ");"
   End If
   
   'Retorna o resultado da atualização
   Adicionar_Dados_Conta = dbData.Execute(sSQL)
End Function

Private Sub txtDuplic_GotFocus()
   SelectControl txtDuplic
End Sub

Private Sub mskVenc_KeyPress(KeyAscii As Integer)
   mskVenc.Mask = "##/##/##"
End Sub

Private Sub mskVenc_LostFocus()
   If mskVenc.Text = "" Or mskVenc.Text = "__/__/__" Then
      mskVenc.Mask = ""
      mskVenc.Text = ""
      Exit Sub
   Else
      If Not IsDate(mskVenc.Text) Then
         ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
         mskVenc.SetFocus
         SelectControl mskVenc
      End If
   End If
End Sub




Private Sub txtVlrUnit_GotFocus()
   SelectControl txtVlrUnit
End Sub

Private Sub txtVlrUnit_KeyPress(KeyAscii As Integer)
   KeyAscii = aNumeros(KeyAscii, True)
End Sub

Private Sub txtValorHaver_GotFocus()
   SelectControl txtValorHaver
End Sub

Private Sub txtValorHaver_LostFocus()
   'cboFonteHaver.SetFocus
   If txtValorHaver.Text = "" Then
      txtValorHaver.Text = Format(0, "##,##0.00")
   Else
      txtValorHaver.Text = Format(txtValorHaver, "##,##0.00")
   End If
End Sub

Private Sub txtVlrUnit_Validate(Cancel As Boolean)
If txtVlrUnit.Text = "" Then txtVlrUnit.Text = Format(0, "##,##0.00")
End Sub
