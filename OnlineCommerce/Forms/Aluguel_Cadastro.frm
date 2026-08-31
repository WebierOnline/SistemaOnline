VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Aluguel_Cadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ALUGUEL"
   ClientHeight    =   9585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14730
   ForeColor       =   &H00008000&
   Icon            =   "Aluguel_Cadastro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   14730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   60
      ScaleHeight     =   705
      ScaleWidth      =   14565
      TabIndex        =   51
      Top             =   60
      Width           =   14595
      Begin VB.TextBox txtCodFuncionario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   13260
         TabIndex        =   143
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contrato:"
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
         Height          =   285
         Left            =   10800
         TabIndex        =   84
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label lblCodigo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   12000
         TabIndex        =   81
         Top             =   180
         Width           =   1155
      End
      Begin VB.Image Image1 
         Height          =   600
         Left            =   420
         Picture         =   "Aluguel_Cadastro.frx":23D2
         Top             =   60
         Width           =   600
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ALUGUEL"
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
         Left            =   1320
         TabIndex        =   52
         Top             =   180
         Width           =   1500
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8355
      Left            =   60
      TabIndex        =   44
      Top             =   840
      Width           =   14595
      _ExtentX        =   25744
      _ExtentY        =   14737
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "CADASTRO"
      TabPicture(0)   =   "Aluguel_Cadastro.frx":2BBE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCodPedido"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdImprimirContato"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdNovo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdSalvar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdExcluir"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdAlterar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdCancelar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdCadastrarCliente"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Picture1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "CONSULTA"
      TabPicture(1)   =   "Aluguel_Cadastro.frx":2BDA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtCONtotal"
      Tab(1).Control(1)=   "txtCONquant"
      Tab(1).Control(2)=   "Picture3"
      Tab(1).Control(3)=   "GridConsulta"
      Tab(1).Control(4)=   "cmdFecharAluguel"
      Tab(1).Control(5)=   "Label13"
      Tab(1).Control(6)=   "Label12"
      Tab(1).ControlCount=   7
      Begin VB.TextBox txtCONtotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   -61860
         TabIndex        =   56
         Top             =   7980
         Width           =   1335
      End
      Begin VB.TextBox txtCONquant 
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
         Left            =   -61860
         TabIndex        =   55
         Top             =   7680
         Width           =   1335
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   915
         Left            =   -74940
         ScaleHeight     =   885
         ScaleWidth      =   14325
         TabIndex        =   54
         Top             =   420
         Width           =   14355
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Critérios"
            Height          =   675
            Left            =   4980
            TabIndex        =   68
            Top             =   60
            Width           =   4935
            Begin ChamaleonBtn.chameleonButton cmdFinal 
               Height          =   315
               Left            =   4260
               TabIndex        =   74
               Tag             =   "Calendario"
               Top             =   240
               Visible         =   0   'False
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
               MICON           =   "Aluguel_Cadastro.frx":2BF6
               PICN            =   "Aluguel_Cadastro.frx":2C12
               PICH            =   "Aluguel_Cadastro.frx":4F65
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdInicial 
               Height          =   315
               Left            =   2040
               TabIndex        =   73
               Tag             =   "Calendario"
               Top             =   240
               Visible         =   0   'False
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
               MICON           =   "Aluguel_Cadastro.frx":72B8
               PICN            =   "Aluguel_Cadastro.frx":72D4
               PICH            =   "Aluguel_Cadastro.frx":9627
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.ComboBox cboNome 
               Height          =   315
               Left            =   720
               TabIndex        =   79
               Top             =   240
               Visible         =   0   'False
               Width           =   3855
            End
            Begin VB.TextBox txtCodClienteCons 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   4080
               TabIndex        =   78
               Top             =   240
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.ComboBox cboAno 
               Height          =   315
               Left            =   3120
               Sorted          =   -1  'True
               TabIndex        =   76
               Top             =   240
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.ComboBox cboMES 
               Height          =   315
               ItemData        =   "Aluguel_Cadastro.frx":B97A
               Left            =   1320
               List            =   "Aluguel_Cadastro.frx":B97C
               TabIndex        =   75
               Top             =   240
               Visible         =   0   'False
               Width           =   1755
            End
            Begin MSMask.MaskEdBox Mask2 
               Height          =   315
               Left            =   3300
               TabIndex        =   69
               Top             =   240
               Visible         =   0   'False
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Mask1 
               Height          =   315
               Left            =   1080
               TabIndex        =   70
               Top             =   240
               Visible         =   0   'False
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin VB.Label lblCONnome 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nome:"
               Height          =   195
               Left            =   180
               TabIndex        =   80
               Top             =   240
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.Label lblCONmes 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "E&scolha o mês:"
               Height          =   195
               Left            =   180
               TabIndex        =   77
               Top             =   300
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.Label lblCONint1 
               AutoSize        =   -1  'True
               Caption         =   "Da&ta Inicial:"
               Height          =   195
               Left            =   180
               TabIndex        =   72
               Top             =   300
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.Label lblCONint2 
               AutoSize        =   -1  'True
               Caption         =   "Data &Final:"
               Height          =   195
               Left            =   2460
               TabIndex        =   71
               Top             =   300
               Visible         =   0   'False
               Width           =   765
            End
         End
         Begin VB.Frame frm 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Status"
            Height          =   675
            Left            =   3240
            TabIndex        =   66
            Top             =   60
            Width           =   1695
            Begin VB.ComboBox cboCONStatus 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   60
               TabIndex        =   67
               TabStop         =   0   'False
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Ordem"
            Height          =   675
            Left            =   1680
            TabIndex        =   64
            Top             =   60
            Width           =   1515
            Begin VB.ComboBox cboOrdem 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   60
               TabIndex        =   65
               TabStop         =   0   'False
               Top             =   240
               Width           =   1395
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Filtro"
            Height          =   675
            Left            =   60
            TabIndex        =   62
            Top             =   60
            Width           =   1575
            Begin VB.ComboBox cboFiltro 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   60
               TabIndex        =   63
               TabStop         =   0   'False
               Top             =   240
               Width           =   1455
            End
         End
         Begin ChamaleonBtn.chameleonButton cmdExibirConsulta 
            Height          =   615
            Left            =   9960
            TabIndex        =   60
            Top             =   120
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   1085
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
            BCOL            =   32768
            BCOLO           =   32768
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Aluguel_Cadastro.frx":B97E
            PICN            =   "Aluguel_Cadastro.frx":B99A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdImprimirConsulta 
            Height          =   615
            Left            =   11520
            TabIndex        =   87
            Top             =   120
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   1085
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
            MICON           =   "Aluguel_Cadastro.frx":C274
            PICN            =   "Aluguel_Cadastro.frx":C290
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
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   7875
         Left            =   120
         ScaleHeight     =   7845
         ScaleWidth      =   12645
         TabIndex        =   45
         Top             =   420
         Width           =   12675
         Begin VB.Frame frmProrrogacao 
            Caption         =   "Prorrogação"
            Height          =   1815
            Left            =   6420
            TabIndex        =   109
            Top             =   5940
            Width           =   6195
            Begin VB.TextBox txtQuantItem 
               Height          =   285
               Left            =   2100
               TabIndex        =   141
               Top             =   0
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.TextBox txtTotalAluguelDescProrro 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   3060
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   137
               TabStop         =   0   'False
               Top             =   1140
               Width           =   855
            End
            Begin VB.TextBox txtDescAluguelProrro 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   2400
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   136
               Top             =   1140
               Width           =   615
            End
            Begin VB.TextBox txtTotalAluguelProrro 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   1500
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   134
               TabStop         =   0   'False
               Top             =   1140
               Width           =   915
            End
            Begin VB.TextBox txtEntradaAdiar 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   3960
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   133
               Top             =   1140
               Width           =   795
            End
            Begin VB.TextBox txtValorRealDescProrro 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3840
               MaxLength       =   40
               TabIndex        =   128
               Top             =   0
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtDescProrro 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3300
               MaxLength       =   40
               TabIndex        =   125
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox txtTotalProrro 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   3960
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   124
               TabStop         =   0   'False
               Top             =   480
               Width           =   855
            End
            Begin ChamaleonBtn.chameleonButton chameleonButton5 
               Height          =   315
               Left            =   3000
               TabIndex        =   111
               TabStop         =   0   'False
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
               MICON           =   "Aluguel_Cadastro.frx":C5AA
               PICN            =   "Aluguel_Cadastro.frx":C5C6
               PICH            =   "Aluguel_Cadastro.frx":E919
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.TextBox txtDiasProrrogar 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0E0FF&
               Height          =   315
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   116
               Top             =   1140
               Width           =   495
            End
            Begin VB.TextBox txtValorAluguellProrro 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   120
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   110
               TabStop         =   0   'False
               Top             =   1140
               Width           =   915
            End
            Begin MSMask.MaskEdBox mskDataProrrogar 
               Height          =   315
               Left            =   2160
               TabIndex        =   112
               Top             =   480
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   12632319
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskDataFinalLocacaoProrro 
               Height          =   315
               Left            =   1140
               TabIndex        =   118
               Top             =   480
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskDataInicioProrro 
               Height          =   315
               Left            =   2820
               TabIndex        =   120
               Top             =   0
               Visible         =   0   'False
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin ChamaleonBtn.chameleonButton cmdProrrogacao 
               Height          =   315
               Left            =   4860
               TabIndex        =   122
               ToolTipText     =   "Remove"
               Top             =   540
               Visible         =   0   'False
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "Prorrogar"
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
               MICON           =   "Aluguel_Cadastro.frx":10C6C
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdAdiacao 
               Height          =   315
               Left            =   4860
               TabIndex        =   131
               ToolTipText     =   "Remove"
               Top             =   960
               Visible         =   0   'False
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "Antecipar"
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
               MICON           =   "Aluguel_Cadastro.frx":10C88
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSMask.MaskEdBox mskDataInicialLocacaoProrro 
               Height          =   315
               Left            =   120
               TabIndex        =   132
               Top             =   480
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Desc."
               Height          =   195
               Left            =   3300
               TabIndex        =   140
               Top             =   240
               Width           =   420
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Total"
               Height          =   195
               Left            =   3060
               TabIndex        =   139
               Top             =   900
               Width           =   360
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Desc."
               Height          =   195
               Left            =   2340
               TabIndex        =   138
               Top             =   900
               Width           =   420
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Total"
               Height          =   195
               Left            =   1500
               TabIndex        =   135
               Top             =   900
               Width           =   360
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Entrada"
               Height          =   195
               Left            =   3960
               TabIndex        =   129
               Top             =   900
               Width           =   555
            End
            Begin VB.Label Label40 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Desc."
               Height          =   195
               Left            =   3240
               TabIndex        =   127
               Top             =   1260
               Width           =   420
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Restante"
               Height          =   195
               Left            =   3960
               TabIndex        =   126
               Top             =   240
               Width           =   645
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Inicio da Prorrogação"
               Height          =   195
               Left            =   4320
               TabIndex        =   121
               Top             =   120
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Data Final"
               Height          =   195
               Left            =   1140
               TabIndex        =   119
               Top             =   240
               Width           =   720
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Aluguel"
               Height          =   195
               Left            =   120
               TabIndex        =   117
               Top             =   900
               Width           =   525
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Data"
               Height          =   195
               Left            =   2160
               TabIndex        =   115
               Top             =   240
               Width           =   345
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Locação"
               Height          =   195
               Left            =   120
               TabIndex        =   114
               Top             =   240
               Width           =   630
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Dias"
               Height          =   195
               Left            =   1020
               TabIndex        =   113
               Top             =   900
               Width           =   315
            End
         End
         Begin VB.Frame frmPagamento 
            Caption         =   "Pagamento"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   1035
            Left            =   6420
            TabIndex        =   145
            Top             =   6780
            Visible         =   0   'False
            Width           =   6195
            Begin VB.TextBox txtDesc 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   5220
               Locked          =   -1  'True
               TabIndex        =   43
               Text            =   "0"
               Top             =   660
               Width           =   915
            End
            Begin VB.OptionButton opt2forma 
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   5640
               TabIndex        =   152
               TabStop         =   0   'False
               Top             =   180
               Width           =   495
            End
            Begin VB.OptionButton opt1forma 
               Caption         =   "1 "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   5160
               TabIndex        =   151
               TabStop         =   0   'False
               Top             =   180
               Value           =   -1  'True
               Width           =   495
            End
            Begin VB.CheckBox chkPgtoAuto 
               Caption         =   "Pgto Automático"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   38
               Top             =   240
               Value           =   1  'Checked
               Width           =   1515
            End
            Begin VB.ComboBox cboformaPgto2 
               Height          =   315
               Left            =   2760
               TabIndex        =   41
               Top             =   660
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.ComboBox cboformaPgto1 
               Height          =   315
               Left            =   120
               TabIndex        =   39
               Top             =   660
               Width           =   1515
            End
            Begin VB.TextBox txtValor2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4140
               Locked          =   -1  'True
               TabIndex        =   42
               Text            =   "0"
               Top             =   660
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.TextBox txtValor1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1680
               TabIndex        =   40
               Text            =   "0"
               Top             =   660
               Width           =   1035
            End
            Begin VB.Label lblDesc 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Desconto"
               Height          =   195
               Left            =   5220
               TabIndex        =   153
               Top             =   460
               Width           =   690
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Qtde. Formas:"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   4080
               TabIndex        =   150
               Top             =   120
               Width           =   975
            End
            Begin VB.Label lblformaPgto2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Forma Pgto 2"
               Height          =   195
               Left            =   2760
               TabIndex        =   149
               Top             =   460
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.Label lblformaPgto1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Forma Pgto 1"
               Height          =   195
               Left            =   120
               TabIndex        =   148
               Top             =   460
               Width           =   945
            End
            Begin VB.Label lblValor2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor"
               Height          =   195
               Left            =   4140
               TabIndex        =   147
               Top             =   460
               Visible         =   0   'False
               Width           =   360
            End
            Begin VB.Label lblValor1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor:"
               Height          =   195
               Left            =   1680
               TabIndex        =   146
               Top             =   460
               Width           =   405
            End
         End
         Begin VB.Frame frmCliente 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Locatário"
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
            Height          =   975
            Left            =   60
            TabIndex        =   48
            Top             =   60
            Width           =   12495
            Begin VB.ComboBox cboSituacao 
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   120
               TabIndex        =   1
               TabStop         =   0   'False
               Top             =   480
               Width           =   1995
            End
            Begin VB.ComboBox cboCidadeObra 
               Height          =   315
               Left            =   10080
               TabIndex        =   4
               Top             =   480
               Width           =   2295
            End
            Begin VB.TextBox txtDescricaoObra 
               Height          =   315
               Left            =   7200
               TabIndex        =   3
               Top             =   480
               Width           =   2835
            End
            Begin VB.ComboBox txtCliente 
               Height          =   315
               Left            =   2160
               TabIndex        =   2
               Top             =   480
               Width           =   4995
            End
            Begin VB.TextBox txtCodCliente 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5880
               TabIndex        =   49
               Top             =   240
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Situação:"
               Height          =   195
               Left            =   120
               TabIndex        =   142
               Top             =   240
               Width           =   675
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Cidade:"
               Height          =   195
               Left            =   10080
               TabIndex        =   83
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Obra:"
               Height          =   195
               Left            =   7200
               TabIndex        =   82
               Top             =   240
               Width           =   390
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cliente:"
               Height          =   195
               Left            =   2160
               TabIndex        =   50
               Top             =   240
               Width           =   525
            End
         End
         Begin VB.Frame frmReferente 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Aluguel"
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
            Height          =   4935
            Left            =   60
            TabIndex        =   46
            Top             =   1020
            Width           =   12555
            Begin VB.Frame frmCodicoes 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Condições do aluguel"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   915
               Left            =   120
               TabIndex        =   94
               Top             =   1200
               Width           =   12375
               Begin VB.ComboBox txtTipoCobranca 
                  BackColor       =   &H00C0FFFF&
                  Height          =   315
                  Left            =   120
                  TabIndex        =   9
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   915
               End
               Begin VB.ComboBox cboFormaEntrada 
                  BackColor       =   &H00C0FFFF&
                  Height          =   315
                  Left            =   9540
                  TabIndex        =   22
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   1755
               End
               Begin VB.TextBox txtEntradaReal 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Height          =   315
                  Left            =   8760
                  Locked          =   -1  'True
                  MaxLength       =   40
                  TabIndex        =   21
                  TabStop         =   0   'False
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   735
               End
               Begin VB.TextBox txtEntrada 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
                  Height          =   315
                  Left            =   7860
                  MaxLength       =   40
                  TabIndex        =   20
                  Top             =   480
                  Width           =   855
               End
               Begin VB.TextBox txtSubTotal 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Height          =   315
                  Left            =   6960
                  Locked          =   -1  'True
                  MaxLength       =   40
                  TabIndex        =   19
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   855
               End
               Begin VB.TextBox txtQuant 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   4560
                  Locked          =   -1  'True
                  TabIndex        =   16
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   795
               End
               Begin VB.TextBox txtTotal 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   11340
                  Locked          =   -1  'True
                  MaxLength       =   40
                  TabIndex        =   23
                  TabStop         =   0   'False
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   855
               End
               Begin VB.TextBox txtDesconto 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
                  Height          =   315
                  Left            =   6300
                  MaxLength       =   40
                  TabIndex        =   18
                  Top             =   480
                  Width           =   615
               End
               Begin VB.TextBox txtValor 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   5400
                  Locked          =   -1  'True
                  MaxLength       =   40
                  TabIndex        =   17
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   855
               End
               Begin ChamaleonBtn.chameleonButton chameleonButton1 
                  Height          =   315
                  Left            =   1920
                  TabIndex        =   11
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
                  MICON           =   "Aluguel_Cadastro.frx":10CA4
                  PICN            =   "Aluguel_Cadastro.frx":10CC0
                  PICH            =   "Aluguel_Cadastro.frx":13013
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin MSMask.MaskEdBox mskInicio 
                  Height          =   315
                  Left            =   1080
                  TabIndex        =   10
                  Top             =   480
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   12648447
                  PromptChar      =   "_"
               End
               Begin ChamaleonBtn.chameleonButton chameleonButton2 
                  Height          =   315
                  Left            =   3660
                  TabIndex        =   14
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
                  MICON           =   "Aluguel_Cadastro.frx":15366
                  PICN            =   "Aluguel_Cadastro.frx":15382
                  PICH            =   "Aluguel_Cadastro.frx":176D5
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin MSMask.MaskEdBox mskFinal 
                  Height          =   315
                  Left            =   2820
                  TabIndex        =   13
                  Top             =   480
                  Width           =   1155
                  _ExtentX        =   2037
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   12648447
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox mskHoraInicio 
                  Height          =   315
                  Left            =   2220
                  TabIndex        =   12
                  Top             =   480
                  Width           =   555
                  _ExtentX        =   979
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   12648447
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox mskHoraFinal 
                  Height          =   315
                  Left            =   3960
                  TabIndex        =   15
                  Top             =   480
                  Width           =   555
                  _ExtentX        =   979
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   12648447
                  PromptChar      =   "_"
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Tipo"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   156
                  Top             =   240
                  Width           =   315
               End
               Begin VB.Label lbl5 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Data Inicial"
                  Height          =   195
                  Left            =   1080
                  TabIndex        =   155
                  Top             =   240
                  Width           =   795
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Data Final"
                  Height          =   195
                  Left            =   2820
                  TabIndex        =   154
                  Top             =   240
                  Width           =   720
               End
               Begin VB.Label lblFormaEntrada 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000009&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Forma de Pgto:"
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
                  Left            =   9540
                  TabIndex        =   103
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   1305
               End
               Begin VB.Label Label25 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Entrada(%)"
                  Height          =   195
                  Left            =   7860
                  TabIndex        =   102
                  Top             =   240
                  Width           =   765
               End
               Begin VB.Label Label24 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Subtotal"
                  Height          =   195
                  Left            =   6960
                  TabIndex        =   101
                  Top             =   240
                  Width           =   585
               End
               Begin VB.Label lbl1 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Total"
                  Height          =   195
                  Left            =   11340
                  TabIndex        =   98
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   360
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Qtde.Dias"
                  Height          =   195
                  Left            =   4560
                  TabIndex        =   97
                  Top             =   240
                  Width           =   705
               End
               Begin VB.Label Label10 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Desc."
                  Height          =   195
                  Left            =   6300
                  TabIndex        =   96
                  Top             =   240
                  Width           =   420
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Valor"
                  Height          =   195
                  Left            =   5400
                  TabIndex        =   95
                  Top             =   240
                  Width           =   360
               End
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Equipamento"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   915
               Left            =   120
               TabIndex        =   88
               Top             =   240
               Width           =   12375
               Begin VB.TextBox txtValorAluguel 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3840
                  Locked          =   -1  'True
                  TabIndex        =   6
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   915
               End
               Begin VB.ComboBox cboEquipamento 
                  BackColor       =   &H00C0FFFF&
                  Height          =   315
                  Left            =   120
                  TabIndex        =   5
                  Top             =   480
                  Width           =   3675
               End
               Begin VB.TextBox txtCodEquip 
                  Appearance      =   0  'Flat
                  Height          =   285
                  Left            =   3180
                  TabIndex        =   89
                  Top             =   180
                  Visible         =   0   'False
                  Width           =   615
               End
               Begin VB.TextBox txtQuantAlugada 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
                  Height          =   315
                  Left            =   4800
                  TabIndex        =   7
                  Top             =   480
                  Width           =   615
               End
               Begin VB.TextBox txtTotalAluguel 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   5460
                  Locked          =   -1  'True
                  TabIndex        =   8
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   975
               End
               Begin VB.Label lbl6 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Item:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   93
                  Top             =   240
                  Width           =   345
               End
               Begin VB.Label lbl7 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Valor Diária"
                  Height          =   195
                  Left            =   3840
                  TabIndex        =   92
                  Top             =   240
                  Width           =   810
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Quant."
                  Height          =   195
                  Left            =   4800
                  TabIndex        =   91
                  Top             =   240
                  Width           =   480
               End
               Begin VB.Label Label20 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Total"
                  Height          =   195
                  Left            =   5460
                  TabIndex        =   90
                  Top             =   240
                  Width           =   360
               End
            End
            Begin MSFlexGridLib.MSFlexGrid GridProdutos 
               Height          =   1935
               Left            =   120
               TabIndex        =   25
               Top             =   2520
               Width           =   12375
               _ExtentX        =   21828
               _ExtentY        =   3413
               _Version        =   393216
               SelectionMode   =   1
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin ChamaleonBtn.chameleonButton cmdAdicionarProduto 
               Height          =   315
               Left            =   120
               TabIndex        =   24
               ToolTipText     =   "Adiciona"
               Top             =   2160
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "Adicionar"
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
               FCOL            =   16384
               FCOLO           =   16384
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "Aluguel_Cadastro.frx":19A28
               PICN            =   "Aluguel_Cadastro.frx":19A44
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdRemoverProduto 
               Height          =   315
               Left            =   1380
               TabIndex        =   26
               ToolTipText     =   "Remove"
               Top             =   2160
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "Remover"
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
               MICON           =   "Aluguel_Cadastro.frx":19DDE
               PICN            =   "Aluguel_Cadastro.frx":19DFA
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdProrrogar 
               Height          =   315
               Left            =   1920
               TabIndex        =   104
               TabStop         =   0   'False
               ToolTipText     =   "Remove"
               Top             =   4500
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "Prorrogar Entrega >>"
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
               MICON           =   "Aluguel_Cadastro.frx":1A194
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdDevolver 
               Height          =   315
               Left            =   3720
               TabIndex        =   105
               TabStop         =   0   'False
               ToolTipText     =   "Remove"
               Top             =   4500
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "Devolver Integral"
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
               MICON           =   "Aluguel_Cadastro.frx":1A1B0
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdDevolverParcial 
               Height          =   315
               Left            =   5520
               TabIndex        =   106
               TabStop         =   0   'False
               ToolTipText     =   "Remove"
               Top             =   4500
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "Devolver Parcial"
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
               MICON           =   "Aluguel_Cadastro.frx":1A1CC
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonBtn.chameleonButton cmdAdiar 
               Height          =   315
               Left            =   120
               TabIndex        =   130
               TabStop         =   0   'False
               ToolTipText     =   "Remove"
               Top             =   4500
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "<< Antecipar Entrega"
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
               MICON           =   "Aluguel_Cadastro.frx":1A1E8
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label lblSomaReferente 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
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
               Left            =   11580
               TabIndex        =   47
               Top             =   4560
               Width           =   915
            End
         End
         Begin MSFlexGridLib.MSFlexGrid grid_Parcelas 
            Height          =   1755
            Left            =   180
            TabIndex        =   123
            TabStop         =   0   'False
            Top             =   6000
            Width           =   6195
            _ExtentX        =   10927
            _ExtentY        =   3096
            _Version        =   393216
            BackColor       =   12648447
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin VB.Frame frmDevolucao 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Devolução"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   855
            Left            =   6420
            TabIndex        =   99
            Top             =   6660
            Visible         =   0   'False
            Width           =   6195
            Begin VB.TextBox txtQuantDev 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3840
               TabIndex        =   36
               Top             =   480
               Width           =   1035
            End
            Begin ChamaleonBtn.chameleonButton chameleonButton3 
               Height          =   315
               Left            =   2760
               TabIndex        =   34
               TabStop         =   0   'False
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
               MICON           =   "Aluguel_Cadastro.frx":1A204
               PICN            =   "Aluguel_Cadastro.frx":1A220
               PICH            =   "Aluguel_Cadastro.frx":1C573
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSMask.MaskEdBox mskDevolver 
               Height          =   315
               Left            =   1860
               TabIndex        =   33
               Top             =   480
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin ChamaleonBtn.chameleonButton cmdDevolverItem 
               Height          =   315
               Left            =   4980
               TabIndex        =   37
               ToolTipText     =   "Remove"
               Top             =   480
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               BTYPE           =   3
               TX              =   "Devolver"
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
               MICON           =   "Aluguel_Cadastro.frx":1E8C6
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSMask.MaskEdBox mskDataFinalLocacao 
               Height          =   315
               Left            =   120
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   480
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskDevolverHora 
               Height          =   315
               Left            =   3120
               TabIndex        =   35
               Top             =   480
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskDataFinalLocacaoHora 
               Height          =   315
               Left            =   1140
               TabIndex        =   32
               Top             =   480
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Data Final"
               Height          =   195
               Left            =   120
               TabIndex        =   108
               Top             =   240
               Width           =   720
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Devolução"
               Height          =   195
               Left            =   1860
               TabIndex        =   107
               Top             =   240
               Width           =   780
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Quant. Devol."
               Height          =   195
               Left            =   3840
               TabIndex        =   100
               Top             =   240
               Width           =   990
            End
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdCadastrarCliente 
         Height          =   615
         Left            =   12840
         TabIndex        =   53
         Top             =   3720
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Clientes"
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
         MICON           =   "Aluguel_Cadastro.frx":1E8E2
         PICN            =   "Aluguel_Cadastro.frx":1E8FE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSFlexGridLib.MSFlexGrid GridConsulta 
         Height          =   6195
         Left            =   -74940
         TabIndex        =   57
         Top             =   1440
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   10927
         _Version        =   393216
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelar 
         Height          =   615
         Left            =   12840
         TabIndex        =   28
         Top             =   1740
         Width           =   1635
         _ExtentX        =   2884
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
         MICON           =   "Aluguel_Cadastro.frx":1F1D8
         PICN            =   "Aluguel_Cadastro.frx":1F1F4
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
         Left            =   12840
         TabIndex        =   29
         Top             =   2400
         Width           =   1635
         _ExtentX        =   2884
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
         MICON           =   "Aluguel_Cadastro.frx":20F86
         PICN            =   "Aluguel_Cadastro.frx":20FA2
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
         Left            =   12840
         TabIndex        =   30
         Top             =   3060
         Width           =   1635
         _ExtentX        =   2884
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
         MICON           =   "Aluguel_Cadastro.frx":22D34
         PICN            =   "Aluguel_Cadastro.frx":22D50
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
         Left            =   12840
         TabIndex        =   27
         Top             =   1080
         Width           =   1635
         _ExtentX        =   2884
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
         MICON           =   "Aluguel_Cadastro.frx":24AE2
         PICN            =   "Aluguel_Cadastro.frx":24AFE
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
         Left            =   12840
         TabIndex        =   0
         Top             =   420
         Width           =   1635
         _ExtentX        =   2884
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
         MICON           =   "Aluguel_Cadastro.frx":26890
         PICN            =   "Aluguel_Cadastro.frx":268AC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdImprimirContato 
         Height          =   615
         Left            =   12840
         TabIndex        =   85
         Top             =   4380
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Contrato"
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
         MICON           =   "Aluguel_Cadastro.frx":2863E
         PICN            =   "Aluguel_Cadastro.frx":2865A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdFecharAluguel 
         Height          =   315
         Left            =   -74940
         TabIndex        =   144
         Top             =   7680
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Fechar Aluguel"
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
         MICON           =   "Aluguel_Cadastro.frx":28974
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblCodPedido 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   13080
         TabIndex        =   86
         Top             =   5160
         Width           =   1155
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Atual:"
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
         Left            =   -63660
         TabIndex        =   59
         Top             =   8040
         Width           =   1755
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
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
         Left            =   -63660
         TabIndex        =   58
         Top             =   7740
         Width           =   1755
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   61
      Top             =   9315
      Width           =   14730
      _ExtentX        =   25982
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   14896
            Text            =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
            TextSave        =   "Desenv.: Online.Info Sistemas - Tel.: (89) 9 8817-7036"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "06:02"
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
Attribute VB_Name = "Aluguel_Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moCombo As cComboHelper

Dim i As Integer
Dim UltimoParcela As Integer
Dim varCodItem As Long                  'item da os
Dim vQuantDias As Integer
Dim date1 As Date
Dim date2 As Date

Dim sSQL As String
Dim r As ADODB.Recordset
Dim printSQL As String

Dim vQuantDiasAdiar As Integer
Dim vTipoDevolver As Integer

Dim vQuantHoras As Integer              'calcular as horas acionadas apos termino
Dim vValorTotalHoras As Currency        'calcular as horas acionadas apos termino
Dim vSituacao As String
Dim vValorParcelaDevolucao As Currency    'valor em aberto da parcela do item sendo devolvido



Private Sub CalcularDiaria()
Dim varQtdeDias As Integer
Dim varValorAluguel As Currency
Dim varValorTotal As Currency
Dim varDesc As Currency
Dim varSubtotal As Currency

If txtQuant.Text = "" Then Exit Sub
If txtTotalAluguel.Text = "" Then Exit Sub
If txtDesconto.Text = "" Then varDesc = 0: txtDesconto.Text = Format(0, "##,##0.00") Else varDesc = txtDesconto.Text

varQtdeDias = txtQuant.Text
varValorAluguel = txtTotalAluguel.Text

varSubtotal = varValorAluguel * varQtdeDias
txtValor.Text = Format(varSubtotal, "##,##0.00")

varSubtotal = txtValor.Text
If txtValor.Text = "" Then Exit Sub
varValorTotal = varSubtotal - varDesc
txtSubTotal = Format(varValorTotal, "##,##0.00")

Dim varEntrada As Double
If txtEntrada.Text = "" Then varEntrada = 0: txtEntrada.Text = Format(0, "##,##0.00") Else varEntrada = txtEntrada.Text

txtEntradaReal.Text = Format(((CCur(varValorTotal) * CCur(varEntrada)) / 100), ocMONEY)
txtTotal.Text = Format(CCur(varValorTotal) - ((CCur(varValorTotal) * CCur(varEntrada)) / 100), ocMONEY)
End Sub


Private Sub calculardiasAdiar()
Dim Result As Integer

If Not IsDate(mskDataInicialLocacaoProrro) Then Exit Sub
If Not IsDate(mskDataProrrogar) Then Exit Sub

date1 = CDate(mskDataInicialLocacaoProrro.Text)
date2 = CDate(mskDataProrrogar.Text)

Result = DateDiff("d", date1, date2)

If date1 = date2 Then
    Result = Result + 1
Else
    Result = Result
End If

txtDiasProrrogar.Text = Result
End Sub



Private Sub CalcularQuantDias()
Dim date1 As Date
Dim date2 As Date
Dim Result As Long

If txtTipoCobranca.Text = "HORA" Then
    If Not IsDate(mskInicio) Then Exit Sub
    If Not IsDate(mskFinal) Then Exit Sub
    If mskHoraInicio.Text = "" Then Exit Sub
    If mskHoraFinal.Text = "" Then Exit Sub

    date1 = CDate(mskInicio.Text & " " & mskHoraInicio.Text)
    date2 = CDate(mskFinal.Text & " " & mskHoraFinal.Text)

    Result = DateDiff("h", date1, date2)
    If Result < 1 Then Result = 1

    txtQuant.Text = Result
Else
    If Not IsDate(mskInicio) Then Exit Sub
    If Not IsDate(mskFinal) Then Exit Sub

    date1 = CDate(mskInicio.Text)
    date2 = CDate(mskFinal.Text)

    Result = DateDiff("d", date1, date2)

    If mskInicio.Text = mskFinal.Text Then
        Result = 1
    End If

    txtQuant.Text = Result
End If
End Sub

Private Sub CalcularTotalAdiar()
Dim vQuantDiasAdiarTudo As Integer
If txtDiasProrrogar.Text = "" Then Exit Sub

vQuantDiasAdiarTudo = txtDiasProrrogar.Text

'txtValorAluguellProrro
'txtDiasProrrogar
'txtTotalAluguelProrro
'txtDescAluguelProrro
'txtTotalAluguelDescProrro

'calcular o total
Dim vValorDiariaAdiar As Currency
Dim vDescAdiar As Double
Dim vDescInicial As Double
Dim vSubtotalAdiar As Currency

If txtValorAluguellProrro.Text = "" Then txtValorAluguellProrro.Text = "0"
If txtDiasProrrogar.Text = "" Then txtDiasProrrogar.Text = "0"
If txtDescProrro.Text = "" Then txtDescProrro.Text = "0"

txtValorAluguellProrro.Text = Format(GridProdutos.TextMatrix(i, 5), "##,##0.00")
vDescInicial = GridProdutos.TextMatrix(i, 8)
txtDescAluguelProrro = Format(vDescInicial, "##,##0.00")

vValorDiariaAdiar = GridProdutos.TextMatrix(i, 5)
vDescAdiar = txtDescProrro.Text

vSubtotalAdiar = vValorDiariaAdiar * vQuantDiasAdiarTudo
txtTotalAluguelProrro = Format(vSubtotalAdiar, "##,##0.00")
vSubtotalAdiar = vSubtotalAdiar - vDescInicial
txtTotalAluguelDescProrro = Format(vSubtotalAdiar, "##,##0.00")
vSubtotalAdiar = vSubtotalAdiar - vDescAdiar

'saber o valo da nova parcela desncontando a entrada
Dim vValorNovaParcela As Currency
Dim vValorEntradaProrro As Currency

If txtEntradaAdiar.Text = "" Then txtEntradaAdiar.Text = "0"
vValorEntradaProrro = txtEntradaAdiar.Text
vValorNovaParcela = vSubtotalAdiar - vValorEntradaProrro
txtTotalProrro.Text = Format(vValorNovaParcela, "##,##0.00")
End Sub

Private Sub CalcularValorTotal()
Dim varValorAluguel As Currency
Dim varQuantAlugada As Integer
Dim varTotalAluguel As Currency

If txtValorAluguel.Text = "" Then Exit Sub
If txtQuantAlugada.Text = "" Then txtQuantAlugada.Text = "1"

varValorAluguel = txtValorAluguel.Text
varQuantAlugada = txtQuantAlugada.Text
varTotalAluguel = varValorAluguel * varQuantAlugada
txtTotalAluguel = Format(varTotalAluguel, ocMONEY)
End Sub


Private Sub ConsultarUltimaParcela()
i = GridProdutos.Row

If lblCodPedido.Caption = "" Then Exit Sub
Dim vItem As Integer

If GridProdutos.rows >= 2 Then
    vItem = GridProdutos.TextMatrix(i, 1)
Else
    vItem = 1
End If

sSQL = "SELECT ISNULL(MAX(numero), 0) as UltimoNumero " & _
        "FROM parcelas WHERE (cod_pedido = " & lblCodPedido.Caption & ") and (OS_ITEM = " & vItem & ");"

Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then UltimoParcela = r("UltimoNumero")

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub



Sub FormatarGrid(rTabela As ADODB.Recordset)
'Dim i As Integer
Dim j As Integer

With GridConsulta
   .Clear
   .Cols = 11
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 800
   .ColWidth(2) = 3800
   .ColWidth(3) = 2700
   .ColWidth(4) = 1300
   .ColWidth(5) = 1200
   .ColWidth(6) = 1200
   .ColWidth(7) = 1200
   .ColWidth(8) = 900
   .ColWidth(9) = 0
   .ColWidth(10) = 700
   
   .TextMatrix(0, 1) = "CONT."
   .TextMatrix(0, 2) = "CLIENTE"
   .TextMatrix(0, 3) = "OBRA"
   .TextMatrix(0, 4) = "CIDADE"
   .TextMatrix(0, 5) = "VLR INICIAL"
   .TextMatrix(0, 6) = "VLR ATUAL"
   .TextMatrix(0, 7) = "EM ABERTO"
   .TextMatrix(0, 8) = "STATUS"
   .TextMatrix(0, 9) = "EXCLUIDO"
   .TextMatrix(0, 10) = "EQUIP"

   .Redraw = False
   
   'colocar os cabeçalho em negrito
   For i = 0 To .Cols - 1
      .Col = i
      .Row = 0
      .CellFontBold = True
   Next
   
   'ALINHAMENTO
   '.ColAlignment(2) = 1
   
   'Centralizar o titulo
   For i = 0 To .Cols - 1
      .Row = 0
      .Col = i
      .CellAlignment = flexAlignCenterCenter
   Next
   
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         .TextMatrix(.rows - 1, 1) = Format(rTabela("varCodAluguel"), "000000")
         .TextMatrix(.rows - 1, 2) = rTabela("NOME")
         .TextMatrix(.rows - 1, 3) = ValidateNull(rTabela("OBRA"))
         .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("CIDADEOBRA"))
         .TextMatrix(.rows - 1, 5) = Format(ValidateNull(rTabela("VALORINICIAL")), ocMONEY)
         .TextMatrix(.rows - 1, 6) = Format(ValidateNull(rTabela("varSomaParcelasTodas")), ocMONEY)
         .TextMatrix(.rows - 1, 7) = Format(ValidateNull(rTabela("varSomaParcelasAbertas")), ocMONEY)
         .TextMatrix(.rows - 1, 8) = rTabela("VAR_STATUS")
         .TextMatrix(.rows - 1, 9) = rTabela("varexcluido")
         .TextMatrix(.rows - 1, 10) = rTabela("varQuantItens")

         
         rTabela.MoveNext
         .rows = .rows + 1
      Loop
   End If
   
    For i = 1 To .rows - 1
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
   
   .Redraw = True
   .rows = .rows - 1

End With

txtCONtotal.Text = Format(SomaGrid(GridConsulta, 6), ocMONEY)
End Sub

Private Sub LimparObjetos_Prorrogacao()
mskDataInicialLocacaoProrro.Mask = ""
mskDataInicialLocacaoProrro.Text = ""
mskDataFinalLocacaoProrro.Mask = ""
mskDataFinalLocacaoProrro.Text = ""
mskDataInicioProrro.Mask = ""
mskDataInicioProrro.Text = ""
mskDataProrrogar.Mask = ""
mskDataProrrogar.Text = ""
txtDiasProrrogar.Text = ""
txtValorAluguellProrro.Text = ""
txtDescProrro.Text = ""
txtTotalProrro.Text = ""
txtValorRealDescProrro.Text = ""
End Sub

Private Sub LimparObjetosDevolucao()
mskDataFinalLocacao.Mask = ""
mskDataFinalLocacao.Text = ""
mskDevolver.Mask = ""
mskDevolver.Text = ""
txtQuantDev.Text = ""
End Sub

Private Sub PreencherFormaPgtoDevolucao(cbo As ComboBox)
cbo.Clear
cbo.AddItem "1 - DINHEIRO"
cbo.AddItem "3 - CARTÃO - DÉBITO"
cbo.AddItem "4 - CARTÃO - CRÉDITO"
cbo.AddItem "5 - CHEQUE"
cbo.AddItem "7 - TRANSFERÊNCIA"
cbo.AddItem "8 - DEPOSITO"
cbo.AddItem "9 - FINANCEIRA"
cbo.AddItem "10 - PIX"
End Sub

Private Sub AtualizarVisibilidadeFormaPgto()
Dim vMostrarForma2 As Boolean
vMostrarForma2 = (opt2forma.Value = True)

lblformaPgto2.Visible = vMostrarForma2
cboformaPgto2.Visible = vMostrarForma2
lblValor2.Visible = vMostrarForma2
txtValor2.Visible = vMostrarForma2

If vMostrarForma2 Then
    lblDesc.Left = 5220
    txtDesc.Left = 5220
Else
    lblDesc.Left = 2760
    txtDesc.Left = 2760
End If
End Sub

Private Sub AtualizarPagamentoDevolucao()
Dim rPgto As ADODB.Recordset

i = GridProdutos.Row
If lblCodigo.Caption = "" Or GridProdutos.TextMatrix(i, 1) = "" Then Exit Sub

sSQL = "SELECT ISNULL(SUM(VALOR_FINAL), 0) AS vSoma, ISNULL(SUM(DESCONTO), 0) AS vSomaDesc FROM parcelas WHERE (cod_os = " & lblCodigo.Caption & ") and OS_ITEM = " & GridProdutos.TextMatrix(i, 1) & " and STATUS = 0;"
Set rPgto = dbData.OpenRecordset(sSQL)

If Not rPgto.BOF Then
    vValorParcelaDevolucao = rPgto("vSoma")
    txtDesc.Text = Format(rPgto("vSomaDesc"), "##,##0.00")
Else
    vValorParcelaDevolucao = 0
    txtDesc.Text = Format(0, "##,##0.00")
End If
If rPgto.State <> 0 Then rPgto.Close
Set rPgto = Nothing

AtualizarVisibilidadeFormaPgto

If opt2forma.Value = True Then
    txtValor1.Text = Format(0, "##,##0.00")
    txtValor2.Text = Format(vValorParcelaDevolucao, "##,##0.00")
Else
    txtValor1.Text = Format(vValorParcelaDevolucao, "##,##0.00")
    txtValor2.Text = Format(0, "##,##0.00")
    If chkPgtoAuto.Value = vbChecked Then
        PreencherFormaPgtoDevolucao cboformaPgto1
        cboformaPgto1.Text = "10 - PIX"
    End If
End If
End Sub

Private Sub chkPgtoAuto_Click()
Dim vHabilitado As Boolean
vHabilitado = (chkPgtoAuto.Value = vbChecked)

opt1forma.Enabled = vHabilitado
opt2forma.Enabled = vHabilitado
cboformaPgto1.Enabled = vHabilitado
cboformaPgto2.Enabled = vHabilitado
txtValor1.Enabled = vHabilitado
txtValor2.Enabled = vHabilitado
txtDesc.Enabled = vHabilitado

If vHabilitado Then AtualizarPagamentoDevolucao
End Sub

Private Sub opt1forma_Click()
AtualizarPagamentoDevolucao
End Sub

Private Sub opt2forma_Click()
AtualizarPagamentoDevolucao
End Sub

Private Sub cboformaPgto1_GotFocus()
Dim varTexto As String
varTexto = cboformaPgto1.Text
PreencherFormaPgtoDevolucao cboformaPgto1
cboformaPgto1.Text = varTexto
moCombo.AttachTo cboformaPgto1
End Sub

Private Sub cboformaPgto2_GotFocus()
Dim varTexto As String
varTexto = cboformaPgto2.Text
PreencherFormaPgtoDevolucao cboformaPgto2
cboformaPgto2.Text = varTexto
moCombo.AttachTo cboformaPgto2
End Sub

Private Sub txtValor1_Change()
If opt2forma.Value = True Then
    Dim vValor1Atual As Currency
    vValor1Atual = Val(Replace(Replace(txtValor1.Text, ".", ""), ",", "."))

    Dim vValor2Calc As Currency
    vValor2Calc = vValorParcelaDevolucao - vValor1Atual
    If vValor2Calc < 0 Then vValor2Calc = 0

    txtValor2.Text = Format(vValor2Calc, "##,##0.00")
End If
End Sub

Private Sub txtValor1_LostFocus()
Dim vValor1Atual As Currency
If txtValor1.Text = "" Then
    vValor1Atual = 0
Else
    vValor1Atual = Val(Replace(Replace(txtValor1.Text, ".", ""), ",", "."))
End If

If vValor1Atual < 0 Then vValor1Atual = 0
If vValor1Atual > vValorParcelaDevolucao Then vValor1Atual = vValorParcelaDevolucao

txtValor1.Text = Format(vValor1Atual, "##,##0.00")

If opt2forma.Value = True Then
    txtValor2.Text = Format(vValorParcelaDevolucao - vValor1Atual, "##,##0.00")
End If
End Sub

Public Function SomaGridItens(var_Grid As MSFlexGrid, Col As Integer) As Currency
'Dim i As Integer
Dim Valor As Currency

Valor = 0
For i = 0 To var_Grid.rows - 1
   If IsNumeric(var_Grid.TextMatrix(i, Col)) Then
     'If var_Grid.TextMatrix(i, 15) = "NÃO" Then
         Valor = Valor + CDbl(var_Grid.TextMatrix(i, Col))
     'End If
   End If
Next

SomaGridItens = Valor
End Function

Public Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Currency
'Dim i As Integer
Dim Valor As Currency

Valor = 0
For i = 0 To var_Grid.rows - 1
   If IsNumeric(var_Grid.TextMatrix(i, Col)) Then
      Valor = Valor + CDbl(var_Grid.TextMatrix(i, Col))
   End If
Next

SomaGrid = Valor
End Function

Private Sub FormatarGridProdutos(rTabela As ADODB.Recordset)
'Dim i As Integer
Dim j As Integer

With GridProdutos
   .Clear
   .Cols = 20
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 300
   .ColWidth(2) = 2900
   .ColWidth(3) = 800
   .ColWidth(4) = 450
   .ColWidth(5) = 800
   .ColWidth(6) = 500
   .ColWidth(7) = 800
   .ColWidth(8) = 630
   .ColWidth(9) = 800
   .ColWidth(10) = 800
   .ColWidth(11) = 800
   .ColWidth(12) = 0   'esse
   .ColWidth(13) = 0   'esse

   .ColWidth(14) = 800
   .ColWidth(15) = 500
   .ColWidth(16) = 800
   .ColWidth(17) = 500
   .ColWidth(18) = 500  'esse
   .ColWidth(19) = 500  'esse


   .TextMatrix(0, 1) = "ITEM"
   .TextMatrix(0, 2) = "EQUIPAMENTO"
   .TextMatrix(0, 3) = "R$"
   .TextMatrix(0, 4) = "QTD"
   .TextMatrix(0, 5) = "VALOR"
   .TextMatrix(0, 6) = "DIAS"
   .TextMatrix(0, 7) = "TOTAL"
   .TextMatrix(0, 8) = "DESC"
   .TextMatrix(0, 9) = "SUBTOTAL"
   .TextMatrix(0, 10) = "ENTRADA"
   .TextMatrix(0, 11) = "RESTANTE"
   .TextMatrix(0, 12) = "EQUIP"
   .TextMatrix(0, 13) = "DEV"
   .TextMatrix(0, 14) = "INICIO"
   .TextMatrix(0, 15) = "HORA"
   .TextMatrix(0, 16) = "FINAL"
   .TextMatrix(0, 17) = "HORA"
   .TextMatrix(0, 18) = "excluido"
   .TextMatrix(0, 19) = "Prorrogado"

   
   'colocar os cabeçalho em negrito
   For i = 0 To .Cols - 1
      .Col = i
      .Row = 0
      .CellFontBold = True
   Next i

   .Redraw = False
   
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
        
         .TextMatrix(.rows - 1, 1) = rTabela("ITEM")
         .TextMatrix(.rows - 1, 2) = ValidateNull(rTabela("descricao")) & " /  " & ValidateNull(rTabela("MODELO")) & " / " & ValidateNull(rTabela("FABRICANTE"))
         .TextMatrix(.rows - 1, 3) = Format(rTabela("VALOR_UND"), ocMONEY)
         .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela("QUANT_ALUGADA"))
         .TextMatrix(.rows - 1, 5) = Format(rTabela("TOTAL_ALUGADA"), ocMONEY)
         .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("QUANT"))
         .TextMatrix(.rows - 1, 7) = Format(rTabela("VALOR"), ocMONEY)
         .TextMatrix(.rows - 1, 8) = Format(rTabela("DESCONTO"), ocMONEY)
         .TextMatrix(.rows - 1, 9) = Format(rTabela("SUBTOTAL"), ocMONEY)
         .TextMatrix(.rows - 1, 10) = Format(rTabela("ENTRADA"), ocMONEY)
         .TextMatrix(.rows - 1, 11) = Format(rTabela("VALOR_FINAL"), ocMONEY)
         .TextMatrix(.rows - 1, 12) = ValidateNull(rTabela("COD_EQUIP"))
         .TextMatrix(.rows - 1, 13) = ValidateNull(rTabela("varDevolvido"))
         
         .TextMatrix(.rows - 1, 14) = Format(rTabela("DATA_INICIO"), ocDATA2)
         .TextMatrix(.rows - 1, 15) = Format(rTabela("HORA_INICIO"), ocHRMN)
         .TextMatrix(.rows - 1, 16) = Format(rTabela("DATA_FINAL"), ocDATA2)
         .TextMatrix(.rows - 1, 17) = Format(rTabela("HORA_FINAL"), ocHRMN)
         .TextMatrix(.rows - 1, 18) = ValidateNull(rTabela("varExcluido"))
         .TextMatrix(.rows - 1, 19) = ValidateNull(rTabela("varprorrogado"))
         
         rTabela.MoveNext
         .rows = .rows + 1
      Loop
   End If
   

   
   'And .TextMatrix(i, 15) = "NÃO" Then
    '
    '      ElseIf .TextMatrix(i, 14) = "NÃO" And .TextMatrix(i, 15) = "SIM" Then
    '         .CellForeColor = vbRed
   
    For i = 1 To .rows - 1
       For j = 0 To .Cols - 1
          .Col = j
          .Row = i
    
          If .TextMatrix(i, 13) = "SIM" Then
             .CellForeColor = vbBlue
          ElseIf .TextMatrix(i, 18) = "SIM" Then
             .CellForeColor = vbRed
          ElseIf .TextMatrix(i, 19) = "SIM" Then
             .CellForeColor = &H8000&
          Else
             .CellForeColor = vbBlack
          End If
          
       Next
    Next
   .rows = .rows - 1
   .Redraw = True
End With

Dim soma As Currency
soma = 0

With GridProdutos
   For i = 1 To .rows - 1
      If .TextMatrix(i, 18) = "NÃO" Then
         soma = soma + CCur(.TextMatrix(i, 9))
      End If
   Next
End With

lblSomaReferente.Caption = Format(soma, ocMONEY)

End Sub

Private Sub LimparGridProdutos()
'Dim i As Integer

With GridProdutos
   .Clear
   .Cols = 16
   .rows = 1
   
   .ColWidth(0) = 0
   .ColWidth(1) = 0
   .ColWidth(2) = 3150
   .ColWidth(3) = 700
   .ColWidth(4) = 800
   .ColWidth(5) = 850
   .ColWidth(6) = 630
   .ColWidth(7) = 850
   .ColWidth(8) = 630
   .ColWidth(9) = 650
   .ColWidth(10) = 850
   .ColWidth(11) = 700
   .ColWidth(12) = 850
   .ColWidth(13) = 0
   .ColWidth(14) = 0
   .ColWidth(15) = 0
   
   .TextMatrix(0, 1) = "ITEM"
   .TextMatrix(0, 2) = "EQUIPAMENTO"
   .TextMatrix(0, 3) = "TIPO"
   .TextMatrix(0, 4) = "VALOR"
   .TextMatrix(0, 5) = "INICIO"
   .TextMatrix(0, 6) = "HORA"
   .TextMatrix(0, 7) = "FINAL"
   .TextMatrix(0, 8) = "HORA"
   .TextMatrix(0, 9) = "QTDE"
   .TextMatrix(0, 10) = "SUBT"
   .TextMatrix(0, 11) = "DESC"
   .TextMatrix(0, 12) = "TOTAL"
   .TextMatrix(0, 13) = "cod_equip"
   .TextMatrix(0, 14) = "DEV"
   .TextMatrix(0, 15) = "EXC"
   
   'colocar os cabeçalho em negrito
   For i = 0 To .Cols - 1
      .Col = i
      .Row = 0
      .CellFontBold = True
   Next
   
   .ColAlignment(1) = 1
   .ColAlignment(2) = 2
End With
End Sub

Private Sub PreencherComboStatus()
cboCONStatus.Clear
cboCONStatus.AddItem "TODOS"
cboCONStatus.AddItem "ABERTO"
cboCONStatus.AddItem "FECHADO"

If cboCONStatus.Text = "" Then cboCONStatus.ListIndex = 1
moCombo.AttachTo cboCONStatus
End Sub

Private Sub PreencherGridProdutos()
'    sSQL = "SELECT *, (CASE WHEN excluido = 1 THEN 'SIM' ELSE 'NÃO' END) as varExcluido, (CASE WHEN devolvido = 1 THEN 'SIM' ELSE 'NÃO' END) as varDevolvido FROM aluguel_cadastro_itens INNER JOIN aluguel_cadastro_equipamento ON aluguel_cadastro_itens.COD_EQUIP = aluguel_cadastro_equipamento.COD_EQUIP WHERE (COD_LOCACAO = " & lblCodigo.Caption & ");"
If txtTipoCobranca.Text = "DIA" Then
    sSQL = "SELECT Aluguel_Cadastro_Equipamento.DESCRICAO, Aluguel_Cadastro_Equipamento.FABRICANTE, Aluguel_Cadastro_Equipamento.MODELO, COD_LOCACAO, ITEM, TIPO_LOCACAO, aluguel_cadastro_itens.COD_EQUIP, aluguel_cadastro_itens.ENTRADA, aluguel_cadastro_itens.VALOR, VALOR_UND, aluguel_cadastro_itens.QUANT_ALUGADA, aluguel_cadastro_itens.TOTAL_ALUGADA, DATA_INICIO, HORA_INICIO, DATA_FINAL, HORA_FINAL, QUANT, VALOR_FINAL, DESCONTO, SUBTOTAL, (CASE WHEN aluguel_cadastro_itens.devolvido = 1 THEN 'SIM' ELSE 'NÃO' END) as varDevolvido, (CASE WHEN aluguel_cadastro_itens.excluido = 1 THEN 'SIM' ELSE 'NÃO' END) as varExcluido, (CASE WHEN aluguel_cadastro_itens.PRORROGADO = 1 THEN 'SIM' ELSE 'NÃO' END) as varProrrogado FROM aluguel_cadastro_itens INNER JOIN aluguel_cadastro_equipamento ON aluguel_cadastro_itens.COD_EQUIP = aluguel_cadastro_equipamento.COD_EQUIP WHERE (COD_LOCACAO = " & lblCodigo.Caption & ") ORDER BY ITEM ;"
Else
    sSQL = "SELECT Aluguel_Cadastro_Equipamento.DESCRICAO, Aluguel_Cadastro_Equipamento.FABRICANTE, Aluguel_Cadastro_Equipamento.MODELO, COD_LOCACAO, ITEM, TIPO_LOCACAO, aluguel_cadastro_itens.COD_EQUIP, aluguel_cadastro_itens.ENTRADA, aluguel_cadastro_itens.VALOR, VALOR_UND, aluguel_cadastro_itens.QUANT_ALUGADA, aluguel_cadastro_itens.TOTAL_ALUGADA, DATA_INICIO, HORA_INICIO, DATA_FINAL, HORA_FINAL, QUANT, VALOR_FINAL, DESCONTO, SUBTOTAL, (CASE WHEN aluguel_cadastro_itens.devolvido = 1 THEN 'SIM' ELSE 'NÃO' END) as varDevolvido, (CASE WHEN aluguel_cadastro_itens.excluido = 1 THEN 'SIM' ELSE 'NÃO' END) as varExcluido, (CASE WHEN aluguel_cadastro_itens.PRORROGADO = 1 THEN 'SIM' ELSE 'NÃO' END) as varProrrogado FROM aluguel_cadastro_itens INNER JOIN aluguel_cadastro_equipamento ON aluguel_cadastro_itens.COD_EQUIP = aluguel_cadastro_equipamento.COD_EQUIP WHERE (COD_LOCACAO = " & lblCodigo.Caption & ") ORDER BY ITEM ;"
End If
'Debug.Print sSQL
Set r = dbData.OpenRecordset(sSQL)

FormatarGridProdutos r
If r.State <> 0 Then r.Close
End Sub


Private Function Inserir_Dados() As Boolean
'Dim sSQL As String

'Comando de inclusão
sSQL = "INSERT INTO a_receber (" & _
   "cod_receber, cod_cliente, cliente) VALUES ("

sSQL = sSQL & _
   lblCodigo.Caption & ", " & txtCodCliente.Text & ", '" & txtCliente.Text & "')"

'Retorna o resultado da inclusão
Inserir_Dados = dbData.Execute(sSQL)
End Function

Private Sub AutoNumeracao_Aluguel()
'Dim sSQL As String
'Dim r As ADODB.Recordset

sSQL = "SELECT ISNULL(MAX(CODIGO), 0) AS ultimo FROM aluguel_cadastro;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then lblCodigo.Caption = Format(r("ultimo") + 1, "000000")
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub



Private Sub Limpar_Objetos()
cboSituacao.Text = ""
txtCodCliente.Text = ""
txtCliente.Text = ""
txtDescricaoObra.Text = ""
cboCidadeObra.Text = ""
lblCodigo.Caption = ""
lblCodPedido.Caption = ""
lblSomaReferente.Caption = ""
frmCliente.Enabled = False
frmReferente.Enabled = False
LimparGridProdutos
End Sub

Private Sub cboAno_GotFocus()
Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
'Dim i As Integer
  
cboAno.Clear

iAno = Year(Date)
FirstYear = iAno - 2
LastYear = iAno + 2

For i = LastYear To FirstYear Step -1
   cboAno.AddItem i
Next

'For i = iAno To FirstYear Step -1
'   cboAno.AddItem i
'Next
'
'iAno = iAno + 1
'For i = iAno To LastYear
'   cboAno.AddItem i
'Next
End Sub

Private Sub cboCidadeObra_GotFocus()
'Dim sSQL As String
'Dim r As ADODB.Recordset

cboCidadeObra.Clear

sSQL = "SELECT DISTINCT CIDADEOBRA FROM aluguel_cadastro ORDER BY CIDADEOBRA;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboCidadeObra.AddItem ValidateNull(r("CIDADEOBRA"))
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

moCombo.AttachTo cboCidadeObra
End Sub


Private Sub cboCidadeObra_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cboCONStatus_Change()
   cboCONStatus_Click
End Sub

Private Sub cboCONStatus_Click()
   'cmdExibirConsulta_Click 'desativei
End Sub

Private Sub cboCONStatus_GotFocus()
   PreencherComboStatus
End Sub

Private Sub cboCONStatus_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cboFiltro_Change()
cboFiltro_LostFocus
End Sub

Private Sub cboFiltro_Click()
cboFiltro_LostFocus
End Sub


Private Sub cboFiltro_GotFocus()
cboFiltro.Clear
cboFiltro.AddItem "TODOS"
cboFiltro.AddItem "MÊS"
cboFiltro.AddItem "PERIODO"
cboFiltro.AddItem "CLIENTE"
moCombo.AttachTo cboFiltro
End Sub


Private Sub cboFiltro_LostFocus()
If cboFiltro.Text = "TODOS" Then
    lblCONmes.Visible = False
    cboMes.Visible = False
    cboAno.Visible = False
    lblCONint1.Visible = False
    Mask1.Visible = False
    lblCONint2.Visible = False
    Mask2.Visible = False
    cmdInicial.Visible = False
    cmdFinal.Visible = False
    lblCONnome.Visible = False
    cboNome.Visible = False
    'cmdExibirConsulta_Click
ElseIf cboFiltro.Text = "MÊS" Then
   lblCONmes.Visible = True
   cboMes.Visible = True
   cboAno.Visible = True
   lblCONint1.Visible = False
   Mask1.Visible = False
   lblCONint2.Visible = False
   Mask2.Visible = False
    cmdInicial.Visible = False
    cmdFinal.Visible = False
   lblCONnome.Visible = False
   cboNome.Visible = False
   cboMes.SetFocus
ElseIf cboFiltro.Text = "PERIODO" Then
   lblCONmes.Visible = False
   cboMes.Visible = False
   cboAno.Visible = False
   lblCONint1.Visible = True
   Mask1.Visible = True
   lblCONint2.Visible = True
   Mask2.Visible = True
   Mask1.SetFocus
    cmdInicial.Visible = True
    cmdFinal.Visible = True
   lblCONnome.Visible = False
   cboNome.Visible = False
ElseIf cboFiltro.Text = "CLIENTE" Then
   lblCONmes.Visible = False
   cboMes.Visible = False
   cboAno.Visible = False
   lblCONint1.Visible = False
   Mask1.Visible = False
   lblCONint2.Visible = False
   Mask2.Visible = False
    cmdInicial.Visible = False
    cmdFinal.Visible = False
   lblCONnome.Visible = True
   cboNome.Visible = True
   cboNome.SetFocus
End If
End Sub


Private Sub cboFormaEntrada_GotFocus()
cboFormaEntrada.Clear
cboFormaEntrada.AddItem "DINHEIRO"
cboFormaEntrada.AddItem "CHEQUE"
cboFormaEntrada.AddItem "CARTÃO - DÉBITO"
cboFormaEntrada.AddItem "CARTÃO - CRÉDITO"
cboFormaEntrada.AddItem "DEPOSITO"
cboFormaEntrada.AddItem "TRANSFERÊNCIA"
cboFormaEntrada.AddItem "PIX"
   
If cboFormaEntrada.ListCount <> 0 Then cboFormaEntrada.ListIndex = 0
moCombo.AttachTo cboFormaEntrada
End Sub


Private Sub cboMes_GotFocus()
   Dim vMes As Integer
   
   cboMes.Clear
   
   For vMes = 1 To 12
      cboMes.AddItem StrConv(MonthName(vMes), vbProperCase)
   Next
   
   moCombo.AttachTo cboMes
End Sub

Private Sub cboNome_GotFocus()
'Dim sSQL As String
'Dim r As ADODB.Recordset

cboNome.Clear

sSQL = "SELECT DISTINCT nome, codigo FROM cliente ORDER BY nome;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboNome.AddItem ValidateNull(r("nome"))
   cboNome.ItemData(cboNome.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

moCombo.AttachTo cboNome
End Sub

Private Sub cboNome_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      'cmdExibirConsulta_Click
   End If
   
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboNome_LostFocus()
   On Error GoTo TrataErro
   
   If cboNome.Text = "" Then txtCodClienteCons.Text = "": Exit Sub
   txtCodClienteCons = cboNome.ItemData(cboNome.ListIndex)
   Exit Sub
   
TrataErro:
   If Err.Number = 381 Then cboNome.Text = ""
End Sub

Private Sub cboOrdem_Change()
'cmdExibirConsulta_Click
End Sub

Private Sub cboOrdem_GotFocus()
cboOrdem.Clear
cboOrdem.AddItem "VENC."
cboOrdem.AddItem "CLIENTE"
moCombo.AttachTo cboFiltro
End Sub


Private Sub cboEquipamento_GotFocus()
'Dim sSQL As String
'Dim r As ADODB.Recordset

cboEquipamento.Clear

sSQL = "SELECT DISTINCT descricao, MODELO, FABRICANTE, cod_equip, (quant_estoque - quant_alugada) as varDisponivel FROM aluguel_cadastro_equipamento ORDER BY descricao;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboEquipamento.AddItem ValidateNull(r("descricao")) & " /  " & ValidateNull(r("MODELO")) & " / " & ValidateNull(r("FABRICANTE")) & " / " & r("varDisponivel")
   cboEquipamento.ItemData(cboEquipamento.NewIndex) = r("cod_equip")
   r.MoveNext
Loop
'rTabela("descricao") & " /  " & rTabela("var_tam") & " / " & rTabela("var_fab")
 If r.State <> 0 Then r.Close
 Set r = Nothing

moCombo.AttachTo cboEquipamento
End Sub

Private Sub cboEquipamento_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub SelecionarEquipamento()
On Error GoTo TrataErroSelEquip

If cboEquipamento.Text = "" Then txtCodEquip.Text = "": Exit Sub

txtCodEquip = cboEquipamento.ItemData(cboEquipamento.ListIndex)

txtTipoCobranca.Text = "DIA"
txtQuantAlugada.Text = "1"

sSQL = "SELECT COD_EQUIP, descricao, VALOR_DIA as varValorDiaAluguel FROM  aluguel_cadastro_equipamento WHERE (COD_EQUIP = " & txtCodEquip.Text & ");"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    txtValorAluguel.Text = Format(ValidateNull(r("varValorDiaAluguel")), "##,##0.00")
    frmCodicoes.Enabled = True
End If
If r.State <> 0 Then r.Close
Set r = Nothing

mskHoraInicio.Text = Format(Now, "hh:mm")
mskInicio.Text = Format(Date, "dd/mm/yy")
mskFinal.Text = Format(DateAdd("d", 1, mskInicio.Text), "dd/mm/yy")
mskHoraFinal.Text = mskHoraInicio.Text

CalcularValorTotal
CalcularQuantDias
CalcularDiaria

Exit Sub

TrataErroSelEquip:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cboEquipamento_Click()
SelecionarEquipamento
End Sub

Private Sub cboEquipamento_LostFocus()
SelecionarEquipamento
End Sub

Private Sub cboSituacao_GotFocus()
cboSituacao.Clear
cboSituacao.AddItem "ABERTO"
cboSituacao.AddItem "FECHADO"
   
If cboSituacao.ListCount <> 0 Then cboSituacao.ListIndex = 0
moCombo.AttachTo cboSituacao
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

mskInicio = Format(varData, "dd/mm/yy")   'Exibe a data no campo
mskInicio_LostFocus
End Sub

Private Sub chameleonButton2_Click()
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

mskFinal = Format(varData, "dd/mm/yy")   'Exibe a data no campo
mskFinal_LostFocus
mskFinal.SetFocus
End Sub


Private Sub chameleonButton3_Click()
Dim varData As Variant
Dim fCal As Calendario

varData = Empty                    'Inicializa a variável

Set fCal = New Calendario      'Cria o form de calendário
fCal.Left = (Screen.Width - Me.Width) / 2
fCal.Top = 1000
fCal.Show vbModal

varData = fCal.DateSelected    'Recupera a data selecionada

Unload fCal                           'Fecha o form
Set fCal = Nothing                   'Destrói a variável

If Not IsDate(varData) Then Exit Sub   'Valida a data
If varData = 0 Then Exit Sub

mskDevolver = Format(varData, "dd/mm/yy")   'Exibe a data no campo
mskDevolver_LostFocus
End Sub


Private Sub chameleonButton5_Click()
Dim varData As Variant
Dim fCal As Calendario

varData = Empty                    'Inicializa a variável

Set fCal = New Calendario      'Cria o form de calendário
fCal.Left = (Screen.Width - Me.Width) / 2
fCal.Top = 1000
fCal.Show vbModal

'fCal.Top = (Aluguel_Cadastro.Height - Me.Height) / 2
varData = fCal.DateSelected    'Recupera a data selecionada

Unload fCal                           'Fecha o form
Set fCal = Nothing                   'Destrói a variável

If Not IsDate(varData) Then Exit Sub   'Valida a data
If varData = 0 Then Exit Sub

mskDataProrrogar = Format(varData, "dd/mm/yy")   'Exibe a data no campo
'mskDataProrrogar_LostFocus
End Sub


Private Sub cmdAdiacao_Click()
i = GridProdutos.Row

If CDate(mskDataProrrogar.Text) > CDate(mskDataFinalLocacaoProrro.Text) Then
    MsgBox "A data de retrocesso é maior que a data de devolução", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If CCur(txtEntradaAdiar.Text) > CCur(txtTotalAluguelDescProrro.Text) Then
    MsgBox "O valor da entrada é maior que o valor total do aluguel", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

'saber a quantidades de dias entre duas datas
    If Not IsDate(mskDataInicialLocacaoProrro) Then Exit Sub
    If Not IsDate(mskDataProrrogar) Then Exit Sub
    
    date1 = CDate(mskDataInicialLocacaoProrro.Text)
    date2 = CDate(mskDataProrrogar.Text)
    
    vQuantDias = DateDiff("d", date1, date2)
    vQuantDias = vQuantDias
    
    If mskDataInicialLocacaoProrro.Text = mskDataProrrogar.Text Then
        vQuantDias = 1
    Else
        vQuantDias = vQuantDias
    End If
            
            'calcular o total
            Dim vValorDiariaAdiarT As Currency
            vValorDiariaAdiarT = txtValorAluguellProrro.Text
            
            Dim vValorTotalAdiarT As Currency
            vValorTotalAdiarT = vValorDiariaAdiarT * vQuantDias
            
            Dim vValorDescAdiarT As Double
            vValorDescAdiarT = txtDescProrro.Text
            
            vValorTotalAdiarT = vValorTotalAdiarT - vValorDescAdiarT
                        
            Dim ValorEntradaAdiarT As Currency
            ValorEntradaAdiarT = txtEntradaAdiar.Text
            
            Dim ValorTotalAdiarMenosEntradaT As Currency
            ValorTotalAdiarMenosEntradaT = vValorTotalAdiarT - ValorEntradaAdiarT
            
            If ValorTotalAdiarMenosEntradaT < 0 Then
                ValorTotalAdiarMenosEntradaT = 0
            Else
                ValorTotalAdiarMenosEntradaT = ValorTotalAdiarMenosEntradaT
            End If

If ValorTotalAdiarMenosEntradaT = 0 Then
    MsgBox "O valor do retrocesso é menor que o valor da entrada", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

'mudar o item
    Dim vDescNovo As Currency
    vDescNovo = txtDescProrro.Text
    dbData.Execute "UPDATE Aluguel_Cadastro_Itens SET DATA_FINAL = CONVERT(DATETIME, '" & Format(mskDataProrrogar, ocDATA) & "', 103), QUANT = " & vQuantDias & ", VALOR = (TOTAL_ALUGADA * " & vQuantDias & ") , SUBTOTAL = ((TOTAL_ALUGADA * " & vQuantDias & ") - DESCONTO) - " & vDescNovo & ", VALOR_FINAL = (((TOTAL_ALUGADA * " & vQuantDias & ") - DESCONTO) - Entrada) - " & vDescNovo & ", ADIADO = 1, DESCONTO = (DESCONTO + " & vDescNovo & ") WHERE (item = " & GridProdutos.TextMatrix(i, 1) & ") and (COD_LOCACAO = " & lblCodigo.Caption & ");"

'Apagar as parcelas que nao seja a entrada
    dbData.Execute "DELETE FROM parcelas WHERE (cod_os = " & lblCodigo.Caption & ") and (OS_ITEM = " & GridProdutos.TextMatrix(i, 1) & ") and (status = 0)"

'CRIAR A SEGUNDA PARCELA
'saber quantos dias de adiação entre datas sem contar a entrada
    If Not IsDate(mskDataInicialLocacaoProrro) Then Exit Sub
    Dim vDataAposEntrada As Date
    vDataAposEntrada = Format(DateAdd("d", Val(1), mskDataInicialLocacaoProrro.Text), "dd/mm/yy")
        
    Dim vQuantDiasAdiar As Integer
    Dim vDataFinalAdiacao As Date
    
    If Not IsDate(mskDataProrrogar) Then Exit Sub
    vDataFinalAdiacao = CDate(mskDataProrrogar)
    
    If Not IsDate(mskDataInicialLocacaoProrro) Then Exit Sub
    Dim vDataInicialLocacao As Date
    vDataInicialLocacao = CDate(mskDataInicialLocacaoProrro)
    
    If vDataFinalAdiacao = vDataInicialLocacao Then
        vQuantDiasAdiar = "1"
    ElseIf vDataFinalAdiacao < vDataInicialLocacao Then
        MsgBox "Não é possivel adiar para uma data anterior a locação!", vbInformation, "Aviso do Sistema"
        Exit Sub
    Else
        If vDataAposEntrada = vDataFinalAdiacao Then
            vQuantDiasAdiar = "1"
        Else
            vQuantDiasAdiar = DateDiff("d", vDataAposEntrada, vDataFinalAdiacao)
        End If
    End If

CalcularTotalAdiar

Gerar_ParcelasSegundaAdiacao

LimparObjetos_Prorrogacao
frmProrrogacao.Visible = False

'MostrarParcelas
LimparGrid_Parcelas
PreencherGridProdutos
End Sub

Private Sub cmdAdiar_Click()
frmProrrogacao.Visible = True
frmProrrogacao.Caption = "Antecipação"
frmDevolucao.Visible = False
frmPagamento.Visible = False
cmdAdiacao.Visible = True
cmdProrrogacao.Visible = False
LimparObjetos_Prorrogacao

i = GridProdutos.Row
mskDataInicialLocacaoProrro.Text = GridProdutos.TextMatrix(i, 14)
mskDataFinalLocacaoProrro.Text = GridProdutos.TextMatrix(i, 16)
mskDataInicioProrro.Text = Format(DateAdd("d", Val(0), mskDataFinalLocacaoProrro.Text), "dd/mm/yy")
mskDataProrrogar.Text = Format(Date, "dd/mm/yy")

txtEntradaAdiar.Text = Format(GridProdutos.TextMatrix(i, 10), "##,##0.00")

varCodItem = GridProdutos.TextMatrix(i, 1)
'txtDescProrro.Text = GridProdutos.TextMatrix(i, 8)
'txtValorRealDescProrro.Text = GridProdutos.TextMatrix(i, 10)
'txtValorAluguellProrro.Text = GridProdutos.TextMatrix(i, 5)

'calcular a quantidade de dias
'calculardiasAdiar

'calcular o valor final com descont
calculardiasAdiar
CalcularTotalAdiar
End Sub
Private Sub cmdAdicionarProduto_Click()
'Dim sSQL As String
'Dim r As ADODB.Recordset

Dim varCodEntrada As Long

If txtCodCliente.Text = "" Then
   ShowMsg "Falta escolher o cliente!", vbInformation
   txtCliente.SetFocus
   Exit Sub
End If

If lblCodigo.Caption = "" Or cboEquipamento.Text = "" Or txtValorAluguel.Text = "" Then Exit Sub

'atualizar os equipamentos alugados
sSQL = "SELECT (QUANT_ESTOQUE - QUANT_ALUGADA) AS varDisponivel FROM Aluguel_Cadastro_Equipamento WHERE (COD_EQUIP = " & txtCodEquip.Text & ");"
Set r = dbData.OpenRecordset(sSQL)

Dim varQuantAtual As Integer
Dim varQuantDisponivel As Integer

varQuantAtual = txtQuantAlugada.Text
varQuantDisponivel = r("varDisponivel")

If varQuantAtual > varQuantDisponivel Then
    MsgBox "A quantidade de equipamentos escolhidos é maior que a quantidade disponível!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

varCodItem = 1
   
If txtTipoCobranca.Text = "DIA" Then
    If txtQuant.Text = "" Then MsgBox "Preencha o restante das informações!", vbInformation, "Aviso do Sistema": Exit Sub

    'indice do codigo
    sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultimo_codigo FROM aluguel_cadastro_itens;"
    Set r = dbData.OpenRecordset(sSQL)
    
    If Not r.BOF Then varCodEntrada = r("ultimo_codigo") + 1
    If r.State <> 0 Then r.Close
    Set r = Nothing
    
    'indice do item
    Dim IndiceItem As Integer
    
    sSQL = "SELECT ISNULL(MAX(item), 0) AS ultimo_item FROM aluguel_cadastro_itens WHERE COD_LOCACAO = " & lblCodigo.Caption & ";"
    Set r = dbData.OpenRecordset(sSQL)
    
    If Not r.BOF Then varCodItem = r("ultimo_item") + 1
    If r.State <> 0 Then r.Close
    Set r = Nothing
    
    'inserir
    dbData.Execute "INSERT INTO aluguel_cadastro_itens (Codigo, COD_LOCACAO, Item, TIPO_LOCACAO, COD_EQUIP, VALOR_UND, TOTAL_ALUGADA, QUANT_ALUGADA, DATA_INICIO, HORA_INICIO, DATA_FINAL, HORA_FINAL, QUANT, VALOR_FINAL, DESCONTO, VALOR, ENTRADA, SUBTOTAL) VALUES (" & _
          varCodEntrada & ", " & lblCodigo.Caption & ", " & varCodItem & ", '" & txtTipoCobranca.Text & "', " & txtCodEquip.Text & ", " & Replace(CCur(txtValorAluguel.Text), ",", ".") & ", " & Replace(CCur(txtTotalAluguel.Text), ",", ".") & ", " & txtQuantAlugada.Text & ", CONVERT(DATETIME, '" & Format(mskInicio, ocDATA) & "', 103), '" & mskHoraInicio & "', CONVERT(DATETIME, '" & Format(mskFinal, ocDATA) & "', 103), '" & mskHoraFinal & "', " & txtQuant.Text & ", " & Replace(CCur(txtTotal.Text), ",", ".") & ", " & Replace(CCur(txtDesconto.Text), ",", ".") & ", " & Replace(CCur(txtValor.Text), ",", ".") & ", " & Replace(CCur(txtEntradaReal.Text), ",", ".") & ", " & Replace(CCur(txtSubTotal.Text), ",", ".") & ");"

Else
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultimo_item FROM aluguel_cadastro_itens;"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then varCodItem = r("ultimo_item") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing

   dbData.Execute "INSERT INTO aluguel_cadastro_itens (codigo, CODIGO, cod_produto, preco, quantidade, data, maquina, descricao, tipo_venda) VALUES (" & _
         varCodItem & ", " & lblCodigo.Caption & ", " & txtCodEquip.Text & ", " & Replace(CCur(txtValorAluguel.Text), ",", ".") & ", 1, CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), '" & _
         IIf(StatusBar1.Panels(3).Text = "", "CAIXA01", StatusBar1.Panels(3).Text) & "', '" & cboEquipamento.Text & "', 'BALCAO');"
End If

'colocar a quantidade de itens alugados
dbData.Execute "UPDATE aluguel_cadastro_equipamento SET QUANT_ALUGADA = QUANT_ALUGADA + " & txtQuantAlugada.Text & " WHERE (COD_EQUIP = " & txtCodEquip.Text & ");"

PreencherGridProdutos

If txtEntrada.Text <> "0,00" Then
    Gerar_ParcelasEntrada
    If txtEntrada <> "100,00" Then
        Gerar_ParcelasSegunda
    End If
Else
    Gerar_ParcelasSegunda
End If

LimparObjetosItens
cboEquipamento.SetFocus
End Sub



Private Sub LimparObjetosItens()
txtValorAluguel.Text = ""
cboEquipamento.Text = ""
txtCodEquip.Text = ""
mskInicio.Text = ""
mskInicio.Mask = ""
mskHoraInicio.Mask = ""
mskHoraInicio.Text = ""
mskFinal.Mask = ""
mskFinal.Text = ""
mskHoraFinal.Mask = ""
mskHoraFinal.Text = ""
txtTipoCobranca.Text = ""
txtQuant.Text = ""
txtTotal.Text = ""
txtValor.Text = ""
txtDesconto.Text = ""
txtQuantAlugada.Text = ""
txtTotalAluguel.Text = ""
txtEntrada.Text = ""
txtSubTotal.Text = ""
txtEntradaReal.Text = ""
txtTipoCobranca.Text = "DIA"
End Sub



Private Sub cmdAlterar_Click()
If cboSituacao.Text = "ABERTO" Then
    vSituacao = False
ElseIf cboSituacao.Text = "FECHADO" Then
    vSituacao = True
End If

If lblCodigo.Caption = "" Or txtCliente.Text = "" Or lblSomaReferente.Caption = "" Then Exit Sub

dbData.Execute "UPDATE aluguel_cadastro SET CODIGO = " & lblCodigo.Caption & ", cod_cliente = " & txtCodCliente.Text & ", OBRA = '" & txtDescricaoObra.Text & "', CIDADEOBRA = '" & cboCidadeObra.Text & "', " & _
               "STATUS = '" & vSituacao & "'  " & _
               "WHERE (CODIGO = " & lblCodigo.Caption & ");"

dbData.Execute "UPDATE pedidos SET cod_cliente = " & txtCodCliente.Text & ", " & _
      "data_compra = CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), tipo_desc = 'P', valor_desc = 0, " & _
      "subtotal = " & Replace(CCur(lblSomaReferente.Caption), ",", ".") & ", total = " & Replace(CCur(lblSomaReferente.Caption), ",", ".") & ", " & _
      "tipo_pagamento = 'À Prazo', pagamento = 'AVULSO', " & _
      "status_pedido = 1, " & _
      "maquina = '" & StatusBar1.Panels(2).Text & "', CodCaixa = " & StatusBar1.Panels(4).Text & ", CAIXA = '" & StatusBar1.Panels(3).Text & "', Entrada = 0, COD_FUNCIONARIO = " & txtCodFuncionario.Text & ", TIPO_ACRESCIMO = 'P', VALOR_ACRESCIMO = 0, ValorDescReal = 0, ValorAcrescReal = 0 WHERE (cod_pedido = " & lblCodPedido.Caption & ");"

Limpar_Objetos
LimparGridProdutos
LimparObjetosItens
Form_Load
cmdExibirConsulta_Click
frmDevolucao.Visible = False
frmPagamento.Visible = False
End Sub


Private Sub cmdCadastrarCliente_Click()
Clientes_Cadastro.Show 1
End Sub

Private Sub cmdCancelar_Click()
dbData.Execute "DELETE FROM aluguel_cadastro WHERE (CODIGO = " & lblCodigo.Caption & ");"
dbData.Execute "DELETE FROM pedidos WHERE (COD_PEDIDO = " & lblCodPedido.Caption & ");"
dbData.Execute "DELETE FROM Aluguel_Cadastro_Itens WHERE (COD_LOCACAO = " & lblCodigo.Caption & ");"
dbData.Execute "DELETE FROM parcelas WHERE (COD_PEDIDO = " & lblCodPedido.Caption & ");"

'desfazer as quantidade de produtos alocados
Dim i As Integer
grid_Parcelas.Col = 0

For i = 1 To GridProdutos.rows - 1
    GridProdutos.Row = i
    dbData.Execute "UPDATE aluguel_cadastro_equipamento SET QUANT_ALUGADA =  QUANT_ALUGADA - " & GridProdutos.TextMatrix(i, 4) & " WHERE (COD_EQUIP = " & GridProdutos.TextMatrix(i, 12) & ");"
Next

frmCliente.Enabled = False
frmReferente.Enabled = False
cmdNovo.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdAdicionarProduto.Enabled = False
cmdRemoverProduto.Enabled = False
LimparGridProdutos
LimparGrid_Parcelas
LimparObjetosItens
Limpar_Objetos
Form_Load
End Sub

Public Function Verifica_Dia(DIA, var_Mes)
   Dim diasDoMes As Variant
   
   DIA = Val(DIA)
   diasDoMes = Array(31, 28, 30, 30, 31, 30, 31, 30, 30, 31, 30, 31)
   
   If DIA = 31 Then
      Verifica_Dia = diasDoMes(var_Mes - 1)
   Else
      Verifica_Dia = DIA
   End If
End Function


Private Sub Gerar_ParcelasHora()
Dim lNovoCod As Long
Dim Var_NumParc As Integer

Dim varTipoCartao As String
varTipoCartao = "NULL"

Dim var_PAGAMENTO As String
var_PAGAMENTO = "NULL"

i = GridProdutos.Row

Dim vItem As Integer
vItem = GridProdutos.TextMatrix(i, 1)

ConsultarUltimaParcela

Var_NumParc = UltimoParcela + 1

lNovoCod = Autonumeracao_Parcelas
   '
dbData.Execute "INSERT INTO parcelas (codigo, cod_pedido, COD_OS, OS_ITEM, numero, data, hora, status, valor, VALOR_FINAL, multa, desconto, tipo, caixa, codcaixa) VALUES (" & _
      lNovoCod & ", " & lblCodPedido.Caption & ", " & lblCodigo.Caption & ", " & vItem & ", " & Var_NumParc & ", CONVERT(DATETIME, '" & Format(mskDevolver, ocDATA) & "', 103), '" & Format(Now, ocHRMN) & "', 0, " & _
      Replace(CCur(vValorTotalHoras), ",", ".") & ", " & Replace(CCur(vValorTotalHoras), ",", ".") & ", 0, 0, 'ALUGUEL', '" & StatusBar1.Panels(3).Text & "', " & varCodCaixa & ");"
End Sub

Private Sub Gerar_ParcelasProrrogacao()
Dim lNovoCod As Long
Dim Var_NumParc As Integer

Dim varTipoCartao As String
varTipoCartao = "NULL"

Dim var_PAGAMENTO As String
var_PAGAMENTO = "NULL"

i = GridProdutos.Row

Dim vItem As Integer
vItem = GridProdutos.TextMatrix(i, 1)

ConsultarUltimaParcela

Var_NumParc = UltimoParcela + 1

lNovoCod = Autonumeracao_Parcelas
   '
dbData.Execute "INSERT INTO parcelas (codigo, cod_pedido, COD_OS, OS_ITEM, numero, data, hora, status, valor, VALOR_FINAL, multa, desconto, tipo, caixa, codcaixa) VALUES (" & _
      lNovoCod & ", " & lblCodPedido.Caption & ", " & lblCodigo.Caption & ", " & vItem & ", " & Var_NumParc & ", CONVERT(DATETIME, '" & Format(mskDataProrrogar, ocDATA) & "', 103), '" & Format(Now, ocHRMN) & "', 0, " & _
      Replace(CCur(txtTotalProrro.Text), ",", ".") & ", " & Replace(CCur(txtTotalProrro.Text), ",", ".") & ", 0, " & Replace(CCur(txtValorRealDescProrro.Text), ",", ".") & ", 'ALUGUEL', '" & StatusBar1.Panels(3).Text & "', " & varCodCaixa & ");"
End Sub


Private Sub Gerar_ParcelasSegundaAdiacao()
Dim lNovoCod As Long
Dim Var_NumParc As Integer

Dim vValorRestante As Currency
Dim vDescAtual As Currency
Dim vTotalAtual As Currency

vDescAtual = txtDescProrro.Text
vTotalAtual = txtTotalProrro.Text
vValorRestante = vDescAtual + vTotalAtual

Dim varTipoCartao As String
varTipoCartao = "NULL"

Dim var_PAGAMENTO As String
var_PAGAMENTO = "NULL"

ConsultarUltimaParcela

Var_NumParc = UltimoParcela + 1

lNovoCod = Autonumeracao_Parcelas

dbData.Execute "INSERT INTO parcelas (codigo, cod_pedido, COD_OS, OS_ITEM, numero, data, hora, status, valor, VALOR_FINAL, multa, desconto, tipo, caixa, codcaixa) VALUES (" & _
      lNovoCod & ", " & lblCodPedido.Caption & ", " & lblCodigo.Caption & ", " & varCodItem & ", " & Var_NumParc & ", CONVERT(DATETIME, '" & Format(mskDataProrrogar, ocDATA) & "', 103), '" & Format(Now, ocHRMN) & "', 0, " & _
      Replace(CCur(vValorRestante), ",", ".") & ", " & Replace(CCur(txtTotalProrro.Text), ",", ".") & ", 0, " & Replace(CCur(txtDescProrro.Text), ",", ".") & ", 'ALUGUEL', '" & StatusBar1.Panels(3).Text & "', " & varCodCaixa & ");"
End Sub

Private Sub Gerar_ParcelasSegunda()
Dim lNovoCod As Long
Dim Var_NumParc As Integer

Dim varTipoCartao As String
varTipoCartao = "NULL"

If cboFormaEntrada.Text = "CARTÃO - DÉBITO" Then
   varTipoCartao = "'D'"
ElseIf cboFormaEntrada.Text = "CARTÃO - CRÉDITO" Then
   varTipoCartao = "'C'"
Else
    varTipoCartao = "NULL"
End If

Dim var_PAGAMENTO As String
If cboFormaEntrada.Text = "DINHEIRO" Then
   var_PAGAMENTO = "DINHEIRO"
ElseIf cboFormaEntrada.Text = "CARTÃO - DÉBITO" Then
   var_PAGAMENTO = "CARTAO"
ElseIf cboFormaEntrada.Text = "CARTÃO - CRÉDITO" Then
   var_PAGAMENTO = "CARTAO"
ElseIf cboFormaEntrada.Text = "CHEQUE" Then
   var_PAGAMENTO = "CHEQUE"
ElseIf cboFormaEntrada.Text = "TRANSFERÊNCIA" Then
   var_PAGAMENTO = "TRANSFERENCIA"
ElseIf cboFormaEntrada.Text = "DEPOSITO" Then
   var_PAGAMENTO = "DEPOSITO"
ElseIf cboFormaEntrada.Text = "PIX" Then
   var_PAGAMENTO = "PIX"
End If

ConsultarUltimaParcela

Dim vDescVazio As Currency
vDescVazio = 0

Var_NumParc = UltimoParcela + 1

lNovoCod = Autonumeracao_Parcelas
   '
dbData.Execute "INSERT INTO parcelas (codigo, cod_pedido, COD_OS, OS_ITEM, numero, data, hora, status, valor, VALOR_FINAL, multa, desconto, tipo, caixa, codcaixa, DIAS_ATRAZO, JUROS, HAVER) VALUES (" & _
      lNovoCod & ", " & lblCodPedido.Caption & ", " & lblCodigo.Caption & ", " & varCodItem & ", " & Var_NumParc & ", CONVERT(DATETIME, '" & Format(mskFinal, ocDATA) & "', 103), '" & Format(Now, ocHRMN) & "', 0, " & _
      Replace(CCur(txtTotal.Text), ",", ".") & ", " & Replace(CCur(txtTotal.Text), ",", ".") & ", 0, " & Replace(CCur(vDescVazio), ",", ".") & ", 'ALUGUEL', '" & StatusBar1.Panels(3).Text & "', " & varCodCaixa & ", 0, 0, 0);"
End Sub
Private Sub Gerar_ParcelasEntrada()
Dim lNovoCod As Long
Dim Var_NumParc As Integer

Dim varTipoCartao As String
varTipoCartao = "NULL"

If cboFormaEntrada.Text = "CARTÃO - DÉBITO" Then
   varTipoCartao = "'D'"
ElseIf cboFormaEntrada.Text = "CARTÃO - CRÉDITO" Then
   varTipoCartao = "'C'"
Else
    varTipoCartao = "NULL"
End If

Dim var_PAGAMENTO As String
If cboFormaEntrada.Text = "DINHEIRO" Then
   var_PAGAMENTO = "DINHEIRO"
ElseIf cboFormaEntrada.Text = "CARTÃO - DÉBITO" Then
   var_PAGAMENTO = "CARTAO"
ElseIf cboFormaEntrada.Text = "CARTÃO - CRÉDITO" Then
   var_PAGAMENTO = "CARTAO"
ElseIf cboFormaEntrada.Text = "CHEQUE" Then
   var_PAGAMENTO = "CHEQUE"
ElseIf cboFormaEntrada.Text = "TRANSFERÊNCIA" Then
   var_PAGAMENTO = "TRANSFERENCIA"
ElseIf cboFormaEntrada.Text = "DEPOSITO" Then
   var_PAGAMENTO = "DEPOSITO"
ElseIf cboFormaEntrada.Text = "PIX" Then
   var_PAGAMENTO = "PIX"
End If

ConsultarUltimaParcela

Dim vDescVazio As Currency
vDescVazio = 0

Var_NumParc = UltimoParcela + 1

lNovoCod = Autonumeracao_Parcelas
   '
dbData.Execute "INSERT INTO parcelas (codigo, cod_pedido, COD_OS, OS_ITEM, numero, data, hora, status, valor, VALOR_FINAL, multa, desconto, tipo, pagamento, FORMA_PGTO, TIPO_CARTAO, caixa, codcaixa, COD_FUNCIONARIO, DIAS_ATRAZO, JUROS, HAVER) VALUES (" & _
      lNovoCod & ", " & lblCodPedido.Caption & ", " & lblCodigo.Caption & ", " & varCodItem & ", " & Var_NumParc & ", CONVERT(DATETIME, '" & Format(mskInicio, ocDATA) & "', 103), '" & Format(Now, ocHRMN) & "', 1, " & _
      Replace(CCur(txtEntradaReal.Text), ",", ".") & ", " & Replace(CCur(txtEntradaReal.Text), ",", ".") & ", 0, " & Replace(CCur(vDescVazio), ",", ".") & ", 'ALUGUEL', CONVERT(DATETIME, '" & Format(mskInicio, ocDATA) & "', 103), '" & var_PAGAMENTO & "', " & varTipoCartao & ", '" & StatusBar1.Panels(3).Text & "', " & varCodCaixa & ", " & txtCodFuncionario.Text & ", 0, 0, 0);"
End Sub
Private Function Autonumeracao_Parcelas() As Long
Dim lRet As Long

lRet = 0
sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultima_parcela FROM parcelas;"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then lRet = r("ultima_parcela") + 1
If r.State <> 0 Then r.Close
Set r = Nothing

Autonumeracao_Parcelas = lRet
End Function

Private Sub cmdDevolver_Click()
frmProrrogacao.Visible = False
frmDevolucao.Visible = True
LimparObjetosDevolucao

'Dim i As Integer

i = GridProdutos.Row

'If GridProdutos.TextMatrix(i, 2) = "" Then Exit Sub

If GridProdutos.TextMatrix(i, 13) = "SIM" Then
    MsgBox "Esse equipamento já foi devolvido!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If


If GridProdutos.TextMatrix(i, 18) = "SIM" Then
    MsgBox "Esse equipamento já foi removido do contrato!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

vTipoDevolver = 1
frmDevolucao.Visible = True
frmPagamento.Visible = True
mskDevolver = Format(Date, "dd/mm/yy")   'Exibe a data no campo
mskDevolverHora = Format(Now, "hh:mm")   'Exibe a data no campo
txtQuantDev.Text = GridProdutos.TextMatrix(i, 4)
mskDataFinalLocacao.Text = GridProdutos.TextMatrix(i, 16)
mskDataFinalLocacaoHora.Text = GridProdutos.TextMatrix(i, 17)
txtQuantDev.Locked = True
AtualizarPagamentoDevolucao
txtQuantDev.SetFocus
End Sub



Private Function ExtrairFormaPagamento(ByVal vTextoCombo As String) As String
If vTextoCombo = "1 - DINHEIRO" Then
   ExtrairFormaPagamento = "DINHEIRO"
ElseIf vTextoCombo = "3 - CARTÃO - DÉBITO" Then
   ExtrairFormaPagamento = "CARTAO"
ElseIf vTextoCombo = "4 - CARTÃO - CRÉDITO" Then
   ExtrairFormaPagamento = "CARTAO"
ElseIf vTextoCombo = "5 - CHEQUE" Then
   ExtrairFormaPagamento = "CHEQUE"
ElseIf vTextoCombo = "7 - TRANSFERÊNCIA" Then
   ExtrairFormaPagamento = "TRANSFERENCIA"
ElseIf vTextoCombo = "8 - DEPOSITO" Then
   ExtrairFormaPagamento = "DEPOSITO"
ElseIf vTextoCombo = "9 - FINANCEIRA" Then
   ExtrairFormaPagamento = "FINANCEIRA"
ElseIf vTextoCombo = "10 - PIX" Then
   ExtrairFormaPagamento = "PIX"
Else
   ExtrairFormaPagamento = "DINHEIRO"
End If
End Function

Private Function ExtrairTipoCartao(ByVal vTextoCombo As String) As String
If vTextoCombo = "3 - CARTÃO - DÉBITO" Then
   ExtrairTipoCartao = "'D'"
ElseIf vTextoCombo = "4 - CARTÃO - CRÉDITO" Then
   ExtrairTipoCartao = "'C'"
Else
   ExtrairTipoCartao = "NULL"
End If
End Function

Private Sub RegistrarPagamentoParcela(ByVal vCodParcRP As Long, ByVal vValorPagoRP As Currency, ByVal vFormaRP As String, ByVal vTipoCartaoRP As String)
dbData.Execute "UPDATE parcelas SET status = 1, valor_final = " & Replace(CStr(vValorPagoRP), ",", ".") & ", desconto = 0, tipo = 'PARCELA', tipo_cartao = " & vTipoCartaoRP & ", forma_pgto = '" & vFormaRP & "', COD_FUNCIONARIO = " & txtCodFuncionario.Text & ", pagamento = CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), hora = '" & Format(Now, ocHRMN) & "', caixa = '" & StatusBar1.Panels(3).Text & "', CODCAIXA = " & varCodCaixa & " WHERE (codigo = " & vCodParcRP & ");"
End Sub

Private Sub RegistrarHaverParcela(ByVal vCodParcRH As Long, ByVal vNumeroParcRH As Integer, ByVal vVencimentoParcRH As String, ByVal vValorParcelaRH As Currency, ByVal vValorHaverRH As Currency, ByVal vFormaRH As String, ByVal vTipoCartaoRH As String)
Dim rHaver As ADODB.Recordset
Dim lNovoCodHaver As Long

sSQL = "SELECT ISNULL(MAX(codigo), 0) AS codigo_haver FROM parcelas_haver WITH (UPDLOCK, HOLDLOCK);"
Set rHaver = dbData.OpenRecordset(sSQL)
lNovoCodHaver = rHaver("codigo_haver") + 1
If rHaver.State <> 0 Then rHaver.Close
Set rHaver = Nothing

dbData.Execute "INSERT INTO parcelas_haver (codigo, cod_parcela, numero, vencimento, haver, valor_parcela, valor_haver, hora, forma_pgto, caixa, CODCAIXA, tipo, tipo_cartao, COD_FUNCIONARIO) VALUES (" & _
   lNovoCodHaver & ", " & vCodParcRH & ", " & vNumeroParcRH & ", CONVERT(DATETIME, '" & vVencimentoParcRH & "', 103), CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), " & _
   Replace(CStr(vValorParcelaRH), ",", ".") & ", " & Replace(CStr(vValorHaverRH), ",", ".") & ", '" & Format(Now, ocHRMN) & "', '" & vFormaRH & "', '" & StatusBar1.Panels(3).Text & "', " & varCodCaixa & ", 'PARCELA', " & vTipoCartaoRH & ", " & txtCodFuncionario.Text & ");"

dbData.Execute "UPDATE parcelas SET haver = 1 WHERE (codigo = " & vCodParcRH & ");"
End Sub

Private Sub BaixarParcelasDevolucao()
If cboformaPgto1.Text = "" Then
    MsgBox "Escolha a forma de pagamento 1!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If
If opt2forma.Value = True And cboformaPgto2.Text = "" Then
    MsgBox "Escolha a forma de pagamento 2!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

'le todas as parcelas em aberto ANTES de abrir a transacao
Dim rParc As ADODB.Recordset
Dim colCodigo As New Collection
Dim colNumero As New Collection
Dim colVencimento As New Collection
Dim colValor As New Collection

sSQL = "SELECT codigo, numero, DATA, VALOR_FINAL FROM parcelas WHERE (cod_os = " & lblCodigo.Caption & ") and OS_ITEM = " & GridProdutos.TextMatrix(i, 1) & " and STATUS = 0 ORDER BY numero;"
Set rParc = dbData.OpenRecordset(sSQL)

Do While Not rParc.EOF
    colCodigo.Add CLng(rParc("codigo"))
    colNumero.Add CInt(rParc("numero"))
    colVencimento.Add Format(rParc("DATA"), ocDATA)
    colValor.Add CCur(rParc("VALOR_FINAL"))
    rParc.MoveNext
Loop
If rParc.State <> 0 Then rParc.Close
Set rParc = Nothing

Dim vSaldoForma1 As Currency
Dim vSaldoForma2 As Currency
Dim vNomeForma1 As String
Dim vNomeForma2 As String
Dim vTipoCartao1 As String
Dim vTipoCartao2 As String

vSaldoForma1 = Val(Replace(Replace(txtValor1.Text, ".", ""), ",", "."))
vNomeForma1 = ExtrairFormaPagamento(cboformaPgto1.Text)
vTipoCartao1 = ExtrairTipoCartao(cboformaPgto1.Text)

If opt2forma.Value = True Then
    vSaldoForma2 = Val(Replace(Replace(txtValor2.Text, ".", ""), ",", "."))
    vNomeForma2 = ExtrairFormaPagamento(cboformaPgto2.Text)
    vTipoCartao2 = ExtrairTipoCartao(cboformaPgto2.Text)
Else
    vSaldoForma2 = 0
End If

Dim bTrans As Boolean
On Error GoTo ErrHandlerBaixaParc

dbData.Execute "BEGIN TRANSACTION"
bTrans = True

Dim vIdxParc As Integer
Dim vFaltaParc As Currency

For vIdxParc = 1 To colCodigo.Count
    vFaltaParc = colValor(vIdxParc)

    Do While vFaltaParc > 0 And (vSaldoForma1 > 0 Or vSaldoForma2 > 0)
        If vSaldoForma1 > 0 Then
            If vSaldoForma1 >= vFaltaParc Then
                RegistrarPagamentoParcela colCodigo(vIdxParc), vFaltaParc, vNomeForma1, vTipoCartao1
                vSaldoForma1 = vSaldoForma1 - vFaltaParc
                vFaltaParc = 0
            Else
                RegistrarHaverParcela colCodigo(vIdxParc), colNumero(vIdxParc), colVencimento(vIdxParc), colValor(vIdxParc), vSaldoForma1, vNomeForma1, vTipoCartao1
                vFaltaParc = vFaltaParc - vSaldoForma1
                vSaldoForma1 = 0
            End If
        ElseIf vSaldoForma2 > 0 Then
            If vSaldoForma2 >= vFaltaParc Then
                RegistrarPagamentoParcela colCodigo(vIdxParc), vFaltaParc, vNomeForma2, vTipoCartao2
                vSaldoForma2 = vSaldoForma2 - vFaltaParc
                vFaltaParc = 0
            Else
                RegistrarHaverParcela colCodigo(vIdxParc), colNumero(vIdxParc), colVencimento(vIdxParc), colValor(vIdxParc), vSaldoForma2, vNomeForma2, vTipoCartao2
                vFaltaParc = vFaltaParc - vSaldoForma2
                vSaldoForma2 = 0
            End If
        End If
    Loop
Next vIdxParc

dbData.Execute "COMMIT TRANSACTION"
bTrans = False
Exit Sub

ErrHandlerBaixaParc:
Dim vMsgErroBaixa As String
vMsgErroBaixa = Err.Description
If bTrans Then dbData.Execute "ROLLBACK TRANSACTION"
MsgBox "Erro ao dar baixa nas parcelas: " & vMsgErroBaixa, vbCritical, "Aviso do Sistema"
End Sub

Private Sub cmdDevolverItem_Click()
If GridProdutos.TextMatrix(i, 13) = "SIM" Then
    MsgBox "Esse equipamento já foi devolvido!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If GridProdutos.TextMatrix(i, 18) = "SIM" Then
    MsgBox "Esse equipamento já foi removido do contrato!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If mskDevolver.Text < mskDataFinalLocacao.Text Then
    MsgBox "Cliente devolveu antes do prazo, favor corrigir a prorrogação!", vbInformation, "Aviso do Sistema"
    Exit Sub
ElseIf mskDevolver.Text > mskDataFinalLocacao.Text Then
    MsgBox "Cliente devolveu após o prazo, favor corrigir a prorrogação!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

i = GridProdutos.Row

If ShowMsg("Deseja devolver o equipamento: " & GridProdutos.TextMatrix(i, 2) & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub

'calcular as horas adicionais
Dim vHoraInicial As Date
Dim vHoraFinal As Date
Dim vQuantHoras As Integer

vHoraInicial = TimeValue(mskDataFinalLocacaoHora.Text)
vHoraFinal = TimeValue(mskDevolverHora.Text)

Dim minutos As Long
minutos = DateDiff("n", vHoraInicial, vHoraFinal)

'chamar a função
If vHoraFinal < vHoraInicial Then
    vQuantHoras = 0
Else
    vQuantHoras = GetHora(minutos)
End If

If vQuantHoras >= 1 Then
    If ShowMsg("Você passou seu horário de entrega em: " & vQuantHoras & " hora(s)." & vbCrLf & "Isso levará a um custo adicional." & vbCrLf & "Deseja Continuar ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
        
        'saber o valor da hora
        If GridProdutos.TextMatrix(i, 12) = "" Then Exit Sub

        sSQL = "SELECT VALOR_HORA FROM aluguel_cadastro_equipamento WHERE (COD_EQUIP = " & GridProdutos.TextMatrix(i, 12) & ");"
        Set r = dbData.OpenRecordset(sSQL)
        
        Dim vValorHora As Currency
        
        If Not r.BOF Then
            vValorHora = r("VALOR_HORA")
        End If
        
    'saber o montante de horas
    vValorTotalHoras = vValorHora * vQuantHoras
    
    'adicionar as parcelas
    Gerar_ParcelasHora
    
    'exibir no grid
    MostrarParcelas
End If


'========= deevolver

If vTipoDevolver = 1 Then               'devolver tudo
    'marcar o equipamento como devolvido
    dbData.Execute "UPDATE Aluguel_Cadastro_Itens SET devolvido =  1, DEVOLUCAO = CONVERT(DATETIME, '" & Format(mskDevolver, ocDATA) & "', 103) WHERE (item = " & GridProdutos.TextMatrix(i, 1) & ") and (COD_LOCACAO = " & lblCodigo.Caption & ");"
    
    'marcar o equipamento como livre
    'dbData.Execute "UPDATE aluguel_cadastro_equipamento SET contrato =  0, alugado = 0 WHERE (COD_EQUIP = " & GridProdutos.TextMatrix(i, 13) & ");"
    
    'Remover a parcela desse equipamento
    'dbData.Execute "DELETE FROM parcelas WHERE (cod_os = " & lblCodigo.Caption & ") and OS_ITEM = " & GridProdutos.TextMatrix(i, 1) & " ;"
    
    'atualizar os equipamentos alugados
    sSQL = "SELECT QUANT_ESTOQUE, QUANT_ALUGADA FROM Aluguel_Cadastro_Equipamento WHERE (COD_EQUIP = " & GridProdutos.TextMatrix(i, 12) & ");"
    Set r = dbData.OpenRecordset(sSQL)
    
    Dim varQuantAtual As Integer
    Dim varQuantAlugada As Integer
    Dim varNovaQuant As Integer
    
    varQuantAtual = r("QUANT_ALUGADA")
    varQuantAlugada = GridProdutos.TextMatrix(i, 4)
    varNovaQuant = varQuantAtual - varQuantAlugada
    
    'colocar a quantidade de itens alugados
    dbData.Execute "UPDATE aluguel_cadastro_equipamento SET QUANT_ALUGADA =  " & varNovaQuant & " WHERE (COD_EQUIP = " & GridProdutos.TextMatrix(i, 12) & ");"
    
    'Remover o equipamento
    'dbData.Execute "DELETE FROM Aluguel_Cadastro_Itens WHERE (item = " & GridProdutos.TextMatrix(i, 1) & ") and (COD_LOCACAO = " & lblCodigo.Caption & ");"

ElseIf vTipoDevolver = 2 Then           'devolver parcial
        '====MUDAR O ITEM EXISTE ====================================
        'saber a quantidades de dias entre duas datas
        If Not IsDate(mskDataInicialLocacaoProrro) Then Exit Sub
        If Not IsDate(mskDevolver) Then Exit Sub

        date1 = CDate(mskDataInicialLocacaoProrro.Text)
        date2 = CDate(mskDevolver.Text)

        vQuantDias = DateDiff("d", date1, date2)

        If mskDataInicialLocacaoProrro.Text = mskDataProrrogar.Text Then
            vQuantDias = 1
        End If

        'saber a quantidade devolvida e a quantida restante
        Dim vQuantAlugada As Integer
        Dim vQuantDevolvida As Integer
        Dim vQuantRestante As Integer
        vQuantAlugada = GridProdutos.TextMatrix(i, 4)
        vQuantDevolvida = txtQuantDev
        vQuantRestante = vQuantAlugada - vQuantDevolvida

        If vQuantRestante <= 0 Then
            'devolveu tudo que restava - marca o item como totalmente devolvido
            dbData.Execute "UPDATE Aluguel_Cadastro_Itens SET QUANT_ALUGADA = 0, devolvido = 1, DEVOLUCAO = CONVERT(DATETIME, '" & Format(mskDevolver, ocDATA) & "', 103) WHERE (item = " & GridProdutos.TextMatrix(i, 1) & ") and (COD_LOCACAO = " & lblCodigo.Caption & ");"
        Else
            'atualiza o item com a quantidade restante e os novos totais
            dbData.Execute "UPDATE Aluguel_Cadastro_Itens SET QUANT_ALUGADA = " & vQuantRestante & ", TOTAL_ALUGADA = (VALOR_UND * " & vQuantRestante & "), QUANT = " & vQuantDias & ", Valor = ((VALOR_UND * " & vQuantRestante & ") * " & vQuantDias & "), SUBTOTAL = ((VALOR_UND * " & vQuantRestante & ") * " & vQuantDias & ") - DESCONTO, VALOR_FINAL = (((VALOR_UND * " & vQuantRestante & ") * " & vQuantDias & ") - DESCONTO) - Entrada, DATA_FINAL = CONVERT(DATETIME, '" & Format(mskDevolver, ocDATA) & "', 103) WHERE (item = " & GridProdutos.TextMatrix(i, 1) & ") and (COD_LOCACAO = " & lblCodigo.Caption & ");"
        End If

        'da baixa na quantidade alugada do equipamento (so a parte devolvida agora)
        sSQL = "SELECT QUANT_ALUGADA FROM Aluguel_Cadastro_Equipamento WHERE (COD_EQUIP = " & GridProdutos.TextMatrix(i, 12) & ");"
        Set r = dbData.OpenRecordset(sSQL)

        Dim varQuantAtualEquipParcial As Integer
        Dim varNovaQuantEquipParcial As Integer

        varQuantAtualEquipParcial = r("QUANT_ALUGADA")
        varNovaQuantEquipParcial = varQuantAtualEquipParcial - vQuantDevolvida

        dbData.Execute "UPDATE aluguel_cadastro_equipamento SET QUANT_ALUGADA = " & varNovaQuantEquipParcial & " WHERE (COD_EQUIP = " & GridProdutos.TextMatrix(i, 12) & ");"

End If

'consultar parcelas em aberto do equipamento
sSQL = "SELECT COUNT(codigo) AS vQuantParc, ISNULL(SUM(VALOR_FINAL), 0) AS vSomaParc " & _
       "FROM parcelas WHERE (cod_os = " & lblCodigo.Caption & ") and OS_ITEM = " & GridProdutos.TextMatrix(i, 1) & " and STATUS = 0;"
Set r = dbData.OpenRecordset(sSQL)

Dim vQuantParcAberto As Long
Dim vSomaParcAberto As Currency
vQuantParcAberto = r("vQuantParc")
vSomaParcAberto = r("vSomaParc")
If r.State <> 0 Then r.Close
Set r = Nothing

If vQuantParcAberto > 0 Then
    If ShowMsg("Esse equipamento possui: " & vQuantParcAberto & " parcela(s) em aberto, totalizando " & Format(vSomaParcAberto, ocMONEY) & vbCrLf & "Você poderá devolver o equipamento sem dar baixa nas parcelas." & vbCrLf & "Deseja dar baixa nas parcelas agora ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
        If chkPgtoAuto.Value = vbChecked Then
            BaixarParcelasDevolucao
        Else
            Load Parcelas
            Parcelas.txtCodCliente.Text = txtCodCliente.Text
            Parcelas.cboCliente.Text = txtCliente.Text
            vClienteEncontrado = True
            Parcelas.Show 1
        End If
    End If
End If

vTipoDevolver = 0
PreencherGridProdutos
LimparObjetosDevolucao
frmDevolucao.Visible = False
frmPagamento.Visible = False

'fechar o contrato após devolver todos
'Dim i As Integer
Dim vQuantAtivo As Integer
vQuantAtivo = 0

For i = 1 To GridProdutos.rows - 1
    If GridProdutos.TextMatrix(i, 13) = "NÃO" Then
        If GridProdutos.TextMatrix(i, 18) = "NÃO" Then
            vQuantAtivo = vQuantAtivo + 1
        End If
    End If
Next

If vQuantAtivo = 0 Then
    dbData.Execute "UPDATE aluguel_cadastro SET STATUS = 1 WHERE (CODIGO = " & lblCodigo.Caption & ");"
End If

LimparGrid_Parcelas
End Sub
Private Sub cmdDevolverParcial_Click()
txtQuantDev.Locked = False
frmProrrogacao.Visible = False
frmDevolucao.Visible = True
LimparObjetosDevolucao

'Dim i As Integer

i = GridProdutos.Row

mskDataInicialLocacaoProrro.Text = GridProdutos.TextMatrix(i, 14)

If GridProdutos.TextMatrix(i, 13) = "SIM" Then
    MsgBox "Esse equipamento já foi devolvido!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If


If GridProdutos.TextMatrix(i, 18) = "SIM" Then
    MsgBox "Esse equipamento já foi removido do contrato!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

vTipoDevolver = 2
frmDevolucao.Visible = True
frmPagamento.Visible = True
mskDevolver = Format(Date, "dd/mm/yy")   'Exibe a data no campo
mskDevolverHora = Format(Now, "hh:mm")   'Exibe a hora no campo
txtQuantDev.Text = GridProdutos.TextMatrix(i, 4)
mskDataFinalLocacao.Text = GridProdutos.TextMatrix(i, 16)
mskDataFinalLocacaoHora.Text = GridProdutos.TextMatrix(i, 17)
AtualizarPagamentoDevolucao
txtQuantDev.SetFocus
End Sub

Private Sub cmdExcluir_Click()
'If Tela_Principal.txtNivel.Text <> "1" Then MsgBox "Seu nível de acesso não permite essa operação!", vbInformation, "Aviso do Sistema": Exit Sub
'Dim i As Integer

If lblCodigo.Caption = "" Then Exit Sub

If ShowMsg("Excluir esse contrato?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub

dbData.Execute "UPDATE aluguel_cadastro SET excluido =  1 WHERE (CODIGO = " & lblCodigo.Caption & ");"
dbData.Execute "UPDATE Aluguel_Cadastro_Itens SET excluido =  1, devolvido = 0 WHERE (COD_LOCACAO = " & lblCodigo.Caption & ");"
dbData.Execute "DELETE FROM parcelas WHERE (COD_PEDIDO = " & lblCodPedido.Caption & ");"

'desfazer as quantidade de produtos alocados
Dim i As Integer
grid_Parcelas.Col = 0

For i = 1 To GridProdutos.rows - 1
    GridProdutos.Row = i
    dbData.Execute "UPDATE aluguel_cadastro_equipamento SET QUANT_ALUGADA =  QUANT_ALUGADA - " & GridProdutos.TextMatrix(i, 4) & " WHERE (COD_EQUIP = " & GridProdutos.TextMatrix(i, 12) & ");"
Next

frmCliente.Enabled = False
frmReferente.Enabled = False
cmdNovo.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdImprimirContato.Enabled = False
cmdAdicionarProduto.Enabled = False
cmdRemoverProduto.Enabled = False
'cmdDevolver.Enabled = False

Limpar_Objetos
LimparGridProdutos
LimparObjetosItens
LimparGrid_Parcelas

Form_Load
cmdExibirConsulta_Click
frmDevolucao.Visible = False
frmPagamento.Visible = False
End Sub

Private Sub cmdExibirConsulta_Click()
Dim totalRegistros As Long

Dim vStatus As String
If cboCONStatus.Text = "ABERTO" Then
    vStatus = "and Aluguel_Cadastro_1.STATUS = 0"
ElseIf cboCONStatus.Text = "FECHADO" Then
    vStatus = "and Aluguel_Cadastro_1.STATUS = 1"
Else
    vStatus = " "
End If

sSQL = "SELECT DISTINCT Aluguel_Cadastro_1.CODIGO AS varCodAluguel, Aluguel_Cadastro_1.COD_CLIENTE, cliente.Nome, Aluguel_Cadastro_1.OBRA, Aluguel_Cadastro_1.CIDADEOBRA, " & _
"(CASE Aluguel_Cadastro_1.STATUS WHEN 0 THEN 'ABERTO' ELSE 'FECHADO' END) AS var_STATUS, " & _
"(CASE WHEN Aluguel_Cadastro_1.excluido = 1 THEN 'SIM' ELSE 'NÃO' END) AS varExcluido, " & _
"VALORINICIAL, " & _
"(SELECT ISNULL(SUM(VALOR_FINAL), 0) FROM parcelas WHERE (COD_OS = Aluguel_Cadastro_1.CODIGO)) AS varSomaParcelasTodas, " & _
"(SELECT ISNULL(SUM(VALOR_FINAL), 0) FROM parcelas AS parcelas_1 WHERE (COD_OS = Aluguel_Cadastro_1.CODIGO) AND (STATUS = 0)) AS varSomaParcelasAbertas, " & _
"(SELECT COUNT(Aluguel_Cadastro_Itens_1.CODIGO) FROM Aluguel_Cadastro INNER JOIN Aluguel_Cadastro_Itens AS Aluguel_Cadastro_Itens_1 ON Aluguel_Cadastro.CODIGO = Aluguel_Cadastro_Itens_1.COD_LOCACAO WHERE (Aluguel_Cadastro_Itens_1.COD_LOCACAO = Aluguel_Cadastro_1.CODIGO) AND (Aluguel_Cadastro_Itens_1.DEVOLVIDO = 0) AND (Aluguel_Cadastro_Itens_1.EXCLUIDO = 0)) AS varQuantItens " & _
"FROM Aluguel_Cadastro AS Aluguel_Cadastro_1 INNER JOIN cliente ON Aluguel_Cadastro_1.COD_CLIENTE = cliente.CODIGO "

'"(SELECT SUM(VALOR_FINAL) FROM Aluguel_Cadastro_Itens WHERE (COD_LOCACAO = Aluguel_Cadastro_1.CODIGO) AND (EXCLUIDO = 0)) AS varSomaParcelas, " & _

If cboFiltro.Text = "TODOS" Then
    sSQL = sSQL & "WHERE Aluguel_Cadastro_1.CODIGO > 0 " & vStatus & " ORDER BY varCodAluguel"
ElseIf cboFiltro.Text = "MÊS" Then
    sSQL = sSQL & "WHERE (MONTH(DATACADASTRO) = " & cboMes.ListIndex + 1 & ") AND (YEAR(DATACADASTRO) = " & cboAno & ") " & vStatus & " ORDER BY varCodAluguel"
ElseIf cboFiltro.Text = "PERIODO" Then
    sSQL = sSQL & "WHERE (DATACADASTRO >= CONVERT(DATETIME, '" & Format(Mask1.Text, ocDATA) & "', 103)) AND (DATACADASTRO <= CONVERT(DATETIME, '" & Format(Mask2.Text, ocDATA) & "', 103)) " & vStatus & " ORDER BY varCodAluguel"
ElseIf cboFiltro.Text = "CLIENTE" Then
    sSQL = sSQL & "WHERE (cod_cliente = " & txtCodClienteCons.Text & ") " & vStatus & " ORDER BY varCodAluguel"
End If

Set r = dbData.OpenRecordset(sSQL, totalRegistros)
'Debug.Print sSQL
printSQL = sSQL

FormatarGrid r

'If r.State <> 0 Then r.Close
'Set r = Nothing

'MOSTRAR A QUANTIDADE REGISTROS
txtCONquant.Text = Format(totalRegistros, "00")
End Sub

Private Sub cmdFecharAluguel_Click()
i = GridConsulta.Row
If GridConsulta.TextMatrix(i, 7) <> "0,00" Or GridConsulta.TextMatrix(i, 10) <> "0" Then Exit Sub

dbData.Execute "UPDATE aluguel_cadastro SET STATUS = 1  " & _
               "WHERE (CODIGO = " & GridConsulta.TextMatrix(i, 1) & ");"

cmdExibirConsulta_Click
End Sub

Private Sub cmdFinal_Click()
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

Mask2 = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdImprimirConsulta_Click()
'colocar o nome da maquina na barra de status
Dim var_Impressora As String
Dim oIni As Ini
Dim r As ADODB.Recordset

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
Set oIni = Nothing

Me.Hide

Set r = dbData.OpenRecordset(printSQL)

Set REL_Aluguel_Listadecontratos.Relatorio.Recordset = r
REL_Aluguel_Listadecontratos.dfQuant.Caption = "QUANTIDADE: " & txtCONquant.Text
REL_Aluguel_Listadecontratos.dfTotal.Caption = "TOTAL: " & txtCONtotal.Text
REL_Aluguel_Listadecontratos.lblTitulo.Caption = "RELATÓRIO - CONTRATOS DE ALUGUEIS"

'If cboFiltro.Text = "TODOS" Then
'   REL_Aluguel_Listadecontratos.dfTipo.Caption = "Tipo: Todos os registros"
'ElseIf cboFiltro.Text = "PERIODO" Then
'   REL_Aluguel_Listadecontratos.dfTipo.Caption = "Tipo: Intervalo de " & Mask1.Text & " à " & Mask2.Text
'ElseIf cboFiltro.Text = "MÊS" Then
'   REL_Aluguel_Listadecontratos.dfTipo.Caption = "Tipo: Mês = " & cboMes.Text & "/" & cboAno.Text
'ElseIf cboFiltro.Text = "CLIENTE" Then
'   REL_Aluguel_Listadecontratos.dfTipo.Caption = "Cliente = " & cboNome.Text
'Else
'   REL_Aluguel_Listadecontratos.dfTipo.Caption = "Tipo:"
'End If

REL_Aluguel_Listadecontratos.Relatorio.NomeImpressora = var_Impressora
REL_Aluguel_Listadecontratos.Relatorio.Ativar
Unload REL_Aluguel_Listadecontratos

Me.Show 1
End Sub

Private Sub cmdImprimirContato_Click()
Dim r As ADODB.Recordset
Dim r_empresa As ADODB.Recordset
Dim r_cliente As ADODB.Recordset
Dim r_contrato As ADODB.Recordset
Dim r_Itens As ADODB.Recordset

If lblCodigo.Caption = "" Or txtCodCliente.Text = "" Then Exit Sub
'colocar o nome da maquina na barra de status
'Dim var_Impressora As String
'Dim oIni As Ini

'Set oIni = New Ini
'oIni.Arquivo = appPathApp & "config.ini"
'var_Impressora = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")
'Set oIni = Nothing

sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
Set r_empresa = dbData.OpenRecordset(sSQL)

sSQL = "SELECT nome, codigo, endereco, bairro, cidade, estado, numero, cpf FROM cliente where codigo = " & txtCodCliente.Text & ";"
Set r_cliente = dbData.OpenRecordset(sSQL)

sSQL = "SELECT CIDADEOBRA, obra FROM aluguel_cadastro WHERE (CODIGO = " & lblCodigo.Caption & ");"
Set r_contrato = dbData.OpenRecordset(sSQL)

sSQL = "SELECT *, aluguel_cadastro_itens.cod_equip as varCodEquip, aluguel_cadastro_itens.QUANT_ALUGADA as vQuant FROM aluguel_cadastro_itens INNER JOIN aluguel_cadastro_equipamento ON aluguel_cadastro_itens.COD_EQUIP = aluguel_cadastro_equipamento.COD_EQUIP WHERE (COD_LOCACAO = " & lblCodigo.Caption & ") and EXCLUIDO = 0;"
Set r_Itens = dbData.OpenRecordset(sSQL)
'Debug.Print sSQL

Me.Hide

'Set r = dbData.OpenRecordset(printSQL)

Set REL_ContratoAluguel.Relatorio.Recordset = r_Itens

If Not IsNull(r_empresa("caminho")) Then
      If Dir$(r_empresa("caminho")) <> "" Then Set REL_ContratoAluguel.imgLogo.Picture = LoadPicture(r_empresa("caminho"))
End If

REL_ContratoAluguel.lbl1.Caption = "A " & r_empresa("razao") & " com CNPJ " & r_empresa("cnpj") & " sedeada no endereço à " & r_empresa("endereco") & ", nº " & r_empresa("numero") & ", bairro " & r_empresa("bairro") & " na cidade " & r_empresa("cidade") & "-" & r_empresa("estado") & " doravante denominado LOCADOR e do outro lado o(a) " & r_cliente("nome") & " com sede no município " & r_cliente("cidade") & "-" & r_cliente("estado") & ", estabelecida na sua residência, " & r_cliente("endereco") & ", Bairro: " & r_cliente("bairro") & ", inscrita no CNPJ/CPF sob o nº " & r_cliente("cpf") & ". doravante denominada LOCATÁRIO, ambas as partes aqui representadas por quem de direito, tem justo e contratado entre si a locação dos equipamentos abaixo discriminados; mediante as cláusulas e condições a seguir estipuladas."
REL_ContratoAluguel.lbl2.Caption = "1.1 O equipamento ora locado, será utilizado pelo próprio LOCADOR para exercer suas funções na obra " & UCase(r_contrato("obra")) & " localizada em sua residência no MUNICIPIO DE " & UCase(r_contrato("CIDADEOBRA")) & " a serviço do LOCATÁRIO."
REL_ContratoAluguel.lbl3.Caption = "2.1 A locatária pagará ao locador a quantia de R$ " & lblSomaReferente.Caption & " por " & GridProdutos.TextMatrix(1, 6) & " dia(s) por tempo determinado, com reajuste Semestral, de acordo com a variação do IPC, ou outro índice que estiver em vigor autorizado pelo Governo Federal. O aluguel mensal constitui o pagamento pelo uso do equipamento e será devido a partir do dia da assinatura do presente. "
REL_ContratoAluguel.lbl4.Caption = "4.1 O presente contrato é estabelecido por prazo " & GridProdutos.TextMatrix(1, 6) & " dia(s) podendo ser renovado automaticamente por igual período tantas vezes forem necessários."
REL_ContratoAluguel.lbl5.Caption = "" & r_empresa("cidade") & "-" & r_empresa("estado") & ", " & Format(GridProdutos.TextMatrix(1, 14), "dd") & " de " & Format(GridProdutos.TextMatrix(1, 14), "mmmm") & " de " & Format(GridProdutos.TextMatrix(1, 14), "yyyy") & "."
REL_ContratoAluguel.lbl6.Caption = "" & r_empresa("razao") & ""
REL_ContratoAluguel.lbl7.Caption = "CNPJ " & r_empresa("cnpj") & ""
REL_ContratoAluguel.lbl8.Caption = "" & r_cliente("NOME") & ""
REL_ContratoAluguel.lbl9.Caption = "CPF/CNPJ " & r_cliente("CPF") & ""
REL_ContratoAluguel.lblContrato.Caption = "CONTRATO Nº " & lblCodigo.Caption & "/2020"
'REL_ContratoAluguel.lbl3.Caption = " "

'REL_ContratoAluguel.lblTitulo.Caption = "RELATÓRIO DE CAIXA - VENDAS À PRAZO"
'REL_ContratoAluguel.rfQuant.Caption = lbl1.Caption
'REL_ContratoAluguel.rfSubTotal.Caption = Format(lbl2.Caption, ocMONEY)
'REL_ContratoAluguel.rfEntrada.Caption = Format(lbl3.Caption, ocMONEY)
'REL_ContratoAluguel.rftotal.Caption = Format(lbl4.Caption, ocMONEY)

'REL_ContratoAluguel.rfData.Caption = Format(StatusBar1.Panels(5).Text, "dd/mm/yy")
'REL_ContratoAluguel.rfCodCaixa.Caption = varFluxoCodCaixa
'REL_ContratoAluguel.rfNomeCaixa.Caption = varFluxoNomeCaixa

'REL_ContratoAluguel.Relatorio.NomeImpressora = var_Impressora
REL_ContratoAluguel.Relatorio.Ativar
Unload REL_ContratoAluguel

Me.Show 1

If r_Itens.State <> 0 Then r_Itens.Close
Set r_Itens = Nothing
End Sub

Private Sub cmdInicial_Click()
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

Mask1 = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdNovo_Click()
Limpar_Objetos
LimparObjetosItens
txtTipoCobranca.Text = "DIA"
LimparGridProdutos
SSTab1.Tab = 0

AutoNumeracao_Pedidos
dbData.Execute "INSERT INTO pedidos (cod_pedido, tipo_pedido, status_pedido, orcamento, reaberto, cancelado) VALUES (" & lblCodPedido.Caption & ", 'ALUGUEL', 0, 0, 0, 0);"

AutoNumeracao_Aluguel
dbData.Execute "INSERT INTO aluguel_cadastro (codigo, status, cod_pedido) VALUES (" & lblCodigo.Caption & ", 0, " & lblCodPedido.Caption & ");"

frmCliente.Enabled = True
frmReferente.Enabled = True

cmdNovo.Enabled = False
cmdSalvar.Enabled = True
cmdCancelar.Enabled = True
cmdImprimirContato.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False

cmdAdicionarProduto.Enabled = True
cmdRemoverProduto.Enabled = True
cmdDevolver.Enabled = False
frmDevolucao.Visible = False
frmPagamento.Visible = False
cboSituacao.Text = "ABERTO"

LimparGrid_Parcelas
'frmDevolucao.Visible = False
frmProrrogacao.Visible = False
frmCodicoes.Enabled = False
End Sub

Private Sub AutoNumeracao_Pedidos()
'Dim sSQL As String
'Dim r As ADODB.Recordset

sSQL = "SELECT ISNULL(MAX(cod_pedido), 0) AS ultimo FROM pedidos;"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then lblCodPedido.Caption = Format(r("ultimo") + 1, "000000")
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub


Private Sub cmdProrrogacao_Click()
If txtDiasProrrogar.Text = 0 Then Exit Sub

    If CDate(mskDataProrrogar.Text) < CDate(mskDataFinalLocacaoProrro.Text) Then
        ShowMsg "A data é inferior ao prorrogação anterior!", vbInformation
        Exit Sub
    ElseIf CDate(mskDataInicioProrro.Text) = CDate(mskDataProrrogar.Text) Then
        ShowMsg "Já existe uma prorrogação com essa data!", vbInformation
        Exit Sub
    End If

'saber a quantidades de dias entre duas datas
If Not IsDate(mskDataInicialLocacaoProrro) Then Exit Sub
If Not IsDate(mskDataProrrogar) Then Exit Sub

date1 = CDate(mskDataInicialLocacaoProrro.Text)
date2 = CDate(mskDataProrrogar.Text)

vQuantDias = DateDiff("d", date1, date2)
vQuantDias = vQuantDias

If mskDataInicialLocacaoProrro.Text = mskDataProrrogar.Text Then
    vQuantDias = 1
Else
    vQuantDias = vQuantDias
End If

Gerar_ParcelasProrrogacao
MostrarParcelas

i = GridProdutos.Row
'colocar a quantidade de itens alugados
dbData.Execute "UPDATE Aluguel_Cadastro_Itens SET DATA_FINAL = CONVERT(DATETIME, '" & Format(mskDataProrrogar, ocDATA) & "', 103), QUANT = " & vQuantDias & ", VALOR = (TOTAL_ALUGADA * " & vQuantDias & ") , SUBTOTAL = (TOTAL_ALUGADA * " & vQuantDias & ") - DESCONTO, VALOR_FINAL = ((TOTAL_ALUGADA * " & vQuantDias & ") - DESCONTO) - Entrada, PRORROGADO = 1 WHERE (item = " & GridProdutos.TextMatrix(i, 1) & ") and (COD_LOCACAO = " & lblCodigo.Caption & ");"
PreencherGridProdutos

LimparObjetos_Prorrogacao
LimparGrid_Parcelas
frmProrrogacao.Visible = False
End Sub

Private Sub cmdProrrogar_Click()
frmProrrogacao.Visible = True
frmProrrogacao.Caption = "Prorrogativa"
frmDevolucao.Visible = False
frmPagamento.Visible = False
cmdAdiacao.Visible = False
cmdProrrogacao.Visible = True
LimparObjetos_Prorrogacao

i = GridProdutos.Row
mskDataInicialLocacaoProrro.Text = GridProdutos.TextMatrix(i, 14)
mskDataFinalLocacaoProrro.Text = GridProdutos.TextMatrix(i, 16)
txtEntradaAdiar.Text = Format(GridProdutos.TextMatrix(i, 10), "##,##0.00")
mskDataInicioProrro.Text = Format(DateAdd("d", Val(0), mskDataFinalLocacaoProrro.Text), "dd/mm/yy")

'calcular o periodo a ser prorrogado
Dim date1 As Date
Dim date2 As Date
Dim Result As Integer

If Not IsDate(mskDataInicialLocacaoProrro) Then Exit Sub
If Not IsDate(mskDataFinalLocacaoProrro) Then Exit Sub

date1 = CDate(mskDataInicialLocacaoProrro.Text)
date2 = CDate(mskDataFinalLocacaoProrro.Text)

Result = DateDiff("d", date1, date2)

If date1 = date2 Then
    Result = Result + 1
Else
    Result = Result
End If

Dim vDataProrrogarSugerida As Date
vDataProrrogarSugerida = DateAdd("d", Val(Result), mskDataInicioProrro.Text)

If vDataProrrogarSugerida < Date Then
    vDataProrrogarSugerida = Date
End If

Result = DateDiff("d", CDate(mskDataInicioProrro.Text), vDataProrrogarSugerida)
If Result < 1 Then Result = 1

txtDiasProrrogar.Text = Result
mskDataProrrogar.Text = Format(vDataProrrogarSugerida, "dd/mm/yy")

'calcular valor da diaria
Dim vValorDiariaProro As Currency
vValorDiariaProro = GridProdutos.TextMatrix(i, 5)  'valor de 1 diaria
Dim vDescProrro As Double
Dim vSubtotalProrro As Currency
vSubtotalProrro = vValorDiariaProro * Result
txtValorAluguellProrro.Text = Format(vSubtotalProrro, "##,##0.00")   'valor do subtotal

'calcular o desconto
Dim vValorDescProrro As Currency
If txtValorRealDescProrro.Text = "" Or txtValorRealDescProrro.Text = "0,00" Then
    vValorDescProrro = 0
    txtDescProrro.Text = "0,00"
    txtValorRealDescProrro.Text = "0,00"
Else
    vValorDescProrro = txtValorRealDescProrro.Text
End If

Dim vTotalProrro As Currency
vTotalProrro = txtValorAluguellProrro - vValorDescProrro
txtTotalProrro = Format(vTotalProrro, "##,##0.00")

'frmProrrogacao.Visible = True
'frmDevolucao.Visible = False
'cmdAdiacao.Visible = False
'cmdProrrogacao.Visible = True
'LimparObjetos_Prorrogacao

'i = GridProdutos.Row
'mskDataInicialLocacaoProrro.Text = GridProdutos.TextMatrix(i, 14)
'mskDataFinalLocacaoProrro.Text = GridProdutos.TextMatrix(i, 16)
'mskDataInicioProrro.Text = Format(DateAdd("d", Val(0), mskDataFinalLocacaoProrro.Text), "dd/mm/yy")
'mskDataProrrogar.Text = Format(Date, "dd/mm/yy")

'txtEntradaAdiar.Text = Format(GridProdutos.TextMatrix(i, 10), "##,##0.00")

'varCodItem = GridProdutos.TextMatrix(i, 1)
''txtDescProrro.Text = GridProdutos.TextMatrix(i, 8)
''txtValorRealDescProrro.Text = GridProdutos.TextMatrix(i, 10)
''txtValorAluguellProrro.Text = GridProdutos.TextMatrix(i, 5)

''calcular a quantidade de dias
''calculardiasAdiar

''calcular o valor final com descont
'calculardiasAdiar
'CalcularTotalAdiar
End Sub

Private Sub cmdRemoverProduto_Click()

i = GridProdutos.Row

If GridProdutos.TextMatrix(i, 13) = "SIM" Then
    MsgBox "Esse equipamento já foi devolvido!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

If GridProdutos.TextMatrix(i, 18) = "SIM" Then
    MsgBox "Esse equipamento já foi removido do contrato!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

'If GridProdutos.TextMatrix(i, 2) = "" Then Exit Sub
If ShowMsg("Deseja remover o equipamento: " & GridProdutos.TextMatrix(i, 2) & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub

'marcar o equipamento como excluido
dbData.Execute "UPDATE Aluguel_Cadastro_Itens SET excluido =  1 WHERE (item = " & GridProdutos.TextMatrix(i, 1) & ") and (COD_LOCACAO = " & lblCodigo.Caption & ");"

'marcar o equipamento como livre
'dbData.Execute "UPDATE aluguel_cadastro_equipamento SET contrato =  0, alugado = 0 WHERE (COD_EQUIP = " & GridProdutos.TextMatrix(i, 13) & ");"

'Remover a parcela desse equipamento
dbData.Execute "DELETE FROM parcelas WHERE (cod_os = " & lblCodigo.Caption & ") and OS_ITEM = " & GridProdutos.TextMatrix(i, 1) & " ;"

'atualizar os equipamentos alugados
i = GridProdutos.Row

sSQL = "SELECT QUANT_ESTOQUE, QUANT_ALUGADA FROM Aluguel_Cadastro_Equipamento WHERE (COD_EQUIP = " & GridProdutos.TextMatrix(i, 12) & ");"
Set r = dbData.OpenRecordset(sSQL)

Dim varQuantAtual As Integer
Dim varQuantAlugada As Integer
Dim varNovaQuant As Integer

varQuantAtual = r("QUANT_ALUGADA")
varQuantAlugada = GridProdutos.TextMatrix(i, 4)
varNovaQuant = varQuantAtual - varQuantAlugada

'colocar a quantidade de itens alugados
dbData.Execute "UPDATE aluguel_cadastro_equipamento SET QUANT_ALUGADA =  " & varNovaQuant & " WHERE (COD_EQUIP = " & GridProdutos.TextMatrix(i, 12) & ");"

'Remover o equipamento
'dbData.Execute "DELETE FROM Aluguel_Cadastro_Itens WHERE (item = " & GridProdutos.TextMatrix(i, 1) & ") and (COD_LOCACAO = " & lblCodigo.Caption & ");"

PreencherGridProdutos
End Sub

Private Sub cmdSalvar_Click()
If lblCodigo.Caption = "" Or txtCliente.Text = "" Or lblSomaReferente.Caption = "" Then Exit Sub

If cboSituacao.Text = "ABERTO" Then
    vSituacao = False
ElseIf cboSituacao.Text = "FECHADO" Then
    vSituacao = True
End If

dbData.Execute "UPDATE aluguel_cadastro SET CODIGO = " & lblCodigo.Caption & ", cod_cliente = " & txtCodCliente.Text & ", OBRA = '" & txtDescricaoObra.Text & "', CIDADEOBRA = '" & cboCidadeObra.Text & "', " & _
               "STATUS = '" & vSituacao & "', VALORINICIAL =  " & Replace(CCur(lblSomaReferente.Caption), ",", ".") & ", DATACADASTRO = CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103) " & _
               "WHERE (CODIGO = " & lblCodigo.Caption & ");"

dbData.Execute "UPDATE pedidos SET cod_cliente = " & txtCodCliente.Text & ", " & _
      "data_compra = CONVERT(DATETIME, '" & Format(Date, ocDATA) & "', 103), tipo_desc = 'P', valor_desc = 0, " & _
      "subtotal = " & Replace(CCur(lblSomaReferente.Caption), ",", ".") & ", total = " & Replace(CCur(lblSomaReferente.Caption), ",", ".") & ", " & _
      "tipo_pagamento = 'À Prazo', pagamento = 'AVULSO', " & _
      "status_pedido = 1, " & _
      "maquina = '" & StatusBar1.Panels(2).Text & "', CodCaixa = " & StatusBar1.Panels(4).Text & ", CAIXA = '" & StatusBar1.Panels(3).Text & "', Entrada = 0, COD_FUNCIONARIO = " & txtCodFuncionario.Text & ", TIPO_ACRESCIMO = 'P', VALOR_ACRESCIMO = 0, ValorDescReal = 0, ValorAcrescReal = 0 WHERE (cod_pedido = " & lblCodPedido.Caption & ");"
      
    
'colocar equipamentos como alugado e o numero de contrato
'For i = 1 To GridProdutos.Rows - 1
'   dbData.Execute "UPDATE aluguel_cadastro_equipamento SET contrato =  " & lblCodigo.Caption & ", alugado = 1 WHERE (COD_EQUIP = " & GridProdutos.TextMatrix(i, 13) & ");"
'Next

Limpar_Objetos
LimparGrid_Parcelas
Form_Load
cmdExibirConsulta_Click
End Sub

Private Sub Calcular_Prazo()
'If cboPrazo.Text = "" Then Exit Sub
If Not IsDate(mskInicio.Text) Then Exit Sub
If mskHoraInicio.Text = "" Then Exit Sub

'If optMultiplica.Value = True Then
   'mskInicio.Text = Format(mskCompra, "dd/mm/yy")
mskFinal.Text = Format(DateAdd("d", Val(1) - 1, mskInicio.Text), "dd/mm/yy")
If mskHoraFinal.Text = "" Then mskHoraFinal.Text = mskHoraInicio
mskFinal_LostFocus
'ElseIf optDivide.Value = True Then
   'mskInicio.Text = Format(mskCompra, "dd/mm/yy")
   'mskFinal.Text = Format(DateAdd("m", Val(cboQuantParc.Text), mskInicio.Text), "dd/mm/yy")'
'   mskFinal.Text = Format(DateAdd("m", Val(cboQuantParc.Text) - 1, mskInicio.Text), "dd/mm/yy")
'End If
End Sub

Function GetHora(ByVal totalMinutos As Long) As String
GetHora = CDbl(CLng(totalMinutos / 60)) '& ":00"
End Function


Private Sub Form_Load()
Set moCombo = New cComboHelper
cmdNovo.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
'cmdImprimirConsulta.Enabled = False
cmdImprimirContato.Enabled = False
cmdAdicionarProduto.Enabled = False
cmdRemoverProduto.Enabled = False
cmdDevolver.Enabled = False
cmdAdiar.Enabled = False
cmdDevolverParcial.Enabled = False
cmdProrrogar.Enabled = False
'frmDevolucao.Visible = False
frmProrrogacao.Visible = False
frmCodicoes.Enabled = False

LimparGridProdutos
PreencherComboStatus
SSTab1.Tab = 0
cboFiltro.Text = "TODOS"
vTipoDevolver = 0
cmdExibirConsulta_Click
StatusBar1.Panels(5).Text = Format(Date, "dd/mm/yy")

If vCodFunc = 0 Then
    txtCodFuncionario.Text = "1"
Else
    txtCodFuncionario.Text = vCodFunc
End If

'colocar o nome da maquina na barra de status
Dim var_Caixa As String
Dim var_Maquina As String

'abre o ini
Dim oIni As Ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"

'nome da caixa
var_Caixa = oIni.LerTexto("DADOS_CAIXA", "caixa")
StatusBar1.Panels(3).Text = var_Caixa

'nome da caixa
var_Maquina = oIni.LerTexto("DADOS_MAQUINA", "maquina")
StatusBar1.Panels(2).Text = var_Maquina
Set oIni = Nothing
'StatusBar1.Panels(4).Text = Format(ValidateNull(r("codcaixa")), "00000")

MostrarCodCaixa
'ConsultarCaixaAtual
End Sub
Private Sub MostrarCodCaixa()
sSQL = "SELECT *, CASE status WHEN 0 THEN 'ABERTO' ELSE 'FECHADO' END AS varStatus " & _
       "FROM caixa_dia " & _
       "WHERE (caixa = '" & StatusBar1.Panels(3).Text & "') and caixa_dia.status = 0;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then
    varCodCaixa = ValidateNull(r("codcaixa"))
    'cmdImprimir.Enabled = True
    'cmdImprimirResumido.Enabled = True
    'cmdAbrirCaixa.Visible = False
    'cmdFecharCaixa.Visible = True
    'cmdTroco.Enabled = True
    StatusBar1.Panels(4).Text = Format(ValidateNull(r("codcaixa")), "00000")
    'StatusBar1.Panels(4).Text = r("VARSTATUS")
    'vStatusCaixaAtual = r("VARSTATUS")
Else
    'If varFluxoCaixa = False Then
        varCodCaixa = 0
        'cmdImprimir.Enabled = False
        'cmdImprimirResumido.Enabled = False
        'cmdAbrirCaixa.Visible = True
        'cmdFecharCaixa.Visible = False
        'cmdTroco.Enabled = False
        StatusBar1.Panels(4).Text = Format(0, "00000")
        'StatusBar1.Panels(4).Text = "FECHADO"
        'vStatusCaixaAtual = "FECHADO"
    'Else
    '    varCodCaixa = StatusBar1.Panels(3).Text
    'End If
End If
End Sub
Private Sub ConsultarCaixaAtual()
sSQL = "SELECT * " & _
       "FROM caixa_dia " & _
       "WHERE (caixa = '" & StatusBar1.Panels(3).Text & "') and caixa_dia.status = 0;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then
    varCodCaixa = ValidateNull(r("codcaixa"))
Else
    varCodCaixa = 0
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If cmdSalvar.Enabled = True Then
    MsgBox "O Aluguel iniciado ainda não foi salvo", vbInformation, "Aviso do Sistema"
    Cancel = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set moCombo = Nothing
End Sub

Private Sub MostrarParcelas()
i = GridProdutos.Row
'mskDataInicialLocacaoProrro.Text = GridProdutos.TextMatrix(i, 14)

If lblCodigo.Caption = "" Then Exit Sub
If GridProdutos.TextMatrix(i, 1) = "" Then Exit Sub

If GridProdutos.rows >= 2 Then
    sSQL = "SELECT DATA, PAGAMENTO, VALOR, DESCONTO, VALOR_FINAL, CASE status WHEN 0 THEN 'À PAGAR' ELSE 'PAGO' END AS varStatus, FORMA_PGTO, CODCAIXA, CAIXA " & _
       "FROM parcelas WHERE (cod_os = " & lblCodigo.Caption & ") and OS_ITEM = " & GridProdutos.TextMatrix(i, 1) & " ORDER BY numero;"
    
    Set r = dbData.OpenRecordset(sSQL)
    
    FormatarGrid_Parcelas r
    
    If r.State <> 0 Then r.Close
    Set r = Nothing
End If
End Sub
Private Sub LimparGrid_Parcelas()
Dim i As Integer

With grid_Parcelas
   .Clear
   .Cols = 8
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 800
   .ColWidth(2) = 800
   .ColWidth(3) = 850
   .ColWidth(4) = 800
   .ColWidth(5) = 1000
   .ColWidth(6) = 800
   .ColWidth(7) = 800
   
   .TextMatrix(0, 1) = "Venc."
   .TextMatrix(0, 2) = "Valor"
   .TextMatrix(0, 3) = "Status"
   .TextMatrix(0, 4) = "Pgto"
   .TextMatrix(0, 5) = "Forma."
   .TextMatrix(0, 6) = "Caixa"
   .TextMatrix(0, 7) = "Cód.Cx"
   
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
   Next i
   
   .rows = .rows - 1
End With
End Sub

Private Sub GridConsulta_Click()
i = GridConsulta.Row
If GridConsulta.TextMatrix(i, 7) = "0,00" And GridConsulta.TextMatrix(i, 10) = "0" Then
    cmdFecharAluguel.Enabled = True
Else
    cmdFecharAluguel.Enabled = False
End If
End Sub

Private Sub gridConsulta_DblClick()

i = GridConsulta.Row
If GridConsulta.TextMatrix(i, 7) = "SIM" Then
    cmdNovo.Enabled = True
    cmdSalvar.Enabled = False
    cmdCancelar.Enabled = False
    cmdAlterar.Enabled = False
    cmdExcluir.Enabled = False
    frmCliente.Enabled = True
    frmReferente.Enabled = False
    cmdImprimirContato.Enabled = False
    cmdAdicionarProduto.Enabled = False
    cmdRemoverProduto.Enabled = False
    'cmdDevolver.Enabled = False
Else
    cmdNovo.Enabled = True
    cmdSalvar.Enabled = False
    cmdCancelar.Enabled = False
    cmdAlterar.Enabled = True
    cmdExcluir.Enabled = True
    frmCliente.Enabled = True
    frmReferente.Enabled = True
    cmdImprimirContato.Enabled = True
    cmdAdicionarProduto.Enabled = True
    cmdRemoverProduto.Enabled = True
   'cmdDevolver.Enabled = True
End If

frmProrrogacao.Visible = False
frmDevolucao.Visible = False
frmPagamento.Visible = False
frmCodicoes.Enabled = False

txtTipoCobranca.Text = "DIA"
lblCodigo.Caption = ""
lblCodigo.Caption = GridConsulta.TextMatrix(GridConsulta.RowSel, 1)

SSTab1.Tab = 0
End Sub

Private Sub GridProdutos_Click()
LimparObjetos_Prorrogacao
LimparObjetosDevolucao
frmProrrogacao.Visible = False
frmDevolucao.Visible = False
frmPagamento.Visible = False

i = GridProdutos.Row
'mskDataInicialLocacaoProrro.Text = GridProdutos.TextMatrix(i, 14)

If lblCodigo.Caption = "" Then Exit Sub
If GridProdutos.TextMatrix(i, 1) = "" Then Exit Sub
    
If GridProdutos.rows >= 2 Then
i = GridProdutos.Row
    sSQL = "SELECT DATA, PAGAMENTO, VALOR, DESCONTO, VALOR_FINAL, CASE status WHEN 0 THEN 'À PAGAR' ELSE 'PAGO' END AS varStatus, FORMA_PGTO, CODCAIXA, CAIXA " & _
       "FROM parcelas WHERE (cod_os = " & lblCodigo.Caption & ") and OS_ITEM = " & GridProdutos.TextMatrix(i, 1) & " order by numero;"
    'Debug.Print sSQL
    Set r = dbData.OpenRecordset(sSQL)
    
    FormatarGrid_Parcelas r
    
    If r.State <> 0 Then r.Close
    Set r = Nothing
End If

If GridProdutos.TextMatrix(i, 13) = "SIM" Or GridProdutos.TextMatrix(i, 18) = "SIM" Then
    cmdDevolver.Enabled = False
    cmdDevolverParcial.Enabled = False
    cmdProrrogar.Enabled = False
    cmdAdiar.Enabled = False
Else
    cmdDevolver.Enabled = True
    cmdDevolverParcial.Enabled = True
    cmdProrrogar.Enabled = True
    cmdAdiar.Enabled = True
End If
End Sub

Private Sub FormatarGrid_Parcelas(rTabela As ADODB.Recordset)
Dim i As Integer

With grid_Parcelas
   .Clear
   .Cols = 10
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 800
   .ColWidth(2) = 800
   .ColWidth(3) = 850
   .ColWidth(4) = 800
   .ColWidth(5) = 1000
   .ColWidth(6) = 800
   .ColWidth(7) = 900
   .ColWidth(8) = 0
   .ColWidth(9) = 0
   
   .TextMatrix(0, 1) = "Venc."
   .TextMatrix(0, 2) = "Valor"
   .TextMatrix(0, 3) = "Desc"
   .TextMatrix(0, 4) = "Total"
   .TextMatrix(0, 5) = "Status"
   .TextMatrix(0, 6) = "Pgto"
   .TextMatrix(0, 7) = "Forma."
   .TextMatrix(0, 8) = "Caixa"
   .TextMatrix(0, 9) = "Cód.Cx"
   
   'colocar os cabeçalho em negrito
   For i = 0 To .Cols - 1
      .Col = i
      .Row = 0
      .CellFontBold = True
   Next
   
   'ALINHAMENTO
   '.ColAlignment(2) = 1
   
   'centralizar o titulo
   For i = 0 To .Cols - 1
      .Row = 0
      .Col = i
      .CellAlignment = flexAlignCenterCenter
   Next i
   
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         .TextMatrix(.rows - 1, 1) = Format(rTabela("DATA"), "dd/mm/yy")
         .TextMatrix(.rows - 1, 2) = Format(rTabela("VALOR"), ocMONEY)
         .TextMatrix(.rows - 1, 3) = Format(rTabela("DESCONTO"), ocMONEY)
         .TextMatrix(.rows - 1, 4) = Format(rTabela("VALOR_FINAL"), ocMONEY)
         .TextMatrix(.rows - 1, 5) = rTabela("varSTATUS")
         .TextMatrix(.rows - 1, 6) = Format(rTabela("PAGAMENTO"), "dd/mm/yy")
         .TextMatrix(.rows - 1, 7) = ValidateNull(rTabela("FORMA_PGTO"))
         .TextMatrix(.rows - 1, 8) = ValidateNull(rTabela("CAIXA"))
         .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("CODCAIXA"))
         rTabela.MoveNext
         .rows = .rows + 1
      Loop
   End If
   
   .rows = .rows - 1
End With
End Sub

Private Sub lblCodigo_Change()
If lblCodigo.Caption = "" Then Exit Sub

'If cmdExcluir.Enabled = True Then
sSQL = "SELECT CODIGO, COD_CLIENTE, OBRA, CIDADEOBRA, STATUS, EXCLUIDO, COD_PEDIDO, STATUS FROM aluguel_cadastro WHERE (CODIGO = " & lblCodigo.Caption & ");"
Set r = dbData.OpenRecordset(sSQL)

 If Not r.BOF Then
     If CBool(r("STATUS")) = True Then
         cboSituacao.Text = "FECHADO"
     Else
         cboSituacao.Text = "ABERTO"
     End If
     txtDescricaoObra.Text = ValidateNull(r("obra"))
     cboCidadeObra.Text = ValidateNull(r("CIDADEOBRA"))
     lblCodPedido.Caption = r("COD_PEDIDO")
     txtCodCliente.Text = r("cod_cliente")
End If

If r.State <> 0 Then r.Close
Set r = Nothing

LimparGridProdutos
PreencherGridProdutos
LimparGrid_Parcelas
'End If
End Sub

Private Sub lblSomaReferente_Change()
'txtTotal.Text = lblSomaReferente.Caption
'cboQuantParc = 1
'txtParc_GotFocus
End Sub

Private Sub MASK1_KeyPress(KeyAscii As Integer)
   Mask1.Mask = "##/##/##"
End Sub

Private Sub Mask1_LostFocus()
   If Mask1.Text = "__/__/__" Then
      Mask1.Mask = ""
      Mask1.Text = ""
      Exit Sub
   ElseIf Mask1.Text = "" Then
      Mask1.Mask = ""
      Mask1.Text = ""
      Exit Sub
   ElseIf Not IsDate(Mask1) Then
      ShowMsg "Data Inválida", vbExclamation
      Mask1.SetFocus
      Exit Sub
   End If
End Sub

Private Sub Mask2_KeyPress(KeyAscii As Integer)
   Mask2.Mask = "##/##/##"
End Sub

Private Sub Mask2_LostFocus()
   If Mask2.Text = "__/__/__" Then
      Mask2.Mask = ""
      Mask2.Text = ""
      Exit Sub
   ElseIf Mask2.Text = "" Then
      Mask2.Mask = ""
      Mask2.Text = ""
      Exit Sub
   ElseIf Not IsDate(Mask2) Then
      ShowMsg "Data Inválida", vbExclamation
      Mask2.SetFocus
      Exit Sub
   End If
End Sub

Private Sub mskDataFinalLocacaoHora_KeyPress(KeyAscii As Integer)
mskDataFinalLocacaoHora.Mask = "##:##"
End Sub


Private Sub mskDataProrrogar_Change()
mskDataProrrogar_LostFocus
End Sub

Private Sub mskDataProrrogar_LostFocus()
If cmdProrrogacao.Visible = True Then
    'calcular os dias
    Dim Result As Integer
    
    If Not IsDate(mskDataInicioProrro) Then Exit Sub
    If Not IsDate(mskDataProrrogar) Then Exit Sub
    
    date1 = CDate(mskDataInicioProrro.Text)
    date2 = CDate(mskDataProrrogar.Text)
    
    Result = DateDiff("d", date1, date2)
    Result = Result
    
    If CDate(mskDataProrrogar.Text) < CDate(mskDataInicioProrro.Text) Then
        ShowMsg "A data é inferior ao prorrogação anterior!", vbInformation
        txtDiasProrrogar.Text = 0
        Exit Sub
    Else
        If CDate(mskDataInicioProrro.Text) = CDate(mskDataProrrogar.Text) Then
            txtDiasProrrogar.Text = 1
            Result = 1
        Else
            txtDiasProrrogar.Text = Result
        End If
    End If
    
    'calcular valores
    
    Dim vValorDiariaProro As Currency
    
    vValorDiariaProro = GridProdutos.TextMatrix(i, 5)  'valor de 1 diaria
    
    'calcular totais
    Dim vSubtotalProrro As Currency
    
    vSubtotalProrro = vValorDiariaProro * Result

    txtValorAluguellProrro.Text = Format(vSubtotalProrro, "##,##0.00")   'valor do subtotal
    
    'calcular o desconto
    Dim vValorDescProrro As Currency
    If txtValorRealDescProrro.Text = "" Or txtValorRealDescProrro.Text = "0,00" Then
        vValorDescProrro = 0
        txtDescProrro.Text = "0,00"
        txtValorRealDescProrro.Text = "0,00"
    Else
        vValorDescProrro = txtValorRealDescProrro.Text
    End If
    
    Dim vTotalProrro As Currency
    vTotalProrro = txtValorAluguellProrro - vValorDescProrro
    
    txtTotalProrro = Format(vTotalProrro, "##,##0.00")
ElseIf cmdAdiacao.Visible = True Then
    calculardiasAdiar
    CalcularTotalAdiar
End If
End Sub


Private Sub mskDevolver_Change()
'Dim date1 As Date
'Dim date2 As Date
'Dim Result As Integer

'If Not IsDate(mskDataFinalLocacao) Then Exit Sub
'If Not IsDate(mskDevolver) Then Exit Sub

'date1 = CDate(mskDataFinalLocacao.Text)
'date2 = CDate(mskDevolver.Text)

'Result = DateDiff("d", date1, date2)
'Result = Result

'If mskDataFinalLocacao.Text = mskDevolver.Text Then
'    txtDiasLocacao.Text = 1
'Else
'    txtDiasLocacao.Text = Result
'End If

'calcular valor
'Dim vValorAluguel As Currency
'Dim vQuantDias As Integer

'vValorAluguel = txtTotalDev.Text
'vQuantDias = txtDiasLocacao.Text

'txtSubtotalLocacao.Text = (vValorAluguel * vQuantDias)

'calcular o valor restante
Dim vValorEntrada As Currency

'vValorEntrada = txtEntradaLocacao.Text
'txtRestanteLocacao.Text = (txtTotalDev - vValorEntrada)
End Sub

Private Sub mskDevolver_GotFocus()
SelectControl mskDevolver
End Sub

Private Sub mskDevolver_LostFocus()
If mskDevolver.Text = "" Or mskDevolver.Text = "__/__/__" Then
   mskDevolver.Mask = ""
   mskDevolver.Text = ""
   Exit Sub
Else
    If Not IsDate(mskDevolver.Text) Then
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      SelectControl mskDevolver
   End If
End If
End Sub


Private Sub mskDevolverHora_GotFocus()
SelectControl mskDevolverHora
End Sub

Private Sub mskDevolverHora_KeyPress(KeyAscii As Integer)
mskDevolverHora.Mask = "##:##"
End Sub


Private Sub mskFinal_Change()
'If mskFinal.Text <> "" Then mskFinal_LostFocus
End Sub

Private Sub mskFinal_GotFocus()
SelectControl mskFinal
End Sub


Private Sub mskFinal_KeyPress(KeyAscii As Integer)
mskFinal.Mask = "##/##/##"
End Sub


Private Sub mskFinal_LostFocus()
If mskFinal.Text = "" Or mskFinal.Text = "__/__/__" Then
   mskFinal.Mask = ""
   mskFinal.Text = ""
   Exit Sub
Else
    If Not IsDate(mskFinal.Text) Then
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      SelectControl mskFinal
   End If
End If

'If optPorDatas.Value = True Then
'    cboQuantParc.Enabled = False
    'chameleonButton2.Enabled = True
    'mskFinal.Enabled = True
'    CalcularQuantDias
'Else
'    cboQuantParc.Enabled = True
    'chameleonButton2.Enabled = False
    'mskFinal.Enabled = False
'End If

CalcularQuantDias
CalcularDiaria
End Sub


Private Sub mskHoraFinal_GotFocus()
SelectControl mskHoraFinal
End Sub


Private Sub mskHoraFinal_KeyPress(KeyAscii As Integer)
mskHoraFinal.Mask = "##:##"
End Sub

Private Sub mskHoraFinal_LostFocus()
If txtTipoCobranca.Text = "HORA" Then
    CalcularQuantDias
    CalcularDiaria
End If
End Sub


Private Sub mskHoraInicio_GotFocus()
SelectControl mskHoraInicio
End Sub


Private Sub mskHoraInicio_KeyPress(KeyAscii As Integer)
mskHoraInicio.Mask = "##:##"
End Sub


Private Sub mskHoraInicio_LostFocus()
Calcular_Prazo
End Sub


Private Sub mskInicio_Change()
Calcular_Prazo
'If mskInicio.Text <> "" Then mskInicio_LostFocus
End Sub

Private Sub mskInicio_GotFocus()
If mskInicio.Text = "" Then
    mskHoraInicio.Text = Format(Now, "hh:mm")
    mskInicio.Text = Format(Date, "dd/mm/yy")
End If

SelectControl mskInicio
End Sub

Private Sub mskInicio_KeyPress(KeyAscii As Integer)
mskInicio.Mask = "##/##/##"
End Sub

Private Sub mskInicio_LostFocus()
If mskInicio.Text = "" Or mskInicio.Text = "__/__/__" Then
   mskInicio.Mask = ""
   mskInicio.Text = ""
   Exit Sub
Else
    If Not IsDate(mskInicio.Text) Then
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      SelectControl mskInicio
   End If
End If

If IsDate(mskInicio) = True Then
    If mskFinal.Text = "" Then
        mskFinal.Text = Format(DateAdd("d", Val(1), mskInicio.Text), "dd/mm/yy")
        mskHoraFinal.Text = mskHoraInicio
    End If
End If

'If optPorDatas.Value = True Then
'    cboQuantParc.Enabled = False
'    chameleonButton2.Enabled = True
'    mskFinal.Enabled = True
'    CalcularQuantDias
'Else
'    cboQuantParc.Enabled = True
'    chameleonButton2.Enabled = False
'    mskFinal.Enabled = False
'End If
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
If cmdSalvar.Enabled = True Then
    MsgBox "O Aluguel iniciado ainda não foi salvo", vbInformation, "Aviso do Sistema"
    SSTab1.Tab = 0
End If
End Sub

Private Sub txtCliente_GotFocus()
'Dim sSQL As String
'Dim r As ADODB.Recordset
Dim itemAtual As String
Dim codAtual As String

itemAtual = txtCliente.Text
codAtual = txtCodCliente.Text
txtCliente.Clear

sSQL = "SELECT DISTINCT nome, codigo FROM cliente ORDER BY nome;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   txtCliente.AddItem r("nome")
   txtCliente.ItemData(txtCliente.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

txtCliente.Text = itemAtual
txtCodCliente.Text = codAtual
moCombo.AttachTo txtCliente
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCliente_LostFocus()
   On Error GoTo TrataErro
   If txtCliente.Text = "" Then txtCodCliente.Text = "": Exit Sub
   txtCodCliente = txtCliente.ItemData(txtCliente.ListIndex)
   Exit Sub
   
TrataErro:
  ' If Err.Number = 381 Then txtCodCliente.Text = ""
End Sub

Private Sub TxtCodCliente_Change()
If txtCodCliente.Text = "" Then Exit Sub

If cmdExcluir.Enabled = True Then
   sSQL = "SELECT codigo, nome FROM cliente WHERE (codigo = " & txtCodCliente.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then txtCliente.Text = r("nome")
   'If r.State <> 0 Then r.Close
   'Set r = Nothing
End If
End Sub

Private Sub txtDesconto_GotFocus()
SelectControl txtDesconto
End Sub

Private Sub txtDesconto_LostFocus()
If txtDesconto.Text = "" Then txtDesconto.Text = Format(0, "##,##0.00") Else txtDesconto.Text = Format(txtDesconto, "##,##0.00")
CalcularDiaria
End Sub


Private Sub txtDescProrro_Change()
txtDescProrro_LostFocus
End Sub

Private Sub txtDescProrro_GotFocus()
SelectControl txtDescProrro
End Sub


Private Sub txtDescProrro_LostFocus()
'txtValorRealDescProrro.Text = Format(((CCur(txtValorAluguellProrro) * CCur(txtDescProrro)) / 100), ocMONEY)
CalcularTotalAdiar
End Sub


Private Sub txtDescricaoObra_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtDiasProrrogar_Change()
CalcularTotalAdiar
End Sub

Private Sub txtEntrada_Change()
If txtEntrada.Text = "0,00" Or txtEntrada.Text = "" Then
    txtEntradaReal.Visible = False
    lblFormaEntrada.Visible = False
    lbl1.Visible = False
    cboFormaEntrada.Visible = False
    txtTotal.Visible = False
Else
    txtEntradaReal.Visible = True
    lblFormaEntrada.Visible = True
    lbl1.Visible = True
    cboFormaEntrada.Visible = True
    txtTotal.Visible = True
End If
End Sub

Private Sub txtEntrada_GotFocus()
SelectControl txtEntrada
End Sub


Private Sub txtEntrada_LostFocus()
If txtEntrada.Text = "" Then txtEntrada.Text = Format(0, "##,##0.00") Else txtEntrada.Text = Format(txtEntrada, "##,##0.00")
CalcularDiaria
End Sub



Private Sub txtQuantAlugada_Change()
CalcularDiaria
End Sub

Private Sub txtQuantAlugada_GotFocus()
SelectControl txtQuantAlugada
End Sub

Private Sub txtQuantAlugada_LostFocus()
CalcularValorTotal
CalcularDiaria
End Sub


Private Sub txtQuantDev_GotFocus()
SelectControl txtQuantDev
End Sub


Private Sub txtTipoCobranca_GotFocus()
txtTipoCobranca.Clear
txtTipoCobranca.AddItem "DIA"
txtTipoCobranca.AddItem "HORA"
   
If txtTipoCobranca.ListCount <> 0 Then txtTipoCobranca.ListIndex = 0
moCombo.AttachTo txtTipoCobranca
End Sub

Private Sub txtTipoCobranca_Click()
If txtCodEquip.Text <> "" Then
    If txtTipoCobranca.Text = "DIA" Then
        sSQL = "SELECT VALOR_DIA as varValorDiaAluguel FROM aluguel_cadastro_equipamento WHERE (COD_EQUIP = " & txtCodEquip.Text & ");"
    Else
        sSQL = "SELECT VALOR_HORA as varValorDiaAluguel FROM aluguel_cadastro_equipamento WHERE (COD_EQUIP = " & txtCodEquip.Text & ");"
    End If
    Set r = dbData.OpenRecordset(sSQL)
    If Not r.BOF Then txtValorAluguel.Text = Format(ValidateNull(r("varValorDiaAluguel")), "##,##0.00")
    If r.State <> 0 Then r.Close
    Set r = Nothing
End If

If txtTipoCobranca.Text = "HORA" Then
    If IsDate(mskInicio) Then mskFinal.Text = mskInicio.Text
    If mskHoraInicio.Text <> "" Then mskHoraFinal.Text = Format(TimeValue(mskHoraInicio.Text) + TimeSerial(1, 0, 0), "hh:mm")
Else
    If IsDate(mskInicio) Then mskFinal.Text = Format(DateAdd("d", 1, mskInicio.Text), "dd/mm/yy")
    mskHoraFinal.Text = mskHoraInicio.Text
End If

CalcularValorTotal
CalcularQuantDias
CalcularDiaria
End Sub


Private Sub txtTotal_Change()
'Call CalcularParcelas
'Call Calcular_Prazo
End Sub

Private Sub txtTotal_LostFocus()
   If txtTotal.Text = "" Then txtTotal = Format(0, "##,##0.00") Else txtTotal = Format(txtTotal, "##,##0.00")
End Sub

Private Sub txtValorAluguel_Change()
CalcularValorTotal
End Sub

Private Sub txtValorAluguel_LostFocus()
If txtValorAluguel.Text = "" Then txtValorAluguel = Format(0, "##,##0.00") Else txtValorAluguel = Format(txtValorAluguel, "##,##0.00")
End Sub

Private Sub txtValorRealDescProrro_Change()
Dim vValorDescProrro As Currency

If txtValorRealDescProrro.Text = "" Or txtValorRealDescProrro.Text = "0,00" Then
    vValorDescProrro = 0
    txtDescProrro.Text = "0,00"
    txtValorRealDescProrro.Text = "0,00"
Else
    vValorDescProrro = txtValorRealDescProrro.Text
End If

Dim vTotalProrro As Currency

If txtValorAluguellProrro.Text = "" Then txtValorAluguellProrro.Text = "0,00"
vTotalProrro = txtValorAluguellProrro - vValorDescProrro

txtTotalProrro = Format(vTotalProrro, "##,##0.00")
End Sub


