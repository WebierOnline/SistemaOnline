VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Frm_NFCe 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NFCe - Nota Fiscal Consumidor Eletronica"
   ClientHeight    =   10515
   ClientLeft      =   735
   ClientTop       =   1455
   ClientWidth     =   13455
   KeyPreview      =   -1  'True
   LinkTopic       =   "Frm_NFCe"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10515
   ScaleWidth      =   13455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   13365
      TabIndex        =   67
      Top             =   0
      Width           =   13395
      Begin VB.TextBox txtCodNota 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   11280
         TabIndex        =   91
         TabStop         =   0   'False
         ToolTipText     =   "Cód da Nota"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtCodPedido 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   12360
         TabIndex        =   90
         TabStop         =   0   'False
         ToolTipText     =   "Cód do Pedido"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NFCe - Nota Fiscal Consumidor Eletrônica"
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
         Left            =   1605
         TabIndex        =   68
         Top             =   180
         Width           =   6330
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   540
         Picture         =   "Frm_NFCe.frx":0000
         Top             =   0
         Width           =   750
      End
   End
   Begin TabDlg.SSTab Frm_NFCe 
      Height          =   9375
      Left            =   0
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   840
      Width           =   13395
      _ExtentX        =   23627
      _ExtentY        =   16536
      _Version        =   393216
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Nota Fiscal (NFCe)"
      TabPicture(0)   =   "Frm_NFCe.frx":090B
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdConsultar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdRecalcularNF"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdAlterar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdSalvar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdCancelarNota"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdTransmitir"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdImprimir"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdExcluir"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdSair"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdNovo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdCancelar"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "frmNota"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "frmTransporte"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "frmItens"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Consulta"
      TabPicture(1)   =   "Frm_NFCe.frx":0927
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label14"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label11"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label15"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label16"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label19"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label20"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label21"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label8"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label7"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label22"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label18"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label17"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label10"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label12"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label13"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label9"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "lblTotalNota"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "GridNotas"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Command1"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Text17"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Frame3"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Text15"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Text16"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Text19"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Text21"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Text22"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Text1"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "txtValorFrete"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "cboTransporte"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Text23"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Text9"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Text10"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Text11"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Text12"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Text13"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Text18"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "txtCodTransporte"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).ControlCount=   37
      TabCaption(2)   =   "Pedidos"
      TabPicture(2)   =   "Frm_NFCe.frx":0943
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblQuantPedidos"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "GridPedidos"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "frmFiltrosPedidos"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.TextBox txtCodTransporte 
         Height          =   285
         Left            =   -67020
         MaxLength       =   50
         TabIndex        =   117
         Top             =   1980
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox Text18 
         Height          =   315
         Left            =   -68625
         MaxLength       =   50
         TabIndex        =   116
         Top             =   2940
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox Text13 
         Height          =   315
         Left            =   -70560
         MaxLength       =   50
         TabIndex        =   114
         Top             =   2940
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox Text12 
         Height          =   315
         Left            =   -71685
         MaxLength       =   50
         TabIndex        =   113
         Top             =   2940
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox Text11 
         Height          =   315
         Left            =   -72690
         MaxLength       =   50
         TabIndex        =   112
         Top             =   2940
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox Text10 
         Height          =   315
         Left            =   -73560
         MaxLength       =   50
         TabIndex        =   111
         Top             =   2940
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox Text9 
         Height          =   315
         Left            =   -66330
         MaxLength       =   8
         TabIndex        =   110
         Top             =   2280
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox Text23 
         Height          =   315
         Left            =   -65400
         MaxLength       =   2
         TabIndex        =   109
         Top             =   2280
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.ComboBox cboTransporte 
         Height          =   315
         Left            =   -70320
         TabIndex        =   108
         Top             =   2280
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.TextBox txtValorFrete 
         Height          =   315
         Left            =   -64440
         MaxLength       =   50
         TabIndex        =   107
         Top             =   2280
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         Height          =   525
         Left            =   -65280
         MultiLine       =   -1  'True
         TabIndex        =   105
         Top             =   1440
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.TextBox Text22 
         Height          =   315
         Left            =   -67605
         MaxLength       =   50
         TabIndex        =   99
         Top             =   1680
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox Text21 
         Height          =   315
         Left            =   -66405
         MaxLength       =   50
         TabIndex        =   98
         Top             =   1680
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.TextBox Text19 
         Height          =   315
         Left            =   -69870
         MaxLength       =   50
         TabIndex        =   97
         Top             =   1680
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.TextBox Text16 
         Height          =   315
         Left            =   -68730
         MaxLength       =   50
         TabIndex        =   96
         Top             =   1680
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.TextBox Text15 
         Height          =   315
         Left            =   -72540
         MaxLength       =   50
         TabIndex        =   95
         Top             =   1680
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.Frame Frame3 
         Caption         =   "Filtros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74880
         TabIndex        =   148
         Top             =   8160
         Width           =   13155
         Begin VB.Frame Frame5 
            Caption         =   "Filtrar por:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   180
            TabIndex        =   162
            Top             =   240
            Width           =   4035
            Begin VB.ComboBox cboFiltroNota 
               Height          =   315
               Left            =   960
               TabIndex        =   163
               Top             =   300
               Width           =   2715
            End
            Begin VB.Label Label50 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Escolher:"
               Height          =   195
               Left            =   180
               TabIndex        =   164
               Top             =   360
               Width           =   660
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
            Height          =   735
            Left            =   4320
            TabIndex        =   149
            Top             =   240
            Width           =   5535
            Begin VB.ComboBox cboConNotaMes 
               Height          =   315
               Left            =   2580
               TabIndex        =   168
               Top             =   300
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.ComboBox cboConNotaAno 
               Height          =   315
               Left            =   600
               TabIndex        =   167
               Top             =   300
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.ComboBox cboConNotaCliente 
               Height          =   315
               Left            =   840
               TabIndex        =   166
               Top             =   300
               Visible         =   0   'False
               Width           =   3885
            End
            Begin VB.TextBox txtConNotaCodCliente 
               Height          =   315
               Left            =   4740
               TabIndex        =   151
               Top             =   300
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.TextBox txtConNotaNumNota 
               Height          =   315
               Left            =   1200
               TabIndex        =   150
               Top             =   300
               Visible         =   0   'False
               Width           =   1875
            End
            Begin MSMask.MaskEdBox mskConNotaFinal 
               Height          =   315
               Left            =   2760
               TabIndex        =   152
               Top             =   300
               Visible         =   0   'False
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "dd/mm/yy"
               PromptChar      =   "_"
            End
            Begin ChamaleonBtn.chameleonButton cmdConNotaCal2 
               Height          =   315
               Left            =   3780
               TabIndex        =   153
               Top             =   300
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
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
               BCOL            =   13160660
               BCOLO           =   13160660
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "Frm_NFCe.frx":095F
               PICN            =   "Frm_NFCe.frx":097B
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSMask.MaskEdBox mskConNotaInicial 
               Height          =   315
               Left            =   720
               TabIndex        =   154
               Top             =   300
               Visible         =   0   'False
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "dd/mm/yy"
               PromptChar      =   "_"
            End
            Begin ChamaleonBtn.chameleonButton cmdConNotaCal1 
               Height          =   315
               Left            =   1740
               TabIndex        =   155
               Top             =   300
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
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
               BCOL            =   13160660
               BCOLO           =   13160660
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "Frm_NFCe.frx":2D5D
               PICN            =   "Frm_NFCe.frx":2D79
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label lblConNotaAno 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ano:"
               Height          =   195
               Left            =   180
               TabIndex        =   161
               Top             =   360
               Visible         =   0   'False
               Width           =   330
            End
            Begin VB.Label lblConNotaMes 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mês:"
               Height          =   195
               Left            =   2160
               TabIndex        =   160
               Top             =   360
               Visible         =   0   'False
               Width           =   345
            End
            Begin VB.Label lblConNotaCliente 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cliente:"
               Height          =   195
               Left            =   180
               TabIndex        =   159
               Top             =   360
               Visible         =   0   'False
               Width           =   525
            End
            Begin VB.Label lblConNotaInicial 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Inicial:"
               Height          =   195
               Left            =   180
               TabIndex        =   158
               Top             =   360
               Visible         =   0   'False
               Width           =   450
            End
            Begin VB.Label lblConNotaFinal 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Final:"
               Height          =   195
               Left            =   2280
               TabIndex        =   157
               Top             =   360
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.Label lblConNotaNumNota 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Num. Nota:"
               Height          =   195
               Left            =   180
               TabIndex        =   156
               Top             =   360
               Visible         =   0   'False
               Width           =   810
            End
         End
         Begin ChamaleonBtn.chameleonButton cmdExibirConNotas 
            Height          =   495
            Left            =   9960
            TabIndex        =   165
            Top             =   360
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
            MICON           =   "Frm_NFCe.frx":515B
            PICN            =   "Frm_NFCe.frx":5177
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
      Begin VB.TextBox Text17 
         Height          =   315
         Left            =   -69540
         MaxLength       =   50
         TabIndex        =   115
         Top             =   2940
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Frame frmFiltrosPedidos 
         Caption         =   "Filtros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   -74880
         TabIndex        =   71
         Top             =   8100
         Width           =   13155
         Begin VB.Frame Frame2 
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
            Height          =   735
            Left            =   4320
            TabIndex        =   75
            Top             =   240
            Width           =   5535
            Begin VB.TextBox txtConCodPedido 
               Height          =   315
               Left            =   1020
               TabIndex        =   92
               Top             =   240
               Visible         =   0   'False
               Width           =   1875
            End
            Begin VB.TextBox txtCodClientePedidos 
               Height          =   315
               Left            =   4680
               TabIndex        =   88
               Top             =   240
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.ComboBox cboClientePedidos 
               Height          =   315
               Left            =   720
               TabIndex        =   82
               Top             =   240
               Visible         =   0   'False
               Width           =   3885
            End
            Begin VB.ComboBox cboMesPedidos 
               Height          =   315
               Left            =   540
               TabIndex        =   77
               Top             =   240
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.ComboBox cboAnoPedidos 
               Height          =   315
               Left            =   2340
               Sorted          =   -1  'True
               TabIndex        =   76
               Top             =   240
               Visible         =   0   'False
               Width           =   1155
            End
            Begin MSMask.MaskEdBox mskFinalPedidos 
               Height          =   315
               Left            =   2640
               TabIndex        =   80
               Top             =   260
               Visible         =   0   'False
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "dd/mm/yy"
               PromptChar      =   "_"
            End
            Begin ChamaleonBtn.chameleonButton cmdCalPedidos2 
               Height          =   315
               Left            =   3660
               TabIndex        =   81
               Top             =   260
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
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
               BCOL            =   13160660
               BCOLO           =   13160660
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "Frm_NFCe.frx":5A51
               PICN            =   "Frm_NFCe.frx":5A6D
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSMask.MaskEdBox mskInicialPedidos 
               Height          =   315
               Left            =   720
               TabIndex        =   84
               Top             =   260
               Visible         =   0   'False
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "dd/mm/yy"
               PromptChar      =   "_"
            End
            Begin ChamaleonBtn.chameleonButton cmdCalPedidos1 
               Height          =   315
               Left            =   1740
               TabIndex        =   87
               Top             =   260
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
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
               BCOL            =   13160660
               BCOLO           =   13160660
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "Frm_NFCe.frx":7E4F
               PICN            =   "Frm_NFCe.frx":7E6B
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label lblConsCodPedido 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cód. Pedido:"
               Height          =   195
               Left            =   60
               TabIndex        =   93
               Top             =   240
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.Label lblFinalPedidos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Final:"
               Height          =   195
               Left            =   2220
               TabIndex        =   86
               Top             =   260
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.Label lblInicialPedidos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Inicial:"
               Height          =   195
               Left            =   180
               TabIndex        =   85
               Top             =   260
               Visible         =   0   'False
               Width           =   450
            End
            Begin VB.Label lblClientePedidos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cliente:"
               Height          =   195
               Left            =   120
               TabIndex        =   83
               Top             =   240
               Visible         =   0   'False
               Width           =   525
            End
            Begin VB.Label lblMesPedidos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mês:"
               Height          =   195
               Left            =   120
               TabIndex        =   79
               Top             =   240
               Visible         =   0   'False
               Width           =   345
            End
            Begin VB.Label lblAnoPedidos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ano:"
               Height          =   195
               Left            =   1980
               TabIndex        =   78
               Top             =   240
               Visible         =   0   'False
               Width           =   330
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Filtrar por:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   180
            TabIndex        =   72
            Top             =   240
            Width           =   4035
            Begin VB.ComboBox cboIndicePedidos 
               Height          =   315
               Left            =   960
               TabIndex        =   73
               Top             =   300
               Width           =   2715
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Escolher:"
               Height          =   195
               Left            =   180
               TabIndex        =   74
               Top             =   360
               Width           =   660
            End
         End
         Begin ChamaleonBtn.chameleonButton Command2 
            Height          =   495
            Left            =   9960
            TabIndex        =   94
            Top             =   360
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
            MICON           =   "Frm_NFCe.frx":A24D
            PICN            =   "Frm_NFCe.frx":A269
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
      Begin VB.CommandButton Command1 
         Caption         =   "Mostrar Todos"
         Height          =   255
         Left            =   -71580
         TabIndex        =   63
         Top             =   7920
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.Frame frmItens 
         Caption         =   "Itens da Nota"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3435
         Left            =   120
         TabIndex        =   47
         Top             =   5820
         Width           =   10935
         Begin VB.TextBox txtEdit 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   5040
            TabIndex        =   89
            Top             =   1800
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.TextBox txtNCM 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            Left            =   1860
            MaxLength       =   10
            TabIndex        =   70
            ToolTipText     =   "NCM"
            Top             =   840
            Width           =   1965
         End
         Begin VB.TextBox txtDesc 
            Height          =   315
            Left            =   9000
            MaxLength       =   10
            TabIndex        =   19
            Top             =   480
            Width           =   690
         End
         Begin VB.TextBox txtCST 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   28
            ToolTipText     =   "CST"
            Top             =   840
            Width           =   510
         End
         Begin VB.TextBox txtUnid 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            Left            =   120
            MaxLength       =   10
            TabIndex        =   26
            ToolTipText     =   "Unid de Venda"
            Top             =   840
            Width           =   510
         End
         Begin VB.TextBox txtValor 
            Height          =   315
            Left            =   7260
            MaxLength       =   8
            TabIndex        =   17
            Top             =   480
            Width           =   1020
         End
         Begin VB.TextBox txtSubTotal 
            Height          =   315
            Left            =   9720
            MaxLength       =   8
            TabIndex        =   20
            Top             =   480
            Width           =   1080
         End
         Begin VB.TextBox txtQuant 
            Height          =   315
            Left            =   8340
            MaxLength       =   10
            TabIndex        =   18
            Top             =   480
            Width           =   630
         End
         Begin VB.TextBox txtICMS 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            Left            =   3840
            MaxLength       =   10
            TabIndex        =   29
            ToolTipText     =   "ICMS"
            Top             =   840
            Width           =   570
         End
         Begin VB.TextBox txtCFOP 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            MaxLength       =   10
            TabIndex        =   27
            ToolTipText     =   "CFOP"
            Top             =   840
            Width           =   630
         End
         Begin VB.CheckBox chkCodBarra 
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   195
         End
         Begin VB.CheckBox chkDesc 
            Height          =   195
            Left            =   1740
            TabIndex        =   50
            Top             =   240
            Width           =   195
         End
         Begin VB.TextBox txtCodProduto 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5040
            TabIndex        =   48
            Top             =   240
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtCodBarra 
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   480
            Width           =   1575
         End
         Begin VB.ComboBox cboDescricao 
            Height          =   315
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   480
            Width           =   5475
         End
         Begin MSFlexGridLib.MSFlexGrid GridNotasItens 
            Height          =   1695
            Left            =   120
            TabIndex        =   22
            Top             =   1260
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   2990
            _Version        =   393216
            Appearance      =   0
         End
         Begin ChamaleonBtn.chameleonButton cmdAdicionarItem 
            Height          =   315
            Left            =   8040
            TabIndex        =   21
            Top             =   900
            Width           =   1335
            _ExtentX        =   2355
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
            MICON           =   "Frm_NFCe.frx":AB43
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdRemoverItem 
            Height          =   315
            Left            =   9420
            TabIndex        =   23
            Top             =   900
            Width           =   1335
            _ExtentX        =   2355
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
            MICON           =   "Frm_NFCe.frx":AB5F
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblTotalDesc 
            Alignment       =   1  'Right Justify
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
            Height          =   195
            Left            =   9360
            TabIndex        =   144
            Top             =   3000
            Width           =   225
         End
         Begin VB.Label lblSubTotal 
            Alignment       =   1  'Right Justify
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
            Height          =   195
            Left            =   7320
            TabIndex        =   143
            Top             =   3000
            Width           =   225
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desc."
            Height          =   195
            Left            =   9000
            TabIndex        =   69
            Top             =   240
            Width           =   420
         End
         Begin VB.Label MostraStatus 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   120
            TabIndex        =   64
            Top             =   3000
            Width           =   6705
         End
         Begin VB.Label lblValorNota 
            Alignment       =   1  'Right Justify
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
            Height          =   195
            Left            =   10545
            TabIndex        =   62
            Top             =   3000
            Width           =   225
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde"
            Height          =   195
            Left            =   8340
            TabIndex        =   61
            Top             =   240
            Width           =   345
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor"
            Height          =   195
            Left            =   7260
            TabIndex        =   60
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SubTotal"
            Height          =   195
            Left            =   9720
            TabIndex        =   59
            Top             =   240
            Width           =   645
         End
         Begin VB.Label lblCodFabrica 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. de Barra"
            Height          =   195
            Left            =   360
            TabIndex        =   49
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição"
            Height          =   195
            Left            =   1980
            TabIndex        =   25
            Top             =   240
            Width           =   720
         End
      End
      Begin VB.Frame frmTransporte 
         Caption         =   "Assinatura / Transmissão"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   45
         Top             =   4620
         Width           =   10935
         Begin VB.TextBox Text32 
            BackColor       =   &H00C0C0FF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   8160
            TabIndex        =   138
            TabStop         =   0   'False
            Top             =   540
            Width           =   2595
         End
         Begin VB.TextBox Text31 
            BackColor       =   &H00C0C0FF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   5700
            TabIndex        =   137
            TabStop         =   0   'False
            Top             =   540
            Width           =   2415
         End
         Begin VB.TextBox Text30 
            BackColor       =   &H00C0C0FF&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   180
            TabIndex        =   136
            TabStop         =   0   'False
            Top             =   540
            Width           =   5475
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Recibo"
            Height          =   195
            Left            =   8160
            TabIndex        =   141
            Top             =   300
            Width           =   510
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Protocolo"
            Height          =   195
            Left            =   5700
            TabIndex        =   140
            Top             =   300
            Width           =   675
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Chave de Acesso"
            Height          =   195
            Left            =   180
            TabIndex        =   139
            Top             =   300
            Width           =   1260
         End
      End
      Begin VB.Frame frmNota 
         Caption         =   "Nota Fiscal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   120
         TabIndex        =   32
         Top             =   420
         Width           =   10935
         Begin ChamaleonBtn.chameleonButton cmdCal2 
            Height          =   315
            Left            =   9720
            TabIndex        =   172
            Tag             =   "Calendario"
            Top             =   540
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
            MICON           =   "Frm_NFCe.frx":AB7B
            PICN            =   "Frm_NFCe.frx":AB97
            PICH            =   "Frm_NFCe.frx":CEEA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdCal1 
            Height          =   315
            Left            =   8400
            TabIndex        =   171
            Tag             =   "Calendario"
            Top             =   540
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
            MICON           =   "Frm_NFCe.frx":F23D
            PICN            =   "Frm_NFCe.frx":F259
            PICH            =   "Frm_NFCe.frx":115AC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.ComboBox cboConsumidorFinal 
            Height          =   315
            Left            =   6120
            TabIndex        =   14
            Top             =   3660
            Width           =   1875
         End
         Begin VB.ComboBox cboTipoContribuinte 
            Height          =   315
            Left            =   3300
            TabIndex        =   13
            Top             =   3660
            Width           =   2775
         End
         Begin MSMask.MaskEdBox mskHora 
            Height          =   315
            Left            =   10080
            TabIndex        =   6
            Top             =   540
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskSaida 
            Height          =   315
            Left            =   8760
            TabIndex        =   5
            Top             =   540
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskEmissao 
            Height          =   315
            Left            =   7440
            TabIndex        =   4
            Top             =   540
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtCliIBGE 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Height          =   315
            Left            =   6930
            TabIndex        =   142
            ToolTipText     =   "Cód. IBGE"
            Top             =   1500
            Width           =   1125
         End
         Begin VB.TextBox txtCliIE 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Height          =   315
            Left            =   9720
            TabIndex        =   135
            ToolTipText     =   "IE"
            Top             =   1500
            Width           =   1125
         End
         Begin VB.TextBox txtCliCPF 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Height          =   315
            Left            =   8100
            TabIndex        =   134
            ToolTipText     =   "CNPJ/CPF"
            Top             =   1500
            Width           =   1605
         End
         Begin VB.TextBox txtCliUF 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Height          =   315
            Left            =   6600
            TabIndex        =   133
            ToolTipText     =   "Estado"
            Top             =   1500
            Width           =   315
         End
         Begin VB.TextBox txtCliCidade 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Height          =   315
            Left            =   5220
            TabIndex        =   132
            ToolTipText     =   "Cidade"
            Top             =   1500
            Width           =   1365
         End
         Begin VB.TextBox txtCliBairro 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Height          =   315
            Left            =   3960
            TabIndex        =   131
            ToolTipText     =   "Bairro"
            Top             =   1500
            Width           =   1245
         End
         Begin VB.ComboBox cboModFrete 
            Height          =   315
            Left            =   120
            TabIndex        =   12
            Top             =   3660
            Width           =   3135
         End
         Begin VB.TextBox txtCliNum 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Height          =   315
            Left            =   3480
            TabIndex        =   129
            ToolTipText     =   "Número"
            Top             =   1500
            Width           =   435
         End
         Begin VB.TextBox txtCliEndereco 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Height          =   315
            Left            =   120
            TabIndex        =   128
            ToolTipText     =   "Endereço"
            Top             =   1500
            Width           =   3345
         End
         Begin VB.ComboBox cboDestOperacao 
            Height          =   315
            Left            =   900
            TabIndex        =   2
            Top             =   540
            Width           =   2475
         End
         Begin VB.TextBox txtNumNota 
            Height          =   315
            Left            =   120
            MaxLength       =   50
            TabIndex        =   1
            Top             =   540
            Width           =   720
         End
         Begin VB.TextBox TxtCodCliente 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3840
            MaxLength       =   10
            TabIndex        =   33
            Top             =   900
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.TextBox txtInfComple 
            Height          =   525
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   8
            Text            =   "Frm_NFCe.frx":138FF
            Top             =   2160
            Width           =   10680
         End
         Begin VB.ComboBox cboIndicadorPagamento 
            Height          =   315
            Left            =   120
            TabIndex        =   9
            Top             =   3000
            Width           =   3135
         End
         Begin VB.ComboBox cboFormatoDANFe 
            Height          =   315
            ItemData        =   "Frm_NFCe.frx":13963
            Left            =   3300
            List            =   "Frm_NFCe.frx":13965
            TabIndex        =   10
            Top             =   3000
            Width           =   2415
         End
         Begin VB.ComboBox cboTipoEmissao 
            Height          =   315
            Left            =   5760
            TabIndex        =   11
            Top             =   3000
            Width           =   5055
         End
         Begin VB.ComboBox CboCliente 
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   1140
            Width           =   5175
         End
         Begin VB.ComboBox cboNatureza 
            Height          =   315
            Left            =   3420
            TabIndex        =   3
            Top             =   540
            Width           =   3975
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Consumidor Final"
            Height          =   195
            Left            =   6120
            TabIndex        =   147
            Top             =   3420
            Width           =   1200
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Contribuinte"
            Height          =   195
            Left            =   3300
            TabIndex        =   146
            Top             =   3420
            Width           =   1425
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Modalidade do Frete"
            Height          =   195
            Left            =   120
            TabIndex        =   130
            Top             =   3420
            Width           =   1455
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hora"
            Height          =   195
            Left            =   9885
            TabIndex        =   44
            Top             =   285
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. Saida"
            Height          =   195
            Left            =   8700
            TabIndex        =   43
            Top             =   285
            Width           =   660
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. Emissão"
            Height          =   195
            Left            =   7500
            TabIndex        =   42
            Top             =   285
            Width           =   840
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Natureza"
            Height          =   195
            Left            =   3420
            TabIndex        =   41
            Top             =   300
            Width           =   765
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Destino"
            Height          =   195
            Index           =   0
            Left            =   945
            TabIndex        =   40
            Top             =   300
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NF Num."
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   39
            Top             =   300
            Width           =   630
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome do cliente"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   900
            Width           =   1155
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Informações Complementares"
            Height          =   195
            Left            =   60
            TabIndex        =   37
            Top             =   1920
            Width           =   2085
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Indicador Forma Pagto"
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   2760
            Width           =   1605
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Formato Impressão DANFE"
            Height          =   195
            Left            =   3300
            TabIndex        =   35
            Top             =   2760
            Width           =   1920
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Emissão NFe"
            Height          =   195
            Left            =   5760
            TabIndex        =   34
            Top             =   2760
            Width           =   1515
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelar 
         Height          =   615
         Left            =   11100
         TabIndex        =   31
         Top             =   1860
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
         MICON           =   "Frm_NFCe.frx":13967
         PICN            =   "Frm_NFCe.frx":13983
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
         Left            =   11100
         TabIndex        =   0
         Top             =   540
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
         MICON           =   "Frm_NFCe.frx":15715
         PICN            =   "Frm_NFCe.frx":15731
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdSair 
         Height          =   615
         Left            =   11100
         TabIndex        =   52
         Top             =   8640
         Width           =   2175
         _ExtentX        =   3836
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
         MICON           =   "Frm_NFCe.frx":174C3
         PICN            =   "Frm_NFCe.frx":174DF
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
         Left            =   11100
         TabIndex        =   53
         Top             =   3180
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
         MICON           =   "Frm_NFCe.frx":19271
         PICN            =   "Frm_NFCe.frx":1928D
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
         Height          =   615
         Left            =   11100
         TabIndex        =   54
         Top             =   5340
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&DANFe"
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
         MICON           =   "Frm_NFCe.frx":1B01F
         PICN            =   "Frm_NFCe.frx":1B03B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdTransmitir 
         Height          =   615
         Left            =   11100
         TabIndex        =   55
         Top             =   4680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Transmitir"
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
         MICON           =   "Frm_NFCe.frx":1CDCD
         PICN            =   "Frm_NFCe.frx":1CDE9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdCancelarNota 
         Height          =   615
         Left            =   11100
         TabIndex        =   56
         Top             =   6000
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Cancelar NFe"
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
         MICON           =   "Frm_NFCe.frx":1EB7B
         PICN            =   "Frm_NFCe.frx":1EB97
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
         Left            =   11100
         TabIndex        =   24
         Top             =   1200
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
         MICON           =   "Frm_NFCe.frx":1F1D8
         PICN            =   "Frm_NFCe.frx":1F1F4
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
         Left            =   11100
         TabIndex        =   57
         Top             =   2520
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
         MICON           =   "Frm_NFCe.frx":20F86
         PICN            =   "Frm_NFCe.frx":20FA2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSFlexGridLib.MSFlexGrid GridPedidos 
         Height          =   7215
         Left            =   -74880
         TabIndex        =   65
         Top             =   480
         Width           =   13155
         _ExtentX        =   23204
         _ExtentY        =   12726
         _Version        =   393216
         TextStyleFixed  =   1
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin ChamaleonBtn.chameleonButton cmdRecalcularNF 
         Height          =   615
         Left            =   11100
         TabIndex        =   145
         Top             =   7560
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Recalcular"
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
         MICON           =   "Frm_NFCe.frx":22D34
         PICN            =   "Frm_NFCe.frx":22D50
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSFlexGridLib.MSFlexGrid GridNotas 
         Height          =   7395
         Left            =   -74880
         TabIndex        =   46
         Top             =   480
         Width           =   13155
         _ExtentX        =   23204
         _ExtentY        =   13044
         _Version        =   393216
         TextStyleFixed  =   1
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin ChamaleonBtn.chameleonButton cmdConsultar 
         Height          =   615
         Left            =   11100
         TabIndex        =   170
         Top             =   6660
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Consultar"
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
         MICON           =   "Frm_NFCe.frx":24AE2
         PICN            =   "Frm_NFCe.frx":24AFE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblTotalNota 
         Alignment       =   1  'Right Justify
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
         Height          =   195
         Left            =   -61980
         TabIndex        =   169
         Top             =   7920
         Width           =   225
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Placa"
         Height          =   195
         Left            =   -66300
         TabIndex        =   127
         Top             =   2040
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Numero"
         Height          =   195
         Left            =   -70575
         TabIndex        =   126
         Top             =   2700
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca"
         Height          =   195
         Left            =   -71685
         TabIndex        =   125
         Top             =   2700
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde Vol."
         Height          =   195
         Left            =   -73560
         TabIndex        =   123
         Top             =   2700
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peso bruto"
         Height          =   195
         Left            =   -69570
         TabIndex        =   122
         Top             =   2700
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peso liquido"
         Height          =   195
         Left            =   -68625
         TabIndex        =   121
         Top             =   2700
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UF"
         Height          =   195
         Left            =   -65445
         TabIndex        =   120
         Top             =   2040
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transportadora"
         Height          =   195
         Left            =   -70320
         TabIndex        =   119
         Top             =   2040
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fatura"
         Height          =   195
         Left            =   -65280
         TabIndex        =   106
         Top             =   1200
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B. ICMS"
         Height          =   195
         Left            =   -67605
         TabIndex        =   104
         Top             =   1440
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B. ICMS Sub."
         Height          =   195
         Left            =   -66405
         TabIndex        =   103
         Top             =   1440
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seguro"
         Height          =   195
         Left            =   -69870
         TabIndex        =   102
         Top             =   1440
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Despesas"
         Height          =   195
         Left            =   -68730
         TabIndex        =   101
         Top             =   1440
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mensagem"
         Height          =   195
         Left            =   -72540
         TabIndex        =   100
         Top             =   1440
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Especie"
         Height          =   195
         Left            =   -72690
         TabIndex        =   124
         Top             =   2700
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor do Frete"
         Height          =   195
         Left            =   -64440
         TabIndex        =   118
         Top             =   2040
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblQuantPedidos 
         Alignment       =   1  'Right Justify
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
         Height          =   195
         Left            =   -61980
         TabIndex        =   66
         Top             =   7740
         Width           =   225
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   58
      Top             =   10245
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17251
            Text            =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
            TextSave        =   "Desenv.: Online.Info - Informática  - Tel.: (89) 3544-2553"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2470
            MinWidth        =   2470
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "12:08"
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
Attribute VB_Name = "Frm_NFCe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private moCombo As cComboHelper
Public TbNotas As New ADODB.Recordset
Public TbConsulta As New ADODB.Recordset
Public TbNotaPedido As New ADODB.Recordset
Dim Tb As New ADODB.Recordset
Dim Titulo, Book As Variant, NomeTabela
Dim TbProduto As New ADODB.Recordset
Private iRow As Long, iCol As Long
Private Sub AtualizarGrid_Itens()
  Dim i As Integer
   Dim sSQL As String
   
   For i = 1 To GridNotasItens.Rows - 1
      If GridNotasItens.TextMatrix(i, 1) <> "" Then
         dbData.Execute "UPDATE TbNFCe_Itens SET CFOP = " & GridNotasItens.TextMatrix(i, 5) & ", ICMSCST = '" & GridNotasItens.TextMatrix(i, 4) & "', CodNcm = '" & GridNotasItens.TextMatrix(i, 3) & "'  WHERE IdNFProd = " & txtCodNota.Text & " AND IdNFProd_Item = " & GridNotasItens.TextMatrix(i, 11) & ""
         dbData.Execute "UPDATE Produtos SET NCM = '" & GridNotasItens.TextMatrix(i, 3) & "'  WHERE CODIGO = " & GridNotasItens.TextMatrix(i, 1) & ""
      End If
   Next
End Sub

Private Sub GravarPedido()
flag = False

'On Error GoTo Err_Grava

Dim sSQL As String
Dim r As ADODB.Recordset
Dim totalRegistros As Long

'If txtCodPedido = "" Then Exit Sub

'preencher objetos da nota com o pedido
sSQL = "SELECT pedidos.*, cliente.codigo, cliente.nome as VarNome FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente WHERE pedidos.cod_pedido = " & txtCodPedido & ";"
Set r = dbData.OpenRecordset(sSQL, totalRegistros)

If Not r.BOF Then Mostrar_Pedido r

If r.State <> 0 Then r.Close
Set r = Nothing


'If TxtCodCliente.Text = "" Then MsgBox "O campo CLIENTE é obrigatório.", vbCritical, "Online Commerce": TxtCodCliente.SetFocus: Exit Sub
If cboModFrete.Text = "" Then MsgBox "o campo Modalidade do frete é obrigatório.", vbCritical, "Online Commerce": cboModFrete.SetFocus: Exit Sub
If cboDestOperacao.Text = "" Then MsgBox "O campo código CFOP é obrigatório.", vbCritical, "Online Commerce": cboDestOperacao.SetFocus: Exit Sub
'If Text15.Text = "" Then MsgBox "O campo mensagem é obrigatório.", vbCritical, "Online Commerce": Text15.SetFocus: Exit Sub

If txtCodPedido.Text = "0" Then

Else
    RsOpen TbNotas, "SELECT * FROM TbNFCe"
    TbNotas.AddNew
End If

flag = True

Load_Data
TbNotas.Update
vgDb.CommitTrans

Load_Controls

TransformarPedidoemNFE  'TESTE

SomarProdutosNota

PreencherGridNotas

cmdNovo.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = True
cmdExcluir.Enabled = False
'Command5.Enabled = True
cmdTransmitir.Enabled = True

'Clear_Controls
LimparCamposItens

End Sub



Sub Load_Data_Itens()
Dim sSQL As String, seq As Integer

If txtCodNota.Text = "" Then Exit Sub

    sSQL = "SELECT MAX(IdNFProd_Item) r FROM TbNFCe_Itens WHERE IdNFProd = " & Val(txtCodNota.Text)
    seq = SQLExecutaRetorno(sSQL, "r", 0) + 1
    Tb("IdNFProd") = Format(txtCodNota.Text, "@")
    Tb("IdNFProd_Item") = seq
    Tb("IDProduto") = Format(txtCodProduto, "@")
    Tb("DescricaoProduto") = UCase(Format(cboDescricao, "@"))
    Tb("CFOP") = Format(txtCFOP, "@")
    Tb("CodNcm") = Format(txtNCM, "@")
    Tb("ICMSCST") = Right(Format(txtCST, "@"), 3)
    Tb("UN") = UCase(Format(txtUnid, "@"))
    Tb("ValorUnit") = CDbl(Format(txtValor, "@"))
    'Tb("ValorTotalBruto") = CDbl(Format(txtSubTotal, "@"))
    'Tb("tipodesconto") = Format(1, "@")
    Tb("Desconto") = CDbl(Format(txtDesc, "@"))
    'Tb("Valordesconto") = CDbl(Format(txtDesc, "@"))
    If txtQuant.Text <> "" Then Tb("QtdeMov") = CDbl(Format(txtQuant, "@")) Else Tb("QuantidadeComercial") = Format(1, "@")
    If txtICMS.Text <> "" Then Tb("Aliq_Icms") = CDbl(Format(txtICMS, "@"))
    If txtICMS.Text <> "" Then Tb("Bc_Icms") = CDbl(Format(txtSubTotal, "@"))
    If txtICMS.Text <> "" Then Tb("Vlr_Icms") = Round(Format(txtSubTotal, "@") * (Format(txtICMS, "@") / 100), 2)
End Sub

Private Sub Calcular_Total()
Dim var_Quant As Double
Dim var_VALOR As Currency, var_Total As Currency

If txtQuant.Text = "" Then var_Quant = 1 Else var_Quant = txtQuant.Text
If txtValor.Text = "" Then var_VALOR = 0 Else var_VALOR = txtValor.Text

var_Total = var_VALOR * var_Quant
txtSubTotal.Text = Format(var_Total, ocMONEY)
End Sub
Sub Clear_Controls()
   Limpa_Tudo Me
End Sub

Private Sub LimparCamposItens()
txtCodBarra.Text = ""
cboDescricao.Text = ""
txtCodProduto.Text = ""
txtCST = Format("0", "@")
txtNCM = Format("0", "@")
txtValor = Format("0", "@")
txtSubTotal = Format("0", "@")
txtQuant = Format("1", "@")
txtICMS = Format("0", "@")
txtCFOP = Format("0", "@")
txtDesc = Format("0", "@")
chkDesc.Value = 0
chkDesc.Value = 1
End Sub

Sub Load_Data()
'On Error GoTo erro
    vgDb.BeginTrans
    If Text17.Text = "" Then Text17.Text = "0"
    If Text18.Text = "" Then Text18.Text = "0"
    TbNotas("NFeIdentificadorDestino") = IIf(IsNull(Format(Left(cboDestOperacao.Text, 1), "@")) Or Vazio(Format(Left(cboDestOperacao.Text, 1), "@")), 1, Format(Left(cboDestOperacao.Text, 1), "@"))
    TbNotas("IdNFProd") = Format(txtCodNota.Text, "@")
    'TbNotas("SerieNF") = 1
    'TbNotas("InscricaoEstadual") = 0
    'TbNotas("Suframa") = 0
    'TbNotas("CNPJ_CPF") = 0
    'TbNotas("Logradouro") = 0
    'TbNotas("numero") = 0
    'TbNotas("CodigoIBGE") = 0
    'TbNotas("Bairro") = 0
    'TbNotas("Complemento") = 0
    'TbNotas("Municipio") = 0
    'TbNotas("UF") = 0
    TbNotas("Linha1") = Format(Left(txtInfComple, 60), "@")
    TbNotas("NaturezaOperacao") = Format(cboNatureza, "@")
    TbNotas("TipoNF") = "S"
    TbNotas("DataEmissao") = IIf(TbNotas("DataEmissao") = Empty, Format(Date, "dd/mm/yyyy"), Format(mskEmissao, "@"))
    TbNotas("DataSaidaEntrada") = IIf(TbNotas("DataSaidaEntrada") = Empty, Format(Date, "dd/mm/yyyy"), Format(mskSaida, "@"))
    TbNotas("HoraSaida") = IIf(TbNotas("HoraSaida") = Empty, Format(Time(), "HH:MM:ss"), Format(mskHora, "@"))
    
    'TbNotas("TranspCodigo") = IIf(IsNull(Format(txtCodTransporte, "@")) Or Vazio(Format(txtCodTransporte, "@")), 0, Format(txtCodTransporte, "@"))
    TbNotas("NomeTrasnportador") = Format(cboTransporte, "@")
    TbNotas("Placa_Veiculo") = Format(txtPlaca, "@")
    TbNotas("Qtde_Trasnp") = "1" 'Format(Text10, "@")
    TbNotas("Especie_Transp") = Format(Text11, "@")
    TbNotas("Marca_Trasnp") = Format(Text12, "@")
    TbNotas("Num_Transp") = 0 'Format(txtVolNumeracao, "@")
    'TbNotas("CodigoObservacao") = IIf(IsNull(Format(Text15, "@")) Or Vazio(Format(Text15, "@")), 0, Format(Text15, "@"))
    TbNotas("Valor_Frete") = IIf(Vazio(txtValorFrete), 0, CDbl(Format(txtValorFrete, "##0.00")))
    TbNotas("Valor_Seguro") = IIf(IsNull(Format(Text19, "@")) Or Vazio(Format(Text19, "@")), 0, CDbl(Format(Text19, "##0.00")))
    TbNotas("OutrasDespesasAces") = IIf(IsNull(Format(Text16, "@")) Or Vazio(Format(Text16, "@")), 0, CDbl(Format(Text16, "##0.00")))
    TbNotas("NumeNota") = Format(txtNumNota, "@")
    TbNotas("NFeCodigoNota") = IIf(TbNotas("NFeCodigoNota") = 0, GeraCodigoNota, TbNotas("NFeCodigoNota"))
    TbNotas("PesoBruto_Transp") = IIf(IsNull(Format(Text17, "@")) Or Vazio(Format(Text17, "@")), 0, CDbl(Format(Text17, "##0.000")))
    TbNotas("PesoLiq_Transp") = IIf(IsNull(Format(Text18, "@")) Or Vazio(Format(Text18, "@")), 0, CDbl(Format(Text18, "##0.000")))
    TbNotas("BaseCalc_ICMS") = IIf(IsNull(Format(Text22, "@")) Or Vazio(Format(Text22, "@")), 0, CDbl(Format(Text22, "##0.000")))
    TbNotas("BaseCalc_ICSM_Subst") = IIf(IsNull(Format(Text21, "@")) Or Vazio(Format(Text21, "@")), 0, CDbl(Format(Text21, "##0.00")))
    TbNotas("UF_Mot_Transp") = Format(Text23, "@")
    TbNotas("IDCliente") = IIf(IsNull(Format(txtCodCliente, "@")) Or Vazio(Format(txtCodCliente, "@")), 0, Format(txtCodCliente, "@"))
    TbNotas("NomeRazSocial") = Format(cboCliente, "@")
    TbNotas("Frete_Por_Conta") = IIf(IsNull(Format(Left(cboModFrete.Text, 1), "@")) Or Vazio(Format(Left(cboModFrete.Text, 1), "@")), 9, Format(Left(cboModFrete.Text, 1), "@"))
    TbNotas("NFeIndicadorFormaPagto") = Format(cboIndicadorPagamento.Text, "@")
    TbNotas("NFeFormatoImpressaoDANFe") = Format(cboFormatoDANFe.Text, "@")
    TbNotas("NFeTipoEmissao") = Format(cboTipoEmissao.Text, "@")
    TbNotas("Num_OS_VD_Origem") = Format(txtCodPedido.Text, "@")
    TbNotas("NFeIdentificadorDestino") = IIf(IsNull(Format(Left(cboDestOperacao.Text, 1), "@")) Or Vazio(Format(Left(cboDestOperacao.Text, 1), "@")), 1, Format(Left(cboDestOperacao.Text, 1), "@"))
    TbNotas("NFeIndicadorIEDestinatario") = IIf(IsNull(Format(Left(cboTipoContribuinte.Text, 1), "@")) Or Vazio(Format(Left(cboTipoContribuinte.Text, 1), "@")), 1, Format(Left(cboTipoContribuinte.Text, 1), "@"))
    TbNotas("NFeConsumidorFinal") = IIf(IsNull(Format(Left(cboConsumidorFinal.Text, 1), "@")) Or Vazio(Format(Left(cboConsumidorFinal.Text, 1), "@")), 1, Format(Left(cboConsumidorFinal.Text, 1), "@"))
   
    'cboConsumidorFinal
    Exit Sub

Resume

'erro:
'    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce"
'    vgDb.RollbackTrans
'    Exit Sub
End Sub

Public Sub Load_Controls()
'On Error GoTo erro
    txtCodCliente = Format(TbNotas("IDCliente"), "@")
    cboCliente = Format(TbNotas("NomeRazSocial"), "@")
    
    If TbNotas("NFeIdentificadorDestino") = 1 Then
        cboDestOperacao.Text = "1 - Operação Interna"
    ElseIf TbNotas("CodigoNatureza") = 2 Then
        cboDestOperacao.Text = "2 - Operação Interestadual"
    ElseIf TbNotas("CodigoNatureza") = 3 Then
        cboDestOperacao.Text = "3 - Operação com Exterior"
    End If

    If TbNotas("NFeIndicadorIEDestinatario") = 1 Then
        cboTipoContribuinte.Text = "1 - CONTRIBUINTE ICMS"
    ElseIf TbNotas("NFeIndicadorIEDestinatario") = 2 Then
        cboTipoContribuinte.Text = "2 - CONTRIBUINTE ISENTO"
    ElseIf TbNotas("NFeIndicadorIEDestinatario") = 9 Then
        cboTipoContribuinte.Text = "9 - NÃO CONTRIBUINTE"
    End If

    If TbNotas("NFeConsumidorFinal") = False Then
        cboConsumidorFinal.Text = "0 - NÃO"
    ElseIf TbNotas("NFeConsumidorFinal") = True Then
        cboConsumidorFinal.Text = "1 - SIM"
    End If

'cboConsumidorFinal
    'cboDestOperacao = Format(TbNotas("CodigoNatureza"), "@")
    
    Text1 = Format(TbNotas("Linha1"), "@")
    cboNatureza = Format(TbNotas("NaturezaOperacao"), "@")
    mskEmissao = Format(TbNotas("DataEmissao"), "@")
    mskSaida = Format(TbNotas("DataSaidaEntrada"), "@")
    mskHora = Format(TbNotas("HoraSaida"), "@")
    txtCodTransporte = "" 'Format(TbNotas("TranspCodigo"), "@")
    cboTransporte = Format(TbNotas("NomeTrasnportador"), "@")
    txtPlaca = Format(TbNotas("Placa_Veiculo"), "@")
    Text10 = Format(TbNotas("Qtde_Trasnp"), "@")
    Text11 = Format(TbNotas("Especie_Transp"), "@")
    Text12 = Format(TbNotas("Marca_Trasnp"), "@")
    txtVolNumeracao = Format(TbNotas("Num_Transp"), "@")
    Text15 = Format(TbNotas("Linha2"), "@")
    txtNumNota = Format(TbNotas("NumeNota"), "@")
    Text19 = Format(TbNotas("Valor_Seguro"), "##,##0.00")
    Text16 = Format(TbNotas("OutrasDespesasAces"), "##,##0.00")
    txtValorFrete = Format(TbNotas("Valor_Frete"), "##,##0.00")
    Text22 = Format(TbNotas("BaseCalc_ICMS"), "##,##0.00")
    Text21 = Format(TbNotas("BaseCalc_ICSM_Subst"), "##,##0.00")
    Text17 = Format(TbNotas("PesoBruto_Transp"), "@")
    Text18 = Format(TbNotas("PesoLiq_Transp"), "@")
    Text23 = Format(TbNotas("UF_Mot_Transp"), "@")

    If Left(TbNotas("Frete_Por_Conta"), 1) = 0 Then
        cboModFrete.Text = "0 - POR CONTA DO EMITENTE"
    ElseIf Left(TbNotas("Frete_Por_Conta"), 1) = 1 Then
        cboModFrete.Text = "1 - POR CONTA DE AMBOS"
    ElseIf Left(TbNotas("Frete_Por_Conta"), 1) = 2 Then
        cboModFrete.Text = "2 - POR CONTA DO TERCEIROS"
    ElseIf Left(TbNotas("Frete_Por_Conta"), 1) = 9 Then
        cboModFrete.Text = "9 - SEM FRETE"
    End If
    
    txtCodNota = Format(TbNotas("IdNFProd"), "@")
    Text30 = Format(TbNotas("NFCeChaveAcesso"), "@")
    Text31 = Format(TbNotas("NFCeProtocolo"), "@")
    Text32 = Format(TbNotas("NFCeRecibo"), "@")
    cboIndicadorPagamento.Text = Format(TbNotas("NFeIndicadorFormaPagto"), "@")
    cboFormatoDANFe.Text = Format(TbNotas("NFeFormatoImpressaoDANFe"), "@")
    cboTipoEmissao.Text = Format(TbNotas("NFeTipoEmissao"), "@")
    txtCodPedido = Format(TbNotas("Num_OS_VD_Origem"), "@")
    txtInfComple.Text = Format(TbNotas("Linha1"), "@")

    'habilita/desabilita botões
    'cmdCancelarNota.Enabled = False
    cmdImprimir.Enabled = False
    
    If TbNotas("NFCeEnviada") = False Then
        cmdSalvar.Enabled = False
        cmdCancelar.Enabled = False
        cmdAlterar.Enabled = True
        cmdExcluir.Enabled = True
        cmdTransmitir.Enabled = True
        'cmdCancelarNota.Enabled = False
        cmdImprimir.Enabled = False
    Else
       cmdSalvar.Enabled = False
       cmdCancelar.Enabled = False
       cmdAlterar.Enabled = False
       cmdExcluir.Enabled = False
       cmdTransmitir.Enabled = False
       'cmdCancelarNota.Enabled = True
       cmdImprimir.Enabled = True
    'End If
    'ElseIf TbNotas("Enviada") = False And TbNotas("Cancelada") = False Then
    '    cmdTransmitir.Enabled = True
    End If
    
    If TbNotas("NFCeCancelada") Then
       'cmdCancelarNota.Enabled = False
    End If
    
    MostraStatus = MostraStatus_F9()
Exit Sub

Resume

'erro:
'MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub

Private Sub Mostrar_ItensNota()
Dim sSQL As String, enviada As Boolean
Dim totalRegistros As Long
    
    On Error GoTo ErrLoad
    
    sSQL = "SELECT * FROM TbNFCe_Itens WHERE IdNFProd = " & Val(txtCodNota.Text)
    RsOpen Tb, sSQL
    
    If Tb.RecordCount > 0 Then totalRegistros = Tb.RecordCount
    
    
'    enviada = SQLExecutaRetorno("SELECT Enviada FROM NotaFiscal WHERE CodigoNota = " & Val(Frm_NFCe.txtCodNota.Text), "Enviada", 0)
'
'    If enviada Then
'       cboDestOperacao.Enabled = False
'       Text3.Enabled = False
'       mskEmissao.Enabled = False
'       mskSaida.Enabled = False
'       mskHora.Enabled = False
'       Text7.Enabled = False
'       Text8.Enabled = False
'       txtPlaca.Enabled = False
'       Text10.Enabled = False
'    End If
    
    LimparGridItensNota
    DoEvents
    FormatarGridItensNota Tb
    Exit Sub
    
ErrLoad:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub FormatarGridItensNota(rTabela As ADODB.Recordset)
   Dim i As Integer, j As Integer
   
   With GridNotasItens
      .Visible = False
      .Redraw = False
      
      .Clear
      .Cols = 12
      .Rows = 2
      
      .ColWidth(0) = 200
      .ColWidth(1) = 0
      .ColWidth(2) = 3500
      .ColWidth(3) = 1200
      .ColWidth(4) = 500
      .ColWidth(5) = 800
      .ColWidth(6) = 500
      .ColWidth(7) = 1000
      .ColWidth(8) = 800
      .ColWidth(9) = 800
      .ColWidth(10) = 1000
      .ColWidth(11) = 0

      
      'CodigoProduto, NomeProduto, CST, Unidade, Qtde, Valor, SubTotal
      .TextMatrix(0, 1) = "CÓD."
      .TextMatrix(0, 2) = "DESCRIÇÃO"
      .TextMatrix(0, 3) = "NCM"
      .TextMatrix(0, 4) = "CST"
      .TextMatrix(0, 5) = "CFOP"
      .TextMatrix(0, 6) = "UND"
      .TextMatrix(0, 7) = "VALOR"
      .TextMatrix(0, 8) = "QTDE"
      .TextMatrix(0, 9) = "DESC."
      .TextMatrix(0, 10) = "TOTAL"
      .TextMatrix(0, 11) = "ITEM"
      
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
      
      i = 1
      
      'ALINHAMENTO
      .ColAlignment(0) = 1
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 1
      .ColAlignment(6) = 1
      .ColAlignment(7) = 2
      .ColAlignment(8) = 2
      .ColAlignment(9) = 2
      .ColAlignment(10) = 2
      .ColAlignment(11) = 1
      
      'CodigoProduto, NomeProduto, CST, Unidade, Qtde, Valor, SubTotal
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = rTabela("IDProduto")
            .TextMatrix(.Rows - 1, 2) = rTabela("DescricaoProduto")
            .TextMatrix(.Rows - 1, 3) = rTabela("CodNcm")
            .TextMatrix(.Rows - 1, 4) = rTabela("ICMSCST")
            .TextMatrix(.Rows - 1, 5) = rTabela("CFOP")
            .TextMatrix(.Rows - 1, 6) = rTabela("UN")
            .TextMatrix(.Rows - 1, 7) = Format(rTabela("ValorUnit"), ocMONEY)
            .TextMatrix(.Rows - 1, 8) = Format(rTabela("QtdeMov"), ocMONEY)
            .TextMatrix(.Rows - 1, 9) = Format(rTabela("Desconto"), ocMONEY)
            .TextMatrix(.Rows - 1, 10) = Format(rTabela("QtdeMov") * rTabela("ValorUnit"), ocMONEY)
            .TextMatrix(.Rows - 1, 11) = rTabela("IdNFProd_Item")
            
            rTabela.MoveNext
            .Rows = .Rows + 1
            i = i + 1
         Loop
      End If
      
      .Rows = .Rows - 1
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 2
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 3
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
              
     'GridNotasItens.ColWidth(0) = 400
      'GridNotasItens.Rows = 11
      GridNotasItens.Col = 0
            
      .Visible = True
      .Redraw = True
      
      SomarGridItens
   End With
End Sub

Private Sub PreencherGridNotas()
Dim totalRegistros As Long

On Error GoTo ErrLoad

RsOpen TbConsulta, "SELECT *, (Valor_NF_Prod - DescontoPromocional) Valor_NF, " & _
                "(CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 0 THEN 'Enviada' ELSE (CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) AS Status " & _
                "FROM TbNFCe order by NumeNota desc"
                
If TbConsulta.RecordCount > 0 Then totalRegistros = TbConsulta.RecordCount

LimparGridNotas
FormatarGridNotas TbConsulta

Exit Sub
Resume

ErrLoad:
    MsgBox Err.Description, vbCritical
    Err.Clear
    Set TbConsulta = Nothing
End Sub

Private Sub SomarProdutosNota()
Dim sSQL As String, vTotal As Double

On Error GoTo erro

    sSQL = "SELECT ISNULL(SUM(ValorTotalBruto), 0) r FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
    vTotal = SQLExecutaRetorno(sSQL, "r", 0)
    
    sSQL = "UPDATE NotaFiscal SET ValorProdutos = " & FSQL(vTotal, 2) & ", ValorNota = " & FSQL(vTotal, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text)
    SQLExecuta sSQL
    
    Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "SistemasNFe": Exit Sub
End Sub

Private Sub TransformarPedidoemNFE()
Dim sSQL As String
'Dim r As ADODB.Recordset
'Dim totalRegistros As Long

'If txtCodPedido = "" Then Exit Sub

'preencher objetos da nota com o pedido
'sSQL = "SELECT pedidos.*, cliente.codigo, cliente.nome as VarNome FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente WHERE pedidos.cod_pedido = " & txtCodPedido & ";"
'Set r = dbData.OpenRecordset(sSQL, totalRegistros)

'If Not r.BOF Then Mostrar_Pedido r

'If r.State <> 0 Then r.Close
'Set r = Nothing

Dim tblItensPedido As ADODB.Recordset

'Atualiza a base de dados (funcionando)
Dim VarCodNota As Integer
VarCodNota = CInt(txtCodNota.Text)

sSQL = "INSERT INTO TbNFCe_Itens ( " & _
        "IDProduto, " & _
        "DescricaoProduto, " & _
        "CFOP, " & _
        "CodNcm, " & _
        "ICMSCST, " & _
        "UN, " & _
        "ValorUnit, " & _
        "Desconto, " & _
        "QtdeMov, " & _
        "Aliq_Icms, " & _
        "Bc_Icms, " & _
        "Vlr_Icms,  " & _
        "IdNFProd_Item, " & _
        "IdNFProd " & _
        " ) " & _
        "SELECT pedidos_itens.cod_produto, produtos.descricao, produtos.cfop, produtos.ncm, produtos.icmscst, produtos.unid_medida, pedidos_itens.preco, 0, pedidos_itens.quantidade, 0, (pedidos_itens.preco * pedidos_itens.quantidade) as varVBC, 0, pedidos_itens.item, " & VarCodNota & " " & _
        "FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto INNER JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
        "WHERE pedidos_itens.COD_PEDIDO = " & txtCodPedido.Text & ";"
'MsgBox sSQL
dbData.Execute sSQL

'preencher o grid dos itens com o pedido
sSQL = "SELECT * FROM TbNFCe WHERE CodigoNota = " & Val(txtCodNota.Text)
RsOpen Tb, sSQL

FormatarGridItensNota Tb

'MOSTRAR A QUANTIDADE REGISTROS
'lblQuantPedidos.Caption = Format(totalRegistros, "00")
End Sub

Private Sub cboAno_Change()

End Sub

Private Sub cboAno_GotFocus()
SelectControl mskFim
End Sub


Private Sub cboCategoria_Change()

End Sub

Private Sub cboCategoria_GotFocus()
SelectControl mskFim
End Sub


Private Sub cboAnoPedidos_GotFocus()
Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
Dim i As Integer

cboAnoPedidos.Clear

iAno = Year(Date)
FirstYear = iAno - 2
LastYear = iAno + 2

For i = FirstYear To LastYear
   cboAnoPedidos.AddItem i
Next

moCombo.AttachTo cboAnoPedidos
End Sub


Private Sub cboClientePedidos_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
'Dim itemAtual As String
'Dim codAtual As String

'itemAtual = CboCliente.Text
'codAtual = TxtCodCliente.Text
cboClientePedidos.Clear

sSQL = "SELECT DISTINCT nome, codigo FROM cliente ORDER BY nome;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboClientePedidos.AddItem r("nome")
   cboClientePedidos.ItemData(cboClientePedidos.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

'CboCliente.Text = itemAtual
'TxtCodCliente.Text = codAtual
moCombo.AttachTo cboClientePedidos
End Sub


Private Sub cboClientePedidos_LostFocus()
On Error GoTo TrataErro

If cboClientePedidos.Text = "" Then txtCodClientePedidos.Text = "": Exit Sub
If cboClientePedidos.ListIndex = -1 Then txtCodClientePedidos.Text = "": Exit Sub

txtCodClientePedidos = cboClientePedidos.ItemData(cboClientePedidos.ListIndex)

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub


Private Sub cboConNotaAno_GotFocus()
Dim iAno As Integer, FirstYear As Integer, LastYear As Integer
Dim i As Integer

cboConNotaAno.Clear

iAno = Year(Date)
FirstYear = iAno - 2
LastYear = iAno + 2

For i = FirstYear To LastYear
   cboConNotaAno.AddItem i
Next

moCombo.AttachTo cboConNotaAno
End Sub


Private Sub cboConNotaCliente_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
'Dim itemAtual As String
'Dim codAtual As String

'itemAtual = CboCliente.Text
'codAtual = TxtCodCliente.Text
cboConNotaCliente.Clear

sSQL = "SELECT DISTINCT nome, codigo FROM cliente ORDER BY nome;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboConNotaCliente.AddItem r("nome")
   cboConNotaCliente.ItemData(cboConNotaCliente.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

'CboCliente.Text = itemAtual
'TxtCodCliente.Text = codAtual
moCombo.AttachTo cboConNotaCliente
End Sub


Private Sub cboConNotaCliente_LostFocus()
On Error GoTo TrataErro

If cboConNotaCliente.Text = "" Then txtConNotaCodCliente.Text = "": Exit Sub
If cboConNotaCliente.ListIndex = -1 Then txtConNotaCodCliente.Text = "": Exit Sub

txtConNotaCodCliente = cboConNotaCliente.ItemData(cboConNotaCliente.ListIndex)

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub


Private Sub cboConNotaMes_GotFocus()
cboConNotaMes.Clear

cboConNotaMes.AddItem "Janeiro"
cboConNotaMes.AddItem "Fevereiro"
cboConNotaMes.AddItem "Março"
cboConNotaMes.AddItem "Abril"
cboConNotaMes.AddItem "Maio"
cboConNotaMes.AddItem "Junho"
cboConNotaMes.AddItem "Julho"
cboConNotaMes.AddItem "Agosto"
cboConNotaMes.AddItem "Setembro"
cboConNotaMes.AddItem "Outubro"
cboConNotaMes.AddItem "Novembro"
cboConNotaMes.AddItem "Dezembro"

moCombo.AttachTo cboConNotaMes
End Sub


Private Sub cboConsumidorFinal_GotFocus()
Dim VarText As String
VarText = cboConsumidorFinal.Text

cboConsumidorFinal.Clear
cboConsumidorFinal.AddItem "0 - NÃO"
cboConsumidorFinal.AddItem "1 - SIM"

If cboConsumidorFinal.Text = "" Then cboConsumidorFinal.Text = VarText

moCombo.AttachTo cboConsumidorFinal
End Sub


Private Sub cboDestOperacao_GotFocus()
Dim VarText As String
VarText = cboDestOperacao.Text

cboDestOperacao.Clear
cboDestOperacao.AddItem "1 - Operação Interna"
cboDestOperacao.AddItem "2 - Operação Interestadual"
cboDestOperacao.AddItem "3 - Operação com Exterior"

If cboDestOperacao.Text = "" Then cboDestOperacao.Text = VarText

moCombo.AttachTo cboDestOperacao
End Sub


Private Sub cboFiltroNota_Click()
lblConNotaNumNota.Visible = False
txtConNotaNumNota.Visible = False
lblConNotaCliente.Visible = False
cboConNotaCliente.Visible = False
txtConNotaCodCliente.Visible = False
lblConNotaAno.Visible = False
lblConNotaMes.Visible = False
cboConNotaAno.Visible = False
cboConNotaMes.Visible = False
lblConNotaInicial.Visible = False
lblConNotaFinal.Visible = False
mskConNotaInicial.Visible = False
mskConNotaFinal.Visible = False
cmdConNotaCal1.Visible = False
cmdConNotaCal2.Visible = False

If cboFiltroNota.Text = "TODAS" Then
    Exit Sub
ElseIf cboFiltroNota.Text = "NUM. NOTA" Then
    lblConNotaNumNota.Visible = True
    txtConNotaNumNota.Visible = True
ElseIf cboFiltroNota.Text = "CLIENTE" Then
    lblConNotaCliente.Visible = True
    cboConNotaCliente.Visible = True
    'txtConNotaCodCliente.Visible = True
ElseIf cboFiltroNota.Text = "DATAS" Then
    lblConNotaInicial.Visible = True
    lblConNotaFinal.Visible = True
    mskConNotaInicial.Visible = True
    mskConNotaFinal.Visible = True
    cmdConNotaCal1.Visible = True
    cmdConNotaCal2.Visible = True
ElseIf cboFiltroNota.Text = "MENSAL" Then
    lblConNotaAno.Visible = True
    lblConNotaMes.Visible = True
    cboConNotaAno.Visible = True
    cboConNotaMes.Visible = True
Else
End If
End Sub


Private Sub cboFiltroNota_GotFocus()
cboFiltroNota.Clear
cboFiltroNota.AddItem "TODAS"
cboFiltroNota.AddItem "NUM. NOTA"
cboFiltroNota.AddItem "CLIENTE"
cboFiltroNota.AddItem "DATAS"
cboFiltroNota.AddItem "MENSAL"

moCombo.AttachTo cboFiltroNota
End Sub


Private Sub cboIndicePedidos_Click()
lblInicialPedidos.Visible = False
mskInicialPedidos.Visible = False
cmdCalPedidos1.Visible = False
lblFinalPedidos.Visible = False
mskFinalPedidos.Visible = False
cmdCalPedidos2.Visible = False
lblClientePedidos.Visible = False
cboClientePedidos.Visible = False
lblConsCodPedido.Visible = False
txtConCodPedido.Visible = False
lblMesPedidos.Visible = False
cboMesPedidos.Visible = False
lblAnoPedidos.Visible = False
cboAnoPedidos.Visible = False

If cboIndicePedidos.Text = "PEDIDO" Then
    lblConsCodPedido.Visible = True
    txtConCodPedido.Visible = True
ElseIf cboIndicePedidos.Text = "CLIENTE" Then
    lblClientePedidos.Visible = True
    cboClientePedidos.Visible = True
ElseIf cboIndicePedidos.Text = "DATAS" Then
    lblInicialPedidos.Visible = True
    mskInicialPedidos.Visible = True
    cmdCalPedidos1.Visible = True
    lblFinalPedidos.Visible = True
    mskFinalPedidos.Visible = True
    cmdCalPedidos2.Visible = True
ElseIf cboIndicePedidos.Text = "MENSAL" Then
    lblMesPedidos.Visible = True
    cboMesPedidos.Visible = True
    lblAnoPedidos.Visible = True
    cboAnoPedidos.Visible = True
Else
End If
End Sub

Private Sub cboIndicePedidos_GotFocus()
cboIndicePedidos.Clear
cboIndicePedidos.AddItem "PEDIDO"
cboIndicePedidos.AddItem "CLIENTE"
cboIndicePedidos.AddItem "DATAS"
cboIndicePedidos.AddItem "MENSAL"

moCombo.AttachTo cboIndicePedidos
End Sub


Private Sub cboMes_Change()

End Sub

Private Sub cboMes_GotFocus()
SelectControl mskFim
End Sub


Private Sub cboIndicePedidos_Validate(Cancel As Boolean)
cboIndicePedidos_Click
End Sub

Private Sub cboMesPedidos_GotFocus()
cboMesPedidos.Clear

cboMesPedidos.AddItem "Janeiro"
cboMesPedidos.AddItem "Fevereiro"
cboMesPedidos.AddItem "Março"
cboMesPedidos.AddItem "Abril"
cboMesPedidos.AddItem "Maio"
cboMesPedidos.AddItem "Junho"
cboMesPedidos.AddItem "Julho"
cboMesPedidos.AddItem "Agosto"
cboMesPedidos.AddItem "Setembro"
cboMesPedidos.AddItem "Outubro"
cboMesPedidos.AddItem "Novembro"
cboMesPedidos.AddItem "Dezembro"

moCombo.AttachTo cboMesPedidos
End Sub


Private Sub cboModFrete_GotFocus()
Dim VarText As String
VarText = cboModFrete.Text

cboModFrete.Clear
cboModFrete.AddItem "0 - POR CONTA DO EMITENTE"
cboModFrete.AddItem "1 - POR CONTA DE AMBOS"
cboModFrete.AddItem "2 - POR CONTA DO TERCEIROS"
cboModFrete.AddItem "9 - SEM FRETE"

If cboModFrete.Text = "" Then cboModFrete.Text = VarText

'cboModFrete.AddItem ""
moCombo.AttachTo cboModFrete
End Sub


Private Sub cboTipoContribuinte_GotFocus()
Dim VarText As String
VarText = cboTipoContribuinte.Text

cboTipoContribuinte.Clear
cboTipoContribuinte.AddItem "1 - CONTRIBUINTE ICMS"
cboTipoContribuinte.AddItem "2 - CONTRIBUINTE ISENTO"
cboTipoContribuinte.AddItem "9 - NÃO CONTRIBUINTE"

If cboTipoContribuinte.Text = "" Then cboTipoContribuinte.Text = VarText

'cboModFrete.AddItem ""
moCombo.AttachTo cboTipoContribuinte
End Sub
Private Sub cmdAdicionarItem_Click()
Dim sSQL As String, vTotal As Double

'On Error GoTo erro
If txtCodNota.Text = "" Then Exit Sub
If txtCodProduto.Text = "" Then Exit Sub
If txtSubTotal.Text = "" Then Exit Sub
If Len(txtNCM.Text) < 8 Then ShowMsg "NCM INCORRETO!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbExclamation

    sSQL = "SELECT * FROM TbNFCe_Itens WHERE IdNFProd = " & Val(txtCodNota.Text)
    RsOpen Tb, sSQL


    vgDb.BeginTrans
    
    'insere os dados itens
    Tb.AddNew
    Load_Data_Itens
    Tb.Update
    
    vgDb.CommitTrans
    
    Call cmdRecalcularNF_Click
    
    'Limpa_Tudo Me ' limpa tudo
    
    'sSQL = "SELECT ISNULL(SUM(ValorTotalBruto), 0) r FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
    'vTotal = SQLExecutaRetorno(sSQL, "r", 0)
    
    'sSQL = "UPDATE NotaFiscal SET ValorProdutos = " & FSQL(vTotal, 2) & ", ValorNota = " & FSQL(vTotal, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text)
    'SQLExecuta sSQL
    
    'EXIBIR NO GRID
    sSQL = "SELECT * FROM TbNFCe_Itens WHERE IdNFProd = " & Val(txtCodNota.Text)
    RsOpen Tb, sSQL
    
    FormatarGridItensNota Tb
    
    LimparCamposItens
    
    KeyCode = 0
    If chkCodBarra.Value = 1 Then txtCodBarra.SetFocus
'End If
Exit Sub
'erro:
'MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "SistemasNFe": Exit Sub
End Sub
Private Sub SomarGridItens()
Dim Total As Currency, SUBTOTAL As Currency, Desc As Currency
Dim i As Integer

SUBTOTAL = 0
Desc = 0
Total = 0

'Sub-Total
With GridNotasItens
   For i = 1 To .Rows - 1
      .Col = 0
      .Row = i
      
         .Col = 6
         SUBTOTAL = SUBTOTAL + (.TextMatrix(.Row, 7) * .TextMatrix(.Row, 8))
        '.Col = 10
         Desc = Desc + .TextMatrix(.Row, 9)
         '.Col = 11
         Total = Total + .TextMatrix(.Row, 10)
   Next
End With


lblSubtotal.Caption = Format(SUBTOTAL, ocMONEY)
lblTotalDesc.Caption = Format(Desc, ocMONEY)
lblValorNota.Caption = Format(Total, ocMONEY)
End Sub

Private Sub cmdAlterar_Click()
Dim sSQL As String
'Dim enviada As Boolean
'Dim totalRegistros As Long
flag = False

'On Error GoTo Err_Grava

If txtCodCliente.Text = "" Then MsgBox "O campo código do cliente é obrigatório.", vbCritical, "Online Commerce": txtCodCliente.SetFocus: Exit Sub
If cboModFrete.Text = "" Then MsgBox "o campo Modalidade do frete é obrigatório.", vbCritical, "Online Commerce": cboModFrete.SetFocus: Exit Sub
If cboDestOperacao.Text = "" Then MsgBox "O campo código CFOP é obrigatório.", vbCritical, "Online Commerce": cboDestOperacao.SetFocus: Exit Sub
If txtCodNota.Text = "" Then MsgBox "Essa Nota Fiscal não existe!", vbCritical, "Online Commerce": cboDestOperacao.SetFocus: Exit Sub
    
If TbNotas.EditMode = 2 Then
   resp = MsgBox("Confirma inclusão ?", 36, Titulo)
   flag = True
   If resp <> 6 Then Exit Sub
Else
   resp = MsgBox("Confirma alteração ?", 36, Titulo)
   flag = False
   If resp <> 6 Then Exit Sub
   
   'TbNotas.Edit
End If

Load_Data
TbNotas.Update

vgDb.CommitTrans

Load_Controls

SomarProdutosNota

PreencherGridNotas

cmdNovo.Enabled = True
cmdSalvar.Visible = True
cmdCancelar.Visible = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdImprimir.Enabled = False
cmdTransmitir.Enabled = False
'cmdCancelarNota.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False

Clear_Controls
LimparCamposItens

Exit Sub
Resume
'Err_Grava:
'    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce"
End Sub

Private Function Atualizar_Dados() As Boolean
   'A atualização deve ser feita utilizando o comando UPDATE do sql
   'e não mais usando o método .Update do Recordset
   
   'Não se deve comparar se o campo está vazio ou não, pois dessa forma não
   'haverá atualização quando for necessário apagar alguma informação
   
   Dim sSQL As String
   
   'Comando de atualização
   sSQL = "UPDATE setor SET setor = '" & cboSetor.Text & "' WHERE (cod_setor = " & txtCodigo.Text & ");"
   
   'Retorna o resultado da atualização
   Atualizar_Dados = dbData.Execute(sSQL)
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

mskEmissao = Format(varData, "dd/mm/yyyy")   'Exibe a data no campo
End Sub

Private Sub cmdCal2_Click()
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

mskSaida = Format(varData, "dd/mm/yyyy")   'Exibe a data no campo
End Sub


Private Sub cmdCalPedidos1_Click()
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

mskInicialPedidos = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdCalPedidos2_Click()
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

mskFinalPedidos = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub


Private Sub cmdCancelarNota_Click()
Dim Justificativa As String
 If MsgBox("Tem certeza que deseja Cancelar a Nota Fiscal?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
 Justificativa = InputBox("Informe a Justificativa do Cancelamento da Nota:", "Cancelamento da Nota", "DESISTENCIA DA COMPRA")
 vsNumeroNota = Val(txtCodNota.Text)
 iRetorno = CancelaNFCe(TbNotas("NFCeChaveAcesso"), TbNotas("NFCeProtocolo"), Justificativa, True)
 If iRetorno Then
    SQL = "UPDATE TbNFCe SET " & _
          "NFCeCancelada = 1, " & _
          "NFCeCanceladaProtocolo = " & NFeNumeroProtocolo & ", " & _
          "NFCeCanceladaJustificativa = '" & Justificativa & "' " & _
          "WHERE IdNFProd = " & Val(txtCodNota.Text)
    vgDb.Execute SQL
    RsOpen TbNotas, "SELECT *, " & _
                    "(CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 0 THEN 'Enviada' ELSE (CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) AS Status " & _
                    "FROM TbNFCe WHERE IdNFProd = " & Val(txtCodNota.Text)
    Load_Controls
    FormatarGridNotas TbNotas
 End If
End Sub

Private Sub CboCliente_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim itemAtual As String
Dim codAtual As String

itemAtual = cboCliente.Text
codAtual = txtCodCliente.Text
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

cboCliente.Text = itemAtual
txtCodCliente.Text = codAtual
moCombo.AttachTo cboCliente
End Sub

Private Sub CboCliente_LostFocus()
On Error GoTo TrataErro

If cboCliente.Text = "" Then txtCodCliente.Text = "": Exit Sub
If cboCliente.ListIndex = -1 Then txtCodCliente.Text = "": Exit Sub

txtCodCliente = cboCliente.ItemData(cboCliente.ListIndex)

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub


Private Sub cboDescricao_GotFocus()
   moCombo.AttachTo cboDescricao
   
Dim sSQL As String
Dim r As ADODB.Recordset

'Limpa a lista
If chkDesc.Value = 1 Then cboDescricao.Clear

sSQL = "SELECT DISTINCT descricao, codigo FROM produtos ORDER BY descricao;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboDescricao.AddItem ValidateNull(r("descricao"))
    cboDescricao.ItemData(cboDescricao.NewIndex) = r("codigo")
   r.MoveNext
Loop

moCombo.AttachTo cboDescricao
End Sub


Private Sub cboDescricao_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub cboDescricao_LostFocus()
If chkDesc.Value = 1 Then
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   txtCodBarra.Text = ""
   
   If cboDescricao.Text = "" Then txtCodProduto.Text = "": Exit Sub
   If cboDescricao.ListIndex = -1 Then txtCodProduto.Text = "": Exit Sub

   txtCodProduto = cboDescricao.ItemData(cboDescricao.ListIndex)
   
   If txtCodProduto.Text = "" Then Exit Sub
   
   sSQL = "SELECT codigo, descricao, cod_barra, unid_medida, ncm, CFOP, ICMSCST, ICMSAliq  FROM produtos WHERE (codigo = " & txtCodProduto.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then
    If txtCodBarra.Text = "" Then txtCodBarra.Text = r("cod_barra")
    txtUnid.Text = ValidateNull(r("unid_medida"))
    txtCFOP.Text = ValidateNull(r("CFOP"))
    txtCST.Text = ValidateNull(r("ICMSCST"))
    txtNCM.Text = ValidateNull(r("ncm"))
    txtICMS.Text = Format(ValidateNull(r("ICMSAliq")), "##,##0.00")
   End If
   
   MostrarValorVenda
   txtQuant.SetFocus

   If r.BOF Then ShowMsg "Produto não cadastrado.", vbExclamation
   If r.State <> 0 Then r.Close
End If
End Sub
Private Sub MostrarValorVenda()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim vrVenda As Currency
If txtCodProduto.Text = "" Then Exit Sub

'mostrar o ultimo preço de compra
sSQL = "SELECT TOP 1 venda FROM produtos_entrada_itens WHERE (codigo_produto = " & txtCodProduto & ") ORDER BY codigo DESC;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then vrVenda = r("venda")
If r.State <> 0 Then r.Close
Set r = Nothing

txtValor.Text = Format(vrVenda, ocMONEY)
End Sub

Private Sub cboNatureza_GotFocus()
Dim VarText As String
VarText = cboNatureza.Text

cboNatureza.Clear
cboNatureza.AddItem "VENDA DE MERCADORIA"
cboNatureza.AddItem ""

If cboNatureza.Text = "" Then cboNatureza.Text = VarText

moCombo.AttachTo cboNatureza
End Sub

Private Sub cboTransporte_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim itemAtual As String
Dim codAtual As String

itemAtual = cboTransporte.Text
codAtual = txtCodTransporte.Text
cboTransporte.Clear

sSQL = "SELECT DISTINCT fantasia, codigo FROM transportadora ORDER BY fantasia;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboTransporte.AddItem r("fantasia")
   cboTransporte.ItemData(cboTransporte.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

cboTransporte.Text = itemAtual
txtCodTransporte.Text = codAtual
moCombo.AttachTo cboTransporte
End Sub


Private Sub cboTransporte_LostFocus()
On Error GoTo TrataErro

If cboTransporte.Text = "" Then txtCodTransporte.Text = "": Exit Sub
If cboTransporte.ListIndex = -1 Then txtCodTransporte.Text = "": Exit Sub

txtCodTransporte = cboTransporte.ItemData(cboTransporte.ListIndex)

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub


Private Sub chkCodBarra_Click()
If chkCodBarra.Value = 1 Then chkDesc.Value = 0
txtCodBarra.Locked = False
cboDescricao.Text = ""
If txtCodBarra.Locked = False Then cboDescricao.Locked = True
txtCodBarra.SetFocus
End Sub

Private Sub chkDesc_Click()
If chkDesc.Value = 1 Then chkCodBarra.Value = 0
cboDescricao.Locked = False
txtCodBarra.Text = ""
If cboDescricao.Locked = False Then txtCodBarra.Locked = True
cboDescricao.SetFocus
End Sub


Private Sub cmdCancelar_Click()
On Error GoTo Err_Cancela

cmdNovo.Enabled = True
cmdSalvar.Enabled = False
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdImprimir.Enabled = False
cmdTransmitir.Enabled = False
'cmdCancelarNota.Enabled = False

If TbNotas.EditMode <> 0 Then TbNotas.CancelUpdate

Limpa_Tudo Me
txtInfComple.Text = "EMPRESA ME OU EPP OPTANTE PELO SIMPLES NACIONAL NÃO GERA DIREITO A CREDITO FISCAL DE ICMS OU ISS."
Exit Sub

Err_Cancela:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub

Private Sub cmdConNotaCal1_Click()
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

mskConNotaInicial = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdConNotaCal2_Click()
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

mskConNotaFinal = Format(varData, "dd/mm/yy")   'Exibe a data no campo
End Sub

Private Sub cmdConsultar_Click()
    If (Text32.Text = Empty) Or (Text32.Text = "") Then Exit Sub
    If (Text31.Text = Empty) Or (Text31.Text = "") Then
       ConsultaReciboNFCe Text32.Text, Text30.Text, "1", True
    Else
       consultaNFCe Text30.Text
       If cStat = 100 Then
          SQL = "UPDATE TbNFCe SET " & _
                "NFceEnviada = 1, " & _
                "NFCeProtocolo = " & NFeNumeroProtocolo & ", " & _
                "NFCeProtocoloDataHora = '" & NFeDataHora & "' " & _
                "WHERE IdNFProd = " & Val(txtCodNota.Text)
          vgDb.Execute SQL
          'SQL = "INSERT INTO NotaFiscalRecibos (CodigoNota, NumeroProtocolo, DataHora) Values " & _
          '      "(" & Val(txtCodNota.Text) & ", " & NFeNumeroProtocolo & ", '" & NFeDataHora & "')"
          'vgDb.Execute SQL
       End If
    End If
    RsOpen TbNotas, "SELECT *, " & _
                    "(CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 0 THEN 'Enviada' ELSE (CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) AS Status " & _
                    "FROM TbNFCe WHERE IdNFProd = " & Val(txtCodNota.Text)
    Load_Controls
    FormatarGridNotas TbNotas
End Sub

Private Sub cmdExibirConNotas_Click()
Dim totalRegistros As Long

On Error GoTo ErrLoad

If cboFiltroNota.Text = "TODAS" Then
    RsOpen TbConsulta, "SELECT *, (Valor_NF_Prod - DescontoPromocional) Valor_NF, " & _
                    "(CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 0 THEN 'Enviada' ELSE (CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) AS Status " & _
                    "FROM TbNFCe order by NumeNota desc"
ElseIf cboFiltroNota.Text = "NUM. NOTA" Then
    RsOpen TbConsulta, "SELECT *, (Valor_NF_Prod - DescontoPromocional) Valor_NF, " & _
                    "(CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 0 THEN 'Enviada' ELSE (CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) AS Status " & _
                    "FROM TbNFCe WHERE NumeNota = " & txtConNotaNumNota & " order by NumeNota desc"
ElseIf cboFiltroNota.Text = "CLIENTE" Then
    RsOpen TbConsulta, "SELECT *, (Valor_NF_Prod - DescontoPromocional) Valor_NF, " & _
                    "(CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 0 THEN 'Enviada' ELSE (CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) AS Status " & _
                    "FROM TbNFCe WHERE IDCliente = " & txtConNotaCodCliente.Text & " order by NumeNota desc"
ElseIf cboFiltroNota.Text = "DATAS" Then
    RsOpen TbConsulta, "SELECT *, (Valor_NF_Prod - DescontoPromocional) Valor_NF, " & _
                    "(CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 0 THEN 'Enviada' ELSE (CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) AS Status " & _
                    "FROM TbNFCe WHERE (DataEmissao >= CONVERT(DATETIME, '" & Format(mskConNotaInicial.Text, ocDATA) & "', 103)) AND (DataEmissao <= CONVERT(DATETIME, '" & Format(mskConNotaFinal.Text, ocDATA) & "', 103)) order by NumeNota desc"
ElseIf cboFiltroNota.Text = "MENSAL" Then
    RsOpen TbConsulta, "SELECT *, (Valor_NF_Prod - DescontoPromocional) Valor_NF, " & _
                    "(CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 0 THEN 'Enviada' ELSE (CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) AS Status " & _
                    "FROM TbNFCe WHERE (MONTH(DataEmissao) = " & cboConNotaMes.ListIndex + 1 & ") AND (YEAR(DataEmissao) = " & cboConNotaAno & ") order by NumeNota desc"
End If

If TbConsulta.RecordCount > 0 Then totalRegistros = TbConsulta.RecordCount
lblTotalNota.Caption = Format(totalRegistros, "00")

LimparGridNotas
FormatarGridNotas TbConsulta

Exit Sub
Resume

ErrLoad:
    MsgBox Err.Description, vbCritical
    Err.Clear
    Set TbConsulta = Nothing
End Sub

Private Sub cmdNovo_Click()
On Error GoTo ErrLoad

'pegando o numero correto da nota
'Dim var_NumeroNota As Integer
Dim ConsultaSQL As String
Dim tbNota As ADODB.Recordset

      ConsultaSQL = "SELECT ISNULL(MAX(NumeNota), 0) AS Maior_nota FROM TbNFCe"
      Set tbNota = dbData.OpenRecordset(ConsultaSQL)
      'If Not tbNota.BOF Then var_NumeroNota = tbNota("ultima_nota") + 1
      



'preecher objetos do form
Dim totalRegistros As Long
RsOpen TbNotas, "SELECT *, " & _
                "(CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 0 THEN 'Enviada' ELSE (CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) AS Status " & _
                "FROM TbNFCe"
                
If TbNotas.RecordCount > 0 Then totalRegistros = TbNotas.RecordCount

    Clear_Controls
    
    txtCodPedido.Text = 0
    
    If TbNotas.EOF And TbNotas.BOF Then
        txtNumNota.Text = tbNota("Maior_nota") + 1
        txtCodNota.Text = "1"
    Else
        TbNotas.MoveLast
        txtNumNota.Text = tbNota("Maior_nota") + 1
        txtCodNota.Text = TbNotas("CodigoNota") + 1
    End If
    
    
    cboIndicadorPagamento.Text = "0 - Pagamento à vista"
    cboFormatoDANFe.Text = "1 - Retrato"
    cboTipoEmissao.Text = "1 - Normal"
    cboModFrete.Text = ""
    txtValorFrete.Text = "0,00"
    Text16.Text = "0,00"
    Text17.Text = "0,00"
    Text18.Text = "0,00"
    Text19.Text = "0,00"
    Text21.Text = "0,00"
    Text22.Text = "0,00"
    txtInfComple.Text = "EMPRESA ME OU EPP OPTANTE PELO SIMPLES NACIONAL NÃO GERA DIREITO A CREDITO FISCAL DE ICMS OU ISS."
    LimparCamposItens
    
    TbNotas.AddNew
    cmdNovo.Enabled = False
    cmdSalvar.Enabled = True
    cmdCancelar.Enabled = True
    cmdImprimir.Enabled = False
    cmdTransmitir.Enabled = False
    'cmdCancelarNota.Enabled = False
    cmdAlterar.Enabled = False
    cmdExcluir.Enabled = False
    cboDestOperacao.SetFocus

Exit Sub
Resume

ErrLoad:
    MsgBox Err.Description, vbCritical
    Err.Clear
    Set TbNotas = Nothing
End Sub

Private Sub cmdRecalcularNF_Click()
Dim sSQL As String
    sSQL = "UPDATE TbNFCe SET " & _
           "Valor_NF_Prod = tb.TotalBruto, " & _
           "DescontoPromocional = tb.Desconto " & _
           "FROM (SELECT ISNULL(SUM((QtdeMov * ValorUnit)), 0) TotalBruto, ISNULL(SUM(Desconto), 0) Desconto FROM TbNFCe_Itens WHERE IdNFProd = " & Val(txtCodNota.Text) & ") tb " & _
           "WHERE IdNFProd = " & Val(txtCodNota.Text)
    SQLExecuta sSQL
End Sub

Private Sub cmdRemoverItem_Click()
Dim sSQL As String, vTotal As Double

'On Error GoTo erro

    'sSQL = "SELECT * FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)
    'RsOpen Tb, sSQL


    'vgDb.BeginTrans
    
    'Tb.AddNew 'insere os dados
    'Load_Data_Itens
    'Tb.Update
    
    'vgDb.CommitTrans
    
    'Limpa_Tudo Me ' limpa tudo
    If ShowMsg("Deseja remover o item: " & GridNotasItens.TextMatrix(GridNotasItens.Row, 2) & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub

    
    dbData.Execute "DELETE FROM TbNFCe_Itens WHERE (IDProduto = " & GridNotasItens.TextMatrix(GridNotasItens.Row, 1) & ") AND (IdNFProd_Item = " & GridNotasItens.TextMatrix(GridNotasItens.Row, 11) & ");"

    Call cmdRecalcularNF_Click
    
   ' sSQL = "SELECT ISNULL(SUM(ValorTotalBruto), 0) r FROM NotaFiscalItens WHERE CodigoNota = " & Val(Frm_NFCe.txtCodNota.Text)
    'vTotal = SQLExecutaRetorno(sSQL, "r", 0)
    
    'sSQL = "UPDATE NotaFiscal SET ValorProdutos = " & FSQL(vTotal, 2) & ", ValorNota = " & FSQL(vTotal, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text)

    'SQLExecuta sSQL
    
    sSQL = "SELECT * FROM TbNFCe_Itens WHERE IDNFProd = " & Val(txtCodNota.Text)
    RsOpen Tb, sSQL
    
    FormatarGridItensNota Tb
    
'    lblValorNota.Caption = Format(Tb("vTotal"), ocMONEY)
    
    KeyCode = 0
    If chkDesc.Value = 1 Then cboDescricao.SetFocus
'End If
Exit Sub
'erro:
'MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "SistemasNFe": Exit Sub
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSalvar_Click()
flag = False

'On Error GoTo Err_Grava

'If TxtCodCliente.Text = "" Then MsgBox "O campo CLIENTE é obrigatório.", vbCritical, "Online Commerce": CboCliente.SetFocus: Exit Sub
If cboModFrete.Text = "" Then MsgBox "o campo Modalidade do frete é obrigatório.", vbCritical, "Online Commerce": cboModFrete.SetFocus: Exit Sub
If cboDestOperacao.Text = "" Then MsgBox "O campo código CFOP é obrigatório.", vbCritical, "Online Commerce": cboDestOperacao.SetFocus: Exit Sub
'If Text15.Text = "" Then MsgBox "O campo mensagem é obrigatório.", vbCritical, "Online Commerce": Text15.SetFocus: Exit Sub

If txtCodPedido.Text = "0" Then

Else
    RsOpen TbNotas, "SELECT * FROM TbNFCe"
    TbNotas.AddNew
End If

If TbNotas.EditMode = 2 Then
   resp = MsgBox("Confirma inclusão ?", 36, Titulo)
   flag = True
   If resp <> 6 Then Exit Sub
Else
   resp = MsgBox("Confirma alteração ?", 36, Titulo)
   flag = False
   If resp <> 6 Then Exit Sub
   'TbNotas.Edit
End If

Load_Data
TbNotas.Update
vgDb.CommitTrans

Load_Controls

SomarProdutosNota

PreencherGridNotas

cmdNovo.Enabled = True
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
'Command5.Enabled = True
cmdTransmitir.Enabled = False

Clear_Controls
LimparCamposItens

Exit Sub
Resume
'Err_Grava:
'    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce"
End Sub
Private Sub cmdExcluir_Click()
'On Error GoTo Err_Delete

If MostraStatus.Caption <> "Transmitida/Autorizada" Or MostraStatus.Caption <> "Cancelada" Or MostraStatus.Caption <> "Denegada" Then
    resp = MsgBox("Confirma a exclusão ?", 36, Titulo)
    If resp <> 6 Then Exit Sub
    
    dbData.Execute "DELETE FROM TbNFCe_Itens WHERE IdNFProd = " & Val(txtCodNota.Text)
    dbData.Execute "DELETE FROM TbNFCe WHERE  IdNFProd = " & Val(txtCodNota.Text)
    PreencherGridNotas
    Limpa_Tudo Me
    txtInfComple.Text = "EMPRESA ME OU EPP OPTANTE PELO SIMPLES NACIONAL NÃO GERA DIREITO A CREDITO FISCAL DE ICMS OU ISS."
Else
    MsgBox "Não é possivel excluir uma nota Cancelada, Trasmitida ou Denegada!", vbCritical, "Online Commerce": Exit Sub
End If

Exit Sub

'Err_Delete:
'    If Err = 3021 Then
'        Clear_Controls
'        Exit Sub
'    End If
'    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub

Private Sub cmdImprimir_Click()
   On Error GoTo DeuErro
     Dim sistNFe As snfe.Util
     Set sistNFe = New snfe.Util
     
     dirXML = SQLExecutaRetorno("SELECT DiretorioXML FROM empresa", "DiretorioXML", App.path)
     xCaminhoXML = dirXML & "\nfe\arquivos\procNFe\" & Format(mskEmissao, "yyyymm") & "\" & Text30 & "-procNFe.xml"
     xCaminhoPDF = dirXML & "\nfe\arquivos\PDF\NFe" & Text30 & ".pdf"
     
     If Not Existe(xCaminhoXML) Then consultaNFCe Text30, True
     
     If Not Existe(xCaminhoXML) Then Exit Sub
     
     iRetorno = ConfiguraDLLNFeNFCe(65, "1", sistNFe)
     Call sistNFe.DANFCeImprimir(xCaminhoXML, False, "", True, xCaminhoPDF, 0, False, False, "")
     
     Exit Sub
DeuErro:
    If InStr(1, Err.Description, "Exception") > 0 Then
       iRetorno = ConfiguraDLLNFeNFCe(65, "1", sistNFe)
       Call sistNFe.DANFCeImprimir(xCaminhoXML, False, "", True, xCaminhoPDF, 0, False, False, "")
    Else
       MsgBox Err.Description, vbInformation
    End If
    Err.Clear
End Sub

Private Sub cmdTransmitir_Click()
Call cmdRecalcularNF_Click
DoEvents
iRetorno = TransmitirNFCe(Val(txtCodNota.Text), "1", True)
If iRetorno Then
   SQL = "UPDATE TbNFCe SET " & _
         "NFCeEnviada = 1 " & _
         "WHERE IDNFProd = " & Val(txtCodNota.Text)
   vgDb.Execute SQL
   RsOpen TbNotas, "SELECT *, " & _
                   "(CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 0 THEN 'Enviada' ELSE (CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) AS Status " & _
                   "FROM TbNFCe WHERE IdNFProd = " & Val(txtCodNota.Text)
   Load_Controls
   FormatarGridNotas TbNotas
End If
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Combo1_GotFocus()

End Sub


Private Sub Combo2_Change()

End Sub

Private Sub Combo3_Change()

End Sub

Private Sub Combo4_Change()

End Sub

Private Sub Command1_Click()
PreencherGridNotas
End Sub

Private Sub Command2_Click()
If cboIndicePedidos.Text = "" Then Exit Sub

   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim totalRegistros As Long
   
If cboIndicePedidos.Text = "PEDIDO" Then
    sSQL = "SELECT cliente.codigo, pedidos.cod_cliente, cliente.nome as var_Nome, pedidos.tipo_pagamento AS var_tipoPGTO, pedidos.cod_pedido AS var_codped, pedidos.data_compra as var_DTCompra, pedidos.total AS var_total, (CASE WHEN NotaFiscal.cod_pedido = pedidos.cod_pedido THEN 'SIM' ELSE 'NÃO' END) AS Status " & _
           "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente LEFT JOIN NotaFiscal ON NotaFiscal.cod_pedido = pedidos.cod_pedido WHERE pedidos.cod_pedido = " & txtConCodPedido & ";"
ElseIf cboIndicePedidos.Text = "CLIENTE" Then
    sSQL = "SELECT cliente.codigo, pedidos.cod_cliente, cliente.nome as var_Nome, pedidos.tipo_pagamento AS var_tipoPGTO, pedidos.cod_pedido AS var_codped, pedidos.data_compra as var_DTCompra, pedidos.total AS var_total, (CASE WHEN NotaFiscal.cod_pedido = pedidos.cod_pedido THEN 'SIM' ELSE 'NÃO' END) AS Status " & _
           "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente LEFT JOIN NotaFiscal ON NotaFiscal.cod_pedido = pedidos.cod_pedido WHERE (cliente.codigo = " & txtCodClientePedidos.Text & ") ORDER BY pedidos.cod_pedido;"
ElseIf cboIndicePedidos.Text = "DATAS" Then
    sSQL = "SELECT cliente.codigo, pedidos.cod_cliente, cliente.nome as var_Nome, pedidos.tipo_pagamento AS var_tipoPGTO, pedidos.cod_pedido AS var_codped, pedidos.data_compra as var_DTCompra, pedidos.total AS var_total, (CASE WHEN NotaFiscal.cod_pedido = pedidos.cod_pedido THEN 'SIM' ELSE 'NÃO' END) AS Status " & _
           "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente LEFT JOIN NotaFiscal ON NotaFiscal.cod_pedido = pedidos.cod_pedido WHERE (pedidos.data_compra >= CONVERT(DATETIME, '" & Format(mskInicialPedidos.Text, ocDATA) & "', 103)) AND (pedidos.data_compra <= CONVERT(DATETIME, '" & Format(mskFinalPedidos.Text, ocDATA) & "', 103)) ORDER BY pedidos.cod_pedido;"
ElseIf cboIndicePedidos.Text = "MENSAL" Then
    sSQL = "SELECT cliente.codigo, pedidos.cod_cliente, cliente.nome as var_Nome, pedidos.tipo_pagamento AS var_tipoPGTO, pedidos.cod_pedido AS var_codped, pedidos.data_compra as var_DTCompra, pedidos.total AS var_total, (CASE WHEN NotaFiscal.cod_pedido = pedidos.cod_pedido THEN 'SIM' ELSE 'NÃO' END) AS Status " & _
           "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente LEFT JOIN NotaFiscal ON NotaFiscal.cod_pedido = pedidos.cod_pedido WHERE (MONTH(pedidos.data_compra) = " & cboMesPedidos.ListIndex + 1 & ") AND (YEAR(pedidos.data_compra) = " & cboAnoPedidos & ") ORDER BY pedidos.cod_pedido;"
End If
   
   Set r = dbData.OpenRecordset(sSQL, totalRegistros)
   FormatarGridPedidos r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   'MOSTRAR A QUANTIDADE REGISTROS
   lblQuantPedidos.Caption = Format(totalRegistros, "00")
End Sub

Private Sub FormatarGridPedidos(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With GridPedidos
       .Clear
       .Cols = 7
       .Rows = 2
           
       .ColWidth(0) = 0
       .ColWidth(1) = 900
       .ColWidth(2) = 900
       .ColWidth(3) = 1100
       .ColWidth(4) = 4000
       .ColWidth(5) = 2000
       .ColWidth(6) = 2000
 
      
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .TextMatrix(0, 1) = "PEDIDO"
      .TextMatrix(0, 2) = "NFE"
      .TextMatrix(0, 3) = "COMPRA"
      .TextMatrix(0, 4) = "CLIENTE"
      .TextMatrix(0, 5) = "TIPO"
      .TextMatrix(0, 6) = "VALOR"

      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = rTabela("var_codped")
            .TextMatrix(.Rows - 1, 2) = rTabela("status")
            .TextMatrix(.Rows - 1, 3) = Format(rTabela("var_dtcompra"), "dd/mm/yy")
            .TextMatrix(.Rows - 1, 4) = rTabela("var_Nome")
            .TextMatrix(.Rows - 1, 5) = rTabela("var_tipoPGTO")
            .TextMatrix(.Rows - 1, 6) = Format(rTabela("var_total"), ocMONEY)
            
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
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
   
'   lblValor.Caption = Format(SomaGrid(GridPedidos, 5), ocMONEY)
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Label47_Click()

End Sub

Private Sub Label48_Click()

End Sub


Private Sub Label49_Click()

End Sub

Private Sub Label44_Click()

End Sub

Private Sub mskConNotaFinal_GotFocus()
SelectControl mskConNotaFinal
End Sub

Private Sub mskConNotaFinal_KeyPress(KeyAscii As Integer)
mskConNotaFinal.Mask = "##/##/##"
End Sub


Private Sub mskConNotaFinal_LostFocus()
If mskConNotaFinal.Text = "" Or mskConNotaFinal.Text = "__/__/__" Then
   mskConNotaFinal.Mask = ""
   mskConNotaFinal.Text = ""
   Exit Sub
Else
   If IsDate(mskConNotaFinal.Text) Then
      'cmdLocalizar.SetFocus
   Else
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      mskConNotaFinal.SetFocus
      SelectControl mskConNotaFinal
   End If
End If
End Sub


Private Sub mskConNotaInicial_GotFocus()
SelectControl mskConNotaInicial
End Sub


Private Sub mskConNotaInicial_KeyPress(KeyAscii As Integer)
mskConNotaInicial.Mask = "##/##/##"
End Sub


Private Sub mskConNotaInicial_LostFocus()
If mskConNotaInicial.Text = "" Or mskConNotaInicial.Text = "__/__/__" Then
   mskConNotaInicial.Mask = ""
   mskConNotaInicial.Text = ""
   Exit Sub
Else
   If IsDate(mskConNotaInicial.Text) Then
      'cmdLocalizar.SetFocus
   Else
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      mskConNotaInicial.SetFocus
      SelectControl mskConNotaInicial
   End If
End If

End Sub

Private Sub mskEmissao_KeyPress(KeyAscii As Integer)
mskEmissao.Mask = "##/##/####"
End Sub

Private Sub mskHora_KeyPress(KeyAscii As Integer)
mskHora.Mask = "##:##:##"
End Sub

Private Sub mskSaida_KeyPress(KeyAscii As Integer)
mskSaida.Mask = "##/##/####"
End Sub

Private Sub TxtCodCliente_Change()
Dim TbClientes As New ADODB.Recordset

If txtCodCliente.Text = "" Then Exit Sub

RsOpen TbClientes, "SELECT * FROM cliente WHERE codigo = " & Val(txtCodCliente.Text)
If TbClientes.EOF And TbClientes.BOF Then
    'MsgBox "Código do cliente não foi localizado no sistema. Verifique.", vbCritical, "Online Commerce": Exit Sub
Else
    txtCliEndereco.Text = ValidateNull(TbClientes("endereco"))
    txtCliNum.Text = ValidateNull(TbClientes("numero"))
    txtCliBairro.Text = ValidateNull(TbClientes("bairro"))
    txtCliCidade.Text = ValidateNull(TbClientes("cidade"))
    txtCliUF.Text = ValidateNull(TbClientes("estado"))
    txtCliIBGE.Text = ValidateNull(TbClientes("CodigoIBGE"))
    txtCliCPF.Text = ValidateNull(TbClientes("cpf"))
    txtCliIE.Text = ValidateNull(TbClientes("ie"))
End If
End Sub

Private Sub txtDesc_Validate(Cancel As Boolean)
Calcular_Desconto
End Sub

Private Sub txtEdit_KeyUp(KeyCode As Integer, Shift As Integer)
   'Exit Sub
   If KeyCode = 38 Then
      If GridNotasItens.Row - 1 = 0 Then ShowMsg "VOCÊ JÁ ESTÁ NA PRIMEIRA LINHA !!!", vbExclamation: Exit Sub
      GridNotasItens.Row = iRow - 1
      GridNotasItens.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
      GridNotasItens_Click
   
   ElseIf KeyCode = 40 Then
      If GridNotasItens.Rows = GridNotasItens.Row + 1 Then ShowMsg "VOCÊ JÁ ESTÁ NA ULTIMA LINHA !!!", vbExclamation: Exit Sub
      GridNotasItens.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
      GridNotasItens.Row = iRow + 1
      GridNotasItens_Click
   End If
End Sub
Private Sub txtEdit_LostFocus()
   GridNotasItens.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
   txtEdit.Visible = False

AtualizarGrid_Itens
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then SendK vbKeyTab
If KeyCode = vbKeyUp Then
    SendK vbKeyTab
    KeyCode = 0
ElseIf KeyCode = vbKeyDown Then
    SendK vbKeyTab
    KeyCode = 0
End If

If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
Set moCombo = New cComboHelper
Frm_NFCe.Tab = 0

Me.Left = (Tela_Principal.ScaleWidth - Me.Width) / 2
Me.Top = (Tela_Principal.ScaleHeight - Me.Height) / 2

cboIndicadorPagamento.Clear
cboIndicadorPagamento.AddItem "0 - Pagamento à vista", 0
cboIndicadorPagamento.AddItem "1 - Pagamento à prazo", 1
cboIndicadorPagamento.AddItem "2 - Outros", 2

'0 - Sem geração de DANFE|1 - DANFE normal, Retrato|2 - DANFE normal, Paisagem|3 - DANFE Simplificado|4 - DANFE NFC-e|5 - DANFE NFC-e em mensagem eletrônica
cboFormatoDANFe.Clear
cboFormatoDANFe.AddItem "0 - Sem geração de DANFE", 0
cboFormatoDANFe.AddItem "1 - DANFE normal, Retrato", 0
cboFormatoDANFe.AddItem "2 - DANFE normal, Paisagem", 2
cboFormatoDANFe.AddItem "3 - DANFE Simplificado", 3
cboFormatoDANFe.AddItem "4 - DANFE NFC-e", 4
cboFormatoDANFe.AddItem "5 - DANFE NFC-e em mensagem eletrônica", 5

'1 - Emissão normal|2 - Contingência FS-IA|3 - Contingência SCAN|4  - Contingência DPEC|5 - Contingência FS-DA|6 - Contingência SVC-AN|7 - Contingência SVC-RS|9 - Contingência off-line da NFC-e
cboTipoEmissao.Clear
cboTipoEmissao.AddItem "1 - Normal", 0
cboTipoEmissao.AddItem "2 - Contingência FS", 1
cboTipoEmissao.AddItem "3 - Contingência SCAN", 2
cboTipoEmissao.AddItem "4 - Contingência DPEC", 3
cboTipoEmissao.AddItem "5 - Contingência FS-DA", 4
cboTipoEmissao.AddItem "6 - Contingência SVC-AN", 5
cboTipoEmissao.AddItem "7 - Contingência SVC-RS", 6
cboTipoEmissao.AddItem "9 - Contingência off-line da NFC-e", 7

cmdNovo.Enabled = True
cmdSalvar.Enabled = False
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdImprimir.Enabled = False
cmdTransmitir.Enabled = False
'cmdCancelarNota.Enabled = False


End Sub

Private Sub GridNotas_DblClick()
'Clear_Controls
'LimparCamposItens
If cmdSalvar.Enabled = True Then
    MsgBox "Existem um NFCe em aberto, Salve-a ou Cancele-a!", vbExclamation, "Online Commerce": Frm_NFCe.Tab = 0: Exit Sub
Else
    RsOpen TbNotas, "SELECT *, " & _
                    "(CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 0 THEN 'Enviada' ELSE (CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) AS Status " & _
                    "FROM TbNFCe WHERE IdNFProd = " & GridNotas.TextMatrix(GridNotas.Row, 1)
    Load_Controls
    Frm_NFCe.Tab = 0
End If
End Sub

Private Sub GridNotasItens_Click()
Dim i As Integer

For i = 3 To 5
   If GridNotasItens.ColSel = i Then
      txtEdit.Move GridNotasItens.Left + GridNotasItens.CellLeft, GridNotasItens.Top + GridNotasItens.CellTop, GridNotasItens.CellWidth, GridNotasItens.CellHeight
      txtEdit.Text = GridNotasItens.TextMatrix(GridNotasItens.Row, GridNotasItens.Col)
      txtEdit.Visible = True
      txtEdit.SetFocus
      txtEdit.SelStart = 0
      txtEdit.SelLength = Len(txtEdit.Text)
      iRow = GridNotasItens.Row
      iCol = GridNotasItens.Col
   End If
Next
End Sub

Private Sub GridPedidos_DblClick()
If ShowMsg("Deseja realmente transformar o pedido: " & GridPedidos.TextMatrix(GridPedidos.Row, 1) & " em Nota Fiscal?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
If GridPedidos.TextMatrix(GridPedidos.Row, 2) = "SIM" Then MsgBox "Esse pedido já foi transformado em NFe!", vbInformation, "Online Commerce": Exit Sub

txtCodPedido.Text = (GridPedidos.TextMatrix(GridPedidos.Row, 1))
    
'TransformarPedidoemNFE
GravarPedido

cmdNovo.Enabled = False
cmdSalvar.Enabled = False
cmdCancelar.Enabled = False
cmdImprimir.Enabled = False
cmdTransmitir.Enabled = True
'cmdCancelarNota.Enabled = False
cmdAlterar.Enabled = True
cmdExcluir.Enabled = True

Frm_NFCe.Tab = 0
End Sub


Private Sub Label26_Click()
chkDesc.Value = 1
chkDesc_Click
End Sub

Private Sub lblCodFabrica_Click()
chkCodBarra.Value = 1
chkCodBarra_Click
End Sub

Public Function SomaGrid(var_Grid As MSFlexGrid, Col As Integer) As Currency
   Dim i As Integer, Valor As Currency
   
   Valor = 0
   For i = 0 To var_Grid.Rows - 1
      If IsNumeric(var_Grid.TextMatrix(i, Col)) Then
         Valor = Valor + CCur(var_Grid.TextMatrix(i, Col))
      End If
   Next
   
   SomaGrid = Valor
End Function

Private Sub mskFinalPedidos_GotFocus()
SelectControl mskFinalPedidos
End Sub

Private Sub mskFinalPedidos_KeyPress(KeyAscii As Integer)
mskFinalPedidos.Mask = "##/##/##"
End Sub


Private Sub mskFinalPedidos_LostFocus()
If mskFinalPedidos.Text = "" Or mskFinalPedidos.Text = "__/__/__" Then
   mskFinalPedidos.Mask = ""
   mskFinalPedidos.Text = ""
   Exit Sub
Else
   If IsDate(mskFinalPedidos.Text) Then
      'cmdLocalizar.SetFocus
   Else
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      mskFinalPedidos.SetFocus
      SelectControl mskFinalPedidos
   End If
End If
End Sub

Private Sub mskInicialPedidos_GotFocus()
SelectControl mskInicialPedidos
End Sub

Private Sub mskInicialPedidos_KeyPress(KeyAscii As Integer)
mskInicialPedidos.Mask = "##/##/##"
End Sub


Private Sub mskInicialPedidos_LostFocus()
If mskInicialPedidos.Text = "" Or mskInicialPedidos.Text = "__/__/__" Then
   mskInicialPedidos.Mask = ""
   mskInicialPedidos.Text = ""
   Exit Sub
Else
   If IsDate(mskInicialPedidos.Text) Then
      'cmdLocalizar.SetFocus
   Else
      ShowMsg "DATA INVÁLIDA!" & vbCrLf & "A data digitada está incompleta ou errada." & vbCrLf & "Verifique e digite novamente.", vbInformation
      mskInicialPedidos.SetFocus
      SelectControl mskInicialPedidos
   End If
End If
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1)
End Sub


Private Sub Text10_GotFocus()
Text10.SelStart = 0
Text10.SelLength = Len(Text10)
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii = Asc(".") Then KeyAscii = Asc(",")
ElseIf KeyAscii = Asc(",") Then KeyAscii = Asc(",")
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub

End Sub

Private Sub Text11_GotFocus()
Text11.SelStart = 0
Text11.SelLength = Len(Text11)
End Sub


Private Sub Text12_GotFocus()
Text12.SelStart = 0
Text12.SelLength = Len(Text12)
End Sub


Private Sub txtVolNumeracao_GotFocus()
txtVolNumeracao.SelStart = 0
txtVolNumeracao.SelLength = Len(txtVolNumeracao)
End Sub


Private Sub Text15_GotFocus()
Text15.SelStart = 0
Text15.SelLength = Len(Text15)
End Sub

Private Sub Text16_GotFocus()
Text16.SelStart = 0
Text16.SelLength = Len(Text16)
End Sub

Private Sub Text17_GotFocus()
Text17.SelStart = 0
Text17.SelLength = Len(Text17)
End Sub

Private Sub Text18_GotFocus()
Text18.SelStart = 0
Text18.SelLength = Len(Text18)
End Sub

Private Sub Text19_GotFocus()
Text19.SelStart = 0
Text19.SelLength = Len(Text19)
End Sub

Private Sub Mostrar_Pedido(rTabela As ADODB.Recordset)
If Not rTabela Is Nothing Then

Dim totalRegistros As Long

    'buscar Numero e codigo da nota (autopreenchimento)
    RsOpen TbNotaPedido, "SELECT *, " & _
                "(CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 0 THEN 'Enviada' ELSE (CASE WHEN NFCeEnviada = 1 AND NFCeCancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) AS Status " & _
                "FROM TbNFCe"
                
    If TbNotaPedido.RecordCount > 0 Then totalRegistros = TbNotaPedido.RecordCount
        
    'Clear_Controls
    
    'INICIO DO PREENCHIMENTOS DOS OBJETOS
    If TbNotaPedido.EOF And TbNotaPedido.BOF Then
        txtNumNota.Text = "1"
        txtCodNota.Text = "1"
    Else
        TbNotaPedido.MoveLast
        txtNumNota.Text = TbNotaPedido("NumeNota") + 1
        txtCodNota.Text = TbNotaPedido("IdNFProd") + 1
    End If

    
'    txtNumNota = Format(rTabela("NumeroNota"), "@")
'    txtCodNota = Format(rTabela("CodigoNota"), "@")
    txtCodCliente = Format(rTabela("IDCliente"), "@")
    cboCliente = Format(rTabela("NomeRazSocial"), "@")
    'cboDestOperacao = Format(5102, "@")
    cboDestOperacao.Text = "1 - Operação Interna"
    'Text1 = Format(rTabela("InformacoesAdicionais"), "@")
    cboNatureza = "VENDA DE MERCADORIA"
    'mskEmissao = Format(Date, "dd/mm/yyyy")
    'mskSaida = Format(Date, "dd/mm/yyyy")
    'mskHora = Format(Time(), "HH:MM:ss")
    txtCodTransporte = Format(0, "@")
    cboTransporte = Format(0, "@")
    txtPlaca = Format(0, "@")
    Text10 = Format(0, "@")
    Text11 = Format(0, "@")
    Text12 = Format(0, "@")
    txtVolNumeracao = Format(0, "@")
    Text15 = Format(0, "@")
    
    Text19 = Format(0, "##,##0.00")
    Text16 = Format(0, "##,##0.00")
    txtValorFrete = Format(0, "##,##0.00")
    Text22 = Format(0, "##,##0.00")
    Text21 = Format(0, "##,##0.00")
    Text17 = Format(0, "@")
    Text18 = Format(0, "@")
    Text23 = Format(0, "@")
    cboModFrete = "9 - SEM FRETE"
    
    Text30 = Format(0, "@")
    Text31 = Format(0, "@")
    Text32 = Format(0, "@")
    cboIndicadorPagamento.Text = "0 - Pagamento à vista"
    cboFormatoDANFe.Text = "4 - DANFE NFC-e"
    cboTipoEmissao.Text = "1 - Normal"
End If
End Sub

Private Sub txtDesc_Change()
'On Error GoTo Erro

If txtDesc.Text = "" Or txtValor.Text = "" Then
   txtDesc.Text = "0"
   SelectControl txtDesc
   Exit Sub
End If

Calcular_Desconto
Exit Sub
   
'Erro:
'   ShowMsg "O valor digitado é inválido!", vbExclamation
'   txtDesc.Text = 0
End Sub

Private Sub Calcular_Desconto()
Dim varTotalUnid As Currency
Dim varTotalSemDesc As Currency
Dim varTotalComdesc As Currency
Dim varQuant As Double
Dim varDesc As Currency

If txtDesc.Text = "" Then txtDesc.Text = "0"
If txtQuant.Text = "" Then txtQuant.Text = "1": SelectControl txtDesc
If txtValor.Text = "" Then txtValor.Text = "0"

varTotalUnid = txtValor.Text

varDesc = txtDesc.Text
varQuant = txtQuant.Text
varTotalSemDesc = CCur(varTotalUnid) * CDbl(varQuant)
varTotalComdesc = CCur(varTotalUnid) * CDbl(varQuant) - CCur(varDesc)

If txtValor.Text = "0" Or txtValor.Text = "0,00" Then Exit Sub
If txtDesc.Text = "" Then txtDesc.Text = Format(0, ocMONEY)

If txtDesc.Text <> "0,00" Then
    txtSubTotal.Text = Format(varTotalComdesc, ocMONEY)
Else
    txtSubTotal.Text = Format(varTotalSemDesc, ocMONEY)
End If
End Sub

Private Sub txtDesc_GotFocus()
SelectControl txtDesc
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub

Private Sub txtDesc_LostFocus()
On Error GoTo erro

If txtDesc.Text = "" Or txtValor.Text = "" Then
   txtDesc.Text = 0
   SelectControl txtDesc
   Exit Sub
End If

Calcular_Desconto
txtDesc.Text = Format(txtDesc.Text, ocMONEY)
Exit Sub
   
erro:
   ShowMsg "O valor digitado é inválido!", vbExclamation
   txtDesc.Text = 0
End Sub

Private Sub txtnumnota_GotFocus()
txtNumNota.SelStart = 0
txtNumNota.SelLength = Len(txtNumNota)
End Sub

Private Sub Text21_GotFocus()
Text21.SelStart = 0
Text21.SelLength = Len(Text21)
End Sub

Private Sub Text22_GotFocus()
Text22.SelStart = 0
Text22.SelLength = Len(Text22)
End Sub

Private Sub Text23_GotFocus()
Text23.SelStart = 0
Text23.SelLength = Len(Text23)
End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtPlaca_GotFocus()
txtPlaca.SelStart = 0
txtPlaca.SelLength = Len(txtPlaca)
End Sub

Private Sub txtPlaca_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtQuant_Validate(Cancel As Boolean)
Calcular_Desconto
End Sub

Private Sub txtSubTotal_GotFocus()
SelectControl txtSubTotal
End Sub


Private Sub txtValor_GotFocus()
SelectControl txtValor
End Sub

Private Sub txtValor_Validate(Cancel As Boolean)
Calcular_Desconto
End Sub

Private Sub txtValorFrete_GotFocus()
txtValorFrete.SelStart = 0
txtValorFrete.SelLength = Len(txtValorFrete)
End Sub

Private Sub txtValorFrete_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii = Asc(".") Then KeyAscii = Asc(",")
ElseIf KeyAscii = Asc(",") Then KeyAscii = Asc(",")
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub

End Sub

Private Sub txtValorFrete_LostFocus()
If txtValorFrete.Text = "" Then txtValorFrete.Text = "0"
Moeda txtValorFrete
End Sub

Private Sub Text15_KeyDown(KeyCode As Integer, Shift As Integer)
Dim TbMSN As New ADODB.Recordset
On Error GoTo erro
If KeyCode = 13 Then
    If Text15.Text = "" Then
        'Frm_LMensagem.Show 1
        Exit Sub
    End If
    RsOpen TbMSN, "select * from ObservacoesNFe where CodigoObservacao = " & Val(Text15.Text)
    If TbMSN.EOF And TbMSN.BOF Then
        MsgBox "Não foi possivel localizar o código da observação solicitada.", vbCritical, "Online Commerce": Exit Sub
    Else
        Text15.Text = TbMSN("CodigoObservacao")
    End If
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub

End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii = Asc(".") Then KeyAscii = Asc(",")
ElseIf KeyAscii = Asc(",") Then KeyAscii = Asc(",")
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub

End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii = Asc(".") Then KeyAscii = Asc(",")
ElseIf KeyAscii = Asc(",") Then KeyAscii = Asc(",")
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub


End Sub

Private Sub Text16_LostFocus()
If Text16.Text = "" Then Text16.Text = "0"
Moeda Text16
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii = Asc(".") Then KeyAscii = Asc(",")
ElseIf KeyAscii = Asc(",") Then KeyAscii = Asc(",")
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub


End Sub

Private Sub Text17_LostFocus()
If Text17.Text = "" Then Text17.Text = "0"
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii = Asc(".") Then KeyAscii = Asc(",")
ElseIf KeyAscii = Asc(",") Then KeyAscii = Asc(",")
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub

End Sub

Private Sub Text18_LostFocus()
If Text18.Text = "" Then Text18.Text = "0"
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii = Asc(".") Then KeyAscii = Asc(",")
ElseIf KeyAscii = Asc(",") Then KeyAscii = Asc(",")
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub

End Sub

Private Sub Text19_LostFocus()
If Text19.Text = "" Then Text19.Text = "0"
Moeda Text19
End Sub

Private Sub cboDestOperacao_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim TbOP As New ADODB.Recordset
'On Error GoTo erro
'If KeyCode = 13 Then
'    If cboDestOperacao.Text = "" Then
'        cboDestOperacao.SetFocus
'        Exit Sub
'    End If
'    RsOpen TbOP, "SELECT * FROM NaturezaOperacaoNF WHERE CodigoNatureza = " & cboDestOperacao.Text
'    If TbOP.EOF And TbOP.BOF Then
        'MsgBox "Não foi possivel localizar o código da OP no sistema.", vbCritical, "Online Commerce": Exit Sub
'    Else
'        cboNatureza.Text = TbOP("descricao")
'        cboNatureza.SetFocus
'    End If
'End If
'Exit Sub
'erro:
'MsgBox "Erro no sistema: " & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub

Private Sub txtnumnota_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii = Asc(".") Then KeyAscii = Asc(",")
ElseIf KeyAscii = Asc(",") Then KeyAscii = Asc(",")
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub

End Sub

Private Sub Text21_LostFocus()
If Text21.Text = "" Then Text21.Text = "0"
Moeda Text21
End Sub

Private Sub Text22_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii = Asc(".") Then KeyAscii = Asc(",")
ElseIf KeyAscii = Asc(",") Then KeyAscii = Asc(",")
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub

End Sub

Private Sub Text22_LostFocus()
If Text22.Text = "" Then Text22.Text = "0"
Moeda Text22
End Sub

Private Sub txtCodBarra_GotFocus()
txtCodBarra.SelStart = 0
txtCodBarra.SelLength = Len(txtCodBarra)
End Sub


Private Sub txtCodBarra_LostFocus()
Dim sSQL As String
Dim r As ADODB.Recordset

If chkCodBarra.Value = 1 Then
    If txtCodBarra.Text = "" Then Exit Sub
     'txtCodProduto.Text = "":
    sSQL = "SELECT codigo AS var_codprod, descricao AS var_desc, tamanho, REF, fabricante, quant_estoque, unid_medida, CFOP, NCM, ICMSCST, ICMSAliq  FROM produtos WHERE (cod_barra = '" & txtCodBarra.Text & "') AND (ativo = 1);"
    Set r = dbData.OpenRecordset(sSQL)
    
    If Not r.BOF Then
       txtCodProduto.Text = r("var_codprod")
       
       If tipoEmpresa = 4 Then
           cboDescricao.Text = ValidateNull(r("var_desc")) & " /  " & ValidateNull(r("tamanho")) & " / " & ValidateNull(r("fabricante")) & " /  " & r("REF")
           'cboDescricao2.Text = ValidateNull(r("var_desc"))
       Else
          cboDescricao.Text = ValidateNull(r("var_desc"))
          txtUnid.Text = ValidateNull(r("unid_medida"))
          txtCFOP.Text = ValidateNull(r("CFOP"))
          txtCST.Text = ValidateNull(r("ICMSCST"))
          txtNCM.Text = ValidateNull(r("NCM"))
          txtICMS.Text = Format(ValidateNull(r("ICMSAliq")), "##,##0.00")
       End If
       
'       txtQuantAtual.Text = ValidateNull(r("quant_Estoque"))
    Else
       ShowMsg "Produto Inexistente!", vbCritical
       txtCodBarra.Text = ""
       txtCodBarra.SetFocus
       Exit Sub
    End If
    
    MostrarValorVenda
    txtQuant.SetFocus
End If

On Local Error Resume Next
End Sub


Private Sub txtCodCliente_KeyDown(KeyCode As Integer, Shift As Integer)
Dim TbClientes As New ADODB.Recordset
On Error GoTo erro
If KeyCode = 13 Then
    If txtCodCliente.Text = "" Then
        txtCodCliente.SetFocus
        Exit Sub
    End If
    RsOpen TbClientes, "SELECT * FROM cliente WHERE codigo = " & Val(txtCodCliente.Text)
    If TbClientes.EOF And TbClientes.BOF Then
        MsgBox "Código do cliente não foi localizado no sistema. Verifique.", vbCritical, "Online Commerce": Exit Sub
    Else
        cboCliente.Text = TbClientes("nome")
        txtCliEndereco.Text = TbClientes("nome")
        txtCliNum.Text = TbClientes("nome")
        txtCliBairro.Text = TbClientes("nome")
        txtCliCidade.Text = TbClientes("nome")
        txtCliUF.Text = TbClientes("nome")
        txtCliCPF.Text = TbClientes("nome")
        txtCliIE.Text = TbClientes("nome")
        
        cboCliente.SetFocus
    End If
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub


Private Sub CboCliente_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboNatureza_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub mskEmissao_GotFocus()
If mskEmissao.Text = "" Then mskEmissao = Format(Date, "dd/mm/yyyy")

mskEmissao.SelStart = 0
mskEmissao.SelLength = Len(mskEmissao)
End Sub

Private Sub mskSaida_GotFocus()
If mskSaida.Text = "" Then mskSaida = Format(Date, "dd/mm/yyyy")
mskSaida.SelStart = 0
mskSaida.SelLength = Len(mskSaida)
End Sub

Private Sub mskHora_GotFocus()
If mskHora.Text = "" Then mskHora = Format(Time(), "HH:MM:ss")
mskHora.SelStart = 0
mskHora.SelLength = Len(mskHora)
End Sub

Private Sub txtCodNota_Change()
Mostrar_ItensNota
End Sub

Private Sub txtCodTransporte_KeyDown(KeyCode As Integer, Shift As Integer)
Dim TbTransportadora As New ADODB.Recordset
On Error GoTo erro
If KeyCode = 13 Then
    If txtCodTransporte.Text = "" Then
        'Frm_LTransportadora.Show 1
        Exit Sub
    End If
    RsOpen TbTransportadora, "select * from transportadora where codigo=" & Val(txtCodTransporte.Text)
    If TbTransportadora.BOF And TbTransportadora.BOF Then
        MsgBox "Não foi possivel localizar o código da transportadora. Favor verifique.", vbCritical, "Online Commerce": Exit Sub
    Else
        txtCodTransporte.Text = TbTransportadora("codigo")
        cboTransporte.Text = TbTransportadora("razao")
        txtCodTransporte.SetFocus
    End If
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub

Private Sub txtCodTransporte_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii = Asc(".") Then KeyAscii = Asc(",")
ElseIf KeyAscii = Asc(",") Then KeyAscii = Asc(",")
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "Online Commerce": Exit Sub
End Sub


Private Function MostraStatus_F9() As String
 On Error Resume Next
 If TbNotas("NFCeEnviada") And Not TbNotas("NFCeCancelada") Then
   MostraStatus.ForeColor = vbBlue
   MostraStatus_F9$ = "Transmitida/Autorizada"  'Deve retornar uma expressão caractere
 ElseIf TbNotas("NFCeCancelada") Then
   MostraStatus.ForeColor = vbRed
   MostraStatus_F9$ = "Cancelada"  'Deve retornar uma expressão caractere
 Else
   MostraStatus.ForeColor = vbBlack
   MostraStatus_F9$ = "Em Digitação"  'Deve retornar uma expressão caractere
 End If
End Function

Private Sub LimparGridNotas()
   Dim i As Integer
   
   With GridNotas
      .Visible = False
      .Redraw = False
      
      .Clear
      .Cols = 8
      .Rows = 2
      
      .ColWidth(0) = 300
      .ColWidth(1) = 500
      .ColWidth(2) = 1000
      .ColWidth(3) = 1100
      .ColWidth(4) = 2000
      .ColWidth(5) = 3000
      .ColWidth(6) = 1500
      .ColWidth(7) = 1100
      
      'CodigoNota, NumeroNota, DataEmissao, NaturezaOperacao, RazaoSocial, ValorNota
      .TextMatrix(0, 1) = "CÓD"
      .TextMatrix(0, 2) = "NUM NF"
      .TextMatrix(0, 3) = "DATA"
      .TextMatrix(0, 4) = "NATUREZA OP"
      .TextMatrix(0, 5) = "DESTINATÁRIO"
      .TextMatrix(0, 6) = "VALOR NOTA"
      .TextMatrix(0, 7) = "STATUS"
      
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
      
      i = 1
      
      'ALINHAMENTO
      .ColAlignment(0) = 1
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 1
      .ColAlignment(6) = 2
      .ColAlignment(7) = 1
      .Rows = .Rows + 1
      
      i = i + 1
      .Rows = .Rows - 1
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 2
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 3
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'GridNotas.ColWidth(0) = 400
      'GridNotas.Rows = 11
      GridNotas.Col = 0
      
      .Visible = True
      .Redraw = True
   End With
End Sub

Private Sub LimparGridItensNota()
   Dim i As Integer
   
   With GridNotasItens
      .Visible = False
      .Redraw = False
      
      .Clear
      .Cols = 8
      .Rows = 2
      
      .ColWidth(0) = 200
      .ColWidth(1) = 2000
      .ColWidth(2) = 1000
      .ColWidth(3) = 1000
      .ColWidth(4) = 1000
      .ColWidth(5) = 1200
      .ColWidth(6) = 1500
      .ColWidth(7) = 1500
      
      'CodigoProduto, NomeProduto, CST, Unidade, Qtde, Valor, SubTotal
      .TextMatrix(0, 1) = "CÓDIGO"
      .TextMatrix(0, 2) = "DESCRIÇÃO"
      .TextMatrix(0, 3) = "CST"
      .TextMatrix(0, 4) = "UND"
      .TextMatrix(0, 5) = "QTDE"
      .TextMatrix(0, 6) = "VALOR"
      .TextMatrix(0, 7) = "SUBTOTAL"
      
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
      
      i = 1
      
      'ALINHAMENTO
      .ColAlignment(0) = 1
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 2
      .ColAlignment(6) = 2
      .ColAlignment(7) = 2
      .Rows = .Rows + 1
      
      i = i + 1
      .Rows = .Rows - 1
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 2
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 3
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'GridNotasItens.ColWidth(0) = 400
      'GridNotasItens.Rows = 11
      GridNotasItens.Col = 0
      
      .Visible = True
      .Redraw = True
   End With
End Sub


Public Sub FormatarGridNotas(rTabela As ADODB.Recordset)
   Dim i As Integer, j As Integer
   
   With GridNotas
      .Visible = False
      .Redraw = False
      
      .Clear
      .Cols = 8
      .Rows = 2
      
      .ColWidth(0) = 300
      .ColWidth(1) = 500
      .ColWidth(2) = 1000
      .ColWidth(3) = 1100
      .ColWidth(4) = 2000
      .ColWidth(5) = 3000
      .ColWidth(6) = 1500
      .ColWidth(7) = 2000
      
      'CodigoNota, NumeroNota, DataEmissao, NaturezaOperacao, RazaoSocial, ValorNota
      .TextMatrix(0, 1) = "CÓD"
      .TextMatrix(0, 2) = "NUM NF"
      .TextMatrix(0, 3) = "DATA"
      .TextMatrix(0, 4) = "NATUREZA OP"
      .TextMatrix(0, 5) = "DESTINATÁRIO"
      .TextMatrix(0, 6) = "VALOR NOTA"
      .TextMatrix(0, 7) = "STATUS"
      
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
      
      i = 1
      
      'ALINHAMENTO
      .ColAlignment(0) = 1
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 1
      .ColAlignment(6) = 2
      .ColAlignment(7) = 2
      
      'CodigoNota, NumeroNota, DataEmissao, NaturezaOperacao, RazaoSocial, ValorNota
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = rTabela("IdNFProd")
            .TextMatrix(.Rows - 1, 2) = Format(rTabela("NumeNota"), "000000")
            .TextMatrix(.Rows - 1, 3) = Format(rTabela("DataEmissao"), "dd/mm/yy")
            .TextMatrix(.Rows - 1, 4) = rTabela("NaturezaOperacao")
            .TextMatrix(.Rows - 1, 5) = rTabela("NomeRazSocial")
            .TextMatrix(.Rows - 1, 6) = Format(rTabela("Valor_NF"), ocMONEY)
            .TextMatrix(.Rows - 1, 7) = rTabela("Status")
            rTabela.MoveNext
            .Rows = .Rows + 1
            i = i + 1
         Loop
      End If
      
      .Rows = .Rows - 1
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 2
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 3
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
              
     'GridNotas.ColWidth(0) = 400
      'GridNotas.Rows = 11
      GridNotas.Col = 0
            
      .Visible = True
      .Redraw = True
   End With
End Sub

Private Sub txtQuant_Change()
If txtQuant.Text = "" Or txtValor.Text = "" Then
   txtQuant.Text = "1"
   SelectControl txtQuant
   Exit Sub
End If

Calcular_Desconto
Exit Sub
End Sub

Private Sub txtQuant_GotFocus()
SelectControl txtQuant
End Sub


Private Sub txtQuant_LostFocus()
'Calcular_Total
Calcular_Desconto
End Sub


Private Sub txtValor_Change()
'Calcular_Total
Calcular_Desconto
End Sub


Private Sub txtValor_LostFocus()
'Calcular_Total
Calcular_Desconto
End Sub


