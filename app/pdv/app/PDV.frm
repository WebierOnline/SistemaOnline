VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSComm32.Ocx"
Begin VB.Form PDV 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11160
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   15375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "PDV.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "PDV.frx":0ECA
   ScaleHeight     =   11160
   ScaleWidth      =   15375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Left            =   1080
      Top             =   10200
   End
   Begin VB.PictureBox picAguarde 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   6360
      Picture         =   "PDV.frx":FEAD
      ScaleHeight     =   1095
      ScaleWidth      =   2895
      TabIndex        =   161
      Top             =   5640
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Frame frmCaixaFechado 
      BackColor       =   &H00C0C0FF&
      Caption         =   "AVISO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6540
      TabIndex        =   148
      Top             =   3720
      Visible         =   0   'False
      Width           =   7755
      Begin ChamaleonBtn.chameleonButton cmdOrcamento 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   4080
         TabIndex        =   149
         Top             =   960
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "2 - ORÇAMENTO"
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
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "PDV.frx":10EE5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAbrirCaixa 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   1500
         TabIndex        =   150
         Top             =   960
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "1 - ABRIR O CAIXA"
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
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "PDV.frx":10F01
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SEU CAIXA ENCONTRA-SE FECHADO!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1140
         TabIndex        =   151
         Top             =   360
         Width           =   5760
      End
   End
   Begin VB.Frame frmAvancado 
      BackColor       =   &H00C0FFFF&
      Caption         =   "MENU AVANÇADO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5355
      Left            =   7980
      TabIndex        =   81
      Top             =   2340
      Visible         =   0   'False
      Width           =   3135
      Begin ChamaleonBtn.chameleonButton cmdAvanClientes 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   180
         TabIndex        =   82
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Cadastro de Clientes"
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
         MICON           =   "PDV.frx":10F1D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAvanProdutos 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   180
         TabIndex        =   83
         Top             =   840
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Cadastro de Produtos"
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
         MICON           =   "PDV.frx":10F39
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAvanFinanceiro 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   180
         TabIndex        =   84
         Top             =   1560
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Financeiro"
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
         MICON           =   "PDV.frx":10F55
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAvanPedReabrir 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   180
         TabIndex        =   85
         Top             =   4320
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Consultar"
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
         MICON           =   "PDV.frx":10F71
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAvanNFCe 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   180
         TabIndex        =   86
         Top             =   4920
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Consultar"
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
         MICON           =   "PDV.frx":10F8D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAvanVendaReiniciar 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   87
         Top             =   4620
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Reiniciar"
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
         MICON           =   "PDV.frx":10FA9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAvanVendaPausar 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   90
         Top             =   4260
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Pausar"
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
         MICON           =   "PDV.frx":10FC5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAvanVendaTransferir 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   91
         Top             =   4980
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Transferir"
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
         MICON           =   "PDV.frx":10FE1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAvanSaidaProd 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   180
         TabIndex        =   142
         Top             =   1200
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Saída de Produtos"
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
         MICON           =   "PDV.frx":10FFD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAvanCarne 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   180
         TabIndex        =   154
         Top             =   3360
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Carnê"
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
         MICON           =   "PDV.frx":11019
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAvanEtiquetas 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   180
         TabIndex        =   155
         Top             =   3720
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Etiquetas"
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
         MICON           =   "PDV.frx":11035
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAvanRecibo 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   180
         TabIndex        =   156
         Top             =   2640
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Recibo"
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
         MICON           =   "PDV.frx":11051
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdAvanRecAvulso 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   180
         TabIndex        =   157
         Top             =   3000
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Recibo Avulso"
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
         MICON           =   "PDV.frx":1106D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdLicenca 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   180
         TabIndex        =   158
         Top             =   1920
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Licença"
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
         MICON           =   "PDV.frx":11089
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblNfce2 
         AutoSize        =   -1  'True
         Caption         =   "!"
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
         Left            =   1860
         TabIndex        =   144
         Top             =   4260
         Width           =   75
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NFCe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   240
         TabIndex        =   93
         Top             =   4680
         Width           =   2685
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "IMPRESSÕES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   240
         TabIndex        =   92
         Top             =   2400
         Width           =   2685
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "VENDAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   240
         TabIndex        =   89
         Top             =   4080
         Width           =   2685
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "GERAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   180
         TabIndex        =   88
         Top             =   240
         Width           =   2685
      End
   End
   Begin MSComctlLib.ListView lstCashBack 
      Height          =   2895
      Left            =   7320
      TabIndex        =   160
      Top             =   3600
      Visible         =   0   'False
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   5106
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Frame frmTipoVenda 
      Caption         =   "Tipo de Venda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   6300
      TabIndex        =   99
      Top             =   3840
      Width           =   8295
      Begin VB.Frame Frame15 
         Caption         =   "ATACADO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   4200
         TabIndex        =   103
         Top             =   360
         Width           =   3915
         Begin ChamaleonBtn.chameleonButton cmdAV 
            CausesValidation=   0   'False
            Height          =   495
            Left            =   180
            TabIndex        =   104
            Top             =   480
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "3 - À VISTA"
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
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "PDV.frx":110A5
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdAP 
            CausesValidation=   0   'False
            Height          =   495
            Left            =   2040
            TabIndex        =   105
            Top             =   480
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "4 - CRÉDITO"
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
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "PDV.frx":110C1
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
      Begin VB.Frame Frame14 
         Caption         =   "VAREJO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   180
         TabIndex        =   100
         Top             =   360
         Width           =   3915
         Begin ChamaleonBtn.chameleonButton cmdVV 
            CausesValidation=   0   'False
            Height          =   495
            Left            =   180
            TabIndex        =   101
            Top             =   480
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "1 - À VISTA"
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
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "PDV.frx":110DD
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonBtn.chameleonButton cmdVP 
            CausesValidation=   0   'False
            Height          =   495
            Left            =   2040
            TabIndex        =   102
            Top             =   480
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "2 - À PRAZO"
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
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "PDV.frx":110F9
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
   Begin VB.Frame frmProdutoNaoCadastrado 
      BackColor       =   &H00C0C0FF&
      Caption         =   "AVISO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6660
      TabIndex        =   131
      Top             =   3840
      Visible         =   0   'False
      Width           =   7755
      Begin ChamaleonBtn.chameleonButton cmdCadastarProduto 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   2640
         TabIndex        =   132
         Top             =   960
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "F2 - CADASTRO RÁPIDO"
         ENAB            =   0   'False
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
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "PDV.frx":11115
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdUsarCadastrado 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   240
         TabIndex        =   134
         Top             =   960
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "F1 - USAR PRODUTO AVULSO"
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
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "PDV.frx":11131
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdNaoCadastrar 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   5160
         TabIndex        =   135
         Top             =   960
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "F3 - NÃO CADASTRAR"
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
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "PDV.frx":1114D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUTO NÃO CADASTRADO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         TabIndex        =   133
         Top             =   360
         Width           =   4620
      End
   End
   Begin VB.Frame frmProdutoAvulso 
      BackColor       =   &H00C0C0FF&
      Caption         =   "PRODUTO AVULSO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   8280
      TabIndex        =   136
      Top             =   4260
      Visible         =   0   'False
      Width           =   4395
      Begin VB.TextBox txtDescProdAvulso 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   140
         Text            =   "PRODUTO AVULSO"
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtValorProdAvulso 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   138
         Top             =   600
         Width           =   1215
      End
      Begin ChamaleonBtn.chameleonButton cmdOKProdAvulso 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   3780
         TabIndex        =   139
         Top             =   600
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "OK"
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
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "PDV.frx":11169
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
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
         Left            =   2520
         TabIndex        =   141
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
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
         TabIndex        =   137
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      TabIndex        =   153
      Top             =   7800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer timerBackup 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   120
      Top             =   10200
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      TabIndex        =   152
      Top             =   7320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtQuant 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2940
      MaxLength       =   6
      TabIndex        =   1
      Top             =   8880
      Width           =   2835
   End
   Begin VB.Frame frmProduto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Informação do Produto"
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
      Height          =   1155
      Left            =   540
      TabIndex        =   124
      Top             =   3780
      Visible         =   0   'False
      Width           =   5115
      Begin VB.TextBox txtInfVenda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         TabIndex        =   129
         Top             =   720
         Width           =   795
      End
      Begin VB.TextBox txtInfMargem 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   128
         Top             =   720
         Width           =   795
      End
      Begin VB.TextBox txtInfCusto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   127
         Top             =   720
         Width           =   795
      End
      Begin VB.TextBox txtInfQuant 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   126
         Top             =   720
         Width           =   795
      End
      Begin VB.TextBox txtInfDesc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   125
         Top             =   360
         Width           =   4875
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "x"
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
         Left            =   4920
         TabIndex        =   130
         Top             =   120
         Width           =   105
      End
   End
   Begin VB.TextBox txtNivel 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10740
      TabIndex        =   120
      Top             =   8940
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtHoraCompra 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8760
      TabIndex        =   117
      Top             =   9840
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.PictureBox frmSenha 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   11160
      ScaleHeight     =   1425
      ScaleWidth      =   2025
      TabIndex        =   95
      Top             =   2340
      Visible         =   0   'False
      Width           =   2055
      Begin VB.ComboBox cboUsuario 
         Height          =   315
         Left            =   120
         TabIndex        =   94
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtNivelUsuario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   110
         Top             =   60
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtCodUsuario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         TabIndex        =   108
         Top             =   60
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtSenha 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   96
         Top             =   1020
         Width           =   1335
      End
      Begin ChamaleonBtn.chameleonButton cmdSenha 
         Height          =   315
         Left            =   1500
         TabIndex        =   97
         Top             =   1020
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
         MICON           =   "PDV.frx":11185
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSMask.MaskEdBox mskCPF 
         Height          =   315
         Left            =   120
         TabIndex        =   119
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   14
         Mask            =   "###.###.###-##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário"
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
         TabIndex        =   109
         Top             =   120
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   195
         Left            =   120
         TabIndex        =   98
         Top             =   780
         Width           =   555
      End
   End
   Begin VB.TextBox txtDataCompra 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8760
      TabIndex        =   107
      Top             =   9540
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.CommandButton cmdMinimizar 
      Caption         =   "-"
      CausesValidation=   0   'False
      Height          =   255
      Left            =   14760
      TabIndex        =   77
      Top             =   0
      Width           =   255
   End
   Begin MSComctlLib.ListView lstBusca 
      Height          =   2895
      Left            =   360
      TabIndex        =   74
      Top             =   6720
      Visible         =   0   'False
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5106
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Frame frmMaquina 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Maquina"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   6240
      TabIndex        =   70
      Top             =   8460
      Visible         =   0   'False
      Width           =   2595
      Begin VB.ComboBox cboMaquina 
         Height          =   315
         Left            =   120
         TabIndex        =   71
         Top             =   240
         Width           =   1815
      End
      Begin ChamaleonBtn.chameleonButton cmdMaqOK 
         Height          =   315
         Left            =   1980
         TabIndex        =   72
         Top             =   240
         Width           =   495
         _ExtentX        =   873
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
         MICON           =   "PDV.frx":111A1
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
   Begin VB.TextBox txtCodFunc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10740
      TabIndex        =   67
      Top             =   8640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtCodBarraPeso 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8760
      TabIndex        =   49
      Top             =   10140
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox txtCodItem 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   11160
      TabIndex        =   48
      Top             =   900
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.TextBox txtUnidMed 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   8820
      TabIndex        =   46
      Top             =   900
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtCodPedido 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13020
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      ToolTipText     =   "Código da Venda"
      Top             =   480
      Width           =   1875
   End
   Begin VB.TextBox txtCodProduto 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   9960
      TabIndex        =   44
      Top             =   900
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "x"
      CausesValidation=   0   'False
      Height          =   255
      Left            =   15060
      TabIndex        =   9
      Top             =   0
      Width           =   255
   End
   Begin VB.TextBox txtTotalGeral 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   675
      Left            =   12000
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   8820
      Width           =   3015
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2940
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   10080
      Width           =   2835
   End
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2940
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7680
      Width           =   2835
   End
   Begin ChamaleonBtn.chameleonButton cmdFinalizarAvista 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6060
      TabIndex        =   5
      Top             =   7980
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Venda à Vista"
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "PDV.frx":111BD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdRemover 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   12360
      TabIndex        =   10
      Top             =   7980
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Remover Item"
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
      MICON           =   "PDV.frx":111D9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdFinalizarPrazo 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   7740
      TabIndex        =   6
      Top             =   7980
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Venda à Prazo"
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "PDV.frx":111F5
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
      Height          =   315
      Left            =   9240
      TabIndex        =   62
      Top             =   9000
      Visible         =   0   'False
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
      MICON           =   "PDV.frx":11211
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdOrçamento 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   9420
      TabIndex        =   7
      Top             =   7980
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Orçamento"
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
      MICON           =   "PDV.frx":1122D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdAvancado 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   13560
      TabIndex        =   11
      Top             =   7980
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Avançado"
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
      MICON           =   "PDV.frx":11249
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtCodBarra 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   360
      TabIndex        =   0
      Top             =   6480
      Width           =   5415
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   106
      Top             =   10890
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13608
            MinWidth        =   1411
            Text            =   "ATALHOS: F1 = PARCELAS, F2 = INFO, F3 = QUANT, F10 = À VISTA, F12 = À PRAZO"
            TextSave        =   "ATALHOS: F1 = PARCELAS, F2 = INFO, F3 = QUANT, F10 = À VISTA, F12 = À PRAZO"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2470
            MinWidth        =   2470
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "21/03/2026"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "16:02"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
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
   Begin ChamaleonBtn.chameleonButton cmdCancelarPedido 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   11100
      TabIndex        =   8
      Top             =   7980
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Cancelar"
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "PDV.frx":11265
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdInfProduto 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   14700
      TabIndex        =   12
      Top             =   7980
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "..."
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
      MICON           =   "PDV.frx":11281
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   240
      Top             =   9240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   2024
      NullDiscard     =   -1  'True
      RThreshold      =   1
      RTSEnable       =   -1  'True
      EOFEnable       =   -1  'True
   End
   Begin VB.Frame frmVendaFechamento 
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
      Height          =   6495
      Left            =   6780
      TabIndex        =   47
      Top             =   1440
      Visible         =   0   'False
      Width           =   7515
      Begin VB.CommandButton Command1 
         Caption         =   "X"
         Height          =   195
         Left            =   7320
         TabIndex        =   118
         Top             =   60
         Width           =   195
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H00C0FFFF&
         Height          =   855
         Left            =   3960
         TabIndex        =   113
         Top             =   2640
         Width           =   3435
         Begin VB.TextBox txtRecebido 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   60
            TabIndex        =   22
            Top             =   420
            Width           =   1875
         End
         Begin VB.TextBox txtTroco 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Left            =   1980
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   420
            Width           =   1335
         End
         Begin VB.Label Label29 
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
            Left            =   1980
            TabIndex        =   115
            Top             =   180
            Width           =   570
         End
         Begin VB.Label Label28 
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
            Left            =   120
            TabIndex        =   114
            Top             =   180
            Width           =   885
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Forma Pagamento"
         Height          =   615
         Left            =   3360
         TabIndex        =   112
         Top             =   240
         Width           =   1695
         Begin VB.ComboBox cboTipoPgto 
            Height          =   315
            Left            =   60
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Quant. de Forma de Pgto"
         Height          =   615
         Left            =   5100
         TabIndex        =   111
         Top             =   240
         Width           =   2295
         Begin VB.ComboBox cboQuantForma 
            Height          =   315
            Left            =   60
            TabIndex        =   15
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Timer tmrDebito 
         Enabled         =   0   'False
         Interval        =   150
         Left            =   180
         Top             =   5880
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Usuário"
         Height          =   615
         Left            =   120
         TabIndex        =   66
         Top             =   240
         Width           =   3195
         Begin VB.TextBox txtFuncAP 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   240
            Width           =   2235
         End
         Begin VB.TextBox txtCodFuncAP 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         Height          =   2115
         Left            =   120
         TabIndex        =   55
         Top             =   3900
         Width           =   7275
         Begin ChamaleonBtn.chameleonButton cmdCal1 
            Height          =   315
            Left            =   3780
            TabIndex        =   33
            TabStop         =   0   'False
            Tag             =   "Calendario"
            Top             =   1680
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
            MICON           =   "PDV.frx":1129D
            PICN            =   "PDV.frx":112B9
            PICH            =   "PDV.frx":1360C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.ComboBox cboformaPgto 
            Height          =   315
            Left            =   4500
            TabIndex        =   28
            Top             =   1080
            Width           =   2175
         End
         Begin VB.ComboBox cboFormaPgtoEntrada 
            Height          =   315
            Left            =   1200
            TabIndex        =   26
            Top             =   1080
            Width           =   2175
         End
         Begin VB.ComboBox cboCliente 
            Height          =   315
            Left            =   120
            TabIndex        =   24
            Top             =   420
            Width           =   7035
         End
         Begin VB.TextBox txtValorParc 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   31
            Text            =   "0"
            Top             =   1680
            Width           =   1155
         End
         Begin VB.ComboBox cboQuantParc 
            Height          =   315
            Left            =   120
            TabIndex        =   29
            Text            =   "1"
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox txtCodCliente 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6360
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtValorRest 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3420
            Locked          =   -1  'True
            TabIndex        =   27
            Text            =   "0"
            Top             =   1080
            Width           =   1035
         End
         Begin VB.TextBox txtEntrada 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   25
            Text            =   "0"
            Top             =   1080
            Width           =   1035
         End
         Begin VB.ComboBox cboPrazo 
            Height          =   315
            Left            =   900
            TabIndex        =   30
            Text            =   "30"
            Top             =   1680
            Width           =   675
         End
         Begin MSMask.MaskEdBox mskInicio 
            Height          =   315
            Left            =   2820
            TabIndex        =   32
            Top             =   1680
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskTermino 
            Height          =   315
            Left            =   4080
            TabIndex        =   34
            Top             =   1680
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label lblFormaParcelas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Forma de Pagamento"
            Height          =   195
            Left            =   4500
            TabIndex        =   123
            Top             =   840
            Width           =   1515
         End
         Begin VB.Label lblFormaEntrada 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Forma de Pagamento"
            Height          =   195
            Left            =   1200
            TabIndex        =   122
            Top             =   840
            Width           =   1515
         End
         Begin VB.Label lblTermino 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Termino:"
            Height          =   195
            Left            =   4080
            TabIndex        =   65
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Parc.:"
            Height          =   195
            Left            =   1620
            TabIndex        =   64
            Top             =   1440
            Width           =   825
         End
         Begin VB.Label lblQtdeParc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quant:"
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   1440
            Width           =   480
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente"
            Height          =   195
            Left            =   120
            TabIndex        =   61
            Top             =   180
            Width           =   480
         End
         Begin VB.Label lblValorParc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Rest."
            Height          =   195
            Left            =   3420
            TabIndex        =   60
            Top             =   840
            Width           =   780
         End
         Begin VB.Label lblQuantParc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prazo:"
            Height          =   195
            Left            =   900
            TabIndex        =   59
            Top             =   1440
            Width           =   450
         End
         Begin VB.Label lblInicio 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio:"
            Height          =   195
            Left            =   2820
            TabIndex        =   58
            Top             =   1440
            Width           =   420
         End
         Begin VB.Label lblEntrada 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor:"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   840
            Width           =   405
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFFF&
         Height          =   1815
         Left            =   3900
         TabIndex        =   50
         Top             =   840
         Width           =   3435
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
            TabIndex        =   21
            ToolTipText     =   "Pressiona a tecla ""ENTER"" para desconto em dinheiro."
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtAcrescDinheiro 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
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
            Left            =   60
            TabIndex        =   38
            Top             =   480
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtDescDinheiro 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
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
            Left            =   60
            TabIndex        =   36
            Top             =   120
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtDescItens 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            Height          =   315
            Left            =   120
            TabIndex        =   121
            Top             =   1440
            Visible         =   0   'False
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
            TabIndex        =   37
            TabStop         =   0   'False
            Text            =   "0,00"
            Top             =   240
            Width           =   1515
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
            TabIndex        =   18
            ToolTipText     =   "Pressiona a tecla ""ENTER"" para desconto em dinheiro."
            Top             =   600
            Width           =   855
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
            TabIndex        =   41
            TabStop         =   0   'False
            Text            =   "0,00"
            Top             =   1320
            Width           =   1455
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   210
            Left            =   1620
            ScaleHeight     =   210
            ScaleWidth      =   915
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   660
            Width           =   915
            Begin VB.OptionButton optDescPorc 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   480
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   0
               Value           =   -1  'True
               Width           =   435
            End
            Begin VB.OptionButton optDescRS 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Caption         =   "R$"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   225
               Left            =   0
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   0
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   210
            Left            =   1620
            ScaleHeight     =   210
            ScaleWidth      =   1035
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   1020
            Width           =   1035
            Begin VB.OptionButton optAscrescPorc 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   225
               Left            =   480
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   0
               Value           =   -1  'True
               Width           =   435
            End
            Begin VB.OptionButton optAscrescRS 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Caption         =   "R$"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   225
               Left            =   0
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   0
               Width           =   495
            End
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desc."
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
            Left            =   1080
            TabIndex        =   54
            Top             =   660
            Width           =   510
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Acresc."
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
            TabIndex        =   69
            Top             =   1020
            Width           =   660
         End
         Begin VB.Label Label30 
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
            TabIndex        =   53
            Top             =   1380
            Width           =   510
         End
         Begin VB.Label Label31 
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
            TabIndex        =   52
            Top             =   300
            Width           =   840
         End
      End
      Begin ChamaleonBtn.chameleonButton cmdFinalizar 
         Height          =   315
         Left            =   5580
         TabIndex        =   35
         Top             =   6060
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
         MICON           =   "PDV.frx":1595F
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
         Left            =   6480
         TabIndex        =   39
         Top             =   6060
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
         MICON           =   "PDV.frx":1597B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdImpOrcamentoCompleto 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   840
         TabIndex        =   116
         Top             =   2280
         Visible         =   0   'False
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Completa"
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
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "PDV.frx":15997
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblInfoDebito 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "POR FAVOR, ENCAMINHE O CLIENTE PARA GERÊNCIA"
         Height          =   195
         Left            =   240
         TabIndex        =   75
         Top             =   6060
         Visible         =   0   'False
         Width           =   4215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      CausesValidation=   0   'False
      Height          =   6255
      Left            =   6120
      TabIndex        =   4
      Top             =   1620
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   11033
      _Version        =   393216
      BackColorBkg    =   16777215
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblMSG1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Há NFCe não Transmitida!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   12915
      TabIndex        =   159
      Top             =   10380
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Label lblAlerta 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".::ALERTAS::."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   13605
      TabIndex        =   147
      Top             =   9660
      Width           =   990
   End
   Begin VB.Label lblRotuloAberto 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caixa Atual Aberto:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   12780
      TabIndex        =   146
      Top             =   9900
      Width           =   1410
   End
   Begin VB.Label lblDataAberturaCaixa 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00/00/0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   14445
      TabIndex        =   145
      Top             =   9900
      Width           =   810
   End
   Begin VB.Label lblNfce1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Há NFCe não Transmitida!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   12915
      TabIndex        =   143
      Top             =   10140
      Width           =   2145
   End
   Begin VB.Label lblQuantTipo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   13260
      TabIndex        =   80
      Top             =   9660
      Width           =   1755
   End
   Begin VB.Label lblTipoVenda 
      Height          =   315
      Left            =   10800
      TabIndex        =   79
      Top             =   10320
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblTipoPedido 
      Height          =   315
      Left            =   10800
      TabIndex        =   78
      Top             =   9600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblInfoBusca 
      BackStyle       =   0  'Transparent
      Caption         =   "PESQUISANDO PRODUTOS. AGUARDE..."
      Height          =   375
      Left            =   360
      TabIndex        =   76
      Top             =   5640
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Label lblEstornar 
      Height          =   315
      Left            =   10800
      TabIndex        =   73
      Top             =   9960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Image imLogoCupom 
      Height          =   1125
      Left            =   6180
      Picture         =   "PDV.frx":159B3
      Top             =   9600
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   420
      TabIndex        =   42
      Top             =   360
      Width           =   12375
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuVendas 
         Caption         =   "Vendas"
         Begin VB.Menu mnuPausar 
            Caption         =   "Pausar"
         End
         Begin VB.Menu mnuReiniciar 
            Caption         =   "Reiniciar"
         End
         Begin VB.Menu mnuTransferir 
            Caption         =   "Transferir"
         End
      End
   End
End
Attribute VB_Name = "PDV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moCombo As New cComboHelper

Dim varValorRealDesc As Currency    'impressão do valor do desconto em dinheiro na impressão do pedido
Dim varValorRealAcresc As Currency  'impressão do valor do acrescimo em dinheiro na impressão do pedido

Public lNovoCod As Long     'usado em varias partes do sistema para autonumeração
Dim vUsandoCashBack As Boolean  'usando no cashback para não limitar desconto no caso de cashback
Dim var_Venda As Currency   'usado na função adicionar produtos
Dim var_Quant As Double     'usado na função adicionar produtos
Dim var_Data As Date        'usado na função adicionar produtos
Dim vQtde As Double         'usado na função adicionar produtos
Dim var_Custo As Currency   'usado na função adicionar produtos

Dim NumCopias As Integer            'usada para saber o numero de copias para imprimir
Dim ii As Integer                   'usada para saber o numero de copias para imprimir
Dim varTipoImpressaoAV As String    'tipo de impressão da venda avista (cupom ou folha)

Dim LOG_NovoCod As Long
Dim vDataFlexivel As Boolean
Public vQuantItensVenda As Integer  'quantidade de registros na venda
Public vDescItensVenda As Currency  'valor de desconto para cada item

Dim vFabBalanca As String    'balança
Dim Retorno As Long          'balança
Dim Peso As String * 5      'balança

Dim NFCe_OK As Boolean   'verificar se os dados estao ok para emitir NFCe
Dim PararFechamentoVenda As Boolean
Dim vBotaoOrcamento As Boolean
Dim vBotaoOrcAtivo As Boolean

Dim CAIXA_FECHADO As Boolean    'caixa
Dim varCodCaixa As Long         'caixa

Dim Msg_Prop As String
Dim VERIFICAR_QUANTIDADE As Boolean
Dim EXISTENCIA_PRODUTO As Boolean
Dim Passou_Limite As Boolean
Dim Cliente_Debito As Boolean
'Dim varLiberarVendaDevedor As Boolean
Dim vValorRececido As Currency
Dim vValorTroco As Currency

Dim strRecebe As String 'Pegar peso balança
Dim vUsarBalanca As String 'confirmar se tem balança no pc
Dim PesoF4 As Boolean

'Dim Liberar As Boolean  'desabilitei pq nao entendi a necessidade

Dim varQuantParc As String  'usado na criação das parcelas
Dim varValorParc As String  'usado na criação das parcelas
Dim i As Integer            'usado na criação das parcelas

Dim totalRegistros As Long

'TABELA CONFIGURAÇÃO - VARIAVEIS
Public TipoValorVenda As String
Public varTipoValorVenda As String
Public varSegurancaAvancada As String
Public vDeclararRecebedor As String
Public vLimitarCompra As String

Public vCashbackAV As String            'cashback A vista SIM/NÃO
Public vCashbackAP As String            'cashback A prazo SIM/NÃO
Public vCashbackValorAV As String       'cashback Valor vista SIM/NÃO
Public vCashbackValorAP As String       'cashback Valor Prazo SIM/NÃO
Public vCashbackLimite As String        'cashback Limite

Public tipoEmpresa As Integer
Public bFechAV As Boolean       'impressão avista
Public iCopiasAV As Integer     'impressão avista
Public bEntregaAV As Boolean    'impressão avista
Public vImprimirVendaAV As Boolean       'impressão avista
Public vConfImprimirVendaAV As Boolean   'impressão avista
Public vTipoImpressaoVendaAV As Integer       'impressão avista
Public bFechAP As Boolean       'impressão aprazo
Public iCopiasAP As Integer     'impressão aprazo
Public bEntregaAP As Boolean    'impressão aprazo
Public vImprimirVendaAP As Boolean       'impressão aprazo
Public vConfImprimirVendaAP As Boolean   'impressão aprazo
Public vTipoImpressaoVendaAP As Integer       'impressão aprazo
Public bFechORC As Boolean      'impressão orçamento
Public iCopiasORC As Integer    'impressão orçamento
Public bEntregaORC As Boolean   'impressão orçamento
Public bImprORC As Boolean      'impressão orçamento
Public bConfImprORC As Boolean  'impressão orçamento
Public iImprORC As Integer      'impressão orçamento
Public vPortaBalanca As String
Public vNFCeImprimir As String  'Perguntar se deseja imprimir a NFCe
Public vNFCeConfImp As String   'Confirmar a impressão de NFCe
Public vNFCeConfCPF As String   'Confirmar a inclussão do cpf na NFCe
Public vNFCeConfPrazo As String   'Confirmar se vai emitir NFCe nas vendas a prazo
Public vNFCeCombinarImp As String   'Combinar Impressão de Pedido com NFCe
Public vTipoParcelaImpressao As Integer   'se será parcelas resumidas ou detalhada no cabeçalho da impressão

Public varConfCartaoCredito As Integer      'ver se tem acrescimento venda cartao credito
Public varConfCartaodebito As Integer       'ver se tem acrescimento venda cartao debito
Public varLoginFunc As Integer              'ver o tipo de identificação de funcionario 1=login 2=codigo

Dim varTipoPgto As String       'venda a prazo e orçamento
Dim varTipoCartao As String     'venda a prazo e orçamento

'arquivo .ini
Public cCfg As ConfigItem
Public oIni As Ini
Public var_Caixa As String
Public var_Maquina As String
'Public var_ImpTermica As String
'Public var_ImpNFCe As String
Public vConfImprimeNFCeLocal As String
Public bIdentMaq As Boolean
'Public varLoginFunc As String
Public varTipoEtiqueta As String
Public varTipoLogin As String

'desconto
Public vTipoDesc As String
Public vLimitarDesc As String
Public vDescCartaoDebito As String
Public vDescCartaoCredito As String
Public vLiberacaoGerente As String
Public vValorDescFixoAV As String
Public vValorDescFixoAP As String
Public vValorDescGradualAV1 As String
Public vValorDescGradualAP1 As String
Public vValorDescGradualAV2 As String
Public vValorDescGradualAP2 As String
Public vValorDescGradualAV3 As String
Public vValorDescGradualAP3 As String
Public vMargemDescGradual1 As String
Public vMargemDescGradual2 As String
Public vMargemDescGradual3 As String

Dim varNomeBotao As String              'variavel usando para saber qual botao acionou a senha
Dim vCodUsuario As Long             'variavel para saber o que pode liberar para o usuario

'buscar o produto na hora de Venda
Public varUnidMed As String
Public varCodBarra As String
Public varCodProdMed As String
Dim varPeso As String

Dim vEtapa As Integer           'desconto gradual
Dim sSQL As String
Dim r As ADODB.Recordset
Dim sSQL2 As String
Dim r2 As ADODB.Recordset
Dim rCliente As ADODB.Recordset
Dim rNFCe As ADODB.Recordset
Dim rNFCeItens As ADODB.Recordset

'abrir o link para pagamento via aplicativo
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Private Const SW_SHOWNORMAL = 1
Private Sub BuscarClienteConsumidor()
If lblEstornar.Caption <> "ESTORNO" And lblEstornar.Caption <> "REIMPRESSÃO" Then
    'If cboCliente.Text = "" Then
        sSQL = "SELECT DISTINCT nome, codigo FROM cliente WHERE codigo = 1;"
        Set r = dbData.OpenRecordset(sSQL)
        If Not r.BOF Then
            cboCliente.Text = r("nome")
            txtCodCliente.Text = r("codigo")
                    
            If r.State <> 0 Then r.Close
            Set r = Nothing
        Else
            MsgBox "Cliente CONSUMIDOR ausente!", vbInformation, "Aviso do Sistema"
            Exit Sub
        End If
    'End If
End If


moCombo.AttachTo cboCliente

End Sub

Private Sub CriarNovoPedido()
vQuantItensVenda = 0
vDescItensVenda = 0

If varTipoValorVenda = 1 Then
    'verificar se o pedido está livre
    Dim var_NroPedido As Long
    var_NroPedido = ExistePedidoLivre
    
    'Nenhum pedido livre
    If var_NroPedido = -1 Then
       txtCodPedido = AutoNumeracao_Pedido
       dbData.Execute "INSERT INTO pedidos (cod_pedido, data_compra, status_pedido, caixa, maquina, cancelado, reaberto, orcamento) VALUES (" & txtCodPedido.Text & ", '" & Format$(Now, "yyyy-dd-MM") & "', 0, '" & var_Caixa & "', '" & var_Maquina & "', 0, 0, 0);"
    Else
       txtCodPedido = var_NroPedido
    End If
    HabilitaObjetosVenda False
    'If txtCodBarra.Enabled = True Then txtCodBarra.SetFocus
    lblTipoVenda.Caption = ""
ElseIf varTipoValorVenda = 2 Then
    'txtCodBarra.Enabled = False
    'If CAIXA_FECHADO = False Then frmTipoVenda.Visible = True
    If txtCodPedido.Text = "" Then
        frmTipoVenda.Visible = True
        HabilitaObjetosVenda True
    Else
        'frmTipoVenda.Visible = False
        'HabilitaObjetosVenda False
        If txtCodBarra.Enabled = True Then txtCodBarra.SetFocus
    End If
    'txtCodBarra.SetFocus
    lblTipoVenda.Caption = ""
End If
End Sub

Private Sub DANFCeImpressao()
'Dim anoEmes As String, Arquivo As String
'Dim sistNFe As snfe.Util
'Dim NomeImpNFCe As String

'    Set sistNFe = New snfe.Util
'    Set cCfg = sysConfig("NOME_IMP_NFCE")
'    NomeImpNFCe = cCfg.Value
'    Set cCfg = Nothing
    
'    dirXML = SQLExecutaRetorno("SELECT DiretorioXML FROM Empresa", "DiretorioXML", App.Path)
'    dirXML = IIf(Right(dirXML, 1) = "\", dirXML, dirXML & "\")
'    xCaminhoXML = dirXML & "nfe\arquivos\procNFe\" & NFeChaveAcesso & "-procNFe.xml"
'    anoEmes = dirXML & "nfe\arquivos\procNFe\" & Format(Date, "yyyymm") & "\"
'    xCaminhoPDF = dirXML & "nfe\arquivos\PDF\NFe" & NFeChaveAcesso & ".pdf"
'    If Not Existe(xCaminhoXML) Then xCaminhoXML = anoEmes & NFeChaveAcesso & "-procNFe.xml"
    
'    Call sistNFe.DANFCeImprimir(xCaminhoXML, "", "", True, NomeImpNFCe, True, xCaminhoPDF, False, False, 0, "")
End Sub



Private Sub ImprimirVendaAPsemPergunta()
If iCopiasAP <> 0 Then
    'If bEntregaAP Then
    '      If ShowMsg("Desesa Imprimir o pedido para ENTREGAR?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
     '        NumCopias = iCopiasAP + 1
     '     Else
     '        NumCopias = iCopiasAP
     '     End If
    'Else
       NumCopias = iCopiasAP
    'End If
Else
    NumCopias = 1
End If

'If vImprimirVendaAP Then       'Confirma se vai ter impressão
'    If vConfImprimirVendaAP Then
'        If ShowMsg("Desesa Imprimir o pedido?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            For ii = 1 To NumCopias
                If vTipoImpressaoVendaAP = 1 Then
                    Imprimir_Pedido
                ElseIf vTipoImpressaoVendaAP = 2 Then
                    Imprimir_CupomSerrilha
                ElseIf vTipoImpressaoVendaAP = 3 Then
                    Imprimir_CupomGuilhotina
                End If
            Next
        'End If
'    Else
'        For ii = 1 To NumCopias
'            If vTipoImpressaoVendaAP = 1 Then
'                Imprimir_Pedido
'            ElseIf vTipoImpressaoVendaAP = 2 Then
'                Imprimir_CupomSerrilha
'            ElseIf vTipoImpressaoVendaAP = 3 Then
'                Imprimir_CupomGuilhotina
'            End If
'        Next
'    End If
'End If  'final do vConfImprimirVendaAP
''End If      'final do vImprimirVendaAP
End Sub
Private Sub ImprimirVendaAP()
If iCopiasAP <> 0 Then
    If bEntregaAP Then
          If ShowMsg("Desesa Imprimir o pedido para ENTREGAR?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
             NumCopias = iCopiasAP + 1
          Else
             NumCopias = iCopiasAP
          End If
    Else
       NumCopias = iCopiasAP
    End If
Else
    NumCopias = 1
End If

If vImprimirVendaAP Then       'Confirma se vai ter impressão
    If vConfImprimirVendaAP Then
        If ShowMsg("Desesa Imprimir o pedido?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            For ii = 1 To NumCopias
                If vTipoImpressaoVendaAP = 1 Then
                    Imprimir_Pedido
                ElseIf vTipoImpressaoVendaAP = 2 Then
                    Imprimir_CupomSerrilha
                ElseIf vTipoImpressaoVendaAP = 3 Then
                    Imprimir_CupomGuilhotina
                End If
            Next
        End If
    Else
        For ii = 1 To NumCopias
            If vTipoImpressaoVendaAP = 1 Then
                Imprimir_Pedido
            ElseIf vTipoImpressaoVendaAP = 2 Then
                Imprimir_CupomSerrilha
            ElseIf vTipoImpressaoVendaAP = 3 Then
                Imprimir_CupomGuilhotina
            End If
        Next
    End If
End If  'final do vConfImprimirVendaAP
'End If      'final do vImprimirVendaAP
End Sub

Private Sub ImprimirVendaAVsemPergunta()
'Numero de copias
If iCopiasAV <> 0 Then    'saber a quantidade de copias
    'If bEntregaAV Then
    '    If ShowMsg("Desesa Imprimir o pedido para ENTREGAR?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
    '        NumCopias = iCopiasAV + 1
    '    Else
    '        NumCopias = iCopiasAV
    '    End If
    'Else
        NumCopias = iCopiasAV
    'End If
Else
    NumCopias = "1"
End If

'tipos de impressões
'If vImprimirVendaAV Then        'se deseja imprimir
'   If vConfImprimirVendaAV Then     'se é para confirmar se deve perguntar sobre a impressão
'      If ShowMsg("Desesa Imprimir o pedido?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            For ii = 1 To NumCopias
                If vTipoImpressaoVendaAV = 1 Then
                    Imprimir_Pedido
                ElseIf vTipoImpressaoVendaAV = 2 Then
                    Imprimir_CupomSerrilha
                ElseIf vTipoImpressaoVendaAV = 3 Then
                    Imprimir_CupomGuilhotina
                End If
            Next
'      End If
'   Else
'        For ii = 1 To NumCopias
'            If vTipoImpressaoVendaAV = 1 Then
'                Imprimir_Pedido
'            ElseIf vTipoImpressaoVendaAV = 2 Then
'                Imprimir_CupomSerrilha
'            ElseIf vTipoImpressaoVendaAV = 3 Then
'                Imprimir_CupomGuilhotina
'            End If
'        Next
'   End If
'End If
End Sub
Private Sub ImprimirVendaAV()
'Numero de copias
If iCopiasAV <> 0 Then    'saber a quantidade de copias
    If bEntregaAV Then
        If ShowMsg("Desesa Imprimir o pedido para ENTREGAR?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
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

'tipos de impressões
If vImprimirVendaAV Then        'se deseja imprimir
   If vConfImprimirVendaAV Then     'se é para confirmar se deve perguntar sobre a impressão
      If ShowMsg("Desesa Imprimir o pedido?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            For ii = 1 To NumCopias
                If vTipoImpressaoVendaAV = 1 Then
                    Imprimir_Pedido
                ElseIf vTipoImpressaoVendaAV = 2 Then
                    Imprimir_CupomSerrilha
                ElseIf vTipoImpressaoVendaAV = 3 Then
                    Imprimir_CupomGuilhotina
                End If
            Next
      End If
   Else
        For ii = 1 To NumCopias
            If vTipoImpressaoVendaAV = 1 Then
                Imprimir_Pedido
            ElseIf vTipoImpressaoVendaAV = 2 Then
                Imprimir_CupomSerrilha
            ElseIf vTipoImpressaoVendaAV = 3 Then
                Imprimir_CupomGuilhotina
            End If
        Next
   End If
End If
End Sub


Private Sub Mostrar_Desconto()
If vTipoDesc = "1" Then
    txtDesc.Text = FormatNumber(0, 2)
ElseIf vTipoDesc = "2" Then
    If cboTipoPgto.Text = "À VISTA" Then
        txtDesc.Text = FormatNumber(vValorDescFixoAV, 2)
    ElseIf cboTipoPgto.Text = "À PRAZO" Then
        txtDesc.Text = FormatNumber(vValorDescFixoAP, 2)
    ElseIf cboTipoPgto.Text = "ORÇAMENTO" Then
        txtDesc.Text = FormatNumber(vValorDescFixoAV, 2)
    End If
ElseIf vTipoDesc = "3" Then
    Dim vSubtotal As Currency
    vSubtotal = txtSubtotal.Text

If vSubtotal <= vMargemDescGradual1 Then
    vEtapa = 1
ElseIf vSubtotal > vMargemDescGradual1 And vSubtotal <= vMargemDescGradual2 Then
    vEtapa = 2
ElseIf vSubtotal > vMargemDescGradual2 Then
    vEtapa = 3
Else
    vEtapa = 1
End If

    If cboTipoPgto.Text = "À VISTA" Then
        If vEtapa = 1 Then
            txtDesc.Text = FormatNumber(vValorDescGradualAV1, 2)
        ElseIf vEtapa = 2 Then
            txtDesc.Text = FormatNumber(vValorDescGradualAV2, 2)
        ElseIf vEtapa = 3 Then
            txtDesc.Text = FormatNumber(vValorDescGradualAV3, 2)
        End If
    ElseIf cboTipoPgto.Text = "À PRAZO" Then
        If vEtapa = 1 Then
            txtDesc.Text = FormatNumber(vValorDescGradualAP1, 2)
        ElseIf vEtapa = 2 Then
            txtDesc.Text = FormatNumber(vValorDescGradualAP2, 2)
        ElseIf vEtapa = 3 Then
            txtDesc.Text = FormatNumber(vValorDescGradualAP3, 2)
        End If
    ElseIf cboTipoPgto.Text = "ORÇAMENTO" Then
        If vEtapa = 1 Then
            txtDesc.Text = FormatNumber(vValorDescGradualAV1, 2)
        ElseIf vEtapa = 2 Then
            txtDesc.Text = FormatNumber(vValorDescGradualAV2, 2)
        ElseIf vEtapa = 3 Then
            txtDesc.Text = FormatNumber(vValorDescGradualAV3, 2)
        End If
    End If
End If
End Sub

Private Sub MostrarCaixaSenha()
Set oCfg = sysConfig("TIPOLOGIN")
If oCfg.Value = "NOME" Then
    frmSenha.Visible = True
    cboUsuario.Visible = True
    mskCPF.Visible = False
    cboUsuario.Text = ""
    txtCodUsuario.Text = ""
    txtSenha.Text = ""
    Label1.Caption = "Usuário:"
Else
    frmSenha.Visible = True
    cboUsuario.Visible = False
    mskCPF.Visible = True
    txtSenha.Text = ""
    mskCPF.Mask = ""
    mskCPF.Text = ""
    mskCPF.Mask = "###.###.###-##"
    Label1.Caption = "CPF:"
    If mskCPF.Visible = True Then mskCPF.SetFocus
End If
End Sub


Public Sub Permissoes()
    'If LerPermissoesUsuario(vCodUsuario, 3) = True Then
    '    Menu_PROD_Simples.Enabled = True
    'Else
    '    Menu_PROD_Simples.Enabled = False
    'End If
    
    'If LerPermissoesUsuario(vCodUsuario, 4) = True Then
    '    Menu_CAD_Usuario.Enabled = True
    'Else
    '    Menu_CAD_Usuario.Enabled = False
    'End If
    'If LerPermissoesUsuario(vCodUsuario, 5) = True Then
    '    Menu_CONS_Fluxo.Enabled = True
    'Else
    '    Menu_CONS_Fluxo.Enabled = False
    'End If
    'If LerPermissoesUsuario(vCodUsuario, 6) = True Then
    '    Menu_CONS_Lancamentos.Enabled = True
    'Else
    '    Menu_CONS_Lancamentos.Enabled = False
    'End If
    'If LerPermissoesUsuario(vCodUsuario, 7) = True Then
    '    Menu_PROD_Entrada.Enabled = True
    'Else
    '    Menu_PROD_Entrada.Enabled = False
    'End If
    'If LerPermissoesUsuario(vCodUsuario, 8) = True Then
    '    Menu_Entrada_Estoque.Enabled = True
    'Else
    '    Menu_Entrada_Estoque.Enabled = False
    'End If
    If LerPermissoesUsuario(vCodFunc, 9) = True Then
        cmdAvanSaidaProd.Enabled = True
    Else
        cmdAvanSaidaProd.Enabled = False
    End If
    
    'If LerPermissoesUsuario(vCodUsuario, 10) = True Then
    '    Menu_CAD_Empresa.Enabled = True
    'Else
    '    Menu_CAD_Empresa.Enabled = False
    'End If
    'If LerPermissoesUsuario(vCodUsuario, 11) = True Then
    '    Menu_SIS_Config.Enabled = True
    'Else
    '    Menu_SIS_Config.Enabled = False
    'End If
    'If LerPermissoesUsuario(vCodUsuario, 12) = True Then
    '    Menu_Fin_APagar.Enabled = True
    '    cmdContasApagar.Enabled = True
    'Else
    '    Menu_Fin_APagar.Enabled = False
    '    cmdContasApagar.Enabled = False
    'End If
    'If LerPermissoesUsuario(vCodUsuario, 13) = True Then
    '    Menu_Fin_Parcelas.Enabled = True
    '    'cmdContasApagar.Enabled = True
    'Else
    '    Menu_Fin_Parcelas.Enabled = False
    '    'cmdContasApagar.Enabled = False
    
    'End If
    'If LerPermissoesUsuario(vCodUsuario, 14) = True Then
    '    Menu_Fin_AReceber.Enabled = True
    'Else
    '    Menu_Fin_AReceber.Enabled = False
    'End If
    'If LerPermissoesUsuario(vCodUsuario, 15) = True Then
    '    Menu_Fin_Sangria.Enabled = True
    'Else
    '    Menu_Fin_Sangria.Enabled = False
    'End If
    'If LerPermissoesUsuario(vCodUsuario, 16) = True Then
    '    Menu_Fin_Retirada.Enabled = True
    'Else
    '    Menu_Fin_Retirada.Enabled = False
    'End If
    'If LerPermissoesUsuario(vCodUsuario, 17) = True Then
    '    Menu_Fin_Suprimentos.Enabled = True
    'Else
    '    Menu_Fin_Suprimentos.Enabled = False
    'End If
    'If LerPermissoesUsuario(vCodUsuario, 18) = True Then
    '    Menu_Fin_Caixa.Enabled = True
    'Else
    '    Menu_Fin_Caixa.Enabled = False
    'End If
    'If LerPermissoesUsuario(vCodUsuario, 19) = True Then
    '    Menu_CONS_Vendas.Enabled = True
    'Else
    '    Menu_CONS_Vendas.Enabled = False
    'End If
    'If LerPermissoesUsuario(vCodUsuario, 20) = True Then
    '    Menu_CONS_VendasPorProdutos.Enabled = True
    'Else
    '    Menu_CONS_VendasPorProdutos.Enabled = False
    'End If
    'If LerPermissoesUsuario(vCodUsuario, 21) = True Then
    '    Menu_CONS_VendasPorProdutosAgrupados.Enabled = True
    'Else
    '    Menu_CONS_VendasPorProdutosAgrupados.Enabled = False
    'End If
    'If LerPermissoesUsuario(vCodUsuario, 22) = True Then
    '    Menu_CONS_EntradaPorProdutos.Enabled = True
    'Else
    '    Menu_CONS_EntradaPorProdutos.Enabled = False
    'End If
    'If LerPermissoesUsuario(vCodUsuario, 23) = True Then
    '    Menu_CONS_EntradaPorProdAgrupadas.Enabled = True
    'Else
    '    Menu_CONS_EntradaPorProdAgrupadas.Enabled = False
    'End If
    'If LerPermissoesUsuario(vCodUsuario, 24) = True Then
    '    Menu_CONS_EntradavsSaida.Enabled = True
    'Else
    '    Menu_CONS_EntradavsSaida.Enabled = False
    'End If
    'If LerPermissoesUsuario(vCodUsuario, 25) = True Then
    '    Menu_CONS_Comissoes.Enabled = True
    'Else
    '    Menu_CONS_Comissoes.Enabled = False
    'End If
    'If LerPermissoesUsuario(vCodUsuario, 26) = True Then
    '    Menu_CONS_Parcelas.Enabled = True
    'Else
    '    Menu_CONS_Parcelas.Enabled = False
    'End If
    
    If LerPermissoesUsuario(vCodFunc, 27) = True Then
        cmdAvanProdutos.Enabled = True
    Else
        cmdAvanProdutos.Enabled = False
    End If

    If LerPermissoesUsuario(vCodFunc, 28) = True Then
        cmdAvanFinanceiro.Enabled = True
    Else
        cmdAvanFinanceiro.Enabled = False
    End If
    

End Sub

Public Function LerPermissoesUsuario(vCodUser As Long, permissao As Long) As Boolean
sSQL = "SELECT Usuario_Acessos.Cod_Permissao FROM Usuario INNER JOIN Usuario_Acessos ON Usuario.Codigo = Usuario_Acessos.Cod_Usuario WHERE (Usuario_Acessos.Cod_Usuario = " & vCodUser & ") AND Usuario_Acessos.Cod_Permissao = " & permissao & ";"
Set r = dbData.OpenRecordset(sSQL)

If r.EOF And r.BOF Then
   LerPermissoesUsuario = False ' não achou a permissao
Else
   LerPermissoesUsuario = True 'aqui achou
End If
End Function
Private Sub Preencher_FormaPgto()
If cboTipoPgto.Text = "À VISTA" Then
    cboFormaPgto.AddItem "1 - DINHEIRO"
    cboFormaPgto.AddItem "3 - CARTÃO - DÉBITO"
    cboFormaPgto.AddItem "4 - CARTÃO - CRÉDITO"
    cboFormaPgto.AddItem "5 - CHEQUE"
    cboFormaPgto.AddItem "7 - TRANSFERÊNCIA"
    cboFormaPgto.AddItem "8 - DEPOSITO"
    cboFormaPgto.AddItem "9 - FINANCEIRA"
    cboFormaPgto.AddItem "10 - PIX"
Else
    cboFormaPgto.AddItem "2 - PROMISSÓRIA"
    cboFormaPgto.AddItem "5 - CHEQUE"
    cboFormaPgto.AddItem "6 - BOLETO"
End If
End Sub



Private Function ExistePedidoLivre() As Long
Dim sSQL As String
Dim r As ADODB.Recordset
Dim lRet As Long

lRet = -1
sSQL = "SELECT cod_pedido FROM pedidos WHERE (data_compra = '" & Format$(Now, "yyyy-dd-MM") & "') AND (status_pedido = 0) AND (reaberto = 0) AND (maquina = '" & var_Maquina & "');"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then lRet = r("cod_pedido")
If r.State <> 0 Then r.Close
Set r = Nothing

ExistePedidoLivre = lRet
End Function

Sub FlashColor()
   Dim i As Integer
   Dim bCor(1 To 2) As OLE_COLOR, fCor(1 To 2) As OLE_COLOR
   Dim Start As Single, Finish As Single
   
   bCor(1) = RGB(225, 0, 0)
   bCor(2) = frmVendaFechamento.BackColor
   fCor(1) = frmVendaFechamento.BackColor
   fCor(2) = RGB(225, 0, 0)
   
   For i = 1 To 2
      Start = Timer
      Finish = Start + 0.5
      Do
         DoEvents
         lblInfoDebito.BackColor = bCor(i)
         lblInfoDebito.ForeColor = fCor(i)
      Loop While Timer < Finish
   Next
   
   Erase bCor
   Erase fCor
End Sub

Private Sub pTransferirVendaCaixa()
   Dim fPDV As ListarPDV
   Dim bExecutado As Boolean
   
   'Avisa o usuáriio sobre a autorização e solicita confirmação
   If ShowMsg("A transferência deste pedido para outro caixa requer autorização." & vbCr & vbCr & _
      "Deseja realmente executar esta operação?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
   
   Set fPDV = New ListarPDV
   fPDV.Show vbModal
   bExecutado = fPDV.Done
   Unload fPDV
   Set fPDV = Nothing
   
   If bExecutado Then
      'Reinicia o form para uma nova venda
      LimparObjetos_Pedido
      txtTotalGeral.Text = ""
      LimparGrid_Pedido
      LimparObjetos_Prazo
      txtCodPedido.Text = ""
      lblEstornar.Caption = ""
      Form_Load
      txtCodBarra.SetFocus
      
      frmAvancado.Visible = False
      frmVendaFechamento.Visible = False
      'Liberar = False
   End If
End Sub

Private Sub Abrir_Pedido_Reimpressao()
Dim sSQL As String
Dim r As ADODB.Recordset, IdNFProd As Long

If txtCodPedido.Text = "" Then Exit Sub
   
cmdFinalizarAvista.Visible = False
cmdFinalizarPrazo.Visible = False
cmdCancelarPedido.Visible = False
cmdRemover.Visible = False
cmdOrçamento.Visible = False
cmdAvancado.Visible = False
frmAvancado.Visible = False
frmSenha.Visible = False
   
   sSQL = "SELECT * FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
    If Not r.BOF Then
      If r("tipo_pagamento") = "À Prazo" Or r("tipo_pagamento") = "À PRAZO" Then
        cboTipoPgto.Text = "À PRAZO"
      Else
        cboTipoPgto.Text = "À VISTA"
      End If
        
        frmVendaFechamento.Visible = True
        
        If r("TIPO_ACRESCIMO") = "R" Then
           optAscrescRS.Value = True
        Else
           optAscrescPorc.Value = True
        End If

         txtAcresc.Text = FormatNumber(r("valor_acrescimo"), 2)
        
        If r("tipo_desc") = "R" Then
           optDescRS.Value = True
        Else
           optDescPorc.Value = True
        End If
        
        txtDesc.Text = FormatNumber(r("valor_desc"), 2)
        
        If r("pagamento") = "DINHEIRO" Then
            cboFormaPgto.Text = "1 - DINHEIRO"
        ElseIf r("pagamento") = "PROMISSORIA" Then
            cboFormaPgto.Text = "2 - PROMISSÓRIA"
        ElseIf r("pagamento") = "CARTAO" And r("TIPO_CARTAO") = "D" Then
            cboFormaPgto.Text = "3 - CARTÃO - DÉBITO"
        ElseIf r("pagamento") = "CARTAO" And r("TIPO_CARTAO") = "C" Then
            cboFormaPgto.Text = "4 - CARTÃO - CRÉDITO"
        ElseIf r("pagamento") = "CHEQUE" Then
            cboFormaPgto.Text = "5 - CHEQUE"
        ElseIf r("pagamento") = "BOLETO" Then
            cboFormaPgto.Text = "6 - BOLETO"
        End If
         
         txtDataCompra.Text = Format(r("data_compra"), "dd/mm/yyyy")
         txtSubtotal.Text = Format(r("subtotal"), ocMONEY)
         txtTotalDesc.Text = Format(r("total"), ocMONEY)
         'txtValorParc.Text = Format(r("valor_parc"), ocMONEY)
         txtEntrada.Text = Format(r("entrada"), ocMONEY)
         'cboPrazo.Text = ValidateNull(r("prazo"))
         'cboQuantParc.Text = ValidateNull(r("parcelas"))
         'mskInicio.Text = Format(r("vencimento"), "dd/mm/yy")
         txtCodFuncAP.Text = r("cod_funcionario")
         txtCodCliente.Text = r("cod_cliente")
         
         cmdFinalizar.Visible = False
         cmdCancelar.Visible = False
         frmVendaFechamento.Enabled = False
   
      sSQL = "SELECT IdNFProd FROM TbNFCe WHERE Num_OS_VD_Origem  = " & txtCodPedido.Text
      IdNFProd = SQLExecutaRetorno(sSQL, "IdNFProd", 0)
      If IdNFProd > 0 Then
         sSQL = "SELECT NFCeChaveAcesso, NFCeProtocolo, NFCeCancelada, NFCeCanceladaProtocolo, NFCeCanceladaJustificativa FROM TbNFCe WHERE IdNFProd = " & IdNFProd
         NFeChaveAcesso = SQLExecutaRetorno(sSQL, "NFCeChaveAcesso", "")
         'DANFCeImpressao
      End If
   End If
End Sub


Private Sub Adicionar_Produto()
'VerificarUnidadeMedidas

'Dim sSQL As String             'desativei no dia 12/04/2024
'Dim r As ADODB.Recordset       'desativei no dia 12/04/2024

If txtCodBarra.Text = "" Or txtCodProduto.Text = "" Then Exit Sub

'localizar o valor do produto
If varTipoValorVenda = 1 Then
    sSQL = "SELECT DISTINCT produtos.codigo AS vCodProduto, produtos.cod_barra, produtos.quant_estoque as vQuant, produtos.UNID_MEDIDA as vUnid, produtos.PEDIRPESO, " & _
    "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda, " & _
    "(SELECT TOP 1 Produtos_Precos.CUSTO FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS vcusto " & _
    "FROM produtos WHERE (produtos.codigo = '" & txtCodProduto & "') AND (produtos.ativo = 1) " & _
    "ORDER BY produtos.codigo;"
    
ElseIf varTipoValorVenda = 2 Then
    If TipoValorVenda = "VV" Then
        sSQL = "SELECT DISTINCT produtos.codigo AS vCodProduto, produtos.cod_barra, produtos.quant_estoque as vQuant, produtos.UNID_MEDIDA as vUnid, produtos.PEDIRPESO, " & _
        "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda, " & _
    "(SELECT TOP 1 Produtos_Precos.CUSTO FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS vcusto " & _
        "FROM produtos WHERE (produtos.codigo = '" & txtCodProduto & "') AND (produtos.ativo = 1) " & _
        "ORDER BY produtos.codigo;"
    ElseIf TipoValorVenda = "VP" Then
        sSQL = "SELECT DISTINCT produtos.codigo AS vCodProduto, produtos.cod_barra, produtos.quant_estoque as vQuant, produtos.UNID_MEDIDA as vUnid, produtos.PEDIRPESO, " & _
        "(SELECT TOP 1 Produtos_Precos.VALOR_VP FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda, " & _
    "(SELECT TOP 1 Produtos_Precos.CUSTO FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS vcusto " & _
        "FROM produtos WHERE (produtos.codigo = '" & txtCodProduto & "') AND (produtos.ativo = 1) " & _
        "ORDER BY produtos.codigo;"
    ElseIf TipoValorVenda = "AV" Then
        sSQL = "SELECT DISTINCT produtos.codigo AS vCodProduto, produtos.cod_barra, produtos.quant_estoque as vQuant, produtos.UNID_MEDIDA as vUnid, produtos.PEDIRPESO, " & _
        "(SELECT TOP 1 Produtos_Precos.VALOR_AV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda, " & _
    "(SELECT TOP 1 Produtos_Precos.CUSTO FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS vcusto " & _
        "FROM produtos WHERE (produtos.codigo = '" & txtCodProduto & "') AND (produtos.ativo = 1) " & _
        "ORDER BY produtos.codigo;"
    ElseIf TipoValorVenda = "AP" Then
        sSQL = "SELECT DISTINCT produtos.codigo AS vCodProduto, produtos.cod_barra, produtos.quant_estoque as vQuant, produtos.UNID_MEDIDA as vUnid, produtos.PEDIRPESO, " & _
        "(SELECT TOP 1 Produtos_Precos.VALOR_AP FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda, " & _
    "(SELECT TOP 1 Produtos_Precos.CUSTO FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS vcusto " & _
        "FROM produtos WHERE (produtos.codigo = '" & txtCodProduto & "') AND (produtos.ativo = 1) " & _
        "ORDER BY produtos.codigo;"
    Else
        sSQL = "SELECT DISTINCT produtos.codigo AS vCodProduto, produtos.cod_barra, produtos.quant_estoque as vQuant, produtos.UNID_MEDIDA as vUnid, produtos.PEDIRPESO, " & _
        "(SELECT TOP 1 Produtos_Precos.VALOR_VP FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda, " & _
    "(SELECT TOP 1 Produtos_Precos.CUSTO FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS vcusto " & _
        "FROM produtos WHERE (produtos.codigo = '" & txtCodProduto & "') AND (produtos.ativo = 1) " & _
        "ORDER BY produtos.codigo;"
    End If
End If
Set r = dbData.OpenRecordset(sSQL)

If r.BOF Then
        MsgBox "PRODUTO NÃO CADASTRADO!", vbCritical, "Alerta"
        vPedirPeso = False
        Exit Sub
Else
    If Left(txtCodBarra.Text, 1) = "2" And Len(txtCodBarra.Text) = 13 Then
        vPedirPeso = False
    Else
        vPedirPeso = Abs(CBool(r("PEDIRPESO")))
    End If
End If

'If Grid.Rows >= 2 Then
'    If Not r.EOF Then
'        vPedirPeso = Abs(CBool(r("PEDIRPESO")))
'    Else
'        vPedirPeso = False
'    End If
'End If

'setar as variaveis com dados do registro
'Dim var_Venda As Currency      'desativei no dia 12/04/2024
'Dim var_Quant As Double        'desativei no dia 12/04/2024
'Dim var_Data As Date           'desativei no dia 12/04/2024
'Dim vQtde As Double            'desativei no dia 12/04/2024
'Dim var_Custo As Currency      'desativei no dia 12/04/2024

If txtCodProduto.Text = "1" Then
    If txtValorProdAvulso.Text = "" Then
        var_Venda = "0"
    Else
        var_Venda = txtValorProdAvulso.Text
    End If
Else
    var_Venda = ValidateNull(r("venda"))
End If

If txtCodProduto.Text = "1" Then
    If txtValorProdAvulso.Text = "" Then
        var_Custo = "0"
    Else
        var_Custo = txtValorProdAvulso.Text
    End If
Else
    var_Custo = ValidateNull(r("vcusto"))
End If

var_Quant = r("vQuant")
var_Data = Date
varUnidMed = r("vUnid")
vQtde = ValidateNull(r("vQuant"))

'======INICIO DA INSERÇÃO=================
'Dim lNovoCod As Long
Dim itemVenda As Long
Dim varIndiceItem As Long

'Verifica se o produto está no grid
itemVenda = existeVenda(txtCodProduto.Text)

'ver se a quantidade que está sendo vendida passo do limite do estoque
If bEstNeg = False Then
    If vQtde <= 0 Then
        If r("vCodProduto") <> 1 Then
            ShowMsg "A quantidade em estoque é insuficiente.", vbExclamation
            LimparObjetos_Produto
            Exit Sub
        End If
    ElseIf vQtde > 0 Then
        If itemVenda <> -1 Then
            Dim varQuantGrid As Double
            Dim varQuantSobrando As Double
            
            'verificar quantidade no grid
            For i = 1 To Grid.Rows - 1
                If Grid.TextMatrix(i, 2) = txtCodProduto.Text Then
                    varQuantGrid = Grid.TextMatrix(i, 5)
                    'Exit Function
                End If
            Next
        End If
            
        If varQuantGrid >= vQtde Then
            MsgBox "Quantidade não disponivel!", vbInformation, "Aviso do Sistema"
            LimparObjetos_Produto
            txtCodBarra.SetFocus
            Exit Sub
        End If
    End If
End If

'pegar o peso
Dim varQuantMedidas As Double
If varUnidMed = "KG" Then
    If Left(txtCodBarra.Text, 1) = "2" And Len(txtCodBarra.Text) = 13 Then
        If varTipoEtiqueta = "5" Then
            varQuantMedidas = Mid(txtCodBarra, 8, 5) / 1000
        ElseIf varTipoEtiqueta = "4" Then
            varQuantMedidas = Mid(txtCodBarra, 8, 5) / 1000
        ElseIf varTipoEtiqueta = "7" Then
            varQuantMedidas = Mid(txtCodBarra, 8, 5) / 1000
        ElseIf varTipoEtiqueta = "2" Then
            varQuantMedidas = Mid(txtCodBarra, 8, 5) / 1000
        End If
    Else
        varQuantMedidas = 1
    End If
Else
    varQuantMedidas = 1
End If

'adicionar registos
If txtCodProduto.Text = "1" Then
       lNovoCod = AutoNumeracao_Itens
       varIndiceItem = AutoNumeracao_Indice
       sSQL = "INSERT INTO pedidos_itens (codigo, cod_pedido, cod_produto, preco, custo, quantidade, data, tipo_venda, item, cancelado, desconto, subtotal, total) VALUES (" & _
          lNovoCod & ", " & txtCodPedido.Text & ", " & txtCodProduto.Text & ", " & Replace(CCur(var_Venda), ",", ".") & ", " & Replace(CCur(var_Custo), ",", ".") & ", " & Replace(varQuantMedidas, ",", ".") & ", '" & Format$(var_Data, "yyyy-dd-MM") & "', 'VENDA', " & varIndiceItem & ", 0, 0, ((" & Replace(CCur(var_Venda), ",", ".") & ")*(" & Replace(varQuantMedidas, ",", ".") & ")), ((" & Replace(CCur(var_Venda), ",", ".") & ")*(" & Replace(varQuantMedidas, ",", ".") & ")));"
Else
    If varUnidMed = "KG" Then
           lNovoCod = AutoNumeracao_Itens
           varIndiceItem = AutoNumeracao_Indice
           sSQL = "INSERT INTO pedidos_itens (codigo, cod_pedido, cod_produto, preco, custo, quantidade, data, tipo_venda, item, cancelado, desconto, subtotal, total) VALUES (" & _
              lNovoCod & ", " & txtCodPedido.Text & ", " & txtCodProduto.Text & ", " & Replace(CCur(var_Venda), ",", ".") & ", " & Replace(CCur(var_Custo), ",", ".") & ", " & Replace(varQuantMedidas, ",", ".") & ", '" & Format$(var_Data, "yyyy-dd-MM") & "', 'VENDA', " & varIndiceItem & ", 0, 0, ((" & Replace(CCur(var_Venda), ",", ".") & ")*(" & Replace(varQuantMedidas, ",", ".") & ")), ((" & Replace(CCur(var_Venda), ",", ".") & ")*(" & Replace(varQuantMedidas, ",", ".") & ")));"
    Else
        If itemVenda = -1 Then
           lNovoCod = AutoNumeracao_Itens
           varIndiceItem = AutoNumeracao_Indice
           sSQL = "INSERT INTO pedidos_itens (codigo, cod_pedido, cod_produto, preco, custo, quantidade, data, tipo_venda, item, cancelado, desconto, subtotal, total) VALUES (" & _
              lNovoCod & ", " & txtCodPedido.Text & ", " & txtCodProduto.Text & ", " & Replace(CCur(var_Venda), ",", ".") & ", " & Replace(CCur(var_Custo), ",", ".") & ", " & Replace(varQuantMedidas, ",", ".") & ", '" & Format$(var_Data, "yyyy-dd-MM") & "', 'VENDA', " & varIndiceItem & ", 0, 0, ((" & Replace(CCur(var_Venda), ",", ".") & ")*(" & Replace(varQuantMedidas, ",", ".") & ")), ((" & Replace(CCur(var_Venda), ",", ".") & ")*(" & Replace(varQuantMedidas, ",", ".") & ")));"
        Else
           sSQL = "UPDATE pedidos_itens SET " & _
              "quantidade = quantidade + " & Replace(varQuantMedidas, ",", ".") & ", SUBTOTAL =(" & Replace(CCur(var_Venda), ",", ".") & ")*(quantidade + " & Replace(varQuantMedidas, ",", ".") & "), TOTAL =(" & Replace(CCur(var_Venda), ",", ".") & ")*(quantidade + " & Replace(varQuantMedidas, ",", ".") & ")  WHERE (codigo = " & Grid.TextMatrix(itemVenda, 1) & ");"
        End If
    End If
End If
   
'Debug.Print sSQL
dbData.Execute sSQL
LimparObjetos_Produto

End Sub

Private Function AutoNumeracao_Indice() As Long
Dim sSQL As String
Dim r As ADODB.Recordset
Dim varIndiceItem As Long

varIndiceItem = 1
sSQL = "SELECT ISNULL(MAX(item), 0) as ultimo_item FROM pedidos_itens where COD_PEDIDO = '" & Val(txtCodPedido.Text) & "' ;"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then varIndiceItem = r("ultimo_item") + 1

If r.State <> 0 Then r.Close
Set r = Nothing

AutoNumeracao_Indice = varIndiceItem
End Function

Function existeVenda(ByVal codProduto As Long) As Long
Dim i As Integer

existeVenda = -1

For i = 1 To Grid.Rows - 1
   If Grid.TextMatrix(i, 2) = codProduto Then
      existeVenda = i
      Exit Function
   End If
Next
End Function

Function ArredondaCentavos(Valor As String) As String
   ArredondaCentavos = ((Int(Valor * 20)) + IIf(Int((Valor * 20)) <> (Valor * 20), 1, 0)) / 20
End Function

Private Function AutoNumeracao_Itens() As Long
   Dim sSQL As String
   Dim r As ADODB.Recordset
   'Dim lNovoCod As Long
   
   lNovoCod = 1
   sSQL = "SELECT ISNULL(MAX(codigo), 0) as ultimo_item FROM pedidos_itens;"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then lNovoCod = r("ultimo_item") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   AutoNumeracao_Itens = lNovoCod
End Function

Private Function AutoNumeracao_LOG() As Long
   Dim sSQL As String
   Dim r As ADODB.Recordset
   'Dim lNovoCod As Long
   
   lNovoCod = 1
   sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultimo_log FROM log;"
   Set r = dbData.OpenRecordset(sSQL)
   If Not r.BOF Then lNovoCod = r("ultimo_log") + 1
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   AutoNumeracao_LOG = lNovoCod
End Function

Private Function Autonumeracao_Cashback() As Long
'Dim sSQL As String
'Dim r As ADODB.Recordset
'Dim lNovoCod As Long

sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultimo_Cashback FROM Pedidos_Cashback;"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then lNovoCod = r("ultimo_Cashback") + 1
If r.State <> 0 Then r.Close
Set r = Nothing

Autonumeracao_Cashback = lNovoCod
End Function

Private Function Autonumeracao_Parcelas() As Long
'Dim sSQL As String
'Dim r As ADODB.Recordset
'Dim lNovoCod As Long

sSQL = "SELECT ISNULL(MAX(codigo), 0) AS ultima_parcela FROM parcelas;"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then lNovoCod = r("ultima_parcela") + 1
If r.State <> 0 Then r.Close
Set r = Nothing

Autonumeracao_Parcelas = lNovoCod
End Function
Private Function AutoNumeracao_Pedido() As Long
'Dim sSQL As String
'Dim r As ADODB.Recordset
'Dim lNovoCod As Long

'pegar o código do ultimo pedido
lNovoCod = 1
sSQL = "SELECT ISNULL(MAX(cod_pedido), 0) AS ultimo FROM pedidos;"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then lNovoCod = r("ultimo") + 1
If r.State <> 0 Then r.Close
Set r = Nothing

AutoNumeracao_Pedido = lNovoCod
End Function



Private Sub Calcular_Desconto()
'SABER O VALOR DO DESCONTO=====================================================
'If cboTipoPgto.Text = "À VISTA" Then
'        If vValorDescFixoAV <> "0,00" Then
'           If txtDesc.Text <> vValorDescFixoAV And txtDesc.Text <> "0,00" Then
'                txtDesc.Text = Format(txtDesc, ocMONEY)
'           ElseIf txtDesc.Text <> vValorDescFixoAV And txtDesc.Text = "0,00" Then
'                txtDesc.Text = Format(vValorDescFixoAV, ocMONEY)
'           ElseIf txtDesc.Text = vValorDescFixoAV And txtDesc.Text = "0,00" Then
'                txtDesc.Text = Format(vValorDescFixoAV, ocMONEY)
'           End If
'        Else
'           txtDesc.Text = Format(txtDesc, ocMONEY)
'        End If
'ElseIf cboTipoPgto.Text = "À PRAZO" Then
'        If vValorDescFixoAP <> "0,00" Then
'           If txtDesc.Text <> vValorDescFixoAP And txtDesc.Text <> "0,00" Then
'                txtDesc.Text = Format(txtDesc, ocMONEY)
'           ElseIf txtDesc.Text <> vValorDescFixoAP And txtDesc.Text = "0,00" Then
'                txtDesc.Text = Format(vValorDescFixoAP, ocMONEY)
'           ElseIf txtDesc.Text = vValorDescFixoAP And txtDesc.Text = "0,00" Then
'                txtDesc.Text = Format(vValorDescFixoAP, ocMONEY)
'           End If
'        Else
'           txtDesc.Text = Format(txtDesc, ocMONEY)
'        End If
'ElseIf cboTipoPgto.Text = "ORÇAMENTO" Then
'    If cboFormaPgto.Text <> "3 - CARTÃO - DÉBITO" And cboFormaPgto.Text <> "4 - CARTÃO - CRÉDITO" Then
'        If vValorDescFixoAV <> "0,00" Then
'           txtDesc.Text = Format(vValorDescFixoAV, ocMONEY)
'        Else
'           txtDesc.Text = Format(txtDesc, ocMONEY)
'        End If
'    Else
'        txtDesc.Text = Format(txtDesc, ocMONEY)
'    End If
'End If

'CALCULAR O VALOR DAS PARCELAS
If txtSubtotal.Text = "" Or txtSubtotal.Text = "0,00" Then Exit Sub
If txtDesc.Text = "" Then txtDesc.Text = FormatNumber(0, 2)
If txtAcresc.Text = "" Then txtAcresc.Text = FormatNumber(0, 2)

Dim varValorSubTotalDebito As Currency
Dim varValorSubTotalCredito As Currency
Dim varSubTotalBruto As Currency
Dim varSubTotalLiquido As Currency

varSubTotalBruto = txtSubtotal.Text

'If cboformaPgto.Text = "3 - CARTÃO - DÉBITO" Then
'   If txtDesc.Text <> "0,00" And txtAcresc.Text = "0,00" Then     'com desconto sem acrescimo
'      If optDescRS.Value = True Then
'         varSubTotalBruto = Format(CCur(txtSubTotal.Text) - CCur(txtDesc.Text), ocMONEY)
'      ElseIf optDescPorc.Value = True Then
'         varSubTotalBruto = Format(CCur(txtSubTotal.Text) - ((CCur(txtSubTotal.Text) * CCur(txtDesc.Text)) / 100), ocMONEY)
'      End If
'   ElseIf txtAcresc.Text <> "0,00" And txtDesc.Text = "0,00" Then    'sem desconto com acrescim0
'      If optAscrescRS.Value = True Then
'         varSubTotalBruto = Format(CCur(txtSubTotal.Text) + CCur(txtAcresc.Text), ocMONEY)
'      ElseIf optAscrescPorc.Value = True Then
'         varSubTotalBruto = Format(CCur(txtSubTotal.Text) + ((CCur(txtSubTotal.Text) * CCur(txtAcresc.Text)) / 100), ocMONEY)
'      End If
'   Else
'      varSubTotalBruto = Format(txtSubTotal.Text, ocMONEY)
'   End If

'    Dim varValorTaxaDebito As Currency
    
'    If varConfCartaodebito = "1" Then
'        Set oCfg = sysConfig("ACRESC_DEBITO_VALOR")
'        varValorTaxaDebito = oCfg.Value
'        Set oCfg = Nothing
        
'        varValorSubTotalDebito = Format(CCur(varSubTotalBruto) + ((CCur(varSubTotalBruto) * CCur(varValorTaxaDebito)) / 100), ocMONEY)
'    Else
'        varValorSubTotalDebito = CCur(varSubTotalBruto)
'    End If

'    txtTotalDesc.Text = Format(varValorSubTotalDebito, ocMONEY)
'    txtValorRest.Text = Format(varValorSubTotalDebito, ocMONEY) 'esse
    
'ElseIf cboformaPgto.Text = "4 - CARTÃO - CRÉDITO" Or cboFormaPgtoEntrada.Text = "4 - CARTÃO - CRÉDITO" Then
'   'txtDesc.Text = "0,00"
'   If txtDesc.Text <> "0,00" And txtAcresc.Text = "0,00" Then     'com desconto sem acrescimo
'      If optDescRS.Value = True Then
'         varSubTotalBruto = Format(CCur(txtSubTotal.Text) - CCur(txtDesc.Text), ocMONEY)
'      ElseIf optDescPorc.Value = True Then
'         varSubTotalBruto = Format(CCur(txtSubTotal.Text) - ((CCur(txtSubTotal.Text) * CCur(txtDesc.Text)) / 100), ocMONEY)
'      End If
'   ElseIf txtAcresc.Text <> "0,00" And txtDesc.Text = "0,00" Then    'sem desconto com acrescim0
'      If optAscrescRS.Value = True Then
'         varSubTotalBruto = Format(CCur(txtSubTotal.Text) + CCur(txtAcresc.Text), ocMONEY)
'      ElseIf optAscrescPorc.Value = True Then
'         varSubTotalBruto = Format(CCur(txtSubTotal.Text) + ((CCur(txtSubTotal.Text) * CCur(txtAcresc.Text)) / 100), ocMONEY)
'      End If
'   Else
'      varSubTotalBruto = Format(txtSubTotal.Text, ocMONEY)
'   End If

'    Dim varValorTaxaCredito As Currency
    
'    If varConfCartaoCredito = "1" Then
'        Set oCfg = sysConfig("ACRESC_CREDITO_VALOR")
'        varValorTaxaCredito = oCfg.Value
'        Set oCfg = Nothing
        
'        varValorSubTotalCredito = Format(CCur(varSubTotalBruto) + ((CCur(varSubTotalBruto) * CCur(varValorTaxaCredito)) / 100), ocMONEY)
'    Else
'        varValorSubTotalCredito = CCur(varSubTotalBruto)
'    End If

'    txtTotalDesc.Text = Format(varValorSubTotalCredito, ocMONEY)
'    txtValorRest.Text = Format(varValorSubTotalCredito, ocMONEY) 'esse
    
    
'Else
   If txtDesc.Text <> "0,00" And txtAcresc.Text = "0,00" Then     'com desconto sem acrescimo
      
      If optDescRS.Value = True Then
         txtTotalDesc.Text = FormatNumber(CCur(txtSubtotal.Text) - CCur(txtDesc.Text), 2)
      ElseIf optDescPorc.Value = True Then
         'txtTotalDesc.Text = Format(CCur(txtSubTotal.Text) - ((CCur(txtSubTotal.Text) * CDbl(txtDesc.Text)) / 100), ocMONEY)
         txtTotalDesc.Text = FormatNumber(CCur(txtSubtotal.Text) - Round(((CCur(txtSubtotal.Text) * CCur(txtDesc.Text)) / 100), 2), 2)
      End If
   ElseIf txtAcresc.Text <> "0,00" And txtDesc.Text = "0,00" Then    'sem desconto com acrescim0
      If optAscrescRS.Value = True Then
         txtTotalDesc.Text = FormatNumber(CCur(txtSubtotal.Text) + CCur(txtAcresc.Text), 2)
      ElseIf optAscrescPorc.Value = True Then
         txtTotalDesc.Text = FormatNumber(CCur(txtSubtotal.Text) + ((CCur(txtSubtotal.Text) * CDbl(txtAcresc.Text)) / 100), 2)
      End If
      
      
   ElseIf txtAcresc.Text <> "0,00" And txtDesc.Text <> "0,00" Then    'com desconto com acrescim0
    'ACRESCIMO
    Dim vDesc As Currency
    Dim vAcresc As Currency
    Dim vTotDescAcresc As Currency
    
    
      If optAscrescRS.Value = True And optDescRS.Value = True Then
         vAcresc = CCur(txtSubtotal.Text) + CCur(txtAcresc.Text)
         vDesc = vAcresc - CCur(txtDesc.Text)
         txtTotalDesc.Text = FormatNumber(vDesc, 2)
         
      ElseIf optAscrescPorc.Value = True And optDescPorc.Value = True Then
        vDesc = CCur(txtSubtotal.Text) - Round(((CCur(txtSubtotal.Text) * CCur(txtDesc.Text)) / 100), 2)
        vAcresc = ((CCur(vDesc) * CDbl(txtAcresc.Text)) / 100)
        vTotDescAcresc = vDesc + vAcresc
        txtTotalDesc.Text = FormatNumber(vTotDescAcresc, 2)
      
      ElseIf optAscrescPorc.Value = True And optDescPorc.Value = False Then
        vDesc = CCur(txtSubtotal.Text) - CCur(txtDesc.Text)
        vAcresc = ((CCur(vDesc) * CDbl(txtAcresc.Text)) / 100)
        vTotDescAcresc = vDesc + vAcresc
        txtTotalDesc.Text = FormatNumber(vTotDescAcresc, 2)
      ElseIf optAscrescPorc.Value = False And optDescPorc.Value = True Then
        vDesc = CCur(txtSubtotal.Text) - Round(((CCur(txtSubtotal.Text) * CCur(txtDesc.Text)) / 100), 2)
        vAcresc = CCur(vDesc) + CCur(txtAcresc.Text)
        'vTotDescAcresc = vDesc + vAcresc
        txtTotalDesc.Text = FormatNumber(vAcresc, 2)
      End If
   Else
      txtTotalDesc.Text = FormatNumber(txtSubtotal.Text, 2)
   End If
'End If
   
   Mostrar_ValorRestante

If optDescRS.Value = True Then      'desconto em dinheiro
    If txtDesc.Text = "0,00" Then
        'txtDescItens.Text = FormatNumber(0, 2)
        vDescItensVenda = FormatNumber(0, 2)
    Else
        'converter o desconto em dinheiro em porcentagem
        If txtTotalDesc.Text = "" Then Exit Sub
        If txtSubtotal.Text = "" Then Exit Sub
        
        Dim varValorDescProc As Double
        Dim A As Currency
        Dim B As Currency
        
        B = txtTotalDesc.Text
        A = txtSubtotal.Text
        
        varValorDescProc = ((B - A) / A) * 100
        vDescItensVenda = Abs(FormatNumber(varValorDescProc, 2))
        vDescItensVenda = FormatNumber(vDescItensVenda, 2)
    End If
   
Else
    'txtDescItens.Text = FormatNumber(txtDesc.Text, 2)
    vDescItensVenda = FormatNumber(txtDesc.Text, 2)
End If

'MsgBox vDescItensVenda

Calcular_Troco
End Sub
Private Sub Calcular_Parcelas()
If txtTotalDesc.Text = "0,00" Or txtValorRest.Text = "0,00" Or cboQuantParc.Text = "" Then Exit Sub

Dim var_ValorRest As Currency
Dim QUANT As Integer
Dim RESULTADO As Currency

var_ValorRest = txtValorRest.Text
If cboQuantParc.Text = "0" Then cboQuantParc.Text = "1"
QUANT = cboQuantParc.Text

RESULTADO = CCur(var_ValorRest / QUANT)
txtValorParc = Format(RESULTADO, ocMONEY)
End Sub

Private Sub Calcular_Total()
Dim var_Quant As Double
Dim var_VALOR As Currency, var_Total As Currency

If LTrim(txtQuant) = "," Then txtQuant.Text = "0,"
If txtQuant.Text = "0," Then
    txtQuant.SelStart = Len(txtQuant.Text)
End If

If LTrim(txtValor) = "," Then txtValor.Text = "0,"
If txtValor.Text = "0," Then
    txtValor.SelStart = Len(txtValor.Text)
End If

If txtQuant.Text = "" Then var_Quant = 1 Else var_Quant = txtQuant.Text
If txtValor.Text = "" Then var_VALOR = 0 Else var_VALOR = txtValor.Text

var_Total = var_VALOR * var_Quant
txtTotal.Text = Format(var_Total, ocMONEY)
End Sub

Private Sub Calcular_Troco()
Dim VAR_GERAL As Currency, VAR_RECEBIDO As Currency, var_Troco As Currency

If txtTotalDesc.Text = "" Or txtRecebido.Text = "" Then Exit Sub

If txtRecebido.Text = "0,00" Or txtRecebido.Text = "" Then
   txtTroco.Text = Format(0, ocMONEY)
   txtRecebido.Text = Format(0, ocMONEY)
Else
   VAR_GERAL = txtTotalDesc.Text
   VAR_RECEBIDO = txtRecebido.Text
    'If VAR_RECEBIDO > VAR_GERAL Then
        var_Troco = VAR_RECEBIDO - VAR_GERAL
    'Else
    '    var_Troco = 0
    'End If
   txtTroco.Text = Format(var_Troco, ocMONEY)
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

Private Sub FormatarGrid_Produtos(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid
      .Clear
      .Cols = 7
      .Rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 0
      .ColWidth(3) = 6050
      .ColWidth(4) = 900
      .ColWidth(5) = 750
      .ColWidth(6) = 900
      
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      .TextMatrix(0, 1) = "COD"
      .TextMatrix(0, 2) = "COD_PRODUTO"
      .TextMatrix(0, 3) = "DESCRIMINAÇÃO"
      .TextMatrix(0, 4) = "PREÇO"
      .TextMatrix(0, 5) = "QTDE"
      .TextMatrix(0, 6) = "TOTAL"
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = rTabela("codigo")
            .TextMatrix(.Rows - 1, 2) = rTabela("cod_produto")
            
            If tipoEmpresa = 4 Then
               .TextMatrix(.Rows - 1, 3) = rTabela("descricao") & " /  " & rTabela("var_tam") & " / " & rTabela("var_fab")
            Else
               .TextMatrix(.Rows - 1, 3) = rTabela("descricao") & " /  " & rTabela("var_fab") & " / " & rTabela("var_ref")
            End If
            
            .TextMatrix(.Rows - 1, 4) = Format(rTabela("preco"), ocMONEY)
            .TextMatrix(.Rows - 1, 5) = rTabela("quantidade")
            .TextMatrix(.Rows - 1, 6) = Format(rTabela("total"), ocMONEY)
            
            rTabela.MoveNext
            .Rows = .Rows + 1
            i = i + 1
         Loop
      End If
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
   
   txtTotalGeral.Text = Format(SomaGrid(Grid, 6), ocMONEY)
End Sub
Private Sub Imprimir_CupomRecibo()
   'On Error GoTo Tratar_Erro
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim rP As ADODB.Recordset
   Dim rI As ADODB.Recordset
   Dim rF As ADODB.Recordset
   
   Dim i As Integer
   Dim f As Integer
   
   If txtCodPedido.Text = "" Then Exit Sub
   
   sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
   Set r = dbData.OpenRecordset(sSQL)
   
   'consultar funcionario do pedido
   Set rP = dbData.OpenRecordset("SELECT cod_funcionario, TIPO_PAGAMENTO, PAGAMENTO FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ");")
   Set rF = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & rP("cod_funcionario") & ");")
   
   'Recupera um número de arquivo disponível
   f = FreeFile()
   
   'pegar o nome da impressora no ini
   'Dim oIni As Ini
   'Dim var_Impressora As String
   
   'Set oIni = New Ini
   'oIni.Arquivo = appPathApp & "config.ini"
   'var_Impressora = oIni.LerTexto("IMPRESSORA_CUPOM", "impressora")
   'Set oIni = Nothing
   
   Dim Prt As Printer
   Dim oldPrinter As String
   
   'Armazena o nome da impressora atual
   oldPrinter = Printer.DeviceName
   
   ' Find and use the printer just selected in the ListBox
   For Each Prt In Printers
      If Prt.DeviceName = var_ImpTermica Then
         Set Printer = Prt
         Exit For
      End If
   Next
   
   'Open "LPT1" For Output As #1
   'Open "\\balcao04\TERMICA" For Output As #f
      

   'Open "LPT1" For Output As #1
   'Open "\\CAIXAO1\termica" For Output As #1
   
      With Printer
         .ScaleMode = vbPixels
         .PaintPicture imLogoCupom.Picture, 100, 0, 372, 150
         
         For i = 1 To 6
            Printer.Print " "
         Next
         
         .ScaleMode = vbCentimeters
         .FontName = "courier new"
         '.PrintQuality = vbPRPQHigh
         
         Fonte 8, False, False
         Printer.Print String(40, "-")
         Fonte 10, True, False
         Printer.Print Tab((35 - Len(r("fantasia"))) / 2); r("fantasia")   'Esse /2 é p/ centralizar
         Fonte 10, False, False
         Printer.Print Tab((35 - Len(r("razao"))) / 2); r("razao")
         Fonte 8, False, False
         Printer.Print r("endereco") & ", " & r("cidade") & "-" & r("estado")
         Printer.Print "FONE: "; r("telefone")                                        '& " - (89) 9986-3739"
         Fonte 8, False, False
         Printer.Print "CNPJ:"; r("cnpj") & "  IE:" & r("ie")
         Printer.Print " "
         
         Fonte 10, True, False
         Printer.Print Tab(10); "CUPOM DE VENDA"
         
         Fonte 8, False, False
         Printer.Print Tab(2); Format(Date, "dd/mm/yy"); " "; Format(Time, "hh:mm"); " "; "CÓD:"; Format(txtCodPedido.Text, "000000"); " "; rF("nome")

         Fonte 8, False, False
         Printer.Print Tab(2); "Tipo de Pgto:"; rP("TIPO_PAGAMENTO"); "  "; "Forma:"; rP("PAGAMENTO")

         Fonte 8, False, False
         Printer.Print String(40, "-")
         Printer.Print Tab(0); "DESCRIÇÃO";
         Printer.Print Tab(20); "PREÇO";
         Printer.Print Tab(26); "QTDE";
         Printer.Print Tab(35); "TOTAL"
         Printer.Print String(40, "-")
         
         sSQL = "SELECT pedidos_itens.codigo, pedidos_itens.cod_pedido, pedidos_itens.preco, pedidos_itens.quantidade, (pedidos_itens.preco * pedidos_itens.quantidade) as total, produtos.descricao " & _
            "FROM pedidos_itens INNER JOIN produtos ON pedidos_itens.cod_produto = produtos.codigo " & _
            "WHERE (pedidos_itens.cod_pedido = " & txtCodPedido.Text & ") ORDER BY pedidos_itens.codigo DESC;"
         Set rI = dbData.OpenRecordset(sSQL)
         
         Do While Not rI.EOF
            '---------------imprime os dados da tabela----------------------------
            Printer.Print Tab(0); rI("descricao");
            Printer.Print Tab(19); Format$(Format$(rI("preco"), "0.00"), "@@@@@@@");
            Printer.Print Tab(26); Format$(Format$(rI("quantidade"), "0.000"), "@@@@@@@");
            Printer.Print Tab(33); Format$(Format$(rI("total"), "0.00"), "@@@@@@@")
            
            rI.MoveNext                 'vai para o proximo registro
         Loop
         
         Printer.Print String(40, "-")
         
         If frmVendaFechamento.Visible = True Then
            'sub-total
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "SubTotal: ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtSubtotal.Text, "0.00"), "@@@@@@@@")
            
            'desconto
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "Desc.: ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtDesc.Text, "0.00"), "@@@@@@@@")
            
            'total
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "Total: ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtTotalDesc.Text, "0.00"), "@@@@@@@@")
            
         ElseIf cboTipoPgto.Text = "À VISTA" Then
            'sub-total
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "SubTotal: ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtSubtotal.Text, "0.00"), "@@@@@@@@")
            
            'desconto
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "Desc.: ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtDesc.Text, "0.00"), "@@@@@@@@")
            
            'total
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "Total: ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtTotalDesc.Text, "0.00"), "@@@@@@@@")
            
            Printer.Print
            
            'Recebido
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "Receb.: ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtRecebido.Text, "0.00"), "@@@@@@@@")
            
            'Troco
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "Troco: ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtTroco.Text, "0.00"), "@@@@@@@@")
         End If
         
         Printer.Print
         
         Fonte 8, False, False
         Printer.Print Tab((40 - Len("ESTE CUPOM NÃO TEM VALOR FISCAL")) / 2); "ESTE CUPOM NÃO TEM VALOR FISCAL"
         Fonte 8, False, False
         Printer.Print Tab((40 - Len("Obrigado pela preferência")) / 2); "Obrigado pela preferência"
         
         For i = 1 To 4
               Printer.Print " "
         Next
         
         Printer.Print Tab((40 - Len("______________________________________")) / 2); "______________________________________"
         Printer.Print Tab((40 - Len(cboCliente.Text)) / 2); cboCliente.Text
         Printer.Print Tab((40 - Len("VENCIMENTO:" & mskInicio.Text)) / 2); mskInicio.Text
         
         'For i = 1 To 10
         'Print #f, ""
         'Next
        
      Close #f
      .EndDoc
      'rsPedidos.Close
      'rsFunc.Close
      'RS.Close
      'BD.Close
   End With
   
Tratar_Erro:
   ' Atribui a impressora inicial
   'For Each Prt In Printers
   '   If Prt.DeviceName = oldPrinter Then
   '      Set Printer = Prt
   '      Exit For
   '   End If
   'Next
   
   If Not r Is Nothing Then If r.State <> 0 Then r.Close
   If Not rP Is Nothing Then If rP.State <> 0 Then rP.Close
   If Not rI Is Nothing Then If rI.State <> 0 Then rI.Close
   If Not rF Is Nothing Then If rF.State <> 0 Then rF.Close
   
   'If Err.Number = 52 Then
    '  ShowMsg "Impressora não esta pronta ou está com problemas, Verifique !!!", vbInformation
    '  Printer.KillDoc
    '  Exit Sub
   'End If
End Sub

Private Sub Imprimir_CupomGuilhotina2()
   'On Error GoTo Tratar_Erro
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim rP As ADODB.Recordset
   Dim rI As ADODB.Recordset
   Dim rF As ADODB.Recordset
   
   Dim i As Integer
   Dim f As Integer
   
   If txtCodPedido.Text = "" Then Exit Sub
   
   sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
   Set r = dbData.OpenRecordset(sSQL)
   
   'consultar funcionario do pedido
   Set rP = dbData.OpenRecordset("SELECT cod_funcionario, TIPO_PAGAMENTO, PAGAMENTO FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ");")
   Set rF = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & rP("cod_funcionario") & ");")
   
   'Recupera um número de arquivo disponível
   f = FreeFile()
   
   'pegar o nome da impressora no ini
   'Dim oIni As Ini 'desativei aqui 09/11/22
   'Dim var_ImpTermica As String
   
   Set oIni = New Ini
   oIni.Arquivo = appPathApp & "config.ini"
   var_ImpTermica = oIni.LerTexto("IMPRESSORA_TERMICA", "impressora")
   Set oIni = Nothing
   
   Dim Prt As Printer
   Dim oldPrinter As String
   
   'Armazena o nome da impressora atual
   oldPrinter = Printer.DeviceName
   
   ' Find and use the printer just selected in the ListBox
   For Each Prt In Printers
      If Prt.DeviceName = var_ImpTermica Then
         Set Printer = Prt
         Exit For
      End If
   Next
   
   'Open "LPT1" For Output As #1
   'Open "\\balcao04\TERMICA" For Output As #f
      

   'Open "LPT1" For Output As #1
   'Open "\\CAIXAO1\termica" For Output As #1
   
      With Printer
         .ScaleMode = vbPixels
         .PaintPicture imLogoCupom.Picture, 100, 0, 372, 150
         
         For i = 1 To 6
            Printer.Print " "
         Next
         
         .ScaleMode = vbCentimeters
         .FontName = "courier new"
         '.PrintQuality = vbPRPQHigh
         
         Fonte 8, False, False
         Printer.Print String(40, "-")
         Fonte 10, True, False
         Printer.Print Tab((35 - Len(r("fantasia"))) / 2); r("fantasia")   'Esse /2 é p/ centralizar
         Fonte 10, False, False
         'Printer.Print Tab((35 - Len(r("razao"))) / 2); r("razao")
         Printer.Print " "
         Fonte 8, False, False
         Printer.Print r("endereco") & ", " & r("cidade") & "-" & r("estado")
         Printer.Print "FONE: "; r("telefone")                                        '& " - (89) 9986-3739"
         Fonte 8, False, False
         Printer.Print "CNPJ:"; r("cnpj") & "  IE:" & r("ie")
         Printer.Print " "
         
         Fonte 10, True, False
         If cboTipoPgto.Text = "ORÇAMENTO" Then
            Printer.Print Tab(10); "O R Ç A M E N T O"
        Else
            Printer.Print Tab(10); "CUPOM DE VENDA"
        End If
         
         Fonte 8, False, False
         Printer.Print Tab(2); Format(Date, "dd/mm/yy"); " "; Format(Time, "hh:mm"); " "; "CÓD:"; Format(txtCodPedido.Text, "000000"); " "; rF("nome")

         Fonte 8, False, False
         Printer.Print Tab(2); "Tipo de Pgto:"; rP("TIPO_PAGAMENTO"); "  "; "Forma:"; rP("PAGAMENTO")

         Fonte 8, False, False
         Printer.Print String(40, "-")
         Printer.Print Tab(0); "DESCRIÇÃO";
         Printer.Print Tab(20); "PREÇO";
         Printer.Print Tab(26); "QTDE";
         Printer.Print Tab(35); "TOTAL"
         Printer.Print String(40, "-")
         
         sSQL = "SELECT pedidos_itens.codigo, pedidos_itens.cod_pedido, pedidos_itens.preco, pedidos_itens.quantidade, (pedidos_itens.preco * pedidos_itens.quantidade) as total, produtos.descricao " & _
            "FROM pedidos_itens INNER JOIN produtos ON pedidos_itens.cod_produto = produtos.codigo " & _
            "WHERE (pedidos_itens.cod_pedido = " & txtCodPedido.Text & ") ORDER BY pedidos_itens.codigo DESC;"
         Set rI = dbData.OpenRecordset(sSQL)
         
         Do While Not rI.EOF
            '---------------imprime os dados da tabela----------------------------
            Printer.Print Tab(0); rI("descricao");
            Printer.Print Tab(19); Format$(Format$(rI("preco"), "0.00"), "@@@@@@@");
            Printer.Print Tab(26); Format$(Format$(rI("quantidade"), "0.000"), "@@@@@@@");
            Printer.Print Tab(33); Format$(Format$(rI("total"), "0.00"), "@@@@@@@")
            
            rI.MoveNext                 'vai para o proximo registro
         Loop
         
         Printer.Print String(40, "-")
         
         If cboTipoPgto.Text = "À PRAZO" Then
            'sub-total
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "SubTotal: ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtSubtotal.Text, "0.00"), "@@@@@@@@")
            
            'desconto
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "Desc.(%): ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtDesc.Text, "0.00"), "@@@@@@@@")
            
            'total
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "Total: ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtTotalDesc.Text, "0.00"), "@@@@@@@@")
            
         ElseIf cboTipoPgto.Text = "À VISTA" Then
            'sub-total
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "SubTotal: ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtSubtotal.Text, "0.00"), "@@@@@@@@")
            
            'desconto
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "Desc.(%): ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtDesc.Text, "0.00"), "@@@@@@@@")
            
            'total
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "Total: ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtTotalDesc.Text, "0.00"), "@@@@@@@@")
            
            Printer.Print
            
            'Recebido
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "Receb.: ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtRecebido.Text, "0.00"), "@@@@@@@@")
            
            'Troco
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "Troco: ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtTroco.Text, "0.00"), "@@@@@@@@")
         ElseIf cboTipoPgto.Text = "ORÇAMENTO" Then
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "SubTotal: ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtSubtotal.Text, "0.00"), "@@@@@@@@")
            
            'desconto
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "Desc.(%): ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtDesc.Text, "0.00"), "@@@@@@@@")
            
            'total
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "Total: ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtTotalDesc.Text, "0.00"), "@@@@@@@@")
            
            Printer.Print
            
            'Recebido
            'Fonte 8, False, False
            'Printer.Print Tab(0); Tab(20); "Receb.: ";
            
            'Fonte 10, True, False
            'Printer.Print Tab(25); Format$(Format$(txtRecebido.Text, "0.00"), "@@@@@@@@")
            
            'Troco
            'Fonte 8, False, False
            'Printer.Print Tab(0); Tab(20); "Troco: ";
            
            'Fonte 10, True, False
            'Printer.Print Tab(25); Format$(Format$(txtTroco.Text, "0.00"), "@@@@@@@@")
         End If
         
         Printer.Print
         
         Fonte 8, False, False
         Printer.Print Tab((40 - Len("ESTE CUPOM NÃO TEM VALOR FISCAL")) / 2); "ESTE CUPOM NÃO TEM VALOR FISCAL"
         Fonte 8, False, False
         Printer.Print Tab((40 - Len("Obrigado pela preferência")) / 2); "Obrigado pela preferência"
         
         For i = 1 To 4
               Printer.Print " "
         Next
         
        If cboTipoPgto.Text <> "ORÇAMENTO" Then
            Printer.Print Tab((40 - Len("______________________________________")) / 2); "______________________________________"
            Printer.Print Tab((40 - Len(cboCliente.Text)) / 2); cboCliente.Text
            Printer.Print Tab((40 - Len("VENCIMENTO:" & mskInicio.Text)) / 2); mskInicio.Text
        End If
         'For i = 1 To 10
         'Print #f, ""
         'Next
        
      Close #f
      .EndDoc
      'rsPedidos.Close
      'rsFunc.Close
      'RS.Close
      'BD.Close
   End With
   
Tratar_Erro:
   ' Atribui a impressora inicial
   'For Each Prt In Printers
   '   If Prt.DeviceName = oldPrinter Then
   '      Set Printer = Prt
   '      Exit For
   '   End If
   'Next
   
   If Not r Is Nothing Then If r.State <> 0 Then r.Close
   If Not rP Is Nothing Then If rP.State <> 0 Then rP.Close
   If Not rI Is Nothing Then If rI.State <> 0 Then rI.Close
   If Not rF Is Nothing Then If rF.State <> 0 Then rF.Close
   
   'If Err.Number = 52 Then
    '  ShowMsg "Impressora não esta pronta ou está com problemas, Verifique !!!", vbInformation
    '  Printer.KillDoc
    '  Exit Sub
   'End If
End Sub
Private Sub Imprimir_CupomGuilhotina()
   'On Error GoTo Tratar_Erro
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim rP As ADODB.Recordset
   Dim rPR As ADODB.Recordset
   Dim rI As ADODB.Recordset
   Dim rF As ADODB.Recordset
   Dim rParc As ADODB.Recordset
   
   Dim i As Integer
   Dim f As Integer
   
   If txtCodPedido.Text = "" Then Exit Sub
   
   sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
   Set r = dbData.OpenRecordset(sSQL)
   
   'consultar funcionario do pedido
   Set rP = dbData.OpenRecordset("SELECT cod_funcionario, TIPO_PAGAMENTO, PAGAMENTO FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ");")
   Set rPR = dbData.OpenRecordset("SELECT * FROM Pedidos_Recebedor WHERE (cod_pedido = " & txtCodPedido.Text & ");")
   Set rF = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & rP("cod_funcionario") & ");")
   Set rParc = dbData.OpenRecordset("SELECT COD_PEDIDO, NUMERO, PAGAMENTO, DATA, VALOR_FINAL, (CASE WHEN FORMA_PGTO = 'CARTAO' THEN (CASE WHEN TIPO_CARTAO = 'D' THEN 'CARTÃO DÉBITO' ELSE 'CARTÃO CRÉDITO' END) ELSE isnull(FORMA_PGTO, '') END) AS varFormaPgto FROM parcelas WHERE (cod_pedido = " & txtCodPedido.Text & ") order by NUMERO;")
   
   'Recupera um número de arquivo disponível
   f = FreeFile()
   
   'pegar o nome da impressora no ini
   'Dim oIni As Ini 'desativei aqui 09/11/22
   ''Dim var_ImpTermica As String
   
   Set oIni = New Ini
   oIni.Arquivo = appPathApp & "config.ini"
   var_ImpTermica = oIni.LerTexto("IMPRESSORA_TERMICA", "impressora")
   Set oIni = Nothing
   
   Dim Prt As Printer
   Dim oldPrinter As String
   
   'Armazena o nome da impressora atual
   oldPrinter = Printer.DeviceName
   
   ' Find and use the printer just selected in the ListBox
   For Each Prt In Printers
      If Prt.DeviceName = var_ImpTermica Then
         Set Printer = Prt
         Exit For
      End If
   Next
   
      With Printer
         .ScaleMode = vbPixels
         .PaintPicture imLogoCupom.Picture, 100, 0, 372, 150
         
         For i = 1 To 6
            Printer.Print " "
         Next
         
         .ScaleMode = vbCentimeters
         .FontName = "courier new"
         '.PrintQuality = vbPRPQHigh
         
         Fonte 8, False, False
         Printer.Print String(40, "-")
         Fonte 10, True, False
         Printer.Print Tab((35 - Len(r("fantasia"))) / 2); r("fantasia")   'Esse /2 é p/ centralizar
         Fonte 10, False, False
         'Printer.Print Tab((35 - Len(r("razao"))) / 2); r("razao")
         Printer.Print " "
         Fonte 8, False, False
         Printer.Print r("endereco") & ", " & r("cidade") & "-" & r("estado")
         Printer.Print "FONE: "; r("telefone") & " - " & r("celular") & ""                                        '& " - (89) 9986-3739"
         Fonte 8, False, False
         Printer.Print "CNPJ:"; r("cnpj") & "  IE:" & r("ie")
         Printer.Print " "
         
         Fonte 10, True, False
         If cboTipoPgto.Text = "ORÇAMENTO" Then
            Printer.Print Tab(10); "O R Ç A M E N T O"
         ElseIf cboTipoPgto.Text = "CONSIGNADO" Then
            Printer.Print Tab(10); "C O N S I G N A D O"
         Else
            Printer.Print Tab(10); "CUPOM DE VENDA"
         End If
         
         Fonte 8, False, False
         Printer.Print Tab(2); Format(Date, "dd/mm/yy"); " "; Format(Time, "hh:mm"); " "; "CÓD:"; Format(txtCodPedido.Text, "000000"); " "; rF("nome")

         Fonte 8, False, False
         Printer.Print Tab(2); "Tipo de Pgto:"; rP("TIPO_PAGAMENTO"); "  "; "Forma:"; rP("PAGAMENTO")

         Fonte 8, False, False
         Printer.Print String(40, "-")
         Printer.Print Tab(0); "DESCRIÇÃO";
         Printer.Print Tab(20); "PREÇO";
         Printer.Print Tab(26); "QTDE";
         Printer.Print Tab(35); "TOTAL"
         Printer.Print String(40, "-")
         
         sSQL = "SELECT pedidos_itens.codigo, pedidos_itens.cod_pedido, pedidos_itens.preco, pedidos_itens.quantidade, (pedidos_itens.preco * pedidos_itens.quantidade) as total, produtos.descricao " & _
            "FROM pedidos_itens INNER JOIN produtos ON pedidos_itens.cod_produto = produtos.codigo " & _
            "WHERE (pedidos_itens.cod_pedido = " & txtCodPedido.Text & ") ORDER BY pedidos_itens.codigo DESC;"
         Set rI = dbData.OpenRecordset(sSQL)
         
         Do While Not rI.EOF
            '---------------imprime os dados da tabela----------------------------
            Printer.Print Tab(0); rI("descricao");
            Printer.Print Tab(19); Format$(Format$(rI("preco"), "0.00"), "@@@@@@@");
            Printer.Print Tab(26); Format$(Format$(rI("quantidade"), "0.000"), "@@@@@@@");
            Printer.Print Tab(33); Format$(Format$(rI("total"), "0.00"), "@@@@@@@")
            
            rI.MoveNext                 'vai para o proximo registro
         Loop
         
         Printer.Print String(40, "-")

            Fonte 8, False, False
            Printer.Print Tab(0); "*** PARCELAS ***";
            Printer.Print Tab(0); "No.";
            If cboTipoPgto.Text = "À PRAZO" Then
                Printer.Print Tab(5); "VENC.";
            Else
                Printer.Print Tab(5); "PGTO";
            End If
            Printer.Print Tab(17); "VALOR";
            Printer.Print Tab(25); "FORMA"
         
             Do While Not rParc.EOF
                Printer.Print Tab(0); rParc("NUMERO");
                If cboTipoPgto.Text = "À PRAZO" Then
                    Printer.Print Tab(5); Format$(Format$(rParc("DATA"), "dd/mm/yy"), "@@@@@@@");
                Else
                    Printer.Print Tab(5); Format$(Format$(rParc("PAGAMENTO"), "dd/mm/yy"), "@@@@@@@");
                End If
                Printer.Print Tab(15); Format$(Format$(rParc("VALOR_FINAL"), "0.00"), "@@@@@@@");
                Printer.Print Tab(25); rParc("varFormaPgto")
                
                rParc.MoveNext                 'vai para o proximo registro
            Loop
         
            For i = 1 To 1
               Printer.Print " "
            Next
         
         If cboTipoPgto.Text = "À PRAZO" Then
            'sub-total
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "SubTotal: ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtSubtotal.Text, "0.00"), "@@@@@@@@")
            
            'desconto
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "Desc.(%): ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtDesc.Text, "0.00"), "@@@@@@@@")
            
            'total
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "Total: ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtTotalDesc.Text, "0.00"), "@@@@@@@@")
            
         ElseIf cboTipoPgto.Text = "À VISTA" Then
         
            'sub-total
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "SubTotal: ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtSubtotal.Text, "0.00"), "@@@@@@@@")
            
            'desconto
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "Desc.(%): ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtDesc.Text, "0.00"), "@@@@@@@@")
            
            'total
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "Total: ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtTotalDesc.Text, "0.00"), "@@@@@@@@")
            
            Printer.Print
            
            'Recebido
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "Receb.: ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtRecebido.Text, "0.00"), "@@@@@@@@")
            
            'Troco
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "Troco: ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtTroco.Text, "0.00"), "@@@@@@@@")
         ElseIf cboTipoPgto.Text = "ORÇAMENTO" Then
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "SubTotal: ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtSubtotal.Text, "0.00"), "@@@@@@@@")
            
            'desconto
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "Desc.(%): ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtDesc.Text, "0.00"), "@@@@@@@@")
            
            'total
            Fonte 8, False, False
            Printer.Print Tab(0); Tab(20); "Total: ";
            
            Fonte 10, True, False
            Printer.Print Tab(25); Format$(Format$(txtTotalDesc.Text, "0.00"), "@@@@@@@@")
            
            Printer.Print
            
            'Recebido
            'Fonte 8, False, False
            'Printer.Print Tab(0); Tab(20); "Receb.: ";
            
            'Fonte 10, True, False
            'Printer.Print Tab(25); Format$(Format$(txtRecebido.Text, "0.00"), "@@@@@@@@")
            
            'Troco
            'Fonte 8, False, False
            'Printer.Print Tab(0); Tab(20); "Troco: ";
            
            'Fonte 10, True, False
            'Printer.Print Tab(25); Format$(Format$(txtTroco.Text, "0.00"), "@@@@@@@@")
         End If
         
         Printer.Print
         
         Fonte 8, False, False
         If cboTipoPgto.Text <> "CONSIGNADO" Then
            Printer.Print Tab((40 - Len("ESTE CUPOM NÃO TEM VALOR FISCAL")) / 2); "ESTE CUPOM NÃO TEM VALOR FISCAL"
         Else
            Printer.Print Tab((40 - Len("[AVISO]: O produto deverá ser entregue")) / 2); "[AVISO]: O produto deverá ser entregue"
         End If
         
         Fonte 8, False, False
         
         If cboTipoPgto.Text <> "CONSIGNADO" Then
            Printer.Print Tab((40 - Len("Obrigado pela preferência")) / 2); "Obrigado pela preferência"
         Else
            Printer.Print Tab((40 - Len("no máximo em 24 horas.")) / 2); "no máximo em 24 horas."
         End If
         
         For i = 1 To 4
               Printer.Print " "
         Next
         
        If cboTipoPgto.Text <> "ORÇAMENTO" Then
            Printer.Print Tab((40 - Len("______________________________________")) / 2); "______________________________________"
            Printer.Print Tab((40 - Len(cboCliente.Text)) / 2); cboCliente.Text
            If cboTipoPgto.Text <> "CONSIGNADO" Then
                If cboTipoPgto.Text = "À PRAZO" Then
                    Printer.Print Tab((40 - Len("VENCIMENTO:" & mskInicio.Text)) / 2); "Pagar em:" & Format(mskInicio.Text, "dd/mm/yy")
                Else
                    Printer.Print Tab((40 - Len("VENCIMENTO:" & mskInicio.Text)) / 2); "Pago em:" & Format(mskInicio.Text, "dd/mm/yy")
                End If
            Else
                Printer.Print Tab((40 - Len("DATA:" & Date)) / 2); "Recebido em:" & Format(Date, "dd/mm/yy")
            End If
        End If
        
        For i = 1 To 2
            Printer.Print " "
        Next
        
        'DADOS DO RECECEDOR
        If cboTipoPgto.Text <> "ORÇAMENTO" And cboTipoPgto.Text <> "CONSIGNADO" Then
            If vDeclararRecebedor = "SIM" Then
                If Not rPR.EOF Then
                    Printer.Print Tab((40 - Len("______________________________________")) / 2); "______________________________________"
                    Printer.Print Tab((40 - Len(rPR("Recebedor"))) / 2); rPR("Recebedor")
                    Printer.Print Tab((40 - Len("RECEBEDOR")) / 2); "RECEBEDOR"
                    Printer.Print Tab((40 - Len("Recebido em:" & Date)) / 2); "Recebido em:" & Format(Date, "dd/mm/yy")
                End If
            End If
        End If
        

         'For i = 1 To 10
         'Print #f, ""
         'Next
        
      Close #f
      .EndDoc
      'rsPedidos.Close
      'rsFunc.Close
      'RS.Close
      'BD.Close
   End With
   
Tratar_Erro:
   ' Atribui a impressora inicial
   'For Each Prt In Printers
   '   If Prt.DeviceName = oldPrinter Then
   '      Set Printer = Prt
   '      Exit For
   '   End If
   'Next
   
   If Not r Is Nothing Then If r.State <> 0 Then r.Close
   If Not rP Is Nothing Then If rP.State <> 0 Then rP.Close
   If Not rPR Is Nothing Then If rPR.State <> 0 Then rPR.Close
   If Not rI Is Nothing Then If rI.State <> 0 Then rI.Close
   If Not rF Is Nothing Then If rF.State <> 0 Then rF.Close
   If Not rParc Is Nothing Then If rParc.State <> 0 Then rParc.Close
   
   'If Err.Number = 52 Then
    '  ShowMsg "Impressora não esta pronta ou está com problemas, Verifique !!!", vbInformation
    '  Printer.KillDoc
    '  Exit Sub
   'End If
End Sub
Private Sub Imprimir_CupomSerrilhaPrazo()
   'On Error GoTo TrataErro
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim rP As ADODB.Recordset
   Dim rI As ADODB.Recordset
   Dim rF As ADODB.Recordset
   
   Dim i As Integer
   Dim f As Integer
   
   sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
   Set r = dbData.OpenRecordset(sSQL)
   
   If txtCodPedido.Text = "" Then Exit Sub
   
   'consultar funcionario do pedido
   Set rP = dbData.OpenRecordset("SELECT cod_funcionario FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ");")
   Set rF = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & rP("cod_funcionario") & ");")
   
   f = FreeFile()
   
   'Open "LPT2" For Output As #1
   Open "\\BALCAO01\termica" For Output As #f
      Print #f, Chr$(27) & Chr(15)
      Print #f, Spc(0); "----------------------------------------------------------------"
      Print #f, Tab((60 - Len(r("fantasia"))) / 2); r("fantasia")
      Print #f, Tab((60 - Len(r("razao"))) / 2); r("razao")
      Print #f, Tab((60 - Len(r("endereco") & ", " & r("cidade") & "-" & r("estado"))) / 2); r("endereco") & ", " & r("cidade") & "-" & r("estado")
      Print #f, Tab((60 - Len(r("telefone"))) / 2); r("telefone")
      Print #f, Tab((60 - Len(r("cnpj") & "  IE:" & r("ie"))) / 2); r("cnpj") & "  IE:" & r("ie")
      Print #f, ""
      Print #f, Spc(0); Format(Date, "dd/mm/yy"); Spc(3); Format(Time, "hh:mm"); Spc(4); "No. Cupom:"; Spc(1); Format(txtCodPedido.Text, "000000"); Spc(3); "Usuario:"; Spc(1); rF("nome")
      Print #f, ""
      Print #f, Spc(0); "                       C   U   P   O   M                     "
      Print #f, Spc(0); "----------------------------------------------------------------"
      Print #f, Tab(0); "DESCRICAO"; Tab(40); "PRECO"; Tab(48); "QUANT"; Tab(56); "TOTAL"
      Print #f, Spc(0); "----------------------------------------------------------------"
      
         sSQL = "SELECT pedidos_itens.codigo, pedidos_itens.cod_pedido, pedidos_itens.preco, pedidos_itens.quantidade, (pedidos_itens.preco * pedidos_itens.quantidade) as total, produtos.descricao " & _
            "FROM pedidos_itens INNER JOIN produtos ON pedidos_itens.cod_produto = produtos.codigo " & _
            "WHERE (pedidos_itens.cod_pedido = " & txtCodPedido.Text & ") ORDER BY pedidos_itens.codigo DESC;"
         Set rI = dbData.OpenRecordset(sSQL)
      
      Do While Not rI.EOF
         Print #f, Tab(0); rI("descricao"); Tab(38); Format$(Format$(r("preco"), "0.00"), "@@@@@@@"); Tab(46); Format$(Format$(rI("quantidade"), "0.000"), "@@@@@@@"); Tab(54); Format$(Format$(rI("total"), "0.00"), "@@@@@@@")
         rI.MoveNext
      Loop
      
      Print #f, Spc(0); "----------------------------------------------------------------"
      Print #f, Tab(45); "TOTAL: "; Tab(54); Format$(Format$(txtTotalGeral.Text, "0.00"), "@@@@@@@@")
      Print #f, ""
      Print #f, Tab((60 - Len("ESTE CUPOM NAO TEM VALOR FISCAL")) / 2); "ESTE CUPOM NAO TEM VALOR FISCAL"
      Print #f, Tab((60 - Len("Obrigado pela preferencia")) / 2); "Obrigado pela preferencia"
      Print #f, ""
      Print #f, ""
      Print #f, ""
      Print #f, Tab((60 - Len("_________________________________________________")) / 2); "_________________________________________________"
      Print #f, Tab((60 - Len(cboCliente.Text)) / 2); cboCliente.Text
      Print #f, Tab((60 - Len("VENCIMENTO:" & mskInicio.Text)) / 2); mskInicio.Text
      Print #f, ""
      Print #f, ""
      Print #f, ""
      Print #f, ""
      Print #f, ""
      Print #f, ""
      Print #f, ""
      Print #f, ""
      Print #f, ""
      Print #f, ""
   Close #f
   
   If Not r Is Nothing Then If r.State <> 0 Then r.Close
   If Not rP Is Nothing Then If rP.State <> 0 Then rP.Close
   If Not rI Is Nothing Then If rI.State <> 0 Then rI.Close
   If Not rF Is Nothing Then If rF.State <> 0 Then rF.Close
   
   Exit Sub

'TrataErro:
   'MsgBox Err.Description, vbCritical, "Erro no Sistema, Impressora Inoperante"
End Sub

Private Sub Imprimir_CupomSerrilha()
   'On Error GoTo TrataErro
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim rP As ADODB.Recordset
   Dim rI As ADODB.Recordset
   Dim rF As ADODB.Recordset
   
   Dim i As Integer
   Dim f As Integer
   
   sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
   Set r = dbData.OpenRecordset(sSQL)
   
   'consultar funcionario do pedido
   Set rP = dbData.OpenRecordset("SELECT * FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ");")
   Set rF = dbData.OpenRecordset("SELECT codigo, nome FROM funcionario WHERE (codigo = " & rP("cod_funcionario") & ");")
   
   f = FreeFile()
   
   'Open "LPT1" For Output As #1
   Open "\\BALCAO01\termica" For Output As #f
      Print #f, Chr$(27) & Chr(15)
      Print #f, Spc(0); "----------------------------------------------------------------"
      Print #f, Tab((60 - Len(r("fantasia"))) / 2); r("fantasia")
      Print #f, Tab((60 - Len(r("razao"))) / 2); r("razao")
      Print #f, Tab((60 - Len(r("endereco") & ", " & r("cidade") & "-" & r("estado"))) / 2); r("endereco") & ", " & r("cidade") & "-" & r("estado")
      Print #f, Tab((60 - Len(r("telefone"))) / 2); r("telefone")
      Print #f, Tab((60 - Len(r("cnpj") & "  IE:" & r("ie"))) / 2); r("cnpj") & "  IE:" & r("ie")
      Print #f, ""
      Print #f, Spc(0); Format(Date, "dd/mm/yy"); Spc(3); Format(Time, "hh:mm"); Spc(4); "No. Cupom:"; Spc(1); Format(txtCodPedido.Text, "000000"); Spc(3); "Usuario:"; Spc(1); rF("nome")
      Print #f, ""
      Print #f, Spc(0); "                       C   U   P   O   M                     "
      Print #f, Spc(0); "----------------------------------------------------------------"
      Print #f, Tab(0); "DESCRICAO"; Tab(40); "PRECO"; Tab(48); "QUANT"; Tab(56); "TOTAL"
      Print #f, Spc(0); "----------------------------------------------------------------"

         sSQL = "SELECT pedidos_itens.codigo, pedidos_itens.cod_pedido, pedidos_itens.preco, pedidos_itens.quantidade, (pedidos_itens.preco * pedidos_itens.quantidade) as total, produtos.descricao " & _
            "FROM pedidos_itens INNER JOIN produtos ON pedidos_itens.cod_produto = produtos.codigo " & _
            "WHERE (pedidos_itens.cod_pedido = " & txtCodPedido.Text & ") ORDER BY pedidos_itens.codigo DESC;"
         Set rI = dbData.OpenRecordset(sSQL)
      
      Do While Not rI.EOF
         Print #f, Tab(0); rI("descricao"); Tab(38); Format$(Format$(r("preco"), "0.00"), "@@@@@@@"); Tab(46); Format$(Format$(rI("quantidade"), "0.000"), "@@@@@@@"); Tab(54); Format$(Format$(rI("total"), "0.00"), "@@@@@@@")
         rI.MoveNext
      Loop
      
      Print #f, Spc(0); "----------------------------------------------------------------"
      Print #f, Tab(45); "TOTAL: "; Tab(54); Format$(Format$(txtTotalDesc.Text, "0.00"), "@@@@@@@@")
      Print #f, ""
      Print #f, Tab((60 - Len("ESTE CUPOM NAO TEM VALOR FISCAL")) / 2); "ESTE CUPOM NAO TEM VALOR FISCAL"
      Print #f, Tab((60 - Len("Obrigado pela preferencia")) / 2); "Obrigado pela preferencia"
      Print #f, ""
      Print #f, ""
      Print #f, ""
      Print #f, ""
      Print #f, ""
      Print #f, ""
      Print #f, ""
      Print #f, ""
   Close #f
   
   If Not r Is Nothing Then If r.State <> 0 Then r.Close
   If Not rP Is Nothing Then If rP.State <> 0 Then rP.Close
   If Not rI Is Nothing Then If rI.State <> 0 Then rI.Close
   If Not rF Is Nothing Then If rF.State <> 0 Then rF.Close
   

'TrataErro:
   'MsgBox Err.Description, vbCritical, "Erro no Sistema, Impressora Inoperante"
End Sub

Private Sub Fonte(Tamanho As Byte, Negrito As Boolean, Italico As Boolean) 'Altera a fonte
   Printer.FontSize = Tamanho
   Printer.FontBold = Negrito
   Printer.FontItalic = Italico
End Sub

Private Sub LimparGrid_Pedido()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT * FROM pedidos_itens WHERE 0 = 1;"
Set r = dbData.OpenRecordset(sSQL)
FormatarGrid_Produtos r
If Not r.State <> 0 Then r.Close
Set r = Nothing

lblQuantTipo.Caption = ""
End Sub

Private Sub LimparObjetos_Pedido()
'txtTotalGeral.Text = ""
txtRecebido.Text = ""
txtTroco.Text = ""
lblDesc.Caption = ""
txtValor.Text = ""
txtQuant.Text = ""
txtTotal.Text = ""
cmdFinalizarAvista.Visible = True
cmdFinalizarPrazo.Visible = True
cmdCancelarPedido.Visible = True
cmdRemover.Visible = True
If lblEstornar.Caption <> "ESTORNO" Then cmdOrçamento.Visible = True
cmdAvancado.Visible = True
cmdFechar.Visible = True
cmdFinalizar.Visible = True
cmdCancelar.Visible = True
'cmdPausarCompra.Visible = True
frmVendaFechamento.Enabled = True
lblEstornar.Caption = ""
End Sub

Private Sub LimparObjetos_Prazo()
If lblEstornar.Caption = "ESTORNO" Then
   txtEntrada.Text = Format(0, ocMONEY)
   cboPrazo.Text = "30"
   txtValorParc.Text = Format(0, ocMONEY)
   mskInicio.Mask = ""
   mskInicio.Text = ""
   optDescPorc.Value = True
   'txtDesc.Text = "0,00"
   'txtAcresc.Text = "0,00"
   cboQuantParc.Text = "1"
ElseIf lblEstornar.Caption = "REIMPRESSÃO" Then
   Exit Sub
Else
   If cboTipoPgto.Text <> "À VISTA" Then txtCodCliente.Text = ""
   If cboTipoPgto.Text <> "À VISTA" Then cboCliente.Text = ""
   txtEntrada.Text = Format(0, ocMONEY)
   cboPrazo.Text = "30"
   txtValorParc.Text = Format(0, ocMONEY)
   mskInicio.Mask = ""
   mskInicio.Text = ""
   'optDescPorc.Value = True
   txtDesc.Text = "0,00"
   txtAcresc.Text = "0,00"
   cboQuantParc.Text = "1"
   If varLoginFunc = "2" Then txtCodFuncAP.Text = ""
   If varLoginFunc = "2" Then txtFuncAP.Text = ""
End If
End Sub

Private Sub LimparObjetos_Produto()
'lblDesc.Caption = ""
txtTotal.Text = Format(0, ocMONEY)
txtQuant.Text = ""
txtValor.Text = Format(0, ocMONEY)
txtUnidMed.Text = ""
txtCodProduto.Text = ""
txtCodItem.Text = ""
txtCodBarraPeso.Text = ""
txtCodBarra.Text = ""
txtValorProdAvulso.Text = ""
End Sub

Private Sub Mostrar_Descricao_Produto()
'Verifica_QuantEstoque

Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodBarra.Text = "" Then Exit Sub

sSQL = "SELECT produtos.codigo AS var_codprod, produtos.descricao AS var_desc, produtos.unid_medida AS var_unidmed, " & _
   "produtos_entrada_itens.valor_vv AS var_venda, produtos_entrada_itens.codigo_produto, produtos.cod_barra, produtos.ativo, " & _
   "produtos_entrada_itens.codigo FROM produtos INNER JOIN produtos_entrada_itens ON produtos.codigo = produtos_entrada_itens.codigo_produto " & _
   "WHERE (produtos.cod_barra = '" & txtCodBarra.Text & "') AND (produtos.ativo = 1) ORDER BY produtos_entrada_itens.codigo DESC;"

Set r = dbData.OpenRecordset(sSQL)

If r.BOF Then
   ShowMsg "Produto Inexistente!", vbCritical
   LimparObjetos_Produto
   txtCodBarra.SetFocus
   GoTo SairProc
End If

'rs.MoveLast
lblDesc.Caption = r("var_desc")
txtValor.Text = Format(r("var_venda"), ocMONEY)
txtUnidMed.Text = Format(r("var_unidmed"), ocMONEY)

'Set cCfg = sysConfig("TIPO_EMPRESA")
'tipoEmpresa = cCfg.Value
'Set cCfg = Nothing

If tipoEmpresa = 4 Then
   txtQuant.Text = 1
Else
   If Left(txtCodBarraPeso.Text, 1) <> "2" Then
      txtQuant.Text = 1
   Else
      'calcular kilo pelas gramas
      'If txtUnidMed.Text = "KG" Then
      '   txtQuant = Mid(txtCodBarraPeso, 8, 5) / 1000
      'Else
      '   txtQuant = Mid(txtCodBarraPeso, 8, 5)
      'End If

      'calcular kilo pelas preço
      'If txtUnidMed.Text = "KG" Then
      '   txtQuant = Mid(txtCodBarraPeso, 8, 5) / 1000
      'Else
         Dim ValorCompra As Currency
         Dim ValorKilo As Currency
         Dim Peso As Double
         
         ValorKilo = CCur(txtValor.Text)
         ValorCompra = CCur(Mid(txtCodBarraPeso, 8, 5)) / 100
         Peso = ValorCompra / CDbl(ValorKilo)
         
         txtQuant = Peso
         
      'End If
   End If
End If

txtCodProduto.Text = r("var_codprod")
Calcular_Total
   
SairProc:
   If Not r Is Nothing Then If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub Mostrar_Produto_Alterar()
'Dim sSQL As String
'Dim r As ADODB.Recordset

If txtCodItem.Text = "" Then Exit Sub

sSQL = "SELECT pedidos_itens.codigo, pedidos_itens.cod_pedido, pedidos_itens.preco, pedidos_itens.quantidade, produtos.descricao " & _
   "FROM pedidos_itens INNER JOIN produtos ON pedidos_itens.cod_produto = produtos.codigo " & _
   "WHERE (pedidos_itens.codigo = " & txtCodItem.Text & ");"

Set r = dbData.OpenRecordset(sSQL)

 txtQuant.Enabled = True
 
If Not r.BOF Then
    txtCodProduto.Text = r("codigo")
    lblDesc.Caption = r("descricao")
    'txtCodBarra.Text = r("var_codbarra")
    txtValor.Text = Format(r("preco"), ocMONEY)
    txtQuant.Text = Format(r("quantidade"), ocPESO)
    If r.State <> 0 Then r.Close
    Set r = Nothing
    Calcular_Total
    txtQuant.BackColor = &HC0FFC0
    cmdFinalizarAvista.Enabled = False
    cmdFinalizarPrazo.Enabled = False
    cmdOrçamento.Enabled = False
    cmdCancelarPedido.Enabled = False
    cmdRemover.Enabled = False
    cmdAvancado.Enabled = False
    cmdInfProduto.Enabled = False
    txtCodBarra.Enabled = False
    txtQuant.SetFocus
End If


End Sub

Private Sub ConsultarProdutosNFCe()
If txtCodPedido.Text = "" Then Exit Sub
Dim varCodProduto As Long

Dim ProdutoEAN As Boolean
Dim ProdutoNCM As Boolean
Dim ProdutoCFOP As Boolean
Dim ProdutoCST As Boolean

Dim i As Integer

Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT pedidos_itens.codigo, produtos.codigo as varCodProd, produtos.ean, produtos.ncm, produtos.cfop, produtos.icmscst, produtos.descricao, produtos.unid_medida, pedidos_itens.cod_produto " & _
   "FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
   "WHERE (pedidos_itens.cod_pedido = " & txtCodPedido.Text & ") ORDER BY pedidos_itens.codigo DESC;"
Set r = dbData.OpenRecordset(sSQL)


For i = 1 To r.RecordCount
    varCodProduto = r!varCodProd
    
    If r!EAN Is Nothing Or IsNull(r!EAN) Or IsEmpty(r!EAN) Or r!EAN = "" Then
        'NFCe_OK = True
        ProdutoEAN = True
    Else
        If Len(r!EAN) < 13 Then
            MsgBox "Produto " & r!descricao & " nao tem EAN ou não está incorreto"
            If ShowMsg("Deseja atualizar o produto " & r!descricao & " ?", vbInformation + vbYesNo) = vbYes Then
                Load Produtos_Cadastro
                Produtos_Cadastro.SSTab1.Tab = 0
                Produtos_Cadastro.cmdNovo.Enabled = False
                Produtos_Cadastro.cmdSalvar.Enabled = False
                Produtos_Cadastro.cmdCancelar.Enabled = False
                Produtos_Cadastro.cmdAlterar.Enabled = True
                Produtos_Cadastro.cmdExcluir.Enabled = True
                Produtos_Cadastro.txtCodigo.Text = varCodProduto
                Produtos_Cadastro.Show
                ProdutoEAN = False
                'PararFechamentoVenda = True
                'Exit Sub
            Else
                MsgBox "Não será possivel transmitir essa NFCe enquanto não corrigir os erros!", vbInformation, "Aviso do Sistema"
                ProdutoEAN = False
                'PararFechamentoVenda = True
                'Exit Sub
            End If
        Else 'se for igual ou maior
            ProdutoEAN = True
        End If
    End If

    If Len(r!NCM) < 8 Or r!NCM = Empty Then
        MsgBox "Produto " & r!descricao & " nao tem NCM ou não está incorreto"
        If ShowMsg("Deseja atualizar o produto " & r!descricao & " ?", vbInformation + vbYesNo) = vbYes Then
            Load Produtos_Cadastro
            Produtos_Cadastro.SSTab1.Tab = 0
            Produtos_Cadastro.cmdNovo.Enabled = False
            Produtos_Cadastro.cmdSalvar.Enabled = False
            Produtos_Cadastro.cmdCancelar.Enabled = False
            Produtos_Cadastro.cmdAlterar.Enabled = True
            Produtos_Cadastro.cmdExcluir.Enabled = True
            Produtos_Cadastro.txtCodigo.Text = varCodProduto
            Produtos_Cadastro.Show
            ProdutoNCM = False
            'PararFechamentoVenda = True
            'Exit Sub
        Else
            MsgBox "Não será possivel transmitir essa NFCe enquanto nao corrigir os erros!", vbInformation, "Aviso do Sistema"
            'NFCe_OK = False
            ProdutoNCM = False
            'Exit Sub
        End If
    Else
        ProdutoNCM = True
    End If
    
    If Len(r!CFOP) < 4 Or r!CFOP = Empty Then
        MsgBox "Produto " & r!descricao & " nao tem CFOP ou não está incorreto"
        If ShowMsg("Deseja atualizar o produto " & r!descricao & " ?", vbInformation + vbYesNo) = vbYes Then
            Load Produtos_Cadastro
            Produtos_Cadastro.SSTab1.Tab = 0
            Produtos_Cadastro.cmdNovo.Enabled = False
            Produtos_Cadastro.cmdSalvar.Enabled = False
            Produtos_Cadastro.cmdCancelar.Enabled = False
            Produtos_Cadastro.cmdAlterar.Enabled = True
            Produtos_Cadastro.cmdExcluir.Enabled = True
            Produtos_Cadastro.txtCodigo.Text = varCodProduto
            Produtos_Cadastro.Show
            ProdutoCFOP = False
            'PararFechamentoVenda = True
            'Exit Sub
        Else
            MsgBox "Não será possivel transmitir essa NFCe enquanto nao corrigir os erros!", vbInformation, "Aviso do Sistema"
            'NFCe_OK = False
            ProdutoCFOP = False
            'Exit Sub
        End If
    Else
        ProdutoCFOP = True
    End If
    
    If Len(r!icmsCST) < 3 Or r!icmsCST = Empty Then
        MsgBox "Produto " & r!descricao & " nao tem CST ou não está incorreto"
        If ShowMsg("Deseja atualizar o produto " & r!descricao & " ?", vbInformation + vbYesNo) = vbYes Then
            Load Produtos_Cadastro
            Produtos_Cadastro.SSTab1.Tab = 0
            Produtos_Cadastro.cmdNovo.Enabled = False
            Produtos_Cadastro.cmdSalvar.Enabled = False
            Produtos_Cadastro.cmdCancelar.Enabled = False
            Produtos_Cadastro.cmdAlterar.Enabled = True
            Produtos_Cadastro.cmdExcluir.Enabled = True
            Produtos_Cadastro.txtCodigo.Text = varCodProduto
            Produtos_Cadastro.Show
            ProdutoCST = False
            'PararFechamentoVenda = True
            'Exit Sub
        Else
            MsgBox "Não será possivel transmitir essa NFCe enquanto nao corrigir os erros!", vbInformation, "Aviso do Sistema"
            'NFCe_OK = False
            ProdutoCST = False
        End If
    Else
        ProdutoCST = True
    End If
    
r.MoveNext
Next

    If ProdutoEAN = False Or ProdutoNCM = False Or ProdutoCFOP = False Or ProdutoCST = False Then
        PararFechamentoVenda = True
    Else
        PararFechamentoVenda = False
    End If

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub MostrarGrid_Produtos()
If txtCodPedido.Text = "" Then Exit Sub

'Dim sSQL As String
'Dim r As ADODB.Recordset

sSQL = "SELECT pedidos_itens.codigo, produtos.ref AS var_ref, produtos.tamanho AS var_tam, produtos.fabricante AS var_fab, pedidos_itens.cod_produto, produtos.descricao, pedidos_itens.preco, produtos.PEDIRPESO, " & _
   "pedidos_itens.quantidade, (pedidos_itens.preco * pedidos_itens.quantidade) as total FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
   "WHERE (pedidos_itens.cod_pedido = " & txtCodPedido.Text & ") ORDER BY pedidos_itens.codigo DESC;"
Set r = dbData.OpenRecordset(sSQL, totalRegistros)
'Debug.Print sSQL

'If Grid.Rows >= 2 Then
'    If Not r.EOF Then
'        vPedirPeso = Abs(CBool(r("PEDIRPESO")))
'    Else
'        vPedirPeso = False
'    End If
'End If

If r.RecordCount <> 0 Then
    lblQuantTipo.Caption = Format(totalRegistros, "00") & " tipo(s)"
    vQuantItensVenda = totalRegistros
Else
    lblQuantTipo.Caption = ""
    vQuantItensVenda = 0
End If

FormatarGrid_Produtos r

'If vPedirPeso = True Then
'    If frmVendaFechamento.Visible = False Then
'        If frmProdutoNaoCadastrado.Visible = False Then Grid_DblClick
'    End If
'End If

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub
Private Sub Mostrar_ValorRestante()
Dim Valor As Currency
Dim QUANT As Integer
Dim ENTRADA As Currency
Dim RESULTADO As Currency
Dim VALOR_SENTRADA As Currency

If txtEntrada.Text <> "" Then
    If txtEntrada.Text = "0,00" Then ENTRADA = 0 Else ENTRADA = txtEntrada.Text
    If txtTotalDesc.Text = "" Then Valor = 0 Else Valor = txtTotalDesc.Text
    If ENTRADA > Valor Then MsgBox "Valor da entrada superior ao valor da venda!", vbExclamation, "Aviso do Sistema"
    VALOR_SENTRADA = Valor - ENTRADA
    txtValorRest.Text = Format(VALOR_SENTRADA, ocMONEY)
End If
End Sub

Private Sub Teste()
'If vConfImprimeNFCeLocal = "SIM" Then   'SE no arquivo ini tem SIM para imprimir NFCE
        'If NFCe_OK = True Then              'SE dei SIM para imprimir NFCE
'            If vNFCeCombinarImp = "SIM" Then
'                 ImprimirVendaAP
'            End If              'final de vNFCeCombinarImp = "SIM"
'        Else                    'meio do NFCe_OK = false
'            ImprimirVendaAP
'        End If                  'fim do NFCe_OK = True
'Else                            'meio do vConfImprimeNFCeLocal = "SIM"
'    ImprimirVendaAP
'End If                          'fim do vConfImprimeNFCeLocal = "SIM"
End Sub

Private Sub Verifica_Existencia_Produto()
   If txtCodBarra.Text = "" Then Exit Sub
   
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT produtos.cod_barra AS var_codbarra, produtos.unid_medida AS var_unidmed, produtos.codigo AS var_codprod, " & _
      "produtos.descricao AS var_desc, produtos_entrada_itens.venda AS var_venda FROM produtos LEFT JOIN ultimas_entradas ON produtos.codigo = ultimas_entradas.codigo_produto " & _
      "LEFT JOIN produtos_entrada_itens ON ultimas_entradas.codigo_produto = produtos_entrada_itens.codigo_produto AND ultimas_entradas.ultentrada = produtos_entrada_itens.codigo_entrada " & _
      "WHERE (produtos.cod_barra = '" & txtCodBarra.Text & "') AND (produtos.ativo = 1);"
   
   Set r = dbData.OpenRecordset(sSQL)
   EXISTENCIA_PRODUTO = Not r.BOF
   If r.BOF Then ShowMsg "Produto não cadastrado!", vbInformation
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub Verifica_QuantEstoque()
'descobrir o codigo do produto
Dim sSQL As String
Dim r As ADODB.Recordset
Dim vCodProduto As Long

If txtCodBarra.Text = "" Then Exit Sub

sSQL = "SELECT produtos.codigo AS var_codprod, produtos.cod_barra, produtos.ativo, " & _
   "produtos_entrada_itens.codigo FROM produtos INNER JOIN produtos_entrada_itens ON produtos.codigo = produtos_entrada_itens.codigo_produto " & _
   "WHERE (produtos.cod_barra = '" & txtCodBarra.Text & "') AND (produtos.ativo = 1) ORDER BY produtos_entrada_itens.codigo DESC;"

Set r = dbData.OpenRecordset(sSQL)

If r.BOF Then
   ShowMsg "Produto Inexistente!", vbCritical
   LimparObjetos_Produto
   txtCodBarra.SetFocus
   Exit Sub
End If

vCodProduto = r("var_codprod")

'verificar quantidade
Dim vQtde As Double
   
   'Consulta os saldos
   sSQL = "SELECT quant_estoque FROM produtos WHERE (codigo = " & vCodProduto & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then vQtde = ValidateNull(r("quant_estoque"))
      If r.State <> 0 Then r.Close
   Set r = Nothing

'Calcula o saldo atual em estoque
'vQtde = EstoqueVendas(vCodProduto)
   
If vQtde <= 0 Then
   Dim oCfg As ConfigItem
   Dim bEstNeg As Boolean
   
   'Recupera a configuração do estoque
   Set oCfg = sysConfig("ESTOQUE_NEGATIVO")
   bEstNeg = CBool(oCfg.Value)
   Set oCfg = Nothing
   
   If Not bEstNeg Then
      ShowMsg "A quantidade em estoque é insuficiente.", vbExclamation
      LimparObjetos_Produto
      Exit Sub
   End If
End If

End Sub

Private Sub Verificar_Backup()
'desativei para ver colocar no online commerce
'If TimeValue(Now) > TimeValue("12:30:00") Then
'    Dim DataHora As Date
'    MensagemErro = ""
   
'   sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
'   Set r = dbData.OpenRecordset(sSQL)
   
'   DataHora = Now
   
'   If Vazio(r!BackupDataHora) Then
'        timerBackup.Enabled = True
'   ElseIf Day(r!BackupDataHora) < Day(DataHora) Then
'        timerBackup.Enabled = True
'   ElseIf Day(r!BackupDataHora) = Day(DataHora) Then
'        timerBackup.Enabled = False
'   ElseIf Day(r!BackupDataHora) > Day(DataHora) Then
'        timerBackup.Enabled = False
'   End If
'End If
End Sub

Private Sub Verificar_NFCe()
'Dim sSQL As String
'Dim r As ADODB.Recordset

sSQL = "SELECT IdNFProd " & _
       "FROM TbNFCe " & _
       "WHERE (DataEmissao = CONVERT(date, GETDATE())) and TbNFCe.NFCeEnviada = 0 and TbNFCe.NFCeCancelada = 0 and TbNFCe.Inutilizada = 0;"
       'Debug.Print sSQL
Set r = dbData.OpenRecordset(sSQL)

If r.BOF Then
    lblAlerta.Visible = False
    lblNfce1.Visible = False
    lblNfce2.Visible = False
Else
    lblAlerta.Visible = True
    lblNfce1.Visible = True
    lblNfce2.Visible = True
End If
End Sub

Private Sub VerificarConsignado()
'Dim sSQL As String
'Dim r As ADODB.Recordset

sSQL = "SELECT COD_PEDIDO, DATEDIFF(day, DATA_COMPRA, GETDATE()) AS vQuantDias FROM pedidos wHERE (TIPO_PEDIDO = 'CONSIGNADO');"
Set r = dbData.OpenRecordset(sSQL)

Dim vQuantDiasCons As Integer
vQuantDiasCons = r("vQuantDias")

If vQuantDiasCons > 5 Then
    lblMSG1.Visible = True
    lblAlerta.Visible = True
    lblMSG1.Caption = "Há consignado em aberto!"
Else
    lblMSG1.Visible = False
    lblMSG1.Visible = False
    lblMSG1.Caption = ""
End If

If Not r.State <> 0 Then r.Close
Set r = Nothing
End Sub


Private Sub VerificarUnidadeMedidas()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodBarra.Text = "" And txtCodProduto.Text = "" Then Exit Sub

sSQL = "SELECT DISTINCT codigo, cod_barra, ISNULL(UNID_MEDIDA, 0) as vUnid " & _
    "FROM produtos WHERE (produtos.cod_barra = '" & txtCodBarra.Text & "') "
Set r = dbData.OpenRecordset(sSQL)

If r("vUnid") = "KG" Then
    If Left(txtCodBarra.Text, 1) = "2" And Len(txtCodBarra.Text) = 13 Then
        Dim varCodProdMed As String
        If varTipoEtiqueta = "5" Then
            varCodProdMed = Format(Mid(txtCodBarra, 2, 5), "00000")
        ElseIf varTipoEtiqueta = "4" Then
            varCodProdMed = Format(Mid(txtCodBarra, 2, 4), "00000")
        ElseIf varTipoEtiqueta = "7" Then
            varCodProdMed = Format(Mid(txtCodBarra, 4, 4), "00000")
        End If
    End If
Else
    varCodProdMed = txtCodBarra.Text
End If

varUnidMed = r("vUnid")
 
'If Left(txtCodBarra.Text, 1) = "2" And Len(txtCodBarra.Text) = 13 Then
'    Dim varCodProdMed As String
'    If varTipoEtiqueta = "2" Then
'        varCodProdMed = Mid(txtCodBarra, 2, 4)
'    ElseIf varTipoEtiqueta = "4" Then
'        varCodProdMed = Mid(txtCodBarra, 4, 4)
'    End If
'Else
'    varCodProdMed = txtCodBarra.Text
'End If

'Dim sSQL As String
'Dim r As ADODB.Recordset
'Dim varUnidMed As String

'If txtCodBarra.Text = "" And txtCodProduto.Text = "" Then Exit Sub
 
'If Left(txtCodBarra.Text, 1) = "2" And Len(txtCodBarra.Text) = 13 Then
'    Dim varCodProdMed As String
'    If varTipoEtiqueta = "2" Then
'        varCodProdMed = Mid(txtCodBarra, 2, 4)
'    ElseIf varTipoEtiqueta = "4" Then
'        varCodProdMed = Mid(txtCodBarra, 4, 4)
'    End If
    
'    sSQL = "SELECT DISTINCT produtos.codigo, produtos.cod_barra, produtos.UNID_MEDIDA as vUnid " & _
'    "FROM produtos WHERE (produtos.cod_barra = '" & varCodProdMed & "') AND (produtos.ativo = 1) " & _
'    "ORDER BY produtos.codigo;"
'    Set r = dbData.OpenRecordset(sSQL)
    
'    varUnidMed = r("vUnid")
'Else
'    varCodProdMed = txtCodBarra.Text
'End If

End Sub

Private Sub cboCliente_Change()
'cboCliente_LostFocus
End Sub

Private Sub cboCliente_Click()
CboCliente_LostFocus
End Sub

Private Sub CboCliente_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset

'If cboCliente.Text <> "CONSUMIDOR" Then
    If cboCliente.ListCount = 0 Then
       sSQL = "SELECT DISTINCT nome, codigo FROM cliente WHERE STATUS = 1 ORDER BY nome;"
       Set r = dbData.OpenRecordset(sSQL)
       
       Do While Not r.EOF
          cboCliente.AddItem r("nome")
          cboCliente.ItemData(cboCliente.NewIndex) = r("codigo")
          r.MoveNext
       Loop
    End If
'End If

moCombo.AttachTo cboCliente
End Sub

Private Sub CboCliente_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
      CboCliente_LostFocus
      cmdFinalizar_Click
End If
End Sub

Private Sub CboCliente_LostFocus()
On Error GoTo TrataErro

'Dim sSQL As String             'desativei dia 13/05/24
'Dim r As ADODB.Recordset       'desativei dia 13/05/24

If cboCliente.Text <> "CONSUMIDOR" Then
    If Cliente_Debito = True Then
       tmrDebito.Enabled = True
       lblInfoDebito.Visible = True
       Cliente_Debito = True
    Else
       tmrDebito.Enabled = False
       lblInfoDebito.Visible = False
       Cliente_Debito = False
    End If
End If

    If lblEstornar.Caption = "ESTORNO" And lblEstornar.Caption <> "REIMPRESSÃO" Then
       If cboCliente.Text = "" Then txtCodCliente.Text = "": Exit Sub
       'If cboCliente.ListIndex = -1 Then txtCodCliente.Text = "": Exit Sub
       txtCodCliente = cboCliente.ItemData(cboCliente.ListIndex)
       Exit Sub
    End If
    
    If lblEstornar.Caption = "" Then
         If cboCliente.ListIndex = -1 Then
            sSQL = "SELECT codigo, nome FROM cliente WHERE (nome = '" & cboCliente.Text & "');"
            Set r = dbData.OpenRecordset(sSQL)
            
            If Not r.BOF Then
               'cboCliente.Text = r("nome")
               txtCodCliente.Text = r("codigo")
            Else
                txtCodCliente.Text = ""
            End If
            
            If r.State <> 0 Then r.Close
            Set r = Nothing
        Else
            txtCodCliente = cboCliente.ItemData(cboCliente.ListIndex)
        End If
    End If

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub

Private Sub cboformaPgto_Change()
Calcular_Desconto
End Sub

Private Sub cboformaPgto_Click()
Calcular_Desconto
End Sub


Private Sub cboformaPgto_GotFocus()
Dim varTexto As String
varTexto = cboFormaPgto.Text
    cboFormaPgto.Clear
    Preencher_FormaPgto
cboFormaPgto.Text = varTexto
SelectControl cboFormaPgto
moCombo.AttachTo cboFormaPgto
End Sub


Private Sub cboFormaPgto_LostFocus()
If vLimitarDesc = 1 Then
    If cboFormaPgto.Text = "3 - CARTÃO - DÉBITO" Then
        If vDescCartaoDebito = "SIM" And txtDesc.Text <> "0,00" Then
            If ShowMsg("Não é permitido dar desconto para vendas com pagamento em cartão de débito!" & Chr(13) & "Deseja mudar a forma de pagamento ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                cboFormaPgto.Text = ""
                cboFormaPgto.SetFocus
            Else
                MsgBox "O valor do desconto será zerado!", vbInformation, "Aviso do Sistema"
                txtDesc.Text = FormatNumber(0, 2)
                Calcular_Desconto
            End If
        End If
    ElseIf cboFormaPgto.Text = "4 - CARTÃO - CRÉDITO" Then
        If vDescCartaoCredito = "SIM" And txtDesc.Text <> "0,00" Then
            If ShowMsg("Não é permitido dar desconto para vendas com pagamento em cartão de crédito!" & Chr(13) & "Deseja mudar a forma de pagamento ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                cboFormaPgto.Text = ""
                cboFormaPgto.SetFocus
            Else
                MsgBox "O valor do desconto será zerado!", vbInformation, "Aviso do Sistema"
                txtDesc.Text = FormatNumber(0, 2)
                Calcular_Desconto
            End If
        End If
    End If
End If
End Sub

Private Sub cboFormaPgtoEntrada_Change()
Calcular_Desconto
End Sub

Private Sub cboFormaPgtoEntrada_Click()
Calcular_Desconto
End Sub


Private Sub cboFormaPgtoEntrada_GotFocus()
cboFormaPgtoEntrada.AddItem "1 - DINHEIRO"
cboFormaPgtoEntrada.AddItem "3 - CARTÃO - DÉBITO"
cboFormaPgtoEntrada.AddItem "4 - CARTÃO - CRÉDITO"
cboFormaPgtoEntrada.AddItem "5 - CHEQUE"
cboFormaPgtoEntrada.AddItem "6 - BOLETO"
cboFormaPgtoEntrada.AddItem "7 - TRANSFERÊNCIA"
cboFormaPgtoEntrada.AddItem "8 - DEPOSITO"
cboFormaPgtoEntrada.AddItem "10 - PIX"
End Sub


Private Sub cboFormaPgtoEntrada_LostFocus()
If vLimitarDesc = 1 Then
    If cboFormaPgtoEntrada.Text = "3 - CARTÃO - DÉBITO" Then
        If vDescCartaoDebito = "SIM" And txtDesc.Text <> "0,00" Then
            If ShowMsg("Não é permitido dar desconto para vendas com pagamento em cartão de débito!" & Chr(13) & "Deseja mudar a forma de pagamento ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                cboFormaPgtoEntrada.Text = ""
                cboFormaPgtoEntrada.SetFocus
            Else
                MsgBox "O valor do desconto será zerado!", vbInformation, "Aviso do Sistema"
                txtDesc.Text = FormatNumber(0, 2)
                Calcular_Desconto
            End If
        End If
    ElseIf cboFormaPgtoEntrada.Text = "4 - CARTÃO - CRÉDITO" Then
        If vDescCartaoCredito = "SIM" And txtDesc.Text <> "0,00" Then
            If ShowMsg("Não é permitido dar desconto para vendas com pagamento em cartão de crédito!" & Chr(13) & "Deseja mudar a forma de pagamento ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                cboFormaPgtoEntrada.Text = ""
                cboFormaPgtoEntrada.SetFocus
            Else
                MsgBox "O valor do desconto será zerado!", vbInformation, "Aviso do Sistema"
                txtDesc.Text = FormatNumber(0, 2)
                Calcular_Desconto
            End If
        End If
    End If
End If
End Sub

Private Sub cboMaquina_GotFocus()
   cboMaquina.Clear
   cboMaquina.AddItem "CAIXA01"
   cboMaquina.AddItem "CAIXA02"
   cboMaquina.AddItem "CAIXA03"
   moCombo.AttachTo cboMaquina
End Sub

Private Sub cboPrazo_Change()
Calcular_Prazo
End Sub

Private Sub cboPrazo_Click()
   Calcular_Prazo
End Sub

Private Sub cboPrazo_GotFocus()
Dim varTexto As String
varTexto = cboPrazo.Text
   cboPrazo.Clear
   cboPrazo.AddItem "5"
   cboPrazo.AddItem "10"
   cboPrazo.AddItem "15"
   cboPrazo.AddItem "20"
   cboPrazo.AddItem "30"
cboPrazo.Text = varTexto
SelectControl cboPrazo
   moCombo.AttachTo cboPrazo
End Sub

Private Sub cboQuantForma_Change()
cboQuantForma_LostFocus
Mostrar_ValorRestante
End Sub

Private Sub cboQuantForma_Click()
cboQuantForma_Change
End Sub


Private Sub cboQuantForma_GotFocus()
Dim varTexto As String
varTexto = cboQuantForma.Text
    cboQuantForma.Clear
If cboTipoPgto.Text = "À VISTA" Then
    cboQuantForma.AddItem "1 - FORMA"
    cboQuantForma.AddItem "2 - FORMAS"
ElseIf cboTipoPgto.Text = "À PRAZO" Then
    cboQuantForma.AddItem "1 - SEM ENTRADA"
    cboQuantForma.AddItem "2 - COM ENTRADA"
End If

cboQuantForma.Text = varTexto
SelectControl cboQuantForma
moCombo.AttachTo cboQuantForma
End Sub


Private Sub cboQuantForma_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
      cmdFinalizar_Click
End If
End Sub

Private Sub cboQuantForma_LostFocus()
If cboQuantForma.Text = "1 - FORMA" Then
    lblEntrada.Enabled = False
    txtEntrada.Enabled = False
    lblFormaEntrada.Enabled = False
    cboFormaPgtoEntrada.Enabled = False
    cboFormaPgto.Enabled = True
    lblFormaParcelas.Enabled = True
    lblValorParc.Enabled = True
    txtValorRest.Enabled = True
    'txtValorRest.Locked = True
ElseIf cboQuantForma.Text = "2 - FORMAS" Then
    lblEntrada.Enabled = True
    txtEntrada.Enabled = True
    lblFormaEntrada.Enabled = True
    cboFormaPgtoEntrada.Enabled = True
    cboFormaPgto.Enabled = True
    lblFormaParcelas.Enabled = True
    lblValorParc.Enabled = True
    txtValorRest.Enabled = True
    'txtValorRest.Locked = True
ElseIf cboQuantForma.Text = "1 - SEM ENTRADA" Then
    lblEntrada.Enabled = False
    txtEntrada.Enabled = False
    lblFormaEntrada.Enabled = False
    cboFormaPgtoEntrada.Enabled = False
    cboFormaPgto.Enabled = True
    lblFormaParcelas.Enabled = True
    lblValorParc.Enabled = True
    txtValorRest.Enabled = True
    txtEntrada.Text = Format(0, ocMONEY)
    'txtValorRest.Locked = True
ElseIf cboQuantForma.Text = "2 - COM ENTRADA" Then
    lblEntrada.Enabled = True
    txtEntrada.Enabled = True
    lblFormaEntrada.Enabled = True
    cboFormaPgtoEntrada.Enabled = True
    cboFormaPgto.Enabled = True
    lblFormaParcelas.Enabled = True
    lblValorParc.Enabled = True
    txtValorRest.Enabled = True
    'txtValorRest.Locked = True
End If
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
Dim varTexto As String
varTexto = cboQuantParc.Text

   cboQuantParc.Clear
   For i = 1 To 15
      cboQuantParc.AddItem i
   Next
cboQuantParc.Text = varTexto
SelectControl cboQuantParc
moCombo.AttachTo cboQuantParc
End Sub

Private Sub cboQuantParc_LostFocus()
Calcular_Parcelas
Calcular_Prazo
End Sub
Private Sub cboTipoPgto_Change()
If cboTipoPgto.Text = "À VISTA" Then
    txtEntrada.Enabled = False
    cboPrazo.Enabled = False
    txtValorRest.Enabled = False
    txtValorParc.Enabled = False
    Label7.Enabled = False
    lblEntrada.Enabled = False
    lblQuantParc.Enabled = False
    lblValorParc.Enabled = False
    
    lblQtdeParc.Enabled = False
    cboQuantParc.Enabled = False
    
    
    lblInicio.Enabled = False
    lblInicio.Caption = "Data"
    mskInicio.Enabled = False
    cmdCal1.Enabled = False
    
    lblTermino.Visible = False
    mskTermino.Visible = False
    
    cboFormaPgtoEntrada.Enabled = False
    lblFormaEntrada.Enabled = False
    BuscarClienteConsumidor
    cboFormaPgto.Text = "1 - DINHEIRO"
    cboFormaPgtoEntrada.Text = "1 - DINHEIRO"
ElseIf cboTipoPgto.Text = "À PRAZO" Then
    txtEntrada.Enabled = True
    cboPrazo.Enabled = True
    txtValorRest.Enabled = True
    txtValorParc.Enabled = True
    Label7.Enabled = True
    lblEntrada.Enabled = True
    lblQuantParc.Enabled = True
    lblValorParc.Enabled = True
    
    lblQtdeParc.Enabled = True
    cboQuantParc.Enabled = True
    
    lblInicio.Enabled = True
    lblInicio.Caption = "Inicio"
    mskInicio.Enabled = True
    
    
    lblTermino.Enabled = True
    mskTermino.Enabled = True
    lblTermino.Visible = True
    mskTermino.Visible = True
    
    cboFormaPgtoEntrada.Enabled = True
    lblFormaEntrada.Enabled = True
    cboFormaPgto.Text = "1 - PROMISSÓRIA"
    cboFormaPgtoEntrada.Text = "1 - DINHEIRO"
ElseIf cboTipoPgto.Text = "ORÇAMENTO" Then
    txtEntrada.Enabled = False
    cboPrazo.Enabled = True
    txtValorRest.Enabled = False
    txtValorParc.Enabled = False
    lblEntrada.Enabled = False
    lblQuantParc.Enabled = False
    lblValorParc.Enabled = False
    Label7.Enabled = False
    
    lblQuantParc.Enabled = True
    lblQtdeParc.Enabled = True
    cboQuantParc.Enabled = True
    
    lblInicio.Enabled = True
    lblInicio.Caption = "Data"
    mskInicio.Enabled = True
    cmdCal1.Enabled = True
    
    If cboQuantParc = "1" Then
        lblTermino.Enabled = False
        mskTermino.Enabled = False
        lblTermino.Visible = False
        mskTermino.Visible = False
    Else
        lblTermino.Enabled = False
        mskTermino.Enabled = False
        lblTermino.Visible = True
        mskTermino.Visible = True
    End If
    
    BuscarClienteConsumidor
    cboFormaPgtoEntrada.Enabled = True
    lblFormaEntrada.Enabled = True
    cboFormaPgto.Text = "1 - DINHEIRO"
    cboFormaPgtoEntrada.Text = "1 - DINHEIRO"
End If

End Sub
Private Sub cboUsuario_GotFocus()
Dim sSQL As String
Dim r As ADODB.Recordset
cboUsuario.Clear
sSQL = "SELECT codigo, login FROM usuario WHERE (visivel = 1) ORDER BY login;"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   cboUsuario.AddItem r("login")
   cboUsuario.ItemData(cboUsuario.NewIndex) = r("codigo")
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing

moCombo.AttachTo cboUsuario
End Sub
Private Sub cboUsuario_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub cboUsuario_LostFocus()
On Error GoTo TrataErro

If cboUsuario.Text = "" Then txtCodUsuario.Text = "": txtNivelUsuario.Text = "": Exit Sub
If cboUsuario.ListIndex = -1 Then txtCodUsuario.Text = "": txtNivelUsuario.Text = "": Exit Sub
txtCodUsuario = cboUsuario.ItemData(cboUsuario.ListIndex)

TrataErro:
   If Err.Number = 381 Then Exit Sub
End Sub


Private Sub PegarPesoUrano()
Dim comandobalanca As String
comandobalanca = Chr(4)
enviaComandoSerial (comandobalanca)
End Sub








Private Sub cmdAbrirCaixa_Click()
   Me.Hide
   Caixa_Fechamento.Show 1
   Unload PDV
End Sub

Private Sub cmdAP_Click()
TipoValorVenda = "AP"
frmTipoVenda.Visible = False
'verificar se o pedido está livre
Dim var_NroPedido As Long
var_NroPedido = ExistePedidoLivre

'Nenhum pedido livre
If var_NroPedido = -1 Then
   txtCodPedido = AutoNumeracao_Pedido
   dbData.Execute "INSERT INTO pedidos (cod_pedido, data_compra, status_pedido, caixa, maquina, cancelado, reaberto, orcamento) VALUES (" & txtCodPedido.Text & ", '" & Format$(Now, "yyyy-dd-MM") & "', 0, '" & var_Caixa & "', '" & var_Maquina & "', 0, 0, 0);"
Else
   txtCodPedido = var_NroPedido
End If

HabilitaObjetosVenda False
txtCodBarra.SetFocus
End Sub

Private Sub cmdAV_Click()
TipoValorVenda = "AV"
frmTipoVenda.Visible = False
'verificar se o pedido está livre
Dim var_NroPedido As Long
var_NroPedido = ExistePedidoLivre

'Nenhum pedido livre
If var_NroPedido = -1 Then
   txtCodPedido = AutoNumeracao_Pedido
   dbData.Execute "INSERT INTO pedidos (cod_pedido, data_compra, status_pedido, caixa, maquina, reaberto, cancelado, orcamento) VALUES (" & txtCodPedido.Text & ", '" & Format$(Now, "yyyy-dd-MM") & "', 0, '" & var_Caixa & "', '" & var_Maquina & "', 0, 0, 0);"
Else
   txtCodPedido = var_NroPedido
End If

HabilitaObjetosVenda False
txtCodBarra.SetFocus

    'If LerPermissoesUsuario(vCodUsuario, 18) = True Then
     '    Menu_Fin_Caixa.Enabled = True
     'Else
End Sub

Private Sub cmdAvancado_Click()
If varSegurancaAvancada = "SIM" Then
    'MostrarCaixaSenha
Else
    If vCodFunc = 0 Then
        vCodFunc = txtCodFuncAP.Text
        Permissoes
    End If
End If

If frmAvancado.Visible = False Then
   frmAvancado.Visible = True
   frmTipoVenda.Visible = False
   HabilitaObjetosVenda True
   txtSenha.Text = ""
Else
   frmAvancado.Visible = False
   frmSenha.Visible = False
   
   If varTipoValorVenda = 2 Then
     If CAIXA_FECHADO = False Then frmTipoVenda.Visible = True
   End If
   
   If CAIXA_FECHADO = False Then
        If varTipoValorVenda <> 2 Then
            HabilitaObjetosVenda False
            txtCodBarra.SetFocus
        End If
    End If
End If
End Sub

Private Sub cmdAvancado_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If Button = vbLeftButton Then
'        Dim NumeroDeItensNoMenu As Integer
'        NumeroDeItensNoMenu = 1
'        Me.PopupMenu mnuMenu, , cmdAvancado.Left, cmdAvancado.Top - cmdAvancado.Height - (330 * (NumeroDeItensNoMenu - 1)) + 70
'    End If
End Sub


Private Sub cmdAvanCarne_Click()
Carne.Show 1
End Sub

Private Sub cmdAvanClientes_Click()
varNomeBotao = "Cliente"
If varSegurancaAvancada = "SIM" Then
    Set oCfg = sysConfig("TIPOLOGIN")
    If oCfg.Value = "NOME" Then
        frmSenha.Visible = True
        cboUsuario.Visible = True
        mskCPF.Visible = False
        cboUsuario.Text = ""
        txtCodUsuario.Text = ""
        txtSenha.Text = ""
        Label1.Caption = "Usuário:"
    Else
        frmSenha.Visible = True
        cboUsuario.Visible = False
        mskCPF.Visible = True
        txtSenha.Text = ""
        mskCPF.Mask = ""
        mskCPF.Text = ""
        mskCPF.Mask = "###.###.###-##"
        Label1.Caption = "CPF:"
        If mskCPF.Visible = True Then mskCPF.SetFocus
    End If
Else
    Clientes_Cadastro.Show 1
    cboCliente.Clear
    txtSenha.Text = ""
    frmSenha.Visible = False
    frmAvancado.Visible = False
    HabilitaObjetosVenda False
End If

End Sub

Private Sub cmdAvanEtiquetas_Click()
Etiquetas_Impressao.Show 1
End Sub

Private Sub cmdAvanFinanceiro_Click()
varNomeBotao = "Financeiro"
If varSegurancaAvancada = "SIM" Then
    Set oCfg = sysConfig("TIPOLOGIN")
    If oCfg.Value = "NOME" Then
        frmSenha.Visible = True
        cboUsuario.Visible = True
        mskCPF.Visible = False
        cboUsuario.Text = ""
        txtCodUsuario.Text = ""
        txtSenha.Text = ""
        Label1.Caption = "Usuário:"
        If cboUsuario.Visible = True Then cboUsuario.SetFocus
    Else
        frmSenha.Visible = True
        cboUsuario.Visible = False
        mskCPF.Visible = True
        txtSenha.Text = ""
        mskCPF.Mask = ""
        mskCPF.Text = ""
        mskCPF.Mask = "###.###.###-##"
        Label1.Caption = "CPF:"
        If mskCPF.Visible = True Then mskCPF.SetFocus
    End If
Else
    vChamouCaixa = "PDV"
    Me.Hide
    Principal_Caixa.Show 1
    'If txtCodFunc.Text <> "" Then Principal_Caixa.txtCodFunc.Text = txtCodFunc.Text Else Principal_Caixa.txtCodFunc.Text = "1"
    txtSenha.Text = ""
    frmSenha.Visible = False
    frmAvancado.Visible = False
End If
HabilitaObjetosVenda True
End Sub

Private Sub cmdAvanNFCe_Click()
varNomeBotao = "NFCe"
If varSegurancaAvancada = "SIM" Then
    Set oCfg = sysConfig("TIPOLOGIN")
    If oCfg.Value = "NOME" Then
        frmSenha.Visible = True
        cboUsuario.Visible = True
        mskCPF.Visible = False
        cboUsuario.Text = ""
        txtCodUsuario.Text = ""
        txtSenha.Text = ""
        Label1.Caption = "Usuário:"
    Else
        frmSenha.Visible = True
        cboUsuario.Visible = False
        mskCPF.Visible = True
        txtSenha.Text = ""
        mskCPF.Mask = ""
        mskCPF.Text = ""
        mskCPF.Mask = "###.###.###-##"
        Label1.Caption = "CPF:"
        If mskCPF.Visible = True Then mskCPF.SetFocus
    End If
Else
    NFCe_Consultar.Show
    txtSenha.Text = ""
    frmSenha.Visible = False
    frmAvancado.Visible = False
End If
HabilitaObjetosVenda True
End Sub

Private Sub cmdAvanPedReabrir_Click()
If Grid.Rows >= 2 And lblEstornar.Caption = "ESTORNO" Then
    MsgBox "Existe um estorno em aberto. Finalize-o!", vbInformation, "Aviso do Sistema"
    frmAvancado.Visible = False
    HabilitaObjetosVenda False
    txtCodBarra.SetFocus
    Exit Sub
ElseIf Grid.Rows >= 2 And lblEstornar.Caption <> "ESTORNO" Then
    MsgBox "Existe um pedido em aberto. Finalize ou cancele o pedido!", vbInformation, "Aviso do Sistema"
    frmAvancado.Visible = False
    HabilitaObjetosVenda False
    txtCodBarra.SetFocus
    Exit Sub
Else
    varNomeBotao = "Pedidos"
    
    If varSegurancaAvancada = "SIM" Then
        MostrarCaixaSenha
    Else
        'Estonar.LerPermissoesUsuario vCodUsuario
        cboUsuario.Text = ""
        txtCodUsuario.Text = ""
        txtSenha.Text = ""
        frmAvancado.Visible = False
        Estonar.Hide
        Estonar.lblCodUser1.Visible = True
        Estonar.lblCodUser2.Visible = True
        Estonar.lblUser1.Visible = True
        Estonar.lblUser2.Visible = True
        Estonar.lblCodUser2.Caption = txtCodFunc.Text
        Estonar.lblUser2.Caption = PDV.StatusBar1.Panels(3).Text
        Estonar.Show 1
    End If
End If
HabilitaObjetosVenda True
End Sub

Private Sub cmdAvanProdutos_Click()
varNomeBotao = "Produtos"
If varSegurancaAvancada = "SIM" Then
    Set oCfg = sysConfig("TIPOLOGIN")
    If oCfg.Value = "NOME" Then
        frmSenha.Visible = True
        cboUsuario.Visible = True
        mskCPF.Visible = False
        cboUsuario.Text = ""
        txtCodUsuario.Text = ""
        txtSenha.Text = ""
        Label1.Caption = "Usuário:"
    Else
        frmSenha.Visible = True
        cboUsuario.Visible = False
        mskCPF.Visible = True
        txtSenha.Text = ""
        mskCPF.Mask = ""
        mskCPF.Text = ""
        mskCPF.Mask = "###.###.###-##"
        Label1.Caption = "CPF:"
        If mskCPF.Visible = True Then mskCPF.SetFocus
    End If
Else
    'Estonar.PegarCodUsuario (r("codigo"))   'desativei pq  tenho que entender pq tem esse código aqui
    Produtos_Cadastro.Show 1
    txtSenha.Text = ""
    frmSenha.Visible = False
    frmAvancado.Visible = False
End If
HabilitaObjetosVenda True
End Sub

Private Sub cmdAvanRecAvulso_Click()
Recibos_Avulso.Show 1
End Sub

Private Sub cmdAvanRecibo_Click()
Recibo.Show 1
End Sub

Private Sub cmdAvanSaidaProd_Click()
Produtos_Saida_Estoque.Show 1
frmAvancado.Visible = False
HabilitaObjetosVenda True
End Sub

Private Sub cmdAvanVendaPausar_Click()
'Solicita confirmação do usuário
If ShowMsg("Confirma a operação de pausa nesta venda?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub

'Atualiza o status do pedido
dbData.Execute "UPDATE pedidos SET status_pedido = -1, data_compra = '" & Format$(Now, "yyyy-dd-MM") & "', caixa = '" & StatusBar1.Panels(2).Text & "', maquina = '" & var_Maquina & "' WHERE (cod_pedido = " & txtCodPedido & ");"

'Reinicia o form para uma nova venda
LimparObjetos_Pedido
txtTotalGeral.Text = ""
LimparGrid_Pedido
LimparObjetos_Prazo
txtCodPedido.Text = ""
lblEstornar.Caption = ""
Form_Load
txtCodBarra.SetFocus

frmVendaFechamento.Visible = False
'Liberar = False
End Sub

Private Sub cmdAvanVendaReiniciar_Click()
If Grid.Rows >= 2 And txtTotalGeral.Text <> "" Then
   If ShowMsg("Existe uma venda em aberto. Deseja sair e cancelar a venda?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
      frmAvancado.Visible = False
      frmSenha.Visible = False
      Exit Sub
   End If
      
   dbData.Execute "DELETE FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ");"
   dbData.Execute "DELETE FROM pedidos_itens WHERE (cod_pedido = " & txtCodPedido.Text & ");"
   LimparGrid_Pedido
End If

Dim rPos As RECT
Dim lLft As Long, lTop As Long
Dim fVda As ReinicarVenda
Dim xList As ListItem

Dim sSQL As String
Dim r As ADODB.Recordset
Dim bCancel As Boolean
Dim lNroPedido As Long

Set fVda = New ReinicarVenda
Load fVda

'GetWindowRect cmdOKOpcoes.hwnd, rPos
lLft = rPos.Right * Screen.TwipsPerPixelX - fVda.Width
lTop = rPos.Top * Screen.TwipsPerPixelY - fVda.Height

'Carrega os pedidos pausados
sSQL = "SELECT cod_pedido, data_compra, total FROM pedidos WHERE (status_pedido = -1) AND (caixa = '" & StatusBar1.Panels(2).Text & "');"
Set r = dbData.OpenRecordset(sSQL)

Do While Not r.EOF
   Set xList = fVda.lvwPed.ListItems.Add
   xList.Text = r("cod_pedido")
   '" & Format$(r("data_compra"), "yyyy-dd-MM") & "'
   xList.SubItems(1) = Format$(r("data_compra"), "yyyy-dd-MM")
   xList.SubItems(2) = Format$(r("total"), ocMONEY)
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing
   
fVda.Move lLft, lTop
fVda.Show vbModal

bCancel = fVda.Cancelled
lNroPedido = fVda.OrderNumber

If bCancel Then Exit Sub

Unload fVda
Set fVda = Nothing

'Aqui abre a venda
frmAvancado.Visible = False
txtCodPedido = lNroPedido
MudarPedidoReaberto
MostrarGrid_Produtos
End Sub

Private Sub cmdAvanVendaTransferir_Click()
pTransferirVendaCaixa
End Sub

Private Sub Verificar_Caixa()
sSQL = "SELECT * " & _
       "FROM caixa_dia " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and caixa_dia.status = 0;"
Set r = dbData.OpenRecordset(sSQL)

If r.BOF Then
    varCodCaixa = CInt(0)
    StatusBar1.Panels(7).Text = Format(varCodCaixa, "0000")
    CAIXA_FECHADO = True
Else
    If CDate(r("DATA_ABERTURA")) <> Date Then
        MsgBox "A data do caixa aberto é diferente da data atual!", vbInformation, "Aviso do Sistema"
        lblAlerta.Visible = True
        lblRotuloAberto.Visible = True
        lblDataAberturaCaixa.Visible = True
    Else
        lblAlerta.Visible = False
        lblRotuloAberto.Visible = False
        lblDataAberturaCaixa.Visible = False
    End If
    varCodCaixa = CInt(r("codcaixa"))
    StatusBar1.Panels(7).Text = Format(varCodCaixa, "0000")
    lblDataAberturaCaixa.Caption = CDate(r("DATA_ABERTURA"))
    CAIXA_FECHADO = False
End If

If CAIXA_FECHADO = True Then
    If vBotaoOrcamento = False Then
        txtCodBarra.Enabled = False
        txtValor.Enabled = False
        txtQuant.Enabled = False
        txtTotal.Enabled = False
        txtTotalGeral.Enabled = False
        cmdAlterar.Enabled = False
        cmdFinalizarAvista.Enabled = False
        cmdFinalizarPrazo.Enabled = False
        cmdOrçamento.Enabled = False
        cmdCancelarPedido.Enabled = False
        cmdRemover.Enabled = False
        cmdAvancado.Enabled = False
        cmdInfProduto.Enabled = False
        Grid.Enabled = False
        frmCaixaFechado.Visible = True
        Exit Sub
    Else
        txtCodBarra.Enabled = True
        txtValor.Enabled = True
        txtQuant.Enabled = True
        txtTotal.Enabled = True
        txtTotalGeral.Enabled = True
        cmdAlterar.Enabled = True
        cmdFinalizarAvista.Enabled = False
        cmdFinalizarPrazo.Enabled = False
        cmdOrçamento.Enabled = True
        cmdCancelarPedido.Enabled = True
        cmdRemover.Enabled = True
        cmdAvancado.Enabled = False
        cmdInfProduto.Enabled = False
        frmCaixaFechado.Visible = False
        Grid.Enabled = True
        StatusBar1.Panels(7).Text = Format(0, "0000")
    End If
Else
    If varTipoValorVenda = 2 Then
        If lblEstornar.Caption <> "ESTORNO" Then
            txtDataCompra.Text = Format(Date, "dd/mm/yyyy")
            CriarNovoPedido
        End If
    Else
        'HabilitaObjetosVenda False     'desabilitei no dia 21/05/2024 desconto habilitava essa opção
   End If
End If
End Sub


Private Sub cmdCadastarProduto_Click()
frmProdutoNaoCadastrado.Visible = False
txtCodBarra.Enabled = True
txtQuant.Enabled = True
Produtos_CadastoRapido.Show 1
End Sub

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

mskInicio = Format(varData, "dd/mm/yy")   'Exibe a data no campo
vDataFlexivel = True
Calcular_Prazo
End Sub

Private Sub cmdCancelar_Click()
If lblEstornar.Caption = "ESTORNO" Then
        HabilitaObjetosVenda False
        LimparObjetos_Prazo
        frmVendaFechamento.Visible = False
        txtTotalGeral.Text = Format(txtSubtotal.Text, ocMONEY)
        cmdFinalizarAvista.Enabled = True
        cmdFinalizarPrazo.Enabled = True
        cmdOrçamento.Enabled = False
        cmdCancelarPedido.Enabled = False
        cmdRemover.Enabled = True
        cmdAvancado.Enabled = False
        cmdInfProduto.Enabled = True
        Grid.Enabled = True
        txtCodBarra.Enabled = True
        txtValor.Enabled = True
        txtQuant.Enabled = True
        txtTotal.Enabled = True
        If txtCodBarra.Enabled = True Then txtCodBarra.SetFocus
Else
    If vBotaoOrcAtivo = True Then
        HabilitaObjetosVenda True
        vBotaoOrcamento = False
        vBotaoOrcAtivo = False
        LimparObjetos_Prazo
        frmVendaFechamento.Visible = False
        Form_Load
    Else
        HabilitaObjetosVenda False
        frmVendaFechamento.Visible = False
        txtTotalGeral.Text = Format(txtSubtotal.Text, ocMONEY)
        cmdFinalizarAvista.Enabled = True
        cmdFinalizarPrazo.Enabled = True
        cmdOrçamento.Enabled = True
        cmdCancelarPedido.Enabled = True
        cmdRemover.Enabled = True
        cmdAvancado.Enabled = True
        cmdInfProduto.Enabled = True
        Grid.Enabled = True
        txtCodBarra.Enabled = True
        txtValor.Enabled = True
        txtQuant.Enabled = True
        txtTotal.Enabled = True
        LimparObjetos_Prazo
        If txtCodBarra.Enabled = True Then txtCodBarra.SetFocus
    End If
End If
vUsandoCashBack = False
End Sub

Private Sub cmdCancelarPedido_Click()
If txtQuant.BackColor = &HC0FFC0 Then
    txtQuant_KeyPress (13)
End If

If lblEstornar.Caption <> "ESTORNO" Then
   If Grid.Rows >= 2 Then
      If ShowMsg("Existe uma compra em aberto. Deseja cancelar essa compra?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        dbData.Execute "DELETE FROM pedidos_itens WHERE (cod_pedido = " & txtCodPedido.Text & ");"
        dbData.Execute "DELETE FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ");"
    End If
    'HabilitaObjetosVenda False
    LimparObjetos_Pedido
    txtTotalGeral.Text = ""
    LimparGrid_Pedido
    LimparObjetos_Prazo
    txtCodPedido.Text = ""
    lblEstornar.Caption = ""
    frmVendaFechamento.Visible = False
    
    
    If vBotaoOrcAtivo = True Then
        HabilitaObjetosVenda True
        vBotaoOrcamento = False
        vBotaoOrcAtivo = False
        Form_Load
    Else
        CriarNovoPedido
        If varTipoValorVenda = 1 Then txtCodBarra.SetFocus
    End If
    
    
    'CriarNovoPedido
    'If varTipoValorVenda = 1 Then txtCodBarra.SetFocus
Else
    MsgBox "Não é permitido cancelar esse pedido dessa forma!", vbInformation, "Aviso do Sistema"
    HabilitaObjetosVenda False
    If varTipoValorVenda = 1 Then txtCodBarra.SetFocus
    Exit Sub
End If

vTipoEdicao = ""
'Dim varPosicao As Boolean
'varPosicao = InStr(1, txtCodBarra.Text, "*")
'MsgBox varPosicao

'Dim seuTextBox As String
'Dim qtdeCapturada As String
'Dim barcodCapturado As String
'Dim arrayString As Variant

'seuTextBox = txtCodBarra.Text

'arrayString = Split(seuTextBox, "*", -1)

'qtdeCapturada = arrayString(0)
'barcodCapturado = arrayString(1)
'MsgBox qtdeCapturada
'MsgBox barcodCapturado
End Sub

Private Sub Retorna_Produtos_Estoque()
If lblTipoPedido.Caption <> "ORÇAMENTO" Then
   For i = 1 To Grid.Rows - 1
      dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque + " & Replace(CDbl(Grid.TextMatrix(i, 5)), ",", ".") & " WHERE (codigo = " & Grid.TextMatrix(i, 2) & ");"
   Next
End If
End Sub

Private Sub MudarPedidoReaberto()
sSQL = "UPDATE pedidos SET status_pedido = 0, reaberto = 1, maquina = '" & var_Maquina & "' WHERE (cod_pedido = " & txtCodPedido.Text & ");"
dbData.Execute sSQL

'dbData.Execute "UPDATE Pedidos_Reabertura SET status_pedido = 0 WHERE (cod_pedido = " & txtCodPedido.Text & ");"
End Sub


Private Sub cmdFechar_Click()
'If txtCodPedido.Text = "" Then Exit Sub

If lblEstornar.Caption = "ESTORNO" Then
    MsgBox "Não é possivel sair de um estorno em aberto, Finalize-o!", vbCritical, "Aviso do Sistema"
    Exit Sub
ElseIf lblEstornar.Caption = "REIMPRESSÃO" Then
    MsgBox "Não é possivel sair de uma reimpressão em aberto, Finalize-a!", vbCritical, "Aviso do Sistema"
    Exit Sub
Else
   If Grid.Rows >= 2 Then
      If ShowMsg("Existe uma compra em aberto. Deseja sair e cancelar a compra?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        dbData.Execute "DELETE FROM pedidos_itens WHERE (cod_pedido = " & txtCodPedido.Text & ");"
        dbData.Execute "DELETE FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ");"
      End
   
   ElseIf txtCodPedido.Text <> "" Then
      dbData.Execute "DELETE FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ");"
   End If
End If

'tirar o sistema da memoria
EncerrarPrograma
End Sub

Private Sub cmdFinalizar_Click()
If txtTotalGeral.Text = "" Then Exit Sub
If txtCodPedido.Text = "" Then Exit Sub
If cboTipoPgto.Text = "" Then Exit Sub
If txtRecebido.Text = "" Then txtRecebido.Text = Format(0, ocMONEY)
If txtTroco.Text = "" Then txtTroco.Text = Format(0, ocMONEY)
If txtCodCliente.Text = "" Then cboCliente.Text = "": cboCliente.SetFocus: Exit Sub
If cboTipoPgto.Text = "À PRAZO" And txtCodCliente = "1" Then MsgBox "IDENTIFIQUE O CLIENTE DA COMPRA!", vbExclamation, "Aviso do sistema": cboCliente.SetFocus: Exit Sub
If txtFuncAP.Text = "" Then ShowMsg "Digite o código do funcionário!", vbInformation: txtCodFuncAP.SetFocus: Exit Sub
If txtCodPedido.Text = "" Then MsgBox "Cód. Pedido em Branco": Exit Sub
Dim vValorVenda As Currency
Dim ValorCash As Currency
Dim vCashbackValidade As Date

If cboQuantForma.Text = "2 - FORMAS" Then
    If txtEntrada.Text = "" Or txtEntrada.Text = "0,00" Then
        ShowMsg "Você esqueceu de colocar um dos valores!", vbInformation: txtEntrada.SetFocus: Exit Sub
    End If
End If

cmdFinalizar.Enabled = False

'Dim lNovoCod As Long
Dim varHora As String       'Saber a hora do pagamento da parcela

'Usando na NFCe
Dim vCPF As String
Dim sistNFe As snfe.Util
Dim NFCeContingencia As Boolean

'*****DESATIVEI DIA 29/01/24
''verificar se o caixa ainda está aberto
'sSQL = "SELECT * " & _
'       "FROM caixa_dia " & _
'       "WHERE (codcaixa = " & varCodCaixa & ") AND (caixa = '" & StatusBar1.Panels(2).Text & "');"
'Set r = dbData.OpenRecordset(sSQL)

'If Not r.EOF Then
'    If r("status") = 0 Then
'        'MsgBox "aberto"
'    Else
'        Verificar_Caixa
'        If CAIXA_FECHADO = True Then
'            MsgBox "Não existe nenhum caixa aberto para essa venda!", vbInformation, "Aviso do Sistema"
'            Exit Sub
'        End If
'    End If
'Else
'    Verificar_Caixa
'    If CAIXA_FECHADO = True Then
'        MsgBox "Não existe nenhum caixa aberto para essa venda!", vbInformation, "Aviso do Sistema"
'        Exit Sub
'    End If
'End If


'VERIFICAR SE EMITE NFCE =============================================
NFCe_OK = False
PararFechamentoVenda = False

sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
Set r = dbData.OpenRecordset(sSQL)
NFCeContingencia = r!NFCeOffline


If cboTipoPgto.Text <> "ORÇAMENTO" And cboTipoPgto.Text <> "CONSIGNADO" Then
    If vConfImprimeNFCeLocal = "SIM" Then    'se essa maquina irá imprimir nfce localmente
        'definir a impressora da nfce
        'Dim oIni As Ini 'desativei aqui 09/11/22
        Set oIni = New Ini
        oIni.Arquivo = appPathApp & "config.ini"
        var_ImpNFCe = oIni.LerTexto("IMPRESSORA_NFCE", "impressora")
        Set oIni = Nothing
        
        Dim Prt As Printer
        Dim oldPrinter As String
        
        oldPrinter = Printer.DeviceName
        
        For Each Prt In Printers
           If Prt.DeviceName = var_ImpNFCe Then
              Set Printer = Prt
              Exit For
           End If
        Next
        
        If cboTipoPgto.Text = "À VISTA" Then    'vendas à vista
            If vNFCeConfImp = "SIM" Then
                If MsgBox("Impressora Pronta?", vbQuestion + vbYesNo, "NFCe") = vbYes Then
                    If txtAcresc.Text = "0,00" Then
                        If PararFechamentoVenda = True Then
                            NFCe_OK = False
                            Exit Sub
                        Else
                            NFCe_OK = True
                        End If
                    Else
                        MsgBox "Não é possível gerar NFCE de uma venda com acréscimo!", vbExclamation, "Aviso do Sistema"
                        NFCe_OK = False
                    End If
                Else
                    NFCe_OK = False
                End If
            Else
                If PararFechamentoVenda = True Then
                    NFCe_OK = False
                    Exit Sub
                Else
                    NFCe_OK = True
                End If
            End If
        Else                                    'vendas à prazo
            If vNFCeConfPrazo = "SIM" Then
                If vNFCeConfImp = "SIM" Then
                    If MsgBox("Impressora Pronta?", vbQuestion + vbYesNo, "NFCe") = vbYes Then
                        If txtAcresc.Text = "0,00" Then
                            If PararFechamentoVenda = True Then
                                NFCe_OK = False
                                Exit Sub
                            Else
                                NFCe_OK = True
                            End If
                        Else
                            MsgBox "Não é possível gerar NFCE de uma venda com acréscimo!", vbExclamation, "Aviso do Sistema"
                            NFCe_OK = False
                        End If
                    Else
                        NFCe_OK = False
                    End If
                Else
                    If PararFechamentoVenda = True Then
                        NFCe_OK = False
                        Exit Sub
                    Else
                        NFCe_OK = True
                    End If
                End If
            Else
                NFCe_OK = False
            End If
        End If
    Else
        NFCe_OK = False
    End If
End If

If NFCe_OK = True Then
    Dim EncontroErroNFCe As Boolean
End If


'TIPO DE CARTAO PARCELAS===========================================
Dim varTipoCartao As String
'varTipoCartao = "NULL" 'desativei em 29/01/2024
If cboFormaPgto.Text = "3 - CARTÃO - DÉBITO" Then
   varTipoCartao = "'D'"
ElseIf cboFormaPgto.Text = "4 - CARTÃO - CRÉDITO" Then
   varTipoCartao = "'C'"
Else
    varTipoCartao = "NULL"
End If

'FORMA DE PAGAMENTO RESTANTE============================================
Dim var_PAGAMENTO As String
If cboFormaPgto.Text = "1 - DINHEIRO" Then
   var_PAGAMENTO = "DINHEIRO"
ElseIf cboFormaPgto.Text = "2 - PROMISSÓRIA" Then
   var_PAGAMENTO = "PROMISSORIA"
ElseIf cboFormaPgto.Text = "3 - CARTÃO - DÉBITO" Then
   var_PAGAMENTO = "CARTAO"
ElseIf cboFormaPgto.Text = "4 - CARTÃO - CRÉDITO" Then
   var_PAGAMENTO = "CARTAO"
ElseIf cboFormaPgto.Text = "5 - CHEQUE" Then
   var_PAGAMENTO = "CHEQUE"
ElseIf cboFormaPgto.Text = "6 - BOLETO" Then
   var_PAGAMENTO = "BOLETO"
ElseIf cboFormaPgto.Text = "7 - TRANSFERÊNCIA" Then
   var_PAGAMENTO = "TRANSFERENCIA"
ElseIf cboFormaPgto.Text = "8 - DEPOSITO" Then
   var_PAGAMENTO = "DEPOSITO"
ElseIf cboFormaPgto.Text = "9 - FINANCEIRA" Then
   var_PAGAMENTO = "FINANCEIRA"
ElseIf cboFormaPgto.Text = "10 - PIX" Then
   var_PAGAMENTO = "PIX"
End If

'se houver 2 opçoes de pagamento
If cboQuantForma.Text = "2 - FORMAS" Or cboQuantForma.Text = "2 - COM ENTRADA" Then
    'TIPO DE CARTAO ENTRADA============================================
    Dim varTipoCartaoEntrada As String
    varTipoCartaoEntrada = "NULL"
    
    If cboFormaPgtoEntrada.Text = "3 - CARTÃO - DÉBITO" Then
       varTipoCartaoEntrada = "'D'"
    ElseIf cboFormaPgtoEntrada.Text = "4 - CARTÃO - CRÉDITO" Then
       varTipoCartaoEntrada = "'C'"
    Else
        varTipoCartaoEntrada = "NULL"
    End If
    
    'FORMA DE PAGAMENTO ENTRADA ===============================================
    Dim var_PGTO_Entrada As String
    If cboFormaPgtoEntrada.Text = "1 - DINHEIRO" Then
       var_PGTO_Entrada = "DINHEIRO"
    ElseIf cboFormaPgtoEntrada.Text = "2 - PROMISSÓRIA" Then
       var_PGTO_Entrada = "PROMISSORIA"
    ElseIf cboFormaPgtoEntrada.Text = "3 - CARTÃO - DÉBITO" Then
       var_PGTO_Entrada = "CARTAO"
    ElseIf cboFormaPgtoEntrada.Text = "4 - CARTÃO - CRÉDITO" Then
       var_PGTO_Entrada = "CARTAO"
    ElseIf cboFormaPgtoEntrada.Text = "5 - CHEQUE" Then
       var_PGTO_Entrada = "CHEQUE"
    ElseIf cboFormaPgtoEntrada.Text = "6 - BOLETO" Then
       var_PGTO_Entrada = "BOLETO"
    ElseIf cboFormaPgtoEntrada.Text = "7 - TRANSFERÊNCIA" Then
       var_PGTO_Entrada = "TRANSFERENCIA"
    ElseIf cboFormaPgtoEntrada.Text = "8 - DEPOSITO" Then
       var_PGTO_Entrada = "DEPOSITO"
    ElseIf cboFormaPgtoEntrada.Text = "9 - FINANCEIRA" Then
       var_PGTO_Entrada = "FINANCEIRA"
    ElseIf cboFormaPgtoEntrada.Text = "10 - PIX" Then
       var_PGTO_Entrada = "PIX"
    Else
        var_PGTO_Entrada = "DINHEIRO"
    End If
End If

'quantidade de formas de pagamento
Dim varDivisaoPgto As String
varDivisaoPgto = cboQuantForma.Text

'calcular desconto em dinheiro ===================================
If optDescRS.Value = True Then
    If txtDesc.Text = "0,00" Then
        varValorRealDesc = FormatNumber(0, 2)
    Else
        varValorRealDesc = FormatNumber(txtDesc.Text, 2)
    End If
ElseIf optDescPorc.Value = True Then
    If txtDesc.Text = "0,00" Then
        varValorRealDesc = FormatNumber(0, 2)
    Else
        varValorRealDesc = FormatNumber(((txtSubtotal.Text * txtDesc.Text) / 100), 2)
    End If
End If

'calcular acrescimo em dinheiro==========================================================
If optAscrescRS.Value = True Then
    If txtAcresc.Text = "0,00" Then
        varValorRealAcresc = FormatNumber(0, 2)
    Else
        varValorRealAcresc = FormatNumber(CCur(txtAcresc.Text), 2)
    End If
ElseIf optAscrescPorc.Value = True Then
    If txtAcresc.Text = "0,00" Then
        varValorRealAcresc = FormatNumber(0, 2)
    Else
        varValorRealAcresc = FormatNumber(((CCur(txtSubtotal.Text) * CCur(txtAcresc.Text)) / 100), 2)
    End If
End If

'declarar quem receber os produtos
If vDeclararRecebedor = "SIM" Then
    Dim vRevebedor As String
    If ShowMsg("Deseja declarar o recebedor?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        vRevebedor = InputBox("Informe o nome do recebedor:", "ENTREGA DAS MERCADORIAS", "")
    Else
        vRevebedor = cboCliente.Text
    End If
    
    dbData.Execute "INSERT INTO pedidos_recebedor (cod_pedido, recebedor) VALUES (" & txtCodPedido.Text & ", '" & vRevebedor & "');"
    vRevebedor = ""
End If

'valor de troco
If txtRecebido.Text < txtTotalDesc.Text Then
   vValorRececido = "0.00"
   vValorTroco = "0.00"
Else
   vValorRececido = txtRecebido.Text
   vValorTroco = txtTroco.Text
End If

'DESCONTO - VARIAVEIS
If vDescItensVenda <> "0,00" Then
    Dim vSomaDescItens As Currency      'soma todos os descontos dos itens da venda em real
    Dim vValorDescVenda As Currency     'consulto quanto é para ser o valor do desconto em real
    Dim vValorSobraDesc As Currency     'valor que sobrou ou faltou no desconto
End If

If cboTipoPgto.Text = "À PRAZO" Then
    Dim var_Vencimento As Date
    Dim Var_NumParc As Integer
    Dim arrayParc() As Currency
    If txtCodCliente = "1" Then MsgBox "IDENTIFIQUE O CLIENTE DA COMPRA!", vbExclamation, "Aviso do sistema": Exit Sub
    'tabela configurações
    
    If bFechAP Then
       If ShowMsg("Deseja finalizar essa compra?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
    End If
    
    'função para checar o limite do cliente..
    If vLimitarCompra = "SIM" Then
        If txtCodCliente.Text <> "1" Then
             Verificar_Limite
        End If
    Else
        Passou_Limite = False
    End If
        'If Passou_Limite = True Then Exit Sub
           
    'Solicita autorização da gerência
    If txtCodCliente.Text <> "1" Then
         If Passou_Limite Or Cliente_Debito Then
            Dim fLib As LiberarVenda
            Dim bCancel As Boolean
            Dim lGerente As Long
            
            Set fLib = New LiberarVenda
            Load fLib
            
            fLib.Show vbModal
            bCancel = fLib.Cancelled
            lGerente = fLib.Gerente
                  
            Unload fLib
            Set fLib = Nothing
            
            If bCancel Then
               cboCliente.Text = ""
               txtCodCliente.Text = ""
               Exit Sub
            End If
            
            'Aqui voc seta o campo com o codigo do gerente que liberou
            'txtCodGerente = lGerente
            ''Else
            ''    varLiberarVendaDevedor = True
         End If
     End If
           
        ''If varLiberarVendaDevedor = True Then
        
    'colocar a data da Ultima compra de cada produro
    'For i = 1 To Grid.Rows - 1
    '   dbData.Execute "UPDATE produtos SET ult_compra = '" & Format$(Date, "yyyy-dd-MM") & "' WHERE (codigo = " & Grid.TextMatrix(i, 2) & ");"
    'Next
     
                 '"TROCO = " & Replace(CCur(txtTroco.Text), ",", ".") & ", " & _
                 '"RECEBIDO = " & Replace(CCur(txtRecebido.Text), ",", ".") & ", " & _

           'ATUALIZANDO A TABELA PEDIDOS
            sSQL = "UPDATE pedidos SET " & _
                 "cod_pedido = " & txtCodPedido.Text & ", " & _
                 "cod_cliente = " & txtCodCliente.Text & ", " & _
                 "data_compra = '" & Format$(txtDataCompra, "yyyy-dd-MM") & "', " & _
                 "tipo_desc = '" & IIf(optDescRS.Value = True, "R", "P") & "', " & _
                 "valor_desc = " & Replace(CCur(txtDesc.Text), ",", ".") & ", " & _
                 "ValorDescReal = " & Replace(CCur(varValorRealDesc), ",", ".") & ", " & _
                 "ValorAcrescReal = " & Replace(CCur(varValorRealAcresc), ",", ".") & ", " & _
                 "TIPO_ACRESCIMO = '" & IIf(optAscrescRS.Value = True, "R", "P") & "', " & _
                 "VALOR_ACRESCIMO = " & Replace(CCur(txtAcresc.Text), ",", ".") & ", " & _
                 "TROCO = " & Replace(CCur(vValorTroco), ",", ".") & ", " & _
                 "RECEBIDO = " & Replace(CCur(vValorRececido), ",", ".") & ", " & _
                 "entrada = " & Replace(CCur(txtEntrada.Text), ",", ".") & ", " & _
                 "subtotal = " & Replace(CCur(txtSubtotal.Text), ",", ".") & ", " & _
                 "total = " & Replace(CCur(txtTotalDesc.Text), ",", ".") & ", " & _
                 "tipo_pagamento = 'À Prazo', pagamento = '" & var_PAGAMENTO & "', tipo_cartao = " & varTipoCartao & ", " & _
                 "cod_funcionario = " & txtCodFuncAP.Text & ", " & _
                 "tipo_pedido = 'VENDA', " & _
                 "caixa = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "', " & _
                 "MAQUINA = '" & IIf(StatusBar1.Panels(4).Text = "", "PDV01", StatusBar1.Panels(4).Text) & "', " & _
                 "codcaixa = " & varCodCaixa & ", " & _
                 "status_pedido = 1 " & _
                 "WHERE (cod_pedido = " & txtCodPedido.Text & ");"
              dbData.Execute sSQL
           
           
           '********PARCELAS***********
           
           'COM ENTRADA =========================================================================
           If txtEntrada.Text <> "0,00" And txtValorParc.Text <> "0,00" Then
             
                'criar a entrada
                lNovoCod = Autonumeracao_Parcelas
              
                dbData.Execute "INSERT INTO parcelas (codigo, cod_pedido, numero, data, valor, status, TIPO, DIAS_ATRAZO, JUROS, MULTA, DESCONTO, COD_FUNCIONARIO, FORMA_PGTO) VALUES (" & _
                   lNovoCod & ", " & txtCodPedido.Text & ", 1, '" & Format$(txtDataCompra, "yyyy-dd-MM") & "', " & _
                   Replace(CCur(txtEntrada.Text), ",", ".") & ", 0, 'VENDA', 0, 0, 0, 0, " & txtCodFuncAP.Text & ", '" & var_PAGAMENTO & "');"
              
                'criar da segunda parcela em diante
                If vDataFlexivel = False Then
                  If cboPrazo.Text = "30" Then
                      var_Vencimento = Format(DateAdd("m", Val(1), mskInicio.Text), "dd/mm/yy")
                  Else
                      var_Vencimento = Format(DateAdd("d", Val(cboPrazo.Text), mskInicio.Text), "dd/mm/yy")
                  End If
                Else
                  var_Vencimento = Format(mskInicio.Text, "dd/mm/yy")
                End If
                
                Var_NumParc = 2
                
                CalcularParcelas (CCur(txtTotalDesc) - CCur(txtEntrada)), CInt(cboQuantParc), arrayParc
                
                For i = 1 To CInt(cboQuantParc)
                   lNovoCod = Autonumeracao_Parcelas
                   
                   dbData.Execute "INSERT INTO parcelas (codigo, cod_pedido, numero, data, valor, status, VALOR_FINAL, TIPO, DIAS_ATRAZO, JUROS, MULTA, DESCONTO, COD_FUNCIONARIO, forma_pgto) VALUES (" & _
                      lNovoCod & ", " & txtCodPedido.Text & ", " & Var_NumParc & ", '" & Format$(var_Vencimento, "yyyy-dd-MM") & "', " & _
                      Replace(arrayParc(i), ",", ".") & ", 0, " & Replace(arrayParc(i), ",", ".") & ", 'VENDA', 0, 0, 0, 0, " & txtCodFuncAP.Text & ", '" & var_PAGAMENTO & "');"
                   
                  If cboPrazo.Text = "30" Then
                      var_Vencimento = Format(DateAdd("m", Val(1), var_Vencimento), "dd/mm/yy")
                  Else
                      var_Vencimento = Format(DateAdd("d", Val(cboPrazo.Text), var_Vencimento), "dd/mm/yy")
                  End If
                   
                   Var_NumParc = Var_NumParc + 1
                Next
              
            
            ElseIf txtEntrada.Text = "0,00" And txtValorParc.Text <> "0,00" Then     'SEM ENTRADA ==
                
                'parcelas
                var_Vencimento = CDate(mskInicio.Text)
                Var_NumParc = 1
                
                CalcularParcelas CCur(txtTotalDesc), CInt(cboQuantParc), arrayParc
                
                'criar as parcelas
                For i = 1 To CInt(cboQuantParc)
                   lNovoCod = Autonumeracao_Parcelas
                   
                   dbData.Execute "INSERT INTO parcelas (codigo, cod_pedido, numero, data, valor, status, VALOR_FINAL, TIPO, DIAS_ATRAZO, JUROS, MULTA, DESCONTO, COD_FUNCIONARIO, FORMA_PGTO) VALUES (" & _
                      lNovoCod & ", " & txtCodPedido.Text & ", " & Var_NumParc & ", '" & Format$(var_Vencimento, "yyyy-dd-MM") & "', " & _
                      Replace(arrayParc(i), ",", ".") & ", 0, " & Replace(arrayParc(i), ",", ".") & ", 'VENDA', 0, 0, 0, 0, " & txtCodFuncAP.Text & ", '" & var_PAGAMENTO & "');"
                   
                   'var_Vencimento = Format(DateAdd("m", Val(1), var_Vencimento), "dd/mm/yy")
                  If cboPrazo.Text = "30" Then
                      var_Vencimento = Format(DateAdd("m", Val(1), var_Vencimento), "dd/mm/yy")
                  Else
                      'var_Vencimento = Format(DateAdd("m", Val(1), mskInicio.Text), "dd/mm/yy")
                      var_Vencimento = Format(DateAdd("d", Val(cboPrazo.Text), var_Vencimento), "dd/mm/yy")
                  End If
                  
                   Var_NumParc = Var_NumParc + 1
                Next
           End If
           
            'verificar duplicidade de parcelas
            sSQL = "SELECT CODIGO, COD_PEDIDO, NUMERO FROM parcelas " & _
                   "WHERE (cod_pedido = " & txtCodPedido.Text & ") AND (numero = 1);"
            Set r = dbData.OpenRecordset(sSQL)
            
            'Dim vSomaDescItens As Currency
            If Not r.EOF Then
                If r.RecordCount = 2 Then
                    'MsgBox "tem 2"
                    dbData.Execute "DELETE FROM parcelas WHERE (CODIGO =(SELECT MAX(CODIGO) FROM parcelas WHERE (cod_pedido = " & txtCodPedido.Text & ") AND (numero = 1)) );"
                End If
            End If
           
            'compra com estorno pega a data e hora do estorno
            If lblEstornar.Caption = "ESTORNO" Then
                If txtHoraCompra.Text = "00:00" Or txtHoraCompra.Text = "" Then
                    varHora = Format(Now, ocHORA)
                Else
                    varHora = Format(txtHoraCompra, ocHORA)
                End If
            Else
                varHora = Format(Now, ocHORA)
            End If

           'dar baixa na parcela de entrada ou compra à vista
           If lblEstornar.Caption = "ESTORNO" Then
                If txtEntrada.Text <> "0,00" Then
                   dbData.Execute "UPDATE parcelas SET " & _
                      "status = 1, valor_final = " & Replace(CCur(txtEntrada.Text), ",", ".") & ", " & _
                      "pagamento = '" & Format$(txtDataCompra, "yyyy-dd-MM") & "', " & _
                      "hora = '" & Format(varHora, ocHORA) & "', " & _
                      "forma_pgto = '" & var_PGTO_Entrada & "', " & _
                      "DIAS_ATRAZO = 0, " & _
                      "JUROS = 0, " & _
                      "MULTA = 0, " & _
                      "DESCONTO = 0, " & _
                      "tipo = 'PARCELA', tipo_cartao = " & varTipoCartaoEntrada & ", " & _
                      "CODCAIXA = " & varCodCaixa & ", " & _
                      "caixa = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "', COD_FUNCIONARIO = " & txtCodFuncAP.Text & " " & _
                      "WHERE (cod_pedido = " & txtCodPedido.Text & ") AND (numero = 1);"
                End If
                    
                'dar baixa nas parcelas de de cartão
                If cboFormaPgto.Text = "3 - CARTÃO - DÉBITO" Or cboFormaPgto.Text = "4 - CARTÃO - CRÉDITO" Then
                   dbData.Execute "update parcelas set pagamento = '" & Format$(txtDataCompra, "yyyy-dd-MM") & "', Status = 1, valor_final = VALOR, hora = '" & Format(varHora, ocHORA) & "', forma_pgto = 'CARTAO', caixa = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "', CODCAIXA = " & varCodCaixa & ", DIAS_ATRAZO = 0, JUROS = 0, MULTA = 0, DESCONTO = 0, COD_FUNCIONARIO = " & txtCodFuncAP.Text & "  WHERE (cod_pedido = " & txtCodPedido.Text & ")"
                End If
           Else
                If txtEntrada.Text <> "0,00" Then
                   dbData.Execute "UPDATE parcelas SET " & _
                      "status = 1, valor_final = " & Replace(CCur(txtEntrada.Text), ",", ".") & ", " & _
                      "pagamento = '" & Format$(txtDataCompra, "yyyy-dd-MM") & "', " & _
                      "hora = '" & Format(Now, ocHORA) & "', " & _
                      "forma_pgto = '" & var_PGTO_Entrada & "', " & _
                      "DIAS_ATRAZO = 0, " & _
                      "JUROS = 0, " & _
                      "MULTA = 0, " & _
                      "DESCONTO = 0, " & _
                      "tipo = 'PARCELA', tipo_cartao = " & varTipoCartaoEntrada & ", " & _
                      "CODCAIXA = " & varCodCaixa & ", " & _
                      "caixa = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "', COD_FUNCIONARIO = " & txtCodFuncAP.Text & " " & _
                      "WHERE (cod_pedido = " & txtCodPedido.Text & ") AND (numero = 1);"
                End If
                
                'dar baixa nas parcelas de de cartão
                If cboFormaPgto.Text = "3 - CARTÃO - DÉBITO" Or cboFormaPgto.Text = "4 - CARTÃO - CRÉDITO" Then
                   dbData.Execute "update parcelas set pagamento = '" & Format$(txtDataCompra, "yyyy-dd-MM") & "', Status = 1, valor_final = VALOR, hora = '" & Format(Now, ocHORA) & "', forma_pgto = 'CARTAO', caixa = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "', CODCAIXA = " & varCodCaixa & ", DIAS_ATRAZO = 0, JUROS = 0, MULTA = 0, DESCONTO = 0, COD_FUNCIONARIO = " & txtCodFuncAP.Text & " WHERE (cod_pedido = " & txtCodPedido.Text & ")"
                End If
           End If
           
           txtHoraCompra.Text = ""
           
           'Colocando a data da ultima compra
           'execSQL "UPDATE CLIENTE SET Ultima_Compra = #" & Format(Date, "MM/dd/yyyy") & "# WHERE CODIGO = " & txtCodCliente.Text

        'calcular subtotal de cada item
        'sSQL = "UPDATE pedidos_itens SET subtotal = preco * quantidade where (cod_pedido = " & txtCodPedido.Text & ")"
        'dbData.Execute sSQL
    
        'calcular desconto de cada item
        If vDescItensVenda <> "0,00" Then
            'adiciona em cada item do pedido o valor do desconto
            sSQL = "UPDATE pedidos_itens SET desconto = (subtotal * " & Replace(CDbl(vDescItensVenda), ",", ".") & " / 100), total = subtotal - (subtotal * " & Replace(CDbl(vDescItensVenda), ",", ".") & " / 100), data = '" & Format$(txtDataCompra, "yyyy-dd-MM") & "' where (cod_pedido = " & txtCodPedido.Text & ")"
            dbData.Execute sSQL
            
            'soma todos os descontos dos itens da venda em real
            sSQL = "SELECT SUM(Desconto) AS varSomaDescItens FROM pedidos_itens WHERE (cod_pedido = " & txtCodPedido.Text & ")"
            Set r = dbData.OpenRecordset(sSQL)
            
            'Dim vSomaDescItens As Currency
            If Not r.EOF Then
                vSomaDescItens = FormatNumber(ValidateNull(r("varSomaDescItens")), 2)
            End If
            
            'consulto quanto é para ser o valor do desconto em real
            sSQL = "SELECT ValorDescReal FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ")"
            Set r = dbData.OpenRecordset(sSQL)
            
            'Dim vValorDescVenda As Currency
            If Not r.EOF Then
                vValorDescVenda = FormatNumber(ValidateNull(r("ValorDescReal")), 2)
            End If
            
            'se o valor total do desconto for maior que a soma dos desconto dos itens da venda
            If vValorDescVenda < vSomaDescItens Then
                vValorSobraDesc = CCur(vSomaDescItens - vValorDescVenda)
                sSQL = "UPDATE pedidos_itens SET Desconto = Desconto - " & Replace(CCur(vValorSobraDesc), ",", ".") & ", Total = Total + " & Replace(CCur(vValorSobraDesc), ",", ".") & " " & _
                        "WHERE (CODIGO = " & _
                "(SELECT MAX(CODIGO) FROM pedidos_itens WHERE (cod_pedido = " & txtCodPedido.Text & ")))"
                dbData.Execute sSQL
            ElseIf vValorDescVenda > vSomaDescItens Then
                vValorSobraDesc = CCur(vValorDescVenda - vSomaDescItens)
                sSQL = "UPDATE pedidos_itens SET Desconto = Desconto + " & Replace(CCur(vValorSobraDesc), ",", ".") & ", Total = Total - " & Replace(CCur(vValorSobraDesc), ",", ".") & " " & _
                        "WHERE (CODIGO = " & _
                "(SELECT MAX(CODIGO) FROM pedidos_itens WHERE (cod_pedido = " & txtCodPedido.Text & ")))"
                dbData.Execute sSQL
            End If
        Else
            If lblEstornar.Caption = "ESTORNO" Then
                sSQL = "UPDATE pedidos_itens SET Desconto = '0.00', Total = Subtotal " & _
                        "WHERE (cod_pedido = " & txtCodPedido.Text & ")"
                dbData.Execute sSQL
            End If
        End If
        
        'Retirar da tabela PRODUTOS as QUANTIDADES mencionadas no grid
        For i = 1 To Grid.Rows - 1
           dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque - " & Replace(CDbl(Grid.TextMatrix(i, 5)), ",", ".") & " WHERE (codigo = " & Grid.TextMatrix(i, 2) & ");"
        Next
        
        'CASHBACK
        If txtCodCliente.Text <> "1" Then
            If vCashbackAP = "SIM" Then
                'Cashback dar baixa
                If vUsandoCashBack = True Then
                    dbData.Execute "UPDATE Pedidos_Cashback SET VALOR_ABATIDO = VALOR_CASHBACK, ABATIDO = 1, DATA_ABATIDO = '" & Format$(Date, "yyyy-dd-MM") & "', COD_PEDIDOABATIDO = " & txtCodPedido.Text & ", COD_FUNCIONARIO = " & txtCodFuncAP.Text & " WHERE (COD_CLIENTE = " & txtCodCliente.Text & ") and ABATIDO = 0 and INVALIDO = 0;"
                End If
                
                'criar novo cashback
                'valor do cashback
                vValorVenda = CCur(txtTotalDesc.Text)
                ValorCash = (vValorVenda * vCashbackValorAP) / 100
                
                'validade do cashback
                vCashbackValidade = Format(DateAdd("d", Val(vCashbackLimite), Date), "dd/mm/yy")
                
                    lNovoCod = Autonumeracao_Cashback
                    sSQL = "INSERT INTO Pedidos_Cashback (CODIGO, COD_PEDIDO, VALOR_VENDA, VALOR_CASHBACK, VALOR_ABATIDO, ABATIDO, VALIDADE, INVALIDO, COD_FUNCIONARIO, COD_CLIENTE) VALUES (" & _
                    lNovoCod & ", " & txtCodPedido.Text & ", " & Replace(CCur(txtTotalDesc.Text), ",", ".") & ", " & Replace(CCur(ValorCash), ",", ".") & ", 0, 0, '" & Format$(vCashbackValidade, "yyyy-dd-MM") & "', 0, " & txtCodFuncAP.Text & ", " & txtCodCliente.Text & ");"
                    dbData.Execute sSQL
            End If
        End If
    
        'inicio do CUPOM FISCAL
        If vConfImprimeNFCeLocal = "SIM" Then
            If NFCe_OK = True Then
            
                    'verifica se já existe um cupom emitdo para esse pedido
                    sSQL2 = "SELECT Num_OS_VD_Origem FROM TbNFCe WHERE (Num_OS_VD_Origem = " & txtCodPedido & ");"
                    Set r2 = dbData.OpenRecordset(sSQL2)
                    
                    If Not r2.BOF Then
                        MsgBox "NFCe para esse pedido já foi criada"
                        GoTo SemGerarNFCe
                    Else  'se nao existir nfce
           
                        'consultar o cpf/cnpj do cliente
                        sSQL2 = "SELECT CPF, CODIGO FROM cliente WHERE (codigo = " & txtCodCliente.Text & ");"
                        Set rCliente = dbData.OpenRecordset(sSQL2)
                        
                        If Not rCliente.EOF Then
                            vCPF = RetirarMascaras(ValidateNull(rCliente!CPF))
                            'vCPF = rCliente!CPF
                        Else
                            vCPF = ""
                        End If
                        
                        'validar CPF
                        Select Case Len(vCPF)
                            Case 0
                                If Len(vCPF) = 0 Then
                                    vCPF = Empty
                                Else
                                    vCPF = ""
                                End If
                                'KeyCode = 0
                            Case 14
CNPJDigitadoErrado:
                                If Validar_CNPJ(vCPF) = False Then
                                        If vNFCeConfCPF = "SIM" Then
                                            If ShowMsg("Deseja inserir o CNPJ no NFCe?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                                                vCPF = InputBox("Informe o CNPJ do cliente:", "EMISSÃO DE NFCe", "")
                                                If Not Vazio(vCPF) Then
                                                    If Len(vCPF) = 11 Then
                                                        If Validar_CPF(vCPF) = False Then
                                                            MsgBox "CPF Informado não é valido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
                                                            GoTo CPFDigitadoErrado
                                                        Else
                                                            vCPF = Format(vCPF, "000\.000\.000\-00")
                                                        End If
                                                    ElseIf Len(vCPF) = 14 Then
                                                        If Validar_CNPJ(vCPF) = False Then
                                                            MsgBox "CNPJ Informado não é valido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
                                                            GoTo CNPJDigitadoErrado
                                                        Else
                                                            vCPF = Format(vCPF, "00\.000\.000\/0000\-00")
                                                        End If
                                                    End If
                                                Else
                                                    vCPF = ""
                                                End If
                                            Else
                                                vCPF = ""           'se na msgbox colocar NÃO quer colocar cpf
                                            End If
                                        End If
                                End If
                            Case 11
CPFDigitadoErrado:
                                If Validar_CPF(vCPF) = False Then
                                        If vNFCeConfCPF = "SIM" Then
                                            If ShowMsg("Deseja inserir o CPF no NFCe?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                                                vCPF = InputBox("Informe o CPF do cliente:", "EMISSÃO DE NFCe", "")
                                                If Not Vazio(vCPF) Then
                                                    If Len(vCPF) = 11 Then
                                                        If Validar_CPF(vCPF) = False Then
                                                            MsgBox "CPF Informado não é valido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
                                                            GoTo CPFDigitadoErrado
                                                        Else
                                                            vCPF = Format(vCPF, "000\.000\.000\-00")
                                                        End If
                                                    ElseIf Len(vCPF) = 14 Then
                                                        If Validar_CNPJ(vCPF) = False Then
                                                            MsgBox "CNPJ Informado não é valido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
                                                            GoTo CNPJDigitadoErrado
                                                        Else
                                                            vCPF = Format(vCPF, "00\.000\.000\/0000\-00")
                                                        End If
                                                    End If
                                                Else
                                                    vCPF = ""       'se o cpf for vazio
                                                End If
                                            Else
                                                vCPF = ""           'se na msgbox colocar NÃO quer colocar cpf
                                            End If
                                        End If
                                End If
                            Case Is < 11
                                'MsgBox "CPF Informado não é valido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
                                'mskCNPJ.SetFocus
                        End Select
                       
                        If Len(vCPF) = 11 Then
                            vCPF = Format(vCPF, "000\.000\.000\-00")
                        ElseIf Len(vCPF) = 14 Then
                            vCPF = Format(vCPF, "00\.000\.000\/0000\-00")
                        Else
                            vCPF = ""
                        End If
                        
                        'coloca o cpf correto no cliente
                        dbData.Execute "UPDATE cliente SET cpf = '" & vCPF & "' WHERE (codigo = " & txtCodCliente.Text & ")"
                        
                        'preenche a tabela de itens da NFCe
                        sSQL = "EXEC NFCeIncluir " & txtCodPedido.Text
                        dbData.Execute sSQL
                    End If
                    
                    If rCliente!Codigo = 1 Then
                        dbData.Execute "UPDATE cliente SET cpf = '' WHERE (codigo = 1)"
                    End If
                    
                    If rCliente.State <> 0 Then rCliente.Close
                    Set rCliente = Nothing
                    
                    If r2.State <> 0 Then r2.Close
                    Set r2 = Nothing
                    
                    'criando as parcelas
                    sSQL = "SELECT IdNFProd FROM TbNFCe WHERE Num_OS_VD_Origem  = " & txtCodPedido.Text
                    Set rNFCe = dbData.OpenRecordset(sSQL)
                    
                    If rNFCe.RecordCount > 0 Then
                       sSQL = "INSERT INTO [TbNFCe_Faturas] " & _
                              "([IdNFProd] " & _
                              ",[IDParcela] " & _
                              ",[TipoPgto] " & _
                              ",[Vencimento] " & _
                              ",[Valor] " & _
                              ",[IdBandeira] " & _
                              ",[CartaoNumeroAutorizacao]) " & _
                              "SELECT " & rNFCe!IdNFProd & " " & _
                              "     ,NUMERO " & _
                              "     ,dbo.NFCeFormaPagto(FORMA_PGTO, TIPO_CARTAO) " & _
                              "     ,DATA " & _
                              "     ,VALOR " & _
                              "     ,'01' " & _
                              "     ,'' " & _
                              "FROM [parcelas] " & _
                              "WHERE COD_PEDIDO = " & txtCodPedido.Text
                       dbData.Execute sSQL
                       
                       'verificando os itens do pedido
                       sSQL = "SELECT IdNFProd, IdNFProd_Item, IDProduto, CodBarras, DescricaoProduto, CodNcm, CFOP, Bc_Icms, ICMSCST, IPICST, COFINSCST, PISCST, UN " & _
                              "FROM TbNFCe_Itens " & _
                              "WHERE (IdNFProd = " & rNFCe!IdNFProd & ");"
                       Set rNFCeItens = dbData.OpenRecordset(sSQL)
                       
                       'Dim EncontroErroNFCe As Boolean
                       EncontroErroNFCe = False
                       
                        For i = 1 To rNFCeItens.RecordCount
                            
                            'NCM..........
                            If rNFCeItens!CodBarras <> "SEM GTIN" Then
                                If Len(rNFCeItens!CodBarras) > 13 Or Len(rNFCeItens!CodBarras) < 8 Then
                                    EncontroErroNFCe = True
                                Else
                                    EncontroErroNFCe = False
                                End If
                            Else
                                EncontroErroNFCe = False
                            End If
                            
                            If EncontroErroNFCe = True Then GoTo Continuar
                                                        
                            'CFOP..........
                            If rNFCeItens!CFOP <> Empty Then
                                If Len(rNFCeItens!CFOP) > 4 Or Len(rNFCeItens!CFOP) < 4 Then
                                    EncontroErroNFCe = True
                                Else
                                    EncontroErroNFCe = False
                                End If
                            Else
                                EncontroErroNFCe = False
                            End If
                            
                            If EncontroErroNFCe = True Then GoTo Continuar
                            
                            'ICMS CST..........
                            If rNFCeItens!icmsCST <> Empty Then
                                If Len(rNFCeItens!icmsCST) > 3 Or Len(rNFCeItens!icmsCST) < 3 Then
                                    EncontroErroNFCe = True
                                Else
                                    EncontroErroNFCe = False
                                End If
                            Else
                                EncontroErroNFCe = False
                            End If
                            
                            If EncontroErroNFCe = True Then GoTo Continuar

                            'PIS CST..........
                            If rNFCeItens!pisCST <> Empty Then
                                If Len(rNFCeItens!pisCST) > 2 Or Len(rNFCeItens!pisCST) < 2 Then
                                    EncontroErroNFCe = True
                                Else
                                    EncontroErroNFCe = False
                                End If
                            Else
                                EncontroErroNFCe = False
                            End If
                            
                            If EncontroErroNFCe = True Then GoTo Continuar

                            'COFINS CST..........
                            If rNFCeItens!cofinsCST <> Empty Then
                                If Len(rNFCeItens!cofinsCST) > 2 Or Len(rNFCeItens!cofinsCST) < 2 Then
                                    EncontroErroNFCe = True
                                Else
                                    EncontroErroNFCe = False
                                End If
                            Else
                                EncontroErroNFCe = False
                            End If
                            
                            If EncontroErroNFCe = True Then GoTo Continuar
                            
                            'NCM..........
                            If rNFCeItens!CodNcm <> Empty Then
                                If Len(rNFCeItens!CodNcm) > 8 Or Len(rNFCeItens!CodNcm) < 8 Then
                                    EncontroErroNFCe = True
                                Else
                                    EncontroErroNFCe = False
                                End If
                            Else
                                EncontroErroNFCe = False
                            End If
                            
                            If EncontroErroNFCe = True Then GoTo Continuar
                            
                            'UNIDADE DE MEDIDA..........
                            If rNFCeItens!UN <> Empty Then
                                If Len(rNFCeItens!UN) > 2 Or Len(rNFCeItens!UN) < 1 Then
                                    EncontroErroNFCe = True
                                Else
                                    EncontroErroNFCe = False
                                End If
                            Else
                                EncontroErroNFCe = False
                            End If
                            
                            If EncontroErroNFCe = True Then GoTo Continuar
                        
                        rNFCeItens.MoveNext
                        Next
                       
Continuar:
                'transmitir o cupom fiscal - NFCe
                If EncontroErroNFCe = False Then
                       DoEvents
                       iRetorno = TransmitirNFCe(rNFCe!IdNFProd, "1", Not NFCeContingencia, "65")
                       
                       If iRetorno Then
                          Set sistNFe = New snfe.Util
                          
                          iRetorno = ConfiguraDLLNFeNFCe(65, "1", sistNFe)
                            If vNFCeImprimir = "SIM" Then
                                If Not NFCeContingencia Then
                                   Call sistNFe.DANFCeImprimir(xCaminhoXML, True, var_ImpNFCe, True, xCaminhoPDF, 0, False, False, "")
                                Else
                                   Call sistNFe.DANFCeOFFImprimir(xCaminhoXML, True, var_ImpNFCe, True, xCaminhoPDF, 0, False, False, "")
                                End If
                            End If
                       End If
                    'End If
                Else
                    If vNFCeCombinarImp = "NÃO" Then
                        If NFCe_OK = True Then
                            ImprimirVendaAPsemPergunta
                        Else
                            ImprimirVendaAP
                        End If
                        'ImprimirVendaAP
                    End If
                End If
                    
                    'dbData.Execute "UPDATE cliente SET cpf = '' WHERE (codigo = 1)"

                'If vConfImprimeNFCeLocal = "SIM" Then   'SE no arquivo ini tem SIM para imprimir NFCE
                        'If NFCe_OK = True Then              'SE dei SIM para imprimir NFCE
                            If vNFCeCombinarImp = "SIM" Then
                                 ImprimirVendaAP
                            End If              'final de vNFCeCombinarImp = "SIM"
                        Else                    'meio do NFCe_OK = false
                            ImprimirVendaAP
                        End If                  'fim do NFCe_OK = True
                Else                            'meio do vConfImprimeNFCeLocal = "SIM"
                    ImprimirVendaAP
                End If
    Else        'MEIO vConfImprimeNFCeLocal = "SIM"
SemGerarNFCe:
        ImprimirVendaAP
    End If      'FIM vConfImprimeNFCeLocal = "SIM"

           dbData.Execute "UPDATE Pedidos_Reabertura SET STATUS_PEDIDO = 1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ") AND (DATA = (SELECT MAX(DATA) FROM Pedidos_Reabertura AS Pedidos_Reabertura_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & "))) AND (HORA = (SELECT MAX(HORA) FROM Pedidos_Reabertura AS Pedidos_Reabertura_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")));"

            LimparObjetos_Pedido
            txtTotalGeral.Text = ""
            LimparGrid_Pedido
            LimparObjetos_Prazo
            txtCodPedido.Text = ""
            lblEstornar.Caption = ""
            'frmVendaFechamento.Visible = False
            frmVendaFechamento.Visible = False
            txtDataCompra.Text = Format(Date, "dd/mm/yyyy")
            CriarNovoPedido
            If varTipoValorVenda = 1 Then txtCodBarra.SetFocus

            
            
            
            
ElseIf cboTipoPgto.Text = "À VISTA" Then
           If bFechAV Then
              If ShowMsg("Deseja finalizar essa compra?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
           End If
           
           'ATUALIZANDO A TABELA PEDIDOS
           sSQL = "UPDATE pedidos SET " & _
              "cod_pedido = " & txtCodPedido.Text & ", " & _
              "cod_cliente = " & txtCodCliente.Text & ", " & _
              "data_compra = '" & Format$(txtDataCompra, "yyyy-dd-MM") & "', " & _
              "tipo_desc = '" & IIf(optDescRS.Value = True, "R", "P") & "', " & _
              "valor_desc = " & Replace(CCur(txtDesc.Text), ",", ".") & ", " & _
              "ValorDescReal = " & Replace(CCur(varValorRealDesc), ",", ".") & ", " & _
              "ValorAcrescReal = " & Replace(CCur(varValorRealAcresc), ",", ".") & ", " & _
              "TIPO_ACRESCIMO = '" & IIf(optAscrescRS.Value = True, "R", "P") & "', " & _
              "VALOR_ACRESCIMO = " & Replace(CCur(txtAcresc.Text), ",", ".") & ", " & _
              "TROCO = " & Replace(CCur(vValorTroco), ",", ".") & ", " & _
              "RECEBIDO = " & Replace(CCur(vValorRececido), ",", ".") & ", " & _
              "subtotal = " & Replace(CCur(txtSubtotal.Text), ",", ".") & ", " & _
              "total = " & Replace(CCur(txtTotalDesc.Text), ",", ".") & ", " & _
              "tipo_pagamento = 'À Vista', pagamento = '" & varDivisaoPgto & "', tipo_cartao = " & varTipoCartao & ", " & _
              "cod_funcionario = " & txtCodFuncAP.Text & ",  " & _
              "tipo_pedido = 'VENDA', caixa = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "', maquina = '" & IIf(StatusBar1.Panels(4).Text = "", "PDV01", StatusBar1.Panels(4).Text) & "', codcaixa = " & varCodCaixa & ", " & _
              "status_pedido = 1 " & _
              "WHERE (cod_pedido = " & txtCodPedido.Text & ");"
           dbData.Execute sSQL
           
           '===========================================CRIAR E DAR BAIXA EM PARCELAS ==========================================
                        If cboTipoPgto.Text = "À VISTA" And cboQuantForma.Text = "1 - FORMA" Then
                            'autonumeração das parcelas
                            lNovoCod = Autonumeracao_Parcelas
                            
                            'Criando as Parcelas
                            sSQL = "INSERT INTO parcelas (codigo, cod_pedido, numero, data, valor, DIAS_ATRAZO, JUROS, MULTA, DESCONTO, COD_FUNCIONARIO) VALUES (" & _
                               lNovoCod & ", " & txtCodPedido.Text & ", 1, '" & Format$(txtDataCompra, "yyyy-dd-MM") & "', " & Replace(CCur(txtTotalDesc.Text), ",", ".") & ", 0, 0, 0, 0, " & txtCodFuncAP.Text & ");"
                            dbData.Execute sSQL
                            
                            'compra com estorno pega a data e hora do estorno
                            If lblEstornar.Caption = "ESTORNO" Then
                                If txtHoraCompra.Text = "00:00" Then
                                    varHora = Format(Now, ocHORA)
                                Else
                                    varHora = Format(txtHoraCompra, ocHORA)
                                End If
                            Else
                                varHora = Format(Now, ocHORA)
                            End If
                            
                            'DAR BAIXA NA PARCELA =====
                            sSQL = "UPDATE parcelas SET " & _
                            "status = 1, " & _
                            "valor_final = " & Replace(CCur(txtTotalDesc.Text), ",", ".") & ", " & _
                            "pagamento = '" & Format$(txtDataCompra, "yyyy-dd-MM") & "', " & _
                            "hora = '" & Format(varHora, ocHORA) & "', " & _
                            "forma_pgto = '" & var_PAGAMENTO & "', " & _
                            "DIAS_ATRAZO = 0, " & _
                            "JUROS = 0, " & _
                            "MULTA = 0, " & _
                            "DESCONTO = 0, " & _
                            "tipo = 'VENDA', " & _
                            "tipo_cartao = " & varTipoCartao & ", " & _
                            "CODCAIXA = " & varCodCaixa & ", " & _
                            "caixa = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "', COD_FUNCIONARIO = " & txtCodFuncAP.Text & " " & _
                            "WHERE (cod_pedido = " & txtCodPedido.Text & ") AND (numero = 1);"
                               
                            dbData.Execute sSQL
                            txtHoraCompra.Text = ""
                        
                        ElseIf cboTipoPgto.Text = "À VISTA" And cboQuantForma.Text = "2 - FORMAS" Then
                            'compra com estorno pega a data e hora do estorno
                            If lblEstornar.Caption = "ESTORNO" Then
                                If txtHoraCompra.Text = "00:00" Or txtHoraCompra.Text = "" Then
                                    varHora = Format(Now, ocHORA)
                                Else
                                    varHora = Format(txtHoraCompra, ocHORA)
                                End If
                            Else
                                varHora = Format(Now, ocHORA)
                            End If
                            
                            'autonumeração das parcelas - PARCELA 1
                            lNovoCod = Autonumeracao_Parcelas
                            
                            'Criando as Parcelas - PARCELA 1
                            If txtEntrada.Text <> "0,00" Then
                                sSQL = "INSERT INTO parcelas (codigo, cod_pedido, numero, data, valor, DIAS_ATRAZO, JUROS, MULTA, DESCONTO, COD_FUNCIONARIO, STATUS, valor_final, pagamento, hora ) VALUES (" & _
                                   lNovoCod & ", " & txtCodPedido.Text & ", 1, '" & Format$(txtDataCompra, "yyyy-dd-MM") & "', " & Replace(CCur(txtEntrada.Text), ",", ".") & ", 0, 0, 0, 0, " & txtCodFuncAP.Text & ", 1, " & Replace(CCur(txtEntrada.Text), ",", ".") & ", '" & Format$(txtDataCompra, "yyyy-dd-MM") & "', '" & Format(varHora, ocHORA) & "' );"
                                dbData.Execute sSQL
                            End If
                            
                                sSQL = "UPDATE parcelas SET " & _
                                "status = 1, " & _
                                "valor_final = " & Replace(CCur(txtEntrada.Text), ",", ".") & ", " & _
                                "pagamento = '" & Format$(txtDataCompra, "yyyy-dd-MM") & "', " & _
                                "hora = '" & Format(varHora, ocHORA) & "', " & _
                                "forma_pgto = '" & var_PGTO_Entrada & "', " & _
                                "DIAS_ATRAZO = 0, " & _
                                "JUROS = 0, " & _
                                "MULTA = 0, " & _
                                "DESCONTO = 0, " & _
                                "tipo = 'VENDA', " & _
                                "tipo_cartao = " & varTipoCartaoEntrada & ", " & _
                                "CODCAIXA = " & varCodCaixa & ", " & _
                                "caixa = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "', COD_FUNCIONARIO = " & txtCodFuncAP.Text & " " & _
                                "WHERE (cod_pedido = " & txtCodPedido.Text & ") AND (numero = 1);"
                                dbData.Execute sSQL
                        
                            'autonumeração das parcelas- PARCELA 2
                            lNovoCod = Autonumeracao_Parcelas
                            
                            'Criando as Parcelas - PARCELA 2
                            If txtValorRest.Text <> "0,00" Then
                                sSQL = "INSERT INTO parcelas (codigo, cod_pedido, numero, data, valor, DIAS_ATRAZO, JUROS, MULTA, DESCONTO, COD_FUNCIONARIO) VALUES (" & _
                                   lNovoCod & ", " & txtCodPedido.Text & ", 2, '" & Format$(txtDataCompra, "yyyy-dd-MM") & "', " & Replace(CCur(txtValorRest.Text), ",", ".") & ", 0, 0, 0, 0, " & txtCodFuncAP.Text & ");"
                                dbData.Execute sSQL
                            End If
                                      

                            'DAR BAIXA NA PARCELA =====  'parcela 1
                            If txtEntrada.Text <> "0,00" Then
                                sSQL = "UPDATE parcelas SET " & _
                                "status = 1, " & _
                                "valor_final = " & Replace(CCur(txtEntrada.Text), ",", ".") & ", " & _
                                "pagamento = '" & Format$(txtDataCompra, "yyyy-dd-MM") & "', " & _
                                "hora = '" & Format(varHora, ocHORA) & "', " & _
                                "forma_pgto = '" & var_PGTO_Entrada & "', " & _
                                "DIAS_ATRAZO = 0, " & _
                                "JUROS = 0, " & _
                                "MULTA = 0, " & _
                                "DESCONTO = 0, " & _
                                "tipo = 'VENDA', " & _
                                "tipo_cartao = " & varTipoCartaoEntrada & ", " & _
                                "CODCAIXA = " & varCodCaixa & ", " & _
                                "caixa = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "', COD_FUNCIONARIO = " & txtCodFuncAP.Text & " " & _
                                "WHERE (cod_pedido = " & txtCodPedido.Text & ") AND (numero = 1);"
                                dbData.Execute sSQL
                            End If
                        
                            'DAR BAIXA NA PARCELA =====  'parcela 2
                            If txtValorRest.Text <> "0,00" Then
                                sSQL = "UPDATE parcelas SET " & _
                                "status = 1, " & _
                                "valor_final = " & Replace(CCur(txtValorRest.Text), ",", ".") & ", " & _
                                "pagamento = '" & Format$(txtDataCompra, "yyyy-dd-MM") & "', " & _
                                "hora = '" & Format(varHora, ocHORA) & "', " & _
                                "forma_pgto = '" & var_PAGAMENTO & "', " & _
                                "DIAS_ATRAZO = 0, " & _
                                "JUROS = 0, " & _
                                "MULTA = 0, " & _
                                "DESCONTO = 0, " & _
                                "tipo = 'VENDA', " & _
                                "tipo_cartao = " & varTipoCartao & ", " & _
                                "CODCAIXA = " & varCodCaixa & ", " & _
                                "caixa = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "', COD_FUNCIONARIO = " & txtCodFuncAP.Text & " " & _
                                "WHERE (cod_pedido = " & txtCodPedido.Text & ") AND (numero = 2);"
                                dbData.Execute sSQL
                            End If
                            
                            txtHoraCompra.Text = ""
                        End If      'fim das opções de 1 e 2 parcelas
                        
            'verificar duplicidade de parcelas
            sSQL = "SELECT CODIGO, COD_PEDIDO, NUMERO FROM parcelas " & _
                   "WHERE (cod_pedido = " & txtCodPedido.Text & ") AND (numero = 1);"
            Set r = dbData.OpenRecordset(sSQL)
            
            'Dim vSomaDescItens As Currency
            If Not r.EOF Then
                If r.RecordCount = 2 Then
                    dbData.Execute "DELETE FROM parcelas WHERE (CODIGO =(SELECT MAX(CODIGO) FROM parcelas WHERE (cod_pedido = " & txtCodPedido.Text & ") AND (numero = 1)) );"
                End If
            End If
           
           'autonumeração das parcelas
        
        'calcular desconto de cada item
        If vDescItensVenda <> "0,00" Then
            'adiciona em cada item do pedido o valor do desconto
            sSQL = "UPDATE pedidos_itens SET desconto = (subtotal * " & Replace(CDbl(vDescItensVenda), ",", ".") & " / 100), total = subtotal - (subtotal * " & Replace(CDbl(vDescItensVenda), ",", ".") & " / 100), data = '" & Format$(txtDataCompra, "yyyy-dd-MM") & "' where (cod_pedido = " & txtCodPedido.Text & ")"
            dbData.Execute sSQL
            
            'soma todos os descontos dos itens da venda em real
            sSQL = "SELECT SUM(Desconto) AS varSomaDescItens FROM pedidos_itens WHERE (cod_pedido = " & txtCodPedido.Text & ")"
            Set r = dbData.OpenRecordset(sSQL)
            
            If Not r.EOF Then
                vSomaDescItens = FormatNumber(ValidateNull(r("varSomaDescItens")), 2)
            End If
            
            'consulto quanto é para ser o valor do desconto em real
            sSQL = "SELECT ValorDescReal FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ")"
            Set r = dbData.OpenRecordset(sSQL)

            If Not r.EOF Then
                vValorDescVenda = FormatNumber(ValidateNull(r("ValorDescReal")), 2)
            End If
            
            'se o valor total do desconto for maior que a soma dos desconto dos itens da venda
            If vValorDescVenda < vSomaDescItens Then
                vValorSobraDesc = CCur(vSomaDescItens - vValorDescVenda)
                sSQL = "UPDATE pedidos_itens SET Desconto = Desconto - " & Replace(CCur(vValorSobraDesc), ",", ".") & ", Total = Total + " & Replace(CCur(vValorSobraDesc), ",", ".") & " " & _
                        "WHERE (CODIGO = " & _
                "(SELECT MAX(CODIGO) FROM pedidos_itens WHERE (cod_pedido = " & txtCodPedido.Text & ")))"
                dbData.Execute sSQL
            ElseIf vValorDescVenda > vSomaDescItens Then
                vValorSobraDesc = CCur(vValorDescVenda - vSomaDescItens)
                sSQL = "UPDATE pedidos_itens SET Desconto = Desconto + " & Replace(CCur(vValorSobraDesc), ",", ".") & ", Total = Total - " & Replace(CCur(vValorSobraDesc), ",", ".") & " " & _
                        "WHERE (CODIGO = " & _
                "(SELECT MAX(CODIGO) FROM pedidos_itens WHERE (cod_pedido = " & txtCodPedido.Text & ")))"
                dbData.Execute sSQL
            End If
        Else
            If lblEstornar.Caption = "ESTORNO" Then
                sSQL = "UPDATE pedidos_itens SET Desconto = '0.00', Total = Subtotal " & _
                        "WHERE (cod_pedido = " & txtCodPedido.Text & ")"
                dbData.Execute sSQL
            End If
        End If
        
        'Retirar da tabela PRODUTOS as QUANTIDADES mencionadas no grid
        For i = 1 To Grid.Rows - 1
           dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque - " & Replace(CDbl(Grid.TextMatrix(i, 5)), ",", ".") & " WHERE (codigo = " & Grid.TextMatrix(i, 2) & ");"
        Next
        
        'CASHBACK
        If txtCodCliente.Text <> "1" Then
            If vCashbackAV = "SIM" Then
            
            'Cashback dar baixa
            If vUsandoCashBack = True Then
                dbData.Execute "UPDATE Pedidos_Cashback SET VALOR_ABATIDO = VALOR_CASHBACK, ABATIDO = 1, DATA_ABATIDO = '" & Format$(Date, "yyyy-dd-MM") & "', COD_PEDIDOABATIDO = " & txtCodPedido.Text & ", COD_FUNCIONARIO = " & txtCodFuncAP.Text & " WHERE (COD_CLIENTE = " & txtCodCliente.Text & ") and ABATIDO = 0 and INVALIDO = 0;"
            End If
            
            'valor do cashback
            vValorVenda = CCur(txtTotalDesc.Text)
            ValorCash = (vValorVenda * vCashbackValorAV) / 100
            
            'validade do cashback
            vCashbackValidade = Format(DateAdd("d", Val(vCashbackLimite), Date), "dd/mm/yy")
            
                lNovoCod = Autonumeracao_Cashback
                sSQL = "INSERT INTO Pedidos_Cashback (CODIGO, COD_PEDIDO, VALOR_VENDA, VALOR_CASHBACK, VALOR_ABATIDO, ABATIDO, VALIDADE, INVALIDO, COD_FUNCIONARIO, COD_CLIENTE) VALUES (" & _
                lNovoCod & ", " & txtCodPedido.Text & ", " & Replace(CCur(txtTotalDesc.Text), ",", ".") & ", " & Replace(CCur(ValorCash), ",", ".") & ", 0, 0, '" & Format$(vCashbackValidade, "yyyy-dd-MM") & "', 0, " & txtCodFuncAP.Text & ", " & txtCodCliente.Text & ");"
                dbData.Execute sSQL
            End If
        End If
        
        
        'Inicio do CUPOM FISCAL - NFCe
        If vConfImprimeNFCeLocal = "SIM" Then   'SE no arquivo ini tem SIM para imprimir NFCE
            If NFCe_OK = True Then              'SE dei SIM para imprimir NFCE
                
                    'verificar se o nfce foi gravado
                    sSQL = "SELECT Num_OS_VD_Origem FROM TbNFCe WHERE (Num_OS_VD_Origem = " & txtCodPedido & ");"
                    Set rNFCe = dbData.OpenRecordset(sSQL)
                    
                    If Not rNFCe.BOF Then
                        MsgBox "NFCe para esse pedido já foi criada"
                        GoTo SemGerarNFCeAV
                    Else    'se nao existir nfce
                       
                        'consultar o cpf/cnpj do cliente
                        sSQL2 = "SELECT CPF, CODIGO FROM cliente WHERE (codigo = " & txtCodCliente.Text & ");"
                        Set rCliente = dbData.OpenRecordset(sSQL2)
                        
                        If Not rCliente.EOF Then
                            If rCliente!Codigo = 1 Then
                            
                            
                                    If vNFCeConfCPF = "SIM" Then
                                            If ShowMsg("Deseja inserir o CPF no NFCe?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                                                vCPF = InputBox("Informe o CPF do cliente:", "EMISSÃO DE NFCe", "")
                                                    If Not Vazio(vCPF) Then
                                                            If Len(vCPF) = 11 Then
                                                                    If Validar_CPF(vCPF) = False Then
                                                                        MsgBox "CPF Informado não é valido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
                                                                        GoTo CPFDigitadoErrado2
                                                                    Else
                                                                        'vCPF = Format(vCPF, "000\.000\.000\-00")
                                                                        vCPF = RetirarMascaras(vCPF)
                                                                        GoTo PularInserirCPF
                                                                    End If
                                                            ElseIf Len(vCPF) = 14 Then
                                                                    If Validar_CNPJ(vCPF) = False Then
                                                                        MsgBox "CNPJ Informado não é valido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
                                                                        GoTo CNPJDigitadoErrado2
                                                                    Else
                                                                        'vCPF = Format(vCPF, "00\.000\.000\/0000\-00")
                                                                        vCPF = RetirarMascaras(vCPF)
                                                                        GoTo PularInserirCPF
                                                                    End If
                                                            End If
                                                    Else
                                                        vCPF = ""       'se o cpf for vazio
                                                    End If
                                            Else
                                                vCPF = ""           'se na msgbox colocar NÃO quer colocar cpf
                                            End If
                                    End If

                            Else
                                vCPF = RetirarMascaras(rCliente!CPF)
                            End If
                        Else
                            vCPF = ""
                        End If
                        
                        'validar CPF
                        Select Case Len(vCPF)
                            Case 0
                                If Len(vCPF) = 0 Then
                                    vCPF = Empty
                                Else
                                    vCPF = ""
                                End If
                                'KeyCode = 0
                            Case 14
CNPJDigitadoErrado2:
                                If Validar_CNPJ(vCPF) = False Then
                                        If vNFCeConfCPF = "SIM" Then
                                            If ShowMsg("Deseja inserir o CNPJ no NFCe?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                                                vCPF = InputBox("Informe o CNPJ do cliente:", "EMISSÃO DE NFCe", "")
                                                If Not Vazio(vCPF) Then
                                                    If Len(vCPF) = 11 Then
                                                        If Validar_CPF(vCPF) = False Then
                                                            MsgBox "CPF Informado não é valido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
                                                            GoTo CPFDigitadoErrado2
                                                        Else
                                                            vCPF = Format(vCPF, "000\.000\.000\-00")
                                                        End If
                                                    ElseIf Len(vCPF) = 14 Then
                                                        If Validar_CNPJ(vCPF) = False Then
                                                            MsgBox "CNPJ Informado não é valido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
                                                            GoTo CNPJDigitadoErrado2
                                                        Else
                                                            vCPF = Format(vCPF, "00\.000\.000\/0000\-00")
                                                        End If
                                                    End If
                                                Else
                                                    vCPF = ""
                                                End If
                                            Else
                                                vCPF = ""           'se na msgbox colocar NÃO quer colocar cpf
                                            End If
                                        End If
                                End If
                            Case 11
CPFDigitadoErrado2:
                                If Validar_CPF(vCPF) = False Then
                                        If vNFCeConfCPF = "SIM" Then
                                            If ShowMsg("Deseja inserir o CPF no NFCe?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                                                vCPF = InputBox("Informe o CPF do cliente:", "EMISSÃO DE NFCe", "")
                                                If Not Vazio(vCPF) Then
                                                    If Len(vCPF) = 11 Then
                                                        If Validar_CPF(vCPF) = False Then
                                                            MsgBox "CPF Informado não é valido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
                                                            GoTo CPFDigitadoErrado2
                                                        Else
                                                            vCPF = Format(vCPF, "000\.000\.000\-00")
                                                        End If
                                                    ElseIf Len(vCPF) = 14 Then
                                                        If Validar_CNPJ(vCPF) = False Then
                                                            MsgBox "CNPJ Informado não é valido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
                                                            GoTo CNPJDigitadoErrado2
                                                        Else
                                                            vCPF = Format(vCPF, "00\.000\.000\/0000\-00")
                                                        End If
                                                    End If
                                                Else
                                                    vCPF = ""       'se o cpf for vazio
                                                End If
                                            Else
                                                vCPF = ""           'se na msgbox colocar NÃO quer colocar cpf
                                            End If
                                        End If
                                End If
                            Case Is < 11
                                'MsgBox "CPF Informado não é valido", vbInformation, "ATENÇÃO! AVISO IMPORTANTE!"
                                'mskCNPJ.SetFocus
                        End Select
                        
PularInserirCPF:
                        If Len(vCPF) = 11 Then
                            vCPF = Format(vCPF, "000\.000\.000\-00")
                        ElseIf Len(vCPF) = 14 Then
                            vCPF = Format(vCPF, "00\.000\.000\/0000\-00")
                        Else
                            vCPF = ""
                        End If
                        
                        'coloca o cpf correto no cliente
                        dbData.Execute "UPDATE cliente SET cpf = '" & vCPF & "' WHERE (codigo = " & txtCodCliente.Text & ")"

                        sSQL = "EXEC NFCeIncluir " & txtCodPedido.Text
                        dbData.Execute sSQL
                    End If
                    
                    If rNFCe.State <> 0 Then rNFCe.Close
                    Set rNFCe = Nothing
                    
                    If rCliente!Codigo = 1 Then
                        dbData.Execute "UPDATE cliente SET cpf = '' WHERE (codigo = 1)"
                    End If
                    
                    If rCliente.State <> 0 Then rCliente.Close
                    Set rCliente = Nothing
                    
                    'If r2.State <> 0 Then r2.Close
                    'Set r2 = Nothing
                    
                    'transmitir nfce
                    sSQL = "SELECT IdNFProd FROM TbNFCe WHERE Num_OS_VD_Origem  = " & txtCodPedido.Text
                    Set rNFCe = dbData.OpenRecordset(sSQL)
                    
                    If rNFCe.RecordCount > 0 Then
                       sSQL = "INSERT INTO [TbNFCe_Faturas] " & _
                              "([IdNFProd] " & _
                              ",[IDParcela] " & _
                              ",[TipoPgto] " & _
                              ",[Vencimento] " & _
                              ",[Valor] " & _
                              ",[IdBandeira] " & _
                              ",[CartaoNumeroAutorizacao]) " & _
                              "SELECT " & rNFCe!IdNFProd & " " & _
                              "     ,NUMERO " & _
                              "     ,dbo.NFCeFormaPagto(FORMA_PGTO, TIPO_CARTAO) " & _
                              "     ,DATA " & _
                              "     ,VALOR " & _
                              "     ,'01' " & _
                              "     ,'' " & _
                              "FROM [parcelas] " & _
                              "WHERE COD_PEDIDO = " & txtCodPedido.Text
                       dbData.Execute sSQL
                    End If
                       DoEvents
                       
                    'verificando os itens do pedido
                       sSQL = "SELECT IdNFProd, IdNFProd_Item, IDProduto, CodBarras, DescricaoProduto, CodNcm, CFOP, Bc_Icms, ICMSCST, IPICST, COFINSCST, PISCST, UN " & _
                              "FROM TbNFCe_Itens " & _
                              "WHERE (IdNFProd = " & rNFCe!IdNFProd & ");"
                       Set rNFCeItens = dbData.OpenRecordset(sSQL)
                       'Debug.Print sSQL
                       
                       EncontroErroNFCe = False
                       
                        For i = 1 To rNFCeItens.RecordCount
                            
                            'NCM..........
                            If rNFCeItens!CodBarras <> "SEM GTIN" Then
                                If Len(rNFCeItens!CodBarras) > 13 Or Len(rNFCeItens!CodBarras) < 8 Then
                                    EncontroErroNFCe = True
                                Else
                                    EncontroErroNFCe = False
                                End If
                            Else
                                EncontroErroNFCe = False
                            End If
                            
                            If EncontroErroNFCe = True Then GoTo ContinuarNFCeAV
                                                        
                            'CFOP..........
                            If rNFCeItens!CFOP <> Empty Then
                                If Len(rNFCeItens!CFOP) > 4 Or Len(rNFCeItens!CFOP) < 4 Then
                                    EncontroErroNFCe = True
                                Else
                                    EncontroErroNFCe = False
                                End If
                            Else
                                EncontroErroNFCe = False
                            End If
                            
                            If EncontroErroNFCe = True Then GoTo ContinuarNFCeAV
                            
                            'ICMS CST..........
                            If rNFCeItens!icmsCST <> Empty Then
                                If Len(rNFCeItens!icmsCST) > 3 Or Len(rNFCeItens!icmsCST) < 3 Then
                                    EncontroErroNFCe = True
                                Else
                                    EncontroErroNFCe = False
                                End If
                            Else
                                EncontroErroNFCe = False
                            End If
                            
                            If EncontroErroNFCe = True Then GoTo ContinuarNFCeAV

                            'PIS CST..........
                            If rNFCeItens!pisCST <> Empty Then
                                If Len(rNFCeItens!pisCST) > 2 Or Len(rNFCeItens!pisCST) < 2 Then
                                    EncontroErroNFCe = True
                                Else
                                    EncontroErroNFCe = False
                                End If
                            Else
                                EncontroErroNFCe = False
                            End If
                            
                            If EncontroErroNFCe = True Then GoTo ContinuarNFCeAV

                            'COFINS CST..........
                            If rNFCeItens!cofinsCST <> Empty Then
                                If Len(rNFCeItens!cofinsCST) > 2 Or Len(rNFCeItens!cofinsCST) < 2 Then
                                    EncontroErroNFCe = True
                                Else
                                    EncontroErroNFCe = False
                                End If
                            Else
                                EncontroErroNFCe = False
                            End If
                            
                            If EncontroErroNFCe = True Then GoTo ContinuarNFCeAV
                            
                            'NCM..........
                            If rNFCeItens!CodNcm <> Empty Then
                                If Len(rNFCeItens!CodNcm) > 8 Or Len(rNFCeItens!CodNcm) < 8 Then
                                    EncontroErroNFCe = True
                                Else
                                    EncontroErroNFCe = False
                                End If
                            Else
                                EncontroErroNFCe = False
                            End If
                            
                            If EncontroErroNFCe = True Then GoTo ContinuarNFCeAV
                            
                            'UNIDADE DE MEDIDA..........
                            If rNFCeItens!UN <> Empty Then
                                If Len(rNFCeItens!UN) > 2 Or Len(rNFCeItens!UN) < 1 Then
                                    EncontroErroNFCe = True
                                Else
                                    EncontroErroNFCe = False
                                End If
                            Else
                                EncontroErroNFCe = False
                            End If
                            
                            If EncontroErroNFCe = True Then GoTo ContinuarNFCeAV
                        
                        rNFCeItens.MoveNext
                        Next
                       
ContinuarNFCeAV:
                       'transmitir o cupom fiscal - NFCe
                If EncontroErroNFCe = False Then
                       'transmitir o cupom NFCe
                       iRetorno = TransmitirNFCe(rNFCe!IdNFProd, "1", Not NFCeContingencia, "65")
                       
                       'impressão da DANFe
                       If iRetorno Then
                          Set sistNFe = New snfe.Util
                          
                          iRetorno = ConfiguraDLLNFeNFCe(65, "1", sistNFe)
                          If vNFCeImprimir = "SIM" Then
                                If Not NFCeContingencia Then
                                   Call sistNFe.DANFCeImprimir(xCaminhoXML, True, var_ImpNFCe, True, xCaminhoPDF, 0, False, False, "")
                                Else
                                   Call sistNFe.DANFCeOFFImprimir(xCaminhoXML, True, var_ImpNFCe, True, xCaminhoPDF, 0, False, False, "")
                                End If
                          End If
                       End If
                    'End If
                Else
                    If vNFCeCombinarImp = "NÃO" Then
                        If NFCe_OK = True Then
                            ImprimirVendaAVsemPergunta
                        Else
                            ImprimirVendaAV
                        End If
                    End If
                End If
                    
                    If vNFCeCombinarImp = "SIM" Then
                        ImprimirVendaAV
                    End If              'final de vNFCeCombinarImp = "SIM"
                Else    'meio do NFCe_OK = false
                    ImprimirVendaAV
                End If
        Else        'meio do vConfImprimeNFCeLocal = "SIM"

SemGerarNFCeAV:
            ImprimirVendaAV
        End If      'fim do vConfImprimeNFCeLocal = "SIM"
        
        dbData.Execute "UPDATE Pedidos_Reabertura SET STATUS_PEDIDO = 1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ") AND (DATA = (SELECT MAX(DATA) FROM Pedidos_Reabertura AS Pedidos_Reabertura_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & "))) AND (HORA = (SELECT MAX(HORA) FROM Pedidos_Reabertura AS Pedidos_Reabertura_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")));"
        LimparObjetos_Pedido
        LimparObjetos_Prazo
        LimparGrid_Pedido
        txtTotalGeral.Text = ""
        txtCodPedido.Text = ""
        lblEstornar.Caption = ""
        txtCodCliente.Text = ""
        cboCliente.Text = ""
        txtDataCompra.Text = Format(Date, "dd/mm/yyyy")
        frmVendaFechamento.Visible = False
        CriarNovoPedido
        If varTipoValorVenda = 1 Then txtCodBarra.SetFocus
        
ElseIf cboTipoPgto.Text = "ORÇAMENTO" Or cboTipoPgto.Text = "CONSIGNADO" Then
  'Dim varValidade As Date
  'varValidade = Format(DateAdd("d", 30, Date), "dd/mm/yy")
  
  Dim vTipoOrcamento As String
  If cboTipoPgto.Text = "CONSIGNADO" Then
    vTipoOrcamento = "CONSIGNADO"
  Else
    vTipoOrcamento = "ORÇAMENTO"
  End If
  
  
  'tabela configurações
  If bFechORC Then
    If cboTipoPgto.Text = "CONSIGNADO" Then
     If ShowMsg("Deseja finalizar esse consignado?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
    Else
     If ShowMsg("Deseja finalizar esse orçamento?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
    End If
  End If
  
    'pedido criando o pedido
              sSQL = "UPDATE pedidos SET " & _
                 "cod_pedido = " & txtCodPedido.Text & ", " & _
                 "cod_cliente = " & txtCodCliente.Text & ", " & _
                 "data_compra = '" & Format$(txtDataCompra, "yyyy-dd-MM") & "', " & _
                 "DATA_ENTREGA = '" & Format$(mskInicio, "yyyy-dd-MM") & "', " & _
                 "tipo_desc = '" & IIf(optDescRS.Value = True, "R", "P") & "', " & _
                 "valor_desc = " & Replace(CCur(txtDesc.Text), ",", ".") & ", " & _
                 "ValorDescReal = " & Replace(CCur(varValorRealDesc), ",", ".") & ", " & _
                 "ValorAcrescReal = " & Replace(CCur(varValorRealAcresc), ",", ".") & ", " & _
                 "TIPO_ACRESCIMO = '" & IIf(optAscrescRS.Value = True, "R", "P") & "', " & _
                 "VALOR_ACRESCIMO = " & Replace(CCur(txtAcresc.Text), ",", ".") & ", " & _
                 "TROCO = " & Replace(CCur(vValorTroco), ",", ".") & ", " & _
                 "RECEBIDO = " & Replace(CCur(vValorRececido), ",", ".") & ", " & _
                 "subtotal = " & Replace(CCur(txtSubtotal.Text), ",", ".") & ", " & _
                 "total = " & Replace(CCur(txtTotalDesc.Text), ",", ".") & ", " & _
                 "tipo_pagamento = 'À Prazo', pagamento = '" & var_PAGAMENTO & "', tipo_cartao = " & varTipoCartao & ", " & _
                 "cod_funcionario = " & txtCodFuncAP.Text & ", " & _
                 "tipo_pedido = '" & vTipoOrcamento & "', " & _
                 "caixa = '" & IIf(StatusBar1.Panels(2).Text = "", "CAIXA01", StatusBar1.Panels(2).Text) & "', " & _
                 "MAQUINA = '" & IIf(StatusBar1.Panels(4).Text = "", "PDV01", StatusBar1.Panels(4).Text) & "', " & _
                 "codcaixa = " & varCodCaixa & ", " & _
                 "status_pedido = 1 " & _
                 "WHERE (cod_pedido = " & txtCodPedido.Text & ");"
    dbData.Execute sSQL

        sSQL = "UPDATE pedidos_itens SET tipo_venda = '" & vTipoOrcamento & "' WHERE (cod_pedido = " & txtCodPedido.Text & ");"
        dbData.Execute sSQL
        
        If cboTipoPgto.Text = "CONSIGNADO" Then
            For i = 1 To Grid.Rows - 1
              dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque - " & Replace(CDbl(Grid.TextMatrix(i, 5)), ",", ".") & " WHERE (codigo = " & Grid.TextMatrix(i, 2) & ");"
            Next
        End If
        
        If vDescItensVenda <> "0,00" Then
            'adiciona em cada item do pedido o valor do desconto
            sSQL = "UPDATE pedidos_itens SET desconto = (subtotal * " & Replace(CDbl(vDescItensVenda), ",", ".") & " / 100), total = subtotal - (subtotal * " & Replace(CDbl(vDescItensVenda), ",", ".") & " / 100), data = '" & Format$(txtDataCompra, "yyyy-dd-MM") & "' where (cod_pedido = " & txtCodPedido.Text & ")"
            dbData.Execute sSQL
            
            'soma todos os descontos dos itens da venda em real
            sSQL = "SELECT SUM(Desconto) AS varSomaDescItens FROM pedidos_itens WHERE (cod_pedido = " & txtCodPedido.Text & ")"
            Set r = dbData.OpenRecordset(sSQL)
            
            'Dim vSomaDescItens As Currency
            If Not r.EOF Then
                vSomaDescItens = FormatNumber(ValidateNull(r("varSomaDescItens")), 2)
            End If
            
            'consulto quanto é para ser o valor do desconto em real
            sSQL = "SELECT ValorDescReal FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ")"
            Set r = dbData.OpenRecordset(sSQL)
            
            'Dim vValorDescVenda As Currency
            If Not r.EOF Then
                vValorDescVenda = FormatNumber(ValidateNull(r("ValorDescReal")), 2)
            End If
            
            'se o valor total do desconto for maior que a soma dos desconto dos itens da venda
            If vValorDescVenda < vSomaDescItens Then
                vValorSobraDesc = CCur(vSomaDescItens - vValorDescVenda)
                sSQL = "UPDATE pedidos_itens SET Desconto = Desconto - " & Replace(CCur(vValorSobraDesc), ",", ".") & ", Total = Total + " & Replace(CCur(vValorSobraDesc), ",", ".") & " " & _
                        "WHERE (CODIGO = " & _
                "(SELECT MAX(CODIGO) FROM pedidos_itens WHERE (cod_pedido = " & txtCodPedido.Text & ")))"
                dbData.Execute sSQL
            ElseIf vValorDescVenda > vSomaDescItens Then
                vValorSobraDesc = CCur(vValorDescVenda - vSomaDescItens)
                sSQL = "UPDATE pedidos_itens SET Desconto = Desconto + " & Replace(CCur(vValorSobraDesc), ",", ".") & ", Total = Total - " & Replace(CCur(vValorSobraDesc), ",", ".") & " " & _
                        "WHERE (CODIGO = " & _
                "(SELECT MAX(CODIGO) FROM pedidos_itens WHERE (cod_pedido = " & txtCodPedido.Text & ")))"
                dbData.Execute sSQL
            End If
        End If

    If iCopiasORC <> 0 Then  'saber a quantidade de copias
       If bEntregaORC Then
          If ShowMsg("Desesa Imprimir o pedido para ENTREGAR?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
             NumCopias = iCopiasORC + 1
          Else
             NumCopias = iCopiasORC
          End If
       Else
          NumCopias = iCopiasORC
       End If
    Else
       NumCopias = 1
    End If
          
        If vImprimirVendaAP Then       'Confirma se vai ter impressão
           If bConfImprORC Then
              If ShowMsg("Desesa Imprimir o pedido?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                 If iImprORC = 1 Then
                    For ii = 1 To NumCopias
                       Imprimir_Pedido
                    Next
                 ElseIf iImprORC = 2 Then
                    For ii = 1 To NumCopias
                       Imprimir_CupomSerrilha
                    Next
                 ElseIf iImprORC = 3 Then
                    For ii = 1 To NumCopias
                       Imprimir_CupomGuilhotina
                    Next
                 End If
              End If
           Else
              If iImprORC = 1 Then
                 For ii = 1 To NumCopias
                    Imprimir_Pedido
                 Next
              ElseIf iImprORC = 2 Then
                 For ii = 1 To NumCopias
                    Imprimir_CupomSerrilha
                 Next
              ElseIf iImprORC = 3 Then
                 For ii = 1 To NumCopias
                    Imprimir_CupomGuilhotina
                 Next
              End If
           End If
        End If
    End If
'COLOCAR O SUBTOTAL, DESCONTO E TOTAL DE CADA ITEM
  
LimparObjetos_Pedido
txtTotalGeral.Text = ""
LimparGrid_Pedido
LimparObjetos_Prazo
txtCodPedido.Text = ""
lblEstornar.Caption = ""
vDescItensVenda = "0,00"
frmVendaFechamento.Visible = False

If vBotaoOrcAtivo = True Then
    HabilitaObjetosVenda True
    vBotaoOrcamento = False
    vBotaoOrcAtivo = False
    Form_Load
Else
  CriarNovoPedido
  If varTipoValorVenda = 1 Then txtCodBarra.SetFocus
End If
'End If
cmdFinalizar.Enabled = True

If CAIXA_FECHADO = True Then
Else
    lblTipoPedido.Caption = ""
    cmdFinalizarAvista.Enabled = True
    cmdFinalizarPrazo.Enabled = True
    cmdOrçamento.Enabled = True
    cmdCancelarPedido.Enabled = True
    cmdRemover.Enabled = True
    cmdAvancado.Enabled = True
    cmdInfProduto.Enabled = True
    Grid.Enabled = True
    txtCodBarra.Enabled = True
    txtValor.Enabled = True
    txtQuant.Enabled = True
    txtTotal.Enabled = True
    vTipoEdicao = ""
End If

vUsandoCashBack = False
End Sub
Private Sub Verificar_Limite()
'Dim sSQL As String
'Dim r As ADODB.Recordset

Dim Limite As Currency
Dim vSomaAbertoNovaCompra As Currency
Dim Parcelas_Abertas As Currency
Dim vValorCompraAtual As Currency
Dim vLimiteAtual As Currency

If txtCodCliente.Text = "" Then Exit Sub
If txtTotalGeral.Text = "" Then Exit Sub

Passou_Limite = False
vValorCompraAtual = CCur(txtTotalGeral.Text)

'ver o limite do cliente
sSQL = "SELECT codigo, limite_credito FROM cliente WHERE (codigo = " & txtCodCliente.Text & ");"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then Limite = r("limite_credito")
If r.State <> 0 Then r.Close
Set r = Nothing

If Limite = 0 Then
   Passou_Limite = False
   Exit Sub
End If

'ver as parcelas em aberto do cliente
Parcelas_Abertas = 0

sSQL = "SELECT pedidos.cod_cliente, ISNULL(SUM(parcelas.valor), 0) AS somas_parcelas " & _
   "FROM parcelas INNER JOIN pedidos ON parcelas.cod_pedido = pedidos.cod_pedido " & _
   "WHERE (pedidos.cod_cliente = " & txtCodCliente.Text & ") AND (parcelas.status = 0) " & _
   "GROUP BY pedidos.cod_cliente;"

Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then Parcelas_Abertas = r("somas_parcelas")
If r.State <> 0 Then r.Close
Set r = Nothing

vLimiteAtual = Limite - Parcelas_Abertas
vSomaAbertoNovaCompra = Parcelas_Abertas + vValorCompraAtual

If Limite <= vSomaAbertoNovaCompra Then
   MsgBox "O limite de compra do cliente foi ultrapassado! " & Chr(13) & " O cliente possui atualmente R$ " & Format(vLimiteAtual, ocMONEY) & " de limite disponível!", vbInformation, "Aviso do Sistema"
   Passou_Limite = True
Else
   Passou_Limite = False
End If
End Sub

Private Sub Imprimir_Pedido()
'DEFINIR IMPRESSORA
'abrindo arquivo .ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"

'nome da maquina
If varImpPDF = True Then
    var_ImpNormal = "Impressora PDF"
Else
    var_ImpNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")
End If

Dim Prt As Printer
Dim oldPrinter As String

'Armazena o nome da impressora atual
oldPrinter = Printer.DeviceName

' Find and use the printer just selected in the ListBox
For Each Prt In Printers
   If Prt.DeviceName = var_ImpNormal Then
      Set Printer = Prt
      Exit For
   End If
Next

If cboTipoPgto.Text = "À PRAZO" Then
    If vQuantItensVenda < 18 Then
        If vTipoParcelaImpressao = 1 Then
            REL_Pedido_Mod05.loadPedidos txtCodPedido.Text
        Else
            REL_Pedido_APrazo.loadPedidos txtCodPedido.Text
        End If
    Else
        codPedido = txtCodPedido.Text

        sSQL = "SELECT produtos.descricao as var_desc, produtos.fabricante as vFab, quantidade, preco, pedidos_itens.subtotal, pedidos_itens.desconto, pedidos_itens.total, produtos.codigo as vCodProd " & _
                "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
                "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
                "WHERE (pedidos_itens.cod_pedido = " & codPedido & ") order by pedidos_itens.Codigo desc"
        Set r = dbData.OpenRecordset(sSQL)
        
        Me.Hide
        
        Set REL_Pedido_Completo.ReportMain1.Recordset = r
        
        REL_Pedido_Completo.txtDHead.Caption = "RELATÓRIO DO PEDIDO Nº " & txtCodPedido.Text
        REL_Pedido_Completo.Mostrar_Parcelas txtCodPedido.Text
        REL_Pedido_Completo.rfSubTotal.Caption = FormatNumber(txtSubtotal.Text, 2)
        REL_Pedido_Completo.txtDescontoRS.Caption = FormatNumber(varValorRealDesc, 2)
        REL_Pedido_Completo.rfTotal.Caption = FormatNumber(txtTotalDesc.Text, 2)
        REL_Pedido_Completo.rfDesc.Caption = FormatNumber(vDescItensVenda, 2)
        
        REL_Pedido_Completo.rfCliente.Caption = cboCliente.Text
        REL_Pedido_Completo.rfData.Caption = txtDataCompra.Text
        REL_Pedido_Completo.rfForma.Caption = cboTipoPgto.Text
        REL_Pedido_Completo.rfFunc.Caption = txtFuncAP.Text
        REL_Pedido_Completo.ReportMain1.NomeImpressora = var_ImpNormal
        REL_Pedido_Completo.ReportMain1.Ativar
        Unload REL_Pedido_Completo
    End If
   '
ElseIf cboTipoPgto.Text = "À VISTA" Then
   'If vTipoParcelaImpressao = 1 Then
    If vQuantItensVenda < 18 Then
        REL_Pedido_Mod06.loadPedidos txtCodPedido.Text
    Else
        codPedido = txtCodPedido.Text

        sSQL = "SELECT produtos.descricao as var_desc, produtos.fabricante as vFab, quantidade, preco, pedidos_itens.subtotal, pedidos_itens.desconto, pedidos_itens.total, produtos.codigo as vCodProd " & _
                "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
                "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
                "WHERE (pedidos_itens.cod_pedido = " & codPedido & ") order by pedidos_itens.Codigo desc"
        Set r = dbData.OpenRecordset(sSQL)
        
        Me.Hide
        
        Set REL_Pedido_Completo.ReportMain1.Recordset = r
        
        REL_Pedido_Completo.txtDHead.Caption = "RELATÓRIO DO PEDIDO Nº " & txtCodPedido.Text
        REL_Pedido_Completo.Mostrar_Parcelas txtCodPedido.Text
        REL_Pedido_Completo.rfSubTotal.Caption = FormatNumber(txtSubtotal.Text, 2)
        REL_Pedido_Completo.rfDesc.Caption = FormatNumber(vDescItensVenda, 2)
        REL_Pedido_Completo.txtDescontoRS.Caption = FormatNumber(varValorRealDesc, 2)
        REL_Pedido_Completo.rfTotal.Caption = FormatNumber(txtTotalDesc.Text, 2)
        
        
        REL_Pedido_Completo.rfCliente.Caption = cboCliente.Text
        REL_Pedido_Completo.rfData.Caption = txtDataCompra.Text
        REL_Pedido_Completo.rfForma.Caption = cboTipoPgto.Text
        REL_Pedido_Completo.rfFunc.Caption = txtFuncAP.Text
        REL_Pedido_Completo.ReportMain1.NomeImpressora = var_ImpNormal
        REL_Pedido_Completo.ReportMain1.Ativar
        Unload REL_Pedido_Completo
    End If
    'Else
    '    REL_Pedido_AVista.loadPedidos txtCodPedido.Text
   'End If
ElseIf cboTipoPgto.Text = "ORÇAMENTO" Or cboTipoPgto.Text = "CONSIGNADO" Then
    If vQuantItensVenda < 18 Then
        REL_Pedido_Orcamento.loadPedidos txtCodPedido.Text
    Else
        codPedido = txtCodPedido.Text

        sSQL = "SELECT produtos.descricao as var_desc, produtos.fabricante as vFab, quantidade, preco, pedidos_itens.subtotal, pedidos_itens.desconto, pedidos_itens.total, produtos.codigo as vCodProd " & _
                "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
                "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
                "WHERE (pedidos_itens.cod_pedido = " & codPedido & ") order by pedidos_itens.Codigo desc"
        Set r = dbData.OpenRecordset(sSQL)
        
        Me.Hide
        
        Set REL_Pedido_Completo.ReportMain1.Recordset = r
        
        REL_Pedido_Completo.txtDHead.Caption = "ORÇAMENTO Nº " & txtCodPedido.Text
        REL_Pedido_Completo.Mostrar_Parcelas txtCodPedido.Text
        REL_Pedido_Completo.rfSubTotal.Caption = FormatNumber(txtSubtotal.Text, 2)
        REL_Pedido_Completo.txtDescontoRS.Caption = FormatNumber(varValorRealDesc, 2)
        REL_Pedido_Completo.rfTotal.Caption = FormatNumber(txtTotalDesc.Text, 2)
        REL_Pedido_Completo.rfDesc.Caption = FormatNumber(vDescItensVenda, 2)
        
        REL_Pedido_Completo.rfCliente.Caption = cboCliente.Text
        REL_Pedido_Completo.rfData.Caption = txtDataCompra.Text
        REL_Pedido_Completo.rfForma.Caption = cboTipoPgto.Text
        REL_Pedido_Completo.rfFunc.Caption = txtFuncAP.Text
        REL_Pedido_Completo.ReportMain1.NomeImpressora = var_ImpNormal
        REL_Pedido_Completo.ReportMain1.Ativar
        Unload REL_Pedido_Completo
    End If
End If
Me.Show
End Sub

Private Sub cmdFinalizaravista_Click()
Dim sSQL As String
Dim r As ADODB.Recordset
Dim varTipoPgto As String
Dim varTipoCartao As String
vDataFlexivel = False
cmdCal1.Enabled = False

If CDate(lblDataAberturaCaixa.Caption) <> Date Then
    If MsgBox("A data do caixa aberto é diferente da data atual. Continuar mesmo assim?", vbQuestion + vbYesNo, "Alerta") = vbYes Then
        If txtTotalGeral.Text = "" Or txtTotalGeral.Text = "0,00" Then Exit Sub
        cboTipoPgto.Text = "À VISTA"
        frmVendaFechamento.Visible = True
        LimparObjetos_Prazo
        txtSubtotal.Text = txtTotalGeral.Text
        txtAcresc.Text = FormatNumber(0, 2)
        
        If lblEstornar.Caption = "ESTORNO" Then
            sSQL = "SELECT * FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ");"
            Set r = dbData.OpenRecordset(sSQL)
            
            'If varLoginFunc = "2" Then txtCodFuncAP.Text = ""
            'txtCodCliente.Text = ""
            
            If Not r.EOF Then
                 
                If r("TIPO_PEDIDO") = "ORÇAMENTO" Then
                    mskInicio.Text = Format(Date, "dd/mm/yy")
                    mskTermino.Text = Format(Date, "dd/mm/yy")
                    txtDataCompra.Text = Format(Date, "dd/mm/yyyy")
                Else
                    txtDataCompra.Text = Format(r("data_compra"), "dd/mm/yyyy")
                    'txtHoraCompra.Text = Format(r("data_compra"), "dd/mm/yyyy")
                    mskInicio.Text = Format(r("data_compra"), "dd/mm/yy")
                    mskTermino.Text = Format(r("data_compra"), "dd/mm/yy")
                End If
                
                txtCodFuncAP.Text = ValidateNull(r("cod_funcionario"))
                txtCodCliente.Text = ValidateNull(r("COD_CLIENTE"))
                'lblTipoPedido.Caption = ValidateNull(r("tipo_pedido"))
                varTipoPgto = ValidateNull(r("pagamento"))
                varTipoCartao = ValidateNull(r("TIPO_CARTAO"))
                
                If r("TIPO_DESC") = "P" Then
                    optDescPorc.Value = True
                Else
                    optDescRS.Value = True
                End If
                
                txtDesc.Text = FormatNumber(r("VALOR_DESC"), 2)
        
                 If varTipoPgto = "DINHEIRO" Then
                     cboFormaPgto.Text = "1 - DINHEIRO"
                 ElseIf varTipoPgto = "PROMISSORIA" Then
                     cboFormaPgto.Text = "2 - PROMISSÓRIA"
                 ElseIf varTipoPgto = "CARTAO" And varTipoCartao = "D" Then
                     cboFormaPgto.Text = "3 - CARTÃO - DÉBITO"
                 ElseIf varTipoPgto = "CARTAO" And varTipoCartao = "C" Then
                     cboFormaPgto.Text = "4 - CARTÃO - CRÉDITO"
                 ElseIf varTipoPgto = "CHEQUE" Then
                     cboFormaPgto.Text = "5 - CHEQUE"
                 ElseIf varTipoPgto = "BOLETO" Then
                     cboFormaPgto.Text = "6 - BOLETO"
                 ElseIf varTipoPgto = "FINANCEIRA" Then
                     cboFormaPgto.Text = "9 - FINANCEIRA"
                 End If
                 
                cboFormaPgto.Text = "1 - DINHEIRO"
                cboQuantForma.Text = "1 - FORMA"
                
                txtRecebido.SetFocus
            End If
                Calcular_Desconto
                'Calcular_Prazo
        Else
            
            cboFormaPgto.Text = "1 - DINHEIRO"
            cboQuantForma.Text = "1 - FORMA"
            optDescPorc.Value = False
            optDescPorc.Value = True
            cboQuantForma_LostFocus
            
            'limpar campo funcionario
            'If varLoginFunc <> "" Then
               If varLoginFunc = "2" Then
                  If lblEstornar.Caption <> "ESTORNO" Then
                     txtCodFuncAP.Text = ""
                     txtFuncAP.Text = ""
                     txtCodFuncAP.SetFocus
                  Else
                     'cboCliente.SetFocus
                  End If
               Else
                  'cboCliente.SetFocus
               End If
            'End If
            
            'If lblEstornar.Caption = "ESTORNO" Then
            
                mskInicio.Text = Format(Date, "dd/mm/yy")
                mskTermino.Text = Format(Date, "dd/mm/yy")
                'optDescPorc.Value = True
                'cboCliente.Text = ""
                BuscarClienteConsumidor    'desabilitei para testar no caskback
                Mostrar_Desconto
                Calcular_Desconto
                'Calcular_Prazo
                If varLoginFunc = "2" Then txtCodFuncAP.Text = ""
                If varLoginFunc = "2" Then txtFuncAP.Text = ""
                If varLoginFunc = "2" Then txtCodFuncAP.SetFocus Else txtRecebido.SetFocus
            
            'HabilitaObjetosVenda True
        End If
        
        cmdFinalizarAvista.Enabled = False
        cmdFinalizarPrazo.Enabled = False
        cmdOrçamento.Enabled = False
        cmdCancelarPedido.Enabled = False
        cmdRemover.Enabled = False
        cmdAvancado.Enabled = False
        cmdInfProduto.Enabled = False
        Grid.Enabled = False
        txtCodBarra.Enabled = False
        txtValor.Enabled = False
        txtQuant.Enabled = False
        txtTotal.Enabled = False
    End If
Else
        If txtTotalGeral.Text = "" Or txtTotalGeral.Text = "0,00" Then Exit Sub
        cboTipoPgto.Text = "À VISTA"
        frmVendaFechamento.Visible = True
        LimparObjetos_Prazo
        txtSubtotal.Text = txtTotalGeral.Text
        txtAcresc.Text = FormatNumber(0, 2)
        
        If lblEstornar.Caption = "ESTORNO" Then
            sSQL = "SELECT * FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ");"
            Set r = dbData.OpenRecordset(sSQL)
            
            'If varLoginFunc = "2" Then txtCodFuncAP.Text = ""
            'txtCodCliente.Text = ""
            
            If Not r.EOF Then

                
                If r("TIPO_PEDIDO") = "ORÇAMENTO" Then
                    mskInicio.Text = Format(Date, "dd/mm/yy")
                    mskTermino.Text = Format(Date, "dd/mm/yy")
                    txtDataCompra.Text = Format(Date, "dd/mm/yyyy")
                Else
                    txtDataCompra.Text = Format(r("data_compra"), "dd/mm/yyyy")
                    'txtHoraCompra.Text = Format(r("data_compra"), "dd/mm/yyyy")
                    mskInicio.Text = Format(r("data_compra"), "dd/mm/yy")
                    mskTermino.Text = Format(r("data_compra"), "dd/mm/yy")
                End If
                
                txtCodFuncAP.Text = ValidateNull(r("cod_funcionario"))
                txtCodCliente.Text = ValidateNull(r("COD_CLIENTE"))
                'lblTipoPedido.Caption = ValidateNull(r("tipo_pedido"))
                varTipoPgto = ValidateNull(r("pagamento"))
                varTipoCartao = ValidateNull(r("TIPO_CARTAO"))
                
                If r("TIPO_DESC") = "P" Then
                    optDescPorc.Value = True
                Else
                    optDescRS.Value = True
                End If
                
                txtDesc.Text = FormatNumber(r("VALOR_DESC"), 3)
        
                 If varTipoPgto = "DINHEIRO" Then
                     cboFormaPgto.Text = "1 - DINHEIRO"
                 ElseIf varTipoPgto = "PROMISSORIA" Then
                     cboFormaPgto.Text = "2 - PROMISSÓRIA"
                 ElseIf varTipoPgto = "CARTAO" And varTipoCartao = "D" Then
                     cboFormaPgto.Text = "3 - CARTÃO - DÉBITO"
                 ElseIf varTipoPgto = "CARTAO" And varTipoCartao = "C" Then
                     cboFormaPgto.Text = "4 - CARTÃO - CRÉDITO"
                 ElseIf varTipoPgto = "CHEQUE" Then
                     cboFormaPgto.Text = "5 - CHEQUE"
                 ElseIf varTipoPgto = "BOLETO" Then
                     cboFormaPgto.Text = "6 - BOLETO"
                 ElseIf varTipoPgto = "FINANCEIRA" Then
                     cboFormaPgto.Text = "9 - FINANCEIRA"
                 End If
                 
                cboFormaPgto.Text = "1 - DINHEIRO"
                cboQuantForma.Text = "1 - FORMA"
                
                txtRecebido.SetFocus
            End If
                Calcular_Desconto
                'Calcular_Prazo
        Else
            
            cboFormaPgto.Text = "1 - DINHEIRO"
            cboQuantForma.Text = "1 - FORMA"
            optDescPorc.Value = False
            optDescPorc.Value = True
            cboQuantForma_LostFocus
            
            'limpar campo funcionario
            'If varLoginFunc <> "" Then
               If varLoginFunc = "2" Then
                  If lblEstornar.Caption <> "ESTORNO" Then
                     txtCodFuncAP.Text = ""
                     txtFuncAP.Text = ""
                     txtCodFuncAP.SetFocus
                  Else
                     'cboCliente.SetFocus
                  End If
               Else
                  'cboCliente.SetFocus
               End If
            'End If
            
            'If lblEstornar.Caption = "ESTORNO" Then
            
                mskInicio.Text = Format(Date, "dd/mm/yy")
                mskTermino.Text = Format(Date, "dd/mm/yy")
                'optDescPorc.Value = True
                'cboCliente.Text = ""
                BuscarClienteConsumidor
                Mostrar_Desconto
                Calcular_Desconto
                'Calcular_Prazo
                If varLoginFunc = "2" Then txtCodFuncAP.Text = ""
                If varLoginFunc = "2" Then txtFuncAP.Text = ""
                If varLoginFunc = "2" Then txtCodFuncAP.SetFocus Else txtRecebido.SetFocus
            
            'HabilitaObjetosVenda True
        End If
        
        cmdFinalizarAvista.Enabled = False
        cmdFinalizarPrazo.Enabled = False
        cmdOrçamento.Enabled = False
        cmdCancelarPedido.Enabled = False
        cmdRemover.Enabled = False
        cmdAvancado.Enabled = False
        cmdInfProduto.Enabled = False
        Grid.Enabled = False
        txtCodBarra.Enabled = False
        txtValor.Enabled = False
        txtQuant.Enabled = False
        txtTotal.Enabled = False
End If
cmdFinalizar.Enabled = True
End Sub

Private Function AbreConexao()
On Error Resume Next
'Retorno = AbrePorta(Val(strPorta), Val(strVelocidade), Val(strDataBits), Val(strParidade))
Retorno = AbrePorta(1, 0, 0, 2)
If Retorno <> 1 Then
    MsgBox "erro ao abrir comunicação com a balança"
End If
End Function
Private Function PegarPesoToledo()
On Error Resume Next

Retorno = PegaPeso(0, Peso, "C:")
If Retorno = 1 Then
    If Peso = "SSSSS" Then
        MsgBox "SOBRE PESO NA BALANÇA", vbInformation
    ElseIf Peso = "IIIII" Or Peso = "00000" Then
        'PegarPesoToledo
        txtQuant.Text = Format(0, ocPESO)
    Else
        txtQuant.Text = Val(Mid(Peso, 1, 2)) & "," & Mid(Peso, 3)
    End If
End If

End Function
Private Function fechaConexao()
On Error Resume Next
Retorno = FechaPorta()

End Function




Private Sub cmdLicenca_Click()
'Dim varNomeProduto As String
'varNomeProduto = 54545454
'ShellExecute hwnd, "open", "https://cosmos.bluesoft.com.br/pesquisar?utf8=" + Chr(95) + "&q=" & varNomeProduto & "", vbNullString, vbNullString, conSwNo

Dim url As String
url = "https://pixgo.org/api/v1/checkout/checkout-public.php?id=ab6cced10452142110b4c8de47ad6755" ' Substitua pelo link desejado
    
' Abre a URL no navegador padrão
ShellExecute Me.hwnd, "open", url, vbNullString, vbNullString, SW_SHOWNORMAL

End Sub


Private Sub Form_Initialize()
'Dim Ver As String * 4       'balança toledo
'VersaoDLL (Ver)             'balança toledo
'Caption = "Toledo Easylink - P05_P05A: Exemplo em Visual Basic - V. " & Ver     'balança toledo
End Sub

Private Sub Form_Terminate()
If vFabBalanca = "TOLEDO" Then
    Call fechaConexao  'balança toledo
End If
End Sub



Private Sub lblMsg1_Click()
Estonar.Show
End Sub

Private Sub MSComm1_OnComm()
If MSComm1.CommEvent = comEvReceive Then
    Delay (100)
    strRecebe = MSComm1.Input
    If strRecebe <> "" Then
     returnOfECF (strRecebe)
    End If
End If
End Sub
Private Sub cmdFinalizarPrazo_Click()
vDataFlexivel = False
cmdCal1.Enabled = True

If CDate(lblDataAberturaCaixa.Caption) <> Date Then
    If MsgBox("A data do caixa aberto é diferente da data atual. Continuar mesmo assim?", vbQuestion + vbYesNo, "Alerta") = vbYes Then
        If txtTotalGeral.Text = "" Or txtTotalGeral.Text = "0,00" Then Exit Sub
        cboTipoPgto.Text = "À PRAZO"
        frmVendaFechamento.Visible = True
        LimparObjetos_Prazo
        txtSubtotal.Text = txtTotalGeral.Text
        txtAcresc.Text = FormatNumber(0, 2)
        HabilitaObjetosVenda True
        
        If lblEstornar.Caption = "ESTORNO" Then

            sSQL = "SELECT * FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ");"
            Set r = dbData.OpenRecordset(sSQL)
            
            'If varLoginFunc = "2" Then txtCodFuncAP.Text = ""
            'txtCodCliente.Text = ""
            
            If Not r.EOF Then
                If r("TIPO_PEDIDO") = "ORÇAMENTO" Then
                    mskInicio.Text = Format(Date, "dd/mm/yy")
                    mskTermino.Text = Format(Date, "dd/mm/yy")
                    txtDataCompra.Text = Format(Date, "dd/mm/yyyy")
                Else
                    txtDataCompra.Text = Format(r("data_compra"), "dd/mm/yyyy")
                    mskInicio.Text = Format(r("data_compra"), "dd/mm/yy")
                    mskTermino.Text = Format(r("data_compra"), "dd/mm/yy")
                End If

                varTipoPgto = ValidateNull(r("pagamento"))
                varTipoCartao = ValidateNull(r("TIPO_CARTAO"))
                
                If r("TIPO_DESC") = "P" Then
                    optDescPorc.Value = True
                Else
                    optDescRS.Value = True
                End If
                
                txtDesc.Text = FormatNumber(r("VALOR_DESC"), 2)
                
                 If varTipoPgto = "DINHEIRO" Then
                     cboFormaPgto.Text = "1 - DINHEIRO"
                 ElseIf varTipoPgto = "PROMISSORIA" Then
                     cboFormaPgto.Text = "2 - PROMISSÓRIA"
                 ElseIf varTipoPgto = "CARTAO" And varTipoCartao = "D" Then
                     cboFormaPgto.Text = "3 - CARTÃO - DÉBITO"
                 ElseIf varTipoPgto = "CARTAO" And varTipoCartao = "C" Then
                     cboFormaPgto.Text = "4 - CARTÃO - CRÉDITO"
                 ElseIf varTipoPgto = "CHEQUE" Then
                     cboFormaPgto.Text = "5 - CHEQUE"
                 ElseIf varTipoPgto = "BOLETO" Then
                     cboFormaPgto.Text = "6 - BOLETO"
                 ElseIf varTipoPgto = "FINANCEIRA" Then
                     cboFormaPgto.Text = "9 - FINANCEIRA"
                 End If
                 
                cboFormaPgto.Text = "2 - PROMISSÓRIA"
                cboQuantForma.Text = "1 - SEM ENTRADA"
                
                txtCodFuncAP.Text = ValidateNull(r("cod_funcionario"))
                txtCodCliente.Text = ValidateNull(r("COD_CLIENTE"))
                
                txtRecebido.SetFocus
                Calcular_Desconto
                Calcular_Parcelas
                Calcular_Prazo
            End If
        Else
            cboFormaPgto.Text = "2 - PROMISSÓRIA"
            cboQuantForma.Text = "1 - SEM ENTRADA"
            'mskInicio.Text = Format(txtDataCompra, "dd/mm/yy")
            Mostrar_ValorRestante
            Calcular_Parcelas
            Calcular_Prazo
            optDescPorc.Value = False
            optDescPorc.Value = True
            
            'mostrar o DESCONTO
            'If vValorDescFixoAP <> "" Then
            '   txtDesc.Text = Format(0, ocMONEY)
            '   optDescPorc.Value = True
            '   txtDesc.Text = Format(vValorDescFixoAP, ocMONEY)
            'Else
            '   txtDesc.Text = Format(0, ocMONEY)
            'End If
            
            'txtAcresc.Text = Format(0, ocMONEY)
            
            
            'limpar campo funcionario
        '    If varLoginFunc <> "" Then
               If varLoginFunc = "2" Then
                  If lblEstornar.Caption <> "ESTORNO" Then
                     txtCodFuncAP.Text = ""
                     txtFuncAP.Text = ""
                     txtCodFuncAP.SetFocus
                  Else
                     txtEntrada.SetFocus
                  End If
               Else
                  'cboCliente.SetFocus
               End If
         '   End If
            
            'Preencher_FormaPgto
            'cboformaPgto.ListIndex = 0
            Mostrar_Desconto
            Calcular_Desconto
            'Calcular_Prazo
            
            If varLoginFunc = "2" Then txtCodFuncAP.Text = ""
            If varLoginFunc = "2" Then txtFuncAP.Text = ""
            If varLoginFunc = "2" Then txtCodFuncAP.SetFocus Else cboCliente.SetFocus
            
            'HabilitaObjetosVenda True
        End If
        
        If lblEstornar.Caption <> "ESTORNO" And lblEstornar.Caption <> "REIMPRESSÃO" Then
           cboCliente.Clear
        End If
        
        HabilitaObjetosVenda False
        
        cmdFinalizarAvista.Enabled = False
        cmdFinalizarPrazo.Enabled = False
        cmdOrçamento.Enabled = False
        cmdCancelarPedido.Enabled = False
        cmdRemover.Enabled = False
        cmdAvancado.Enabled = False
        cmdInfProduto.Enabled = False
        Grid.Enabled = False
        txtCodBarra.Enabled = False
        txtValor.Enabled = False
        txtQuant.Enabled = False
        txtTotal.Enabled = False
    End If
Else
        If txtTotalGeral.Text = "" Or txtTotalGeral.Text = "0,00" Then Exit Sub
        cboTipoPgto.Text = "À PRAZO"
        frmVendaFechamento.Visible = True
        LimparObjetos_Prazo
        txtSubtotal.Text = txtTotalGeral.Text
        txtAcresc.Text = FormatNumber(0, 2)
        HabilitaObjetosVenda True
        
        If lblEstornar.Caption = "ESTORNO" Then
            'If r.State <> 0 Then r.Close
            'Set r = Nothing

            sSQL = "SELECT * FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ");"
            Set r = dbData.OpenRecordset(sSQL)
            
            'If varLoginFunc = "2" Then txtCodFuncAP.Text = ""
            'txtCodCliente.Text = ""
            
            'Debug.Print sSQL
            
            If Not r.EOF Then
                'MsgBox r("pagamento")
                varTipoPgto = ValidateNull(r("pagamento"))
                varTipoCartao = ValidateNull(r("TIPO_CARTAO"))
                If r("TIPO_DESC") = "P" Then
                    optDescPorc.Value = True
                Else
                    optDescRS.Value = True
                End If
                
                txtDesc.Text = FormatNumber(r("VALOR_DESC"), 2)
                
                 If varTipoPgto = "DINHEIRO" Then
                     cboFormaPgto.Text = "1 - DINHEIRO"
                 ElseIf varTipoPgto = "PROMISSORIA" Then
                     cboFormaPgto.Text = "2 - PROMISSÓRIA"
                 ElseIf varTipoPgto = "CARTAO" And varTipoCartao = "D" Then
                     cboFormaPgto.Text = "3 - CARTÃO - DÉBITO"
                 ElseIf varTipoPgto = "CARTAO" And varTipoCartao = "C" Then
                     cboFormaPgto.Text = "4 - CARTÃO - CRÉDITO"
                 ElseIf varTipoPgto = "CHEQUE" Then
                     cboFormaPgto.Text = "5 - CHEQUE"
                 ElseIf varTipoPgto = "BOLETO" Then
                     cboFormaPgto.Text = "6 - BOLETO"
                 ElseIf varTipoPgto = "FINANCEIRA" Then
                     cboFormaPgto.Text = "9 - FINANCEIRA"
                 End If
                 
                cboFormaPgto.Text = "2 - PROMISSÓRIA"
                cboQuantForma.Text = "1 - SEM ENTRADA"
                
                If r("TIPO_PEDIDO") = "ORÇAMENTO" Then
                    mskInicio.Text = Format(Date, "dd/mm/yy")
                    mskTermino.Text = Format(Date, "dd/mm/yy")
                    txtDataCompra.Text = Format(Date, "dd/mm/yyyy")
                Else
                    txtDataCompra.Text = Format(r("data_compra"), "dd/mm/yyyy")
                    mskInicio.Text = Format(r("data_compra"), "dd/mm/yy")
                    mskTermino.Text = Format(r("data_compra"), "dd/mm/yy")
                End If
                
                txtCodFuncAP.Text = ValidateNull(r("cod_funcionario"))
                txtCodCliente.Text = ValidateNull(r("COD_CLIENTE"))
                'lblTipoPedido.Caption = ValidateNull(r("tipo_pedido"))

                

                
                txtRecebido.SetFocus
                Calcular_Desconto
                Calcular_Parcelas
                Calcular_Prazo
            End If
        Else
            cboFormaPgto.Text = "2 - PROMISSÓRIA"
            cboQuantForma.Text = "1 - SEM ENTRADA"
            'mskInicio.Text = Format(txtDataCompra, "dd/mm/yy")
            Mostrar_ValorRestante
            Calcular_Parcelas
            Calcular_Prazo
            optDescPorc.Value = False
            optDescPorc.Value = True
            
            'mostrar o DESCONTO
            'If vValorDescFixoAP <> "" Then
            '   txtDesc.Text = Format(0, ocMONEY)
            '   optDescPorc.Value = True
            '   txtDesc.Text = Format(vValorDescFixoAP, ocMONEY)
            'Else
            '   txtDesc.Text = Format(0, ocMONEY)
            'End If
            
            'txtAcresc.Text = Format(0, ocMONEY)
            
            
            'limpar campo funcionario
        '    If varLoginFunc <> "" Then
               If varLoginFunc = "2" Then
                  If lblEstornar.Caption <> "ESTORNO" Then
                     txtCodFuncAP.Text = ""
                     txtFuncAP.Text = ""
                     txtCodFuncAP.SetFocus
                  Else
                     txtEntrada.SetFocus
                  End If
               Else
                  'cboCliente.SetFocus
               End If
         '   End If
            
            'Preencher_FormaPgto
            'cboformaPgto.ListIndex = 0
            Mostrar_Desconto
            Calcular_Desconto
            'Calcular_Prazo
            
            If varLoginFunc = "2" Then txtCodFuncAP.Text = ""
            If varLoginFunc = "2" Then txtFuncAP.Text = ""
            If varLoginFunc = "2" Then txtCodFuncAP.SetFocus Else cboCliente.SetFocus
            
            'HabilitaObjetosVenda True
        End If
        
        If lblEstornar.Caption <> "ESTORNO" And lblEstornar.Caption <> "REIMPRESSÃO" Then
           cboCliente.Clear
        End If
        
        HabilitaObjetosVenda False
        
        cmdFinalizarAvista.Enabled = False
        cmdFinalizarPrazo.Enabled = False
        cmdOrçamento.Enabled = False
        cmdCancelarPedido.Enabled = False
        cmdRemover.Enabled = False
        cmdAvancado.Enabled = False
        cmdInfProduto.Enabled = False
        Grid.Enabled = False
        txtCodBarra.Enabled = False
        txtValor.Enabled = False
        txtQuant.Enabled = False
        txtTotal.Enabled = False
End If
cmdFinalizar.Enabled = True
End Sub

Private Sub cmdImpOrcamentoCompleto_Click()
'Dim sSQL As String
'Dim r As ADODB.Recordset

sSQL = "SELECT produtos.codigo as var_cod, produtos.descricao as var_desc, quantidade, preco, ISNULL((quantidade * preco), 0) AS total, produtos.codigo " & _
         "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
         "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
         "WHERE (pedidos_itens.cod_pedido = " & txtCodPedido.Text & ") order by pedidos_itens.Codigo desc"
Set r = dbData.OpenRecordset(sSQL)
   
'colocar o nome da caixa na barra de status
Dim var_Impressora As String
'Dim oIni As Ini    'desativei aqui 09/11/22

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Impressora = oIni.LerTexto("DADOS_IMPRESSORA", "impressora")
Set oIni = Nothing

Me.Hide

Set REL_Orcamento_Completo.ReportMain1.Recordset = r

REL_Orcamento_Completo.txtDHead.Caption = "ORÇAMENTO COMPLETO - PEDIDO " & txtCodPedido.Text
REL_Orcamento_Completo.txtData.Caption = "DATA: " & StatusBar1.Panels(5).Text
'REL_Orcamento_Completo.rfDesc.Caption = Format(txtTotalCartao.Text, "#,##0.00")
REL_Orcamento_Completo.rfTotal.Caption = Format(txtTotalGeral.Text, "#,##0.00")

'REL_Orcamento_Completo.Relatorio.NomeImpressora = var_Impressora
REL_Orcamento_Completo.ReportMain1.Ativar
Unload REL_Orcamento_Completo

Me.Show 1
End Sub

Private Sub cmdInfProduto_Click()
'On Error GoTo erro
If frmProduto.Visible = True Then
    frmProduto.Visible = False
Else
    frmProduto.Visible = True
End If

Dim sSQL As String
Dim r As ADODB.Recordset

If Grid.Rows >= 2 Then
    If Grid.TextMatrix(Grid.Row, 2) = "" Then Exit Sub
    
    sSQL = "SELECT * FROM produtos WHERE (codigo = " & Grid.TextMatrix(Grid.Row, 2) & ");"
    Set r = dbData.OpenRecordset(sSQL)
    
    txtInfCusto.Text = ""
    txtInfVenda.Text = ""
    txtInfMargem.Text = ""
    txtInfDesc.Text = ""
    txtInfQuant.Text = ""
    
    If Not r.EOF Then
        txtInfDesc.Text = ValidateNull(r("descricao"))
        txtInfQuant.Text = ValidateNull(r("quant_estoque"))
    End If
    
    sSQL = "SELECT TOP 1 * FROM Produtos_Precos WHERE (COD_PRODUTO = " & Grid.TextMatrix(Grid.Row, 2) & ") order by CODIGO desc;"
    Set r = dbData.OpenRecordset(sSQL)
    
    If Not r.EOF Then
        txtInfCusto.Text = Format$(r("CUSTO"), ocMONEY)
        txtInfVenda.Text = Format$(r("VALOR_VV"), ocMONEY)
        txtInfMargem.Text = FormatNumber(r("MARGEM_VV"), 2) & "%"
    End If
    
    If r.State <> 0 Then r.Close
    Set r = Nothing
Else
    Exit Sub
End If
   
'erro:
'   ShowMsg "Não existe nenhum produto para ser excluido!", vbExclamation
'   Exit Sub
End Sub

Private Sub cmdMaqOK_Click()
StatusBar1.Panels(2).Text = cboMaquina.Text
HabilitaObjetosVenda False
frmMaquina.Visible = False
txtCodBarra.SetFocus
End Sub

Private Sub cmdMinimizar_Click()
'CloseWindow (hwnd)
Me.WindowState = 1
End Sub



Private Sub cmdNaoCadastrar_Click()
frmProdutoNaoCadastrado.Visible = False
txtCodBarra.Enabled = True
txtQuant.Enabled = True
txtQuant.Text = ""
txtCodBarra.Text = ""
txtCodBarra.SetFocus
End Sub

Private Sub cmdOKProdAvulso_Click()
frmProdutoAvulso.Visible = False
txtCodBarra.Enabled = True
txtQuant.Enabled = True
txtCodBarra.Text = "00001"
txtCodBarra.SetFocus
SendKey ocKEYENTER
End Sub

Private Sub cmdOrcamento_Click()
HabilitaObjetosVenda True
vBotaoOrcAtivo = True
Form_Load
End Sub

Private Sub cmdOrçamento_Click()
If lblEstornar.Caption = "REIMPRESSÃO" Then Exit Sub
If txtTotalGeral.Text = "" Or txtTotalGeral.Text = "0,00" Then Exit Sub

If tipoEmpresa = 4 Then
    cboTipoPgto.Text = "CONSIGNADO"
Else
    cboTipoPgto.Text = "ORÇAMENTO"
End If

frmVendaFechamento.Visible = True
LimparObjetos_Prazo
txtSubtotal.Text = txtTotalGeral.Text
txtAcresc.Text = FormatNumber(0, 2)
HabilitaObjetosVenda True

If vTipoEdicao <> "EDITAR" Then
    'frmVendaFechamento.Visible = True
    'LimparObjetos_Prazo
    'txtSubTotal.Text = txtTotalGeral.Text
    optDescPorc.Value = True
    txtAcresc.Text = FormatNumber(0, 2)
    txtDesc.Text = FormatNumber(0, 2)
    cboFormaPgto.Text = "1 - DINHEIRO"
    cboQuantForma.Text = "1 - FORMA"
    cboFormaPgtoEntrada.Enabled = False

    mskInicio.Text = Format(Date, "dd/mm/yy")
    mskTermino.Text = Format(Date, "dd/mm/yy")
    'optDescPorc.Value = True
    'cboCliente.Text = ""
    If tipoEmpresa = 4 Then
        txtCodCliente.Text = ""
    Else
        BuscarClienteConsumidor
    End If
    If varLoginFunc = "2" Then txtCodFuncAP.Text = ""
    If varLoginFunc = "2" Then txtFuncAP.Text = ""
    If varLoginFunc = "2" Then txtCodFuncAP.SetFocus Else txtRecebido.SetFocus
    Calcular_Desconto
Else
    sSQL = "SELECT * FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ");"
    Set r = dbData.OpenRecordset(sSQL)
            
    If Not r.EOF Then
        If r("TIPO_PEDIDO") = "ORÇAMENTO" Then
            mskInicio.Text = Format(Date, "dd/mm/yy")
            mskTermino.Text = Format(Date, "dd/mm/yy")
            txtDataCompra.Text = Format(Date, "dd/mm/yyyy")
        Else
            txtDataCompra.Text = Format(r("data_compra"), "dd/mm/yyyy")
            mskInicio.Text = Format(r("data_compra"), "dd/mm/yy")
            mskTermino.Text = Format(r("data_compra"), "dd/mm/yy")
        End If
        
        varTipoPgto = ValidateNull(r("pagamento"))
        varTipoCartao = ValidateNull(r("TIPO_CARTAO"))
        
        If r("TIPO_DESC") = "P" Then
            optDescPorc.Value = True
        Else
            optDescRS.Value = True
        End If
        
        txtDesc.Text = FormatNumber(r("VALOR_DESC"), 2)
        
        txtCodFuncAP.Text = ValidateNull(r("cod_funcionario"))
        txtCodCliente.Text = ValidateNull(r("COD_CLIENTE"))
         
        cboFormaPgto.Text = "1 - DINHEIRO"
        cboQuantForma.Text = "1 - FORMA"
        
        'txtRecebido.SetFocus
        Calcular_Desconto
        'Calcular_Parcelas
        'Calcular_Prazo
    End If
End If

cmdFinalizarAvista.Enabled = False
cmdFinalizarPrazo.Enabled = False
cmdOrçamento.Enabled = False
cmdCancelarPedido.Enabled = False
cmdRemover.Enabled = False
cmdAvancado.Enabled = False
cmdInfProduto.Enabled = False
Grid.Enabled = False
txtCodBarra.Enabled = False
txtValor.Enabled = False
txtQuant.Enabled = False
txtTotal.Enabled = False
End Sub

Private Sub cmdRemover_Click()
On Error GoTo erro

If Grid.Rows > 1 Then
   If Grid.TextMatrix(Grid.Row, 1) = "" Then GoSub erro
   If ShowMsg("Deseja remover o produto: " & Grid.TextMatrix(Grid.Row, 3) & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
   
   dbData.Execute "DELETE FROM pedidos_itens WHERE (codigo = " & Grid.TextMatrix(Grid.Row, 1) & ") AND (cod_produto = " & Grid.TextMatrix(Grid.Row, 2) & ");"
   
   MostrarGrid_Produtos
   'Calcular_Valor_Geral
   Calcular_Troco
   If txtCodBarra.Enabled = True Then txtCodBarra.SetFocus
   Exit Sub
Else
    Exit Sub
End If
   
erro:
   ShowMsg "Não existe nenhum produto para ser excluido!", vbExclamation
   Exit Sub
End Sub

Private Sub cmdSenha_Click()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtSenha.Text = "" Then ShowMsg "ACESSO NEGADO!" & vbCrLf & "Senha obrigatória", vbInformation: Exit Sub

'If Liberar = True Then
   'chkLiberarVenda.Enabled = True
   'chkLiberarVenda.Value = Checked
'   txtSenha.Text = ""
'Else

    Set oCfg = sysConfig("TIPOLOGIN")
    
    If oCfg.Value = "NOME" Then
        If txtCodUsuario.Text = "" Then ShowMsg "ACESSO NEGADO!" & vbCrLf & "Usuário obrigatório", vbInformation: Exit Sub
        sSQL = "SELECT codigo, password, nivel, login FROM Usuario WHERE (password = '" & txtSenha.Text & "') AND (codigo = " & txtCodUsuario.Text & ");"
    Else
        sSQL = "SELECT codigo, password, nivel, cpf FROM Usuario WHERE (password = '" & txtSenha.Text & "') AND (cpf = '" & mskCPF.Text & "');"
    End If
    
   Set r = dbData.OpenRecordset(sSQL)
   
    If Not r.EOF Then
        LimparGrid_Pedido
        If varNomeBotao = "Pedidos" Then
            'Load Estonar
            Estonar.PegarCodUsuario (r("codigo"))
            cboUsuario.Text = ""
            txtCodUsuario.Text = ""
            txtSenha.Text = ""
            Estonar.Hide
                Estonar.lblCodUser1.Visible = True
                Estonar.lblCodUser2.Visible = True
                Estonar.lblUser1.Visible = True
                Estonar.lblUser2.Visible = True
                Estonar.lblCodUser2.Caption = Format(r("codigo"))
                Estonar.lblUser2.Caption = r("login")
            vCodFunc = Format(r("codigo"))
            'vCodUsuario = vCodFunc
            Estonar.Show 1
        ElseIf varNomeBotao = "Financeiro" Then
            'vChamouCaixa = "PDV"
            'Me.Hide
            vCodFunc = Format(r("codigo"))
            'Principal_Caixa.Show 1
            vChamouCaixa = "PDV"
            Me.Hide
            Principal_Caixa.Show 1
        ElseIf varNomeBotao = "NFCe" Then
            NFCe_Consultar.Show 1
        ElseIf varNomeBotao = "Cliente" Then
            Clientes_Cadastro.Show 1
        ElseIf varNomeBotao = "Produtos" Then
            Produtos_Cadastro.Show 1
        ''ElseIf varNomeBotao = "ClienteDebito" Then
        ''    varLiberarVendaDevedor = True
        End If
      
      txtSenha.Text = ""
      frmSenha.Visible = False
      frmAvancado.Visible = False
      HabilitaObjetosVenda False
   Else
      ShowMsg "ACESSO NEGADO!" & vbCrLf & "Você não tem nivel de acesso a esse recurso", vbInformation
      txtSenha.Text = ""
      frmSenha.Visible = False
      frmAvancado.Visible = False
      HabilitaObjetosVenda False
      If txtCodBarra.Enabled = True Then txtCodBarra.SetFocus
      ''varLiberarVendaDevedor = False
  End If
'End If
End Sub

Private Sub cmdUsarCadastrado_Click()
frmProdutoNaoCadastrado.Visible = False
frmProdutoAvulso.Visible = True
txtCodBarra.Enabled = True
txtQuant.Enabled = True
txtValorProdAvulso.SetFocus
End Sub

Private Sub cmdVP_Click()
TipoValorVenda = "VP"
frmTipoVenda.Visible = False
'verificar se o pedido está livre
Dim var_NroPedido As Long
var_NroPedido = ExistePedidoLivre

'Nenhum pedido livre
If var_NroPedido = -1 Then
   txtCodPedido = AutoNumeracao_Pedido
   dbData.Execute "INSERT INTO pedidos (cod_pedido, data_compra, status_pedido, caixa, maquina, reaberto, cancelado, orcamento) VALUES (" & txtCodPedido.Text & ", '" & Format$(Now, "yyyy-dd-MM") & "', 0, '" & var_Caixa & "', '" & var_Maquina & "', 0, 0, 0);"
Else
   txtCodPedido = var_NroPedido
End If

HabilitaObjetosVenda False
txtCodBarra.SetFocus
End Sub

Private Sub cmdVV_Click()
TipoValorVenda = "VV"
frmTipoVenda.Visible = False
    'verificar se o pedido está livre
    Dim var_NroPedido As Long
    var_NroPedido = ExistePedidoLivre
    
    'Nenhum pedido livre
    If var_NroPedido = -1 Then
       txtCodPedido = AutoNumeracao_Pedido
       dbData.Execute "INSERT INTO pedidos (cod_pedido, data_compra, status_pedido, caixa, maquina, reaberto, cancelado, orcamento) VALUES (" & txtCodPedido.Text & ", '" & Format$(Now, "yyyy-dd-MM") & "', 0, '" & var_Caixa & "', '" & var_Maquina & "', 0, 0, 0);"
    Else
       txtCodPedido = var_NroPedido
    End If

HabilitaObjetosVenda False
txtCodBarra.SetFocus
End Sub

Private Sub Command1_Click()
'HabilitaObjetosVenda False
'LimparObjetos_Prazo
'frmVendaFechamento.Visible = False
'txtTotalGeral.Text = Format(txtSubTotal.Text, ocMONEY)
'Calcular_Valor_Geral
cmdCancelar_Click
End Sub


Private Sub Form_Activate()
If frmProdutoNaoCadastrado.Visible = False And frmVendaFechamento.Visible = False And frmTipoVenda.Visible = False And frmAvancado.Visible = False Then
    If txtCodBarra.Enabled = False Then
        HabilitaObjetosVenda False
    End If
End If

If vPedirPeso = False Then
    If txtCodBarra.Enabled = True Then txtCodBarra.SetFocus
End If

If txtCodPedido.Text = "" Then
    If varTipoValorVenda = 2 Then
      If CAIXA_FECHADO = False Then frmTipoVenda.Visible = True
      HabilitaObjetosVenda True
    End If
End If

Verificar_Caixa
Verificar_NFCe
'Verificar_Backup 'desativei para colocar no online commerce
End Sub
Private Sub Calcular_Prazo()
If cboPrazo.Text = "" Then Exit Sub
Dim vDataInicialCerta As Date

If txtDataCompra.Text = "" Then txtDataCompra.Text = Format(Date, "dd/mm/yyyy")
If mskInicio.Text = "" Then mskInicio.Text = Format(txtDataCompra, "dd/mm/yy")

If vDataFlexivel = True Then
    vDataInicialCerta = Format(mskInicio.Text, "dd/mm/yy")
Else
    vDataInicialCerta = Format(txtDataCompra.Text, "dd/mm/yy")
End If

If cboPrazo.Text = "30" Then
    If txtEntrada.Text = "0,00" Or txtEntrada.Text = "" Then
        If vDataFlexivel = False Then
            mskInicio.Text = Format(DateAdd("m", Val(1), vDataInicialCerta), "dd/mm/yy")
        Else
            mskInicio.Text = Format(vDataInicialCerta, "dd/mm/yy")
        End If
        
        mskTermino.Text = Format(DateAdd("m", Val(cboQuantParc.Text) - 1, mskInicio.Text), "dd/mm/yy")
    Else
            mskInicio.Text = Format(vDataInicialCerta, "dd/mm/yy")
            If vDataFlexivel = False Then
                mskTermino.Text = Format(DateAdd("m", Val(cboQuantParc.Text), mskInicio.Text), "dd/mm/yy")
            Else
                mskTermino.Text = Format(DateAdd("m", Val(cboQuantParc.Text) - 1, mskInicio.Text), "dd/mm/yy")
            End If
    End If
Else
    If txtEntrada.Text = "0,00" Or txtEntrada.Text = "" Then
       mskInicio.Text = Format(DateAdd("d", Val(cboPrazo.Text), txtDataCompra), "dd/mm/yy")
       If cboQuantParc.Text = "1" Then
            mskTermino.Mask = ""
            mskTermino.Text = ""
        Else
            'mskTermino.Text = Format(DateAdd("d", Val(cboPrazo.Text) - 1, mskInicio.Text), "dd/mm/yy")
            mskTermino.Text = Format(DateAdd("d", Val(cboPrazo.Text), mskInicio.Text), "dd/mm/yy")
        End If
    Else
       mskInicio.Text = Format(txtDataCompra, "dd/mm/yy")
       Dim vDataFim As Date
       'vDataFim = Format(DateAdd("d", Val(cboPrazo.Text) - 1, mskInicio.Text), "dd/mm/yy")
       vDataFim = Format(DateAdd("d", Val(cboPrazo.Text), mskInicio.Text), "dd/mm/yy")
       mskTermino.Text = Format(DateAdd("m", Val(cboQuantParc.Text), vDataFim), "dd/mm/yy")
    End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF10 Then
   cmdFinalizaravista_Click
ElseIf KeyCode = vbKeyF12 Then
   cmdFinalizarPrazo_Click
ElseIf KeyCode = vbKeyF5 Then
    Parcelas_Consulta.Show
ElseIf KeyCode = vbKeyF6 Then
    If txtCodCliente <> "1" And txtCodCliente.Text <> "" Then
    If vCashbackAV = "SIM" Or vCashbackAP = "SIM" Then
        If lstCashBack.Visible = False Then
            dbData.Execute "UPDATE Pedidos_Cashback SET INVALIDO = 1 WHERE (COD_CLIENTE = " & txtCodCliente.Text & ") and ABATIDO = 0 and INVALIDO = 0 and VALIDADE < CONVERT(DATETIME, CONVERT(date, GETDATE()));"

            lstCashBack.Visible = True
            Dim ListaCash As ListItem
            lstCashBack.FullRowSelect = True
            lstCashBack.LabelEdit = lvwManual
            lstCashBack.Visible = True
            lstCashBack.View = lvwReport
            lstCashBack.HideSelection = False
            lstCashBack.ListItems.Clear
            
            lstCashBack.ColumnHeaders.Clear
            lstCashBack.ColumnHeaders.Add , , "CÓDIGO", 1200
            lstCashBack.ColumnHeaders.Add , , "CÓD.VENDA", 1200
            lstCashBack.ColumnHeaders.Add , , "VLR.VENDA", 1200
            lstCashBack.ColumnHeaders.Add , , "CASHBACK", 1200
            lstCashBack.ColumnHeaders.Add , , "VALIDADE", 1200
            
            sSQL = "SELECT CODIGO, COD_PEDIDO, VALOR_VENDA, VALOR_CASHBACK, VALIDADE " & _
                    "From Pedidos_Cashback " & _
                    "Where (COD_CLIENTE = " & txtCodCliente.Text & " ) and ABATIDO = 0 and INVALIDO = 0 " & _
                    "ORDER BY VALIDADE;"
            
            Set r = dbData.OpenRecordset(sSQL)
            
            If Not r Is Nothing Then
               Do While Not r.EOF
                  'primeira coluna
                  Set ListaCash = lstCashBack.ListItems.Add(, , r("CODIGO"))
                  'segunda e terceira coluna, que são sub itens da coluna 1
                    ListaCash.SubItems(1) = ValidateNull(r("COD_PEDIDO"))
                    ListaCash.SubItems(2) = FormatNumber(ValidateNull(r("VALOR_VENDA")), 2)
                    ListaCash.SubItems(3) = FormatNumber(ValidateNull(r("VALOR_CASHBACK")), 2)
                    ListaCash.SubItems(4) = ValidateNull(r("VALIDADE"))
                  r.MoveNext
               Loop
               
               If r.State <> 0 Then r.Close
               Set r = Nothing
            End If
            
            With lstCashBack
                For i = 1 To .ListItems.Count
                    ListaCash.ListSubItems(3).Bold = True
                    ListaCash.ListSubItems(3).ForeColor = vbRed
                Next i
            End With
        Else
            lstCashBack.Visible = False
        End If
    End If
    End If
ElseIf KeyCode = vbKeyF7 Then
    If txtCodCliente <> "1" And txtCodCliente.Text <> "" Then
    If vCashbackAV = "SIM" Or vCashbackAP = "SIM" Then
        dbData.Execute "UPDATE Pedidos_Cashback SET INVALIDO = 1 WHERE (COD_CLIENTE = " & txtCodCliente.Text & ") and ABATIDO = 0 and INVALIDO = 0 and VALIDADE < CONVERT(DATETIME, CONVERT(date, GETDATE()));"

        sSQL = "SELECT Sum(VALOR_CASHBACK) as vValorSomaCash " & _
                "From Pedidos_Cashback " & _
                "Where (COD_CLIENTE = " & txtCodCliente.Text & " ) and ABATIDO = 0 and INVALIDO = 0 "
        
        Set r = dbData.OpenRecordset(sSQL)
        
        If Not r Is Nothing Then
            optDescRS.Value = True
            txtDesc.Text = FormatNumber(ValidateNull(r("vValorSomaCash")), 2)
            If txtDesc.Text > 0 Then vUsandoCashBack = True Else vUsandoCashBack = False
           If r.State <> 0 Then r.Close
           Set r = Nothing
        Else
            vUsandoCashBack = False
        End If
    End If
    End If
ElseIf KeyCode = vbKeyF1 Then
    If frmProdutoNaoCadastrado.Visible = True Then
        cmdUsarCadastrado_Click
    Else
        Parcelas.Show 1
    End If
ElseIf KeyCode = vbKeyF2 Then
    If frmProdutoNaoCadastrado.Visible = True Then
        cmdCadastarProduto_Click
    Else
        cmdInfProduto_Click
    End If
ElseIf KeyCode = vbKeyF3 Then
    If frmProdutoNaoCadastrado.Visible = True Then
        cmdNaoCadastrar_Click
        txtQuant.Text = ""
    Else
        Grid_DblClick
    End If
ElseIf KeyCode = vbKeyF4 Then
    If vUsarBalanca = "SIM" Then
        If txtCodProduto.Text = "" Then Exit Sub
        If vFabBalanca = "URANO" Then
            Call PegarPesoUrano
        ElseIf vFabBalanca = "TOLEDO" Then
            Call PegarPesoToledo
        ElseIf vFabBalanca = "BALMAK" Then
            Call PegarPesoToledo
        End If
        'txtQuant.Text = " 0,683"
        PesoF4 = True
    End If
ElseIf KeyCode = vbKeyDelete Then
   If frmVendaFechamento.Visible = False Then
      cmdRemover_Click
   End If
ElseIf KeyCode = vbKey1 Then
    If frmTipoVenda.Visible = True Then
        cmdVV_Click
   End If
ElseIf KeyCode = vbKey2 Then
    If frmTipoVenda.Visible = True Then
        cmdVP_Click
    End If
ElseIf KeyCode = vbKey3 Then
    If frmTipoVenda.Visible = True Then
        cmdAV_Click
    End If
ElseIf KeyCode = vbKey4 Then
    If frmTipoVenda.Visible = True Then
        cmdAP_Click
    End If
End If
End Sub

Private Sub Form_Load()
'On Error GoTo Tratar_Erro

If vBotaoOrcAtivo = True Then
    vBotaoOrcamento = True
Else
    vBotaoOrcamento = False
End If

LimparGrid_Pedido

'varTipoValorVenda = 0 'coloquei no dia que criar os caixas multiplos por causa dos erros apos abrir avançado

CAIXA_FECHADO = False

frmTipoVenda.Visible = False
vDataFlexivel = False

varNomeBotao = "" 'variavel usando para saber qual botao acionou a senha

'===========================SETANDO ARQUIVO .INI
'abrindo arquivo .ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"

'Balança
vUsarBalanca = oIni.LerTexto("USAR_BALANCA", "resposta")

'nome da caixa
var_Caixa = oIni.LerTexto("DADOS_CAIXA", "caixa")
StatusBar1.Panels(2).Text = var_Caixa

'nome da Maquina
var_Maquina = oIni.LerTexto("DADOS_MAQUINA", "maquina")
StatusBar1.Panels(4).Text = var_Maquina

'var_ImpTermica = oIni.LerTexto("IMPRESSORA_TERMICA", "impressora")
'var_ImpNormal = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")
'var_ImpNFCe = oIni.LerTexto("IMPRESSORA_NFCE", "impressora")    'desativei aqui 09/11/22
vConfImprimeNFCeLocal = oIni.LerTexto("IMPRIMIR_NFCE", "resposta")
Set oIni = Nothing

'=========================== TABELA CONFIGURAÇÕES
'mostrar fundo do pdv
Dim sFundo As String
Set oCfg = sysConfig("FUNDO_PDV")
sFundo = oCfg.Value
Set oCfg = Nothing

If Existe(sFundo) Then
    If Dir$(sFundo) <> "" Then Set Picture = LoadPicture(sFundo)
    'logomarca impressa do cupom
    Dim sLogo As String
    Set oCfg = sysConfig("LOGO_CUPOM")
    sLogo = oCfg.Value
    Set oCfg = Nothing
    If Dir$(sLogo) <> "" Then Set imLogoCupom.Picture = LoadPicture(sLogo)
End If

'tipo de venda = 1 simples e 2 multiplus preços
Set cCfg = sysConfig("TIPOVALORVENDA")
varTipoValorVenda = cCfg.Value
Set cCfg = Nothing

'============================= VERIFICAR CAIXA
Verificar_Caixa

'If CAIXA_FECHADO = True Then
'    If vBotaoOrcamento = False Then
'        txtCodBarra.Enabled = False
'        txtValor.Enabled = False
'        txtQuant.Enabled = False
'        txtTotal.Enabled = False
'        txtTotalGeral.Enabled = False
'        cmdAlterar.Enabled = False
'        cmdFinalizarAvista.Enabled = False
'        cmdFinalizarPrazo.Enabled = False
'        cmdOrçamento.Enabled = False
'        cmdCancelarPedido.Enabled = False
'        cmdRemover.Enabled = False
'        cmdAvancado.Enabled = False
'        cmdInfProduto.Enabled = False
'        Grid.Enabled = False
'        frmCaixaFechado.Visible = True
'        Exit Sub
'    Else
'        txtCodBarra.Enabled = True
'        txtValor.Enabled = True
'        txtQuant.Enabled = True
'        txtTotal.Enabled = True
'        txtTotalGeral.Enabled = True
'        cmdAlterar.Enabled = True
'        cmdFinalizarAvista.Enabled = False
'        cmdFinalizarPrazo.Enabled = False
'        cmdOrçamento.Enabled = True
'        cmdCancelarPedido.Enabled = True
'        cmdRemover.Enabled = True
'        cmdAvancado.Enabled = False
'        cmdInfProduto.Enabled = False
'        frmCaixaFechado.Visible = False
'        Grid.Enabled = True
'        StatusBar1.Panels(7).Text = Format(0, "0000")
'    End If
'Else
'    If varTipoValorVenda = 2 Then
'        If lblEstornar.Caption <> "ESTORNO" Then
'            txtDataCompra.Text = Format(Date, "dd/mm/yyyy")
'            CriarNovoPedido
'        End If
'    Else
'        HabilitaObjetosVenda False
'   End If
'End If

'===========================SETANDO TABELA CONFIGURAÇÕES
If vUsarBalanca = "SIM" Then
    
    'pegar fabricante da balança
    Set oIni = New Ini
    oIni.Arquivo = appPathApp & "config.ini"
    vFabBalanca = oIni.LerTexto("FAB_BALANCA", "resposta")

    If vFabBalanca = "URANO" Then
        'abrir a porta serial da balança
        Set cCfg = sysConfig("PDV_BALANCA_PORTA")
        vPortaBalanca = cCfg.Value
        Set cCfg = Nothing
        
        If vPortaBalanca = "COM1" Then abrePortaSerial (1)
        If vPortaBalanca = "COM2" Then abrePortaSerial (2)
        If vPortaBalanca = "COM3" Then abrePortaSerial (3)
        If vPortaBalanca = "COM4" Then abrePortaSerial (4)
        If vPortaBalanca = "COM5" Then abrePortaSerial (5)
        If vPortaBalanca = "COM6" Then abrePortaSerial (6)
        If vPortaBalanca = "COM7" Then abrePortaSerial (7)
        If vPortaBalanca = "COM8" Then abrePortaSerial (8)
        If vPortaBalanca = "COM9" Then abrePortaSerial (9)
    ElseIf vFabBalanca = "TOLEDO" Then
        Call AbreConexao
    ElseIf vFabBalanca = "BALMAK" Then
        Call AbreConexao
    End If
    
    'pegar digito inicial
    Set cCfg = sysConfig("INICIAISBALANCA")
    vDigitoInicial = cCfg.Value
    Set cCfg = Nothing
    
    'pegar a quantidade de digitos
    Set cCfg = sysConfig("QTDEDIGITOSBALANCA")
    vQuantDigitos = cCfg.Value
    Set cCfg = Nothing
End If

Set cCfg = sysConfig("NOME_IMP_NFCE")
var_ImpNFCe = cCfg.Value
Set cCfg = Nothing

'se precisa pedi senha nas opções do menu avançado
Set cCfg = sysConfig("SEGURANCAAVANCADA")
varSegurancaAvancada = cCfg.Value
Set cCfg = Nothing

'usar o limite de compra do cliente
Set cCfg = sysConfig("LIMITARCOMPRA")
vLimitarCompra = cCfg.Value
Set cCfg = Nothing

'Cashback À Vista
Set cCfg = sysConfig("CASHBACKAV")
vCashbackAV = cCfg.Value
Set cCfg = Nothing

If vCashbackAV = "SIM" Then
    Set cCfg = sysConfig("CASHBACKVALORAV")
    vCashbackValorAV = cCfg.Value
    Set cCfg = Nothing
End If

'Cashback À Prazo
Set cCfg = sysConfig("CASHBACKAP")
vCashbackAP = cCfg.Value
Set cCfg = Nothing

If vCashbackAP = "SIM" Then
    Set cCfg = sysConfig("CASHBACKVALORAP")
    vCashbackValorAP = cCfg.Value
    Set cCfg = Nothing
End If

'Cashback validade
Set cCfg = sysConfig("CASHBACKVALIDADE")
vCashbackLimite = cCfg.Value
Set cCfg = Nothing


'se precisa pedi senha nas opções do menu avançado
Set cCfg = sysConfig("DECLARARRECEBEDOR")
vDeclararRecebedor = cCfg.Value
Set cCfg = Nothing

Set oCfg = sysConfig("CONF_FECHAMENTO_AV")
bFechAV = CBool(oCfg.Value)
Set oCfg = Nothing

Set oCfg = sysConfig("COPIAS_AV")
iCopiasAV = CInt(oCfg.Value)
Set oCfg = Nothing

Set oCfg = sysConfig("ENTREGA_AV")
bEntregaAV = CBool(oCfg.Value)
Set oCfg = Nothing

Set oCfg = sysConfig("IMP_AV")
vImprimirVendaAV = CBool(oCfg.Value)
Set oCfg = Nothing

If vImprimirVendaAV = True Then
    Set oCfg = sysConfig("CONF_IMPRESSAO_AV")
    vConfImprimirVendaAV = CBool(oCfg.Value)
    Set oCfg = Nothing

    Set oCfg = sysConfig("IMPRIMIR_AV")
    vTipoImpressaoVendaAV = CInt(oCfg.Value)
    Set oCfg = Nothing
    
'    If vTipoImpressaoVendaAV = 1 Then
'        varTipoImpressaoAV = Imprimir_Pedido
'    ElseIf vTipoImpressaoVendaAV = 2 Then
'        varTipoImpressaoAV = Imprimir_CupomSerrilha
'    ElseIf vTipoImpressaoVendaAV = 3 Then
'        varTipoImpressaoAV = Imprimir_CupomGuilhotina
'    End If
End If

Set oCfg = sysConfig("CONF_FECHAMENTO_AP")
bFechAP = CBool(oCfg.Value)
Set oCfg = Nothing

Set oCfg = sysConfig("COPIAS_AP")
iCopiasAP = CInt(oCfg.Value)
Set oCfg = Nothing

Set oCfg = sysConfig("ENTREGA_AP")
bEntregaAP = CBool(oCfg.Value)
Set oCfg = Nothing

Set oCfg = sysConfig("IMP_AP")
vImprimirVendaAP = CBool(oCfg.Value)
Set oCfg = Nothing

Set oCfg = sysConfig("CONF_IMPRESSAO_AP")
vConfImprimirVendaAP = CBool(oCfg.Value)
Set oCfg = Nothing

Set oCfg = sysConfig("IMPRIMIR_AP")
vTipoImpressaoVendaAP = CInt(oCfg.Value)
Set oCfg = Nothing

Set oCfg = sysConfig("CONF_FECHAMENTO_ORC")
bFechORC = CBool(oCfg.Value)
Set oCfg = Nothing

Set oCfg = sysConfig("COPIAS_ORC")
iCopiasORC = CInt(oCfg.Value)
Set oCfg = Nothing

Set oCfg = sysConfig("ENTREGA_ORC")
bEntregaORC = CBool(oCfg.Value)
Set oCfg = Nothing

Set oCfg = sysConfig("IMP_ORC")
bImprORC = CBool(oCfg.Value)
Set oCfg = Nothing

Set oCfg = sysConfig("CONF_IMPRESSAO_ORC")
bConfImprORC = CBool(oCfg.Value)
Set oCfg = Nothing

Set oCfg = sysConfig("IMPRIMIR_ORC")
iImprORC = CInt(oCfg.Value)
Set oCfg = Nothing

Set oCfg = sysConfig("TIPOIMPRESSAOPARCELAS")
vTipoParcelaImpressao = CInt(oCfg.Value)
Set oCfg = Nothing



If vConfImprimeNFCeLocal = "SIM" Then
    Set oCfg = sysConfig("CONFIMPNFCE")
    vNFCeConfImp = oCfg.Value
    Set oCfg = Nothing
    
    Set oCfg = sysConfig("IMPRIMINFCE")
    vNFCeImprimir = oCfg.Value
    Set oCfg = Nothing

    Set oCfg = sysConfig("CONFCPFNFCE")
    vNFCeConfCPF = oCfg.Value
    Set oCfg = Nothing
    
    Set oCfg = sysConfig("CONFPRAZONFCE")
    vNFCeConfPrazo = oCfg.Value
    Set oCfg = Nothing

    Set oCfg = sysConfig("COMBINARIMPNFCE")
    vNFCeCombinarImp = oCfg.Value
    Set oCfg = Nothing
End If


'tipo de empresa
Set cCfg = sysConfig("TIPO_EMPRESA")
tipoEmpresa = cCfg.Value
Set cCfg = Nothing

'se tem acrescimo vendendo no cartão a debito
Set cCfg = sysConfig("ACRESC_DEBITO")
varConfCartaodebito = cCfg.Value
Set cCfg = Nothing

'se tem acrescimo vendendo no cartão a Credito
Set cCfg = sysConfig("ACRESC_CREDITO")
varConfCartaoCredito = cCfg.Value
Set cCfg = Nothing

'Limitar os descontos aos clientes
Set oCfg = sysConfig("LIMITEDESCONTO")
vLimitarDesc = oCfg.Value
Set oCfg = Nothing

If vLimitarDesc = 1 Then
    'ativar a permissão da senha do gerente para liberar um valor de desconto maior
    Set oCfg = sysConfig("LIMITEGERENTE")
    vLiberacaoGerente = oCfg.Value
    Set oCfg = Nothing
    
    'Não aceitar desconto em vendas com cartão de débito
    Set oCfg = sysConfig("DESCCARTAODEDITO")
    vDescCartaoDebito = oCfg.Value
    Set oCfg = Nothing
    
    'Não aceitar desconto em vendas com cartão de crédito
    Set oCfg = sysConfig("DESCCARTAOCREDITO")
    vDescCartaoCredito = oCfg.Value
    Set oCfg = Nothing
    
End If
    
Set oCfg = sysConfig("TIPODESCONTO")
vTipoDesc = oCfg.Value
Set oCfg = Nothing
    
If vTipoDesc = "1" Then
    Set oCfg = sysConfig("DESC_AV")
    vValorDescFixoAV = oCfg.Value
    Set oCfg = Nothing
    
    Set oCfg = sysConfig("DESC_AP")
    vValorDescFixoAP = oCfg.Value
    Set oCfg = Nothing
ElseIf vTipoDesc = "2" Then
    Set oCfg = sysConfig("DESC_AV")
    vValorDescFixoAV = oCfg.Value
    Set oCfg = Nothing
    
    Set oCfg = sysConfig("DESC_AP")
    vValorDescFixoAP = oCfg.Value
    Set oCfg = Nothing
ElseIf vTipoDesc = "3" Then
    Set oCfg = sysConfig("DESC_MARGEM_AV1")
    vMargemDescGradual1 = oCfg.Value
    Set oCfg = Nothing
    
    Set oCfg = sysConfig("DESC_MARGEM_AV2")
    vMargemDescGradual2 = oCfg.Value
    Set oCfg = Nothing
    
    Set oCfg = sysConfig("DESC_MARGEM_AV3")
    vMargemDescGradual3 = oCfg.Value
    Set oCfg = Nothing
    
    Set oCfg = sysConfig("DESC_AV1")
    vValorDescGradualAV1 = oCfg.Value
    Set oCfg = Nothing
    
    Set oCfg = sysConfig("DESC_AV2")
    vValorDescGradualAV2 = oCfg.Value
    Set oCfg = Nothing
    
    Set oCfg = sysConfig("DESC_AV3")
    vValorDescGradualAV3 = oCfg.Value
    Set oCfg = Nothing
    
    Set oCfg = sysConfig("DESC_AP1")
    vValorDescGradualAP1 = oCfg.Value
    Set oCfg = Nothing
    
    Set oCfg = sysConfig("DESC_AP2")
    vValorDescGradualAP2 = oCfg.Value
    Set oCfg = Nothing

    Set oCfg = sysConfig("DESC_AP3")
    vValorDescGradualAP3 = oCfg.Value
    Set oCfg = Nothing
End If

Set oCfg = sysConfig("IDENT_PDV") 'tipo de login - ver se precisa se indentificar ao abrir o PDV
varLoginFunc = CInt(oCfg.Value)
Set oCfg = Nothing

Set oCfg = sysConfig("INICIAISETIQUETAS")
varTipoEtiqueta = oCfg.Value
Set oCfg = Nothing

Set oCfg = sysConfig("TIPOLOGIN")
varTipoLogin = oCfg.Value
Set oCfg = Nothing

'Não vender produtos zerados
Set oCfg = sysConfig("ESTOQUE_NEGATIVO")
bEstNeg = CBool(oCfg.Value)
Set oCfg = Nothing
  
If StatusBar1.Panels(3).Text = "" Then
   If varLoginFunc = 1 Then
      PDV.Hide
      PDV_Senha.Show vbModal
   End If
End If

'pegar a data de venda correta do pedido
If lblEstornar.Caption <> "ESTORNO" Then
    txtDataCompra.Text = Format(Date, "dd/mm/yyyy")
    LimparGrid_Pedido
    If varTipoValorVenda = 1 Then     'aqui
        CriarNovoPedido
    Else                               'aqui
        frmTipoVenda.Visible = True    'aqui
    End If                             'aqui
ElseIf lblEstornar.Caption = "ESTORNO" Then
    frmTipoVenda.Visible = False
End If

'IDENTIFICAR A caixa, saber que caixa vendeu
Set oCfg = sysConfig("IDENT_MAQ")
bIdentMaq = CBool(oCfg.Value)
Set oCfg = Nothing

If bIdentMaq Then
   If StatusBar1.Panels(2).Text = "" Then
        HabilitaObjetosVenda True
        frmMaquina.Visible = True
        'cboMaquina.SetFocus
        Exit Sub
   End If
End If

'IDENTIFICAR O VENDEDOR... limpar campo funcionario
'If sIdentPDV <> "" Then
'   If sIdentPDV = "2" Then
'      If lblEstornar.Caption <> "ESTORNO" Then
'         txtCodFuncAP.Text = ""
'         txtFuncAP.Text = ""
         'txtCodFuncAV.SetFocus
'      Else
'         txtRecebido.SetFocus   'VER DEPOIS
'      End If
'   Else
'      txtRecebido.SetFocus
'   End If
'End If

'variados - configurações de layout
lblInfoBusca.FontName = "Arial"
lblInfoBusca.FontSize = 10
lblInfoBusca.FontBold = True

lblInfoDebito.FontName = "Arial"
lblInfoDebito.FontSize = 10
lblInfoDebito.FontBold = True
lblInfoDebito.ForeColor = RGB(225, 0, 0)

frmAvancado.Visible = False
frmSenha.Visible = False
frmProdutoNaoCadastrado.Visible = False
frmProdutoAvulso.Visible = False
PesoF4 = False
vUsandoCashBack = False

If tipoEmpresa = 4 Then
    cmdOrçamento.Caption = "Consignado"
    'VerificarConsignado
Else
    cmdOrçamento.Caption = "Orçamento"
End If


'Liberar = False

'Tratar_Erro:
'   If Err.Number = 53 Then
'      MsgBox "Logomarca Inexistente!", vbInformation, "Aviso do Sistema"
'      Exit Sub
'   End If
Verificar_NFCe
'vFim = Format(Now, "HH:MM:SS")
'vTempo = vInicio - vFim
'MsgBox vTempo

Randomize Timer   'ver depois, peguei do form de pegar peso da balança
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT * " & _
       "FROM caixa_dia " & _
       "WHERE (caixa = '" & StatusBar1.Panels(2).Text & "') and codcaixa = " & StatusBar1.Panels(7).Text & " and caixa_dia.status = 0;"
Set r = dbData.OpenRecordset(sSQL)

If r.RecordCount <> 0 Then
    MsgBox "Seu Caixa ainda encontra-se aberto!", vbInformation, "Aviso do Sistema"
End If



'If txtCodPedido.Text = "" Then Exit Sub

'If lblEstornar.Caption <> "ESTORNO" Or lblEstornar.Caption <> "REIMPRESSÃO" Then
'   If txtTotalGeral <> "" And CCur(Val(txtTotalGeral)) <> 0 Then
      'Solicita confirmação do usuário
'      If ShowMsg("Existe uma compra em aberto. Deseja sair e cancelar a compra?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'         Cancel = 1
'         Exit Sub
'      End If
      
      'Apaga os itens se houver
'      If Grid.Rows > 1 Then
'         dbData.Execute "DELETE FROM pedidos_itens WHERE (cod_pedido = " & txtCodPedido.Text & ");"
'      End If
      
      'Apaga o pedido atual
'      dbData.Execute "DELETE FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ");"
'   End If
'End If

'Set moCombo = Nothing

'tirar o sistema da memoria
'fechaPortaSerial
KillProcess "PDV" ''DESATIVEI 22/06/2024
''ExitProcess 1
End Sub

Private Sub Grid_DblClick()
If Grid.Rows >= 2 Then
    cmdAlterar.Enabled = True
    txtCodItem.Text = ""
    txtCodItem.Text = (Grid.TextMatrix(Grid.Row, 1))
    txtCodProduto.Text = ""
    txtCodProduto.Text = (Grid.TextMatrix(Grid.Row, 2))
Else
    Exit Sub
End If
End Sub

Private Sub Label8_Click()
frmProduto.Visible = False
End Sub

Private Sub lblEstornar_Change()
If lblEstornar.Caption = "ESTORNO" Then
    If vTipoEdicao = "EDITAR" Then
        cmdOrçamento.Enabled = True
        cmdFinalizarPrazo.Enabled = False
        cmdFinalizarAvista.Enabled = False
    Else
        cmdOrçamento.Enabled = False
        cmdFinalizarPrazo.Enabled = True
        cmdFinalizarAvista.Enabled = True
    End If
Else
    cmdOrçamento.Enabled = True
End If
End Sub

Private Sub lstBusca_ItemClick(ByVal Item As MSComctlLib.ListItem)
If Item Is Nothing Then Exit Sub
Item.Selected = True
Item.EnsureVisible
End Sub

Private Sub lstBusca_KeyPress(KeyAscii As Integer)
txtCodProduto.Text = lstBusca.SelectedItem
txtCodBarra.Text = Format(lstBusca.SelectedItem.ListSubItems.Item(1).Text, "0000")
lblDesc.Caption = lstBusca.SelectedItem.ListSubItems.Item(2).Text
lstBusca.Visible = False
txtCodBarra_Validate False
End Sub

Private Sub lstBusca_KeyUp(KeyCode As Integer, Shift As Integer)
   'If KeyCode = 13 Then
   '   txtCodProduto.Text = lstBusca.SelectedItem
   '   txtCodBarra.Text = lstBusca.SelectedItem.ListSubItems.Item(1).Text
   '   lblDesc.Caption = lstBusca.SelectedItem.ListSubItems.Item(2).Text
      'If txtCodBarra.Text <> "" Then txtCodBarra_Change
   'End If
End Sub

Private Sub mskCPF_GotFocus()
mskCPF.SelStart = 0
mskCPF.SelLength = Len(mskCPF.Text)
End Sub



Private Sub mskInicio_LostFocus()
If cboPrazo.Text = "" Then Exit Sub

If txtEntrada.Text = "0,00" Or txtEntrada.Text = "" Or Not IsDate(mskInicio) = True Then
   mskTermino.Text = Format(DateAdd("m", Val(cboQuantParc.Text) - 1, mskInicio.Text), "dd/mm/yy")
Else
   mskTermino.Text = Format(DateAdd("m", Val(cboQuantParc.Text), mskInicio.Text), "dd/mm/yy")
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

Private Sub mskTermino_Change()
   If Not IsDate(mskTermino.Text) Then Exit Sub
   mskTermino.Mask = "##/##/##"
End Sub

Private Sub mskTermino_GotFocus()
SelectControl mskTermino
End Sub

Private Sub optAscrescPorc_Click()
Calcular_Desconto
If txtAcresc.Enabled = True Then txtAcresc.SetFocus
End Sub

Private Sub optAscrescRS_Click()
Calcular_Desconto
txtAcresc.SetFocus
End Sub

Private Sub optDescPorc_Click()
Calcular_Desconto
If lblEstornar.Caption <> "REIMPRESSÃO" Then
    If frmVendaFechamento.Visible = True Then txtDesc.SetFocus
End If
End Sub

Private Sub optDescRS_Click()
Calcular_Desconto
If txtDesc.Enabled = True Then txtDesc.SetFocus
End Sub

Private Sub StatusBar1_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
   Select Case Panel.Index
      Case 1
         Exit Sub
      Case 2
         frmMaquina.Visible = True
         cboMaquina.SetFocus
      Case 3
         Exit Sub
   End Select
End Sub






Private Sub Timer2_Timer()
Timer2.Enabled = False
End Sub

Private Sub timerBackup_Timer()
'MsgBox "Timer ativo"
Dim DataHora As Date, xCaminhoBK As String
Dim nomeArquivoBK As String
Dim IniciouProcesso As Boolean

   'picAguarde.Visible = False
   DoEvents
   mensagemErro = ""
   iRetorno = False
   IniciouProcesso = False
   
   If IniciouProcesso = False And 1 = 2 Then
        sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
        Set r = dbData.OpenRecordset(sSQL)
        
        If Not r.EOF Then
           dirXML = IIf(Right(r!DiretorioXML, 1) = "\", r!DiretorioXML, r!DiretorioXML & "\")
        Else
           Exit Sub
        End If
        
        xCaminhoBK = dirXML & "backup"
        
        nomeArquivoBK = Retira(r!CNPJ, ".-/ ", UM_A_UM) & ".rar"
        
        DataHora = Now
        
        If Not Existe(xCaminhoBK & "\" & nomeArquivoBK) Then Exit Sub
        
        If Vazio(r!BackupDataHora) Then
            IniciouProcesso = True
           'picAguarde.Visible = True
           DoEvents
           iRetorno = GoogleEnviarArquivo(xCaminhoBK & "\" & nomeArquivoBK)
           'picAguarde.Visible = False
           DoEvents
        ElseIf Day(r!BackupDataHora) < Day(DataHora) Then
            IniciouProcesso = True
           'picAguarde.Visible = True
           DoEvents
           iRetorno = GoogleEnviarArquivo(xCaminhoBK & "\" & nomeArquivoBK)
           'picAguarde.Visible = False
           DoEvents
        ElseIf Format(DataHora, "hh:mm:ss") = CDate("12:30:00") Then
            IniciouProcesso = True
           'picAguarde.Visible = True
           DoEvents
           iRetorno = GoogleEnviarArquivo(xCaminhoBK & "\" & nomeArquivoBK)
           'picAguarde.Visible = False
           DoEvents
        Else
           Exit Sub
        End If
        
        If iRetorno Then
           sSQL = "UPDATE empresa SET BackupDataHora = " & FdthrSQL(DataHora)
           SQLExecuta sSQL
        End If
    End If
    
    IniciouProcesso = False
End Sub


Private Sub tmrDebito_Timer()
   FlashColor
End Sub

Private Sub txtAcresc_GotFocus()
   SelectControl txtAcresc
End Sub

Private Sub txtAcresc_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
   
If KeyAscii = 13 Then
'      txtDescDinheiro.Visible = True
'      txtDescDinheiro.Text = ""
      txtRecebido.SetFocus
End If
End Sub

Private Sub txtAcresc_LostFocus()
On Error GoTo erro
   
If txtAcresc.Text = "" Or txtSubtotal.Text = "" Then
   txtAcresc.Text = FormatNumber(0, 2)
   SelectControl txtAcresc
   Exit Sub
End If

Calcular_Desconto
txtAcresc.Text = FormatNumber(txtAcresc.Text, 2)
Exit Sub
   
erro:
   MsgBox "O valor digitado é inválido!", vbExclamation, "Aviso do Sistema"
   txtAcresc.Text = 0
End Sub

Public Function VerificaValor(vTecla As Integer) As Integer
   ' Função para permitir apenas a digitação de valores
   Select Case vTecla
      Case 8, 44, 48 To 57
         VerificaValor = vTecla
      Case 46
         VerificaValor = 44
      Case Else
         VerificaValor = 0
   End Select
End Function

Private Sub txtAcrescDinheiro_GotFocus()
SelectControl txtAcrescDinheiro
End Sub


Private Sub txtAcrescDinheiro_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
If KeyAscii = 13 Then
    txtAcrescDinheiro_LostFocus
End If
End Sub


Private Sub txtAcrescDinheiro_LostFocus()
Dim ValueTotal As Double
Dim ValueDiscount As Double
Dim Percent As Double

ValueTotal = txtSubtotal.Text

If txtAcrescDinheiro.Text = "" Then
    ValueDiscount = 0
Else
    ValueDiscount = txtAcrescDinheiro.Text
End If

Percent = (ValueDiscount / ValueTotal) * 100
txtAcresc.Text = FormatNumber(Percent, 2)
txtAcrescDinheiro.Visible = False
txtAcresc_LostFocus
If frmVendaFechamento.Visible = True Then txtRecebido.SetFocus
End Sub

Private Sub txtCodBarra_GotFocus()
If txtCodProduto.Text = "" Then
    txtTotal.Text = Format(0, ocMONEY)
    txtQuant.Text = "0"
    txtValor.Text = Format(0, ocMONEY)
End If
End Sub

Private Sub txtCodBarra_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub


Private Sub txtCodBarra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
      SendKey ocKEYTAB
End If
End Sub

Private Sub txtCodBarra_Validate(Cancel As Boolean)
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodPedido.Text = "" Then Exit Sub

lstBusca.Visible = False

Dim varSeVendeNegativo As String
If bEstNeg = False Then
    varSeVendeNegativo = " AND (produtos.quant_estoque > 0)"
Else
    varSeVendeNegativo = " "
End If

ValidarBusca:
If txtCodBarra.Text <> "" And IsNumeric(txtCodBarra.Text) = True Then          'código de barra
    
    If Left(txtCodBarra.Text, 1) = "2" And Len(txtCodBarra.Text) = 13 Then      'código de barra de peso ou inicia com 2
    
        sSQL = "SELECT DISTINCT produtos.codigo AS vCodProduto, produtos.cod_barra, produtos.quant_estoque as vQuant, produtos.UNID_MEDIDA as vUnid, " & _
            "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
            "FROM produtos WHERE produtos.cod_barra = '" & txtCodBarra.Text & "'  " & varSeVendeNegativo & " " & _
            "ORDER BY produtos.codigo;"
        Set r = dbData.OpenRecordset(sSQL)
    
        If Not r.BOF Then
            txtCodProduto.Text = r("vCodProduto")
            'txtUnidMed.Text = r("vUnid")
            varCodBarra = txtCodBarra.Text
        Else
            'codigo da balança===========
            If varTipoEtiqueta = "5" Then
                varCodBarra = Format(Mid(txtCodBarra, 2, 5), "00000")
                'varPeso = Mid(txtCodBarra, 9, 12)
            ElseIf varTipoEtiqueta = "4" Then
                varCodBarra = Format(Mid(txtCodBarra, 2, 4), "00000")
                'varPeso = Mid(txtCodBarra, 8, 5)
            ElseIf varTipoEtiqueta = "7" Then
                varCodBarra = Format(Mid(txtCodBarra, 4, 4), "00000")
                'varPeso = Mid(txtCodBarra, 8, 5)
            ElseIf varTipoEtiqueta = "2" Then
                varCodBarra = Format(Mid(txtCodBarra, 2, 4), "00000")
            End If
            '================================

            
            sSQL = "SELECT DISTINCT produtos.codigo AS vCodProduto, produtos.cod_barra, produtos.quant_estoque as vQuant, produtos.UNID_MEDIDA as vUnid, " & _
                "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
                "FROM produtos WHERE produtos.cod_barra = '" & varCodBarra & "' AND (produtos.ativo = 1) " & varSeVendeNegativo & " " & _
                "ORDER BY produtos.codigo;"
                
            Set r = dbData.OpenRecordset(sSQL)
    
            If Not r.BOF Then
                txtCodProduto.Text = r("vCodProduto")
                'txtUnidMed.Text = r("vUnid")
            Else
                frmProdutoNaoCadastrado.Visible = True
                
                For i = 1 To 3  'aviso sonoro
                    Beep
                Next
                
                txtCodBarra.Enabled = False
                txtQuant.Enabled = False
                'MsgBox "Produto não localizado!", vbInformation, "Aviso do Sistema"
                'txtCodBarra.Text = ""
                'If txtCodBarra.Enabled = True Then txtCodBarra.SetFocus
                Exit Sub
            End If
        End If
    Else                                                                        'produto cod_barra normal
    
        If Len(txtCodBarra.Text) < 13 Then
            If Len(txtCodBarra.Text) < 6 Then
                txtCodBarra.Text = Format(txtCodBarra.Text, "00000")
            Else
                txtCodBarra.Text = txtCodBarra.Text
            End If
        End If
        
        sSQL = "SELECT DISTINCT produtos.codigo AS vCodProduto, produtos.cod_barra, produtos.quant_estoque as vQuant, produtos.UNID_MEDIDA as vUnid, " & _
            "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
            "FROM produtos WHERE produtos.cod_barra = '" & txtCodBarra.Text & "' AND (ATIVO = 1) " & _
            "ORDER BY produtos.codigo;"
            '" & varSeVendeNegativo & "
            'Debug.Print sSQL
        Set r = dbData.OpenRecordset(sSQL)
        
            If Not r.BOF Then
                If r("vCodProduto") <> 1 Then
                    If bEstNeg = False Then         'não vende com estoque negativo
                        If r("vQuant") <= 0 Then
                            'txtQuant.Enabled = False
                            MsgBox "PRODUTO COM ESTOQUE INSUFICIENTE", vbInformation, "Aviso do Sistema"
                            txtCodBarra.Text = ""
                            If txtCodBarra.Enabled = True Then txtCodBarra.SetFocus
                            'txtQuant.Enabled = True
                            Exit Sub
                        End If
                    End If
                End If
            
                txtCodProduto.Text = r("vCodProduto")
                'txtUnidMed.Text = r("vUnid")
                varCodBarra = txtCodBarra.Text
            Else
                frmProdutoNaoCadastrado.Visible = True
                txtCodBarra.Enabled = False
                txtQuant.Enabled = False
                'MsgBox "Produto não localizado!", vbInformation, "Aviso do Sistema"
                'txtCodBarra.Text = ""
                'If txtCodBarra.Enabled = True Then txtCodBarra.SetFocus
                If frmProdutoNaoCadastrado.Visible = True Then Exit Sub
            End If
    End If
    
    Adicionar_Produto
    MostrarGrid_Produtos
    
'    If Grid.Rows >= 2 Then
'        If Not r.EOF Then
'            vPedirPeso = Abs(CBool(r("PEDIRPESO")))
'        Else
'            vPedirPeso = False
'        End If
'    End If

    If vPedirPeso = True Then
        If frmVendaFechamento.Visible = False Then
            If frmProdutoNaoCadastrado.Visible = False Then Grid_DblClick
        End If
    End If
    
    
    If vPedirPeso = False Then txtCodBarra.Text = ""
Else                                                                            'produtos digitados sem codigo de barra
      
    Dim ItemLst As ListItem
    Dim fGrid As Object
    Dim bCancel As Boolean
    Dim vProd() As String
    Dim rPos As RECT
    Dim lLft As Long, lTop As Long
            
  If txtCodBarra.Text <> "" Then
    'carrega o label
    DoEvents
    lblInfoBusca.Visible = True
    lblInfoBusca.Refresh
    Screen.MousePointer = vbHourglass
    
    Dim vUltimoValorVenda As String     '===================TER QUE COLOCAR DEPOIS PARA TODOS OS TIPOS DE VENDAS
    vUltimoValorVenda = " (SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) "
         
    'carrega a consulta
    If varTipoValorVenda = 1 Then
        sSQL = "SELECT DISTINCT produtos.codigo AS var_cod, produtos.ref AS var_ref, produtos.tamanho AS var_tam, " & _
        "produtos.fabricante AS var_fab, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, " & _
        "produtos.quant_estoque AS var_quant, produtos.PRATELEIRA AS var_Local, " & _
        "" & vUltimoValorVenda & " AS venda " & _
        "FROM produtos WHERE (descricao LIKE '%" & txtCodBarra.Text & "%') AND (produtos.ativo = 1) and " & vUltimoValorVenda & " > 0 " & varSeVendeNegativo & " " & _
        "ORDER BY descricao;"
    ElseIf varTipoValorVenda = 2 Then
        If TipoValorVenda = "VV" Then
            sSQL = "SELECT DISTINCT produtos.codigo AS var_cod, produtos.ref AS var_ref, produtos.tamanho AS var_tam, " & _
            "produtos.fabricante AS var_fab, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, " & _
            "produtos.quant_estoque AS var_quant, produtos.PRATELEIRA AS var_Local,  " & _
            "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
            "FROM produtos WHERE (descricao LIKE '%" & txtCodBarra.Text & "%') AND (produtos.ativo = 1) " & varSeVendeNegativo & "" & _
            "ORDER BY descricao;"
        ElseIf TipoValorVenda = "VP" Then
            sSQL = "SELECT DISTINCT produtos.codigo AS var_cod, produtos.ref AS var_ref, produtos.tamanho AS var_tam, " & _
            "produtos.fabricante AS var_fab, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, " & _
            "produtos.quant_estoque AS var_quant, produtos.PRATELEIRA AS var_Local,  " & _
            "(SELECT TOP 1 Produtos_Precos.VALOR_VP FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
            "FROM produtos WHERE (descricao LIKE '%" & txtCodBarra.Text & "%') AND (produtos.ativo = 1) " & varSeVendeNegativo & "" & _
            "ORDER BY descricao;"
        ElseIf TipoValorVenda = "AV" Then
            sSQL = "SELECT DISTINCT produtos.codigo AS var_cod, produtos.ref AS var_ref, produtos.tamanho AS var_tam, " & _
            "produtos.fabricante AS var_fab, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, " & _
            "produtos.quant_estoque AS var_quant, produtos.PRATELEIRA AS var_Local,  " & _
            "(SELECT TOP 1 Produtos_Precos.VALOR_AV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
            "FROM produtos WHERE (descricao LIKE '%" & txtCodBarra.Text & "%') AND (produtos.ativo = 1) " & varSeVendeNegativo & "" & _
            "ORDER BY descricao;"
        ElseIf TipoValorVenda = "AP" Then
            sSQL = "SELECT DISTINCT produtos.codigo AS var_cod, produtos.ref AS var_ref, produtos.tamanho AS var_tam, " & _
            "produtos.fabricante AS var_fab, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, " & _
            "produtos.quant_estoque AS var_quant, produtos.PRATELEIRA AS var_Local,  " & _
            "(SELECT TOP 1 Produtos_Precos.VALOR_AP FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
            "FROM produtos WHERE (descricao LIKE '%" & txtCodBarra.Text & "%') AND (produtos.ativo = 1) " & varSeVendeNegativo & "" & _
            "ORDER BY descricao;"
        Else
            sSQL = "SELECT DISTINCT produtos.codigo AS var_cod, produtos.ref AS var_ref, produtos.tamanho AS var_tam, " & _
            "produtos.fabricante AS var_fab, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, " & _
            "produtos.quant_estoque AS var_quant, produtos.PRATELEIRA AS var_Local,  " & _
            "(SELECT TOP 1 Produtos_Precos.VALOR_VP FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
            "FROM produtos WHERE (descricao LIKE '%" & txtCodBarra.Text & "%') AND (produtos.ativo = 1) " & varSeVendeNegativo & "" & _
            "ORDER BY descricao;"
        End If
    End If
            Set r = dbData.OpenRecordset(sSQL)
            
            If r.EOF Then
               vPedirPeso = False
            End If
      
  'End If
      'carrega a caixa de buscar
         GetWindowRect txtCodBarra.hwnd, rPos
         lLft = rPos.Left * Screen.TwipsPerPixelX - 160
         lTop = rPos.Top * Screen.TwipsPerPixelY + txtCodBarra.Height
       
       'tipos de empresa os itens exibidos no grid mudam
         If tipoEmpresa = 4 Then
            Set fGrid = New BuscaGrid
         ElseIf tipoEmpresa = 5 Then
            Set fGrid = New BuscaGrid_Automoveis
         Else
            Set fGrid = New BuscaGrid_Comum
         End If
         
         Load fGrid
         LockWindowUpdate fGrid.lstBusca.hwnd
         
      'If txtCodBarra.Text <> "" Then
         If Not r Is Nothing Then
            Do While Not r.EOF
               'primeira coluna
               Set ItemLst = fGrid.lstBusca.ListItems.Add(, , r("var_cod"))
               'segunda e terceira coluna, que são sub itens da coluna 1
               ItemLst.SubItems(1) = ValidateNull(r("var_codbarra"))
               
               If tipoEmpresa = 4 Then   'sapataria/vestuario
                  ItemLst.SubItems(2) = ValidateNull(r("var_desc"))
                  ItemLst.SubItems(3) = ValidateNull(r("var_ref"))
                  ItemLst.SubItems(4) = ValidateNull(r("var_tam"))
                  ItemLst.SubItems(5) = ValidateNull(r("var_fab"))
                  If Not IsNull(r("var_quant")) Then ItemLst.SubItems(6) = ValidateNull(r("var_quant"))
                  If Not IsNull(r("venda")) Then ItemLst.SubItems(7) = Format(ValidateNull(r("venda")), ocMONEY)
               ElseIf tipoEmpresa = 5 Then   'autopeças
                  ItemLst.SubItems(2) = ValidateNull(r("var_desc")) & " /  " & ValidateNull(r("var_fab"))
                  If Not IsNull(r("var_quant")) Then ItemLst.SubItems(3) = ValidateNull(r("var_quant"))
                  If Not IsNull(r("venda")) Then ItemLst.SubItems(6) = Format(ValidateNull(r("venda")), ocMONEY)
                  If Not IsNull(r("var_local")) Then ItemLst.SubItems(5) = ValidateNull(r("var_local"))
                      'Compartibilidade
                        Dim sSQL_Comp As String
                        Dim var_Comp As String
                        Dim rS2 As ADODB.Recordset
                        
                        sSQL_Comp = "Select MODELO, ANO From PRODUTOS_COMP Where COD_PRODUTO = " & r("var_cod")
                        Set rS2 = dbData.OpenRecordset(sSQL_Comp)
                        
                        Do While Not rS2.EOF
                        var_Comp = var_Comp & rS2!Modelo & "(" & rS2!Ano & "),  "
                        rS2.MoveNext
                        Loop
                        
                        If Not IsNull(var_Comp) Then ItemLst.SubItems(4) = var_Comp
                        var_Comp = ""
                  
               Else     'outros tipos de empresas
                  ItemLst.SubItems(2) = ValidateNull(r("var_desc")) & " /  " & ValidateNull(r("var_fab"))
                  If Not IsNull(r("var_quant")) Then ItemLst.SubItems(3) = ValidateNull(r("var_quant"))
                  If Not IsNull(r("venda")) Then ItemLst.SubItems(4) = Format(ValidateNull(r("venda")), ocMONEY)
               End If
               
               r.MoveNext
            Loop
            
            If r.State <> 0 Then r.Close
            Set r = Nothing
         End If
      'End If
      
         lblInfoBusca.Visible = False
         Screen.MousePointer = vbDefault
         
         LockWindowUpdate 0
         fGrid.Move lLft, lTop
         fGrid.Show vbModal
         
         bCancel = fGrid.Cancelled
         vProd = fGrid.InfoProduct
         
         Unload fGrid
         Set fGrid = Nothing
         
         If Not bCancel Then
            txtCodProduto.Text = vProd(1) 'lstBusca.SelectedItem
            txtCodBarra.Text = vProd(2)   'lstBusca.SelectedItem.ListSubItems.Item(1).Text
            lblDesc.Caption = vProd(3)    'lstBusca.SelectedItem.ListSubItems.Item(2).Text
            
            Cancel = True
            GoTo ValidarBusca
         End If
      'End If
   'End If
  Else
    Adicionar_Produto
    MostrarGrid_Produtos
    
    If vPedirPeso = True Then
        If frmVendaFechamento.Visible = False Then
            If frmProdutoNaoCadastrado.Visible = False Then Grid_DblClick
        End If
    End If
    
End If
End If

If vPedirPeso = False Then
   txtCodBarra = ""
   lblDesc.Caption = ""
   txtQuant.Text = "0"
   If txtCodBarra.Enabled = True Then txtCodBarra.SetFocus
   Cancel = True
   Exit Sub
Else
    txtCodBarra.Enabled = False
    If frmProdutoNaoCadastrado.Visible = False Then txtQuant.SetFocus
    SelectControl txtQuant
End If
End Sub

Private Sub TxtCodCliente_Change()
'Dim sSQL As String
'Dim r As ADODB.Recordset

Dim oCfg As ConfigItem
Dim bBloq As Boolean
Dim iDiasBloq As Integer

tmrDebito.Enabled = False
lblInfoDebito.Visible = False
Cliente_Debito = False

If lblEstornar.Caption = "ESTORNO" Or lblEstornar.Caption = "REIMPRESSÃO" Then
    If txtCodCliente.Text = "" Then Exit Sub

   sSQL = "SELECT codigo, nome FROM cliente WHERE (codigo = " & txtCodCliente.Text & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then
      cboCliente.Text = r("nome")
      txtCodCliente.Text = r("codigo")
   End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End If

Set oCfg = sysConfig("BLOQUEIAR_CLIENTE")
bBloq = CBool(oCfg.Value)
Set oCfg = Nothing

If bBloq Then
    'ver quantos dias apos vencimento é para bloqueiar
    Set oCfg = sysConfig("DIAS_BLOQUEIO")
    If oCfg.Value = "" Then
        iDiasBloq = 0
    Else
        iDiasBloq = CInt(oCfg.Value)
    End If
    
    If txtCodCliente.Text <> "1" Then
           Dim var_Venc As Date
           Dim var_SomaDatas As Long
           Dim var_Dias As Integer
           
           var_Dias = iDiasBloq
           var_SomaDatas = 0
            
            If txtCodCliente.Text = "" Then Exit Sub
        
           sSQL = "SELECT TOP 1 parcelas.data, parcelas.cod_pedido, pedidos.cod_pedido, cliente.codigo, pedidos.cod_cliente, parcelas.status " & _
              "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente INNER JOIN parcelas ON parcelas.cod_pedido = pedidos.cod_pedido " & _
              "WHERE (cliente.codigo = " & txtCodCliente.Text & ") AND (parcelas.status = 0) ORDER BY parcelas.data;"
           
           Set r = dbData.OpenRecordset(sSQL)
           
           Do While Not r.EOF       'desabilitei para teste
              var_Venc = r("data")
              var_SomaDatas = DateDiff("d", var_Venc, Date)
              r.MoveNext
           Loop
           
           If r.State <> 0 Then r.Close
           Set r = Nothing
           
         If var_SomaDatas >= var_Dias Then
             'ShowMsg "Favor encaminhar o cliente a gerencia!", vbInformation
             'MsgBox ("ESTE CLIENTE POSSUI PARCELA(S) VENCIDA(S) COM " & var_SomaDatas & " DIAS " & vbCrLf & "O LIMITE É DE " & var_Dias & " DIAS"), vbInformation
             
             lblInfoDebito.Visible = True
             tmrDebito.Enabled = True
             Cliente_Debito = True      'Atribui o cliente em débito
         Else
             Cliente_Debito = False
         End If
    Else
        Cliente_Debito = False
    End If
End If

Set oCfg = Nothing
End Sub
Private Sub txtCodFuncAP_Change()
Dim sSQL As String
Dim r As ADODB.Recordset
If txtCodFuncAP.Text <> "" Then txtCodFunc.Text = txtCodFuncAP.Text Else txtFuncAP.Text = ""
If txtCodFuncAP.Text = "" Then Exit Sub

If varLoginFunc = "2" Then txtFuncAP.Text = ""
sSQL = "SELECT codigo, nome, sobrenome FROM funcionario WHERE (codigo = " & txtCodFuncAP.Text & ");"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then
    txtFuncAP.Text = r("nome")
    'txtFuncAP.Text = r("nome") & " " & r("sobrenome")
Else
    txtFuncAP.Text = ""
End If
If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub txtCodFuncAP_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
If KeyAscii = 13 Then
      cmdFinalizar_Click
End If
End Sub

Private Sub txtCodItem_Change()
Mostrar_Produto_Alterar
End Sub

Private Sub txtCodPedido_Change()
If lblEstornar.Caption = "ESTORNO" Then
    'Dim varHoraParc As String

    'abre o caixa
    'Abrir_Caixa
    
    'mudar status do pedido
    MudarPedidoReaberto
    
    'Mostra os produtos do pedido
    MostrarGrid_Produtos
    
    'Retornar a quantidade de produtos ao estoque
    Retorna_Produtos_Estoque
    
    'Apaga as parcelas do pedido
    If lblTipoPedido.Caption <> "ORÇAMENTO" Then
        sSQL = "SELECT TOP 1 HORA FROM parcelas WHERE (cod_pedido = " & txtCodPedido.Text & ") ORDER BY CODIGO;"
        Set r = dbData.OpenRecordset(sSQL)
        
        If Not r.EOF Then
            txtHoraCompra.Text = Format(ValidateNull(r("HORA")), ocHRMN)
        End If
        
        dbData.Execute "DELETE FROM parcelas WHERE (cod_pedido = " & txtCodPedido.Text & ");"
    End If
    'criar log
    'Autonumeracao_LOG
    'execSQL "INSERT INTO LOG (CODIGO, COD_PEDIDO, JANELA, ACAO, DATA, HORA, FUNCIONARIO) VALUES(" & X & ", " & txtCodPedido.Text & ", 'PDV', 'EXCLUIR', #" & Format(Date, "dd/mm/yy") & "#, #" & Format(Now, "hh:mm") & "#, 'maria' )"
    
    'cmdAbrirVenda.Visible = False
    'cmdExcluirVenda.Visible = False
    'cmdCancelarEstorno.Visible = False
    If vTipoEdicao = "EDITAR" Then
        cmdFinalizarAvista.Enabled = False
        cmdFinalizarPrazo.Enabled = False
        cmdOrçamento.Enabled = True
    Else
        cmdFinalizarAvista.Enabled = True
        cmdFinalizarPrazo.Enabled = True
        cmdOrçamento.Enabled = False
    End If
    cmdRemover.Enabled = True
    cmdAvancado.Enabled = False
    cmdCancelarPedido.Enabled = False
    cmdFechar.Enabled = True
    HabilitaObjetosVenda False
    'txtCodBarra.SetFocus
ElseIf lblEstornar.Caption = "REIMPRESSÃO" Then
   MostrarGrid_Produtos
   'If lblEstornar.Caption = "REIMPRESSÃO" Then
        cboTipoPgto.Text = ""
        Abrir_Pedido_Reimpressao
   'End If
End If
End Sub

Private Sub txtCodProduto_Change()
If txtCodProduto.Text = "" Then Exit Sub
If cmdAlterar.Enabled = True Then Exit Sub
End Sub

Private Sub txtCodUsuario_Change()
Dim sSQL As String
Dim r As ADODB.Recordset

If txtCodUsuario.Text = "" Then Exit Sub

sSQL = "SELECT codigo, nivel FROM usuario WHERE (codigo = " & txtCodUsuario.Text & ");"
Set r = dbData.OpenRecordset(sSQL)

If Not r.EOF Then
    vCodUsuario = r("codigo")
Else
    vCodUsuario = 0
End If

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub txtDesc_GotFocus()
SelectControl txtDesc
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)

If KeyAscii = 13 Then
'      txtDescDinheiro.Visible = True
'      txtDescDinheiro.Text = ""
      txtRecebido.SetFocus
End If
End Sub

Private Sub txtDesc_LostFocus()
'On Error GoTo erro
Dim vDesc As Double
   
If txtDesc.Text = "" Or txtSubtotal.Text = "" Then
    txtDesc.Text = FormatNumber(0, 2)
    vDesc = 0
Else
    If optDescRS.Value = True Then
        Dim ValueTotal As Double
        Dim ValueDiscount As Double
        Dim Percent As Double
        
        ValueTotal = txtSubtotal.Text
        
        If txtDesc.Text = "" Then
            ValueDiscount = 0
        Else
            ValueDiscount = txtDesc.Text
        End If
        
        Percent = (ValueDiscount * 100) / ValueTotal
        vDesc = FormatNumber(Percent, 2)
    ElseIf optDescPorc.Value = True Then
        vDesc = txtDesc.Text
    End If
End If

If vLiberacaoGerente = "SIM" Then
    Dim fLib As LiberarVenda
    Dim bCancel As Boolean
    Dim lGerente As Long
End If


If vTipoDesc = "1" Then     'desconto manual
    If vLimitarDesc = 1 Then    'se o ckeck de limitar o desconto estiver ativo
        If cboTipoPgto.Text = "À VISTA" Then
            If vDesc > vValorDescFixoAV Then
                If vUsandoCashBack = False Then
                    If vLiberacaoGerente = "SIM" Then
                        Set fLib = New LiberarVenda
                        Load fLib
                        
                        fLib.Show vbModal
                        bCancel = fLib.Cancelled
                        lGerente = fLib.Gerente
                              
                        Unload fLib
                        Set fLib = Nothing
                        
                        If bCancel Then
                            txtDesc.Text = "0,00"
                            txtDesc.SetFocus
                            Exit Sub
                        Else
                            txtRecebido.SetFocus
                        End If
                    Else
                        MsgBox "Desconto maior que o permitido pela empresa!", vbInformation, "Aviso do Sistema"
                        txtDesc.Text = FormatNumber(0, 2)
                    End If
                Else
                    txtRecebido.SetFocus
                End If
            End If
        ElseIf cboTipoPgto.Text = "À PRAZO" Then
            If vDesc > vValorDescFixoAP Then
                If vUsandoCashBack = False Then
                    If vLiberacaoGerente = "SIM" Then
                        Set fLib = New LiberarVenda
                        Load fLib
                        
                        fLib.Show vbModal
                        bCancel = fLib.Cancelled
                        lGerente = fLib.Gerente
                              
                        Unload fLib
                        Set fLib = Nothing
                        
                        If bCancel Then
                            'Unload LiberarVenda
                            txtDesc.Text = "0,00"
                            txtDesc.SetFocus
                            Exit Sub
                        Else
                            txtRecebido.SetFocus
                        End If
                    Else
                        MsgBox "Desconto maior que o permitido pela empresa!", vbInformation, "Aviso do Sistema"
                        txtDesc.Text = FormatNumber(0, 2)
                    End If
                Else
                    txtRecebido.SetFocus
                End If
            End If
        Else
            If frmVendaFechamento.Visible = True Then txtRecebido.SetFocus
        End If
    Else
    End If
ElseIf vTipoDesc = "2" Then     'desconto fixo
    If vLimitarDesc = 1 Then    'se o ckeck de limitar o desconto estiver ativo
        If cboTipoPgto.Text = "À VISTA" Then
            If vDesc > vValorDescFixoAV Then
                If vUsandoCashBack = False Then
                    If vLiberacaoGerente = "SIM" Then
                        Set fLib = New LiberarVenda
                        Load fLib
                        
                        fLib.Show vbModal
                        bCancel = fLib.Cancelled
                        lGerente = fLib.Gerente
                              
                        Unload fLib
                        Set fLib = Nothing
                        
                        If bCancel Then
                            'Unload LiberarVenda
                            txtDesc.Text = "0,00"
                            txtDesc.SetFocus
                            Exit Sub
                        Else
                            txtRecebido.SetFocus
                        End If
                    Else
                        MsgBox "Desconto maior que o permitido pela empresa!", vbInformation, "Aviso do Sistema"
                        txtDesc.Text = FormatNumber(0, 2)
                    End If
                Else
                    txtRecebido.SetFocus
                End If
            End If
        ElseIf cboTipoPgto.Text = "À PRAZO" Then
            If vDesc > vValorDescFixoAP Then
                If vUsandoCashBack = False Then
                        If vLiberacaoGerente = "SIM" Then
                            Set fLib = New LiberarVenda
                            Load fLib
                            
                            fLib.Show vbModal
                            bCancel = fLib.Cancelled
                            lGerente = fLib.Gerente
                                  
                            Unload fLib
                            Set fLib = Nothing
                            
                            If bCancel Then
                                'Unload LiberarVenda
                                txtDesc.Text = "0,00"
                                txtDesc.SetFocus
                                Exit Sub
                            Else
                                txtRecebido.SetFocus
                            End If
                        Else
                            MsgBox "Desconto maior que o permitido pela empresa!", vbInformation, "Aviso do Sistema"
                            txtDesc.Text = FormatNumber(0, 2)
                        End If
                    Else
                        txtRecebido.SetFocus
                    End If
            End If
        Else
         txtRecebido.SetFocus
        End If
    Else
    End If
ElseIf vTipoDesc = "3" Then     'desconto gradativo
    Dim vValorDescGradual As Currency
    If cboTipoPgto.Text = "À VISTA" Then
        If vEtapa = 1 Then
            vValorDescGradual = vValorDescGradualAV1
        ElseIf vEtapa = 2 Then
            vValorDescGradual = vValorDescGradualAV2
        ElseIf vEtapa = 3 Then
            vValorDescGradual = vValorDescGradualAV3
        End If
    Else
        If vEtapa = 1 Then
            vValorDescGradual = vValorDescGradualAP1
        ElseIf vEtapa = 2 Then
            vValorDescGradual = vValorDescGradualAP2
        ElseIf vEtapa = 3 Then
            vValorDescGradual = vValorDescGradualAP3
        End If
    End If

    If vLimitarDesc = 1 Then
        If vDesc > vValorDescGradual Then
        MsgBox "Desconto maior que o permitido pela empresa!", vbInformation, "Aviso do Sistema"
        txtDesc.Text = FormatNumber(vValorDescGradual, 2)
        End If
    End If
End If

'não dar desconto para vendas no cartão de débito
If vDescCartaoDebito = "SIM" And txtDesc.Text <> "0,00" Then
    If cboFormaPgto.Text = "3 - CARTÃO - DÉBITO" Or cboFormaPgtoEntrada.Text = "3 - CARTÃO - DÉBITO" Then
        MsgBox "Não é permitido dar desconto para vendas com pagamento em cartão de débito!" & Chr(13) & "Mude a forma de pagamento!", vbInformation, "Aviso do Sistema"
        txtDesc.Text = FormatNumber(0, 2)
    End If
End If

'não dar desconto para vendas no cartão de crédito
If vDescCartaoCredito = "SIM" And txtDesc.Text <> "0,00" Then
    If cboFormaPgto.Text = "4 - CARTÃO - CRÉDITO" Or cboFormaPgtoEntrada.Text = "4 - CARTÃO - CRÉDITO" Then
        MsgBox "Não é permitido dar desconto para vendas com pagamento em cartão de crédito!" & Chr(13) & "Mude a forma de pagamento!", vbInformation, "Aviso do Sistema"
        txtDesc.Text = FormatNumber(0, 2)
    End If
End If

Calcular_Desconto

txtDesc.Text = FormatNumber(txtDesc.Text, 2)

Exit Sub
   
'erro:
'   ShowMsg "O valor digitado é inválido!", vbExclamation
'   txtDesc.Text = 0
End Sub

Private Sub txtDescDinheiro_GotFocus()
SelectControl txtDescDinheiro
End Sub

Private Sub txtDescDinheiro_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
If KeyAscii = 13 Then
    txtDescDinheiro_LostFocus
End If
End Sub

Private Sub txtDescDinheiro_LostFocus()
Dim ValueTotal As Double
Dim ValueDiscount As Double
Dim Percent As Double

ValueTotal = txtSubtotal.Text

If txtDescDinheiro.Text = "" Then
    ValueDiscount = 0
Else
    ValueDiscount = txtDescDinheiro.Text
End If

Percent = (ValueDiscount / ValueTotal) * 100
txtDesc.Text = FormatNumber(Percent, 2)
txtDescDinheiro.Visible = False
txtDesc_LostFocus
If frmVendaFechamento.Visible = True Then txtRecebido.SetFocus
End Sub


Private Sub txtEntrada_Change()
'txtEntrada_Click
End Sub

Private Sub txtEntrada_Click()
'If txtTotalGeral.Text = "" Then
'   Exit Sub
'Else
'   Mostrar_ValorRestante
'   Calcular_Parcelas
 '  Calcular_Prazo
'End If
End Sub

Private Sub txtEntrada_GotFocus()
SelectControl txtEntrada
End Sub

Private Sub txtEntrada_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub

Private Sub txtEntrada_LostFocus()
'txtEntrada_Click
If txtEntrada = "" Then txtEntrada = Format(0, ocMONEY) Else txtEntrada = Format(txtEntrada, ocMONEY)
If txtTotalGeral.Text = "" Then
   Exit Sub
Else
   Mostrar_ValorRestante
   Calcular_Parcelas
   Calcular_Prazo
End If
End Sub

Private Sub txtQuant_Change()
If txtQuant.Text = "" Or txtQuant.Text = "0" Then Exit Sub
Calcular_Total
End Sub
Private Sub txtQuant_GotFocus()
SelectControl txtQuant
If Left(txtCodBarraPeso.Text, 1) = "2" Then SendKeys "{ENTER}"
End Sub
Private Sub txtQuant_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 194 Or KeyCode = 190 Then
   SendKeys "{BKSP}"
   SendKeys ","
End If
End Sub
Private Sub txtQuant_KeyPress(KeyAscii As Integer)
'Dim lNovoCod As Long
'saber se pode vender negagivo
Dim oCfg As ConfigItem
Dim bEstNeg As Boolean

'Recupera a configuração do estoque
Set oCfg = sysConfig("ESTOQUE_NEGATIVO")
bEstNeg = CBool(oCfg.Value)
Set oCfg = Nothing

 'variaveis de add produto
 Dim sSQL As String
 Dim itemVenda As Long
 
 'variaveis de verificar quant.
 Dim r As ADODB.Recordset
 Dim vCodProduto As Long
 
 'ADICIONAR O PRODUTO
 If txtCodProduto.Text = "" Then Exit Sub
 If txtValor.Text = "" Then Exit Sub

If txtCodProduto.Text = "" Then txtQuant.Locked = True Else txtQuant.Locked = False
KeyAscii = aNumeros(KeyAscii, True)

If KeyAscii = 13 Then
    'ADICIONAR O PRODUTO
    If txtQuant.Text = "" Then txtQuant.Text = 1
     
     'verificar quantidade
     vCodProduto = txtCodProduto.Text
     
    Dim vQtde As Double
   
   'Consulta os saldos
   sSQL = "SELECT quant_estoque, codigo FROM produtos WHERE (codigo = " & vCodProduto & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   If Not r.BOF Then vQtde = ValidateNull(r("quant_estoque"))
    'If r.State <> 0 Then r.Close
    'Set r = Nothing

    'Verifica se o produto foi vendido
    itemVenda = existeVenda(txtCodProduto.Text)
        
    If vQtde <= 0 Then
       
       If r("codigo") <> 1 Then
        If Not bEstNeg Then
           ShowMsg "A quantidade em estoque é insuficiente.", vbExclamation
           LimparObjetos_Produto
           txtCodBarra.SetFocus
           Exit Sub
        End If
      End If
       
        dbData.Execute "UPDATE pedidos_itens SET quantidade = " & Replace(txtQuant.Text, ",", ".") & ", preco = " & Replace(CCur(txtValor.Text), ",", ".") & ", subtotal = (" & Replace(CCur(txtValor.Text), ",", ".") & ") * (" & Replace(txtQuant.Text, ",", ".") & "), total = (" & Replace(CCur(txtValor.Text), ",", ".") & ") * (" & Replace(txtQuant.Text, ",", ".") & ") WHERE (codigo = " & txtCodItem.Text & ");"
        cmdAlterar.Enabled = False
    
    ElseIf vQtde > 0 Then
        
        If itemVenda <> -1 Then
            Dim varQuantGrid As Double
            Dim varQuantAdicionada As Double
            Dim varQuantSobrando As Double
            
            'verificar quantidade no grid
            For i = 1 To Grid.Rows - 1
                If Grid.TextMatrix(i, 2) = txtCodProduto.Text Then
                    varQuantGrid = Grid.TextMatrix(i, 5)
                    'Exit Function
                End If
            Next
        End If
   
    varQuantSobrando = vQtde - varQuantGrid
    varQuantAdicionada = txtQuant.Text
        
        If r("codigo") <> 1 Then
            If Not bEstNeg Then
                If varQuantAdicionada > vQtde Then
                    MsgBox "Quantidade não disponivel!", vbInformation, "Aviso do Sistema"
                    LimparObjetos_Produto
                    cmdAlterar.Enabled = False
                    txtCodBarra.Enabled = True
                    txtCodBarra.SetFocus
                    Exit Sub
                End If
            End If
        End If
        
        dbData.Execute "UPDATE pedidos_itens SET quantidade = " & Replace(txtQuant.Text, ",", ".") & ", preco = " & Replace(CCur(txtValor.Text), ",", ".") & ", subtotal = (" & Replace(CCur(txtValor.Text), ",", ".") & ") * (" & Replace(txtQuant.Text, ",", ".") & "), total = (" & Replace(CCur(txtValor.Text), ",", ".") & ") * (" & Replace(txtQuant.Text, ",", ".") & ") WHERE (codigo = " & txtCodItem.Text & ");"
        cmdAlterar.Enabled = False
        lblDesc.Caption = ""
      End If
      
      'MOSTRAR NA GRADE
      MostrarGrid_Produtos
      
        LimparObjetos_Produto
        txtQuant.BackColor = &H80000005
        cmdAlterar.Enabled = False
        txtCodBarra.Enabled = True
        cmdFinalizarAvista.Enabled = True
        cmdFinalizarPrazo.Enabled = True
        cmdOrçamento.Enabled = True
        cmdCancelarPedido.Enabled = True
        cmdRemover.Enabled = True
        cmdAvancado.Enabled = True
        cmdInfProduto.Enabled = True
        If txtCodProduto.Text = "" Then txtQuant.BackColor = &H80000005: txtQuant.Text = "0": txtCodBarra.Enabled = True: txtCodBarra.SetFocus
        If txtCodBarra.Enabled = True Then txtCodBarra.SetFocus
End If
   
PesoF4 = False

End Sub

Private Sub txtQuant_LostFocus()
'If txtQuant.BackColor = &HC0FFC0 Then txtQuant.SetFocus
End Sub

Private Sub txtRecebido_Change()
Calcular_Troco
End Sub

Private Sub txtRecebido_GotFocus()
   SelectControl txtRecebido
End Sub

Private Sub txtRecebido_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
If KeyAscii = 13 Then
      cmdFinalizar_Click
End If
End Sub

Private Sub txtRecebido_LostFocus()
If txtRecebido.Text = "" Then
    txtRecebido.Text = Format(0, ocMONEY)
    txtTroco.Text = Format(0, ocMONEY)
Else
    txtRecebido.Text = Format(txtRecebido.Text, ocMONEY)
End If
Calcular_Troco
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cmdSenha_Click
End Sub


Private Sub txtSubTotal_Validate(Cancel As Boolean)
Calcular_Desconto
End Sub


Private Sub txtTotal_Change()
If txtTotal.Text <> "0,00" Then Exit Sub

'If PesoF4 = True Then
'    txtQuant.SetFocus
'    SendKeys "{ENTER}"
'End If
End Sub

Private Sub txtTotalGeral_Change()
   txtRecebido_Change
End Sub





Private Sub txtValor_Change()
Calcular_Total
End Sub

Private Sub txtValor_GotFocus()
SelectControl txtValor
End Sub


Private Sub txtValor_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
If LTrim(txtValor) = "," Then txtValor.Text = "0,"
Calcular_Total
If KeyAscii = 13 Then
    txtQuant.SetFocus
    Beep
End If
End Sub


Private Sub txtValor_LostFocus()
txtValor.Text = Format(txtValor, ocMONEY)
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
   If txtValorParc = "" Then txtValorParc = Format(0, ocMONEY) Else txtValorParc = Format(txtValorParc, ocMONEY)
End Sub

Private Sub txtValorProdAvulso_GotFocus()
SelectControl txtValorProdAvulso
End Sub

Private Sub txtValorProdAvulso_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdOKProdAvulso_Click
End If
End Sub


Private Sub txtValorProdAvulso_LostFocus()
If txtValorProdAvulso.Text = "" Then txtValorProdAvulso.Text = Format(0, ocMONEY) Else txtValorProdAvulso.Text = Format(txtValorProdAvulso.Text, ocMONEY)
End Sub


Private Sub txtValorRest_Change()
   Calcular_Parcelas
End Sub

Private Sub txtValorRest_GotFocus()
SelectControl txtValorRest
End Sub


Private Sub txtValorRest_KeyPress(KeyAscii As Integer)
KeyAscii = aNumeros(KeyAscii, True)
End Sub


