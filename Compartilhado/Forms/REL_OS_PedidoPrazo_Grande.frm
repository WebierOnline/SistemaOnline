VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.Ocx"
Begin VB.Form REL_OS_PedidoPrazo_Grande 
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   13170
   LinkTopic       =   "Form1"
   ScaleHeight     =   192.881
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   232.304
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   4035
      Left            =   0
      Top             =   10905
      Width           =   13170
      _ExtentX        =   23230
      _ExtentY        =   7117
      Tipo            =   7
      Begin ReportX.ReportField frNome 
         Height          =   420
         Left            =   2700
         TabIndex        =   273
         Top             =   900
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   741
         Linhas          =   24
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField rfValorNum 
         Height          =   360
         Left            =   8820
         TabIndex        =   274
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
         Linhas          =   7
         Caption         =   "0,00"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         BackColor       =   14737632
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField2 
         Height          =   345
         Left            =   1920
         TabIndex        =   275
         Top             =   300
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   609
         Caption         =   "N O T A  P R O M I S S Ó R I A"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField10 
         Height          =   240
         Left            =   4140
         TabIndex        =   276
         Top             =   3555
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   423
         Linhas          =   10
         Caption         =   "ASSINATURA"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField12 
         Height          =   420
         Left            =   2280
         TabIndex        =   277
         Top             =   900
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   741
         Linhas          =   2
         Caption         =   "Eu, "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField14 
         Height          =   360
         Left            =   8400
         TabIndex        =   278
         Top             =   360
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   635
         Caption         =   "R$"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField17 
         Height          =   420
         Left            =   2280
         TabIndex        =   279
         Top             =   1320
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   741
         Linhas          =   2
         Caption         =   "pagarei por essa única via de Nota Promissória ou a sua ordem a quantia de"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField frValorEst 
         Height          =   420
         Left            =   2280
         TabIndex        =   280
         Top             =   1740
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   741
         Linhas          =   24
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField19 
         Height          =   420
         Left            =   2280
         TabIndex        =   281
         Top             =   2160
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         Linhas          =   2
         Caption         =   "em moeda corrente deste País."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField rfData 
         Height          =   420
         Left            =   4920
         TabIndex        =   282
         Top             =   2580
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   741
         Linhas          =   2
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin VB.Shape Shape7 
         BorderWidth     =   2
         Height          =   3795
         Left            =   420
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   10635
      End
      Begin VB.Line Line9 
         X1              =   9300
         X2              =   3840
         Y1              =   3540
         Y2              =   3540
      End
      Begin VB.Image Image1 
         Height          =   3510
         Left            =   660
         Picture         =   "REL_OS_PedidoPrazo_Grande.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1275
      End
   End
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   11880
      TabIndex        =   44
      Top             =   6420
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Titulo          =   ""
      LarguraPapel    =   209
      AlturaPapel     =   146
      NomeImpressora  =   "IMPRESSORA1"
      Registrado      =   0   'False
      Visualizar      =   0   'False
   End
   Begin ReportX.ReportSection ReportSection2 
      Align           =   1  'Align Top
      Height          =   10905
      Left            =   0
      Top             =   0
      Width           =   13170
      _ExtentX        =   23230
      _ExtentY        =   19235
      Ordem           =   1
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   17
         Left            =   1440
         TabIndex        =   234
         Top             =   6780
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   18
         Left            =   1440
         TabIndex        =   235
         Top             =   7005
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   19
         Left            =   1440
         TabIndex        =   236
         Top             =   7230
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   20
         Left            =   1440
         TabIndex        =   237
         Top             =   7455
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   21
         Left            =   1440
         TabIndex        =   238
         Top             =   7680
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   22
         Left            =   1440
         TabIndex        =   239
         Top             =   7905
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   23
         Left            =   1440
         TabIndex        =   240
         Top             =   8130
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   24
         Left            =   1440
         TabIndex        =   241
         Top             =   8355
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   25
         Left            =   1440
         TabIndex        =   242
         Top             =   8580
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   26
         Left            =   1440
         TabIndex        =   243
         Top             =   8805
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   0
         Left            =   10440
         TabIndex        =   120
         Top             =   2940
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   0
         Left            =   9700
         TabIndex        =   103
         Top             =   2940
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   78
         Top             =   2940
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   9
         Left            =   8620
         TabIndex        =   43
         Top             =   4965
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   9
         Left            =   6840
         TabIndex        =   42
         Top             =   4965
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   9
         Left            =   7860
         TabIndex        =   41
         Top             =   4965
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   9
         Left            =   1440
         TabIndex        =   40
         Top             =   4965
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   8
         Left            =   8620
         TabIndex        =   39
         Top             =   4740
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   8
         Left            =   6840
         TabIndex        =   38
         Top             =   4740
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   8
         Left            =   7860
         TabIndex        =   37
         Top             =   4740
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   8
         Left            =   1440
         TabIndex        =   36
         Top             =   4740
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   7
         Left            =   8620
         TabIndex        =   35
         Top             =   4515
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   7
         Left            =   6840
         TabIndex        =   34
         Top             =   4515
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   7
         Left            =   7860
         TabIndex        =   33
         Top             =   4515
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   7
         Left            =   1440
         TabIndex        =   32
         Top             =   4515
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   6
         Left            =   8620
         TabIndex        =   31
         Top             =   4290
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   6
         Left            =   6840
         TabIndex        =   30
         Top             =   4290
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   6
         Left            =   7860
         TabIndex        =   29
         Top             =   4290
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   6
         Left            =   1440
         TabIndex        =   28
         Top             =   4290
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   5
         Left            =   8620
         TabIndex        =   27
         Top             =   4065
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   5
         Left            =   6840
         TabIndex        =   26
         Top             =   4065
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   5
         Left            =   7860
         TabIndex        =   25
         Top             =   4065
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   5
         Left            =   1440
         TabIndex        =   24
         Top             =   4065
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   4
         Left            =   8620
         TabIndex        =   23
         Top             =   3840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   4
         Left            =   6840
         TabIndex        =   22
         Top             =   3840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   4
         Left            =   7860
         TabIndex        =   21
         Top             =   3840
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   4
         Left            =   1440
         TabIndex        =   20
         Top             =   3840
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   3
         Left            =   8620
         TabIndex        =   19
         Top             =   3615
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   3
         Left            =   6840
         TabIndex        =   18
         Top             =   3615
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   3
         Left            =   7860
         TabIndex        =   17
         Top             =   3615
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   3
         Left            =   1440
         TabIndex        =   16
         Top             =   3615
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   2
         Left            =   8620
         TabIndex        =   15
         Top             =   3390
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   2
         Left            =   6840
         TabIndex        =   14
         Top             =   3390
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   2
         Left            =   7860
         TabIndex        =   13
         Top             =   3390
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   2
         Left            =   1440
         TabIndex        =   12
         Top             =   3390
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   1
         Left            =   8620
         TabIndex        =   11
         Top             =   3165
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   1
         Left            =   6840
         TabIndex        =   10
         Top             =   3165
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   1
         Left            =   7860
         TabIndex        =   9
         Top             =   3165
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   1
         Left            =   1440
         TabIndex        =   8
         Top             =   3165
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   0
         Left            =   8620
         TabIndex        =   7
         Top             =   2940
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   0
         Left            =   6840
         TabIndex        =   6
         Top             =   2940
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   0
         Left            =   7860
         TabIndex        =   5
         Top             =   2940
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   0
         Left            =   1440
         TabIndex        =   4
         Top             =   2940
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField68 
         Height          =   300
         Left            =   1440
         TabIndex        =   0
         Top             =   2580
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   529
         Caption         =   "DISCRIMINAÇÃO"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField39 
         Height          =   300
         Left            =   7860
         TabIndex        =   1
         Top             =   2580
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   529
         Caption         =   "QTDA"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField40 
         Height          =   300
         Left            =   6840
         TabIndex        =   2
         Top             =   2580
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         Caption         =   "VALOR"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField47 
         Height          =   300
         Left            =   8620
         TabIndex        =   3
         Top             =   2580
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         Caption         =   "SUBTOTAL"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField rf2 
         Height          =   270
         Left            =   2880
         TabIndex        =   45
         Top             =   420
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   476
         Caption         =   ""
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField rf1 
         Height          =   390
         Left            =   2880
         TabIndex        =   46
         Top             =   60
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   688
         Caption         =   ""
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField rf3 
         Height          =   270
         Left            =   2880
         TabIndex        =   47
         Top             =   630
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   476
         Caption         =   ""
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField rf4 
         Height          =   275
         Left            =   2880
         TabIndex        =   48
         Top             =   840
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   476
         Caption         =   ""
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   10
         Left            =   1440
         TabIndex        =   49
         Top             =   5190
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   11
         Left            =   1440
         TabIndex        =   50
         Top             =   5415
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   12
         Left            =   1440
         TabIndex        =   51
         Top             =   5640
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   13
         Left            =   1440
         TabIndex        =   52
         Top             =   5865
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   10
         Left            =   6840
         TabIndex        =   53
         Top             =   5190
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   11
         Left            =   6840
         TabIndex        =   54
         Top             =   5415
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   12
         Left            =   6840
         TabIndex        =   55
         Top             =   5640
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   13
         Left            =   6840
         TabIndex        =   56
         Top             =   5865
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   10
         Left            =   7860
         TabIndex        =   57
         Top             =   5190
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   11
         Left            =   7860
         TabIndex        =   58
         Top             =   5415
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   12
         Left            =   7860
         TabIndex        =   59
         Top             =   5640
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   13
         Left            =   7860
         TabIndex        =   60
         Top             =   5865
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   10
         Left            =   8620
         TabIndex        =   61
         Top             =   5190
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   11
         Left            =   8620
         TabIndex        =   62
         Top             =   5415
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   12
         Left            =   8620
         TabIndex        =   63
         Top             =   5640
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   13
         Left            =   8620
         TabIndex        =   64
         Top             =   5865
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   14
         Left            =   1440
         TabIndex        =   65
         Top             =   6090
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   14
         Left            =   6840
         TabIndex        =   66
         Top             =   6090
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   14
         Left            =   7860
         TabIndex        =   67
         Top             =   6090
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   14
         Left            =   8620
         TabIndex        =   68
         Top             =   6090
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   15
         Left            =   1440
         TabIndex        =   69
         Top             =   6315
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   15
         Left            =   6840
         TabIndex        =   70
         Top             =   6315
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   15
         Left            =   7860
         TabIndex        =   71
         Top             =   6315
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   15
         Left            =   8620
         TabIndex        =   72
         Top             =   6315
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   16
         Left            =   1440
         TabIndex        =   73
         Top             =   6540
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   16
         Left            =   6840
         TabIndex        =   74
         Top             =   6540
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   16
         Left            =   7860
         TabIndex        =   75
         Top             =   6540
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   16
         Left            =   8620
         TabIndex        =   76
         Top             =   6540
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField3 
         Height          =   300
         Left            =   180
         TabIndex        =   77
         Top             =   2580
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "TIPO"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   79
         Top             =   3165
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   2
         Left            =   180
         TabIndex        =   80
         Top             =   3390
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   3
         Left            =   180
         TabIndex        =   81
         Top             =   3615
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   4
         Left            =   180
         TabIndex        =   82
         Top             =   3840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   5
         Left            =   180
         TabIndex        =   83
         Top             =   4065
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   6
         Left            =   180
         TabIndex        =   84
         Top             =   4290
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   7
         Left            =   180
         TabIndex        =   85
         Top             =   4515
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   8
         Left            =   180
         TabIndex        =   86
         Top             =   4740
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   9
         Left            =   180
         TabIndex        =   87
         Top             =   4965
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   10
         Left            =   180
         TabIndex        =   88
         Top             =   5190
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   11
         Left            =   180
         TabIndex        =   89
         Top             =   5415
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   12
         Left            =   180
         TabIndex        =   90
         Top             =   5640
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   13
         Left            =   180
         TabIndex        =   91
         Top             =   5865
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   14
         Left            =   180
         TabIndex        =   92
         Top             =   6090
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   15
         Left            =   180
         TabIndex        =   93
         Top             =   6315
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   16
         Left            =   180
         TabIndex        =   94
         Top             =   6540
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField lblSubTotal 
         Height          =   270
         Left            =   9450
         TabIndex        =   95
         Top             =   10050
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   476
         Caption         =   "Subtotal:"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtSubtotal 
         Height          =   270
         Left            =   10320
         TabIndex        =   96
         Top             =   10050
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   476
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         BackColor       =   12632256
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField lblDesconto 
         Height          =   270
         Left            =   9660
         TabIndex        =   97
         Top             =   10320
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   476
         Caption         =   "Desc.:"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesconto 
         Height          =   270
         Left            =   10320
         TabIndex        =   98
         Top             =   10320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   476
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField ReportField13 
         Height          =   270
         Left            =   9660
         TabIndex        =   99
         Top             =   10590
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   476
         Caption         =   "Total:"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotal 
         Height          =   270
         Left            =   10320
         TabIndex        =   100
         Top             =   10590
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   476
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         BackColor       =   12632256
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField ReportField1 
         Height          =   300
         Left            =   9700
         TabIndex        =   101
         Top             =   2580
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   529
         Caption         =   "DESC."
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField8 
         Height          =   300
         Left            =   10440
         TabIndex        =   102
         Top             =   2580
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         Caption         =   "TOTAL"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   1
         Left            =   9700
         TabIndex        =   104
         Top             =   3165
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   2
         Left            =   9700
         TabIndex        =   105
         Top             =   3390
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   3
         Left            =   9700
         TabIndex        =   106
         Top             =   3615
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   4
         Left            =   9700
         TabIndex        =   107
         Top             =   3840
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   5
         Left            =   9700
         TabIndex        =   108
         Top             =   4065
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   6
         Left            =   9700
         TabIndex        =   109
         Top             =   4290
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   7
         Left            =   9700
         TabIndex        =   110
         Top             =   4515
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   8
         Left            =   9700
         TabIndex        =   111
         Top             =   4740
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   9
         Left            =   9700
         TabIndex        =   112
         Top             =   4965
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   10
         Left            =   9700
         TabIndex        =   113
         Top             =   5190
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   11
         Left            =   9700
         TabIndex        =   114
         Top             =   5415
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   12
         Left            =   9700
         TabIndex        =   115
         Top             =   5640
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   13
         Left            =   9700
         TabIndex        =   116
         Top             =   5865
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   14
         Left            =   9700
         TabIndex        =   117
         Top             =   6090
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   15
         Left            =   9700
         TabIndex        =   118
         Top             =   6315
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   16
         Left            =   9700
         TabIndex        =   119
         Top             =   6540
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   1
         Left            =   10440
         TabIndex        =   121
         Top             =   3165
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   2
         Left            =   10440
         TabIndex        =   122
         Top             =   3390
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   3
         Left            =   10440
         TabIndex        =   123
         Top             =   3615
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   4
         Left            =   10440
         TabIndex        =   124
         Top             =   3840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   5
         Left            =   10440
         TabIndex        =   125
         Top             =   4065
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   6
         Left            =   10440
         TabIndex        =   126
         Top             =   4290
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   7
         Left            =   10440
         TabIndex        =   127
         Top             =   4515
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   8
         Left            =   10440
         TabIndex        =   128
         Top             =   4740
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   9
         Left            =   10440
         TabIndex        =   129
         Top             =   4965
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   10
         Left            =   10440
         TabIndex        =   130
         Top             =   5190
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   11
         Left            =   10440
         TabIndex        =   131
         Top             =   5415
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   12
         Left            =   10440
         TabIndex        =   132
         Top             =   5640
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   13
         Left            =   10440
         TabIndex        =   133
         Top             =   5865
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   14
         Left            =   10440
         TabIndex        =   134
         Top             =   6090
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   15
         Left            =   10440
         TabIndex        =   135
         Top             =   6315
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   16
         Left            =   10440
         TabIndex        =   136
         Top             =   6540
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField ReportField16 
         Height          =   270
         Left            =   180
         TabIndex        =   137
         Top             =   1965
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   476
         Caption         =   "CPF/CNPJ:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtCPF 
         Height          =   270
         Left            =   1080
         TabIndex        =   138
         Top             =   1965
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   476
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField ReportField41 
         Height          =   270
         Left            =   180
         TabIndex        =   139
         Top             =   1200
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   476
         Caption         =   "Cliente:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtCliente 
         Height          =   270
         Left            =   870
         TabIndex        =   140
         Top             =   1200
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   476
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField ReportField4 
         Height          =   270
         Left            =   195
         TabIndex        =   141
         Top             =   1455
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   476
         Caption         =   "End.:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtEnd 
         Height          =   270
         Left            =   660
         TabIndex        =   142
         Top             =   1455
         Width           =   6180
         _ExtentX        =   10901
         _ExtentY        =   476
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField ReportField5 
         Height          =   270
         Left            =   180
         TabIndex        =   143
         Top             =   2220
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   476
         Caption         =   "Ponto de Ref.:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtRef2 
         Height          =   270
         Left            =   1380
         TabIndex        =   144
         Top             =   2220
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   476
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField ReportField7 
         Height          =   270
         Left            =   2820
         TabIndex        =   145
         Top             =   1965
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   476
         Caption         =   "RG:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtRG 
         Height          =   270
         Left            =   3120
         TabIndex        =   146
         Top             =   1965
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   476
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField ReportField9 
         Height          =   270
         Left            =   180
         TabIndex        =   147
         Top             =   1710
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   476
         Caption         =   "Cidade:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtCidade 
         Height          =   270
         Left            =   840
         TabIndex        =   148
         Top             =   1710
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   476
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   17
         Left            =   180
         TabIndex        =   150
         Top             =   6780
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   18
         Left            =   180
         TabIndex        =   151
         Top             =   7005
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   19
         Left            =   180
         TabIndex        =   152
         Top             =   7230
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   20
         Left            =   180
         TabIndex        =   153
         Top             =   7455
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   21
         Left            =   180
         TabIndex        =   154
         Top             =   7680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   22
         Left            =   180
         TabIndex        =   155
         Top             =   7905
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   23
         Left            =   180
         TabIndex        =   156
         Top             =   8130
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   24
         Left            =   180
         TabIndex        =   157
         Top             =   8355
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   25
         Left            =   180
         TabIndex        =   158
         Top             =   8580
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   26
         Left            =   180
         TabIndex        =   159
         Top             =   8805
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   27
         Left            =   180
         TabIndex        =   160
         Top             =   9030
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   28
         Left            =   180
         TabIndex        =   161
         Top             =   9255
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   29
         Left            =   180
         TabIndex        =   162
         Top             =   9480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   17
         Left            =   6840
         TabIndex        =   163
         Top             =   6780
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   18
         Left            =   6840
         TabIndex        =   164
         Top             =   7005
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   19
         Left            =   6840
         TabIndex        =   165
         Top             =   7230
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   20
         Left            =   6840
         TabIndex        =   166
         Top             =   7455
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   21
         Left            =   6840
         TabIndex        =   167
         Top             =   7680
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   22
         Left            =   6840
         TabIndex        =   168
         Top             =   7905
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   23
         Left            =   6840
         TabIndex        =   169
         Top             =   8130
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   24
         Left            =   6840
         TabIndex        =   170
         Top             =   8355
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   25
         Left            =   6840
         TabIndex        =   171
         Top             =   8580
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   26
         Left            =   6840
         TabIndex        =   172
         Top             =   8805
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   27
         Left            =   6840
         TabIndex        =   173
         Top             =   9030
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   28
         Left            =   6840
         TabIndex        =   174
         Top             =   9255
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   29
         Left            =   6840
         TabIndex        =   175
         Top             =   9480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   17
         Left            =   7860
         TabIndex        =   176
         Top             =   6780
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   18
         Left            =   7860
         TabIndex        =   177
         Top             =   7005
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   19
         Left            =   7860
         TabIndex        =   178
         Top             =   7230
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   20
         Left            =   7860
         TabIndex        =   179
         Top             =   7455
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   21
         Left            =   7860
         TabIndex        =   180
         Top             =   7680
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   22
         Left            =   7860
         TabIndex        =   181
         Top             =   7905
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   23
         Left            =   7860
         TabIndex        =   182
         Top             =   8130
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   24
         Left            =   7860
         TabIndex        =   183
         Top             =   8355
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   25
         Left            =   7860
         TabIndex        =   184
         Top             =   8580
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   26
         Left            =   7860
         TabIndex        =   185
         Top             =   8805
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   27
         Left            =   7860
         TabIndex        =   186
         Top             =   9030
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   28
         Left            =   7860
         TabIndex        =   187
         Top             =   9255
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   29
         Left            =   7860
         TabIndex        =   188
         Top             =   9480
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   17
         Left            =   8640
         TabIndex        =   189
         Top             =   6780
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   18
         Left            =   8640
         TabIndex        =   190
         Top             =   7005
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   19
         Left            =   8640
         TabIndex        =   191
         Top             =   7230
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   20
         Left            =   8640
         TabIndex        =   192
         Top             =   7455
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   21
         Left            =   8640
         TabIndex        =   193
         Top             =   7680
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   22
         Left            =   8640
         TabIndex        =   194
         Top             =   7905
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   23
         Left            =   8640
         TabIndex        =   195
         Top             =   8130
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   24
         Left            =   8640
         TabIndex        =   196
         Top             =   8355
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   25
         Left            =   8640
         TabIndex        =   197
         Top             =   8580
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   26
         Left            =   8640
         TabIndex        =   198
         Top             =   8805
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   27
         Left            =   8640
         TabIndex        =   199
         Top             =   9030
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   28
         Left            =   8640
         TabIndex        =   200
         Top             =   9255
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   29
         Left            =   8640
         TabIndex        =   201
         Top             =   9480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   17
         Left            =   9700
         TabIndex        =   202
         Top             =   6780
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   18
         Left            =   9700
         TabIndex        =   203
         Top             =   7005
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   19
         Left            =   9700
         TabIndex        =   204
         Top             =   7230
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   20
         Left            =   9700
         TabIndex        =   205
         Top             =   7455
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   21
         Left            =   9700
         TabIndex        =   206
         Top             =   7680
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   22
         Left            =   9700
         TabIndex        =   207
         Top             =   7905
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   23
         Left            =   9700
         TabIndex        =   208
         Top             =   8130
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   24
         Left            =   9700
         TabIndex        =   209
         Top             =   8355
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   25
         Left            =   9700
         TabIndex        =   210
         Top             =   8580
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   26
         Left            =   9700
         TabIndex        =   211
         Top             =   8805
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   27
         Left            =   9700
         TabIndex        =   212
         Top             =   9030
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   28
         Left            =   9700
         TabIndex        =   213
         Top             =   9255
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   29
         Left            =   9700
         TabIndex        =   214
         Top             =   9480
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   17
         Left            =   10440
         TabIndex        =   215
         Top             =   6780
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   18
         Left            =   10440
         TabIndex        =   216
         Top             =   7005
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   19
         Left            =   10440
         TabIndex        =   217
         Top             =   7230
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   20
         Left            =   10440
         TabIndex        =   218
         Top             =   7455
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   21
         Left            =   10440
         TabIndex        =   219
         Top             =   7680
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   22
         Left            =   10440
         TabIndex        =   220
         Top             =   7905
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   23
         Left            =   10440
         TabIndex        =   221
         Top             =   8130
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   24
         Left            =   10440
         TabIndex        =   222
         Top             =   8355
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   25
         Left            =   10440
         TabIndex        =   223
         Top             =   8580
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   26
         Left            =   10440
         TabIndex        =   224
         Top             =   8805
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   27
         Left            =   10440
         TabIndex        =   225
         Top             =   9030
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   28
         Left            =   10440
         TabIndex        =   226
         Top             =   9255
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   29
         Left            =   10440
         TabIndex        =   227
         Top             =   9480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTipo 
         Height          =   225
         Index           =   30
         Left            =   180
         TabIndex        =   228
         Top             =   9720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtUnit 
         Height          =   225
         Index           =   30
         Left            =   6840
         TabIndex        =   229
         Top             =   9720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   30
         Left            =   7860
         TabIndex        =   230
         Top             =   9720
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   30
         Left            =   8640
         TabIndex        =   231
         Top             =   9720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesco 
         Height          =   225
         Index           =   30
         Left            =   9705
         TabIndex        =   232
         Top             =   9720
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   30
         Left            =   10440
         TabIndex        =   233
         Top             =   9720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   27
         Left            =   1440
         TabIndex        =   244
         Top             =   9030
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   28
         Left            =   1440
         TabIndex        =   245
         Top             =   9255
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   29
         Left            =   1440
         TabIndex        =   246
         Top             =   9480
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   30
         Left            =   1440
         TabIndex        =   247
         Top             =   9705
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   397
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField frTitParc 
         Height          =   255
         Left            =   7020
         TabIndex        =   248
         Top             =   1200
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   450
         Caption         =   "PARCELAS"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField frValorParc 
         Height          =   195
         Index           =   0
         Left            =   8520
         TabIndex        =   249
         Top             =   1500
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   344
         Caption         =   ""
         Alignment       =   1
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
      Begin ReportX.ReportField frVencParc 
         Height          =   195
         Index           =   0
         Left            =   7680
         TabIndex        =   250
         Top             =   1500
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   344
         Caption         =   ""
         Alignment       =   2
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
      Begin ReportX.ReportField frNumParc 
         Height          =   195
         Index           =   0
         Left            =   7260
         TabIndex        =   251
         Top             =   1500
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   344
         Caption         =   ""
         Alignment       =   2
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
      Begin ReportX.ReportField txtVendedor 
         Height          =   240
         Left            =   10200
         TabIndex        =   252
         Top             =   1800
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   423
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtNumero 
         Height          =   285
         Left            =   9420
         TabIndex        =   253
         Top             =   1500
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   503
         Caption         =   "000000"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BordaLagura     =   2
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtData 
         Height          =   240
         Left            =   9420
         TabIndex        =   254
         Top             =   1800
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   423
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BordaLagura     =   2
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtVenda 
         Height          =   300
         Left            =   9420
         TabIndex        =   255
         Top             =   2055
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   ""
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BordaLagura     =   2
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField frNumParc 
         Height          =   195
         Index           =   1
         Left            =   7260
         TabIndex        =   256
         Top             =   1680
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   344
         Caption         =   ""
         Alignment       =   2
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
      Begin ReportX.ReportField frNumParc 
         Height          =   195
         Index           =   2
         Left            =   7260
         TabIndex        =   257
         Top             =   1860
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   344
         Caption         =   ""
         Alignment       =   2
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
      Begin ReportX.ReportField frNumParc 
         Height          =   195
         Index           =   3
         Left            =   7260
         TabIndex        =   258
         Top             =   2040
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   344
         Caption         =   ""
         Alignment       =   2
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
      Begin ReportX.ReportField frVencParc 
         Height          =   195
         Index           =   1
         Left            =   7680
         TabIndex        =   259
         Top             =   1680
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   344
         Caption         =   ""
         Alignment       =   2
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
      Begin ReportX.ReportField frVencParc 
         Height          =   195
         Index           =   2
         Left            =   7680
         TabIndex        =   260
         Top             =   1860
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   344
         Caption         =   ""
         Alignment       =   2
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
      Begin ReportX.ReportField frVencParc 
         Height          =   195
         Index           =   3
         Left            =   7680
         TabIndex        =   261
         Top             =   2040
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   344
         Caption         =   ""
         Alignment       =   2
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
      Begin ReportX.ReportField frValorParc 
         Height          =   195
         Index           =   1
         Left            =   8520
         TabIndex        =   262
         Top             =   1680
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   344
         Caption         =   ""
         Alignment       =   1
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
      Begin ReportX.ReportField frValorParc 
         Height          =   195
         Index           =   2
         Left            =   8520
         TabIndex        =   263
         Top             =   1860
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   344
         Caption         =   ""
         Alignment       =   1
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
      Begin ReportX.ReportField frValorParc 
         Height          =   195
         Index           =   3
         Left            =   8520
         TabIndex        =   264
         Top             =   2040
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   344
         Caption         =   ""
         Alignment       =   1
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
      Begin ReportX.ReportField frNumParc 
         Height          =   195
         Index           =   4
         Left            =   7260
         TabIndex        =   265
         Top             =   2220
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   344
         Caption         =   ""
         Alignment       =   2
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
      Begin ReportX.ReportField frVencParc 
         Height          =   195
         Index           =   4
         Left            =   7680
         TabIndex        =   266
         Top             =   2220
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   344
         Caption         =   ""
         Alignment       =   2
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
      Begin ReportX.ReportField frValorParc 
         Height          =   195
         Index           =   4
         Left            =   8520
         TabIndex        =   267
         Top             =   2220
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   344
         Caption         =   ""
         Alignment       =   1
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
      Begin ReportX.ReportField ReportField6 
         Height          =   255
         Left            =   9420
         TabIndex        =   268
         Top             =   1200
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   450
         Caption         =   "ORDEM"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField frEnt 
         Height          =   195
         Left            =   6960
         TabIndex        =   283
         Top             =   1500
         Visible         =   0   'False
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   344
         Caption         =   "PG"
         Alignment       =   1
         Mostrar         =   0   'False
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
      Begin ReportX.ReportField ReportField11 
         Height          =   270
         Left            =   180
         TabIndex        =   284
         Top             =   10050
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   476
         Caption         =   "Serviços:"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtQuantServicos 
         Height          =   270
         Left            =   900
         TabIndex        =   285
         Top             =   10050
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   476
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField ReportField15 
         Height          =   255
         Left            =   180
         TabIndex        =   286
         Top             =   10320
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   450
         Caption         =   "Peças:"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtQuantPecas 
         Height          =   270
         Left            =   900
         TabIndex        =   287
         Top             =   10320
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   476
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField ReportField21 
         Height          =   270
         Left            =   180
         TabIndex        =   288
         Top             =   10590
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   476
         Caption         =   "Total:"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtQuantGeral 
         Height          =   270
         Left            =   900
         TabIndex        =   289
         Top             =   10590
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   476
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         BackColor       =   12632256
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalServicos 
         Height          =   270
         Left            =   1380
         TabIndex        =   290
         Top             =   10050
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   476
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalPecas 
         Height          =   270
         Left            =   1380
         TabIndex        =   291
         Top             =   10320
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   476
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtTotalPecasServicos 
         Height          =   270
         Left            =   1380
         TabIndex        =   292
         Top             =   10590
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   476
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         BackColor       =   12632256
         AlturaLivre     =   -1  'True
      End
      Begin VB.Shape Shape8 
         BorderWidth     =   2
         Height          =   870
         Left            =   120
         Top             =   10020
         Width           =   2715
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*Troca de produtos até o prazo maximo de 24 horas."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   4080
         TabIndex        =   272
         Top             =   10200
         Width           =   2985
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*Não aceitamos devolução de produtos com a embalagem violada."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   4080
         TabIndex        =   271
         Top             =   10080
         Width           =   3750
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*Clientes com mais de 30 dias de vencidos sujeito a inclusão do nome no SPC e Serasa "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   4080
         TabIndex        =   270
         Top             =   10320
         Width           =   4920
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*Após o vencimento cobrar 0,15% de juros ao dia."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   4080
         TabIndex        =   269
         Top             =   10440
         Width           =   2910
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         Height          =   1395
         Left            =   6960
         Top             =   1140
         Width           =   2355
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         Height          =   1395
         Left            =   9300
         Top             =   1140
         Width           =   2175
      End
      Begin VB.Line Line3 
         X1              =   7860
         X2              =   7860
         Y1              =   10020
         Y2              =   2520
      End
      Begin VB.Line Line6 
         X1              =   6840
         X2              =   6840
         Y1              =   10020
         Y2              =   2520
      End
      Begin VB.Line Line8 
         X1              =   5280
         X2              =   9300
         Y1              =   10680
         Y2              =   10680
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CLIENTE"
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
         Height          =   195
         Left            =   8400
         TabIndex        =   149
         Top             =   10680
         Width           =   825
      End
      Begin VB.Line Line7 
         X1              =   10395
         X2              =   10395
         Y1              =   10020
         Y2              =   2520
      End
      Begin VB.Shape Shape6 
         BorderWidth     =   2
         Height          =   870
         Left            =   9360
         Top             =   10020
         Width           =   2115
      End
      Begin VB.Line Line2 
         X1              =   1395
         X2              =   1395
         Y1              =   10020
         Y2              =   2520
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   11400
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line5 
         X1              =   9660
         X2              =   9660
         Y1              =   10020
         Y2              =   2520
      End
      Begin VB.Line Line4 
         X1              =   8580
         X2              =   8580
         Y1              =   10020
         Y2              =   2520
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         Height          =   7515
         Left            =   120
         Top             =   2520
         Width           =   11355
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   1395
         Left            =   120
         Top             =   1140
         Width           =   6855
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   1095
         Left            =   120
         Top             =   60
         Width           =   11355
      End
      Begin VB.Image imgLogo 
         Height          =   1035
         Left            =   180
         Stretch         =   -1  'True
         Top             =   60
         Width           =   2595
      End
   End
End
Attribute VB_Name = "REL_OS_PedidoPrazo_Grande"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String
Dim r As ADODB.Recordset
Dim rPd As ADODB.Recordset
Dim rOS As ADODB.Recordset
Dim rCl As ADODB.Recordset
Dim rIt As ADODB.Recordset
Dim rTotais As ADODB.Recordset
Dim rPc As ADODB.Recordset
Dim rFu As ADODB.Recordset
Dim rEquip As ADODB.Recordset
'Dim vTipoOS As String
Public cCfg As ConfigItem       'arquivo .ini
Public oIni As Ini              'arquivo .ini
Dim var_ImpNormal As String
Dim i As Integer
Dim Cont As Long
Dim wValorFormatado As String
Public Sub loadPedidos(ByVal Pedido As Long, ByVal Tipo As String)
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

   'Dim Cont As Long
   'Dim wValorFormatado As String
   If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
        Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, PLACA, ANO, KM, COR FROM OS_Equipamento_Auto WHERE (cod_os = " & vCodOS & ");")
   ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
        Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, EQUIPAMENTO FROM OS_Equipamento WHERE (cod_os = " & vCodOS & ");")
   ElseIf vTipoOS = "Comunicação Visual" Then
        Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, EQUIPAMENTO FROM OS_Equipamento WHERE (cod_os = " & vCodOS & ");")
   End If
   
   Set rOS = dbData.OpenRecordset("SELECT cod_pedido, cod_cliente, cod_funcionario, SUBTOTAL, TOTAL, ValorDescReal  FROM os WHERE (cod_os = " & vCodOS & ");")
   Set rPd = dbData.OpenRecordset("SELECT cod_cliente, cod_funcionario, data_compra, total, valor_desc, tipo_desc, subtotal, tipo_pagamento FROM pedidos WHERE (cod_pedido = " & Pedido & ");")
   Set rCl = dbData.OpenRecordset("SELECT nome, ENDERECO, numero, bairro, ponto_de_referencia, Cidade, estado, TELEFONE1, CPF, rg FROM cliente WHERE (codigo = " & rOS("cod_cliente") & ");")
   
   'verificar se existe produtos
   sSQL = "SELECT cod_pedido FROM pedidos_itens WHERE (pedidos_itens.cod_pedido = " & Pedido & ")"
   Set rIt = dbData.OpenRecordset(sSQL)


   sSQL = ""
   
   If rIt.EOF Then      'somente existir serviços
        'SERVIÇOS
        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
              sSQL = sSQL & "SELECT sum(quantidade) as vSomaQuantServ, sum(total) as vSomaValorServ " & _
              "FROM OS_Servicos_Auto WHERE (cod_os = " & vCodOS & ")"
        ElseIf vTipoOS = "Recapadora" Then
            sSQL = sSQL & "SELECT sum(quantidade) as vSomaQuantServ, sum(total) as vSomaValorServ " & _
            "FROM OS_servicos_recapadora WHERE (cod_os = " & vCodOS & ")"
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
              sSQL = sSQL & "SELECT sum(quantidade) as vSomaQuantServ, sum(total) as vSomaValorServ " & _
              "FROM OS_Servicos_Auto WHERE (cod_os = " & vCodOS & ")"
        ElseIf vTipoOS = "Comunicação Visual" Then
              sSQL = sSQL & "SELECT sum(quantidade) as vSomaQuantServ, sum(total) as vSomaValorServ " & _
              "FROM OS_Servicos_Auto WHERE (cod_os = " & vCodOS & ")"
        End If
        
        Set rTotais = dbData.OpenRecordset(sSQL)
        'Debug.Print sSQL
        
        If Not rTotais.EOF Then
            txtQuantServicos.Caption = Format(rTotais("vSomaQuantServ"), "000")
            txtTotalServicos.Caption = FormatNumber(rTotais("vSomaValorServ"), 2)
        Else
            txtQuantServicos.Caption = Format(0, "000")
            txtTotalServicos.Caption = FormatNumber(0, 2)
        End If
        
        txtQuantPecas.Caption = Format(0, "000")
        txtTotalPecas.Caption = FormatNumber(0, 2)
        
        txtQuantGeral.Caption = Format(CInt(txtQuantPecas.Caption) + CInt(txtQuantServicos.Caption), "000")
        txtTotalPecasServicos.Caption = FormatNumber(CCur(txtTotalPecas.Caption) + CCur(txtTotalServicos.Caption), 2)

   
   Else             'produtos e serviços
        'PRODUTOS
        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
             sSQL = "SELECT sum(quantidade) as vSomaQuantProd, sum(total) as vSomaValorProd FROM pedidos_itens  WHERE (pedidos_itens.cod_pedido = " & Pedido & ")"
        ElseIf vTipoOS = "Recapadora" Then
             sSQL = "SELECT sum(quantidade) as vSomaQuantProd, sum(total) as vSomaValorProd FROM pedidos_itens  WHERE (pedidos_itens.cod_pedido = " & Pedido & ")"
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
             sSQL = "SELECT sum(quantidade) as vSomaQuantProd, sum(total) as vSomaValorProd FROM pedidos_itens  WHERE (pedidos_itens.cod_pedido = " & Pedido & ")"
        ElseIf vTipoOS = "Comunicação Visual" Then
             sSQL = "SELECT sum(quantidade) as vSomaQuantProd, sum(total) as vSomaValorProd FROM pedidos_itens  WHERE (pedidos_itens.cod_pedido = " & Pedido & ")"
        End If
        
        Set rTotais = dbData.OpenRecordset(sSQL)
        If Not rTotais.EOF Then
            txtQuantPecas.Caption = Format(rTotais("vSomaQuantProd"), "000")
            txtTotalPecas.Caption = FormatNumber(rTotais("vSomaValorProd"), 2)
        Else
            txtQuantPecas.Caption = Format(0, "000")
            txtTotalPecas.Caption = FormatNumber(0, 2)
        End If
        sSQL = ""
        
        
        'SERVIÇOS
        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
              sSQL = sSQL & "SELECT sum(quantidade) as vSomaQuantServ, sum(total) as vSomaValorServ " & _
              "FROM OS_Servicos_Auto WHERE (cod_os = " & vCodOS & ")"
        ElseIf vTipoOS = "Recapadora" Then
            sSQL = sSQL & "SELECT sum(quantidade) as vSomaQuantServ, sum(total) as vSomaValorServ " & _
            "FROM OS_servicos_recapadora WHERE (cod_os = " & vCodOS & ")"
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
              sSQL = sSQL & "SELECT sum(quantidade) as vSomaQuantServ, sum(total) as vSomaValorServ " & _
              "FROM OS_Servicos_Auto WHERE (cod_os = " & vCodOS & ")"
        ElseIf vTipoOS = "Comunicação Visual" Then
              sSQL = sSQL & "SELECT sum(quantidade) as vSomaQuantServ, sum(total) as vSomaValorServ " & _
              "FROM OS_Servicos_Auto WHERE (cod_os = " & vCodOS & ")"
        End If
        
        Set rTotais = dbData.OpenRecordset(sSQL)
        'Debug.Print sSQL
        
        If Not rTotais.EOF Then
            txtQuantServicos.Caption = Format(rTotais("vSomaQuantServ"), "000")
            txtTotalServicos.Caption = FormatNumber(rTotais("vSomaValorServ"), 2)
        Else
            txtQuantServicos.Caption = Format(0, "000")
            txtTotalServicos.Caption = FormatNumber(0, 2)
        End If
        
        txtQuantGeral.Caption = Format(CInt(txtQuantPecas.Caption) + CInt(txtQuantServicos.Caption), "000")
        txtTotalPecasServicos.Caption = FormatNumber(CCur(txtTotalPecas.Caption) + CCur(txtTotalServicos.Caption), 2)
   End If


   
   If rIt.EOF Then
        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
           sSQL = "SELECT 'SERVIÇO' AS tipo_item, descricao, quantidade, preco, subtotal, codigo, '' as varFabricante, '', '', '', '', '', '', desconto, total " & _
           "FROM OS_Servicos_Auto WHERE (cod_os = " & vCodOS & ")"
        ElseIf vTipoOS = "Recapadora" Then
            sSQL = "SELECT 'SERVIÇO' AS tipo_item, descricao, quantidade, preco, subtotal, codigo, TIPO as var_TipoPneu, SERIE as var_serie, FOGO as var_fogo, ARO as var_aro, BANDA as var_banda, DOTE as var_dote, MEDIDA as var_medida, FABRICANTE as var_fabricante, desconto, total " & _
            "FROM OS_servicos_recapadora WHERE (cod_os = " & vCodOS & ")"
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
           sSQL = "SELECT 'SERVIÇO' AS tipo_item, descricao, quantidade, preco, subtotal, codigo, '' as varFabricante, '', '', '', '', '', '', desconto, total " & _
           "FROM OS_Servicos_Auto WHERE (cod_os = " & vCodOS & ")"
        ElseIf vTipoOS = "Comunicação Visual" Then
           sSQL = "SELECT 'SERVIÇO' AS tipo_item, descricao, quantidade, preco, subtotal, codigo, '' as varFabricante, '', '', '', '', '', '', desconto, total " & _
           "FROM OS_Servicos_Comunicacao WHERE (cod_os = " & vCodOS & ")"
        End If
   Else
        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
             sSQL = "SELECT 'PRODUTO' AS tipo_item, produtos.descricao, pedidos_itens.quantidade, pedidos_itens.preco, subtotal, pedidos_itens.codigo, produtos.Fabricante as varFabricante, desconto, total FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto WHERE (pedidos_itens.cod_pedido = " & Pedido & ")"
        ElseIf vTipoOS = "Recapadora" Then
             sSQL = "SELECT 'PRODUTO' AS tipo_item, produtos.descricao, pedidos_itens.quantidade, pedidos_itens.preco, subtotal, pedidos_itens.codigo, '' as var_TipoPneu, '' as var_serie, '' as var_fogo, '' as var_aro, '' as var_banda, '' as var_dote, '' as var_medida, '' as var_fabricante, desconto, total FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto WHERE (pedidos_itens.cod_pedido = " & Pedido & ")"
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
             sSQL = "SELECT 'PRODUTO' AS tipo_item, produtos.descricao, pedidos_itens.quantidade, pedidos_itens.preco, subtotal, pedidos_itens.codigo, produtos.Fabricante as varFabricante, desconto, total FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto WHERE (pedidos_itens.cod_pedido = " & Pedido & ")"
        ElseIf vTipoOS = "Comunicação Visual" Then
             sSQL = "SELECT 'PRODUTO' AS tipo_item, produtos.descricao, pedidos_itens.quantidade, pedidos_itens.preco, subtotal, pedidos_itens.codigo, produtos.Fabricante as varFabricante, desconto, total FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto WHERE (pedidos_itens.cod_pedido = " & Pedido & ")"
        End If
        
        If UCase(Tipo) = "OFICINA" Then
           sSQL = sSQL & " UNION "
             If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
                   sSQL = sSQL & "SELECT 'SERVIÇO' AS tipo_item, descricao, quantidade, preco, subtotal, codigo, '', desconto, total " & _
                   "FROM OS_Servicos_Auto WHERE (cod_os = " & vCodOS & ")"
             ElseIf vTipoOS = "Recapadora" Then
                 sSQL = sSQL & "SELECT 'SERVIÇO' AS tipo_item, descricao, quantidade, preco, subtotal, codigo, TIPO as var_TipoPneu, SERIE as var_serie, FOGO as var_fogo, ARO as var_aro, BANDA as var_banda, DOTE as var_dote, MEDIDA as var_medida, FABRICANTE as var_fabricante, desconto, total " & _
                 "FROM OS_servicos_recapadora WHERE (cod_os = " & vCodOS & ")"
             ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
                   sSQL = sSQL & "SELECT 'SERVIÇO' AS tipo_item, descricao, quantidade, preco, subtotal, codigo, '', desconto, total " & _
                   "FROM OS_Servicos_Auto WHERE (cod_os = " & vCodOS & ")"
             ElseIf vTipoOS = "Comunicação Visual" Then
                   sSQL = sSQL & "SELECT 'SERVIÇO' AS tipo_item, descricao, quantidade, preco, subtotal, codigo, '', desconto, total " & _
                   "FROM OS_Servicos_Auto WHERE (cod_os = " & vCodOS & ")"
             End If
        End If
   End If
   
   Set rIt = dbData.OpenRecordset(sSQL)
   
   Set rPc = dbData.OpenRecordset("SELECT * FROM parcelas WHERE (cod_pedido = " & Pedido & ") ORDER BY numero;")
   Set rFu = dbData.OpenRecordset("SELECT nome FROM funcionario WHERE (codigo = " & rOS("cod_funcionario") & ");")
    
    'parcelas
    If rPc("numero") = 1 And rPc("status") = True Then
      frEnt.Visible = True
    Else
      frEnt.Visible = False
    End If
   
   
    For i = 0 To 4
      frNumParc(i).Caption = ""
      frVencParc(i).Caption = ""
      frValorParc(i).Caption = ""
   Next
   
   Cont = 0
   
   Do While Not rPc.EOF
      If Cont < 5 Then
         frNumParc(Cont).Caption = Format(rPc("numero"), "00")
         frVencParc(Cont).Caption = Format(rPc("data"), "dd/mm/yy")
         frValorParc(Cont).Caption = Format(rPc("valor"), ocMONEY)
      End If
      Cont = Cont + 1
      rPc.MoveNext
   Loop

    If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Informática" Or vTipoOS = "Celular" Then
        txtNumero.Caption = "Nº " & Format(vCodOS, "000000")
    ElseIf vTipoOS = "Comunicação Visual" Then
        txtNumero.Caption = "Nº " & Format(vCodOS, "000000")
    ElseIf vTipoOS = "Comunicação Visual" Then
        txtNumero.Caption = "Nº " & Format(vCodOS, "000000")
    End If
   
   If rOS("cod_cliente") = 0 Then Exit Sub
   
   'DADOS DO CLIENTE
    txtCliente.Caption = rCl("nome")
    frNome.Caption = rCl("nome") & " - CPF: " & rCl("cpf")   'promissória
    txtEnd.Caption = IIf(IsNull(rCl![Endereco]) = True, "", rCl![Endereco]) & ", " & IIf(IsNull(rCl!Numero) = True, "", rCl!Numero) & " - " & IIf(IsNull(rCl!bairro) = True, "", rCl!bairro)
    txtRef2.Caption = IIf(IsNull(rCl!Ponto_de_referencia) = True, "", rCl!Ponto_de_referencia)
    txtCidade.Caption = IIf(IsNull(rCl!Cidade) = True, "", rCl!Cidade) & "-" & IIf(IsNull(rCl!Estado) = True, "", rCl!Estado) & "   FONE: " & IIf(IsNull(rCl!Telefone1) = True, "", rCl!Telefone1)
    txtCPF.Caption = IIf(IsNull(rCl!CPF) = True, "", rCl!CPF)
    txtRG.Caption = IIf(IsNull(rCl!rg) = True, "", rCl!rg)

    'DADOS DO VEICULO
    'If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
    '    frTitParc.Caption = "VEÍCULO"
    '    txtFabricante.Caption = IIf(IsNull(rEquip!Fabricante) = True, "", rEquip!Fabricante)
    '    txtModelo.Caption = IIf(IsNull(rEquip!Modelo) = True, "", rEquip!Modelo)
    '    txtPlaca.Caption = IIf(IsNull(rEquip!Placa) = True, "", rEquip!Placa)
    '    txtAno.Caption = IIf(IsNull(rEquip!Ano) = True, "", rEquip!Ano)
    '    txtCor.Caption = IIf(IsNull(rEquip!Cor) = True, "", rEquip!Cor)
    '    txtKM.Caption = IIf(IsNull(rEquip!KM) = True, "", rEquip!KM)
    'ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
    '    frTitParc.Caption = "EQUIPAMENTO"
    '    txtFabricante.Caption = IIf(IsNull(rEquip!Equipamento) = True, "", rEquip!Equipamento)
    '    txtModelo.Caption = IIf(IsNull(rEquip!Fabricante) = True, "", rEquip!Fabricante)
    '    txtAno.Caption = IIf(IsNull(rEquip!Modelo) = True, "", rEquip!Modelo)
    '    txtPlaca.Visible = False
    '    txtCor.Visible = False
    '    txtKM.Visible = False
    '    ReportField15.Visible = False
    '    ReportField14.Visible = False
    '    ReportField18.Visible = False
    '    ReportField2.Caption = "Equipamento:"
    '    ReportField10.Caption = "Fabricante:"
    '    ReportField12.Caption = "Modelo:"
    '    ReportField15.Caption = ""
    '    ReportField14.Caption = ""
    '    ReportField18.Caption = ""
    'End If
    
   'DADOS DO PEDIDO
   txtData.Caption = String(1, " ") + Format(rPd("data_compra"), "dd/mm/yy")
   txtVendedor.Caption = rFu("nome")
   txtVenda.Caption = UCase(ValidateNull(rPd("tipo_pagamento")))
   
   'DADOS DO FINANCEIRO
    txtSubtotal.Caption = " " & Format(rOS("SUBTOTAL"), ocMONEY)
    txtDesconto.Caption = " " & Format(rOS("ValorDescReal"), ocMONEY)
    txtTotal.Caption = " " & Format(rOS("TOTAL"), ocMONEY)
    rfValorNum.Caption = String(1, " ") + Format(rOS("total"), ocMONEY) 'promissoria
    frValorEst.Caption = String(1, " ") + Format(rOS("total"), ocMONEY) 'promissoria
    frValorEst.Caption = UCase(NumeroExtenso(rfValorNum.Caption, True)) 'promissoria
   
   'INSIRO OS ITENS
   If Not rIt.EOF Then rIt.MoveLast
   If Not rIt.BOF Then rIt.MoveFirst
   
   Relatorio.NumeroRegistros = Round((rIt.RecordCount / 31) + 0.49)
   Relatorio.NomeImpressora = var_ImpNormal
If varImpPDF = True Then
     Relatorio.Visualizar = False
Else
     Relatorio.Visualizar = True
End If
Relatorio.Ativar
varImpPDF = False
End Sub
Private Sub Form_Load()
On Error GoTo TrataErro

sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
Set r = dbData.OpenRecordset(sSQL)

rf1.Caption = r("fantasia")
rf2.Caption = r("razao")
rf3.Caption = r("endereco") & ", " & r("cidade") & "-" & r("estado")
rf4.Caption = "CNPJ: " & r("cnpj") & " - IE: " & r("ie") & " - CEL.: " & r("CELULAR")

If Not IsNull(r("caminho")) Then
   If Dir$(r("caminho")) <> "" Then Set imgLogo.Picture = LoadPicture(r("caminho"))
End If

If r.State <> 0 Then r.Close
Set r = Nothing
Exit Sub
   
TrataErro:
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Not rPd Is Nothing Then If rPd.State <> 0 Then rPd.Close
   If Not rCl Is Nothing Then If rCl.State <> 0 Then rCl.Close
   If Not rIt Is Nothing Then If rIt.State <> 0 Then rIt.Close
   If Not rPc Is Nothing Then If rPc.State <> 0 Then rPc.Close
   If Not rFu Is Nothing Then If rFu.State <> 0 Then rFu.Close
   If Not rOS Is Nothing Then If rOS.State <> 0 Then rOS.Close
   If Not rEquip Is Nothing Then If rEquip.State <> 0 Then rEquip.Close
End Sub

Private Sub Relatorio_IniciarSecao(ByVal Secao As ReportX.TSecao, ByVal Ordem As Byte)
   Dim i As Integer
   
   'produtos do pedido
   For i = 0 To 30
      txtTipo(i).Caption = ""
      txtDesc(i).Caption = ""
      txtQuant(i).Caption = ""
      txtUnit(i).Caption = ""
      txtTot(i).Caption = ""
      txtDesco(i).Caption = ""
      txtTotalProd(i).Caption = ""
      
      If Not rIt.EOF Then
         txtTipo(i).Caption = rIt("tipo_item")
         
         If vTipoOS = "Recapadora" Then
            If rIt("tipo_item") = "SERVIÇOS" Then
                txtDesc(i).Caption = String(1, " ") + ValidateNull(rIt("descricao")) & " | " & rIt("var_TipoPneu") & " | " & rIt("var_serie") & " | " & rIt("var_fogo") & " | " & rIt("var_aro") & " | " & rIt("var_banda") & " | " & rIt("var_dote") & " | " & rIt("var_medida") & " | " & rIt("var_fabricante") & " "
            Else
                txtDesc(i).Caption = String(1, " ") + ValidateNull(rIt("descricao"))
            End If
         Else
            txtDesc(i).Caption = String(1, " ") + rIt("descricao")
         End If

         txtQuant(i).Caption = rIt("quantidade")
         txtUnit(i).Caption = Format(rIt("preco"), "##,##0.00")
         txtTot(i).Caption = Format(rIt("Subtotal"), ocMONEY)
         txtDesco(i).Caption = Format(rIt("desconto"), ocMONEY)
         txtTotalProd(i).Caption = Format(rIt("total"), ocMONEY)
         
         rIt.MoveNext
      End If
   Next
   Exit Sub
End Sub


