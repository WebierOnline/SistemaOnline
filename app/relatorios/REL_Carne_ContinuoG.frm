VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.Ocx"
Begin VB.Form REL_Carne_ContinuoG 
   Caption         =   "Promissórias"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   9.075
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   20.77
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   60
      TabIndex        =   0
      Top             =   4440
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Divisao         =   2
      Regua           =   -1  'True
      Escala          =   7
      MargemEsquerda  =   0.2
      Titulo          =   ""
      Registrado      =   0   'False
   End
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   4125
      Left            =   0
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7276
      Begin ReportX.ReportField lblCliente1 
         Height          =   270
         Left            =   360
         TabIndex        =   1
         Top             =   1260
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   476
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField lblDuplicata1 
         Height          =   270
         Left            =   360
         TabIndex        =   2
         Top             =   1860
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   476
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField lblVenc1 
         Height          =   270
         Left            =   2460
         TabIndex        =   3
         Top             =   1860
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   476
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField lblPrest1 
         Height          =   270
         Left            =   1800
         TabIndex        =   4
         Top             =   1860
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   476
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField lblVlrPrest1 
         Height          =   270
         Left            =   3600
         TabIndex        =   5
         Top             =   1860
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   476
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField lblAtraso1 
         Height          =   510
         Left            =   1980
         TabIndex        =   6
         Top             =   2520
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   900
         Linhas          =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField lblJuros1 
         Height          =   510
         Left            =   2820
         TabIndex        =   7
         Top             =   2520
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   900
         Linhas          =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField rf2 
         Height          =   180
         Left            =   360
         TabIndex        =   8
         Top             =   480
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   318
         Caption         =   ""
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField rf1 
         Height          =   285
         Left            =   480
         TabIndex        =   9
         Top             =   210
         Width           =   4260
         _ExtentX        =   7514
         _ExtentY        =   503
         Caption         =   ""
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField rf3 
         Height          =   180
         Left            =   360
         TabIndex        =   10
         Top             =   660
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   318
         Caption         =   ""
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField rf4 
         Height          =   180
         Left            =   300
         TabIndex        =   11
         Top             =   840
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   318
         Caption         =   ""
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField lblTotal1 
         Height          =   510
         Left            =   3720
         TabIndex        =   19
         Top             =   2520
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   900
         Linhas          =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField lblDataPgto1 
         Height          =   510
         Left            =   360
         TabIndex        =   21
         Top             =   2520
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   900
         Linhas          =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField rf22 
         Height          =   180
         Left            =   6780
         TabIndex        =   23
         Top             =   570
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   318
         Caption         =   ""
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField rf12 
         Height          =   285
         Left            =   6960
         TabIndex        =   24
         Top             =   300
         Width           =   4260
         _ExtentX        =   7514
         _ExtentY        =   503
         Caption         =   ""
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField rf32 
         Height          =   180
         Left            =   6780
         TabIndex        =   25
         Top             =   750
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   318
         Caption         =   ""
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField rf42 
         Height          =   180
         Left            =   6780
         TabIndex        =   26
         Top             =   930
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   318
         Caption         =   ""
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField lblCliente2 
         Height          =   270
         Left            =   5160
         TabIndex        =   27
         Top             =   1260
         Width           =   6330
         _ExtentX        =   11165
         _ExtentY        =   476
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField lblDuplicata2 
         Height          =   270
         Left            =   5160
         TabIndex        =   28
         Top             =   1860
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   476
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField lblVenc2 
         Height          =   270
         Left            =   7500
         TabIndex        =   29
         Top             =   1860
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   476
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField lblPrest2 
         Height          =   270
         Left            =   6660
         TabIndex        =   30
         Top             =   1860
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   476
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField lblVlrPrest2 
         Height          =   270
         Left            =   8640
         TabIndex        =   31
         Top             =   1860
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   476
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField lblAtraso2 
         Height          =   510
         Left            =   6840
         TabIndex        =   32
         Top             =   2520
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   900
         Linhas          =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField lblJuros2 
         Height          =   510
         Left            =   7740
         TabIndex        =   33
         Top             =   2520
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   900
         Linhas          =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField lblTotal2 
         Height          =   510
         Left            =   9060
         TabIndex        =   34
         Top             =   2520
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   900
         Linhas          =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField lblDataPgto2 
         Height          =   510
         Left            =   5160
         TabIndex        =   35
         Top             =   2520
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   900
         Linhas          =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         AlinhamentoVertical=   1
      End
      Begin VB.Line Line2 
         X1              =   720
         X2              =   4080
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ASSINATURA"
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
         Left            =   1920
         TabIndex        =   46
         Top             =   3660
         Width           =   1035
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ASSINATURA"
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
         Left            =   7740
         TabIndex        =   45
         Top             =   3660
         Width           =   1035
      End
      Begin VB.Line Line1 
         X1              =   6540
         X2              =   9900
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CLIENTE:"
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
         Left            =   5160
         TabIndex        =   44
         Top             =   1020
         Width           =   705
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PEDIDO:"
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
         Left            =   5160
         TabIndex        =   43
         Top             =   1620
         Width           =   615
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PARC.:"
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
         Left            =   6660
         TabIndex        =   42
         Top             =   1620
         Width           =   540
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VALOR:"
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
         Left            =   8700
         TabIndex        =   41
         Top             =   1620
         Width           =   615
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENC.:"
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
         Left            =   7500
         TabIndex        =   40
         Top             =   1620
         Width           =   525
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ATRASO"
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
         Left            =   6840
         TabIndex        =   39
         Top             =   2280
         Width           =   675
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JUROS"
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
         Left            =   7740
         TabIndex        =   38
         Top             =   2280
         Width           =   525
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
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
         Left            =   9060
         TabIndex        =   37
         Top             =   2280
         Width           =   555
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PGTO.:"
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
         Left            =   5160
         TabIndex        =   36
         Top             =   2280
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PGTO.:"
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
         Left            =   360
         TabIndex        =   22
         Top             =   2280
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
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
         Left            =   3720
         TabIndex        =   20
         Top             =   2280
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JUROS"
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
         Left            =   2820
         TabIndex        =   18
         Top             =   2280
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ATRASO"
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
         Left            =   1980
         TabIndex        =   17
         Top             =   2280
         Width           =   675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENC.:"
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
         Left            =   2460
         TabIndex        =   16
         Top             =   1620
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VALOR:"
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
         Left            =   3660
         TabIndex        =   15
         Top             =   1620
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PARC.:"
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
         Left            =   1800
         TabIndex        =   14
         Top             =   1620
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PEDIDO:"
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
         Left            =   360
         TabIndex        =   13
         Top             =   1620
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CLIENTE:"
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
         Left            =   360
         TabIndex        =   12
         Top             =   1020
         Width           =   705
      End
      Begin VB.Shape Shape2 
         Height          =   3795
         Left            =   300
         Shape           =   4  'Rounded Rectangle
         Top             =   180
         Width           =   4635
      End
      Begin VB.Shape Shape1 
         Height          =   3795
         Left            =   5100
         Shape           =   4  'Rounded Rectangle
         Top             =   180
         Width           =   6495
      End
      Begin VB.Image imgLogo 
         Height          =   795
         Left            =   5220
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1515
      End
   End
End
Attribute VB_Name = "REL_Carne_ContinuoG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCliente As Recordset
Dim rsPedidos As Recordset
Dim rsParcelas As Recordset
Dim Cod_Pedido As Integer
Dim totalRegistros As Long


Public Sub loadPromissoria(Pedido As Integer, PARCELA As Integer)
'colocar o nome da maquina na barra de status
Dim var_Impressora As String
'Dim oIni As Ini
Dim vNumeroParcelas As Integer

'abrindo arquivo .ini
Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"

'nome da maquina
var_Impressora = oIni.LerTexto("IMPRESSORA_NORMAL", "impressora")

Dim Prt As Printer
Dim oldPrinter As String

'Armazena o nome da impressora atual
oldPrinter = Printer.DeviceName

' Find and use the printer just selected in the ListBox
For Each Prt In Printers
   If Prt.DeviceName = var_Impressora Then
      Set Printer = Prt
      Exit For
   End If
Next

Cod_Pedido = Pedido

Set rsPedidos = dbData.OpenRecordset("SELECT * FROM pedidos WHERE (cod_pedido = " & Pedido & ");")
Set rsCliente = dbData.OpenRecordset("SELECT * FROM cliente WHERE (codigo = " & rsPedidos("cod_cliente") & ");")

If PARCELA = 0 Then
   Set rsParcelas = dbData.OpenRecordset("SELECT * FROM parcelas WHERE (cod_pedido = " & Pedido & ");", totalRegistros)
Else
   Set rsParcelas = dbData.OpenRecordset("SELECT * FROM parcelas WHERE (cod_pedido = " & Pedido & ") AND (numero = " & PARCELA & ");", totalRegistros)
End If
   vNumeroParcelas = rsParcelas.RecordCount

Relatorio.NumeroRegistros = totalRegistros
Relatorio.NomeImpressora = var_Impressora
Relatorio.Ativar
End Sub

Private Sub Rpx_MsgErro(Numero As Long)
   Dim Msg As String
   
   If Numero < 0 Then
      ' Mensagens de erro previstas
      Select Case Numero - vbObjectError
         Case 1001: Msg = "É necessário existir uma impressora instalada no Windows"
         Case 1002: Msg = "Não há registros a imprimir"
         Case 1003: Msg = "Não foi definida a seção de detalhe do relatório"
         Case 1004: Msg = "A configuração das seções de grupos está incorreta"
         Case 1005: Msg = "Foi definido um cursor do tipo Forward-Only para o recordset do relatório."
         Case 1006: Msg = "A página configurada para o relatório não possuí espaço suficiente para a impressão"
         Case 1007: Msg = "Já existe um relatório em andamento"
      End Select
      
      ShowMsg Msg, vbInformation
   Else
      ' Mensagens não previstas. Isso pode significar um erro
      ' interno no ReportX. Se isso acontecer, por favor reporte isso
      ' através de e-mail para ser corrigido.
      ShowMsg "Erro não previsto:" & Numero & vbCrLf & Error(Numero) & _
         IIf(Err.Number <> 0, vbCrLf + Err.Description, ""), vbCritical
   End If
End Sub

Private Sub Form_Load()
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
Set r = dbData.OpenRecordset(sSQL)

rf1.Caption = r("fantasia")
rf2.Caption = r("razao")
rf3.Caption = r("endereco") & ", " & r("cidade") & "-" & r("estado")
rf4.Caption = "CNPJ: " & r("cnpj") & " - IE: " & r("ie") & " - FONE: " & IIf(r("telefone") = "", r("celular"), r("telefone"))
rf12.Caption = r("fantasia")
rf22.Caption = r("razao")
rf32.Caption = r("endereco") & ", " & r("cidade") & "-" & r("estado")
rf42.Caption = "CNPJ: " & r("cnpj") & " - IE: " & r("ie") & " - FONE: " & IIf(r("telefone") = "", r("celular"), r("telefone"))

If Not IsNull(r("caminho")) Then
   If Dir$(r("caminho")) <> "" Then Set imgLogo.Picture = LoadPicture(r("caminho"))
End If

If r.State <> 0 Then r.Close
Set r = Nothing
Exit Sub
End Sub

Private Sub Relatorio_Erro(ByVal Numero As Long)
   Rpx_MsgErro Numero
End Sub

Private Sub Relatorio_ImprimiuRegistro(Cancelar As Boolean)
   If Not rsParcelas.EOF Then rsParcelas.MoveNext
End Sub

Private Sub Relatorio_IniciarSecao(ByVal Secao As ReportX.TSecao, ByVal Ordem As Byte)

   If Not rsParcelas.EOF Then
      lblCliente1.Caption = " " & rsCliente!nome
      lblDuplicata1.Caption = " " & Format(Cod_Pedido, "00000") & "/" & Format(rsParcelas!Numero, "00")
      lblVenc1.Caption = " " & Format(rsParcelas!Data, "dd/mm/yy")
      lblPrest1.Caption = " " & Format(rsParcelas!Numero, "00") & "/" & Format(totalRegistros, "00")
      lblVlrPrest1.Caption = " " & Format(rsParcelas!Valor, "##,##0.00") & " "
      
      lblCliente2.Caption = " " & rsCliente!nome
      lblDuplicata2.Caption = " " & Format(Cod_Pedido, "00000") & "/" & Format(rsParcelas!Numero, "00")
      lblVenc2.Caption = " " & Format(rsParcelas!Data, "dd/mm/yy")
      lblPrest2.Caption = " " & Format(rsParcelas!Numero, "00") & "/" & Format(totalRegistros, "00")
      lblVlrPrest2.Caption = Format(rsParcelas!Valor, "##,##0.00") & " "
      
      'txtParcela.Caption = rsParcelas!Valor
      
      'Aqui é as data de vencimento
      'txtData.Caption = LCase(NumeroExtenso(Format(rsParcelas!Data, "dd"), False)) & " de " & LCase(Format(rsParcelas!Data, "MMMM")) & " de " & Format(rsParcelas!Data, "yyyy")
      
      'Descomente aqui caso as data lá seja a data atual
      'txtData.Caption = UCase(numeroExtenso(Format(Date, "dd"), False)) & " de " & UCase(Format(Date, "MMMM")) & " de " & Format(Date, "yyyy")
      
      'txtValor.Caption = NumeroExtenso(rsParcelas!Valor)
      
      'DADOS DO CLIENTE
      'txtVencimento.Caption = Format(rsParcelas!Data, "dd/mm/yy")
      'txtEmissao.Caption = Format(Date, "dd/mm/yyyy")
      'txtEmitente.Caption = rsCliente!NOME
      'txtCPF.Caption = IIf(IsNull(rsCliente!cpf) = True, "", rsCliente!cpf)
      'txtEndereco.Caption = rsCliente![ENDERECO] & " , " & rsCliente!BAIRRO
   End If
End Sub


