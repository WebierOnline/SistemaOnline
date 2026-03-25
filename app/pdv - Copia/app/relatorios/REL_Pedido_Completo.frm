VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.Ocx"
Begin VB.Form REL_Pedido_Completo 
   Caption         =   "Form1"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   ScaleHeight     =   85.461
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   216.959
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain ReportMain1 
      Height          =   480
      Left            =   8220
      TabIndex        =   4
      Top             =   3900
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      MargemEsquerda  =   6
      MargemDireita   =   6
      Titulo          =   ""
      NomeImpressora  =   "\\servidor\IMPRESSORA1"
      Registrado      =   0   'False
      Visualizar      =   0   'False
   End
   Begin ReportX.ReportSection ReportSection2 
      Align           =   1  'Align Top
      Height          =   240
      Left            =   0
      Top             =   3135
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   423
      Begin ReportX.ReportField ReportField5 
         Height          =   210
         Left            =   8280
         TabIndex        =   2
         Top             =   0
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   370
         Campo           =   "Subtotal"
         Formato         =   "##,##0.00"
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
      End
      Begin ReportX.ReportField ReportField3 
         Height          =   210
         Left            =   7620
         TabIndex        =   0
         Top             =   0
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   370
         Campo           =   "quantidade"
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
      End
      Begin ReportX.ReportField ReportField2 
         Height          =   210
         Left            =   900
         TabIndex        =   1
         Top             =   0
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   370
         Campo           =   "var_desc"
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
      End
      Begin ReportX.ReportField ReportField7 
         Height          =   210
         Left            =   6660
         TabIndex        =   15
         Top             =   0
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   370
         Campo           =   "preco"
         Formato         =   "##,##0.00"
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
      End
      Begin ReportX.ReportField ReportField1 
         Height          =   210
         Left            =   60
         TabIndex        =   18
         Top             =   0
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   370
         Campo           =   "vCodProd"
         Formato         =   "000000"
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
      End
      Begin ReportX.ReportField ReportField4 
         Height          =   210
         Left            =   5400
         TabIndex        =   19
         Top             =   0
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   370
         Campo           =   "VFab"
         Formato         =   "000000"
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
      End
      Begin ReportX.ReportField ReportField8 
         Height          =   210
         Left            =   9300
         TabIndex        =   20
         Top             =   0
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   370
         Campo           =   "Desconto"
         Formato         =   "##,##0.00"
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
      End
      Begin ReportX.ReportField ReportField9 
         Height          =   210
         Left            =   10020
         TabIndex        =   21
         Top             =   0
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   370
         Campo           =   "total"
         Formato         =   "##,##0.00"
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
      End
   End
   Begin ReportX.ReportSection ReportSection3 
      Align           =   1  'Align Top
      Height          =   1410
      Left            =   0
      Top             =   3375
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   2487
      Tipo            =   7
      Begin ReportX.ReportField ReportField6 
         Height          =   225
         Left            =   120
         TabIndex        =   3
         Top             =   1140
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   397
         Campo           =   "=Página [Pagina] de [Paginas]"
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField lblSubTotal 
         Height          =   270
         Left            =   8880
         TabIndex        =   24
         Top             =   90
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   476
         Caption         =   "Subtotal:"
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
      Begin ReportX.ReportField rfSubTotal 
         Height          =   270
         Left            =   9780
         TabIndex        =   25
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
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
         Left            =   8880
         TabIndex        =   26
         Top             =   375
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   476
         Caption         =   "Desc. %:"
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
      Begin ReportX.ReportField rfDesc 
         Height          =   270
         Left            =   9780
         TabIndex        =   27
         Top             =   375
         Width           =   1215
         _ExtentX        =   2143
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
         Left            =   8880
         TabIndex        =   28
         Top             =   945
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   476
         Caption         =   "Total:"
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
      Begin ReportX.ReportField rfTotal 
         Height          =   270
         Left            =   9780
         TabIndex        =   29
         Top             =   945
         Width           =   1215
         _ExtentX        =   2143
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
      Begin ReportX.ReportField ReportField10 
         Height          =   270
         Left            =   8880
         TabIndex        =   30
         Top             =   660
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   476
         Caption         =   "Desc. R$:"
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
      Begin ReportX.ReportField txtDescontoRS 
         Height          =   270
         Left            =   9780
         TabIndex        =   31
         Top             =   660
         Width           =   1215
         _ExtentX        =   2143
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
      Begin ReportX.ReportField ReportField18 
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   420
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   344
         Caption         =   "....PARCELAS:...."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField rfInicio 
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   780
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   344
         Caption         =   "INICIO:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField rfTermino 
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   975
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   344
         Caption         =   "TÉRMINO:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField rfTotalParc 
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   344
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField rfDataPrim 
         Height          =   195
         Left            =   900
         TabIndex        =   36
         Top             =   780
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   344
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField rfDataUlt 
         Height          =   195
         Left            =   900
         TabIndex        =   37
         Top             =   975
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   344
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField rfValorPrim 
         Height          =   195
         Left            =   1500
         TabIndex        =   38
         Top             =   780
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   344
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField rfValorUlt 
         Height          =   195
         Left            =   1500
         TabIndex        =   39
         Top             =   975
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   344
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField rfNumPrim 
         Height          =   195
         Left            =   660
         TabIndex        =   40
         Top             =   780
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   344
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField rfNumUlt 
         Height          =   195
         Left            =   660
         TabIndex        =   41
         Top             =   975
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   344
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField rfDataEnt 
         Height          =   195
         Left            =   600
         TabIndex        =   42
         Top             =   240
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   344
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField frEnt 
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   344
         Caption         =   "PG"
         Alignment       =   2
         Mostrar         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField11 
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   60
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   344
         Caption         =   "....ENTRADA:...."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField rfValorEnt 
         Height          =   195
         Left            =   1380
         TabIndex        =   45
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   344
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin VB.Shape Shape6 
         BorderWidth     =   2
         Height          =   1215
         Left            =   8820
         Top             =   60
         Width           =   2235
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ASSINATURA"
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
         Left            =   3840
         TabIndex        =   14
         Top             =   1020
         Width           =   3720
      End
      Begin VB.Line Line2 
         X1              =   3840
         X2              =   7560
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         X1              =   11040
         X2              =   60
         Y1              =   0
         Y2              =   0
      End
   End
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   3135
      Left            =   0
      Top             =   0
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   5530
      Tipo            =   2
      Begin ReportX.ReportField rf2 
         Height          =   300
         Left            =   3720
         TabIndex        =   5
         Top             =   540
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   529
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         Height          =   510
         Left            =   3720
         TabIndex        =   6
         Top             =   60
         Width           =   7260
         _ExtentX        =   12806
         _ExtentY        =   900
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Impact"
            Size            =   15.75
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
         Height          =   300
         Left            =   3720
         TabIndex        =   7
         Top             =   780
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   529
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         Height          =   300
         Left            =   3720
         TabIndex        =   8
         Top             =   1020
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   529
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDHead 
         Height          =   390
         Left            =   60
         TabIndex        =   9
         Top             =   1440
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   688
         Caption         =   "RELATORIO COMPLETO DE PEDIDO"
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
         Borda           =   9
         BordaEstilo     =   1
         BackColor       =   14737632
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField rfForma 
         Height          =   210
         Left            =   840
         TabIndex        =   46
         Top             =   2100
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rfCliente 
         Height          =   210
         Left            =   840
         TabIndex        =   47
         Top             =   1860
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   370
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rfData 
         Height          =   210
         Left            =   840
         TabIndex        =   48
         Top             =   2340
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   370
         Formato         =   "DD/MM/YY"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rfFunc 
         Height          =   210
         Left            =   840
         TabIndex        =   49
         Top             =   2580
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   370
         Formato         =   "DD/MM/YY"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line4 
         BorderStyle     =   3  'Dot
         X1              =   60
         X2              =   11100
         Y1              =   3060
         Y2              =   3060
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FORMA:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   60
         TabIndex        =   53
         Top             =   2100
         Width           =   645
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CLIENTE:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   60
         TabIndex        =   52
         Top             =   1860
         Width           =   645
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DATA:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   60
         TabIndex        =   51
         Top             =   2340
         Width           =   645
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FUNC.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   60
         TabIndex        =   50
         Top             =   2580
         Width           =   645
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         X1              =   60
         X2              =   11100
         Y1              =   2820
         Y2              =   2820
      End
      Begin VB.Label Label15 
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
         Left            =   10380
         TabIndex        =   23
         Top             =   2850
         Width           =   555
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESC."
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
         Left            =   9420
         TabIndex        =   22
         Top             =   2850
         Width           =   465
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FABRICANTE"
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
         Left            =   5400
         TabIndex        =   17
         Top             =   2850
         Width           =   1005
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PREÇO"
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
         Left            =   6900
         TabIndex        =   16
         Top             =   2850
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QTDE"
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
         TabIndex        =   13
         Top             =   2850
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIÇÃO"
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
         Left            =   900
         TabIndex        =   12
         Top             =   2850
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CÓD."
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
         Left            =   60
         TabIndex        =   11
         Top             =   2850
         Width           =   390
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUBTOTAL"
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
         Left            =   8340
         TabIndex        =   10
         Top             =   2850
         Width           =   870
      End
      Begin VB.Image imgLogo 
         Height          =   1215
         Left            =   180
         Stretch         =   -1  'True
         Top             =   120
         Width           =   3315
      End
      Begin VB.Shape Shape1 
         Height          =   1335
         Left            =   60
         Shape           =   4  'Rounded Rectangle
         Top             =   60
         Width           =   11055
      End
   End
End
Attribute VB_Name = "REL_Pedido_Completo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rPc As ADODB.Recordset
Public Sub Mostrar_Parcelas(Pedido As Long)
Set rPc = dbData.OpenRecordset("SELECT * FROM parcelas WHERE (cod_pedido = " & Pedido & ") ORDER BY numero;")

If rPc.RecordCount < 1 Then
    'rfNumEnt.Caption = ""
    rfDataEnt.Caption = ""
    rfValorEnt.Caption = ""
    rfNumPrim.Caption = ""
    rfDataPrim.Caption = ""
    rfValorPrim.Caption = ""
    rfNumUlt.Caption = ""
    rfDataUlt.Caption = ""
    rfValorUlt.Caption = ""
    ReportField11.Caption = ""
    ReportField18.Caption = ""
    rfInicio.Caption = ""
    rfTermino.Caption = ""
    
Else
    'rfNumEnt.Caption = ""
    rfDataEnt.Caption = ""
    rfValorEnt.Caption = ""
    rfNumPrim.Caption = ""
    rfDataPrim.Caption = ""
    rfValorPrim.Caption = ""
    rfNumUlt.Caption = ""
    rfDataUlt.Caption = ""
    rfValorUlt.Caption = ""
    ReportField11.Caption = "....ENTRADA:...."
    ReportField18.Caption = "....PARCELAS:...."
    rfInicio.Caption = "INICIO:"
    rfTermino.Caption = "TÉRMINO:"
    
    If rPc("numero") = 1 And CBool(rPc("status")) = True Then
       frEnt.Mostrar = True
        If Not rPc.BOF Then
            rPc.MoveFirst
            frEnt.Caption = "PG"
            'rfNumEnt.Caption = Format(rPc("numero"), "00")
            rfDataEnt.Caption = Format(rPc("data"), "dd/mm/yy")
            rfValorEnt.Caption = FormatNumber(rPc("valor"), 2)
        End If
    Else
       frEnt.Mostrar = True
       frEnt.Caption = FormatNumber(0, 2)
       rfDataEnt.Caption = "Sem Entrada"
    End If
    
    rfTotalParc.Caption = "Total de Parcelas: " & Format(rPc.RecordCount, "00")
    
    If rPc.RecordCount > 1 Then
        If Not rPc.BOF Then
            rPc.MoveFirst
            rfInicio.Caption = "INÍCIO:"
            rfNumPrim.Caption = Format(rPc("numero"), "00")
            rfDataPrim.Caption = Format(rPc("data"), "dd/mm/yy")
            rfValorPrim.Caption = FormatNumber(rPc("valor"), 2)
        End If
        
        If Not rPc.EOF Then
            rPc.MoveLast
            rfTermino.Caption = "TÉRMINO:"
            rfNumUlt.Caption = Format(rPc("numero"), "00")
            rfDataUlt.Caption = Format(rPc("data"), "dd/mm/yy")
            rfValorUlt.Caption = FormatNumber(rPc("valor"), 2)
        End If
    Else
        If Not rPc.BOF Then
            rPc.MoveFirst
            rfInicio.Caption = "VENC.:"
            rfNumPrim.Caption = Format(rPc("numero"), "00")
            rfDataPrim.Caption = Format(rPc("data"), "dd/mm/yy")
            rfValorPrim.Caption = FormatNumber(rPc("valor"), 2)
        End If
        
        If Not rPc.EOF Then
            'rPc.MoveLast
            rfTermino.Caption = ""
            rfNumUlt.Caption = ""
            rfDataUlt.Caption = ""
            rfValorUlt.Caption = ""
        End If
    End If
End If

End Sub

Private Sub Form_Activate()
'Mostrar_Parcelas (rfCodPedido.Caption)
End Sub

Private Sub Form_Load()
   On Error GoTo TrataErro
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
   Set r = dbData.OpenRecordset(sSQL)
   
   rf1.Caption = r("fantasia")
   rf2.Caption = r("razao")
   rf3.Caption = r("endereco") & ", " & r("cidade") & "-" & r("estado")
   rf4.Caption = "CNPJ: " & r("cnpj") & " - IE: " & r("ie") & " - TELEFONE: " & r("telefone")
   
   If Not IsNull(r("caminho")) Then
      If Dir$(r("caminho")) <> "" Then Set imgLogo.Picture = LoadPicture(r("caminho"))
   End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   Exit Sub
   
TrataErro:
End Sub

