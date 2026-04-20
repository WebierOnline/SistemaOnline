VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.Ocx"
Begin VB.Form REL_OS_Completo 
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
   End
   Begin ReportX.ReportSection ReportSection2 
      Align           =   1  'Align Top
      Height          =   240
      Left            =   0
      Top             =   3255
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
         Left            =   1860
         TabIndex        =   1
         Top             =   0
         Width           =   4755
         _ExtentX        =   8387
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
         Left            =   1080
         TabIndex        =   18
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   120
         TabIndex        =   19
         Top             =   0
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
         Campo           =   "vTipo"
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
      Top             =   3495
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   2487
      Tipo            =   7
      Begin ReportX.ReportField ReportField6 
         Height          =   225
         Left            =   120
         TabIndex        =   3
         Top             =   960
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
      Begin ReportX.ReportField ReportField20 
         Height          =   270
         Left            =   180
         TabIndex        =   53
         Top             =   90
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
         TabIndex        =   54
         Top             =   90
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
         Height          =   255
         Left            =   180
         TabIndex        =   55
         Top             =   360
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
         TabIndex        =   56
         Top             =   360
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
      Begin ReportX.ReportField ReportField22 
         Height          =   270
         Left            =   180
         TabIndex        =   57
         Top             =   630
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
         TabIndex        =   58
         Top             =   630
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
         TabIndex        =   59
         Top             =   90
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
         TabIndex        =   60
         Top             =   360
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
         TabIndex        =   61
         Top             =   630
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
      Begin VB.Shape Shape7 
         BorderWidth     =   2
         Height          =   885
         Left            =   120
         Top             =   60
         Width           =   2715
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
      Height          =   3255
      Left            =   0
      Top             =   0
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   5741
      Tipo            =   2
      Begin ReportX.ReportField txtParecer 
         Height          =   870
         Left            =   8700
         TabIndex        =   82
         Top             =   2040
         Visible         =   0   'False
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   1535
         Caption         =   ""
         WordWrap        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BordaLagura     =   2
         AlturaLivre     =   -1  'True
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
      Begin ReportX.ReportField txtParecerTitulo 
         Height          =   270
         Left            =   8700
         TabIndex        =   83
         Top             =   1815
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   476
         Caption         =   "PARECER TÉCNICO:"
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
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtFabricante 
         Height          =   210
         Left            =   5760
         TabIndex        =   41
         Top             =   2100
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   370
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
         BordaLagura     =   2
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField12 
         Height          =   210
         Left            =   4860
         TabIndex        =   42
         Top             =   2078
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   370
         Caption         =   "Fabricante:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin ReportX.ReportField txtModelo 
         Height          =   210
         Left            =   7560
         TabIndex        =   43
         Top             =   2100
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   370
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
         BordaLagura     =   2
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField14 
         Height          =   210
         Left            =   6900
         TabIndex        =   44
         Top             =   2100
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   370
         Caption         =   "Modelo:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin ReportX.ReportField rfForma 
         Height          =   210
         Left            =   840
         TabIndex        =   32
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
         TabIndex        =   33
         Top             =   1860
         Width           =   3915
         _ExtentX        =   6906
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
         Left            =   900
         TabIndex        =   34
         Top             =   2340
         Width           =   1035
         _ExtentX        =   1826
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
         Left            =   3600
         TabIndex        =   35
         Top             =   2580
         Width           =   1155
         _ExtentX        =   2037
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
      Begin ReportX.ReportField frTitParc 
         Height          =   270
         Left            =   4860
         TabIndex        =   40
         Top             =   1815
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   476
         Caption         =   "VEÍCULO:"
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
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtAno 
         Height          =   210
         Left            =   5760
         TabIndex        =   45
         Top             =   2300
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   370
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
         BordaLagura     =   2
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField15 
         Height          =   210
         Left            =   4860
         TabIndex        =   46
         Top             =   2286
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   370
         Caption         =   "Ano:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin ReportX.ReportField txtPlaca 
         Height          =   210
         Left            =   7560
         TabIndex        =   47
         Top             =   2295
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   370
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
         BordaLagura     =   2
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField16 
         Height          =   210
         Left            =   6900
         TabIndex        =   48
         Top             =   2300
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   370
         Caption         =   "Placa:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin ReportX.ReportField txtCor 
         Height          =   210
         Left            =   5760
         TabIndex        =   49
         Top             =   2500
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   370
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
         BordaLagura     =   2
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField17 
         Height          =   210
         Left            =   4860
         TabIndex        =   50
         Top             =   2494
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   370
         Caption         =   "Cor:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin ReportX.ReportField txtKM 
         Height          =   210
         Left            =   7560
         TabIndex        =   51
         Top             =   2505
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   370
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
         BordaLagura     =   2
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField19 
         Height          =   210
         Left            =   6900
         TabIndex        =   52
         Top             =   2500
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   370
         Caption         =   "KM:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin ReportX.ReportField ReportField18 
         Height          =   195
         Left            =   9240
         TabIndex        =   62
         Top             =   2160
         Visible         =   0   'False
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
         Left            =   9240
         TabIndex        =   63
         Top             =   2520
         Visible         =   0   'False
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
         Left            =   9240
         TabIndex        =   64
         Top             =   2715
         Visible         =   0   'False
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
         Left            =   9240
         TabIndex        =   65
         Top             =   2340
         Visible         =   0   'False
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
         Left            =   10020
         TabIndex        =   66
         Top             =   2520
         Visible         =   0   'False
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
         Left            =   10020
         TabIndex        =   67
         Top             =   2715
         Visible         =   0   'False
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
         Left            =   10620
         TabIndex        =   68
         Top             =   2520
         Visible         =   0   'False
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
         Left            =   10620
         TabIndex        =   69
         Top             =   2715
         Visible         =   0   'False
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
         Left            =   9780
         TabIndex        =   70
         Top             =   2520
         Visible         =   0   'False
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
         Left            =   9780
         TabIndex        =   71
         Top             =   2715
         Visible         =   0   'False
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
         Left            =   9720
         TabIndex        =   72
         Top             =   1980
         Visible         =   0   'False
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
         Left            =   9240
         TabIndex        =   73
         Top             =   1980
         Visible         =   0   'False
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
         Left            =   9240
         TabIndex        =   74
         Top             =   1800
         Visible         =   0   'False
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
         Left            =   10500
         TabIndex        =   75
         Top             =   1980
         Visible         =   0   'False
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
      Begin ReportX.ReportField rfTecnico 
         Height          =   210
         Left            =   1020
         TabIndex        =   76
         Top             =   2580
         Width           =   1455
         _ExtentX        =   2566
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
      Begin ReportX.ReportField txtChassi 
         Height          =   210
         Left            =   5760
         TabIndex        =   78
         Top             =   2700
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   370
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
         BordaLagura     =   2
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField24 
         Height          =   210
         Left            =   4860
         TabIndex        =   79
         Top             =   2700
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   370
         Caption         =   "Chassi:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin ReportX.ReportField rfDataSaida 
         Height          =   210
         Left            =   3420
         TabIndex        =   80
         Top             =   2340
         Width           =   1095
         _ExtentX        =   1931
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
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SÁIDA(previsão):"
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
         Left            =   2100
         TabIndex        =   81
         Top             =   2340
         Width           =   1275
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MECÂNICO:"
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
         TabIndex        =   77
         Top             =   2580
         Width           =   840
      End
      Begin VB.Line Line6 
         BorderStyle     =   3  'Dot
         X1              =   8640
         X2              =   8640
         Y1              =   1800
         Y2              =   2940
      End
      Begin VB.Line Line5 
         BorderStyle     =   3  'Dot
         X1              =   4800
         X2              =   4800
         Y1              =   1740
         Y2              =   2940
      End
      Begin VB.Line Line4 
         BorderStyle     =   3  'Dot
         X1              =   60
         X2              =   11100
         Y1              =   3180
         Y2              =   3180
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
         TabIndex        =   39
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
         TabIndex        =   38
         Top             =   1860
         Width           =   645
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ENTRADA:"
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
         TabIndex        =   37
         Top             =   2340
         Width           =   780
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "RECEPÇÃO:"
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
         Left            =   2580
         TabIndex        =   36
         Top             =   2580
         Width           =   870
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         X1              =   60
         X2              =   11100
         Y1              =   2940
         Y2              =   2940
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
         Top             =   2970
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
         Top             =   2970
         Width           =   465
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO"
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
         TabIndex        =   17
         Top             =   2970
         Width           =   375
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
         Top             =   2970
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
         Top             =   2970
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
         Left            =   1860
         TabIndex        =   12
         Top             =   2970
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
         Left            =   1080
         TabIndex        =   11
         Top             =   2970
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
         Top             =   2970
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
Attribute VB_Name = "REL_OS_Completo"
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
   rf4.Caption = "CNPJ: " & r("cnpj") & " - IE: " & r("ie") & " - CEL.: " & r("CELULAR")
   
   If Not IsNull(r("caminho")) Then
      If Dir$(r("caminho")) <> "" Then Set imgLogo.Picture = LoadPicture(r("caminho"))
   End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
   Exit Sub
   
TrataErro:
End Sub

