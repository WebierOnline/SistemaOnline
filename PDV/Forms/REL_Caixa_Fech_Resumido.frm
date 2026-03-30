VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.Ocx"
Begin VB.Form REL_Caixa_Fech_Resumido 
   Caption         =   "Form1"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   ScaleHeight     =   193.146
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   216.959
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain ReportMain1 
      Height          =   480
      Left            =   6540
      TabIndex        =   1
      Top             =   4320
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      MargemEsquerda  =   6
      MargemDireita   =   6
      Titulo          =   ""
      NomeImpressora  =   "IMPRESSORA1"
      Registrado      =   0   'False
   End
   Begin ReportX.ReportSection ReportSection3 
      Align           =   1  'Align Top
      Height          =   2010
      Left            =   0
      Top             =   10995
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   3545
      Tipo            =   7
      Begin ReportX.ReportField ReportField6 
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   1740
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   344
         Campo           =   "=Página [Pagina] de [Paginas]"
         Formula         =   -1  'True
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
      Begin ReportX.ReportField rfHoraA 
         Height          =   195
         Left            =   3780
         TabIndex        =   8
         Top             =   660
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   344
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
      End
      Begin ReportX.ReportField rfCodUsuarioA 
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   660
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   344
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
      End
      Begin ReportX.ReportField rfDataA 
         Height          =   195
         Left            =   2640
         TabIndex        =   10
         Top             =   660
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   344
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
      End
      Begin ReportX.ReportField rfNomeUsuarioA 
         Height          =   195
         Left            =   600
         TabIndex        =   11
         Top             =   660
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   344
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
      End
      Begin ReportX.ReportField rfHoraF 
         Height          =   195
         Left            =   3780
         TabIndex        =   16
         Top             =   1500
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   344
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
      End
      Begin ReportX.ReportField rfCodUsuarioF 
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   1500
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   344
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
      End
      Begin ReportX.ReportField rfDataF 
         Height          =   195
         Left            =   2640
         TabIndex        =   18
         Top             =   1500
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   344
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
      End
      Begin ReportX.ReportField rfNomeUsuarioF 
         Height          =   195
         Left            =   600
         TabIndex        =   19
         Top             =   1500
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   344
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
      End
      Begin ReportX.ReportField rfCaixa 
         Height          =   195
         Left            =   6180
         TabIndex        =   26
         Top             =   1560
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   344
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
      End
      Begin ReportX.ReportField rfCodCaixa 
         Height          =   195
         Left            =   8040
         TabIndex        =   27
         Top             =   1560
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   344
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
      End
      Begin ReportX.ReportField rfSituacao 
         Height          =   195
         Left            =   9780
         TabIndex        =   28
         Top             =   1560
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   344
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
      End
      Begin ReportX.ReportSection ReportSection2 
         Height          =   240
         Left            =   0
         Top             =   -240
         Width           =   12300
         _ExtentX        =   21696
         _ExtentY        =   423
      End
      Begin VB.Line Line5 
         BorderStyle     =   3  'Dot
         X1              =   10980
         X2              =   0
         Y1              =   60
         Y2              =   60
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Situação:"
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
         Left            =   8940
         TabIndex        =   29
         Top             =   1560
         Width           =   810
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa:"
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
         Left            =   5580
         TabIndex        =   25
         Top             =   1560
         Width           =   540
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. Caixa:"
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
         Left            =   7020
         TabIndex        =   24
         Top             =   1560
         Width           =   975
      End
      Begin VB.Shape Shape4 
         Height          =   1635
         Left            =   120
         Top             =   120
         Width           =   4875
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora:"
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
         Left            =   3780
         TabIndex        =   23
         Top             =   1260
         Width           =   480
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "....FECHAMENTO:"
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
         Left            =   240
         TabIndex        =   22
         Top             =   1020
         Width           =   1575
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário:"
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
         Left            =   240
         TabIndex        =   21
         Top             =   1260
         Width           =   705
      End
      Begin VB.Label Label19 
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
         Left            =   2640
         TabIndex        =   20
         Top             =   1260
         Width           =   480
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora:"
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
         Left            =   3780
         TabIndex        =   15
         Top             =   420
         Width           =   480
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "....ABERTURA:"
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
         Left            =   240
         TabIndex        =   14
         Top             =   180
         Width           =   1320
      End
      Begin VB.Label Label5 
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
         Left            =   2640
         TabIndex        =   13
         Top             =   420
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário:"
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
         Left            =   240
         TabIndex        =   12
         Top             =   420
         Width           =   705
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHAMENTO"
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
         Left            =   7740
         TabIndex        =   7
         Top             =   1020
         Width           =   1275
      End
      Begin VB.Line Line2 
         X1              =   7200
         X2              =   9540
         Y1              =   960
         Y2              =   960
      End
   End
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   10828
      Left            =   0
      Top             =   0
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   19103
      Tipo            =   2
      Begin ReportX.ReportField rf2 
         Height          =   300
         Left            =   3720
         TabIndex        =   2
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
         TabIndex        =   3
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
         TabIndex        =   4
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
         TabIndex        =   5
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
         TabIndex        =   6
         Top             =   1440
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   688
         Caption         =   "FECHAMENTO DE CAIXA - RESUMIDO"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
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
      Begin ReportX.ReportField rfVendasTotal 
         Height          =   240
         Left            =   4500
         TabIndex        =   31
         Top             =   2040
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   423
         Formato         =   "##,##0.00"
         Caption         =   ""
         Alignment       =   1
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
      End
      Begin ReportX.ReportField rfVendasDinheiro 
         Height          =   210
         Left            =   4500
         TabIndex        =   32
         Top             =   2280
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfVendasPix 
         Height          =   210
         Left            =   4500
         TabIndex        =   33
         Top             =   2460
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfVendasCartao 
         Height          =   210
         Left            =   4500
         TabIndex        =   34
         Top             =   2640
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfVendasDeposito 
         Height          =   210
         Left            =   4500
         TabIndex        =   35
         Top             =   2820
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfVendasTransferencia 
         Height          =   210
         Left            =   4500
         TabIndex        =   36
         Top             =   3000
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfVendasCheque 
         Height          =   210
         Left            =   4500
         TabIndex        =   37
         Top             =   3180
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfVendasFinanceira 
         Height          =   210
         Left            =   4500
         TabIndex        =   38
         Top             =   3360
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfParcelasTotal 
         Height          =   240
         Left            =   4500
         TabIndex        =   46
         Top             =   3660
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   423
         Formato         =   "##,##0.00"
         Caption         =   ""
         Alignment       =   1
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
      End
      Begin ReportX.ReportField rfParcelasDinheiro 
         Height          =   210
         Left            =   4500
         TabIndex        =   47
         Top             =   3900
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfParcelasPix 
         Height          =   210
         Left            =   4500
         TabIndex        =   48
         Top             =   4080
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfParcelasCartao 
         Height          =   210
         Left            =   4500
         TabIndex        =   49
         Top             =   4260
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfParcelasDeposito 
         Height          =   210
         Left            =   4500
         TabIndex        =   50
         Top             =   4440
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfParcelasTransferencia 
         Height          =   210
         Left            =   4500
         TabIndex        =   51
         Top             =   4620
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfParcelasCheque 
         Height          =   210
         Left            =   4500
         TabIndex        =   52
         Top             =   4800
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfParcelasFinanceira 
         Height          =   210
         Left            =   4500
         TabIndex        =   53
         Top             =   4980
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfHaveresTotal 
         Height          =   240
         Left            =   4500
         TabIndex        =   62
         Top             =   5280
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   423
         Formato         =   "##,##0.00"
         Caption         =   ""
         Alignment       =   1
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
      End
      Begin ReportX.ReportField rfHaveresDinheiro 
         Height          =   210
         Left            =   4500
         TabIndex        =   63
         Top             =   5520
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfHaveresPix 
         Height          =   210
         Left            =   4500
         TabIndex        =   64
         Top             =   5700
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfHaveresCartao 
         Height          =   210
         Left            =   4500
         TabIndex        =   65
         Top             =   5880
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfHaveresDeposito 
         Height          =   210
         Left            =   4500
         TabIndex        =   66
         Top             =   6060
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfHaveresTransferencia 
         Height          =   210
         Left            =   4500
         TabIndex        =   67
         Top             =   6240
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfHaveresCheque 
         Height          =   210
         Left            =   4500
         TabIndex        =   68
         Top             =   6420
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfHaveresFinanceira 
         Height          =   210
         Left            =   4500
         TabIndex        =   69
         Top             =   6600
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfSuprimentoTotal 
         Height          =   240
         Left            =   4500
         TabIndex        =   78
         Top             =   6960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   423
         Formato         =   "##,##0.00"
         Caption         =   ""
         Alignment       =   1
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
      End
      Begin ReportX.ReportField rfSuprimentoDinheiro 
         Height          =   210
         Left            =   4500
         TabIndex        =   79
         Top             =   7200
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfSangriaTotal 
         Height          =   240
         Left            =   4500
         TabIndex        =   82
         Top             =   7560
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   423
         Formato         =   "##,##0.00"
         Caption         =   ""
         Alignment       =   1
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
      End
      Begin ReportX.ReportField rfSangriaDinheiro 
         Height          =   210
         Left            =   4500
         TabIndex        =   83
         Top             =   7800
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfSaldoGeral 
         Height          =   240
         Left            =   3840
         TabIndex        =   86
         Top             =   8160
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   423
         Formato         =   "##,##0.00"
         Caption         =   ""
         Alignment       =   1
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
      End
      Begin ReportX.ReportField rfSaldoFisico 
         Height          =   240
         Left            =   3840
         TabIndex        =   88
         Top             =   8400
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   423
         Formato         =   "##,##0.00"
         Caption         =   ""
         Alignment       =   1
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
      End
      Begin ReportX.ReportField rfVendasQtde 
         Height          =   240
         Left            =   60
         TabIndex        =   91
         Top             =   2040
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   423
         Caption         =   ""
         Alignment       =   1
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
      End
      Begin ReportX.ReportField rfVendasDinheiroQtde 
         Height          =   210
         Left            =   60
         TabIndex        =   92
         Top             =   2280
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfVendasPixQtde 
         Height          =   210
         Left            =   60
         TabIndex        =   93
         Top             =   2460
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfVendasCartaoQtde 
         Height          =   210
         Left            =   60
         TabIndex        =   94
         Top             =   2640
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfVendasDepositoQtde 
         Height          =   210
         Left            =   60
         TabIndex        =   95
         Top             =   2820
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfVendasTransferenciaQtde 
         Height          =   210
         Left            =   60
         TabIndex        =   96
         Top             =   3000
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfVendasChequeQtde 
         Height          =   210
         Left            =   60
         TabIndex        =   97
         Top             =   3180
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfVendasFinanceiraQtde 
         Height          =   210
         Left            =   60
         TabIndex        =   98
         Top             =   3360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfParcelasTotalQtde 
         Height          =   240
         Left            =   60
         TabIndex        =   99
         Top             =   3660
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   423
         Caption         =   ""
         Alignment       =   1
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
      End
      Begin ReportX.ReportField rfParcelasDinheiroQtde 
         Height          =   210
         Left            =   60
         TabIndex        =   100
         Top             =   3900
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfParcelasPixQtde 
         Height          =   210
         Left            =   60
         TabIndex        =   101
         Top             =   4080
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfParcelasCartaoQtde 
         Height          =   210
         Left            =   60
         TabIndex        =   102
         Top             =   4260
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfParcelasDepositoQtde 
         Height          =   210
         Left            =   60
         TabIndex        =   103
         Top             =   4440
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfParcelasTransferenciaQtde 
         Height          =   210
         Left            =   60
         TabIndex        =   104
         Top             =   4620
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfParcelasChequeQtde 
         Height          =   210
         Left            =   60
         TabIndex        =   105
         Top             =   4800
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfParcelasFinanceiraQtde 
         Height          =   210
         Left            =   60
         TabIndex        =   106
         Top             =   4980
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfHaveresTotalQtde 
         Height          =   240
         Left            =   60
         TabIndex        =   107
         Top             =   5280
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   423
         Caption         =   ""
         Alignment       =   1
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
      End
      Begin ReportX.ReportField rfHaveresDinheiroQtde 
         Height          =   210
         Left            =   60
         TabIndex        =   108
         Top             =   5520
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfHaveresPixQtde 
         Height          =   210
         Left            =   60
         TabIndex        =   109
         Top             =   5700
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfHaveresCartaoQtde 
         Height          =   210
         Left            =   60
         TabIndex        =   110
         Top             =   5880
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfHaveresDepositoQtde 
         Height          =   210
         Left            =   60
         TabIndex        =   111
         Top             =   6060
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfHaveresTransferenciaQtde 
         Height          =   210
         Left            =   60
         TabIndex        =   112
         Top             =   6240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfHaveresChequeQtde 
         Height          =   210
         Left            =   60
         TabIndex        =   113
         Top             =   6420
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfHaveresFinanceiraQtde 
         Height          =   210
         Left            =   60
         TabIndex        =   114
         Top             =   6600
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfSuprimentoTotalQtde 
         Height          =   240
         Left            =   60
         TabIndex        =   115
         Top             =   6960
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   423
         Caption         =   ""
         Alignment       =   1
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
      End
      Begin ReportX.ReportField rfSuprimentoDinheiroQtde 
         Height          =   210
         Left            =   60
         TabIndex        =   116
         Top             =   7200
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfSangriaTotalQtde 
         Height          =   240
         Left            =   60
         TabIndex        =   117
         Top             =   7560
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   423
         Caption         =   ""
         Alignment       =   1
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
      End
      Begin ReportX.ReportField rfSangriaDinheiroQtde 
         Height          =   210
         Left            =   60
         TabIndex        =   118
         Top             =   7800
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfResumoDinheiro 
         Height          =   210
         Left            =   4560
         TabIndex        =   120
         Top             =   9060
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfResumoPix 
         Height          =   210
         Left            =   4560
         TabIndex        =   121
         Top             =   9240
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfResumoCartao 
         Height          =   210
         Left            =   4560
         TabIndex        =   122
         Top             =   9420
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfResumoDeposito 
         Height          =   210
         Left            =   4560
         TabIndex        =   123
         Top             =   9600
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfResumoTransferencia 
         Height          =   210
         Left            =   4560
         TabIndex        =   124
         Top             =   9780
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfResumoCheque 
         Height          =   210
         Left            =   4560
         TabIndex        =   125
         Top             =   9960
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfResumoFinanceira 
         Height          =   210
         Left            =   4560
         TabIndex        =   126
         Top             =   10140
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfResumoDinheiroQtde 
         Height          =   210
         Left            =   120
         TabIndex        =   127
         Top             =   9060
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfResumoPixQtde 
         Height          =   210
         Left            =   120
         TabIndex        =   128
         Top             =   9240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfResumoCartaoQtde 
         Height          =   210
         Left            =   120
         TabIndex        =   129
         Top             =   9420
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfResumoDepositoQtde 
         Height          =   210
         Left            =   120
         TabIndex        =   130
         Top             =   9600
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfResumoTransferenciaQtde 
         Height          =   210
         Left            =   120
         TabIndex        =   131
         Top             =   9780
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfResumoChequeQtde 
         Height          =   210
         Left            =   120
         TabIndex        =   132
         Top             =   9960
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfResumoFinanceiraQtde 
         Height          =   210
         Left            =   120
         TabIndex        =   133
         Top             =   10140
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfeExtraAPrazo 
         Height          =   210
         Left            =   10080
         TabIndex        =   141
         Top             =   9300
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfeExtraCanceladas 
         Height          =   210
         Left            =   10080
         TabIndex        =   142
         Top             =   9480
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfeExtraOrcamentos 
         Height          =   210
         Left            =   10080
         TabIndex        =   143
         Top             =   9660
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfeExtraConsignado 
         Height          =   210
         Left            =   10080
         TabIndex        =   144
         Top             =   9840
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfeExtraAPrazoQtde 
         Height          =   210
         Left            =   5640
         TabIndex        =   145
         Top             =   9300
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfeExtraCanceladasQtde 
         Height          =   210
         Left            =   5640
         TabIndex        =   146
         Top             =   9480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfeExtraOrcamentosQtde 
         Height          =   210
         Left            =   5640
         TabIndex        =   147
         Top             =   9660
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfeExtraConsignadoQtde 
         Height          =   210
         Left            =   5640
         TabIndex        =   148
         Top             =   9840
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfRetValor 
         Height          =   210
         Index           =   0
         Left            =   10380
         TabIndex        =   161
         Top             =   2340
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfHoraRet 
         Height          =   210
         Index           =   0
         Left            =   5880
         TabIndex        =   162
         Top             =   2340
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   370
         Formato         =   "hh:mm"
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
      Begin ReportX.ReportField rfDescRet 
         Height          =   210
         Index           =   0
         Left            =   7020
         TabIndex        =   163
         Top             =   2340
         Width           =   3315
         _ExtentX        =   5847
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
      End
      Begin ReportX.ReportField rfHoraRet 
         Height          =   210
         Index           =   1
         Left            =   5880
         TabIndex        =   164
         Top             =   2520
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   370
         Formato         =   "hh:mm"
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
      Begin ReportX.ReportField rfHoraRet 
         Height          =   210
         Index           =   2
         Left            =   5880
         TabIndex        =   165
         Top             =   2700
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   370
         Formato         =   "hh:mm"
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
      Begin ReportX.ReportField rfHoraRet 
         Height          =   210
         Index           =   3
         Left            =   5880
         TabIndex        =   166
         Top             =   2880
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   370
         Formato         =   "hh:mm"
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
      Begin ReportX.ReportField rfDescRet 
         Height          =   210
         Index           =   1
         Left            =   7020
         TabIndex        =   167
         Top             =   2520
         Width           =   3315
         _ExtentX        =   5847
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
      End
      Begin ReportX.ReportField rfDescRet 
         Height          =   210
         Index           =   2
         Left            =   7020
         TabIndex        =   168
         Top             =   2700
         Width           =   3315
         _ExtentX        =   5847
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
      End
      Begin ReportX.ReportField rfDescRet 
         Height          =   210
         Index           =   3
         Left            =   7020
         TabIndex        =   169
         Top             =   2880
         Width           =   3315
         _ExtentX        =   5847
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
      End
      Begin ReportX.ReportField rfRetValor 
         Height          =   210
         Index           =   1
         Left            =   10380
         TabIndex        =   170
         Top             =   2520
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfRetValor 
         Height          =   210
         Index           =   2
         Left            =   10380
         TabIndex        =   171
         Top             =   2700
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfRetValor 
         Height          =   210
         Index           =   3
         Left            =   10380
         TabIndex        =   172
         Top             =   2880
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfFuncRet 
         Height          =   210
         Index           =   0
         Left            =   6480
         TabIndex        =   177
         Top             =   2340
         Width           =   495
         _ExtentX        =   873
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
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField rfFuncRet 
         Height          =   210
         Index           =   1
         Left            =   6480
         TabIndex        =   178
         Top             =   2520
         Width           =   495
         _ExtentX        =   873
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
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField rfFuncRet 
         Height          =   210
         Index           =   2
         Left            =   6480
         TabIndex        =   179
         Top             =   2700
         Width           =   495
         _ExtentX        =   873
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
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField rfFuncRet 
         Height          =   210
         Index           =   3
         Left            =   6480
         TabIndex        =   180
         Top             =   2880
         Width           =   495
         _ExtentX        =   873
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
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField rfExtraTroco 
         Height          =   210
         Left            =   10080
         TabIndex        =   184
         Top             =   9120
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfExtraTrocoQtde 
         Height          =   210
         Left            =   5640
         TabIndex        =   185
         Top             =   9120
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfeExtraAluguel 
         Height          =   210
         Left            =   10080
         TabIndex        =   187
         Top             =   10020
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfeExtraOS 
         Height          =   210
         Left            =   10080
         TabIndex        =   188
         Top             =   10200
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfeExtraAluguelQtde 
         Height          =   210
         Left            =   5640
         TabIndex        =   189
         Top             =   10020
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin ReportX.ReportField rfeExtraOSQtde 
         Height          =   210
         Left            =   5640
         TabIndex        =   190
         Top             =   10200
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
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
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ORDEM DE SERVIÇO................................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6300
         TabIndex        =   192
         Top             =   10200
         Width           =   3705
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ALUGUEL..................................................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6300
         TabIndex        =   191
         Top             =   10020
         Width           =   3735
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TROCO......................................................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6300
         TabIndex        =   186
         Top             =   9120
         Width           =   3735
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FUNC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   660
         TabIndex        =   183
         Top             =   10620
         Width           =   420
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FUNC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6480
         TabIndex        =   181
         Top             =   2100
         Width           =   420
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VALOR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   10560
         TabIndex        =   176
         Top             =   2100
         Width           =   570
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIÇÃO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7020
         TabIndex        =   175
         Top             =   2100
         Width           =   930
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HORA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5880
         TabIndex        =   174
         Top             =   2100
         Width           =   450
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RETIRADAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8040
         TabIndex        =   173
         Top             =   1920
         Width           =   1110
      End
      Begin VB.Line Line4 
         BorderStyle     =   3  'Dot
         X1              =   11040
         X2              =   60
         Y1              =   10440
         Y2              =   10440
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VALOR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4740
         TabIndex        =   160
         Top             =   10620
         Width           =   570
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIÇÃO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1140
         TabIndex        =   159
         Top             =   10620
         Width           =   930
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HORA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   158
         Top             =   10620
         Width           =   450
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SANGRIAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2220
         TabIndex        =   157
         Top             =   10440
         Width           =   1020
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PROJEÇÃO/EXTRA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7440
         TabIndex        =   153
         Top             =   8820
         Width           =   2055
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENDAS À PRAZO...................................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6300
         TabIndex        =   152
         Top             =   9300
         Width           =   3735
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENDAS CANCELADAS............................................"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6300
         TabIndex        =   151
         Top             =   9480
         Width           =   3735
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ORÇAMENTOS..........................................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6300
         TabIndex        =   150
         Top             =   9660
         Width           =   3735
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONSIGNADO............................................................"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6300
         TabIndex        =   149
         Top             =   9840
         Width           =   3735
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         X1              =   5520
         X2              =   5520
         Y1              =   1800
         Y2              =   10440
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FINANCEIRA.............................................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   780
         TabIndex        =   140
         Top             =   10140
         Width           =   3690
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CHEQUE....................................................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   780
         TabIndex        =   139
         Top             =   9960
         Width           =   3720
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TRANSFERÊNCIA......................................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   780
         TabIndex        =   138
         Top             =   9780
         Width           =   3735
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEPOSITO.................................................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   780
         TabIndex        =   137
         Top             =   9600
         Width           =   3720
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CARTÃO...................................................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   780
         TabIndex        =   136
         Top             =   9420
         Width           =   3720
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PIX.............................................................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   780
         TabIndex        =   135
         Top             =   9240
         Width           =   3735
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DINHEIRO...................................................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   780
         TabIndex        =   134
         Top             =   9060
         Width           =   3750
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DETRALHAMENTO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1860
         TabIndex        =   119
         Top             =   8820
         Width           =   2040
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL EM VENDAS..............................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   90
         Top             =   2040
         Width           =   3705
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SALDO FÍSICO.................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   60
         TabIndex        =   89
         Top             =   8400
         Width           =   3735
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         X1              =   10980
         X2              =   0
         Y1              =   8760
         Y2              =   8760
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SALDO GERAL................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   60
         TabIndex        =   87
         Top             =   8160
         Width           =   3720
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL EM SANGRIA.............................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   85
         Top             =   7560
         Width           =   3735
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SANGRIA EM DINHEIRO................................................"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   84
         Top             =   7800
         Width           =   3855
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL EM SUPRIMENTOS...................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   81
         Top             =   6960
         Width           =   3675
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUPRIMENTO EM DINHEIRO................................................"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   80
         Top             =   7200
         Width           =   4110
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HAVER EM FINANCEIRA............................................"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   77
         Top             =   6600
         Width           =   3720
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HAVER EM CHEQUE.................................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   76
         Top             =   6420
         Width           =   3705
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HAVER EM TRANSFERÊNCIA..................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   75
         Top             =   6240
         Width           =   3675
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HAVER EM DEPOSITO..............................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   74
         Top             =   6060
         Width           =   3705
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HAVER EM CARTÃO................................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   73
         Top             =   5880
         Width           =   3705
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HAVER EM PIX.........................................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   72
         Top             =   5700
         Width           =   3675
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HAVER EM DINHEIRO................................................"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   71
         Top             =   5520
         Width           =   3690
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL EM HAVERES............................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   70
         Top             =   5280
         Width           =   3705
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PARCELAS EM FINANCEIRA....................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   61
         Top             =   4980
         Width           =   3690
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PARCELAS EM CHEQUE..........................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   60
         Top             =   4800
         Width           =   3675
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PARCELAS EM TRANSFERÊNCIA............................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   59
         Top             =   4620
         Width           =   3690
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PARCELAS EM DEPOSITO............................................"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   58
         Top             =   4440
         Width           =   3855
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PARCELAS EM CARTÃO.........................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   57
         Top             =   4260
         Width           =   3675
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PARCELAS EM PIX..................................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   56
         Top             =   4080
         Width           =   3645
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PARCELAS EM DINHEIRO.........................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   55
         Top             =   3900
         Width           =   3705
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL EM PARCELAS..........................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   54
         Top             =   3660
         Width           =   3720
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENDAS EM DINHEIRO.............................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   45
         Top             =   2280
         Width           =   3705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENDAS EM PIX........................................................"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   44
         Top             =   2460
         Width           =   3690
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENDAS EM CARTÃO..............................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   43
         Top             =   2640
         Width           =   3720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENDAS EM DEPOSITO............................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   42
         Top             =   2820
         Width           =   3720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENDAS EM TRANSFERÊNCIA................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   41
         Top             =   3000
         Width           =   3690
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENDAS EM CHEQUE..............................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   40
         Top             =   3180
         Width           =   3675
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENDAS EM FINANCEIRA........................................."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   39
         Top             =   3360
         Width           =   3690
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
   Begin ReportX.ReportSection ReportSection4 
      Align           =   1  'Align Top
      Height          =   165
      Left            =   0
      Top             =   10830
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   291
      Begin ReportX.ReportField rfValor 
         Height          =   210
         Left            =   4560
         TabIndex        =   154
         Top             =   0
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   370
         Campo           =   "vSValor"
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
      Begin ReportX.ReportField rfHora 
         Height          =   210
         Left            =   60
         TabIndex        =   155
         Top             =   0
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   370
         Campo           =   "vSHora"
         Formato         =   "hh:mm"
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
      Begin ReportX.ReportField rfDescricao 
         Height          =   210
         Left            =   1140
         TabIndex        =   156
         Top             =   0
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   370
         Campo           =   "vSDescricao"
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
      Begin ReportX.ReportField rfFunc 
         Height          =   210
         Left            =   660
         TabIndex        =   182
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   370
         Campo           =   "vSFunc"
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
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VENDAS EM DINHEIRO.............................................."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   3705
   End
End
Attribute VB_Name = "REL_Caixa_Fech_Resumido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rRet As ADODB.Recordset
Public Sub MostrarRetiradas(vCaixa As String, vCodCaixa As String)
Dim i As Integer
Dim Cont As Long

Set rRet = dbData.OpenRecordset("SELECT DESCRICAO, COD_FUNCIONARIO, HORA, VALOR FROM caixa_retirada WHERE (CAIXA = '" & vCaixa & "') and (CODCAIXA = '" & vCodCaixa & "') ORDER BY CODIGO;")

For i = 0 To 3
    rfHoraRet(i).Caption = ""
    rfFuncRet(i).Caption = ""
    rfDescRet(i).Caption = ""
    rfRetValor(i).Caption = ""
Next

Cont = 0

Do While Not rRet.EOF
   rfHoraRet(Cont).Caption = Format(rRet("HORA"), "hh:mm")
   rfFuncRet(Cont).Caption = Format(rRet("COD_FUNCIONARIO"), "00")
   rfDescRet(Cont).Caption = rRet("DESCRICAO")
   rfRetValor(Cont).Caption = Format(rRet("valor"), ocMONEY)
   Cont = Cont + 1
   rRet.MoveNext
Loop
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



Private Sub rfT1_Change()
End Sub


