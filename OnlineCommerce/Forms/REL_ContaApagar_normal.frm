VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.ocx"
Begin VB.Form REL_ContaApagar_Normal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impressão de Contas à Pagar"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12570
   Icon            =   "REL_ContaApagar_normal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   83.344
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   221.721
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ReportX.ReportSection ReportSection3 
      Align           =   1  'Align Top
      Height          =   1515
      Left            =   0
      Top             =   2610
      Width           =   12570
      _ExtentX        =   22172
      _ExtentY        =   2672
      Tipo            =   7
      Begin ReportX.ReportField dfDesc 
         Height          =   240
         Left            =   180
         TabIndex        =   13
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   423
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
      End
      Begin ReportX.ReportField dfCriterio 
         Height          =   240
         Left            =   180
         TabIndex        =   10
         Top             =   720
         Visible         =   0   'False
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   423
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
      End
      Begin ReportX.ReportField rftotal 
         Height          =   330
         Left            =   9720
         TabIndex        =   30
         Top             =   1140
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         Caption         =   "000"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
      End
      Begin ReportX.ReportField rfhaveres 
         Height          =   330
         Left            =   9720
         TabIndex        =   31
         Top             =   780
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
      End
      Begin ReportX.ReportField rfsubtotal 
         Height          =   330
         Left            =   9720
         TabIndex        =   32
         Top             =   420
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
      End
      Begin ReportX.ReportField rfQuant 
         Height          =   330
         Left            =   9720
         TabIndex        =   41
         Top             =   60
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "QUANT.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   8700
         TabIndex        =   42
         Top             =   60
         Width           =   945
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CRITÉRIO"
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
         TabIndex        =   40
         Top             =   540
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO DE CONSULTA:"
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
         TabIndex        =   39
         Top             =   60
         Width           =   1875
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SUB-TOTAL:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   8220
         TabIndex        =   36
         Top             =   420
         Width           =   1395
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "HAVER(ES):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   8220
         TabIndex        =   35
         Top             =   780
         Width           =   1395
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   8820
         TabIndex        =   34
         Top             =   1140
         Width           =   810
      End
      Begin VB.Line Line3 
         X1              =   3120
         X2              =   7860
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Label Label3 
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
         Left            =   4860
         TabIndex        =   33
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Line Line2 
         BorderStyle     =   3  'Dot
         X1              =   60
         X2              =   11400
         Y1              =   0
         Y2              =   0
      End
   End
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   4140
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Titulo          =   ""
      NomeImpressora  =   "\\BALCAO01\IMPRESSORA1"
      Registrado      =   0   'False
   End
   Begin ReportX.ReportSection ReportSection2 
      Align           =   1  'Align Top
      Height          =   255
      Left            =   0
      Top             =   2355
      Width           =   12570
      _ExtentX        =   22172
      _ExtentY        =   450
      Begin ReportX.ReportField ReportField7 
         Height          =   195
         Left            =   10740
         TabIndex        =   11
         Top             =   0
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   344
         Campo           =   "STATUS"
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
      Begin ReportX.ReportField ReportField2 
         Height          =   195
         Left            =   960
         TabIndex        =   6
         Top             =   0
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   344
         Campo           =   "DESCRICAO"
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
      Begin ReportX.ReportField ReportField3 
         Height          =   195
         Left            =   4080
         TabIndex        =   7
         Top             =   0
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   344
         Campo           =   "FAVORECIDO"
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
      Begin ReportX.ReportField ReportField4 
         Height          =   195
         Left            =   7380
         TabIndex        =   8
         Top             =   0
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   344
         Campo           =   "VALOR_UND"
         Formato         =   "##,##0.00"
         Caption         =   ""
         TipoCampo       =   1
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
      Begin ReportX.ReportField ReportField5 
         Height          =   195
         Left            =   9960
         TabIndex        =   9
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   344
         Campo           =   "DATA_PGTO"
         Formato         =   "dd/mm/yy"
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
      Begin ReportX.ReportField ReportField16 
         Height          =   195
         Left            =   8220
         TabIndex        =   22
         Top             =   0
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   344
         Campo           =   "HAV"
         Formato         =   "##,##0.00"
         TipoCampo       =   1
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
      Begin ReportX.ReportField ReportField17 
         Height          =   195
         Left            =   9000
         TabIndex        =   23
         Top             =   0
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   344
         Campo           =   "var_Liquido"
         Formato         =   "##,##0.00"
         TipoCampo       =   1
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
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   344
         Campo           =   "VENCIMENTO"
         Formato         =   "dd/mm/yy"
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
   End
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   2355
      Left            =   0
      Top             =   0
      Width           =   12570
      _ExtentX        =   22172
      _ExtentY        =   4154
      Tipo            =   2
      Begin ReportX.ReportField rf2 
         Height          =   300
         Left            =   3720
         TabIndex        =   26
         Top             =   540
         Width           =   7425
         _ExtentX        =   13097
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
         TabIndex        =   27
         Top             =   60
         Width           =   7440
         _ExtentX        =   13123
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
         TabIndex        =   28
         Top             =   780
         Width           =   7425
         _ExtentX        =   13097
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
         TabIndex        =   29
         Top             =   1020
         Width           =   7425
         _ExtentX        =   13097
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
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENC."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   2160
         Width           =   570
      End
      Begin VB.Shape Shape1 
         Height          =   1335
         Left            =   60
         Shape           =   4  'Rounded Rectangle
         Top             =   60
         Width           =   11115
      End
      Begin VB.Image imgLogo 
         Height          =   1215
         Left            =   180
         Stretch         =   -1  'True
         Top             =   120
         Width           =   3315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9240
         TabIndex        =   25
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HAVER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8340
         TabIndex        =   24
         Top             =   2160
         Width           =   645
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STATUS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   10740
         TabIndex        =   12
         Top             =   2160
         Width           =   750
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         X1              =   120
         X2              =   11460
         Y1              =   2100
         Y2              =   2100
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PGTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   10140
         TabIndex        =   5
         Top             =   2160
         Width           =   525
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VALOR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7500
         TabIndex        =   4
         Top             =   2160
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FAVORECIDO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4080
         TabIndex        =   3
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIÇÃO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   960
         TabIndex        =   2
         Top             =   2160
         Width           =   1080
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RELATÓRIO DE CONTAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   60
         TabIndex        =   1
         Top             =   1620
         Width           =   11190
      End
   End
   Begin ReportX.ReportField ReportField42 
      Height          =   300
      Left            =   1740
      TabIndex        =   14
      Top             =   480
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   529
      Caption         =   "G. A. DE ANDRADE ME"
      Alignment       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AlturaLivre     =   -1  'True
      AlinhamentoVertical=   1
   End
   Begin ReportX.ReportField ReportField43 
      Height          =   510
      Left            =   1740
      TabIndex        =   15
      Top             =   15
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   900
      Caption         =   "Madereira Santa Maria"
      Alignment       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AlturaLivre     =   -1  'True
      AlinhamentoVertical=   1
   End
   Begin ReportX.ReportField ReportField44 
      Height          =   300
      Left            =   1740
      TabIndex        =   16
      Top             =   780
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   529
      Caption         =   "Rua Bertolínio Filho, 320 - Centro - Telefone: (89) 544 1919 - Uruçuí-PI"
      Alignment       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AlturaLivre     =   -1  'True
      AlinhamentoVertical=   1
   End
   Begin ReportX.ReportField ReportField45 
      Height          =   300
      Left            =   1740
      TabIndex        =   17
      Top             =   1080
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   529
      Caption         =   "CNPJ: 01.208.470/0001-08  -   IE 19.435.747-3  -  E-mail: gilsonalves@gurgueia.com.br"
      Alignment       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AlturaLivre     =   -1  'True
      AlinhamentoVertical=   1
   End
   Begin ReportX.ReportField ReportField8 
      Height          =   300
      Left            =   0
      TabIndex        =   18
      Top             =   465
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   529
      Caption         =   "FRANCISCO KENNED SOARES DE ARAÚJO REIS"
      Alignment       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AlturaLivre     =   -1  'True
      AlinhamentoVertical=   1
   End
   Begin ReportX.ReportField ReportField9 
      Height          =   510
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   900
      Caption         =   "K A R O L L A Y N E   G Á S"
      Alignment       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AlturaLivre     =   -1  'True
      AlinhamentoVertical=   1
   End
   Begin ReportX.ReportField ReportField10 
      Height          =   300
      Left            =   0
      TabIndex        =   20
      Top             =   765
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   529
      Caption         =   "Rua Coelho Rodrigues, 653 - Centro - Telefone: (89) 3544-1385 - Uruçuí-PI"
      Alignment       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AlturaLivre     =   -1  'True
      AlinhamentoVertical=   1
   End
   Begin ReportX.ReportField ReportField11 
      Height          =   300
      Left            =   0
      TabIndex        =   21
      Top             =   1065
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   529
      Caption         =   "CNPJ: 01.634.508/0001-04  -   IE 19.436.943-9"
      Alignment       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AlturaLivre     =   -1  'True
      AlinhamentoVertical=   1
   End
   Begin VB.Image Image2 
      Height          =   1650
      Left            =   0
      Picture         =   "REL_ContaApagar_normal.frx":030A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1725
   End
End
Attribute VB_Name = "REL_ContaApagar_Normal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
