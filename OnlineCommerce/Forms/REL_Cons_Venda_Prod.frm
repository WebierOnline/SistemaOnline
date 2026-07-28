VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.Ocx"
Begin VB.Form REL_Cons_Venda_Prod 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impressão de Vendas por Produto e Serviço"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13875
   Icon            =   "REL_Cons_Venda_Prod.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   80.433
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   244.74
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ReportX.ReportSection ReportSection3 
      Align           =   1  'Align Top
      Height          =   1515
      Left            =   0
      Top             =   2550
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   2672
      Tipo            =   7
      Begin ReportX.ReportField rfCons2 
         Height          =   225
         Left            =   1140
         TabIndex        =   18
         Top             =   660
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   397
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
      End
      Begin ReportX.ReportField ReportField6 
         Height          =   225
         Left            =   9600
         TabIndex        =   19
         Top             =   1140
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   397
         Campo           =   "= Pág. [Pagina] de [Paginas]"
         Caption         =   "ReportField1"
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rfCons1 
         Height          =   225
         Left            =   1860
         TabIndex        =   20
         Top             =   120
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   397
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
      End
      Begin ReportX.ReportField ReportField10 
         Height          =   240
         Left            =   120
         TabIndex        =   21
         Top             =   420
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   423
         Caption         =   "CRITÉRIO:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rfCons3 
         Height          =   225
         Left            =   1140
         TabIndex        =   22
         Top             =   420
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   397
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
      End
      Begin ReportX.ReportField ReportField9 
         Height          =   240
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   423
         Caption         =   "TIPO DE CONSULTA:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField dfQuant 
         Height          =   255
         Left            =   9795
         TabIndex        =   34
         Top             =   120
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
      End
      Begin ReportX.ReportField dfSubtotal 
         Height          =   255
         Left            =   9795
         TabIndex        =   35
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         Formato         =   "##,##0.00"
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
      End
      Begin ReportX.ReportField dfTotal 
         Height          =   255
         Left            =   9795
         TabIndex        =   36
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         Formato         =   "##,##0.00"
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
      End
      Begin ReportX.ReportField dfDesconto 
         Height          =   255
         Left            =   9795
         TabIndex        =   37
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         Formato         =   "##,##0.00"
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Quant:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   9210
         TabIndex        =   41
         Top             =   120
         Width           =   555
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   9015
         TabIndex        =   40
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   8910
         TabIndex        =   39
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   9300
         TabIndex        =   38
         Top             =   840
         Width           =   465
      End
      Begin VB.Line Line2 
         BorderStyle     =   3  'Dot
         X1              =   120
         X2              =   11280
         Y1              =   0
         Y2              =   0
      End
   End
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   4080
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Titulo          =   ""
      NomeImpressora  =   "IMPRESSORA1"
      Registrado      =   0   'False
   End
   Begin ReportX.ReportSection ReportSection2 
      Align           =   1  'Align Top
      Height          =   255
      Left            =   0
      Top             =   2295
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   450
      Begin ReportX.ReportField ReportField7 
         Height          =   195
         Left            =   10800
         TabIndex        =   8
         Top             =   0
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   344
         Campo           =   "vartotal"
         Formato         =   "##,##0.00"
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
      Begin ReportX.ReportField ReportField3 
         Height          =   195
         Left            =   1560
         TabIndex        =   5
         Top             =   0
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   344
         Campo           =   "varData"
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
      Begin ReportX.ReportField ReportField4 
         Height          =   195
         Left            =   9360
         TabIndex        =   6
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   344
         Campo           =   "varsubtotal"
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
         Left            =   10140
         TabIndex        =   7
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   344
         Campo           =   "vardesc"
         Formato         =   "##,##0.000"
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
      Begin ReportX.ReportField ReportField1 
         Height          =   195
         Left            =   8760
         TabIndex        =   23
         Top             =   0
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   344
         Campo           =   "varquant"
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
      Begin ReportX.ReportField ReportField2 
         Height          =   195
         Left            =   2520
         TabIndex        =   24
         Top             =   0
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   344
         Campo           =   "varNome"
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
      Begin ReportX.ReportField ReportField8 
         Height          =   195
         Left            =   7740
         TabIndex        =   25
         Top             =   0
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   344
         Campo           =   "varvalor"
         Formato         =   "##,##0.00"
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
      Begin ReportX.ReportField ReportField11 
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   0
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   344
         Campo           =   "varcodped"
         Formato         =   "000000"
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
      Begin ReportX.ReportField ReportField46 
         Height          =   195
         Left            =   840
         TabIndex        =   32
         Top             =   0
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   344
         Campo           =   "var_CodOS"
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
   End
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   2295
      Left            =   0
      Top             =   0
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   4048
      Tipo            =   2
      Begin ReportX.ReportField rf2 
         Height          =   300
         Left            =   3720
         TabIndex        =   14
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
         TabIndex        =   15
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
         TabIndex        =   16
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
         TabIndex        =   17
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DATA"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1560
         TabIndex        =   31
         Top             =   2040
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QTDE"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8880
         TabIndex        =   28
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIÇÃO"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2520
         TabIndex        =   27
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VALOR"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8160
         TabIndex        =   26
         Top             =   2040
         Width           =   495
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
         Width           =   11655
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   11280
         TabIndex        =   9
         Top             =   2040
         Width           =   450
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         X1              =   60
         X2              =   11760
         Y1              =   1980
         Y2              =   1980
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESC."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   10380
         TabIndex        =   4
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUBTOTAL"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9420
         TabIndex        =   3
         Top             =   2040
         Width           =   750
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PEDIDO"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   2
         Top             =   2040
         Width           =   540
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OS"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   840
         TabIndex        =   33
         Top             =   2040
         Width           =   195
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RELATÓRIO POR PRODUTO E SERVIÇO"
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
         Left            =   2835
         TabIndex        =   1
         Top             =   1500
         Width           =   5520
      End
   End
   Begin ReportX.ReportField ReportField42 
      Height          =   300
      Left            =   1740
      TabIndex        =   10
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
      TabIndex        =   11
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
      TabIndex        =   12
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
      TabIndex        =   13
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
   Begin VB.Image Image2 
      Height          =   1650
      Left            =   0
      Picture         =   "REL_Cons_Venda_Prod.frx":030A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1725
   End
End
Attribute VB_Name = "REL_Cons_Venda_Prod"
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
