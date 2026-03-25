VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.Ocx"
Begin VB.Form REL_Cashback_Semana 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CASHBACK DA SEMANA"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12570
   Icon            =   "REL_Cashback_Semana.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   73.29
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   221.721
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ReportX.ReportSection ReportSection3 
      Align           =   1  'Align Top
      Height          =   1095
      Left            =   0
      Top             =   2550
      Width           =   12570
      _ExtentX        =   22172
      _ExtentY        =   1931
      Tipo            =   7
      Begin ReportX.ReportField rfCons2 
         Height          =   240
         Left            =   60
         TabIndex        =   20
         Top             =   540
         Width           =   3255
         _ExtentX        =   5741
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
      Begin ReportX.ReportField ReportField6 
         Height          =   240
         Left            =   3660
         TabIndex        =   21
         Top             =   660
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   423
         Campo           =   "= [Pagina] de [Paginas]"
         Formula         =   -1  'True
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
      Begin ReportX.ReportField rfCons1 
         Height          =   240
         Left            =   60
         TabIndex        =   22
         Top             =   300
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
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
      Begin ReportX.ReportField dfQuant 
         Height          =   255
         Left            =   9600
         TabIndex        =   23
         Top             =   60
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
      Begin ReportX.ReportField ReportField10 
         Height          =   240
         Left            =   60
         TabIndex        =   24
         Top             =   60
         Visible         =   0   'False
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   423
         Caption         =   "CONSULTAS:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rfCons3 
         Height          =   240
         Left            =   60
         TabIndex        =   26
         Top             =   780
         Width           =   3255
         _ExtentX        =   5741
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
      Begin ReportX.ReportField rfForma 
         Height          =   240
         Left            =   3600
         TabIndex        =   27
         Top             =   360
         Visible         =   0   'False
         Width           =   3240
         _ExtentX        =   5715
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
      Begin ReportX.ReportField ReportField13 
         Height          =   240
         Left            =   3600
         TabIndex        =   28
         Top             =   120
         Visible         =   0   'False
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   423
         Caption         =   "TIPO:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField dfTotal 
         Height          =   255
         Left            =   9600
         TabIndex        =   35
         Top             =   300
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
      Begin VB.Label Label13 
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
         Left            =   9060
         TabIndex        =   36
         Top             =   300
         Width           =   465
      End
      Begin VB.Label Label16 
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
         Left            =   9000
         TabIndex        =   25
         Top             =   60
         Width           =   555
      End
      Begin VB.Line Line2 
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   11160
         Y1              =   0
         Y2              =   0
      End
   End
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   60
      TabIndex        =   0
      Top             =   3660
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
      Width           =   12570
      _ExtentX        =   22172
      _ExtentY        =   450
      Begin ReportX.ReportField ReportField1 
         Height          =   195
         Left            =   60
         TabIndex        =   7
         Top             =   0
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   344
         Campo           =   "varCod"
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
         Left            =   600
         TabIndex        =   8
         Top             =   0
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   344
         Campo           =   "varCodPed"
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
      Begin ReportX.ReportField ReportField3 
         Height          =   195
         Left            =   1380
         TabIndex        =   9
         Top             =   0
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   344
         Campo           =   "varCli"
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
         Left            =   8340
         TabIndex        =   10
         Top             =   0
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   344
         Campo           =   "varDataAbatido"
         Formato         =   "dd/mm/yy"
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
         Left            =   9240
         TabIndex        =   11
         Top             =   0
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   344
         Campo           =   "varCodPedAbatido"
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
      Begin ReportX.ReportField ReportField8 
         Height          =   195
         Left            =   7440
         TabIndex        =   29
         Top             =   0
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   344
         Campo           =   "varAbatido"
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
      Begin ReportX.ReportField ReportField11 
         Height          =   195
         Left            =   6660
         TabIndex        =   30
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   344
         Campo           =   "varVlrCash"
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
      Begin ReportX.ReportField ReportField12 
         Height          =   195
         Left            =   5640
         TabIndex        =   31
         Top             =   0
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   344
         Campo           =   "varValidade"
         Formato         =   "dd/mm/yy"
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
      Begin ReportX.ReportField ReportField7 
         Height          =   195
         Left            =   10200
         TabIndex        =   37
         Top             =   0
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   344
         Campo           =   "varInvalido"
         Caption         =   "varInvalido"
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
      Begin ReportX.ReportField ReportField9 
         Height          =   195
         Left            =   11160
         TabIndex        =   38
         Top             =   0
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   344
         Campo           =   "varCodFunc"
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
      Begin ReportX.ReportField ReportField14 
         Height          =   195
         Left            =   4440
         TabIndex        =   41
         Top             =   0
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   344
         Campo           =   "varCelular"
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
   End
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   2295
      Left            =   0
      Top             =   0
      Width           =   12570
      _ExtentX        =   22172
      _ExtentY        =   4048
      Tipo            =   2
      Begin ReportX.ReportField rf2 
         Height          =   300
         Left            =   3720
         TabIndex        =   16
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
         TabIndex        =   17
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
         TabIndex        =   18
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
         TabIndex        =   19
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
         Caption         =   "CELULAR"
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
         Left            =   4680
         TabIndex        =   42
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FUNC."
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
         Left            =   11160
         TabIndex        =   40
         Top             =   2040
         Width           =   570
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INVALIDO"
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
         Left            =   10200
         TabIndex        =   39
         Top             =   2040
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VALIDADE"
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
         Left            =   5640
         TabIndex        =   34
         Top             =   2040
         Width           =   930
      End
      Begin VB.Label Label2 
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
         Left            =   6720
         TabIndex        =   33
         Top             =   2040
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ABATIDO"
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
         Left            =   7440
         TabIndex        =   32
         Top             =   2040
         Width           =   825
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
         Width           =   11115
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         X1              =   60
         X2              =   11460
         Y1              =   1980
         Y2              =   1980
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CÓD.PED"
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
         TabIndex        =   6
         Top             =   2040
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DATA"
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
         Left            =   8640
         TabIndex        =   5
         Top             =   2040
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CLIENTE"
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
         Left            =   1380
         TabIndex        =   4
         Top             =   2040
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PEDIDO"
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
         Left            =   600
         TabIndex        =   3
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CÓD."
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
         Left            =   60
         TabIndex        =   2
         Top             =   2040
         Width           =   450
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RELATÓRIO DE CASHBACK"
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
         Left            =   3660
         TabIndex        =   1
         Top             =   1500
         Width           =   3870
      End
   End
   Begin ReportX.ReportField ReportField42 
      Height          =   300
      Left            =   1740
      TabIndex        =   12
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
      TabIndex        =   13
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
      TabIndex        =   14
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
      TabIndex        =   15
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
      Picture         =   "REL_Cashback_Semana.frx":030A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1725
   End
End
Attribute VB_Name = "REL_Cashback_Semana"
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

