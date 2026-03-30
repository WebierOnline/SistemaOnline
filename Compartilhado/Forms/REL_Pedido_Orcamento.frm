VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.Ocx"
Begin VB.Form REL_Pedido_Orcamento 
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13170
   LinkTopic       =   "Form1"
   ScaleHeight     =   148.167
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   232.304
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   60
      TabIndex        =   45
      Top             =   7860
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
      Height          =   7755
      Left            =   0
      Top             =   0
      Width           =   13170
      _ExtentX        =   23230
      _ExtentY        =   13679
      Ordem           =   1
      Begin ReportX.ReportField txtTotalProd 
         Height          =   225
         Index           =   0
         Left            =   10440
         TabIndex        =   126
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
         TabIndex        =   109
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
         TabIndex        =   84
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
      Begin ReportX.ReportField txtVendedor 
         Height          =   240
         Left            =   10200
         TabIndex        =   48
         Top             =   1740
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
         TabIndex        =   4
         Top             =   1440
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   503
         Caption         =   ""
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
         TabIndex        =   46
         Top             =   1740
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
      Begin ReportX.ReportField txtValidade 
         Height          =   225
         Left            =   9420
         TabIndex        =   47
         Top             =   1980
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   397
         Caption         =   "VALIDADE"
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
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   9
         Left            =   8620
         TabIndex        =   44
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
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   9
         Left            =   7860
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
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
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   8
         Left            =   7860
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   7
         Left            =   7860
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   6
         Left            =   7860
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   5
         Left            =   7860
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   4
         Left            =   7860
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   3
         Left            =   7860
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   2
         Left            =   7860
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   1
         Left            =   7860
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   0
         Left            =   7860
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   49
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
         TabIndex        =   50
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
         TabIndex        =   51
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
         TabIndex        =   52
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
      Begin ReportX.ReportField ReportField6 
         Height          =   255
         Left            =   9420
         TabIndex        =   53
         Top             =   1200
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   450
         Caption         =   "ORÇAMENTO"
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
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   10
         Left            =   1440
         TabIndex        =   54
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
         TabIndex        =   55
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
         TabIndex        =   56
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
         TabIndex        =   57
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
         TabIndex        =   58
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
         TabIndex        =   59
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
         TabIndex        =   60
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
         TabIndex        =   61
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
         TabIndex        =   62
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
         TabIndex        =   63
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
         TabIndex        =   64
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
         TabIndex        =   65
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
         TabIndex        =   66
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
         TabIndex        =   67
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
         TabIndex        =   68
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
         TabIndex        =   69
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
         TabIndex        =   71
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
         TabIndex        =   72
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
         TabIndex        =   73
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
         TabIndex        =   74
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
         TabIndex        =   75
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
         TabIndex        =   76
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
         TabIndex        =   77
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
         TabIndex        =   78
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
         TabIndex        =   79
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
         TabIndex        =   80
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
         TabIndex        =   81
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
         TabIndex        =   82
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
         TabIndex        =   83
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
         TabIndex        =   85
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
         TabIndex        =   86
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
         TabIndex        =   87
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
         TabIndex        =   88
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
         TabIndex        =   89
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
         TabIndex        =   90
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
         TabIndex        =   91
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
         TabIndex        =   92
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
         TabIndex        =   93
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
         TabIndex        =   94
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
         TabIndex        =   95
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
         TabIndex        =   96
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
         TabIndex        =   97
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
         TabIndex        =   98
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
         TabIndex        =   99
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
         TabIndex        =   100
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
         Left            =   9440
         TabIndex        =   101
         Top             =   6870
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   476
         Caption         =   "Subtotal:"
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
      Begin ReportX.ReportField txtSubtotal 
         Height          =   270
         Left            =   10320
         TabIndex        =   102
         Top             =   6870
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
         Left            =   9640
         TabIndex        =   103
         Top             =   7140
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   476
         Caption         =   "Desc.:"
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
      Begin ReportX.ReportField txtDesconto 
         Height          =   270
         Left            =   10320
         TabIndex        =   104
         Top             =   7140
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
         Left            =   9640
         TabIndex        =   105
         Top             =   7410
         Width           =   645
         _ExtentX        =   1138
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
      Begin ReportX.ReportField txtTotal 
         Height          =   270
         Left            =   10320
         TabIndex        =   106
         Top             =   7410
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
         TabIndex        =   107
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
         TabIndex        =   108
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
         TabIndex        =   110
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
         TabIndex        =   111
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
         TabIndex        =   112
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
         TabIndex        =   113
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
         TabIndex        =   114
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
         TabIndex        =   115
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
         TabIndex        =   116
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
         TabIndex        =   117
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
         TabIndex        =   118
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
         TabIndex        =   119
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
         TabIndex        =   120
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
         TabIndex        =   121
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
         TabIndex        =   122
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
         TabIndex        =   123
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
         TabIndex        =   124
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
         TabIndex        =   125
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
         TabIndex        =   127
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
         TabIndex        =   128
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
         TabIndex        =   129
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
         TabIndex        =   130
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
         TabIndex        =   131
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
         TabIndex        =   132
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
         TabIndex        =   133
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
         TabIndex        =   134
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
         TabIndex        =   135
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
         TabIndex        =   136
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
         TabIndex        =   137
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
         TabIndex        =   138
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
         TabIndex        =   139
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
         TabIndex        =   140
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
         TabIndex        =   141
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
         TabIndex        =   142
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
      Begin ReportX.ReportField txtDataValidade 
         Height          =   240
         Left            =   9420
         TabIndex        =   143
         Top             =   2220
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   423
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
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField16 
         Height          =   270
         Left            =   180
         TabIndex        =   144
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
         TabIndex        =   145
         Top             =   1965
         Width           =   1365
         _ExtentX        =   2408
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
         TabIndex        =   146
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
         TabIndex        =   147
         Top             =   1200
         Width           =   3645
         _ExtentX        =   6429
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
         TabIndex        =   148
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
         TabIndex        =   149
         Top             =   1455
         Width           =   3840
         _ExtentX        =   6773
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
         TabIndex        =   150
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
         TabIndex        =   151
         Top             =   2220
         Width           =   3120
         _ExtentX        =   5503
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
         Left            =   2460
         TabIndex        =   152
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
         Left            =   2760
         TabIndex        =   153
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
      Begin ReportX.ReportField ReportField9 
         Height          =   270
         Left            =   180
         TabIndex        =   154
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
         TabIndex        =   155
         Top             =   1710
         Width           =   3645
         _ExtentX        =   6429
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
      Begin ReportX.ReportField frTitParc 
         Height          =   255
         Left            =   4620
         TabIndex        =   156
         Top             =   1200
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   450
         Caption         =   "VEÍCULO"
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
      Begin ReportX.ReportField txtFabricante 
         Height          =   300
         Left            =   5820
         TabIndex        =   157
         Top             =   1500
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   529
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
      Begin ReportX.ReportField ReportField2 
         Height          =   300
         Left            =   4620
         TabIndex        =   158
         Top             =   1500
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         Caption         =   "Fabricante:"
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
      Begin ReportX.ReportField txtModelo 
         Height          =   300
         Left            =   8040
         TabIndex        =   159
         Top             =   1500
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
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
      Begin ReportX.ReportField ReportField10 
         Height          =   300
         Left            =   7020
         TabIndex        =   160
         Top             =   1500
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         Caption         =   "Modelo:"
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
      Begin ReportX.ReportField txtAno 
         Height          =   300
         Left            =   5460
         TabIndex        =   161
         Top             =   1830
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
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
         Height          =   300
         Left            =   4620
         TabIndex        =   162
         Top             =   1830
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   529
         Caption         =   "Ano:"
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
      Begin ReportX.ReportField txtPlaca 
         Height          =   300
         Left            =   7200
         TabIndex        =   163
         Top             =   1830
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   529
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
         Height          =   300
         Left            =   6660
         TabIndex        =   164
         Top             =   1830
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   529
         Caption         =   "Placa:"
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
      Begin ReportX.ReportField txtCor 
         Height          =   300
         Left            =   5100
         TabIndex        =   165
         Top             =   2160
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   529
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
         Height          =   300
         Left            =   4620
         TabIndex        =   166
         Top             =   2160
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   529
         Caption         =   "Cor:"
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
      Begin ReportX.ReportField txtKM 
         Height          =   300
         Left            =   7200
         TabIndex        =   167
         Top             =   2160
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   529
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
      Begin ReportX.ReportField ReportField18 
         Height          =   300
         Left            =   6660
         TabIndex        =   168
         Top             =   2160
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   529
         Caption         =   "KM:"
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
      Begin ReportX.ReportField ReportField11 
         Height          =   270
         Left            =   180
         TabIndex        =   171
         Top             =   6870
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
         TabIndex        =   172
         Top             =   6870
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
      Begin ReportX.ReportField ReportField19 
         Height          =   255
         Left            =   180
         TabIndex        =   173
         Top             =   7140
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
         TabIndex        =   174
         Top             =   7140
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
         TabIndex        =   175
         Top             =   7410
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
         TabIndex        =   176
         Top             =   7410
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
         TabIndex        =   177
         Top             =   6870
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
         TabIndex        =   178
         Top             =   7140
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
         TabIndex        =   179
         Top             =   7410
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
         Height          =   870
         Left            =   120
         Top             =   6840
         Width           =   2715
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*A assinatura deste orçamento acarretará na autorização da ordem de serviço."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   2880
         TabIndex        =   170
         Top             =   6980
         Width           =   5370
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line8 
         X1              =   5040
         X2              =   9060
         Y1              =   7500
         Y2              =   7500
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
         Left            =   6660
         TabIndex        =   169
         Top             =   7560
         Width           =   825
      End
      Begin VB.Line Line7 
         X1              =   10400
         X2              =   10400
         Y1              =   6840
         Y2              =   2520
      End
      Begin VB.Line Line3 
         X1              =   7860
         X2              =   7860
         Y1              =   6840
         Y2              =   2520
      End
      Begin VB.Line Line6 
         X1              =   6840
         X2              =   6840
         Y1              =   6840
         Y2              =   2520
      End
      Begin VB.Shape Shape6 
         BorderWidth     =   2
         Height          =   870
         Left            =   9360
         Top             =   6840
         Width           =   2115
      End
      Begin VB.Line Line2 
         X1              =   1400
         X2              =   1400
         Y1              =   6840
         Y2              =   2520
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   11400
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         Height          =   1395
         Left            =   9360
         Top             =   1140
         Width           =   2115
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*A validade desse orçamento está sujeito a disponibilidade no estoque."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   2880
         TabIndex        =   70
         Top             =   6840
         Width           =   5370
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line5 
         X1              =   9660
         X2              =   9660
         Y1              =   6840
         Y2              =   2520
      End
      Begin VB.Line Line4 
         X1              =   8580
         X2              =   8580
         Y1              =   6840
         Y2              =   2520
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         Height          =   4335
         Left            =   120
         Top             =   2520
         Width           =   11355
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         Height          =   1395
         Left            =   4560
         Top             =   1140
         Width           =   4815
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   1395
         Left            =   120
         Top             =   1140
         Width           =   4455
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
Attribute VB_Name = "REL_Pedido_Orcamento"
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
If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
     Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, PLACA, ANO, KM, COR FROM OS_Equipamento_Auto WHERE (cod_os = " & vCodOS & ");")
ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
     Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, EQUIPAMENTO FROM OS_Equipamento WHERE (cod_os = " & vCodOS & ");")
ElseIf vTipoOS = "Comunicação Visual" Then
     'Set rEquip = dbData.OpenRecordset("SELECT fabricante, MODELO, EQUIPAMENTO FROM OS_Equipamento WHERE (cod_os = " & vCodOS & ");")
End If

Set rOS = dbData.OpenRecordset("SELECT cod_pedido, cod_cliente, cod_funcionario, SUBTOTAL, TOTAL, ValorDescReal  FROM os WHERE (cod_os = " & vCodOS & ");")
Set rPd = dbData.OpenRecordset("SELECT cod_cliente, cod_funcionario, data_compra, total, valor_desc, tipo_desc, subtotal FROM pedidos WHERE (cod_pedido = " & Pedido & ");")
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
              sSQL = sSQL & "SELECT isnull(sum(quantidade),0) as vSomaQuantServ, isnull(sum(total),0) as vSomaValorServ " & _
              "FROM OS_Servicos_Auto WHERE (cod_os = " & vCodOS & ")"
        ElseIf vTipoOS = "Recapadora" Then
            sSQL = sSQL & "SELECT isnull(sum(quantidade),0) as vSomaQuantServ, isnull(sum(total),0) as vSomaValorServ " & _
            "FROM OS_servicos_recapadora WHERE (cod_os = " & vCodOS & ")"
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
              sSQL = sSQL & "SELECT isnull(sum(quantidade),0) as vSomaQuantServ, isnull(sum(total),0) as vSomaValorServ " & _
              "FROM OS_Servicos_Auto WHERE (cod_os = " & vCodOS & ")"
        ElseIf vTipoOS = "Comunicação Visual" Then
              sSQL = sSQL & "SELECT isnull(sum(quantidade),0) as vSomaQuantServ, isnull(sum(total),0) as vSomaValorServ " & _
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

'Set rPc = dbData.OpenRecordset("SELECT * FROM parcelas WHERE (cod_pedido = " & Pedido & ") ORDER BY numero;")
Set rFu = dbData.OpenRecordset("SELECT nome FROM funcionario WHERE (codigo = " & rOS("cod_funcionario") & ");")

If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Informática" Or vTipoOS = "Celular" Then
     txtNumero.Caption = "Nº " & Format(vCodOS, "000000")
ElseIf vTipoOS = "Comunicação Visual" Then
     txtNumero.Caption = "Nº " & Format(vCodOS, "000000")
ElseIf vTipoOS = "Recapadora" Then
     txtNumero.Caption = "Nº " & Format(vCodOS, "000000")
End If

If rOS("cod_cliente") = 0 Then Exit Sub

'DADOS DO CLIENTE
txtCliente.Caption = rCl("nome")
txtEnd.Caption = IIf(IsNull(rCl![Endereco]) = True, "", rCl![Endereco]) & ", " & IIf(IsNull(rCl!Numero) = True, "", rCl!Numero) & " - " & IIf(IsNull(rCl!bairro) = True, "", rCl!bairro)
txtRef2.Caption = IIf(IsNull(rCl!Ponto_de_referencia) = True, "", rCl!Ponto_de_referencia)
txtCidade.Caption = IIf(IsNull(rCl!Cidade) = True, "", rCl!Cidade) & "-" & IIf(IsNull(rCl!Estado) = True, "", rCl!Estado) & "   FONE: " & IIf(IsNull(rCl!Telefone1) = True, "", rCl!Telefone1)
txtCPF.Caption = IIf(IsNull(rCl!CPF) = True, "", rCl!CPF)
txtRG.Caption = IIf(IsNull(rCl!rg) = True, "", rCl!rg)

 'DADOS DO VEICULO
 If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then
     frTitParc.Caption = "VEÍCULO"
     txtFabricante.Caption = IIf(IsNull(rEquip!Fabricante) = True, "", rEquip!Fabricante)
     txtModelo.Caption = IIf(IsNull(rEquip!Modelo) = True, "", rEquip!Modelo)
     txtPlaca.Caption = IIf(IsNull(rEquip!Placa) = True, "", rEquip!Placa)
     txtAno.Caption = IIf(IsNull(rEquip!Ano) = True, "", rEquip!Ano)
     txtCor.Caption = IIf(IsNull(rEquip!Cor) = True, "", rEquip!Cor)
     txtKM.Caption = IIf(IsNull(rEquip!KM) = True, "", rEquip!KM)
 ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
     frTitParc.Caption = "EQUIPAMENTO"
     txtFabricante.Caption = IIf(IsNull(rEquip!Equipamento) = True, "", rEquip!Equipamento)
     txtModelo.Caption = IIf(IsNull(rEquip!Fabricante) = True, "", rEquip!Fabricante)
     txtAno.Caption = IIf(IsNull(rEquip!Modelo) = True, "", rEquip!Modelo)
     txtPlaca.Visible = False
     txtCor.Visible = False
     txtKM.Visible = False
     ReportField15.Visible = False
     ReportField14.Visible = False
     ReportField18.Visible = False
     ReportField2.Caption = "Equipamento:"
     ReportField10.Caption = "Fabricante:"
     ReportField12.Caption = "Modelo:"
     ReportField15.Caption = ""
     ReportField14.Caption = ""
     ReportField18.Caption = ""
 ElseIf vTipoOS = "Comunicação Visual" Then
     frTitParc.Caption = ""
     txtFabricante.Visible = False
     txtModelo.Visible = False
     txtAno.Visible = False
     txtPlaca.Visible = False
     txtCor.Visible = False
     txtKM.Visible = False
     ReportField15.Visible = False
     ReportField14.Visible = False
     ReportField18.Visible = False
     ReportField2.Caption = ""
     ReportField10.Caption = ""
     ReportField12.Caption = ""
     ReportField15.Caption = ""
     ReportField14.Caption = ""
     ReportField18.Caption = ""
 End If
 
'DADOS DO PEDIDO
 'txtData.Caption = String(1, " ") + Format(rPd("data_compra"), "dd/mm/yy")
txtData.Caption = String(1, " ") + Format(Date, "dd/mm/yy")
txtVendedor.Caption = rFu("nome")
txtDataValidade.Caption = Format(DateAdd("d", Val(5), Date), "dd/mm/yy")

'DADOS DO FINANCEIRO
 txtSubtotal.Caption = " " & Format(rOS("SUBTOTAL"), ocMONEY)
 txtDesconto.Caption = " " & Format(rOS("ValorDescReal"), ocMONEY)
 txtTotal.Caption = " " & Format(rOS("TOTAL"), ocMONEY)

'INSIRO OS ITENS
If Not rIt.EOF Then rIt.MoveLast
If Not rIt.BOF Then rIt.MoveFirst

Relatorio.NumeroRegistros = Round((rIt.RecordCount / 17) + 0.49)
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
   For i = 0 To 16
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
            If rIt("tipo_item") = "SERVIÇO" Then
                txtDesc(i).Caption = String(1, " ") + rIt("descricao") & " | " & rIt("var_TipoPneu") & " | " & rIt("var_serie") & " | " & rIt("var_fogo") & " | " & rIt("var_aro") & " | " & rIt("var_banda") & " | " & rIt("var_dote") & " | " & rIt("var_medida") & " | " & rIt("var_fabricante") & ""
            Else
                txtDesc(i).Caption = String(1, " ") + rIt("descricao")
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

