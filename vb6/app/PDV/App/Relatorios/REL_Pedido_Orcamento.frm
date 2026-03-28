VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.Ocx"
Begin VB.Form REL_Pedido_Orcamento 
   ClientHeight    =   8400
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   ScaleHeight     =   148.167
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   198.967
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   60
      TabIndex        =   46
      Top             =   7860
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Titulo          =   ""
      LarguraPapel    =   209
      AlturaPapel     =   146
      Registrado      =   0   'False
      Visualizar      =   0   'False
   End
   Begin ReportX.ReportSection ReportSection2 
      Align           =   1  'Align Top
      Height          =   8055
      Left            =   0
      Top             =   0
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   14208
      Ordem           =   1
      Begin ReportX.ReportField txtVendedor 
         Height          =   240
         Left            =   9060
         TabIndex        =   49
         Top             =   1800
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   423
         Caption         =   ""
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
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
         Left            =   9060
         TabIndex        =   4
         Top             =   1200
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   503
         Caption         =   "0000"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
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
         Left            =   9060
         TabIndex        =   47
         Top             =   1540
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   423
         Caption         =   ""
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
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
         Height          =   300
         Left            =   9060
         TabIndex        =   48
         Top             =   2055
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "VALIDADE"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
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
         Left            =   8310
         TabIndex        =   44
         Top             =   4905
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   6690
         TabIndex        =   43
         Top             =   4905
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   7575
         TabIndex        =   42
         Top             =   4905
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
         Left            =   180
         TabIndex        =   41
         Top             =   4905
         Width           =   6465
         _ExtentX        =   11404
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
         Left            =   8310
         TabIndex        =   40
         Top             =   4680
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   6690
         TabIndex        =   39
         Top             =   4680
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   7575
         TabIndex        =   38
         Top             =   4680
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
         Left            =   180
         TabIndex        =   37
         Top             =   4680
         Width           =   6465
         _ExtentX        =   11404
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
         Left            =   8310
         TabIndex        =   36
         Top             =   4455
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   6690
         TabIndex        =   35
         Top             =   4455
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   7575
         TabIndex        =   34
         Top             =   4455
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
         Left            =   180
         TabIndex        =   33
         Top             =   4455
         Width           =   6465
         _ExtentX        =   11404
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
         Left            =   8310
         TabIndex        =   32
         Top             =   4230
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   6690
         TabIndex        =   31
         Top             =   4230
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   7575
         TabIndex        =   30
         Top             =   4230
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
         Left            =   180
         TabIndex        =   29
         Top             =   4230
         Width           =   6465
         _ExtentX        =   11404
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
         Left            =   8310
         TabIndex        =   28
         Top             =   4005
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   6690
         TabIndex        =   27
         Top             =   4005
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   7575
         TabIndex        =   26
         Top             =   4005
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
         Left            =   180
         TabIndex        =   25
         Top             =   4005
         Width           =   6465
         _ExtentX        =   11404
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
         Left            =   8310
         TabIndex        =   24
         Top             =   3780
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   6690
         TabIndex        =   23
         Top             =   3780
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   7575
         TabIndex        =   22
         Top             =   3780
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
         Left            =   180
         TabIndex        =   21
         Top             =   3780
         Width           =   6465
         _ExtentX        =   11404
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
         Left            =   8310
         TabIndex        =   20
         Top             =   3555
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   6690
         TabIndex        =   19
         Top             =   3555
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   7575
         TabIndex        =   18
         Top             =   3555
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
         Left            =   180
         TabIndex        =   17
         Top             =   3555
         Width           =   6465
         _ExtentX        =   11404
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
         Left            =   8310
         TabIndex        =   16
         Top             =   3330
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   6690
         TabIndex        =   15
         Top             =   3330
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   7575
         TabIndex        =   14
         Top             =   3330
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
         Left            =   180
         TabIndex        =   13
         Top             =   3330
         Width           =   6465
         _ExtentX        =   11404
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
         Left            =   8310
         TabIndex        =   12
         Top             =   3105
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   6690
         TabIndex        =   11
         Top             =   3105
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   7575
         TabIndex        =   10
         Top             =   3105
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
         Left            =   180
         TabIndex        =   9
         Top             =   3105
         Width           =   6465
         _ExtentX        =   11404
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
         Left            =   8310
         TabIndex        =   8
         Top             =   2880
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   6690
         TabIndex        =   7
         Top             =   2880
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   7575
         TabIndex        =   6
         Top             =   2880
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
         Left            =   180
         TabIndex        =   5
         Top             =   2880
         Width           =   6465
         _ExtentX        =   11404
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
         Left            =   180
         TabIndex        =   0
         Top             =   2550
         Width           =   6465
         _ExtentX        =   11404
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
         Left            =   7575
         TabIndex        =   1
         Top             =   2550
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
         Left            =   6690
         TabIndex        =   2
         Top             =   2550
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   8310
         TabIndex        =   3
         Top             =   2550
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "SUBTOTAL"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin ReportX.ReportField rf2 
         Height          =   270
         Left            =   2880
         TabIndex        =   50
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
         TabIndex        =   51
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
         TabIndex        =   52
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
         TabIndex        =   53
         Top             =   840
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   476
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
      End
      Begin ReportX.ReportField txtDesc 
         Height          =   225
         Index           =   10
         Left            =   180
         TabIndex        =   54
         Top             =   5130
         Width           =   6465
         _ExtentX        =   11404
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
         Left            =   180
         TabIndex        =   55
         Top             =   5350
         Width           =   6465
         _ExtentX        =   11404
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
         Left            =   180
         TabIndex        =   56
         Top             =   5580
         Width           =   6465
         _ExtentX        =   11404
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
         Left            =   180
         TabIndex        =   57
         Top             =   5800
         Width           =   6465
         _ExtentX        =   11404
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
         Left            =   6690
         TabIndex        =   58
         Top             =   5130
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   6690
         TabIndex        =   59
         Top             =   5355
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   6690
         TabIndex        =   60
         Top             =   5580
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   6690
         TabIndex        =   61
         Top             =   5805
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   7575
         TabIndex        =   62
         Top             =   5130
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
         Left            =   7575
         TabIndex        =   63
         Top             =   5355
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
         Left            =   7575
         TabIndex        =   64
         Top             =   5580
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
         Left            =   7575
         TabIndex        =   65
         Top             =   5805
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
         Left            =   8310
         TabIndex        =   66
         Top             =   5130
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   8310
         TabIndex        =   67
         Top             =   5355
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   8310
         TabIndex        =   68
         Top             =   5580
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   8310
         TabIndex        =   69
         Top             =   5805
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   180
         TabIndex        =   73
         Top             =   6030
         Width           =   6465
         _ExtentX        =   11404
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
         Left            =   6690
         TabIndex        =   74
         Top             =   6030
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   7575
         TabIndex        =   75
         Top             =   6030
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
         Left            =   8310
         TabIndex        =   76
         Top             =   6030
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   180
         TabIndex        =   77
         Top             =   6250
         Width           =   6465
         _ExtentX        =   11404
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
         Left            =   6690
         TabIndex        =   78
         Top             =   6255
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   7575
         TabIndex        =   79
         Top             =   6255
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
         Left            =   8310
         TabIndex        =   80
         Top             =   6255
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   180
         TabIndex        =   81
         Top             =   6480
         Width           =   6465
         _ExtentX        =   11404
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
         Left            =   6690
         TabIndex        =   82
         Top             =   6480
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   7575
         TabIndex        =   83
         Top             =   6480
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
         Left            =   8310
         TabIndex        =   84
         Top             =   6480
         Width           =   855
         _ExtentX        =   1508
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
      Begin ReportX.ReportField rfVista 
         Height          =   285
         Left            =   7380
         TabIndex        =   85
         Top             =   1500
         Visible         =   0   'False
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
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
         BordaLagura     =   2
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField rfDescVista 
         Height          =   285
         Left            =   6300
         TabIndex        =   86
         Top             =   1500
         Visible         =   0   'False
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
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
         BordaLagura     =   2
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField rfPrazo 
         Height          =   285
         Left            =   7380
         TabIndex        =   87
         Top             =   2145
         Visible         =   0   'False
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
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
         BordaLagura     =   2
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField rfDescPrazo 
         Height          =   285
         Left            =   6300
         TabIndex        =   88
         Top             =   2145
         Visible         =   0   'False
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
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
         BordaLagura     =   2
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField16 
         Height          =   270
         Left            =   240
         TabIndex        =   89
         Top             =   1965
         Width           =   585
         _ExtentX        =   1032
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
         Left            =   900
         TabIndex        =   90
         Top             =   1980
         Width           =   1605
         _ExtentX        =   2831
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
         Left            =   240
         TabIndex        =   91
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
         Left            =   930
         TabIndex        =   92
         Top             =   1200
         Width           =   4485
         _ExtentX        =   7911
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
      Begin ReportX.ReportField ReportField3 
         Height          =   270
         Left            =   255
         TabIndex        =   93
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
         Left            =   720
         TabIndex        =   94
         Top             =   1455
         Width           =   4680
         _ExtentX        =   8255
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
         Left            =   240
         TabIndex        =   95
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
      Begin ReportX.ReportField txtRef 
         Height          =   270
         Left            =   1440
         TabIndex        =   96
         Top             =   2220
         Width           =   3960
         _ExtentX        =   6985
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
         Left            =   2520
         TabIndex        =   97
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
         Left            =   2820
         TabIndex        =   98
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
         Left            =   240
         TabIndex        =   99
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
         Left            =   900
         TabIndex        =   100
         Top             =   1710
         Width           =   4485
         _ExtentX        =   7911
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
      Begin ReportX.ReportField lblSubTotal 
         Height          =   270
         Left            =   8940
         TabIndex        =   103
         Top             =   6750
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
      Begin ReportX.ReportField txtSubtotal 
         Height          =   270
         Left            =   9840
         TabIndex        =   104
         Top             =   6750
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
         Left            =   8940
         TabIndex        =   105
         Top             =   7035
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
      Begin ReportX.ReportField txtDesconto 
         Height          =   270
         Left            =   9840
         TabIndex        =   106
         Top             =   7040
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
         Left            =   8940
         TabIndex        =   107
         Top             =   7605
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
      Begin ReportX.ReportField txtTotal 
         Height          =   270
         Left            =   9840
         TabIndex        =   108
         Top             =   7600
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
      Begin ReportX.ReportField ReportField1 
         Height          =   270
         Left            =   8940
         TabIndex        =   110
         Top             =   7320
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
         Left            =   9840
         TabIndex        =   111
         Top             =   7320
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
      Begin ReportX.ReportField ReportField2 
         Height          =   300
         Left            =   9240
         TabIndex        =   112
         Top             =   2550
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "DESC"
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
         Index           =   0
         Left            =   9240
         TabIndex        =   113
         Top             =   2880
         Width           =   855
         _ExtentX        =   1508
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
         Index           =   1
         Left            =   9240
         TabIndex        =   114
         Top             =   3105
         Width           =   855
         _ExtentX        =   1508
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
         Index           =   2
         Left            =   9240
         TabIndex        =   115
         Top             =   3330
         Width           =   855
         _ExtentX        =   1508
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
         Index           =   3
         Left            =   9240
         TabIndex        =   116
         Top             =   3555
         Width           =   855
         _ExtentX        =   1508
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
         Index           =   4
         Left            =   9240
         TabIndex        =   117
         Top             =   3780
         Width           =   855
         _ExtentX        =   1508
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
         Index           =   5
         Left            =   9240
         TabIndex        =   118
         Top             =   4005
         Width           =   855
         _ExtentX        =   1508
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
         Index           =   6
         Left            =   9240
         TabIndex        =   119
         Top             =   4230
         Width           =   855
         _ExtentX        =   1508
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
         Index           =   7
         Left            =   9240
         TabIndex        =   120
         Top             =   4455
         Width           =   855
         _ExtentX        =   1508
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
         Index           =   8
         Left            =   9240
         TabIndex        =   121
         Top             =   4680
         Width           =   855
         _ExtentX        =   1508
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
         Index           =   9
         Left            =   9240
         TabIndex        =   122
         Top             =   4905
         Width           =   855
         _ExtentX        =   1508
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
         Index           =   10
         Left            =   9240
         TabIndex        =   123
         Top             =   5130
         Width           =   855
         _ExtentX        =   1508
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
         Index           =   11
         Left            =   9240
         TabIndex        =   124
         Top             =   5355
         Width           =   855
         _ExtentX        =   1508
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
         Index           =   12
         Left            =   9240
         TabIndex        =   125
         Top             =   5580
         Width           =   855
         _ExtentX        =   1508
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
         Index           =   13
         Left            =   9240
         TabIndex        =   126
         Top             =   5805
         Width           =   855
         _ExtentX        =   1508
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
         Index           =   14
         Left            =   9240
         TabIndex        =   127
         Top             =   6030
         Width           =   855
         _ExtentX        =   1508
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
         Index           =   15
         Left            =   9240
         TabIndex        =   128
         Top             =   6255
         Width           =   855
         _ExtentX        =   1508
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
         Index           =   16
         Left            =   9240
         TabIndex        =   129
         Top             =   6480
         Width           =   855
         _ExtentX        =   1508
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
      Begin ReportX.ReportField ReportField4 
         Height          =   300
         Left            =   10180
         TabIndex        =   130
         Top             =   2550
         Width           =   855
         _ExtentX        =   1508
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
      Begin ReportX.ReportField txtTotProd 
         Height          =   225
         Index           =   0
         Left            =   10180
         TabIndex        =   131
         Top             =   2880
         Width           =   855
         _ExtentX        =   1508
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
      Begin ReportX.ReportField txtTotProd 
         Height          =   225
         Index           =   1
         Left            =   10180
         TabIndex        =   132
         Top             =   3105
         Width           =   855
         _ExtentX        =   1508
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
      Begin ReportX.ReportField txtTotProd 
         Height          =   225
         Index           =   2
         Left            =   10180
         TabIndex        =   133
         Top             =   3330
         Width           =   855
         _ExtentX        =   1508
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
      Begin ReportX.ReportField txtTotProd 
         Height          =   225
         Index           =   3
         Left            =   10180
         TabIndex        =   134
         Top             =   3555
         Width           =   855
         _ExtentX        =   1508
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
      Begin ReportX.ReportField txtTotProd 
         Height          =   225
         Index           =   4
         Left            =   10180
         TabIndex        =   135
         Top             =   3780
         Width           =   855
         _ExtentX        =   1508
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
      Begin ReportX.ReportField txtTotProd 
         Height          =   225
         Index           =   5
         Left            =   10180
         TabIndex        =   136
         Top             =   4005
         Width           =   855
         _ExtentX        =   1508
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
      Begin ReportX.ReportField txtTotProd 
         Height          =   225
         Index           =   6
         Left            =   10180
         TabIndex        =   137
         Top             =   4230
         Width           =   855
         _ExtentX        =   1508
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
      Begin ReportX.ReportField txtTotProd 
         Height          =   225
         Index           =   7
         Left            =   10180
         TabIndex        =   138
         Top             =   4455
         Width           =   855
         _ExtentX        =   1508
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
      Begin ReportX.ReportField txtTotProd 
         Height          =   225
         Index           =   8
         Left            =   10180
         TabIndex        =   139
         Top             =   4680
         Width           =   855
         _ExtentX        =   1508
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
      Begin ReportX.ReportField txtTotProd 
         Height          =   225
         Index           =   9
         Left            =   10180
         TabIndex        =   140
         Top             =   4905
         Width           =   855
         _ExtentX        =   1508
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
      Begin ReportX.ReportField txtTotProd 
         Height          =   225
         Index           =   10
         Left            =   10180
         TabIndex        =   141
         Top             =   5130
         Width           =   855
         _ExtentX        =   1508
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
      Begin ReportX.ReportField txtTotProd 
         Height          =   225
         Index           =   11
         Left            =   10180
         TabIndex        =   142
         Top             =   5355
         Width           =   855
         _ExtentX        =   1508
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
      Begin ReportX.ReportField txtTotProd 
         Height          =   225
         Index           =   12
         Left            =   10180
         TabIndex        =   143
         Top             =   5580
         Width           =   855
         _ExtentX        =   1508
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
      Begin ReportX.ReportField txtTotProd 
         Height          =   225
         Index           =   13
         Left            =   10180
         TabIndex        =   144
         Top             =   5805
         Width           =   855
         _ExtentX        =   1508
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
      Begin ReportX.ReportField txtTotProd 
         Height          =   225
         Index           =   14
         Left            =   10180
         TabIndex        =   145
         Top             =   6030
         Width           =   855
         _ExtentX        =   1508
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
      Begin ReportX.ReportField txtTotProd 
         Height          =   225
         Index           =   15
         Left            =   10180
         TabIndex        =   146
         Top             =   6255
         Width           =   855
         _ExtentX        =   1508
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
      Begin ReportX.ReportField txtTotProd 
         Height          =   225
         Index           =   16
         Left            =   10180
         TabIndex        =   147
         Top             =   6480
         Width           =   855
         _ExtentX        =   1508
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
      Begin ReportX.ReportField ReportField6 
         Height          =   255
         Left            =   5580
         TabIndex        =   148
         Top             =   1200
         Width           =   3375
         _ExtentX        =   5953
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
      Begin ReportX.ReportField txtBanco 
         Height          =   180
         Left            =   8580
         TabIndex        =   149
         Top             =   315
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   318
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
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtAgencia 
         Height          =   180
         Left            =   8580
         TabIndex        =   150
         Top             =   510
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   318
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
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtConta 
         Height          =   180
         Left            =   9840
         TabIndex        =   151
         Top             =   510
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   318
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
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField cboTipo 
         Height          =   180
         Left            =   10140
         TabIndex        =   152
         Top             =   315
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   318
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
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtPix 
         Height          =   180
         Left            =   8580
         TabIndex        =   153
         Top             =   900
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   318
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
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtFavorecido 
         Height          =   180
         Left            =   8580
         TabIndex        =   154
         Top             =   700
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   318
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
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField8 
         Height          =   180
         Left            =   8580
         TabIndex        =   155
         Top             =   120
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   318
         Caption         =   "DADOS BANCÁRIO"
         Alignment       =   2
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
      Begin VB.Shape Shape7 
         BorderWidth     =   2
         Height          =   1095
         Left            =   8520
         Top             =   60
         Width           =   2595
      End
      Begin VB.Line Line7 
         X1              =   10140
         X2              =   10140
         Y1              =   6700
         Y2              =   2520
      End
      Begin VB.Line Line2 
         X1              =   9180
         X2              =   9180
         Y1              =   6700
         Y2              =   2520
      End
      Begin VB.Label lblDataRecebido 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ENTREGUE EM: ____/____/_____"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5880
         TabIndex        =   109
         Top             =   1800
         Width           =   2850
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape6 
         BorderWidth     =   2
         Height          =   1215
         Left            =   8880
         Top             =   6720
         Width           =   2235
      End
      Begin VB.Line Line6 
         X1              =   3960
         X2              =   7800
         Y1              =   7440
         Y2              =   7440
      End
      Begin VB.Label LblRecebedor 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RECEBEDOR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5340
         TabIndex        =   102
         Top             =   7440
         Width           =   1050
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "*A validade desse orçamento está descrita no cabeçalho desse orçamento."
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
         Left            =   120
         TabIndex        =   101
         Top             =   6720
         Width           =   5490
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         Height          =   1395
         Left            =   9000
         Top             =   1140
         Width           =   2115
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "*A validade desse orçamento está sujeito a disponibilidade no estoque."
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
         Left            =   120
         TabIndex        =   72
         Top             =   6840
         Width           =   4410
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
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
         Left            =   3420
         TabIndex        =   71
         Top             =   7140
         Visible         =   0   'False
         Width           =   3030
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
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
         Left            =   1200
         TabIndex        =   70
         Top             =   7620
         Visible         =   0   'False
         Width           =   4860
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line5 
         X1              =   8280
         X2              =   8280
         Y1              =   6700
         Y2              =   2520
      End
      Begin VB.Line Line4 
         X1              =   7560
         X2              =   7560
         Y1              =   6700
         Y2              =   2520
      End
      Begin VB.Line Line3 
         X1              =   6660
         X2              =   6660
         Y1              =   6700
         Y2              =   2520
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   11040
         Y1              =   2835
         Y2              =   2835
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         Height          =   4215
         Left            =   120
         Top             =   2520
         Width           =   10995
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         Height          =   1395
         Left            =   5520
         Top             =   1140
         Width           =   3495
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   1395
         Left            =   120
         Top             =   1140
         Width           =   5415
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   1095
         Left            =   120
         Top             =   60
         Width           =   10995
      End
      Begin VB.Image imgLogo 
         Height          =   1035
         Left            =   180
         Stretch         =   -1  'True
         Top             =   60
         Width           =   2595
      End
      Begin VB.Label Label1 
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
         Left            =   1680
         TabIndex        =   45
         Top             =   7500
         Visible         =   0   'False
         Width           =   2910
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "REL_Pedido_Orcamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rPd As ADODB.Recordset
Dim rCl As ADODB.Recordset
Dim rIt As ADODB.Recordset
Dim rPc As ADODB.Recordset
Dim rFu As ADODB.Recordset
Dim sSQL As String
Dim r As ADODB.Recordset

'arquivo .ini
Public cCfg As ConfigItem
Public oIni As Ini
'Public var_ImpNormal As String

Public Sub loadPedidos(Pedido As Long)
Dim Cont As Long
Dim wValorFormatado As String

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

Set rPd = dbData.OpenRecordset("SELECT * FROM pedidos WHERE (cod_pedido = " & Pedido & ");")
Set rCl = dbData.OpenRecordset("SELECT * FROM cliente WHERE (codigo = " & rPd("cod_cliente") & ");")
Set rIt = dbData.OpenRecordset("SELECT produtos.fabricante AS var_fab, pedidos_itens.codigo, pedidos_itens.cod_produto, produtos.descricao, pedidos_itens.preco, pedidos_itens.quantidade, pedidos_itens.Desconto, pedidos_itens.Subtotal, pedidos_itens.Total FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto WHERE (pedidos_itens.cod_pedido = " & Pedido & ") ORDER BY pedidos_itens.codigo DESC;")
Set rPc = dbData.OpenRecordset("SELECT * FROM parcelas WHERE (cod_pedido = " & Pedido & ") ORDER BY numero;")
Set rFu = dbData.OpenRecordset("SELECT * FROM funcionario WHERE (codigo = " & rPd("cod_funcionario") & ");")

txtNumero.Caption = "PEDIDO Nº " & Format(Pedido, "000000")
ReportField6.Caption = rPd("TIPO_PEDIDO")

If rPd("cod_cliente") = 0 Then Exit Sub

'DADOS DO CLIENTE
If UCase(rCl("tipo")) = "FÍSICA" Then
   ReportField16.Caption = "CPF:"
Else
   ReportField16.Caption = "CNPJ:"
End If

'DADOS DO CLIENTE
txtCliente.Caption = ValidateNull(rCl("nome"))
'frNome.Caption = ValidateNull(rCl("nome")) & " - CPF: " & ValidateNull(rCl("cpf"))
txtEnd.Caption = ValidateNull(rCl("endereco")) & " - " & ValidateNull(rCl("bairro"))
txtRef.Caption = ValidateNull(rCl("ponto_de_referencia"))
txtCidade.Caption = ValidateNull(rCl("cidade")) & "-" & ValidateNull(rCl("estado")) & " TEL: " & ValidateNull(rCl("telefone1")) & " / " & ValidateNull(rCl("CELULAR"))
txtCPF.Caption = ValidateNull(rCl("cpf"))
txtRG.Caption = ValidateNull(rCl("rg"))

'DADOS DO PEDIDO
txtData.Caption = "DATA: " & String(1, " ") + Format(rPd("data_compra"), "dd/mm/yy")
txtVendedor.Caption = "FUNC.: " & rFu("nome")
txtValidade.Caption = "VALIDADE: " & Format(rPd("DATA_ENTREGA"), "DD/MM/YY")
'txtValidade.Caption = "VALIDADE: " & Format(DateAdd("d", 10, Date), "DD/MM/YY")

'DADOS DO FINANCEIRO
'Á VISTA
'rfVista.Caption = "TOTAL: " & Format(rPd("total"), ocMONEY)

''If IsNull(rPd("valor_desc")) Then
''   wValorFormatado = "0,00"
''Else
''   If rPd("tipo_desc") = "R" Then
''      wValorFormatado = FormatNumber(rPd("valor_desc"), 2)
''   Else
''      wValorFormatado = FormatNumber(rPd("valor_desc"), 2) & "%"
''   End If
''End If

'rfDescVista.Caption = "DESC: " & wValorFormatado

'Á PRAZO
'rfPrazo.Caption = "TOTAL: " & Format(rPd("subtotal"), ocMONEY)
'rfDescPrazo.Caption = "DESC: " & FormatNumber(0, 2) & "%"

txtSubtotal.Caption = FormatNumber(rPd("subtotal"), 2)


Dim varValorPorcentagem As Currency

'If rPd("tipo_desc") = "R" Then
'    If IsNull(rPd("valor_desc")) Then
'        varValorPorcentagem = FormatNumber(0, 2)
'    Else
'        varValorPorcentagem = FormatNumber(rPd("valor_desc"), 2)
'    End If
If rPd("tipo_desc") = "P" Then
    If IsNull(rPd("valor_desc")) Then
        varValorPorcentagem = FormatNumber(0, 2)
    Else
        varValorPorcentagem = FormatNumber(rPd("VALOR_DESC"), 2)
    End If
Else
    If IsNull(rPd("valor_desc")) Then
        varValorPorcentagem = FormatNumber(0, 2)
    Else
        'varValorPorcentagem = FormatNumber(((rPd("SUBTOTAL") * rPd("VALOR_DESC")) / 100), 2)
        varValorPorcentagem = FormatNumber(Abs((((rPd("TOTAL") - rPd("SUBTOTAL")) / rPd("SUBTOTAL")) * 100)), 2)
    End If
End If

'((rPd("TOTAL") - rPd("SUBTOTAL")) / rPd("SUBTOTAL")) * 100

'If IsNull(rPd("valor_desc")) Then
'   wValorFormatado = "0,00"
'Else
'   If rPd("tipo_desc") = "R" Then
'      wValorFormatado = FormatNumber(rPd("valor_desc"), 2)
'   Else
'      wValorFormatado = FormatNumber(rPd("valor_desc"), 2) & "%"
'   End If
'End If

txtDescontoRS.Caption = FormatNumber(rPd("ValorDescReal"), 2)
txtDesconto.Caption = FormatNumber(varValorPorcentagem, 2)

txtTotal.Caption = String(1, " ") + FormatNumber(rPd("total"), 2)

'INSIRO OS ITENS
If Not rIt.EOF Then rIt.MoveLast
If Not rIt.BOF Then rIt.MoveFirst

Relatorio.NumeroRegistros = Round((rIt.RecordCount / 17) + 0.49)
Relatorio.NomeImpressora = var_ImpNormal
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
rf4.Caption = "CNPJ: " & r("cnpj") & " - IE: " & r("ie") & " - FONE: " & r("telefone") & " - " & r("celular") & ""
txtBanco.Caption = "Banco: " & ValidateNull(r("Banco"))
txtAgencia.Caption = "Agência: " & ValidateNull(r("Agencia"))
txtConta.Caption = "Conta: " & ValidateNull(r("Conta"))
cboTipo.Caption = "Tipo: " & ValidateNull(r("Tipo"))
txtFavorecido.Caption = "Fav.: " & ValidateNull(r("Favorecido"))
txtPix.Caption = "Pix: " & ValidateNull(r("Pix"))

If Not IsNull(r("caminho")) Then
   If Dir$(r("caminho")) <> "" Then Set imgLogo.Picture = LoadPicture(r("caminho"))
End If

If r.State <> 0 Then r.Close
Set r = Nothing
  
TrataErro:
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not rPd Is Nothing Then If rPd.State <> 0 Then rPd.Close
If Not rCl Is Nothing Then If rCl.State <> 0 Then rCl.Close
If Not rIt Is Nothing Then If rIt.State <> 0 Then rIt.Close
If Not rPc Is Nothing Then If rPc.State <> 0 Then rPc.Close
If Not rFu Is Nothing Then If rFu.State <> 0 Then rFu.Close
End Sub

Private Sub Relatorio_IniciarSecao(ByVal Secao As ReportX.TSecao, ByVal Ordem As Byte)
   Dim i As Integer
   
   'produtos do pedido
   For i = 0 To 16
      txtDesc(i).Caption = ""
      txtQuant(i).Caption = ""
      txtUnit(i).Caption = ""
      txtTot(i).Caption = ""
      txtDesco(i).Caption = ""
      txtTotProd(i).Caption = ""
      
      If Not rIt.EOF Then
         txtDesc(i).Caption = String(1, " ") + rIt("descricao")
         txtQuant(i).Caption = rIt("quantidade")
         txtUnit(i).Caption = FormatNumber(rIt("preco"), 2)
         txtTot(i).Caption = FormatNumber((rIt("subtotal")), 2)
         txtDesco(i).Caption = FormatNumber((rIt("Desconto")), 2)
         txtTotProd(i).Caption = FormatNumber((rIt("total")), 2)
         
         rIt.MoveNext
      End If
   Next
   Exit Sub
End Sub

