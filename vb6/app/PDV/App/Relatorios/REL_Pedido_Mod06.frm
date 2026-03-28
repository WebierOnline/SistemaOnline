VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.Ocx"
Begin VB.Form REL_Pedido_Mod06 
   ClientHeight    =   8685
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   12750
   LinkTopic       =   "Form1"
   ScaleHeight     =   153.194
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   224.896
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   60
      TabIndex        =   45
      Top             =   8160
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
      Height          =   8115
      Left            =   0
      Top             =   0
      Width           =   12750
      _ExtentX        =   22490
      _ExtentY        =   14314
      Ordem           =   1
      Begin ReportX.ReportField frTitParc 
         Height          =   255
         Left            =   7500
         TabIndex        =   92
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
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
         Left            =   8700
         TabIndex        =   88
         Top             =   1500
         Width           =   615
         _ExtentX        =   1085
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
         Left            =   7980
         TabIndex        =   84
         Top             =   1500
         Width           =   705
         _ExtentX        =   1244
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
         Index           =   0
         Left            =   7620
         TabIndex        =   80
         Top             =   1500
         Width           =   300
         _ExtentX        =   529
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
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   5
         Left            =   180
         TabIndex        =   59
         Top             =   4005
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   9
         Left            =   180
         TabIndex        =   63
         Top             =   4905
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   8
         Left            =   180
         TabIndex        =   62
         Top             =   4680
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   7
         Left            =   180
         TabIndex        =   61
         Top             =   4455
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   6
         Left            =   180
         TabIndex        =   60
         Top             =   4230
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   4
         Left            =   180
         TabIndex        =   58
         Top             =   3780
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   3
         Left            =   180
         TabIndex        =   57
         Top             =   3555
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   2
         Left            =   180
         TabIndex        =   56
         Top             =   3330
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   54
         Top             =   2880
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtVendedor 
         Height          =   240
         Left            =   10260
         TabIndex        =   53
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
         Left            =   9480
         TabIndex        =   4
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
         Left            =   9480
         TabIndex        =   51
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
         Left            =   9480
         TabIndex        =   52
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
      Begin ReportX.ReportField lblSubTotal 
         Height          =   270
         Left            =   9480
         TabIndex        =   47
         Top             =   6960
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
         Left            =   10380
         TabIndex        =   48
         Top             =   6960
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
         Left            =   9480
         TabIndex        =   49
         Top             =   7260
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   10380
         TabIndex        =   50
         Top             =   7260
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
      Begin ReportX.ReportField ReportField2 
         Height          =   300
         Left            =   180
         TabIndex        =   46
         Top             =   2520
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   529
         Caption         =   "REF."
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
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   9
         Left            =   8760
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
         Left            =   7065
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
         Left            =   7980
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
         Left            =   1060
         TabIndex        =   41
         Top             =   4905
         Width           =   5925
         _ExtentX        =   10451
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
         Left            =   8760
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
         Left            =   7065
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
         Left            =   7980
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
         Left            =   1060
         TabIndex        =   37
         Top             =   4680
         Width           =   5925
         _ExtentX        =   10451
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
         Left            =   8760
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
         Left            =   7065
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
         Left            =   7980
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
         Left            =   1060
         TabIndex        =   33
         Top             =   4455
         Width           =   5925
         _ExtentX        =   10451
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
         Left            =   8760
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
         Left            =   7065
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
         Left            =   7980
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
         Left            =   1060
         TabIndex        =   29
         Top             =   4230
         Width           =   5925
         _ExtentX        =   10451
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
         Left            =   8760
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
         Left            =   7065
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
         Left            =   7980
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
         Left            =   1060
         TabIndex        =   25
         Top             =   4005
         Width           =   5925
         _ExtentX        =   10451
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
         Left            =   8760
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
         Left            =   7065
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
         Left            =   7980
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
         Left            =   1060
         TabIndex        =   21
         Top             =   3780
         Width           =   5925
         _ExtentX        =   10451
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
         Left            =   8760
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
         Left            =   7065
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
         Left            =   7980
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
         Left            =   1060
         TabIndex        =   17
         Top             =   3555
         Width           =   5925
         _ExtentX        =   10451
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
         Left            =   8760
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
         Left            =   7065
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
         Left            =   7980
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
         Left            =   1060
         TabIndex        =   13
         Top             =   3330
         Width           =   5925
         _ExtentX        =   10451
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
         Left            =   8760
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
         Left            =   7065
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
         Left            =   7980
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
         Left            =   1060
         TabIndex        =   9
         Top             =   3105
         Width           =   5925
         _ExtentX        =   10451
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
         Left            =   8760
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
         Left            =   7065
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
         Left            =   7980
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
         Left            =   1060
         TabIndex        =   5
         Top             =   2880
         Width           =   5925
         _ExtentX        =   10451
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
         Left            =   1065
         TabIndex        =   0
         Top             =   2520
         Width           =   5925
         _ExtentX        =   10451
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
         Left            =   7980
         TabIndex        =   1
         Top             =   2520
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
         Left            =   7065
         TabIndex        =   2
         Top             =   2520
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
         Left            =   8760
         TabIndex        =   3
         Top             =   2520
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "SUBTOTAL"
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
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   55
         Top             =   3105
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         TabIndex        =   64
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
         TabIndex        =   65
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
         TabIndex        =   66
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
         TabIndex        =   67
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
      Begin ReportX.ReportField ReportField16 
         Height          =   270
         Left            =   180
         TabIndex        =   68
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
         TabIndex        =   69
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
         Height          =   275
         Left            =   180
         TabIndex        =   70
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
         TabIndex        =   71
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
      Begin ReportX.ReportField ReportField3 
         Height          =   270
         Left            =   195
         TabIndex        =   72
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
         TabIndex        =   73
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
         TabIndex        =   74
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
         Left            =   1380
         TabIndex        =   75
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
         TabIndex        =   76
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
         TabIndex        =   77
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
         TabIndex        =   78
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
         TabIndex        =   79
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
      Begin ReportX.ReportField frNumParc 
         Height          =   195
         Index           =   1
         Left            =   7620
         TabIndex        =   81
         Top             =   1680
         Width           =   300
         _ExtentX        =   529
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
         Left            =   7620
         TabIndex        =   82
         Top             =   1860
         Width           =   300
         _ExtentX        =   529
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
         Left            =   7620
         TabIndex        =   83
         Top             =   2040
         Width           =   300
         _ExtentX        =   529
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
         Left            =   7980
         TabIndex        =   85
         Top             =   1680
         Width           =   705
         _ExtentX        =   1244
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
         Index           =   2
         Left            =   7980
         TabIndex        =   86
         Top             =   1860
         Width           =   705
         _ExtentX        =   1244
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
         Index           =   3
         Left            =   7980
         TabIndex        =   87
         Top             =   2040
         Width           =   705
         _ExtentX        =   1244
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
         Index           =   1
         Left            =   8700
         TabIndex        =   89
         Top             =   1680
         Width           =   615
         _ExtentX        =   1085
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
         Left            =   8700
         TabIndex        =   90
         Top             =   1860
         Width           =   615
         _ExtentX        =   1085
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
         Left            =   8700
         TabIndex        =   91
         Top             =   2040
         Width           =   615
         _ExtentX        =   1085
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
         Left            =   7620
         TabIndex        =   93
         Top             =   2220
         Width           =   300
         _ExtentX        =   529
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
         Left            =   7980
         TabIndex        =   94
         Top             =   2220
         Width           =   705
         _ExtentX        =   1244
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
         Index           =   4
         Left            =   8700
         TabIndex        =   95
         Top             =   2220
         Width           =   615
         _ExtentX        =   1085
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
         Left            =   9480
         TabIndex        =   96
         Top             =   1200
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   450
         Caption         =   "PEDIDO"
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
      Begin ReportX.ReportField ReportField13 
         Height          =   270
         Left            =   9480
         TabIndex        =   97
         Top             =   7560
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   10380
         TabIndex        =   98
         Top             =   7560
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
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   10
         Left            =   180
         TabIndex        =   99
         Top             =   5130
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   11
         Left            =   180
         TabIndex        =   100
         Top             =   5350
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   12
         Left            =   180
         TabIndex        =   101
         Top             =   5580
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   13
         Left            =   180
         TabIndex        =   102
         Top             =   5800
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Index           =   10
         Left            =   1060
         TabIndex        =   103
         Top             =   5130
         Width           =   5925
         _ExtentX        =   10451
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
         Left            =   1060
         TabIndex        =   104
         Top             =   5350
         Width           =   5925
         _ExtentX        =   10451
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
         Left            =   1060
         TabIndex        =   105
         Top             =   5580
         Width           =   5925
         _ExtentX        =   10451
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
         Left            =   1060
         TabIndex        =   106
         Top             =   5800
         Width           =   5925
         _ExtentX        =   10451
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
         Left            =   7065
         TabIndex        =   107
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
         Left            =   7065
         TabIndex        =   108
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
         Left            =   7065
         TabIndex        =   109
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
         Left            =   7065
         TabIndex        =   110
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
         Left            =   7980
         TabIndex        =   111
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
         Left            =   7980
         TabIndex        =   112
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
         Left            =   7980
         TabIndex        =   113
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
         Left            =   7980
         TabIndex        =   114
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
         Left            =   8760
         TabIndex        =   115
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
         Left            =   8760
         TabIndex        =   116
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
         Left            =   8760
         TabIndex        =   117
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
         Left            =   8760
         TabIndex        =   118
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
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   14
         Left            =   180
         TabIndex        =   121
         Top             =   6030
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   1060
         TabIndex        =   122
         Top             =   6030
         Width           =   5925
         _ExtentX        =   10451
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
         Left            =   7065
         TabIndex        =   123
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
         Left            =   7980
         TabIndex        =   124
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
         Left            =   8760
         TabIndex        =   125
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
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   15
         Left            =   180
         TabIndex        =   126
         Top             =   6250
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   1060
         TabIndex        =   127
         Top             =   6250
         Width           =   5925
         _ExtentX        =   10451
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
         Left            =   7065
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
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   15
         Left            =   7980
         TabIndex        =   129
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
         Left            =   8760
         TabIndex        =   130
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
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   16
         Left            =   180
         TabIndex        =   131
         Top             =   6480
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   1060
         TabIndex        =   132
         Top             =   6480
         Width           =   5925
         _ExtentX        =   10451
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
         Left            =   7065
         TabIndex        =   133
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
         Left            =   7980
         TabIndex        =   134
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
         Left            =   8760
         TabIndex        =   135
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
      Begin ReportX.ReportField ReportField1 
         Height          =   300
         Left            =   9690
         TabIndex        =   139
         Top             =   2520
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
         Left            =   9690
         TabIndex        =   140
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
         Left            =   9690
         TabIndex        =   141
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
         Left            =   9690
         TabIndex        =   142
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
         Left            =   9690
         TabIndex        =   143
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
         Left            =   9690
         TabIndex        =   144
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
         Left            =   9690
         TabIndex        =   145
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
         Left            =   9690
         TabIndex        =   146
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
         Left            =   9690
         TabIndex        =   147
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
         Left            =   9690
         TabIndex        =   148
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
         Left            =   9690
         TabIndex        =   149
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
         Left            =   9690
         TabIndex        =   150
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
         Left            =   9690
         TabIndex        =   151
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
         Left            =   9690
         TabIndex        =   152
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
         Left            =   9690
         TabIndex        =   153
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
         Left            =   9690
         TabIndex        =   154
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
         Left            =   9690
         TabIndex        =   155
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
         Left            =   9690
         TabIndex        =   156
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
      Begin ReportX.ReportField ReportField21 
         Height          =   300
         Left            =   10605
         TabIndex        =   157
         Top             =   2520
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
         Left            =   10605
         TabIndex        =   158
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
         Left            =   10605
         TabIndex        =   159
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
         Left            =   10605
         TabIndex        =   160
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
         Left            =   10605
         TabIndex        =   161
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
         Left            =   10605
         TabIndex        =   162
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
         Left            =   10605
         TabIndex        =   163
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
         Left            =   10605
         TabIndex        =   164
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
         Left            =   10605
         TabIndex        =   165
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
         Left            =   10605
         TabIndex        =   166
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
         Left            =   10605
         TabIndex        =   167
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
         Left            =   10605
         TabIndex        =   168
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
         Left            =   10605
         TabIndex        =   169
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
         Left            =   10605
         TabIndex        =   170
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
         Left            =   10605
         TabIndex        =   171
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
         Left            =   10605
         TabIndex        =   172
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
         Left            =   10605
         TabIndex        =   173
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
         Left            =   10605
         TabIndex        =   174
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
      Begin VB.Line Line8 
         X1              =   9660
         X2              =   9660
         Y1              =   6840
         Y2              =   2520
      End
      Begin VB.Line Line7 
         X1              =   10560
         X2              =   10560
         Y1              =   6840
         Y2              =   2520
      End
      Begin VB.Label LblRecebedor2 
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
         Left            =   6840
         TabIndex        =   138
         Top             =   7680
         Width           =   1050
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
         Left            =   2400
         TabIndex        =   137
         Top             =   7620
         Width           =   2850
         WordWrap        =   -1  'True
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
         Left            =   6840
         TabIndex        =   136
         Top             =   7500
         Width           =   1050
      End
      Begin VB.Line Line6 
         X1              =   5460
         X2              =   9300
         Y1              =   7500
         Y2              =   7500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*Não aceitamos devolução de produtos."
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
         TabIndex        =   120
         Top             =   6900
         Width           =   5370
         WordWrap        =   -1  'True
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
         Left            =   120
         TabIndex        =   119
         Top             =   7020
         Width           =   5370
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line5 
         X1              =   8700
         X2              =   8700
         Y1              =   6840
         Y2              =   2520
      End
      Begin VB.Line Line4 
         X1              =   7920
         X2              =   7920
         Y1              =   6840
         Y2              =   2520
      End
      Begin VB.Line Line3 
         X1              =   7020
         X2              =   7020
         Y1              =   6840
         Y2              =   2520
      End
      Begin VB.Line Line2 
         X1              =   1020
         X2              =   1020
         Y1              =   6840
         Y2              =   2520
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   11520
         Y1              =   2835
         Y2              =   2835
      End
      Begin VB.Shape Shape6 
         BorderWidth     =   2
         Height          =   1095
         Left            =   9420
         Top             =   6840
         Width           =   2115
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         Height          =   4335
         Left            =   120
         Top             =   2520
         Width           =   11415
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         Height          =   1395
         Left            =   9420
         Top             =   1140
         Width           =   2115
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         Height          =   1395
         Left            =   7440
         Top             =   1140
         Width           =   1995
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
         Width           =   11415
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
Attribute VB_Name = "REL_Pedido_Mod06"
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
Dim rPR As ADODB.Recordset

'arquivo .ini
Public cCfg As ConfigItem
Public oIni As Ini

Dim sSQL As String
Dim r As ADODB.Recordset
Public vDeclararRecebedor As String

Public Sub loadPedidos(Pedido As Long)
Dim i As Integer
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
Set rPR = dbData.OpenRecordset("SELECT * FROM pedidos_recebedor WHERE (cod_pedido = " & Pedido & ");")
Set rCl = dbData.OpenRecordset("SELECT * FROM cliente WHERE (codigo = " & rPd("cod_cliente") & ");")
Set rIt = dbData.OpenRecordset("SELECT produtos.fabricante AS var_fab, produtos.tamanho AS var_tam, produtos.ref AS var_ref, pedidos_itens.codigo, pedidos_itens.cod_produto, produtos.descricao, pedidos_itens.preco, pedidos_itens.quantidade, pedidos_itens.Desconto, pedidos_itens.Subtotal, pedidos_itens.Total  FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto WHERE (pedidos_itens.cod_pedido = " & Pedido & ") ORDER BY pedidos_itens.codigo DESC;")
Set rPc = dbData.OpenRecordset("SELECT * FROM parcelas WHERE (cod_pedido = " & Pedido & ") ORDER BY numero;")
Set rFu = dbData.OpenRecordset("SELECT * FROM funcionario WHERE (codigo = " & rPd("cod_funcionario") & ");")

For i = 0 To 4
   frNumParc(i).Caption = ""
   frVencParc(i).Caption = ""
   frValorParc(i).Caption = ""
Next

Cont = 0

Do While Not rPc.EOF
   frNumParc(Cont).Caption = Format(rPc("numero"), "00")
   frVencParc(Cont).Caption = Format(rPc("data"), "dd/mm/yy")
   frValorParc(Cont).Caption = Format(rPc("valor"), ocMONEY)
   Cont = Cont + 1
   rPc.MoveNext
Loop

'txtNumero.Caption = "Nº " & Format(PDV.txtCodPedido.Text, "000000")
txtNumero.Caption = "Nº " & Format(Pedido, "000000")

'DADOS DO CLIENTE
txtCliente.Caption = ValidateNull(rCl("nome"))
txtEnd.Caption = ValidateNull(rCl("endereco")) & " - " & ValidateNull(rCl("bairro"))
txtRef.Caption = ValidateNull(rCl("ponto_de_referencia"))
txtCidade.Caption = ValidateNull(rCl("cidade")) & "-" & ValidateNull(rCl("estado")) & " TEL: " & ValidateNull(rCl("telefone1")) & " / " & ValidateNull(rCl("CELULAR"))
txtCPF.Caption = ValidateNull(rCl("cpf"))
txtRG.Caption = ValidateNull(rCl("rg"))

'DADOS DO RECECEDOR
If vDeclararRecebedor = "SIM" Then
    If Not rPR.EOF Then
        lblRecebedor.Visible = True
        lblDataRecebido.Visible = True
        LblRecebedor2.Visible = True
        Line6.Visible = True
        lblRecebedor.Caption = ValidateNull(rPR("Recebedor"))
        lblDataRecebido.Caption = "Recebido em: " & String(1, " ") + Format(rPd("data_compra"), "dd/mm/yy")
    Else
        lblRecebedor.Visible = False
        lblDataRecebido.Visible = False
        LblRecebedor2.Visible = False
        Line6.Visible = False
    End If
Else
    lblRecebedor.Visible = False
    lblDataRecebido.Visible = False
    LblRecebedor2.Visible = False
    Line6.Visible = False
End If

'DADOS DO PEDIDO
txtData.Caption = String(1, " ") + Format(rPd("data_compra"), "dd/mm/yy")
txtVendedor.Caption = rFu("nome")
txtVenda.Caption = UCase(rPd("tipo_pagamento"))

If rPd("tipo_pagamento") = "À Vista" Then
   txtData.Mostrar = True
   txtVendedor.Mostrar = True
   txtVenda.Mostrar = True
   txtSubtotal.Mostrar = True
   txtDesconto.Mostrar = True
   txtTotal.Mostrar = True
   lblSubTotal.Mostrar = True
   lblDesconto.Mostrar = True
Else
   txtData.Mostrar = True
   txtVendedor.Mostrar = True
   txtVenda.Mostrar = True
   txtSubtotal.Mostrar = True
   txtDesconto.Mostrar = True
   lblSubTotal.Mostrar = True
   lblDesconto.Mostrar = True
End If

txtSubtotal.Caption = Format(rPd("subtotal"), ocMONEY)

If IsNull(rPd("valor_desc")) Then
   wValorFormatado = "0,00"
Else
   If rPd("tipo_desc") = "R" Then
      wValorFormatado = Format(rPd("valor_desc"), ocMONEY)
   Else
      wValorFormatado = FormatNumber(rPd("valor_desc"), 2) & "%"
   End If
End If

txtDesconto.Caption = wValorFormatado
txtTotal.Caption = String(1, " ") + Format(rPd("total"), ocMONEY)

If rPd("tipo_pedido") = "BALCAO" Then

ElseIf rPd("tipo_pedido") = "ORÇAMENTO" Then

End If

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
   
'se precisa pedi senha nas opções do menu avançado
Set cCfg = sysConfig("DECLARARRECEBEDOR")
vDeclararRecebedor = cCfg.Value
Set cCfg = Nothing

sSQL = "SELECT TOP 1 * FROM empresa ORDER BY fantasia;"
Set r = dbData.OpenRecordset(sSQL)

rf1.Caption = r("fantasia")
rf2.Caption = r("razao")
rf3.Caption = r("endereco") & ", " & r("cidade") & "-" & r("estado")
rf4.Caption = "CNPJ: " & r("cnpj") & " - IE: " & r("ie") & " - FONE: " & r("telefone") & " - " & r("celular") & ""

If Not IsNull(r("caminho")) Then
   If Dir$(r("caminho")) <> "" Then Set imgLogo.Picture = LoadPicture(r("caminho"))
End If

If r.State <> 0 Then r.Close
Set r = Nothing
Exit Sub

'colocar o nome da maquina na barra de status
Dim var_Impressora As String
'Dim oIni As Ini

Set oIni = New Ini
oIni.Arquivo = appPathApp & "config.ini"
var_Impressora = oIni.LerTexto("IMPRESSORA_FOLHA", "impressora")
Set oIni = Nothing
   
TrataErro:
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Not rPd Is Nothing Then If rPd.State <> 0 Then rPd.Close
   If Not rPR Is Nothing Then If rPR.State <> 0 Then rPR.Close
   If Not rCl Is Nothing Then If rCl.State <> 0 Then rCl.Close
   If Not rIt Is Nothing Then If rIt.State <> 0 Then rIt.Close
   If Not rPc Is Nothing Then If rPc.State <> 0 Then rPc.Close
   If Not rFu Is Nothing Then If rFu.State <> 0 Then rFu.Close
End Sub

Private Sub Relatorio_IniciarSecao(ByVal Secao As ReportX.TSecao, ByVal Ordem As Byte)
   Dim i As Integer

   Dim cCfg As ConfigItem
   Dim tipoEmpresa As Integer
               
   Set cCfg = sysConfig("TIPO_EMPRESA")
   tipoEmpresa = cCfg.Value
   Set cCfg = Nothing
   
   'produtos do pedido
   For i = 0 To 16
      txtDesc(i).Caption = ""
      txtCodProd(i).Caption = ""
      txtQuant(i).Caption = ""
      txtUnit(i).Caption = ""
      txtTot(i).Caption = ""
      txtDesco(i).Caption = ""
      txtTotProd(i).Caption = ""
      
      If Not rIt.EOF Then
      
         If tipoEmpresa = 4 Then
            txtDesc(i).Caption = String(1, " ") + rIt("descricao") & " / " & rIt("var_tam") & " / " & rIt("var_fab")
         Else
            txtDesc(i).Caption = String(1, " ") + rIt("descricao")
         End If
         
         txtCodProd(i).Caption = ValidateNull(rIt("var_ref"))
         'txtQuant(i).Caption = ValidateNull(rIt("quantidade"))
         'txtUnit(i).Caption = Format(rIt("preco"), ocMONEY)
         'txtTot(i).Caption = Format((rIt("preco") * rIt("quantidade")), ocMONEY)
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

