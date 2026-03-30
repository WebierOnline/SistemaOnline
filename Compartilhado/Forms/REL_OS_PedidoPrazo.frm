VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.Ocx"
Begin VB.Form REL_OS_PedidoPrazo 
   ClientHeight    =   8685
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   ScaleHeight     =   153.194
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   198.967
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportSection ReportSection2 
      Align           =   1  'Align Top
      Height          =   8115
      Left            =   0
      Top             =   0
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   14314
      Ordem           =   1
      Begin ReportX.ReportField rf4 
         Height          =   275
         Left            =   2880
         TabIndex        =   68
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
      Begin ReportX.ReportField rf3 
         Height          =   270
         Left            =   2880
         TabIndex        =   67
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
      Begin ReportX.ReportField frEnt 
         Height          =   195
         Left            =   4280
         TabIndex        =   147
         Top             =   1500
         Visible         =   0   'False
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   344
         Caption         =   "PG"
         Alignment       =   1
         Mostrar         =   0   'False
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
      Begin ReportX.ReportField frTitParc 
         Height          =   255
         Left            =   4320
         TabIndex        =   93
         Top             =   1200
         Width           =   2115
         _ExtentX        =   3731
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
         Left            =   5720
         TabIndex        =   89
         Top             =   1500
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   4930
         TabIndex        =   85
         Top             =   1500
         Width           =   765
         _ExtentX        =   1349
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
         Index           =   0
         Left            =   4560
         TabIndex        =   81
         Top             =   1500
         Width           =   375
         _ExtentX        =   661
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
      Begin ReportX.ReportField frNome 
         Height          =   6600
         Left            =   9360
         TabIndex        =   64
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   11642
         Linhas          =   24
         Caption         =   ""
         Alignment       =   2
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
         AlinhamentoVertical=   2
         Rotacao         =   90
      End
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   5
         Left            =   180
         TabIndex        =   59
         Top             =   4005
         Width           =   1410
         _ExtentX        =   2487
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
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   9
         Left            =   180
         TabIndex        =   63
         Top             =   4905
         Width           =   1410
         _ExtentX        =   2487
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
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   8
         Left            =   180
         TabIndex        =   62
         Top             =   4680
         Width           =   1410
         _ExtentX        =   2487
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
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   7
         Left            =   180
         TabIndex        =   61
         Top             =   4455
         Width           =   1410
         _ExtentX        =   2487
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
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   6
         Left            =   180
         TabIndex        =   60
         Top             =   4230
         Width           =   1410
         _ExtentX        =   2487
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
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   4
         Left            =   180
         TabIndex        =   58
         Top             =   3780
         Width           =   1410
         _ExtentX        =   2487
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
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   3
         Left            =   180
         TabIndex        =   57
         Top             =   3555
         Width           =   1410
         _ExtentX        =   2487
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
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   2
         Left            =   180
         TabIndex        =   56
         Top             =   3330
         Width           =   1410
         _ExtentX        =   2487
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
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   54
         Top             =   2880
         Width           =   1410
         _ExtentX        =   2487
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
      Begin ReportX.ReportField txtVendedor 
         Height          =   240
         Left            =   7320
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
         Left            =   6540
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
         Left            =   6540
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
         Left            =   6540
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
         Left            =   6540
         TabIndex        =   47
         Top             =   6960
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
         Left            =   7440
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
         Left            =   6780
         TabIndex        =   49
         Top             =   7260
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   476
         Caption         =   "Desc.:"
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
         Left            =   7440
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
         Top             =   2550
         Width           =   1410
         _ExtentX        =   2487
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
      Begin ReportX.ReportField txtTot 
         Height          =   225
         Index           =   9
         Left            =   7470
         TabIndex        =   44
         Top             =   4905
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   5610
         TabIndex        =   43
         Top             =   4905
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   6735
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
         Left            =   1630
         TabIndex        =   41
         Top             =   4905
         Width           =   3945
         _ExtentX        =   6959
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
         Left            =   7470
         TabIndex        =   40
         Top             =   4680
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   5610
         TabIndex        =   39
         Top             =   4680
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   6735
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
         Left            =   1630
         TabIndex        =   37
         Top             =   4680
         Width           =   3945
         _ExtentX        =   6959
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
         Left            =   7470
         TabIndex        =   36
         Top             =   4455
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   5610
         TabIndex        =   35
         Top             =   4455
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   6735
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
         Left            =   1630
         TabIndex        =   33
         Top             =   4455
         Width           =   3945
         _ExtentX        =   6959
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
         Left            =   7470
         TabIndex        =   32
         Top             =   4230
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   5610
         TabIndex        =   31
         Top             =   4230
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   6735
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
         Left            =   1630
         TabIndex        =   29
         Top             =   4230
         Width           =   3945
         _ExtentX        =   6959
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
         Left            =   7470
         TabIndex        =   28
         Top             =   4005
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   5610
         TabIndex        =   27
         Top             =   4005
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   6735
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
         Left            =   1630
         TabIndex        =   25
         Top             =   4005
         Width           =   3945
         _ExtentX        =   6959
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
         Left            =   7470
         TabIndex        =   24
         Top             =   3780
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   5610
         TabIndex        =   23
         Top             =   3780
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   6735
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
         Left            =   1630
         TabIndex        =   21
         Top             =   3780
         Width           =   3945
         _ExtentX        =   6959
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
         Left            =   7470
         TabIndex        =   20
         Top             =   3555
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   5610
         TabIndex        =   19
         Top             =   3555
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   6735
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
         Left            =   1630
         TabIndex        =   17
         Top             =   3555
         Width           =   3945
         _ExtentX        =   6959
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
         Left            =   7470
         TabIndex        =   16
         Top             =   3330
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   5610
         TabIndex        =   15
         Top             =   3330
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   6735
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
         Left            =   1630
         TabIndex        =   13
         Top             =   3330
         Width           =   3945
         _ExtentX        =   6959
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
         Left            =   7470
         TabIndex        =   12
         Top             =   3105
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   5610
         TabIndex        =   11
         Top             =   3105
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   6735
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
         Left            =   1630
         TabIndex        =   9
         Top             =   3105
         Width           =   3945
         _ExtentX        =   6959
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
         Left            =   7470
         TabIndex        =   8
         Top             =   2880
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   5610
         TabIndex        =   7
         Top             =   2880
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   6735
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
         Left            =   1630
         TabIndex        =   5
         Top             =   2880
         Width           =   3945
         _ExtentX        =   6959
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
         Left            =   1630
         TabIndex        =   0
         Top             =   2550
         Width           =   3945
         _ExtentX        =   6959
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
         Left            =   6735
         TabIndex        =   1
         Top             =   2550
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   529
         Caption         =   "QTDE"
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
         Left            =   5610
         TabIndex        =   2
         Top             =   2550
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   7470
         TabIndex        =   3
         Top             =   2550
         Width           =   1095
         _ExtentX        =   1931
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
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   55
         Top             =   3105
         Width           =   1410
         _ExtentX        =   2487
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
      Begin ReportX.ReportField rf2 
         Height          =   270
         Left            =   2880
         TabIndex        =   65
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
         TabIndex        =   66
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
      Begin ReportX.ReportField ReportField16 
         Height          =   270
         Left            =   180
         TabIndex        =   69
         Top             =   1965
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   476
         Caption         =   "CNPJ:"
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
         Left            =   720
         TabIndex        =   70
         Top             =   1965
         Width           =   1875
         _ExtentX        =   3307
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
         TabIndex        =   71
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
         TabIndex        =   72
         Top             =   1200
         Width           =   3380
         _ExtentX        =   5953
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
         TabIndex        =   73
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
         TabIndex        =   74
         Top             =   1455
         Width           =   3590
         _ExtentX        =   6324
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
         TabIndex        =   75
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
         TabIndex        =   76
         Top             =   2220
         Width           =   2860
         _ExtentX        =   5054
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
         Left            =   2610
         TabIndex        =   77
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
         Left            =   2940
         TabIndex        =   78
         Top             =   1965
         Width           =   1290
         _ExtentX        =   2275
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
         TabIndex        =   79
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
         TabIndex        =   80
         Top             =   1710
         Width           =   3405
         _ExtentX        =   6006
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
         Left            =   4560
         TabIndex        =   82
         Top             =   1680
         Width           =   375
         _ExtentX        =   661
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
         Left            =   4560
         TabIndex        =   83
         Top             =   1860
         Width           =   375
         _ExtentX        =   661
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
         Left            =   4560
         TabIndex        =   84
         Top             =   2040
         Width           =   375
         _ExtentX        =   661
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
         Left            =   4930
         TabIndex        =   86
         Top             =   1680
         Width           =   765
         _ExtentX        =   1349
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
         Index           =   2
         Left            =   4930
         TabIndex        =   87
         Top             =   1860
         Width           =   765
         _ExtentX        =   1349
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
         Index           =   3
         Left            =   4930
         TabIndex        =   88
         Top             =   2040
         Width           =   765
         _ExtentX        =   1349
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
      Begin ReportX.ReportField frValorParc 
         Height          =   195
         Index           =   1
         Left            =   5720
         TabIndex        =   90
         Top             =   1680
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   5720
         TabIndex        =   91
         Top             =   1860
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   5720
         TabIndex        =   92
         Top             =   2040
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   4560
         TabIndex        =   94
         Top             =   2220
         Width           =   375
         _ExtentX        =   661
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
         Left            =   4930
         TabIndex        =   95
         Top             =   2220
         Width           =   765
         _ExtentX        =   1349
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
      Begin ReportX.ReportField frValorParc 
         Height          =   195
         Index           =   4
         Left            =   5720
         TabIndex        =   96
         Top             =   2220
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   6540
         TabIndex        =   97
         Top             =   1200
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   450
         Caption         =   "ORDEM"
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
         Left            =   6780
         TabIndex        =   98
         Top             =   7560
         Width           =   645
         _ExtentX        =   1138
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
         Left            =   7440
         TabIndex        =   99
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
         TabIndex        =   100
         Top             =   5130
         Width           =   1410
         _ExtentX        =   2487
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
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   11
         Left            =   180
         TabIndex        =   101
         Top             =   5350
         Width           =   1410
         _ExtentX        =   2487
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
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   12
         Left            =   180
         TabIndex        =   102
         Top             =   5580
         Width           =   1410
         _ExtentX        =   2487
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
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   13
         Left            =   180
         TabIndex        =   103
         Top             =   5800
         Width           =   1410
         _ExtentX        =   2487
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
         Index           =   10
         Left            =   1630
         TabIndex        =   104
         Top             =   5130
         Width           =   3945
         _ExtentX        =   6959
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
         Left            =   1630
         TabIndex        =   105
         Top             =   5350
         Width           =   3945
         _ExtentX        =   6959
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
         Left            =   1630
         TabIndex        =   106
         Top             =   5580
         Width           =   3945
         _ExtentX        =   6959
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
         Left            =   1630
         TabIndex        =   107
         Top             =   5800
         Width           =   3945
         _ExtentX        =   6959
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
         Left            =   5610
         TabIndex        =   108
         Top             =   5130
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   5610
         TabIndex        =   109
         Top             =   5350
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   5610
         TabIndex        =   110
         Top             =   5580
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   5610
         TabIndex        =   111
         Top             =   5800
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   6735
         TabIndex        =   112
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
         Index           =   12
         Left            =   6735
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
         Left            =   6735
         TabIndex        =   114
         Top             =   5800
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
         Left            =   7470
         TabIndex        =   115
         Top             =   5130
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   7470
         TabIndex        =   116
         Top             =   5350
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   7470
         TabIndex        =   117
         Top             =   5580
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   7470
         TabIndex        =   118
         Top             =   5800
         Width           =   1095
         _ExtentX        =   1931
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
      Begin ReportX.ReportField rfValorNum 
         Height          =   1500
         Left            =   8820
         TabIndex        =   119
         Top             =   240
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   2646
         Linhas          =   7
         Caption         =   "0,00"
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
         Borda           =   15
         BackColor       =   14737632
         AlturaLivre     =   -1  'True
         Rotacao         =   90
      End
      Begin ReportX.ReportField ReportField8 
         Height          =   3960
         Left            =   8760
         TabIndex        =   120
         Top             =   2880
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   6985
         Linhas          =   12
         Caption         =   "N O T A  P R O M I S S Ó R I A"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   11.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlinhamentoVertical=   1
         Rotacao         =   90
      End
      Begin ReportX.ReportField ReportField10 
         Height          =   2100
         Left            =   11220
         TabIndex        =   121
         Top             =   3060
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   3704
         Linhas          =   10
         Caption         =   "ASSINATURA"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
         Rotacao         =   90
      End
      Begin ReportX.ReportField ReportField12 
         Height          =   390
         Left            =   9360
         TabIndex        =   122
         Top             =   6720
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   688
         Linhas          =   2
         Caption         =   "Eu, "
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         Rotacao         =   90
      End
      Begin ReportX.ReportField ReportField14 
         Height          =   405
         Left            =   8820
         TabIndex        =   123
         Top             =   1800
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   714
         Caption         =   "R$"
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
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   2
         Rotacao         =   90
      End
      Begin ReportX.ReportField ReportField17 
         Height          =   6915
         Left            =   9660
         TabIndex        =   124
         Top             =   180
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   12197
         Linhas          =   2
         Caption         =   " pagarei por essa única via de Nota Promissória ou a sua ordem a quantia de"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   2
         Rotacao         =   90
      End
      Begin ReportX.ReportField frValorEst 
         Height          =   6765
         Left            =   10020
         TabIndex        =   125
         Top             =   300
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   11933
         Linhas          =   24
         Caption         =   ""
         Alignment       =   2
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
         AlinhamentoVertical=   2
         Rotacao         =   90
      End
      Begin ReportX.ReportField ReportField19 
         Height          =   3210
         Left            =   10320
         TabIndex        =   126
         Top             =   3840
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   5662
         Linhas          =   2
         Caption         =   "em moeda corrente deste País."
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   2
         Rotacao         =   90
      End
      Begin ReportX.ReportField rfData 
         Height          =   3435
         Left            =   10560
         TabIndex        =   127
         Top             =   300
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   6059
         Linhas          =   2
         Caption         =   ""
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         Rotacao         =   90
      End
      Begin ReportX.ReportField txtQuant 
         Height          =   225
         Index           =   11
         Left            =   6735
         TabIndex        =   131
         Top             =   5350
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
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   14
         Left            =   180
         TabIndex        =   132
         Top             =   6060
         Width           =   1410
         _ExtentX        =   2487
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
         Index           =   14
         Left            =   1635
         TabIndex        =   133
         Top             =   6060
         Width           =   3945
         _ExtentX        =   6959
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
         Left            =   5610
         TabIndex        =   134
         Top             =   6060
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   6735
         TabIndex        =   135
         Top             =   6060
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
         Left            =   7470
         TabIndex        =   136
         Top             =   6060
         Width           =   1095
         _ExtentX        =   1931
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
         TabIndex        =   137
         Top             =   6285
         Width           =   1410
         _ExtentX        =   2487
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
         Index           =   15
         Left            =   1635
         TabIndex        =   138
         Top             =   6285
         Width           =   3945
         _ExtentX        =   6959
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
         Left            =   5610
         TabIndex        =   139
         Top             =   6285
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   6735
         TabIndex        =   140
         Top             =   6285
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
         Left            =   7470
         TabIndex        =   141
         Top             =   6285
         Width           =   1095
         _ExtentX        =   1931
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
         TabIndex        =   142
         Top             =   6480
         Width           =   1410
         _ExtentX        =   2487
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
         Index           =   16
         Left            =   1635
         TabIndex        =   143
         Top             =   6510
         Width           =   3945
         _ExtentX        =   6959
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
         Left            =   5610
         TabIndex        =   144
         Top             =   6510
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   6735
         TabIndex        =   145
         Top             =   6510
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
         Left            =   7470
         TabIndex        =   146
         Top             =   6510
         Width           =   1095
         _ExtentX        =   1931
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
      Begin ReportX.ReportField ReportField11 
         Height          =   270
         Left            =   180
         TabIndex        =   150
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
         TabIndex        =   151
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
      Begin ReportX.ReportField ReportField1 
         Height          =   255
         Left            =   180
         TabIndex        =   152
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
         TabIndex        =   153
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
         TabIndex        =   154
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
         TabIndex        =   155
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
         TabIndex        =   156
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
         TabIndex        =   157
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
         TabIndex        =   158
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
      Begin VB.Shape Shape8 
         BorderWidth     =   2
         Height          =   870
         Left            =   120
         Top             =   6840
         Width           =   2715
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*Não aceitamos devolução de produtos com a embalagem violada."
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
         TabIndex        =   149
         Top             =   6900
         Width           =   3390
         WordWrap        =   -1  'True
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
         TabIndex        =   130
         Top             =   7740
         Visible         =   0   'False
         Width           =   2370
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   2880
         TabIndex        =   129
         Top             =   7020
         Width           =   2730
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*Clientes com mais de 30 dias de vencidos sujeito a inclusão do nome no SPC e Serasa "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         TabIndex        =   128
         Top             =   7260
         Width           =   3360
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image1 
         Height          =   570
         Left            =   8760
         Picture         =   "REL_OS_PedidoPrazo.frx":0000
         Stretch         =   -1  'True
         Top             =   7260
         Width           =   2655
      End
      Begin VB.Line Line6 
         X1              =   11220
         X2              =   11220
         Y1              =   2340
         Y2              =   5880
      End
      Begin VB.Shape Shape7 
         BorderWidth     =   2
         Height          =   7875
         Left            =   8640
         Shape           =   4  'Rounded Rectangle
         Top             =   60
         Width           =   2835
      End
      Begin VB.Line Line5 
         X1              =   7440
         X2              =   7440
         Y1              =   6840
         Y2              =   2520
      End
      Begin VB.Line Line4 
         X1              =   6720
         X2              =   6720
         Y1              =   6840
         Y2              =   2520
      End
      Begin VB.Line Line3 
         X1              =   5580
         X2              =   5580
         Y1              =   6840
         Y2              =   2520
      End
      Begin VB.Line Line2 
         X1              =   1620
         X2              =   1620
         Y1              =   6840
         Y2              =   2520
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   8580
         Y1              =   2830
         Y2              =   2830
      End
      Begin VB.Shape Shape6 
         BorderWidth     =   2
         Height          =   1095
         Left            =   6480
         Top             =   6840
         Width           =   2115
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         Height          =   4335
         Left            =   120
         Top             =   2520
         Width           =   8475
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         Height          =   1395
         Left            =   6480
         Top             =   1140
         Width           =   2115
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         Height          =   1395
         Left            =   4260
         Top             =   1140
         Width           =   2235
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   1395
         Left            =   120
         Top             =   1140
         Width           =   4155
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   1095
         Left            =   120
         Top             =   60
         Width           =   8475
      End
      Begin VB.Image imgLogo 
         Height          =   1035
         Left            =   180
         Stretch         =   -1  'True
         Top             =   120
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   2880
         TabIndex        =   45
         Top             =   7140
         Width           =   2610
         WordWrap        =   -1  'True
      End
   End
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   120
      TabIndex        =   148
      Top             =   8160
      Visible         =   0   'False
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
End
Attribute VB_Name = "REL_OS_PedidoPrazo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String
Dim r As ADODB.Recordset
Dim rPd As ADODB.Recordset
Dim rCl As ADODB.Recordset
Dim rIt As ADODB.Recordset
Dim rTotais As ADODB.Recordset
Dim rPc As ADODB.Recordset
Dim rFu As ADODB.Recordset
Dim rOs As ADODB.Recordset
Public cCfg As ConfigItem       'arquivo .ini
Public oIni As Ini              'arquivo .ini
Dim var_ImpNormal As String

Public Sub loadPedidos(ByVal Pedido As Long, ByVal Tipo As String)
   Dim i As Integer
   Dim Cont As Long
   Dim wValorFormatado As String
   Dim sSQL As String

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
    
    Set rOs = dbData.OpenRecordset("SELECT cod_pedido, cod_cliente, cod_funcionario FROM os WHERE (cod_os = " & vCodOS & ");")
    Set rPd = dbData.OpenRecordset("SELECT * FROM pedidos WHERE (cod_pedido = " & Pedido & ");")
    Set rCl = dbData.OpenRecordset("SELECT * FROM cliente WHERE (codigo = " & rOs("cod_cliente") & ");")
   
   sSQL = "SELECT 'PRODUTO' AS tipo_item, produtos.descricao, pedidos_itens.quantidade, pedidos_itens.preco, (pedidos_itens.preco * pedidos_itens.quantidade) as total, pedidos_itens.codigo, '' as varTipo, '' as varSerie, '' as varFogo, '' as varFabricante, '' as varMedida, '' as varAro, '' as varBanda FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto WHERE (pedidos_itens.cod_pedido = " & Pedido & ")"
   
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
        
        txtQuantGeral.Caption = Format(CInt(txtQuantPecas.Caption) + CInt(txtQuantServicos.Caption), "000")
        txtTotalPecasServicos.Caption = FormatNumber(CCur(txtTotalPecas.Caption) + CCur(txtTotalServicos.Caption), 2)
   End If
   
   If rIt.EOF Then
        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
           sSQL = "SELECT 'SERVIÇO' AS tipo_item, descricao, quantidade, preco, (quantidade * preco) AS total, codigo, '' as varFabricante, '', '', '', '', '', '' " & _
           "FROM OS_Servicos_Auto WHERE (cod_os = " & vCodOS & ")"
        ElseIf vTipoOS = "Recapadora" Then
            sSQL = "SELECT 'SERVIÇO' AS tipo_item, descricao, quantidade, preco, ISNULL((quantidade * preco), 0) AS total, codigo, TIPO as var_TipoPneu, SERIE as var_serie, FOGO as var_fogo, ARO as var_aro, BANDA as var_banda, DOTE as var_dote, MEDIDA as var_medida, FABRICANTE as var_fabricante " & _
            "FROM OS_servicos_recapadora WHERE (cod_os = " & vCodOS & ")"
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
           sSQL = "SELECT 'SERVIÇO' AS tipo_item, descricao, quantidade, preco, (quantidade * preco) AS total, codigo, '' as varFabricante, '', '', '', '', '', '' " & _
           "FROM OS_Servicos_Auto WHERE (cod_os = " & vCodOS & ")"
        ElseIf vTipoOS = "Comunicação Visual" Then
           sSQL = "SELECT 'SERVIÇO' AS tipo_item, descricao, quantidade, preco, (quantidade * preco) AS total, codigo, '' as varFabricante, '', '', '', '', '', '' " & _
           "FROM OS_Servicos_Comunicacao WHERE (cod_os = " & vCodOS & ")"
        End If
   Else
        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
             sSQL = "SELECT 'PRODUTO' AS tipo_item, produtos.descricao, pedidos_itens.quantidade, pedidos_itens.preco, (pedidos_itens.preco * pedidos_itens.quantidade) as total, pedidos_itens.codigo, produtos.Fabricante as varFabricante FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto WHERE (pedidos_itens.cod_pedido = " & Pedido & ")"
        ElseIf vTipoOS = "Recapadora" Then
             sSQL = "SELECT 'PRODUTO' AS tipo_item, produtos.descricao, pedidos_itens.quantidade, pedidos_itens.preco, (pedidos_itens.preco * pedidos_itens.quantidade) as total, pedidos_itens.codigo, '' as var_TipoPneu, '' as var_serie, '' as var_fogo, '' as var_aro, '' as var_banda, '' as var_dote, '' as var_medida, '' as var_fabricante FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto WHERE (pedidos_itens.cod_pedido = " & Pedido & ")"
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
             sSQL = "SELECT 'PRODUTO' AS tipo_item, produtos.descricao, pedidos_itens.quantidade, pedidos_itens.preco, (pedidos_itens.preco * pedidos_itens.quantidade) as total, pedidos_itens.codigo, produtos.Fabricante as varFabricante FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto WHERE (pedidos_itens.cod_pedido = " & Pedido & ")"
        ElseIf vTipoOS = "Comunicação Visual" Then
             sSQL = "SELECT 'PRODUTO' AS tipo_item, produtos.descricao, pedidos_itens.quantidade, pedidos_itens.preco, (pedidos_itens.preco * pedidos_itens.quantidade) as total, pedidos_itens.codigo, produtos.Fabricante as varFabricante FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto WHERE (pedidos_itens.cod_pedido = " & Pedido & ")"
        End If
        
        If UCase(Tipo) = "OFICINA" Then
           sSQL = sSQL & " UNION "
             If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
                   sSQL = sSQL & "SELECT 'SERVIÇO' AS tipo_item, descricao, quantidade, preco, (quantidade * preco) AS total, codigo, '' " & _
                   "FROM OS_Servicos_Auto WHERE (cod_os = " & vCodOS & ")"
             ElseIf vTipoOS = "Recapadora" Then
                 sSQL = sSQL & "SELECT 'SERVIÇO' AS tipo_item, descricao, quantidade, preco, ISNULL((quantidade * preco), 0) AS total, codigo, TIPO as var_TipoPneu, SERIE as var_serie, FOGO as var_fogo, ARO as var_aro, BANDA as var_banda, DOTE as var_dote, MEDIDA as var_medida, FABRICANTE as var_fabricante " & _
                 "FROM OS_servicos_recapadora WHERE (cod_os = " & vCodOS & ")"
             ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
                   sSQL = sSQL & "SELECT 'SERVIÇO' AS tipo_item, descricao, quantidade, preco, (quantidade * preco) AS total, codigo, '' " & _
                   "FROM OS_Servicos_Auto WHERE (cod_os = " & vCodOS & ")"
             ElseIf vTipoOS = "Comunicação Visual" Then
                   sSQL = sSQL & "SELECT 'SERVIÇO' AS tipo_item, descricao, quantidade, preco, (quantidade * preco) AS total, codigo, '' " & _
                   "FROM OS_Servicos_Auto WHERE (cod_os = " & vCodOS & ")"
             End If
        End If
   End If
   
   Set rIt = dbData.OpenRecordset(sSQL)
   Set rPc = dbData.OpenRecordset("SELECT * FROM parcelas WHERE (cod_pedido = " & Pedido & ") ORDER BY numero;")
   Set rFu = dbData.OpenRecordset("SELECT * FROM funcionario WHERE (codigo = " & rPd("cod_funcionario") & ");")
   
   rfData.Caption = "Vencimento: " & Format(rPc("data"), "dd") & " de " & Format(rPc("data"), "mmmm") & " de " & Format(rPc("data"), "yyyy")    'promissoria
   
   For i = 0 To 4
      frNumParc(i).Caption = ""
      frVencParc(i).Caption = ""
      frValorParc(i).Caption = ""
   Next
      
   If rPc("numero") = 1 And rPc("status") = True Then
      frEnt.Visible = True
   Else
      frEnt.Visible = False
   End If
   
   For i = 0 To 4
      frNumParc(i).Caption = ""
      frVencParc(i).Caption = ""
      frValorParc(i).Caption = ""
   Next
   
   Cont = 0
   
   Do While Not rPc.EOF
      If Cont < 5 Then
         frNumParc(Cont).Caption = Format(rPc("numero"), "00")
         frVencParc(Cont).Caption = Format(rPc("data"), "dd/mm/yy")
         frValorParc(Cont).Caption = Format(rPc("valor"), ocMONEY)
      End If
      Cont = Cont + 1
      rPc.MoveNext
   Loop
   
   txtNumero.Caption = "Nº " & Format(vCodOS, "000000")
   
   If UCase(rCl("tipo")) = "FISICA" Then
      ReportField16.Caption = "CPF:"
   Else
      ReportField16.Caption = "CNPJ:"
   End If
   
   'DADOS DO CLIENTE
   txtCliente.Caption = rCl("nome")
   frNome.Caption = rCl("nome") & " - CPF: " & rCl("cpf")   'promissória
   txtEnd.Caption = rCl("endereco") & " - " & rCl("bairro")
   txtRef.Caption = ValidateNull(rCl("ponto_de_referencia"))
   txtCidade.Caption = rCl("cidade") & "-" & rCl("estado") & " TEL: " & rCl("telefone1")
   txtCPF.Caption = ValidateNull(rCl("cpf"))
   txtRG.Caption = ValidateNull(rCl("RG"))
   
   'DADOS DO PEDIDO
   txtData.Caption = String(1, " ") + Format(rPd("data_compra"), "dd/mm/yy")
   txtVendedor.Caption = rFu("nome")
   txtVenda.Caption = UCase(rPd("tipo_pagamento"))
   'txtPagamento.Caption = rsPedidos!PAGAMENTO
   
   'DADOS DAS PARCELAS
   If rPd("tipo_pagamento") = "À Vista" Then
      txtData.Mostrar = True
      txtVendedor.Mostrar = True
      txtVenda.Mostrar = True
      txtSubtotal.Mostrar = True
      txtDesconto.Mostrar = True
      txtTotal.Mostrar = True
      lblSubtotal.Mostrar = True
      lblDesconto.Mostrar = True
   Else
      txtData.Mostrar = True
      txtVendedor.Mostrar = True
      txtVenda.Mostrar = True
      txtSubtotal.Mostrar = True
      txtDesconto.Mostrar = True
      lblSubtotal.Mostrar = True
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
   rfValorNum.Caption = String(1, " ") + Format(rPd("total"), ocMONEY) 'promissoria
   frValorEst.Caption = String(1, " ") + Format(rPd("total"), ocMONEY) 'promissoria
   
   If rPd("tipo_pedido") <> "ORÇAMENTO" Then rPc.MoveLast
   
   frValorEst.Caption = UCase(NumeroExtenso(rfValorNum.Caption, True)) 'promissoria
   
   If Not rIt.BOF Then rIt.MoveFirst
      
      For i = 0 To 16
      txtDesc(i).Caption = ""
      txtCodProd(i).Caption = ""
      txtQuant(i).Caption = ""
      txtUnit(i).Caption = ""
      txtTot(i).Caption = ""
      
      If Not rIt.EOF Then
        txtCodProd(i).Caption = rIt("tipo_item")
        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
            txtDesc(i).Caption = String(1, " ") + ValidateNull(rIt("descricao"))
        ElseIf vTipoOS = "Recapadora" Then
            If rIt("tipo_item") = "SERVIÇOS" Then
                txtDesc(i).Caption = String(1, " ") + ValidateNull(rIt("descricao")) & " | " & rIt("var_TipoPneu") & " | " & rIt("var_serie") & " | " & rIt("var_fogo") & " | " & rIt("var_aro") & " | " & rIt("var_banda") & " | " & rIt("var_dote") & " | " & rIt("var_medida") & " | " & rIt("var_fabricante") & " "
            Else
                txtDesc(i).Caption = String(1, " ") + ValidateNull(rIt("descricao"))
            End If
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
            txtDesc(i).Caption = String(1, " ") + ValidateNull(rIt("descricao"))
        ElseIf vTipoOS = "Comunicação Visual" Then
            txtDesc(i).Caption = String(1, " ") + ValidateNull(rIt("descricao"))
        End If
         
         txtQuant(i).Caption = rIt("quantidade")
         txtUnit(i).Caption = Format(rIt("preco"), ocMONEY)
         txtTot(i).Caption = Format((rIt("preco") * rIt("quantidade")), ocMONEY)
         
         rIt.MoveNext
      End If
   Next
   
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
rf4.Caption = "CNPJ: " & r("cnpj") & " - IE: " & r("ie") & " - FONE: " & r("telefone") & " - " & r("celular") & ""

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
End Sub

