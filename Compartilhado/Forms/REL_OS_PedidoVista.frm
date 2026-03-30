VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.Ocx"
Begin VB.Form REL_OS_PedidoVista 
   ClientHeight    =   8685
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   ScaleHeight     =   153.194
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   191.559
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   60
      TabIndex        =   0
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
   Begin ReportX.ReportSection ReportSection2 
      Align           =   1  'Align Top
      Height          =   8055
      Left            =   0
      Top             =   0
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   14208
      Ordem           =   1
      Begin ReportX.ReportField rf2 
         Height          =   270
         Left            =   2880
         TabIndex        =   1
         Top             =   420
         Width           =   6345
         _ExtentX        =   11192
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
         TabIndex        =   2
         Top             =   630
         Width           =   6345
         _ExtentX        =   11192
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
         Height          =   270
         Left            =   2880
         TabIndex        =   3
         Top             =   840
         Width           =   6345
         _ExtentX        =   11192
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
      Begin ReportX.ReportField frTitParc 
         Height          =   255
         Left            =   7020
         TabIndex        =   4
         Top             =   1200
         Width           =   2235
         _ExtentX        =   3942
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
         Left            =   8520
         TabIndex        =   5
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
         Left            =   7680
         TabIndex        =   6
         Top             =   1500
         Width           =   795
         _ExtentX        =   1402
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
         Left            =   7260
         TabIndex        =   7
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
      Begin ReportX.ReportField txtCodProd 
         Height          =   225
         Index           =   5
         Left            =   180
         TabIndex        =   8
         Top             =   4005
         Width           =   990
         _ExtentX        =   1746
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
         TabIndex        =   9
         Top             =   4905
         Width           =   990
         _ExtentX        =   1746
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
         TabIndex        =   10
         Top             =   4680
         Width           =   990
         _ExtentX        =   1746
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
         TabIndex        =   11
         Top             =   4455
         Width           =   990
         _ExtentX        =   1746
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
         TabIndex        =   12
         Top             =   4230
         Width           =   990
         _ExtentX        =   1746
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
         TabIndex        =   13
         Top             =   3780
         Width           =   990
         _ExtentX        =   1746
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
         TabIndex        =   14
         Top             =   3555
         Width           =   990
         _ExtentX        =   1746
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
         TabIndex        =   15
         Top             =   3330
         Width           =   990
         _ExtentX        =   1746
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
         TabIndex        =   16
         Top             =   2880
         Width           =   990
         _ExtentX        =   1746
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
         Left            =   10200
         TabIndex        =   17
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
         Left            =   9420
         TabIndex        =   18
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
         Left            =   9420
         TabIndex        =   19
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
         Left            =   9420
         TabIndex        =   20
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
         Left            =   9420
         TabIndex        =   21
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
         Left            =   10320
         TabIndex        =   22
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
         Left            =   9660
         TabIndex        =   23
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
         Left            =   10320
         TabIndex        =   24
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
         TabIndex        =   25
         Top             =   2550
         Width           =   990
         _ExtentX        =   1746
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
         Left            =   8880
         TabIndex        =   26
         Top             =   4905
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   7260
         TabIndex        =   27
         Top             =   4905
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   8200
         TabIndex        =   28
         Top             =   4905
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   1260
         TabIndex        =   29
         Top             =   4905
         Width           =   5985
         _ExtentX        =   10557
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
         Left            =   8880
         TabIndex        =   30
         Top             =   4680
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   7260
         TabIndex        =   31
         Top             =   4680
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   8200
         TabIndex        =   32
         Top             =   4680
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   1260
         TabIndex        =   33
         Top             =   4680
         Width           =   5985
         _ExtentX        =   10557
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
         Left            =   8880
         TabIndex        =   34
         Top             =   4455
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   7260
         TabIndex        =   35
         Top             =   4455
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   8200
         TabIndex        =   36
         Top             =   4455
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   1260
         TabIndex        =   37
         Top             =   4455
         Width           =   5985
         _ExtentX        =   10557
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
         Left            =   8880
         TabIndex        =   38
         Top             =   4230
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   7260
         TabIndex        =   39
         Top             =   4230
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   8200
         TabIndex        =   40
         Top             =   4230
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   1260
         TabIndex        =   41
         Top             =   4230
         Width           =   5985
         _ExtentX        =   10557
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
         Left            =   8880
         TabIndex        =   42
         Top             =   4005
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   7260
         TabIndex        =   43
         Top             =   4005
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   8200
         TabIndex        =   44
         Top             =   4005
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   1260
         TabIndex        =   45
         Top             =   4005
         Width           =   5985
         _ExtentX        =   10557
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
         Left            =   8880
         TabIndex        =   46
         Top             =   3780
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   7260
         TabIndex        =   47
         Top             =   3780
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   8200
         TabIndex        =   48
         Top             =   3780
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   1260
         TabIndex        =   49
         Top             =   3780
         Width           =   5985
         _ExtentX        =   10557
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
         Left            =   8880
         TabIndex        =   50
         Top             =   3555
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   7260
         TabIndex        =   51
         Top             =   3555
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   8200
         TabIndex        =   52
         Top             =   3555
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   1260
         TabIndex        =   53
         Top             =   3555
         Width           =   5985
         _ExtentX        =   10557
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
         Left            =   8880
         TabIndex        =   54
         Top             =   3330
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   7260
         TabIndex        =   55
         Top             =   3330
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   8200
         TabIndex        =   56
         Top             =   3330
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   1260
         TabIndex        =   57
         Top             =   3330
         Width           =   5985
         _ExtentX        =   10557
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
         Left            =   8880
         TabIndex        =   58
         Top             =   3105
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   7260
         TabIndex        =   59
         Top             =   3105
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   8200
         TabIndex        =   60
         Top             =   3105
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   1260
         TabIndex        =   61
         Top             =   3105
         Width           =   5985
         _ExtentX        =   10557
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
         Left            =   8880
         TabIndex        =   62
         Top             =   2880
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   7260
         TabIndex        =   63
         Top             =   2880
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   8200
         TabIndex        =   64
         Top             =   2880
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   1260
         TabIndex        =   65
         Top             =   2880
         Width           =   5985
         _ExtentX        =   10557
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
         Left            =   1260
         TabIndex        =   66
         Top             =   2550
         Width           =   5985
         _ExtentX        =   10557
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
         Left            =   8200
         TabIndex        =   67
         Top             =   2550
         Width           =   630
         _ExtentX        =   1111
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
         Left            =   7260
         TabIndex        =   68
         Top             =   2550
         Width           =   885
         _ExtentX        =   1561
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
         Left            =   8880
         TabIndex        =   69
         Top             =   2550
         Width           =   885
         _ExtentX        =   1561
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
         TabIndex        =   70
         Top             =   3105
         Width           =   990
         _ExtentX        =   1746
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
      Begin ReportX.ReportField rf1 
         Height          =   390
         Left            =   2880
         TabIndex        =   71
         Top             =   60
         Width           =   6360
         _ExtentX        =   11218
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
         TabIndex        =   72
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
         TabIndex        =   73
         Top             =   1965
         Width           =   1965
         _ExtentX        =   3466
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
         TabIndex        =   74
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
         TabIndex        =   75
         Top             =   1200
         Width           =   5985
         _ExtentX        =   10557
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
         TabIndex        =   76
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
         TabIndex        =   77
         Top             =   1455
         Width           =   6180
         _ExtentX        =   10901
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
         TabIndex        =   78
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
         TabIndex        =   79
         Top             =   2220
         Width           =   5460
         _ExtentX        =   9631
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
         Left            =   3060
         TabIndex        =   80
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
         Left            =   3360
         TabIndex        =   81
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
         TabIndex        =   82
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
         TabIndex        =   83
         Top             =   1710
         Width           =   5985
         _ExtentX        =   10557
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
         Left            =   7260
         TabIndex        =   84
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
         Left            =   7260
         TabIndex        =   85
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
         Left            =   7260
         TabIndex        =   86
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
         Left            =   7680
         TabIndex        =   87
         Top             =   1680
         Width           =   795
         _ExtentX        =   1402
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
         Left            =   7680
         TabIndex        =   88
         Top             =   1860
         Width           =   795
         _ExtentX        =   1402
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
         Left            =   7680
         TabIndex        =   89
         Top             =   2040
         Width           =   795
         _ExtentX        =   1402
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
         Left            =   8520
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
         Left            =   8520
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
         Left            =   8520
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
         Left            =   7260
         TabIndex        =   93
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
         Left            =   7680
         TabIndex        =   94
         Top             =   2220
         Width           =   795
         _ExtentX        =   1402
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
         Left            =   8520
         TabIndex        =   95
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
         Left            =   9420
         TabIndex        =   96
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
         Left            =   9660
         TabIndex        =   97
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
         Left            =   10320
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
         Width           =   990
         _ExtentX        =   1746
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
         TabIndex        =   100
         Top             =   5350
         Width           =   990
         _ExtentX        =   1746
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
         TabIndex        =   101
         Top             =   5580
         Width           =   990
         _ExtentX        =   1746
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
         TabIndex        =   102
         Top             =   5800
         Width           =   990
         _ExtentX        =   1746
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
         Left            =   1260
         TabIndex        =   103
         Top             =   5130
         Width           =   5985
         _ExtentX        =   10557
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
         Left            =   1260
         TabIndex        =   104
         Top             =   5350
         Width           =   5985
         _ExtentX        =   10557
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
         Left            =   1260
         TabIndex        =   105
         Top             =   5580
         Width           =   5985
         _ExtentX        =   10557
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
         Left            =   1260
         TabIndex        =   106
         Top             =   5800
         Width           =   5985
         _ExtentX        =   10557
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
         Left            =   7260
         TabIndex        =   107
         Top             =   5130
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   7260
         TabIndex        =   108
         Top             =   5355
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   7260
         TabIndex        =   109
         Top             =   5580
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   7260
         TabIndex        =   110
         Top             =   5805
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   8200
         TabIndex        =   111
         Top             =   5130
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   8200
         TabIndex        =   112
         Top             =   5355
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   8200
         TabIndex        =   113
         Top             =   5580
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   8200
         TabIndex        =   114
         Top             =   5805
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   8880
         TabIndex        =   115
         Top             =   5130
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   8880
         TabIndex        =   116
         Top             =   5355
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   8880
         TabIndex        =   117
         Top             =   5580
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   8880
         TabIndex        =   118
         Top             =   5805
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         TabIndex        =   119
         Top             =   6030
         Width           =   990
         _ExtentX        =   1746
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
         Left            =   1260
         TabIndex        =   120
         Top             =   6030
         Width           =   5985
         _ExtentX        =   10557
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
         Left            =   7260
         TabIndex        =   121
         Top             =   6030
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   8200
         TabIndex        =   122
         Top             =   6030
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   8880
         TabIndex        =   123
         Top             =   6030
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         TabIndex        =   124
         Top             =   6250
         Width           =   990
         _ExtentX        =   1746
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
         Left            =   1260
         TabIndex        =   125
         Top             =   6250
         Width           =   5985
         _ExtentX        =   10557
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
         Left            =   7260
         TabIndex        =   126
         Top             =   6255
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   8200
         TabIndex        =   127
         Top             =   6255
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   8880
         TabIndex        =   128
         Top             =   6255
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         TabIndex        =   129
         Top             =   6480
         Width           =   990
         _ExtentX        =   1746
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
         Left            =   1260
         TabIndex        =   130
         Top             =   6480
         Width           =   5985
         _ExtentX        =   10557
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
         Left            =   7260
         TabIndex        =   131
         Top             =   6480
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   8200
         TabIndex        =   132
         Top             =   6480
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Left            =   8880
         TabIndex        =   133
         Top             =   6480
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         Index           =   0
         Left            =   9820
         TabIndex        =   139
         Top             =   2880
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         Index           =   1
         Left            =   9820
         TabIndex        =   140
         Top             =   3105
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         Left            =   9820
         TabIndex        =   141
         Top             =   3330
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         Left            =   9820
         TabIndex        =   142
         Top             =   3555
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         Left            =   9820
         TabIndex        =   143
         Top             =   3780
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         Left            =   9820
         TabIndex        =   144
         Top             =   4005
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         Left            =   9820
         TabIndex        =   145
         Top             =   4230
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         Left            =   9820
         TabIndex        =   146
         Top             =   4455
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         Left            =   9820
         TabIndex        =   147
         Top             =   4680
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         Left            =   9820
         TabIndex        =   148
         Top             =   4905
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         Left            =   9820
         TabIndex        =   149
         Top             =   5130
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         Left            =   9840
         TabIndex        =   150
         Top             =   5355
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         Left            =   9820
         TabIndex        =   151
         Top             =   5580
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         Left            =   9820
         TabIndex        =   152
         Top             =   5805
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         Left            =   9820
         TabIndex        =   153
         Top             =   6030
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         Left            =   9820
         TabIndex        =   154
         Top             =   6255
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         Left            =   9820
         TabIndex        =   155
         Top             =   6480
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Begin ReportX.ReportField txtTotalDesc 
         Height          =   225
         Index           =   0
         Left            =   10540
         TabIndex        =   156
         Top             =   2880
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Begin ReportX.ReportField txtTotalDesc 
         Height          =   225
         Index           =   1
         Left            =   10540
         TabIndex        =   157
         Top             =   3105
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Begin ReportX.ReportField txtTotalDesc 
         Height          =   225
         Index           =   2
         Left            =   10540
         TabIndex        =   158
         Top             =   3330
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Begin ReportX.ReportField txtTotalDesc 
         Height          =   225
         Index           =   3
         Left            =   10540
         TabIndex        =   159
         Top             =   3555
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Begin ReportX.ReportField txtTotalDesc 
         Height          =   225
         Index           =   4
         Left            =   10540
         TabIndex        =   160
         Top             =   3780
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Begin ReportX.ReportField txtTotalDesc 
         Height          =   225
         Index           =   5
         Left            =   10540
         TabIndex        =   161
         Top             =   4005
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Begin ReportX.ReportField txtTotalDesc 
         Height          =   225
         Index           =   6
         Left            =   10540
         TabIndex        =   162
         Top             =   4230
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Begin ReportX.ReportField txtTotalDesc 
         Height          =   225
         Index           =   7
         Left            =   10540
         TabIndex        =   163
         Top             =   4455
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Begin ReportX.ReportField txtTotalDesc 
         Height          =   225
         Index           =   8
         Left            =   10540
         TabIndex        =   164
         Top             =   4680
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Begin ReportX.ReportField txtTotalDesc 
         Height          =   225
         Index           =   9
         Left            =   10540
         TabIndex        =   165
         Top             =   4905
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Begin ReportX.ReportField txtTotalDesc 
         Height          =   225
         Index           =   10
         Left            =   10540
         TabIndex        =   166
         Top             =   5130
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Begin ReportX.ReportField txtTotalDesc 
         Height          =   225
         Index           =   11
         Left            =   10540
         TabIndex        =   167
         Top             =   5355
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Begin ReportX.ReportField txtTotalDesc 
         Height          =   225
         Index           =   12
         Left            =   10540
         TabIndex        =   168
         Top             =   5580
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Begin ReportX.ReportField txtTotalDesc 
         Height          =   225
         Index           =   13
         Left            =   10540
         TabIndex        =   169
         Top             =   5805
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Begin ReportX.ReportField txtTotalDesc 
         Height          =   225
         Index           =   14
         Left            =   10540
         TabIndex        =   170
         Top             =   6030
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Begin ReportX.ReportField txtTotalDesc 
         Height          =   225
         Index           =   15
         Left            =   10540
         TabIndex        =   171
         Top             =   6255
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Begin ReportX.ReportField txtTotalDesc 
         Height          =   225
         Index           =   16
         Left            =   10540
         TabIndex        =   172
         Top             =   6480
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   397
         Caption         =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Begin ReportX.ReportField ReportField1 
         Height          =   300
         Left            =   10560
         TabIndex        =   173
         Top             =   2550
         Width           =   870
         _ExtentX        =   1535
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
      Begin ReportX.ReportField ReportField4 
         Height          =   300
         Left            =   9840
         TabIndex        =   174
         Top             =   2550
         Width           =   630
         _ExtentX        =   1111
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
      Begin ReportX.ReportField ReportField11 
         Height          =   270
         Left            =   180
         TabIndex        =   175
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
         TabIndex        =   176
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
      Begin ReportX.ReportField ReportField8 
         Height          =   255
         Left            =   180
         TabIndex        =   177
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
         TabIndex        =   178
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
         TabIndex        =   179
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
         TabIndex        =   180
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
         TabIndex        =   181
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
         TabIndex        =   182
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
         TabIndex        =   183
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
      Begin VB.Line Line9 
         X1              =   10500
         X2              =   10500
         Y1              =   6840
         Y2              =   2520
      End
      Begin VB.Line Line8 
         X1              =   9780
         X2              =   9780
         Y1              =   6840
         Y2              =   2520
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
         Left            =   6840
         TabIndex        =   138
         Top             =   7740
         Width           =   825
      End
      Begin VB.Line Line6 
         X1              =   5220
         X2              =   9240
         Y1              =   7680
         Y2              =   7680
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PEDIDO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9720
         TabIndex        =   137
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PEÇAS/SERVIÇOS"
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
         Left            =   9720
         TabIndex        =   136
         Top             =   540
         Width           =   1635
      End
      Begin VB.Label Label4 
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
         TabIndex        =   135
         Top             =   6900
         Width           =   3810
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
         TabIndex        =   134
         Top             =   7020
         Width           =   3030
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line5 
         X1              =   8840
         X2              =   8840
         Y1              =   6840
         Y2              =   2520
      End
      Begin VB.Line Line4 
         X1              =   8160
         X2              =   8160
         Y1              =   6840
         Y2              =   2520
      End
      Begin VB.Line Line3 
         X1              =   7260
         X2              =   7260
         Y1              =   6840
         Y2              =   2520
      End
      Begin VB.Line Line2 
         X1              =   1200
         X2              =   1200
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
         Left            =   9360
         Top             =   6840
         Width           =   2115
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         Height          =   4335
         Left            =   120
         Top             =   2520
         Width           =   11355
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         Height          =   1395
         Left            =   9300
         Top             =   1140
         Width           =   2175
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         Height          =   1395
         Left            =   6960
         Top             =   1140
         Width           =   2355
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   1395
         Left            =   120
         Top             =   1140
         Width           =   6855
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
Attribute VB_Name = "REL_OS_PedidoVista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim Cont As Integer
Dim wValorFormatado As Currency
Dim sSQL As String
Dim r As ADODB.Recordset
Dim rPd As ADODB.Recordset
Dim rCl As ADODB.Recordset
Dim rIt As ADODB.Recordset
Dim rTotais As ADODB.Recordset
Dim rPc As ADODB.Recordset
Dim rFu As ADODB.Recordset
Dim rOS As ADODB.Recordset
'Dim oCfg As ConfigItem
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


    Set rOS = dbData.OpenRecordset("SELECT cod_pedido, cod_cliente, cod_funcionario FROM os WHERE (cod_os = " & vCodOS & ");")
    Set rPd = dbData.OpenRecordset("SELECT * FROM pedidos WHERE (cod_pedido = " & Pedido & ");")
    Set rCl = dbData.OpenRecordset("SELECT * FROM cliente WHERE (codigo = " & rOS("cod_cliente") & ");")
    
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
            txtQuantServicos.Caption = Format(ValidateNull(rTotais("vSomaQuantServ")), "000")
            txtTotalServicos.Caption = FormatNumber(ValidateNull(rTotais("vSomaValorServ")), 2)
        Else
            txtQuantServicos.Caption = Format(0, "000")
            txtTotalServicos.Caption = FormatNumber(0, 2)
        End If
        
        txtQuantGeral.Caption = Format(CInt(txtQuantPecas.Caption) + CInt(txtQuantServicos.Caption), "000")
        txtTotalPecasServicos.Caption = FormatNumber(CCur(txtTotalPecas.Caption) + CCur(txtTotalServicos.Caption), 2)
   End If
   
   
   'mostrar produtos no grid
   If rIt.EOF Then
        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
           sSQL = "SELECT 'SERVIÇO' AS tipo_item, descricao, quantidade, preco, subtotal, desconto, total, codigo, '' as varFabricante, '', '', '', '', '', '', desconto, total " & _
           "FROM OS_Servicos_Auto WHERE (cod_os = " & vCodOS & ")"
        ElseIf vTipoOS = "Recapadora" Then
            sSQL = "SELECT 'SERVIÇO' AS tipo_item, descricao, quantidade, preco, subtotal, desconto, total, codigo, TIPO as var_TipoPneu, SERIE as var_serie, FOGO as var_fogo, ARO as var_aro, BANDA as var_banda, DOTE as var_dote, MEDIDA as var_medida, FABRICANTE as var_fabricante, desconto, total " & _
            "FROM OS_servicos_recapadora WHERE (cod_os = " & vCodOS & ")"
        ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
           sSQL = "SELECT 'SERVIÇO' AS tipo_item, descricao, quantidade, preco, subtotal, desconto, total, codigo, '' as varFabricante, '', '', '', '', '', '', desconto, total " & _
           "FROM OS_Servicos_Auto WHERE (cod_os = " & vCodOS & ")"
        ElseIf vTipoOS = "Comunicação Visual" Then
           sSQL = "SELECT 'SERVIÇO' AS tipo_item, descricao, quantidade, preco, subtotal, desconto, total, codigo, '' as varFabricante, '', '', '', '', '', '', desconto, total " & _
           "FROM OS_Servicos_Auto WHERE (cod_os = " & vCodOS & ")"
        End If
   
   'aqui
   Else
        If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
             sSQL = "SELECT 'PRODUTO' AS tipo_item, produtos.descricao, pedidos_itens.quantidade, pedidos_itens.preco, subtotal, desconto, total, pedidos_itens.codigo, produtos.Fabricante as varFabricante FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto WHERE (pedidos_itens.cod_pedido = " & Pedido & ")"
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
                   sSQL = sSQL & "SELECT 'SERVIÇO' AS tipo_item, descricao, quantidade, preco, subtotal, desconto, total, codigo, '' " & _
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
   'Debug.Print sSQL
   
   Set rIt = dbData.OpenRecordset(sSQL)
   
   Set rPc = dbData.OpenRecordset("SELECT * FROM parcelas WHERE (cod_pedido = " & Pedido & ") ORDER BY numero;")
   Set rOS = dbData.OpenRecordset("SELECT * FROM OS WHERE (COD_pedido = " & Pedido & ");")
   Set rFu = dbData.OpenRecordset("SELECT * FROM funcionario WHERE (codigo = " & rOS("COD_FUNCIONARIO") & ");")
   
   'rfData.Caption = "Vencimento: " & Format(rPc("data"), "dd") & " de " & Format(rPc("data"), "mmmm") & " de " & Format(rPc("data"), "yyyy")
   
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
         frValorParc(Cont).Caption = FormatNumber(rPc("valor"), 2)
      End If
      Cont = Cont + 1
      rPc.MoveNext
   Loop
   
   txtNumero.Caption = "Nº " & Format(vCodOS, "000000")
   
   'DADOS DO CLIENTE
   txtCliente.Caption = rCl("nome")
   txtEnd.Caption = rCl("endereco") & " - " & rCl("bairro")
   txtRef.Caption = ValidateNull(rCl("ponto_de_referencia"))
   txtCidade.Caption = rCl("cidade") & "-" & rCl("estado") & " TEL: " & rCl("telefone1")
   txtCPF.Caption = ValidateNull(rCl("cpf"))
   txtRG.Caption = ValidateNull(rCl("RG"))
   
   'DADOS DO PEDIDO
   txtData.Caption = String(1, " ") + Format(rPd("data_compra"), "dd/mm/yy")
   txtVendedor.Caption = rFu("nome")
   txtVenda.Caption = UCase(ValidateNull(rPd("tipo_pagamento")))
   'txtPagamento.Caption = rsPedidos!PAGAMENTO

   If rPd("tipo_pagamento") = "À Vista" Then
      txtData.Mostrar = True
      txtVendedor.Mostrar = True
      txtVenda.Mostrar = True
      'txtPagamento.Mostrar = True
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

'totais
If IsNull(rPd("subtotal")) Then
   txtSubtotal.Caption = Format(OS_Recapadora.txtTotalGeral.Text, ocMONEY)
Else
   txtSubtotal.Caption = Format(rPd("subtotal"), ocMONEY)
End If
  
   If IsNull(rPd("valor_desc")) Then
      wValorFormatado = "0,00"
   Else
      If rPd("tipo_desc") = "R" Then
         wValorFormatado = Format(rPd("ValorDescReal"), ocMONEY)
      Else
        If rPd("valor_desc") = "0" Then
            wValorFormatado = Format(rPd("ValorDescReal"), ocMONEY)
        Else
            'wValorFormatado = Format(rPd("valor_desc"), ocMONEY) & "%"
            'wValorFormatado = FormatNumber(rPd("valor_desc"), 2) & "%"
            wValorFormatado = Format(rPd("ValorDescReal"), ocMONEY)
        End If
      End If
   End If

   txtDesconto.Caption = Format(wValorFormatado, ocMONEY)
  
   txtTotal.Caption = String(1, " ") + Format(rPd("total"), ocMONEY)
   
   
   'INSIRO OS ITENS
   If Not rIt.BOF Then rIt.MoveFirst
   
   'If Not rIt.EOF Then rIt.MoveLast
   'If Not rIt.BOF Then rIt.MoveFirst
   
   'Relatorio.NumeroRegistros = (rIt.RecordCount Mod 17) + ((rIt.RecordCount \ 17) * 17)
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

Private Sub Relatorio_IniciarSecao(ByVal Secao As ReportX.TSecao, ByVal Ordem As Byte)
   If Secao = secDetalhe Then
      'produtos e serviços
      For i = 0 To 16
         txtDesc(i).Caption = ""
         txtCodProd(i).Caption = ""
         txtQuant(i).Caption = ""
         txtUnit(i).Caption = ""
         txtTot(i).Caption = ""
      Next
      
      i = 0
      Do While Not rIt.EOF
         If i >= 17 Then i = 0
         txtCodProd(i).Caption = rIt("tipo_item")
         
         If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Then
            txtDesc(i).Caption = String(1, " ") + ValidateNull(rIt("descricao")) & " | " & rIt("varFabricante")
         ElseIf vTipoOS = "Recapadora" Then
            If rIt("tipo_item") = "SERVIÇO" Then
                txtDesc(i).Caption = String(1, " ") + ValidateNull(rIt("descricao")) & " | " & rIt("var_TipoPneu") & " | " & rIt("var_serie") & " | " & rIt("var_fogo") & " | " & rIt("var_aro") & " | " & rIt("var_banda") & " | " & rIt("var_dote") & " | " & rIt("var_medida") & " | " & rIt("var_fabricante") & " "
            Else
                txtDesc(i).Caption = String(1, " ") + ValidateNull(rIt("descricao")) & " | " & rIt("varFabricante")
            End If
         ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then
            txtDesc(i).Caption = String(1, " ") + ValidateNull(rIt("descricao")) & " | " & rIt("varFabricante")
         ElseIf vTipoOS = "Comunicação Visual" Then
            txtDesc(i).Caption = String(1, " ") + ValidateNull(rIt("descricao")) & " | " & rIt("varFabricante")
         End If
         txtQuant(i).Caption = rIt("quantidade")
         txtUnit(i).Caption = Format(rIt("preco"), ocMONEY)
         txtTot(i).Caption = Format(rIt("subtotal"), ocMONEY)
         txtDesco(i).Caption = Format(rIt("desconto"), ocMONEY)
         txtTotalDesc(i).Caption = Format(rIt("total"), ocMONEY)
         
         i = i + 1
         rIt.MoveNext
      Loop
   End If
End Sub

