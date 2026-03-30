VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.Ocx"
Begin VB.Form REL_OS_Entrada_Celular 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impressão - Recibo"
   ClientHeight    =   9195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12075
   Icon            =   "REL_OS_Entrada_Celular.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   162.19
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   212.99
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   60
      TabIndex        =   0
      Top             =   8220
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Divisao         =   10
      Regua           =   -1  'True
      Titulo          =   ""
      Registrado      =   0   'False
   End
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   8115
      Left            =   0
      Top             =   0
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   14314
      Ordem           =   1
      Begin ReportX.ReportField ReportField15 
         Height          =   420
         Left            =   9180
         TabIndex        =   44
         Top             =   5520
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   741
         Caption         =   ""
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
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField14 
         Height          =   360
         Left            =   9180
         TabIndex        =   43
         Top             =   2940
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         Caption         =   "SENHA:"
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
      End
      Begin ReportX.ReportField lblSituacao 
         Height          =   195
         Index           =   4
         Left            =   7140
         TabIndex        =   39
         Top             =   6120
         Width           =   1875
         _ExtentX        =   3307
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
      Begin ReportX.ReportField lblSituacao 
         Height          =   195
         Index           =   3
         Left            =   7140
         TabIndex        =   38
         Top             =   5880
         Width           =   1875
         _ExtentX        =   3307
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
      Begin ReportX.ReportField lblSituacao 
         Height          =   195
         Index           =   2
         Left            =   7140
         TabIndex        =   37
         Top             =   5640
         Width           =   1875
         _ExtentX        =   3307
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
      Begin ReportX.ReportField lblSituacao 
         Height          =   195
         Index           =   1
         Left            =   7140
         TabIndex        =   36
         Top             =   5400
         Width           =   1875
         _ExtentX        =   3307
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
      Begin ReportX.ReportField lblSituacao 
         Height          =   195
         Index           =   0
         Left            =   7140
         TabIndex        =   35
         Top             =   5160
         Width           =   1875
         _ExtentX        =   3307
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
      Begin ReportX.ReportField ReportField9 
         Height          =   1380
         Left            =   7020
         TabIndex        =   40
         Top             =   5040
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   2434
         Caption         =   ""
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
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField12 
         Height          =   360
         Left            =   7020
         TabIndex        =   41
         Top             =   4680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         Caption         =   "SITUAÇÕES"
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
      End
      Begin ReportX.ReportField txtEquipamento 
         Height          =   420
         Left            =   180
         TabIndex        =   10
         Top             =   3300
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   741
         Caption         =   ""
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
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtModelo 
         Height          =   420
         Left            =   4860
         TabIndex        =   33
         Top             =   3300
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   741
         Caption         =   ""
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
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField10 
         Height          =   360
         Left            =   4920
         TabIndex        =   34
         Top             =   2940
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         Caption         =   "MODELO"
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
      End
      Begin ReportX.ReportField lblAcessorio 
         Height          =   195
         Index           =   4
         Left            =   7140
         TabIndex        =   32
         Top             =   4380
         Width           =   1875
         _ExtentX        =   3307
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
      Begin ReportX.ReportField lblAcessorio 
         Height          =   195
         Index           =   3
         Left            =   7140
         TabIndex        =   31
         Top             =   4140
         Width           =   1875
         _ExtentX        =   3307
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
      Begin ReportX.ReportField lblAcessorio 
         Height          =   195
         Index           =   2
         Left            =   7140
         TabIndex        =   30
         Top             =   3900
         Width           =   1875
         _ExtentX        =   3307
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
      Begin ReportX.ReportField lblAcessorio 
         Height          =   195
         Index           =   1
         Left            =   7140
         TabIndex        =   29
         Top             =   3660
         Width           =   1875
         _ExtentX        =   3307
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
      Begin ReportX.ReportField lblAcessorio 
         Height          =   195
         Index           =   0
         Left            =   7140
         TabIndex        =   28
         Top             =   3420
         Width           =   1875
         _ExtentX        =   3307
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
      Begin ReportX.ReportField ReportField8 
         Height          =   180
         Left            =   540
         TabIndex        =   27
         Top             =   6720
         Width           =   10020
         _ExtentX        =   17674
         _ExtentY        =   318
         Formato         =   "##,##0.00"
         Caption         =   $"REL_OS_Entrada_Celular.frx":030A
         WordWrap        =   -1  'True
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
      Begin ReportX.ReportField ReportField6 
         Height          =   180
         Left            =   180
         TabIndex        =   26
         Top             =   6540
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   318
         Formato         =   "##,##0.00"
         Caption         =   $"REL_OS_Entrada_Celular.frx":03B0
         WordWrap        =   -1  'True
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
      Begin ReportX.ReportField ReportField2 
         Height          =   300
         Left            =   1260
         TabIndex        =   25
         Top             =   7500
         Width           =   4260
         _ExtentX        =   7514
         _ExtentY        =   529
         Formato         =   "##,##0.00"
         Caption         =   "RECEPCIONISTA"
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
         Borda           =   1
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField5 
         Height          =   1380
         Left            =   7020
         TabIndex        =   15
         Top             =   3300
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   2434
         Caption         =   ""
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
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField7 
         Height          =   360
         Left            =   7020
         TabIndex        =   16
         Top             =   2940
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         Caption         =   "ACESSÓRIOS:"
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
      End
      Begin ReportX.ReportField txtEntrada 
         Height          =   300
         Left            =   6420
         TabIndex        =   14
         Top             =   7500
         Width           =   4260
         _ExtentX        =   7514
         _ExtentY        =   529
         Formato         =   "##,##0.00"
         Caption         =   "CLIENTE"
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
         Borda           =   1
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtFuncionario 
         Height          =   360
         Left            =   1260
         TabIndex        =   13
         Top             =   7140
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   635
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
      End
      Begin ReportX.ReportField txtDescricao 
         Height          =   2340
         Left            =   180
         TabIndex        =   12
         Top             =   4080
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   4128
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
         Borda           =   15
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtMarca 
         Height          =   420
         Left            =   2820
         TabIndex        =   11
         Top             =   3300
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   741
         Caption         =   ""
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
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtCliente 
         Height          =   420
         Left            =   180
         TabIndex        =   9
         Top             =   2520
         Width           =   6600
         _ExtentX        =   11642
         _ExtentY        =   741
         Caption         =   ""
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
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDataEntrada 
         Height          =   420
         Left            =   6840
         TabIndex        =   8
         Top             =   2520
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   741
         Caption         =   ""
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
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField11 
         Height          =   360
         Left            =   180
         TabIndex        =   7
         Top             =   3720
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   635
         Caption         =   "DESCRIÇÃO DO CLIENTE:"
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
      End
      Begin ReportX.ReportField rfMarca 
         Height          =   360
         Left            =   2820
         TabIndex        =   6
         Top             =   2940
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         Caption         =   "MARCA"
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
      End
      Begin ReportX.ReportField ReportField4 
         Height          =   360
         Left            =   6840
         TabIndex        =   5
         Top             =   2160
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         Caption         =   "ENTRADA:"
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
      End
      Begin ReportX.ReportField rfEquipamento 
         Height          =   360
         Left            =   180
         TabIndex        =   4
         Top             =   2940
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   635
         Caption         =   "EQUIPAMENTO:"
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
      End
      Begin ReportX.ReportField ReportField1 
         Height          =   360
         Left            =   180
         TabIndex        =   3
         Top             =   2160
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   635
         Caption         =   "CLIENTE:"
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
      End
      Begin ReportX.ReportField txthead 
         Height          =   450
         Left            =   90
         TabIndex        =   1
         Top             =   1500
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   794
         Caption         =   "ENTRADA NA ASSISTÊNCIA"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField rf2 
         Height          =   300
         Left            =   3720
         TabIndex        =   17
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
         TabIndex        =   18
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
         TabIndex        =   19
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
         TabIndex        =   20
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
      Begin ReportX.ReportField ReportField41 
         Height          =   465
         Index           =   0
         Left            =   9180
         TabIndex        =   21
         Top             =   1500
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   820
         Caption         =   "OS:"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtOS 
         Height          =   465
         Left            =   10020
         TabIndex        =   22
         Top             =   1500
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   820
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
      End
      Begin ReportX.ReportField txtSaida 
         Height          =   420
         Left            =   8940
         TabIndex        =   23
         Top             =   2520
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   741
         Caption         =   ""
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
         Borda           =   15
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField3 
         Height          =   360
         Left            =   8940
         TabIndex        =   24
         Top             =   2160
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   635
         Caption         =   "SAÍDA(PREVISÃO):"
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
      End
      Begin ReportX.ReportField ReportField13 
         Height          =   2160
         Left            =   9180
         TabIndex        =   42
         Top             =   3300
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   3810
         Caption         =   ""
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
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField Ret 
         Height          =   5910
         Left            =   60
         TabIndex        =   2
         Top             =   1980
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   10425
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   1935
         Left            =   9300
         Picture         =   "REL_OS_Entrada_Celular.frx":045B
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   1950
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
         Width           =   11415
      End
   End
End
Attribute VB_Name = "REL_OS_Entrada_Celular"
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
   rf4.Caption = "CNPJ: " & r("cnpj") & " - IE: " & r("ie") & " - CEL.: " & r("CELULAR")
   
   If Not IsNull(r("caminho")) Then
      If Dir$(r("caminho")) <> "" Then Set imgLogo.Picture = LoadPicture(r("caminho"))
   End If
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   Exit Sub
   
TrataErro:
End Sub



Public Sub Preencher_Situacao(Pedido As Long)
   Dim i As Integer
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT situacao, cod_os FROM OS_Situacao_Auto WHERE (cod_os = " & Pedido & ");"
   Set r = dbData.OpenRecordset(sSQL)
   
   'limpar
   For i = 0 To 4
      lblSituacao(i).Caption = ""
   Next
   
   'carga
   i = 0
   Do While Not r.EOF
      If i < 5 Then lblSituacao(i).Caption = r("situacao")
      i = i + 1
      r.MoveNext
   Loop
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub
Public Sub Preencher_Acessorios(Pedido As Long)
Dim i As Integer
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT acessorio, cod_os FROM OS_acessorios_Auto WHERE (cod_os = " & Pedido & ");"
Set r = dbData.OpenRecordset(sSQL)

'limpar
For i = 0 To 4
   lblAcessorio(i).Caption = ""
Next

'carga
i = 0
Do While Not r.EOF
   If i < 5 Then lblAcessorio(i).Caption = r("acessorio")
   i = i + 1
   r.MoveNext
Loop

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

