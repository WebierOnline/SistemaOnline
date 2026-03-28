VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.ocx"
Begin VB.Form REL_ContratoAluguel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CONTRATO"
   ClientHeight    =   14940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12570
   Icon            =   "REL_ContratoAluguel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   263.525
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   221.721
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   0
      TabIndex        =   31
      Top             =   8760
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Titulo          =   ""
      NomeImpressora  =   "IMPRESSORA1"
      Registrado      =   0   'False
   End
   Begin ReportX.ReportSection ReportSection3 
      Align           =   1  'Align Top
      Height          =   9645
      Left            =   0
      Top             =   4170
      Width           =   12570
      _ExtentX        =   22172
      _ExtentY        =   17013
      Tipo            =   7
      Ordem           =   1
      Begin ReportX.ReportField ReportField6 
         Height          =   165
         Left            =   60
         TabIndex        =   7
         Top             =   9420
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   291
         Campo           =   "=Página [Pagina] de [Paginas]"
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ATENÇÃO"
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
         Left            =   9420
         TabIndex        =   58
         Top             =   8640
         Width           =   1065
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Será cobrado um valor adicional por cada hora a mais que passar do horario de entrega do equipamento."
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   8520
         TabIndex        =   57
         Top             =   8820
         Width           =   2895
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape1 
         Height          =   915
         Left            =   8400
         Top             =   8580
         Width           =   3075
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "__________________________________________"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5055
         TabIndex        =   54
         Top             =   7380
         Width           =   5985
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "__________________________________________"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   53
         Top             =   7380
         Width           =   5985
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2._________________________________________"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   52
         Top             =   9060
         Width           =   4545
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1._________________________________________"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   51
         Top             =   8820
         Width           =   4545
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TESTEMUNHAS:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   50
         Top             =   8580
         Width           =   4545
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOCATÁRIO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5580
         TabIndex        =   49
         Top             =   8100
         Width           =   4725
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ: 00.000.000/0000-00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5160
         TabIndex        =   48
         Top             =   7860
         Width           =   5385
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BELTRANO"
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
         Left            =   5760
         TabIndex        =   47
         Top             =   7620
         Width           =   4485
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOCADOR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   900
         TabIndex        =   46
         Top             =   8100
         Width           =   4665
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ: 00.665.664/0001-70"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1035
         TabIndex        =   45
         Top             =   7860
         Width           =   4545
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "R&S CONSTRUÇÕES E SERVIÇOS"
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
         Left            =   990
         TabIndex        =   44
         Top             =   7620
         Width           =   4515
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Uruçuí-PI XX de XXXX de 2020."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   43
         Top             =   6900
         Width           =   11415
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"REL_ContratoAluguel.frx":030A
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   180
         TabIndex        =   42
         Top             =   6300
         Width           =   11325
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"REL_ContratoAluguel.frx":03AA
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   180
         TabIndex        =   41
         Top             =   5760
         Width           =   11415
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"REL_ContratoAluguel.frx":047A
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   180
         TabIndex        =   40
         Top             =   5220
         Width           =   11415
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5. RESCISÃO"
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
         Left            =   180
         TabIndex        =   39
         Top             =   4920
         Width           =   11415
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"REL_ContratoAluguel.frx":0515
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   180
         TabIndex        =   38
         Top             =   4260
         Width           =   11415
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4. PRAZO DE VIGÊNCIA DO CONTRATO"
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
         Left            =   180
         TabIndex        =   37
         Top             =   3960
         Width           =   11415
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"REL_ContratoAluguel.frx":05A8
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   180
         TabIndex        =   36
         Top             =   3360
         Width           =   11415
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3.3 O locador deverá manter o equipamento em perfeitas condições de uso."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   35
         Top             =   3060
         Width           =   11415
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"REL_ContratoAluguel.frx":065B
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   180
         TabIndex        =   34
         Top             =   2460
         Width           =   11325
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3.1 A manutenção do equipamento, objeto do presente contrato, é de total responsabilidade do locador."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   33
         Top             =   2160
         Width           =   11415
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2. ALUGUÉIS MENSAIS E REAJUSTES"
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
         Left            =   180
         TabIndex        =   32
         Top             =   720
         Width           =   11415
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"REL_ContratoAluguel.frx":0707
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   180
         TabIndex        =   16
         Top             =   1020
         Width           =   11280
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3. MANUTENÇÃO, ASSISTÊNCIA TÉCNICA, TRANSPORTE E DEVOLUÇÃO"
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
         Left            =   180
         TabIndex        =   15
         Top             =   1860
         Width           =   11295
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"REL_ContratoAluguel.frx":0858
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   180
         TabIndex        =   14
         Top             =   60
         Width           =   11415
         WordWrap        =   -1  'True
      End
   End
   Begin ReportX.ReportSection Det 
      Align           =   1  'Align Top
      Height          =   315
      Left            =   0
      Top             =   3855
      Width           =   12570
      _ExtentX        =   22172
      _ExtentY        =   556
      Begin ReportX.ReportField dfStatus 
         Height          =   225
         Left            =   6900
         TabIndex        =   8
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   397
         Campo           =   "SUBTOTAL"
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
         Borda           =   15
      End
      Begin ReportX.ReportField dfVenc 
         Height          =   225
         Left            =   60
         TabIndex        =   3
         Top             =   0
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   397
         Campo           =   "varCodEquip"
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
         Borda           =   15
      End
      Begin ReportX.ReportField dfClient 
         Height          =   225
         Left            =   780
         TabIndex        =   4
         Top             =   0
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   397
         Campo           =   "DESCRICAO"
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
         Borda           =   15
      End
      Begin ReportX.ReportField dfValor 
         Height          =   225
         Left            =   3720
         TabIndex        =   5
         Top             =   0
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         Campo           =   "MODELO"
         Caption         =   ""
         TipoCampo       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
      End
      Begin ReportX.ReportField dfPgto 
         Height          =   225
         Left            =   6360
         TabIndex        =   6
         Top             =   0
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   397
         Campo           =   "vQuant"
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
         Borda           =   15
      End
      Begin ReportX.ReportField dfHav 
         Height          =   225
         Left            =   4680
         TabIndex        =   9
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   397
         Campo           =   "FABRICANTE"
         Caption         =   ""
         TipoCampo       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   15
      End
      Begin ReportX.ReportField ReportField8 
         Height          =   225
         Left            =   5580
         TabIndex        =   10
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   397
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
         Borda           =   15
      End
      Begin ReportX.ReportField ReportField1 
         Height          =   225
         Left            =   7680
         TabIndex        =   11
         Top             =   0
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   397
         Campo           =   "DESCONTO"
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
         Borda           =   15
      End
      Begin ReportX.ReportField ReportField2 
         Height          =   225
         Left            =   8220
         TabIndex        =   17
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   397
         Campo           =   "VALOR_FINAL"
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
         Borda           =   15
      End
      Begin ReportX.ReportField ReportField3 
         Height          =   225
         Left            =   9000
         TabIndex        =   18
         Top             =   0
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   397
         Campo           =   "DATA_INICIO"
         Formato         =   "DD/MM/YY"
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
         Borda           =   15
      End
      Begin ReportX.ReportField ReportField4 
         Height          =   225
         Left            =   10320
         TabIndex        =   19
         Top             =   0
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   397
         Campo           =   "DATA_FINAL"
         Formato         =   "DD/MM/YY"
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
         Borda           =   15
      End
      Begin ReportX.ReportField ReportField5 
         Height          =   225
         Left            =   9780
         TabIndex        =   55
         Top             =   0
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   397
         Campo           =   "HORA_INICIO"
         Formato         =   "hh:mm"
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
         Borda           =   15
      End
      Begin ReportX.ReportField ReportField7 
         Height          =   225
         Left            =   11100
         TabIndex        =   56
         Top             =   0
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   397
         Campo           =   "HORA_FINAL"
         Formato         =   "hh:mm"
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
         Borda           =   15
      End
   End
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   3855
      Left            =   0
      Top             =   0
      Width           =   12570
      _ExtentX        =   22172
      _ExtentY        =   6800
      Tipo            =   2
      Begin VB.Image imgLogo 
         Height          =   1095
         Left            =   180
         Stretch         =   -1  'True
         Top             =   60
         Width           =   2355
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TERMINO"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   10320
         TabIndex        =   30
         Top             =   3660
         Width           =   780
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INICIO"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   9060
         TabIndex        =   29
         Top             =   3660
         Width           =   585
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "V.FINAL"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   8280
         TabIndex        =   28
         Top             =   3660
         Width           =   705
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESC"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   7740
         TabIndex        =   27
         Top             =   3660
         Width           =   435
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   7080
         TabIndex        =   26
         Top             =   3660
         Width           =   540
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QTDE"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   6360
         TabIndex        =   25
         Top             =   3660
         Width           =   450
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VALOR"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   5580
         TabIndex        =   24
         Top             =   3660
         Width           =   585
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FAB."
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   4680
         TabIndex        =   23
         Top             =   3660
         Width           =   420
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MODELO"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   3720
         TabIndex        =   22
         Top             =   3660
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EQUIPAMENTO"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   840
         TabIndex        =   21
         Top             =   3660
         Width           =   1245
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CÓD."
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   60
         TabIndex        =   20
         Top             =   3660
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"REL_ContratoAluguel.frx":0910
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   13
         Top             =   3060
         Width           =   11355
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. OBJETO E VALOR "
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
         Left            =   120
         TabIndex        =   12
         Top             =   2760
         Width           =   7275
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"REL_ContratoAluguel.frx":09AF
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   180
         TabIndex        =   2
         Top             =   1320
         Width           =   11340
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblContrato 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONTRATO N°XX/2020 "
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
         Left            =   9360
         TabIndex        =   1
         Top             =   960
         Width           =   2100
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONTRATO DE LOCAÇÃO DE EQUIPAMENTOS"
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
         Left            =   3420
         TabIndex        =   0
         Top             =   360
         Width           =   6540
      End
   End
End
Attribute VB_Name = "REL_ContratoAluguel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Relatorio_IniciarSecao(ByVal Secao As ReportX.TSecao, ByVal Ordem As Byte)
    If Secao = secDetalhe Then
        ' Apenas para aumentar a distancia entre os paragrafos.
        Det.Height = Det.Height
    End If
End Sub


