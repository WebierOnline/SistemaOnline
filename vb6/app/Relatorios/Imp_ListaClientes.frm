VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.ocx"
Begin VB.Form Imp_ListaClientes 
   Caption         =   "Form1"
   ClientHeight    =   3105
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   54.769
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   268.817
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   2580
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Orientacao      =   2
      Titulo          =   ""
      NomeImpressora  =   "IMPRESSORA1"
      Registrado      =   0   'False
   End
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   2235
      Left            =   0
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   3942
      Tipo            =   2
      Begin ReportX.ReportField rf2 
         Height          =   300
         Left            =   3720
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   13
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
         TabIndex        =   14
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
      Begin ReportX.ReportField txtDHead 
         Height          =   390
         Left            =   60
         TabIndex        =   15
         Top             =   1560
         Width           =   15315
         _ExtentX        =   27014
         _ExtentY        =   688
         Caption         =   "LISTA DE CLIENTES"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borda           =   9
         BordaEstilo     =   1
         BackColor       =   14737632
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ULT. COMPRA"
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
         Left            =   14100
         TabIndex        =   24
         Top             =   1980
         Width           =   1275
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENDEDOR"
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
         Left            =   12780
         TabIndex        =   22
         Top             =   1980
         Width           =   1050
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CIDADE"
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
         Left            =   9780
         TabIndex        =   20
         Top             =   1980
         Width           =   705
      End
      Begin VB.Label Label1 
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
         Left            =   60
         TabIndex        =   18
         Top             =   1980
         Width           =   750
      End
      Begin VB.Shape Shape1 
         Height          =   1335
         Left            =   60
         Shape           =   4  'Rounded Rectangle
         Top             =   60
         Width           =   15315
      End
      Begin VB.Image imgLogo 
         Height          =   1215
         Left            =   180
         Stretch         =   -1  'True
         Top             =   120
         Width           =   3315
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BAIRRO"
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
         Left            =   8280
         TabIndex        =   10
         Top             =   1980
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CÓD"
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
         Left            =   900
         TabIndex        =   4
         Top             =   1980
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOME"
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
         Left            =   1500
         TabIndex        =   3
         Top             =   1980
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ENDEREÇO"
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
         Left            =   5040
         TabIndex        =   2
         Top             =   1980
         Width           =   1035
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TELEFONE"
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
         Left            =   11580
         TabIndex        =   1
         Top             =   1980
         Width           =   975
      End
   End
   Begin ReportX.ReportSection ReportSection2 
      Align           =   1  'Align Top
      Height          =   240
      Left            =   0
      Top             =   2235
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   423
      Begin ReportX.ReportField ReportField5 
         Height          =   195
         Left            =   8280
         TabIndex        =   9
         Top             =   0
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   344
         Campo           =   "BAIRRO"
         Caption         =   "ReportField2"
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
         Left            =   900
         TabIndex        =   5
         Top             =   0
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   344
         Campo           =   "CODIGO"
         Formato         =   "0000"
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
         Left            =   11580
         TabIndex        =   6
         Top             =   0
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   344
         Campo           =   "TELEFONE1"
         Caption         =   "ReportField2"
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
         Left            =   5040
         TabIndex        =   7
         Top             =   0
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   344
         Campo           =   "ENDERECO"
         Caption         =   "ReportField2"
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
         Left            =   1500
         TabIndex        =   8
         Top             =   0
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   344
         Campo           =   "NOME"
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
         Left            =   60
         TabIndex        =   17
         Top             =   0
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   344
         Campo           =   "var_STATUS"
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
         Left            =   9780
         TabIndex        =   19
         Top             =   0
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   344
         Campo           =   "CIDADE"
         Caption         =   "ReportField2"
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
         Left            =   12780
         TabIndex        =   21
         Top             =   0
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   344
         Campo           =   "VENDEDOR"
         Caption         =   "ReportField2"
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
      Begin ReportX.ReportField ReportField10 
         Height          =   195
         Left            =   14100
         TabIndex        =   23
         Top             =   0
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   344
         Campo           =   "Ultima_Compra"
         Formato         =   "DD/MM/YY"
         Caption         =   "ReportField2"
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
   Begin ReportX.ReportSection ReportSection3 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      Top             =   2475
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   688
      Tipo            =   7
      Begin ReportX.ReportField dfQuant 
         Height          =   195
         Left            =   3120
         TabIndex        =   25
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
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
      Begin ReportX.ReportField ReportField6 
         Height          =   240
         Left            =   60
         TabIndex        =   16
         Top             =   60
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   423
         Campo           =   "=Página [Pagina] de [Paginas]"
         Formula         =   -1  'True
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
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         X1              =   15000
         X2              =   60
         Y1              =   0
         Y2              =   0
      End
   End
End
Attribute VB_Name = "Imp_ListaClientes"
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
