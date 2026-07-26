VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.Ocx"
Begin VB.Form REL_OS_Garantia 
   Caption         =   "Form1"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   ScaleHeight     =   111.125
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   216.959
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain ReportMain1 
      Height          =   480
      Left            =   8220
      TabIndex        =   1
      Top             =   5700
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Titulo          =   ""
      TipoColunas     =   1
      MargemEsquerda  =   1
      MargemDireita   =   1
      Registrado      =   0   'False
      Visualizar      =   0   'False
   End
   Begin ReportX.ReportSection SecaoHeader 
      Align           =   1  'Align Top
      Height          =   1740
      Left            =   0
      Top             =   1035
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   3069
      Tipo            =   2
      Begin ReportX.ReportField txtDHead 
         Height          =   420
         Left            =   630
         TabIndex        =   2
         Top             =   30
         Width           =   11610
         _ExtentX        =   21484
         _ExtentY        =   741
         Caption         =   "TERMO DE GARANTIA DE SERVIÇOS E PEÇAS"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlturaLivre     =   -1  'True
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField ReportField1 
         Height          =   270
         Left            =   630
         TabIndex        =   3
         Top             =   570
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   476
         Caption         =   "Nº DA OS:"
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
      Begin ReportX.ReportField txtNumOSGar 
         Height          =   270
         Left            =   1980
         TabIndex        =   4
         Top             =   570
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   476
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
      Begin ReportX.ReportField ReportField2 
         Height          =   270
         Left            =   8070
         TabIndex        =   5
         Top             =   570
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   476
         Caption         =   "DATA DE SAÍDA:"
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
      Begin ReportX.ReportField txtDataSaidaGar 
         Height          =   270
         Left            =   10020
         TabIndex        =   6
         Top             =   570
         Width           =   730
         _ExtentX        =   2646
         _ExtentY        =   476
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
      Begin ReportX.ReportField ReportField3 
         Height          =   270
         Left            =   630
         TabIndex        =   7
         Top             =   870
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   476
         Caption         =   "CLIENTE:"
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
      Begin ReportX.ReportField txtClienteGar 
         Height          =   270
         Left            =   1980
         TabIndex        =   8
         Top             =   870
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   476
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
      Begin ReportX.ReportField ReportField4 
         Height          =   270
         Left            =   8070
         TabIndex        =   9
         Top             =   870
         Width           =   1350
         _ExtentX        =   2381
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
         AlinhamentoVertical=   1
      End
      Begin ReportX.ReportField txtDocumentoGar 
         Height          =   270
         Left            =   9420
         TabIndex        =   10
         Top             =   870
         Width           =   1330
         _ExtentX        =   3704
         _ExtentY        =   476
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
      Begin ReportX.ReportField lblVeiculoGar 
         Height          =   270
         Left            =   630
         TabIndex        =   11
         Top             =   1170
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   476
         Caption         =   "VEÍCULO:"
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
      Begin ReportX.ReportField txtVeiculoGar 
         Height          =   270
         Left            =   1980
         TabIndex        =   12
         Top             =   1170
         Width           =   3700
         _ExtentX        =   6535
         _ExtentY        =   476
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
      Begin ReportX.ReportField lblPlacaGar 
         Height          =   270
         Left            =   5790
         TabIndex        =   13
         Top             =   1170
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   476
         Caption         =   "PLACA:"
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
      Begin ReportX.ReportField txtPlacaGar 
         Height          =   270
         Left            =   6510
         TabIndex        =   14
         Top             =   1170
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   476
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
      Begin ReportX.ReportField lblKmAtualGar 
         Height          =   270
         Left            =   8010
         TabIndex        =   15
         Top             =   1170
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   476
         Caption         =   "KM ATUAL:"
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
      Begin ReportX.ReportField txtKmAtualGar 
         Height          =   270
         Left            =   9360
         TabIndex        =   16
         Top             =   1170
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   476
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
      Begin VB.Line LinhaTitulo 
         BorderStyle     =   3  'Dot
         X1              =   570
         X2              =   12240
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line LinhaHeader 
         BorderStyle     =   3  'Dot
         X1              =   570
         X2              =   12240
         Y1              =   1680
         Y2              =   1680
      End
   End
   Begin ReportX.ReportSection SecaoDetalhe 
      Align           =   1  'Align Top
      Height          =   240
      Left            =   0
      Top             =   795
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   423
      AutoEncolher    =   -1  'True
      AutoExpandir    =   -1  'True
      Begin ReportX.ReportField rpfParagrafo 
         Height          =   210
         Left            =   630
         TabIndex        =   17
         Top             =   0
         Width           =   10080
         _ExtentX        =   18800
         _ExtentY        =   370
         Campo           =   "Paragrafo"
         Caption         =   ""
         TipoCampo       =   64
         Formula         =   -1  'True
         WordWrap        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TamanhoAuto     =   -1  'True
         Justificar      =   -1  'True
      End
   End
End
Attribute VB_Name = "REL_OS_Garantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private colParagrafos As Collection
Private colNegrito As Collection
Private colTamanho As Collection
Private pIndice As Long
Private pTextoAtual As String

Public Sub LimparParagrafos()
    Set colParagrafos = New Collection
    Set colNegrito = New Collection
    Set colTamanho = New Collection
    pIndice = 0
End Sub

Public Sub AdicionarParagrafo(ByVal sTexto As String, Optional ByVal bNegrito As Boolean = False, Optional ByVal dTamanho As Double = 8.25)
    If colParagrafos Is Nothing Then LimparParagrafos
    colParagrafos.Add sTexto
    colNegrito.Add bNegrito
    colTamanho.Add dTamanho
End Sub

Public Function ContagemParagrafos() As Long
    If colParagrafos Is Nothing Then
        ContagemParagrafos = 0
    Else
        ContagemParagrafos = colParagrafos.Count
    End If
End Function

Private Sub Form_Load()

End Sub

Private Sub ReportMain1_IniciarRelatorio(ByVal Impressora As Boolean, Cancelar As Boolean)
    If Not Impressora Then
        pIndice = 0
    End If
End Sub

Private Sub ReportMain1_ValidarRegistro(Invalido As Boolean, Cancelar As Boolean)
    pIndice = pIndice + 1
    If pIndice >= 1 And pIndice <= colParagrafos.Count Then
        pTextoAtual = colParagrafos(pIndice)
    Else
        pTextoAtual = ""
    End If
End Sub

Private Sub ReportMain1_FormulaCampo(ByVal Campo As String, Valor As Variant)
    Select Case Campo
        Case "Paragrafo": Valor = pTextoAtual
    End Select
End Sub

Private Sub ReportMain1_IniciarSecao(ByVal Secao As ReportX.TSecao, ByVal Ordem As Byte)
    If Secao = secDetalhe Then
        If pIndice >= 1 And pIndice <= colParagrafos.Count Then
            rpfParagrafo.Font.Bold = colNegrito(pIndice)
            rpfParagrafo.Font.Size = colTamanho(pIndice)
        End If
    End If
End Sub

Private Sub ReportMain1_Erro(ByVal Numero As Long)
    MsgBox "Erro ao gerar o Termo de Garantia: " & Numero & vbCrLf & Error(Numero), vbCritical, "Impressão"
End Sub

