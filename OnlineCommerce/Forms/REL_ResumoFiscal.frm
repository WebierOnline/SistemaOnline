VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.Ocx"
Begin VB.Form REL_ResumoFiscal 
   Caption         =   "Resumo Fiscal"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   ScaleHeight     =   52.917
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   216.959
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain ReportMain1 
      Height          =   480
      Left            =   8220
      TabIndex        =   0
      Top             =   2400
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
   Begin ReportX.ReportSection SecaoDetalhe 
      Align           =   1  'Align Top
      Height          =   380
      Left            =   0
      Top             =   0
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   529
      AutoEncolher    =   0  'False
      AutoExpandir    =   -1  'True
      Begin ReportX.ReportField rpfTexto 
         Height          =   380
         Left            =   630
         TabIndex        =   1
         Top             =   0
         Width           =   6500
         _ExtentX        =   11465
         _ExtentY        =   529
         Campo           =   "Texto"
         Caption         =   ""
         TipoCampo       =   64
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rpfLongo 
         Height          =   380
         Left            =   630
         TabIndex        =   3
         Top             =   0
         Width           =   10080
         _ExtentX        =   17780
         _ExtentY        =   670
         Campo           =   "Longo"
         Caption         =   ""
         TipoCampo       =   64
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rpfCentro 
         Height          =   380
         Left            =   630
         TabIndex        =   4
         Top             =   0
         Width           =   10080
         _ExtentX        =   17780
         _ExtentY        =   670
         Campo           =   "Centro"
         Caption         =   ""
         TipoCampo       =   64
         Formula         =   -1  'True
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField rpfValor 
         Height          =   380
         Left            =   7200
         TabIndex        =   2
         Top             =   0
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   529
         Campo           =   "Valor"
         Caption         =   ""
         TipoCampo       =   64
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "REL_ResumoFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Relatorio do Resumo Fiscal (Notas_Adesivas). Cada "linha" do relatorio e um registro do Detalhe:
'um texto a esquerda e um valor a direita (alinhado a direita), com negrito/tamanho por linha.
'Receita do ReportX sem recordset: ver REL_OS_Garantia.frm
Private colTexto As Collection
Private colValor As Collection
Private colNegrito As Collection
Private colTamanho As Collection
Private colCentro As Collection
Private pIndice As Long
Private pTexto As String
Private pValor As String
Private pCentro As Boolean

Public Sub LimparLinhas()
    Set colTexto = New Collection
    Set colValor = New Collection
    Set colNegrito = New Collection
    Set colTamanho = New Collection
    Set colCentro = New Collection
    pIndice = 0
End Sub

'sTexto nunca vazio (linha em branco = " "), senao a linha colapsa com AutoEncolher
Public Sub AdicionarLinha(ByVal sTexto As String, ByVal sValor As String, Optional ByVal bNegrito As Boolean = False, Optional ByVal dTamanho As Double = 9, Optional ByVal bCentro As Boolean = False)
    If colTexto Is Nothing Then LimparLinhas
    If sTexto = "" Then sTexto = " "
    colTexto.Add sTexto
    colValor.Add sValor
    colNegrito.Add bNegrito
    colTamanho.Add dTamanho
    colCentro.Add bCentro
End Sub

Public Function LinhaTexto(ByVal i As Long) As String
    LinhaTexto = colTexto(i)
End Function

Public Function LinhaValor(ByVal i As Long) As String
    LinhaValor = colValor(i)
End Function

Public Function LinhaNegrito(ByVal i As Long) As Boolean
    LinhaNegrito = colNegrito(i)
End Function

Public Function LinhaTamanho(ByVal i As Long) As Double
    LinhaTamanho = colTamanho(i)
End Function

Public Function LinhaCentro(ByVal i As Long) As Boolean
    LinhaCentro = colCentro(i)
End Function

Public Function ContagemLinhas() As Long
    If colTexto Is Nothing Then ContagemLinhas = 0 Else ContagemLinhas = colTexto.Count
End Function

Private Sub ReportMain1_IniciarRelatorio(ByVal Impressora As Boolean, Cancelar As Boolean)
    If Not Impressora Then pIndice = 0
End Sub

Private Sub ReportMain1_ValidarRegistro(Invalido As Boolean, Cancelar As Boolean)
    pIndice = pIndice + 1
    If pIndice >= 1 And pIndice <= ContagemLinhas Then
        pTexto = colTexto(pIndice)
        pValor = colValor(pIndice)
        pCentro = colCentro(pIndice)
    Else
        pTexto = " "
        pValor = ""
        pCentro = False
    End If
End Sub

Private Sub ReportMain1_FormulaCampo(ByVal Campo As String, Valor As Variant)
    Select Case Campo
        'cada linha usa so um dos 3 campos de texto; os outros saem em branco
        Case "Texto": If pValor <> "" Then Valor = pTexto Else Valor = " "
        Case "Longo": If pValor = "" And Not pCentro Then Valor = pTexto Else Valor = " "
        Case "Centro": If pValor = "" And pCentro Then Valor = pTexto Else Valor = " "
        Case "Valor": Valor = pValor
    End Select
End Sub

Private Sub ReportMain1_IniciarSecao(ByVal Secao As ReportX.TSecao, ByVal Ordem As Byte)
    If Secao = secDetalhe Then
        If pIndice >= 1 And pIndice <= ContagemLinhas Then
            rpfTexto.Font.Bold = colNegrito(pIndice)
            rpfLongo.Font.Bold = colNegrito(pIndice)
            rpfCentro.Font.Bold = colNegrito(pIndice)
            rpfValor.Font.Bold = colNegrito(pIndice)
            rpfTexto.Font.Size = colTamanho(pIndice)
            rpfLongo.Font.Size = colTamanho(pIndice)
            rpfCentro.Font.Size = colTamanho(pIndice)
            rpfValor.Font.Size = colTamanho(pIndice)
        End If
    End If
End Sub
