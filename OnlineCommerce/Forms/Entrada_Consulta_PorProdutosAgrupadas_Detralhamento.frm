VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Entrada_Consulta_PorProdutosAgrupadas_Detralhamento 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   Caption         =   "ITENS DO PEDIDO"
   ClientHeight    =   6180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10650
   Icon            =   "Entrada_Consulta_PorProdutosAgrupadas_Detralhamento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4455
      Left            =   120
      TabIndex        =   3
      Top             =   1140
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   7858
      _Version        =   393216
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   120
      ScaleHeight     =   885
      ScaleWidth      =   10365
      TabIndex        =   1
      Top             =   180
      Width           =   10395
      Begin VB.Label lblProduto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   1380
         TabIndex        =   12
         Top             =   420
         Width           =   240
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HISTÓRICOS DE ENTRADAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   1380
         TabIndex        =   2
         Top             =   120
         Width           =   4275
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   240
         Picture         =   "Entrada_Consulta_PorProdutosAgrupadas_Detralhamento.frx":23D2
         Top             =   0
         Width           =   1140
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdFechar 
      Height          =   315
      Left            =   4800
      TabIndex        =   0
      Top             =   5640
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Fechar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Entrada_Consulta_PorProdutosAgrupadas_Detralhamento.frx":8C18
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdExibirNota 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   5640
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "MOSTRAR ENTRADAS"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Entrada_Consulta_PorProdutosAgrupadas_Detralhamento.frx":8C34
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdImprimir 
      Height          =   315
      Left            =   2820
      TabIndex        =   11
      Top             =   5640
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Imprimir"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Entrada_Consulta_PorProdutosAgrupadas_Detralhamento.frx":8C50
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remoção:"
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
      Left            =   7620
      TabIndex        =   10
      Top             =   5820
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Adição:"
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
      Left            =   7860
      TabIndex        =   9
      Top             =   5640
      Width           =   720
   End
   Begin VB.Label lblQuantRemocao 
      Alignment       =   1  'Right Justify
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8700
      TabIndex        =   8
      Top             =   5880
      Width           =   525
   End
   Begin VB.Label lblQuantAdicao 
      Alignment       =   1  'Right Justify
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8700
      TabIndex        =   7
      Top             =   5640
      Width           =   540
   End
   Begin VB.Label lblTotalRemocao 
      Alignment       =   1  'Right Justify
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9300
      TabIndex        =   6
      Top             =   5880
      Width           =   1200
   End
   Begin VB.Label lblTotalAdicao 
      Alignment       =   1  'Right Justify
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9300
      TabIndex        =   5
      Top             =   5640
      Width           =   1200
   End
End
Attribute VB_Name = "Entrada_Consulta_PorProdutosAgrupadas_Detralhamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cCfg As ConfigItem
Dim tipoEmpresa As Integer
Public Sub loadPedidos(ByVal Pedido As Long)
Dim sSQL As String
Dim r As ADODB.Recordset

sSQL = "SELECT descricao FROM produtos WHERE (codigo = " & Pedido & ");"
Set r = dbData.OpenRecordset(sSQL)
If Not r.BOF Then
   lblProduto.Caption = "Produto: " & ValidateNull(r("descricao"))
Else
   lblProduto.Caption = "Produto: ?"
End If
If r.State <> 0 Then r.Close
Set r = Nothing

sSQL = "SELECT Produtos_Quant.Codigo AS vCod, Produtos_Quant.Hora AS vHora, Produtos_Quant.COD_PRODUTO, Produtos_Quant.COD_ENTRADA, Produtos_Quant.FORMA, Produtos_Quant.Tipo, Produtos_Quant.QUANT, Produtos_Quant.Data, ISNULL(funcionario.NOME, '') AS varNomeUsuario " & _
      "FROM Produtos_Quant LEFT JOIN funcionario ON funcionario.CODIGO = Produtos_Quant.cod_usuario " & _
      "WHERE (Produtos_Quant.COD_PRODUTO = " & Pedido & ") " & _
      "ORDER BY Produtos_Quant.data"

Set r = dbData.OpenRecordset(sSQL)

FormatarGrid_Itens r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub cmdExibirNota_Click()
If Grid.Col = 0 Or Grid.Row = 0 Then Exit Sub

If Grid.TextMatrix(Grid.Row, 7) = "ENTRADA" Then
    'Veio de nota fiscal - abre direto a nota exata (coluna 6 = COD_ENTRADA)
    If IsNumeric(Grid.TextMatrix(Grid.Row, 6)) = True Then
        If Grid.TextMatrix(Grid.Row, 6) <> "0000" Then
            Produtos_Entrada.AbrirNotaDireto Val(Grid.TextMatrix(Grid.Row, 6))
            Me.Hide
            Produtos_Entrada.Show
        End If
    End If
Else
    'CADASTRO/AJUSTE - sem nota; abre o historico de ajuste desse produto
    '(coluna 5 = COD_PRODUTO)
    If IsNumeric(Grid.TextMatrix(Grid.Row, 5)) = True Then
        If Grid.TextMatrix(Grid.Row, 5) <> "0000" Then
            Ajuste_Consulta_PorProdutos.FiltrarPorProduto Val(Grid.TextMatrix(Grid.Row, 5))
            Me.Hide
            Ajuste_Consulta_PorProdutos.Show
        End If
    End If
End If
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub
Private Sub FormatarGrid_Itens(rTabela As ADODB.Recordset)
Dim i As Integer

With Grid
   .Clear
   .Cols = 10
   .Rows = 2

   .ColWidth(0) = 150
   .ColWidth(1) = 0
   .ColWidth(2) = 800
   .ColWidth(3) = 700
   .ColWidth(4) = 1300
   .ColWidth(5) = 0
   .ColWidth(6) = 0
   .ColWidth(7) = 1220
   .ColWidth(8) = 1000
   .ColWidth(9) = 1220

   .TextMatrix(0, 1) = "CÓD."
   .TextMatrix(0, 2) = "DATA"
   .TextMatrix(0, 3) = "HORA"
   .TextMatrix(0, 4) = "USUÁRIO"
   .TextMatrix(0, 7) = "FORMA"
   .TextMatrix(0, 8) = "TIPO"
   .TextMatrix(0, 9) = "QUANT."
   .Redraw = False

   'colocar os cabeçalho em negrito
   For i = 0 To .Cols - 1
      .Col = i
      .Row = 0
      .CellFontBold = True
   Next i

   .ColAlignment(1) = 3
   .ColAlignment(2) = 3
   .ColAlignment(3) = 3
   i = 1

   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         .TextMatrix(.Rows - 1, 1) = Format(rTabela("VCOD"), "00000")
         .TextMatrix(.Rows - 1, 2) = Format(rTabela("Data"), "dd/mm/yy")
         If Trim(ValidateNull(rTabela("vHora"))) = "" Then
            .TextMatrix(.Rows - 1, 3) = ""
         ElseIf Format(rTabela("vHora"), "hh:mm:ss") = "00:00:00" Then
            .TextMatrix(.Rows - 1, 3) = ""
         Else
            .TextMatrix(.Rows - 1, 3) = Format(rTabela("vHora"), "hh:mm:ss")
         End If
         .TextMatrix(.Rows - 1, 4) = ValidateNull(rTabela("varNomeUsuario"))
         .TextMatrix(.Rows - 1, 5) = Format(rTabela("COD_PRODUTO"), "0000")
         .TextMatrix(.Rows - 1, 6) = Format(rTabela("COD_ENTRADA"), "0000")
         .TextMatrix(.Rows - 1, 7) = rTabela("FORMA")
         .TextMatrix(.Rows - 1, 8) = rTabela("TIPO")
         .TextMatrix(.Rows - 1, 9) = rTabela("QUANT")

         rTabela.MoveNext
         .Rows = .Rows + 1
         i = i + 1
      Loop
   End If

   'MUDAR COR DE FONTE DA COLUNA
   'For i = 1 To .Rows - 1
   '   .Row = i
   '   .Col = 1
   '   .CellForeColor = &HC0&
   '   .CellFontBold = True
   'Next

   'MUDAR COR DE FONTE DA COLUNA
   For i = 1 To .Rows - 1
      .Row = i
      .Col = 9
      .CellForeColor = &H8000&
      .CellFontBold = True
   Next

    Dim j As Integer
    For i = 1 To .Rows - 1
       For j = 0 To .Cols - 1
          .Col = j
          .Row = i

          If .TextMatrix(i, 8) = "REMOÇÃO" Then
             .CellForeColor = vbRed
          Else
             .CellForeColor = vbBlack
          End If

       Next
    Next

   .Rows = .Rows - 1
   .Row = 0
   .Col = 0
   Grid.Redraw = True
End With

SomaTotais
'lblTotalGeral.Caption = SomaGrid(Grid, 9)
End Sub
Public Function SomaGrid(Grid As MSFlexGrid, Col As Integer) As Currency
Dim i As Integer, Valor As Currency

Valor = 0
For i = 0 To Grid.rows - 1
   If IsNumeric(Grid.TextMatrix(i, Col)) Then
      Valor = Valor + CCur(Grid.TextMatrix(i, Col))
   End If
Next

SomaGrid = Valor
End Function
Private Sub SomaTotais()
Dim soma As Currency
Dim QUANT As Integer
Dim i As Integer

soma = 0
QUANT = 0
With Grid
   For i = 1 To .rows - 1
      If .TextMatrix(i, 8) = "ADIÇÃO" Then
         soma = soma + CCur(.TextMatrix(i, 9))
         QUANT = QUANT + 1
      End If
   Next
End With

lblTotalAdicao.Caption = soma
lblQuantAdicao.Caption = Format(QUANT, "000")

soma = 0
QUANT = 0

With Grid
   For i = 1 To .rows - 1
      If .TextMatrix(i, 8) = "REMOÇÃO" Then
         soma = soma + CCur(.TextMatrix(i, 9))
         QUANT = QUANT + 1
      End If
   Next
End With

lblTotalRemocao.Caption = soma
lblQuantRemocao.Caption = Format(QUANT, "000")


soma = 0
QUANT = 0
End Sub
Private Sub Form_Load()
Set cCfg = sysConfig("TIPO_EMPRESA")
tipoEmpresa = cCfg.Value
Set cCfg = Nothing
End Sub
