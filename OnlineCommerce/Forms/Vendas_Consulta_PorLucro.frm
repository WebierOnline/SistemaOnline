VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Vendas_Consulta_PorLucro 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   Caption         =   "ITENS DO PEDIDO"
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10650
   Icon            =   "Vendas_Consulta_PorLucro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4515
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   7964
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
      Top             =   60
      Width           =   10395
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DETALHAMENTO DE VENDAS COM LUCRO ESTIMADO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1575
         TabIndex        =   2
         Top             =   300
         Width           =   7665
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   240
         Picture         =   "Vendas_Consulta_PorLucro.frx":23D2
         Top             =   0
         Width           =   1140
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdFechar 
      Height          =   495
      Left            =   8640
      TabIndex        =   0
      Top             =   5940
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   873
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
      MICON           =   "Vendas_Consulta_PorLucro.frx":8C18
      PICN            =   "Vendas_Consulta_PorLucro.frx":8C34
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdExibirPedidos 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5880
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Exibir Vendas"
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
      MICON           =   "Vendas_Consulta_PorLucro.frx":8F4E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblDescricao 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Produto:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Top             =   1020
      Width           =   1020
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Produto:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1020
      Width           =   1155
   End
End
Attribute VB_Name = "Vendas_Consulta_PorLucro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cCfg As ConfigItem
Dim tipoEmpresa As Integer
Dim sSQL As String
Dim r As ADODB.Recordset
Dim r2 As ADODB.Recordset
Public Sub loadPedidos2(ByVal Pedido As Long, ByVal DATA As Date)

sSQL = "SELECT pedidos_itens.COD_PEDIDO, pedidos_itens.DATA, pedidos_itens.COD_PRODUTO, produtos.DESCRICAO, pedidos_itens.PRECO, pedidos_itens.QUANTIDADE, pedidos_itens.Subtotal, pedidos_itens.Desconto, pedidos_itens.Total, pedidos_itens.Custo, " & _
       "CASE pedidos_itens.Custo WHEN 0 THEN 0 ELSE (pedidos_itens.Total - pedidos_itens.Custo * pedidos_itens.QUANTIDADE) / (pedidos_itens.Custo * pedidos_itens.QUANTIDADE) * 100 END AS vSomaMARGEM, " & _
       "CASE pedidos_itens.Custo WHEN 0 THEN 0 ELSE pedidos_itens.Total - (pedidos_itens.custo * pedidos_itens.QUANTIDADE) END AS vSomaLUCRO " & _
       "FROM pedidos_itens INNER JOIN produtos ON pedidos_itens.COD_PRODUTO = produtos.CODIGO " & _
       "WHERE (pedidos_itens.cod_produto = " & Pedido & ") AND (pedidos_itens.data = CONVERT(DATETIME, '" & Format(DATA, ocDATA) & "', 103)) " & _
       "ORDER BY pedidos_itens.COD_PEDIDO DESC"

Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
   lblDescricao.Caption = r("DESCRICAO")
End If
 
FormatarGrid_Itens r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub
Public Sub loadPedidos3(ByVal Pedido As Long, ByVal Data1 As Date, ByVal Data2 As Date)

sSQL = "SELECT pedidos_itens.COD_PEDIDO, pedidos_itens.DATA, pedidos_itens.COD_PRODUTO, produtos.DESCRICAO, pedidos_itens.PRECO, pedidos_itens.QUANTIDADE, pedidos_itens.Subtotal, pedidos_itens.Desconto, pedidos_itens.Total, pedidos_itens.Custo, " & _
       "CASE pedidos_itens.Custo WHEN 0 THEN 0 ELSE (pedidos_itens.Total - pedidos_itens.Custo * pedidos_itens.QUANTIDADE) / (pedidos_itens.Custo * pedidos_itens.QUANTIDADE) * 100 END AS vSomaMARGEM, " & _
       "CASE pedidos_itens.Custo WHEN 0 THEN 0 ELSE pedidos_itens.Total - (pedidos_itens.custo * pedidos_itens.QUANTIDADE) END AS vSomaLUCRO " & _
       "FROM pedidos_itens INNER JOIN produtos ON pedidos_itens.COD_PRODUTO = produtos.CODIGO " & _
       "WHERE (pedidos_itens.cod_produto = " & Pedido & ") AND (pedidos_itens.data >= CONVERT(DATETIME, '" & Format(Data1, ocDATA) & "', 103)) AND (pedidos_itens.data <= CONVERT(DATETIME, '" & Format(Data2, ocDATA) & "', 103)) " & _
       "ORDER BY pedidos_itens.COD_PEDIDO DESC"

Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
   lblDescricao.Caption = r("DESCRICAO")
End If
 
FormatarGrid_Itens r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub
Public Sub loadPedidos(ByVal Pedido As Long, ByVal MES As Integer, ByVal ANO As Integer)

sSQL = "SELECT pedidos_itens.COD_PEDIDO, pedidos_itens.DATA, pedidos_itens.COD_PRODUTO, produtos.DESCRICAO, pedidos_itens.PRECO, pedidos_itens.QUANTIDADE, pedidos_itens.Subtotal, pedidos_itens.Desconto, pedidos_itens.Total, pedidos_itens.Custo, " & _
       "CASE pedidos_itens.Custo WHEN 0 THEN 0 ELSE (pedidos_itens.Total - pedidos_itens.Custo * pedidos_itens.QUANTIDADE) / (pedidos_itens.Custo * pedidos_itens.QUANTIDADE) * 100 END AS vSomaMARGEM, " & _
       "CASE pedidos_itens.Custo WHEN 0 THEN 0 ELSE pedidos_itens.Total - (pedidos_itens.custo * pedidos_itens.QUANTIDADE) END AS vSomaLUCRO " & _
       "FROM pedidos_itens INNER JOIN produtos ON pedidos_itens.COD_PRODUTO = produtos.CODIGO " & _
       "WHERE (pedidos_itens.cod_produto = " & Pedido & ") AND (MONTH(pedidos_itens.data) = " & MES & ") AND (YEAR(pedidos_itens.data) = " & ANO & ") " & _
       "ORDER BY pedidos_itens.COD_PEDIDO DESC"

Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
   lblDescricao.Caption = r("DESCRICAO")
End If
 
FormatarGrid_Itens r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub



Private Sub cmdExibirPedidos_Click()
If Grid.Col = 0 Then Exit Sub
If IsNumeric(Grid.TextMatrix(Grid.Row, 1)) = True Then
   If Grid.Col = 1 Then
      If Grid.TextMatrix(Grid.Row, 1) = "" Then Exit Sub
      Parcelas_Consulta_Produtos.loadPedidos Grid.TextMatrix(Grid.Row, 1), " VENDAS "
      Parcelas_Consulta_Produtos.Show 1
   End If
End If

'If Not IsNumeric(Grid.TextMatrix(Grid.Row, 1)) = True Then Exit Sub
'If Grid.TextMatrix(Grid.Row, 1) = "" Or Grid.TextMatrix(Grid.Row, 10) = "" Then Exit Sub

'If Grid.TextMatrix(Grid.Row, 10) <> "0" Then
'   Parcelas_Consulta_Produtos.loadPedidos Grid.TextMatrix(Grid.Row, 1), "OS"
'Else
'   Parcelas_Consulta_Produtos.loadPedidos Grid.TextMatrix(Grid.Row, 1), Grid.TextMatrix(Grid.Row, 7)
'End If

'Parcelas_Consulta_Produtos.Show 1
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub
Private Sub FormatarGrid_Itens(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid
      .Clear
      .Cols = 11
      .rows = 2
      
      .ColWidth(0) = 150
      .ColWidth(1) = 800
      .ColWidth(2) = 1000
      .ColWidth(3) = 900
      .ColWidth(4) = 800
      .ColWidth(5) = 1100
      .ColWidth(6) = 900
      .ColWidth(7) = 1000
      .ColWidth(8) = 1200
      .ColWidth(9) = 1000
      .ColWidth(10) = 1200
      
      .TextMatrix(0, 1) = "PEDIDO"
      .TextMatrix(0, 2) = "DATA"
      .TextMatrix(0, 3) = "PREÇO"
      .TextMatrix(0, 4) = "QTDE"
      .TextMatrix(0, 5) = "SUBTOTAL"
      .TextMatrix(0, 6) = "DESC."
      .TextMatrix(0, 7) = "TOTAL"
      .TextMatrix(0, 8) = "CUSTO UND"
      .TextMatrix(0, 9) = "LUCRO"
      .TextMatrix(0, 10) = "MARGEM %"
      
      .Redraw = False
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next i
      
      .ColAlignment(1) = 3
      .ColAlignment(2) = 3
      i = 1
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = Format(rTabela("cod_pedido"), "000000")
            .TextMatrix(.rows - 1, 2) = Format(rTabela("data"), "dd/mm/yy")
            .TextMatrix(.rows - 1, 3) = FormatNumber(rTabela("PRECO"), 2)
            .TextMatrix(.rows - 1, 4) = rTabela("QUANTIDADE")
            .TextMatrix(.rows - 1, 5) = FormatNumber(rTabela("SUBTOTAL"), 2)
            .TextMatrix(.rows - 1, 6) = FormatNumber(rTabela("DESCONTO"), 2)
            .TextMatrix(.rows - 1, 7) = FormatNumber(rTabela("Total"), 2)
            .TextMatrix(.rows - 1, 8) = FormatNumber(rTabela("CUSTO"), 2)
            .TextMatrix(.rows - 1, 9) = FormatNumber(rTabela("VSOMALUCRO"), 2)
            .TextMatrix(.rows - 1, 10) = FormatNumber(rTabela("VSOMAMARGEM"), 2)

            rTabela.MoveNext
            .rows = .rows + 1
            i = i + 1
         Loop
      End If
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 1
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 9
         .CellForeColor = &H8000&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .rows - 1
         .Row = i
         .Col = 10
         .CellForeColor = &H8000&
         .CellFontBold = True
      Next
      
      .rows = .rows - 1
      Grid.Redraw = True
   End With
   
   'lblTotal.Caption = Format(SomaGrid(Grid, 4), ocMONEY)
   'lblEntrada.Caption = Format(0, ocMONEY)

End Sub

Private Sub Form_Load()
Set cCfg = sysConfig("TIPO_EMPRESA")
tipoEmpresa = cCfg.Value
Set cCfg = Nothing

End Sub

