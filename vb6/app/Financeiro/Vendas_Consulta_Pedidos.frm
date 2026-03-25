VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Vendas_Consulta_Pedidos 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   Caption         =   "ITENS DO PEDIDO"
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10650
   Icon            =   "Vendas_Consulta_Pedidos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4695
      Left            =   120
      TabIndex        =   3
      Top             =   1140
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   8281
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
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DETALHAMENTO POR VENDAS"
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
         Left            =   1575
         TabIndex        =   2
         Top             =   300
         Width           =   4755
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   240
         Picture         =   "Vendas_Consulta_Pedidos.frx":23D2
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
      MICON           =   "Vendas_Consulta_Pedidos.frx":8C18
      PICN            =   "Vendas_Consulta_Pedidos.frx":8C34
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdExibirProdutos 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5880
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Exibir produtos"
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
      MICON           =   "Vendas_Consulta_Pedidos.frx":8F4E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdExibirParcelas 
      Height          =   255
      Left            =   2820
      TabIndex        =   5
      Top             =   5880
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Exibir Parcelas"
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
      MICON           =   "Vendas_Consulta_Pedidos.frx":8F6A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "Vendas_Consulta_Pedidos"
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
   Dim r2 As ADODB.Recordset

   'sSQL = "SELECT cliente.*, pedidos.*,  pedidos_itens.* " & _
         "FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente INNER JOIN pedidos_itens ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
         "WHERE (pedidos_itens.cod_produto = " & Pedido & ")"
    sSQL = "SELECT pedidos.cod_pedido, pedidos.data_compra, pedidos.total, pedidos.tipo_pedido, pedidos.pagamento, pedidos.tipo_pagamento, pedidos_itens.cod_produto, cliente.nome " & _
         "FROM  pedidos INNER JOIN pedidos_itens ON pedidos.cod_pedido = pedidos_itens.cod_pedido INNER JOIN cliente ON pedidos.cod_cliente = cliente.codigo " & _
         "WHERE (pedidos_itens.cod_produto = " & Pedido & ")"

   Set r = dbData.OpenRecordset(sSQL)
   
   FormatarGrid_Itens r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
   
 '  sSQL = "SELECT * FROM pedidos WHERE (cod_pedido = " & Pedido & ");"
 '  Set r = dbData.OpenRecordset(sSQL)
   
 '  If Not r.BOF Then
 '     lblTotal.Caption = Format(r("subtotal"), ocMONEY)
 '     lblTotalGeral.Caption = Format(r("total"), ocMONEY)
      
 '     If r("tipo_desc") = "R" Then
 '        lblTotalDesc.Caption = Format(r("valor_desc"), ocMONEY)
 '     Else
 '        lblTotalDesc.Caption = FormatNumber(r("valor_desc")) & "%"
 '     End If
 '  End If
   
 '  If r.State <> 0 Then r.Close
 '  Set r = Nothing
   
 '  txtCodPedido.Text = Format(Pedido, "000000")
End Sub

Private Sub cmdExibirParcelas_Click()
If Grid.Col = 0 Then Exit Sub
   If IsNumeric(Grid.TextMatrix(Grid.Row, 1)) = True Then
         Vendas_Consulta_Geral_Parcelas.loadInformacoes (Grid.TextMatrix(Grid.Row, 1))
         Vendas_Consulta_Geral_Parcelas.Show 1
   End If
End Sub

Private Sub cmdExibirProdutos_Click()
If Grid.Col = 0 Then Exit Sub
If IsNumeric(Grid.TextMatrix(Grid.Row, 1)) = True Then
   If Grid.Col = 1 Then
      If Grid.TextMatrix(Grid.Row, 1) = "" Then Exit Sub
      Parcelas_Consulta_Produtos.loadPedidos Grid.TextMatrix(Grid.Row, 1), Grid.TextMatrix(Grid.Row, 7)
      Parcelas_Consulta_Produtos.Show 1
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
      .Cols = 8
      .rows = 2
      
      .ColWidth(0) = 150
      .ColWidth(1) = 800
      .ColWidth(2) = 1000
      .ColWidth(3) = 4300
      .ColWidth(4) = 1000
      .ColWidth(5) = 1100
      .ColWidth(6) = 1220
      .ColWidth(7) = 0
      
      .TextMatrix(0, 1) = "PEDIDO"
      .TextMatrix(0, 2) = "DATA"
      .TextMatrix(0, 3) = "NOME DO CLIENTE"
      .TextMatrix(0, 4) = "VALOR"
      .TextMatrix(0, 5) = "FORMA"
      .TextMatrix(0, 6) = "FORMA"
      .TextMatrix(0, 7) = "TIPO"
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
            .TextMatrix(.rows - 1, 2) = Format(rTabela("data_compra"), "dd/mm/yy")
            .TextMatrix(.rows - 1, 3) = UCase(rTabela("nome"))
            .TextMatrix(.rows - 1, 4) = Format(rTabela("total"), ocMONEY)
            .TextMatrix(.rows - 1, 5) = rTabela("tipo_pagamento")
            .TextMatrix(.rows - 1, 6) = rTabela("pagamento")
            .TextMatrix(.rows - 1, 7) = rTabela("tipo_pedido")
            
            
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
         .Col = 4
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

