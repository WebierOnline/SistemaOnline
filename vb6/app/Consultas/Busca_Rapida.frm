VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Busca_Rapida 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4035
      Left            =   60
      TabIndex        =   5
      Top             =   1020
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   7117
      _Version        =   393216
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   6555
      TabIndex        =   1
      Top             =   120
      Width           =   6555
      Begin VB.OptionButton optDesc 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Descrição"
         Height          =   195
         Left            =   2520
         TabIndex        =   4
         Top             =   60
         Width           =   1935
      End
      Begin VB.OptionButton optCodigo 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Código"
         Height          =   195
         Left            =   0
         TabIndex        =   3
         Top             =   60
         Width           =   855
      End
      Begin VB.OptionButton optCodBarra 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cód. de Barra"
         Height          =   195
         Left            =   900
         TabIndex        =   2
         Top             =   60
         Value           =   -1  'True
         Width           =   1395
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3420
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox txtDescricao 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   10635
   End
   Begin VB.Shape Shape1 
      Height          =   5115
      Left            =   0
      Top             =   0
      Width           =   10755
   End
End
Attribute VB_Name = "Busca_Rapida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Data1_Validate(Action As Integer, Save As Integer)
   Save = 0
End Sub

Private Sub Form_Load()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT produtos.codigo AS var_cod, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, " & _
      "produtos.prateleira AS var_prat, produtos.unid_medida AS var_med, produtos.quant_estoque AS var_quant, " & _
      "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
      "FROM produtos " & _
      "WHERE (produtos.ativo = 1) ORDER BY produtos.descricao;"
   
   Set r = dbData.OpenRecordset(sSQL)
   
   Formatar_Grid r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub optCodBarra_Click()
   txtDescricao_Change
   txtDescricao.SetFocus
End Sub

Private Sub optCodigo_Click()
   txtDescricao_Change
   txtDescricao.SetFocus
End Sub

Private Sub optDesc_Click()
   txtDescricao_Change
   txtDescricao.SetFocus
End Sub

Private Sub txtDescricao_Change()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT produtos.codigo AS var_cod, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, " & _
      "produtos.prateleira AS var_prat, produtos.unid_medida AS var_med, produtos.quant_estoque AS var_quant, " & _
      "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
      "FROM produtos " & _
      "WHERE "
   'Monta a consulta base
'   sSQL = "SELECT produtos.codigo AS var_cod, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, " & _
'      "produtos.prateleira AS var_prat, produtos.unid_medida AS var_med, produtos.quant_estoque AS var_quant, " & _
'      "ISNULL(produtos_entrada_itens.venda, 0) AS var_venda FROM produtos LEFT JOIN ultimas_entradas ON produtos.codigo = ultimas_entradas.codigo_produto " & _
'      "LEFT JOIN produtos_entrada_itens ON ultimas_entradas.codigo_produto = produtos_entrada_itens.codigo_produto " & _
'      "AND ultimas_entradas.ultentrada = produtos_entrada_itens.codigo_entrada WHERE "
      
   If optCodigo.Value = True Then
      sSQL = sSQL & "(produtos.codigo LIKE '" & txtDescricao & "%') AND (produtos.ativo = 1) ORDER BY produtos.descricao;"
      
   ElseIf optCodBarra.Value = True Then
      sSQL = sSQL & "(produtos.cod_barra LIKE '" & txtDescricao & "%') AND (produtos.ativo = 1) ORDER BY produtos.descricao;"
      
   ElseIf optDesc.Value = True Then
      sSQL = sSQL & "(produtos.descricao LIKE '" & txtDescricao & "%') AND (produtos.ativo = 1) ORDER BY produtos.descricao;"
      
   End If
   
   Set r = dbData.OpenRecordset(sSQL)
   
   Formatar_Grid r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Grid.SetFocus
      Grid.Row = 1
      Grid.Col = 0
      Grid.ColSel = Grid.Cols - 1
   ElseIf KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub Formatar_Grid(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid
      .Clear
      .Cols = 8
      .Rows = 2
      
      .ColWidth(0) = 0
      
      If optCodigo.Value = True Then
         .ColWidth(1) = 1200
         .ColWidth(2) = 0
      ElseIf optCodBarra.Value = True Then
         .ColWidth(1) = 0
         .ColWidth(2) = 1200
      ElseIf optDesc.Value = True Then
         .ColWidth(1) = 1200
         .ColWidth(2) = 0
      End If
      
      .ColWidth(3) = 5100
      .ColWidth(4) = 1000
      .ColWidth(5) = 1000
      .ColWidth(6) = 1000
      .ColWidth(7) = 1000
      
      .TextMatrix(0, 1) = "CÓDIGO"
      .TextMatrix(0, 2) = "CÓD.BARRA"
      .TextMatrix(0, 3) = "DESCRIÇÃO"
      .TextMatrix(0, 4) = "UNID."
      .TextMatrix(0, 5) = "LOCAL"
      .TextMatrix(0, 6) = "ESTOQUE"
      .TextMatrix(0, 7) = "PREÇO"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next i
      
      'centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      .Redraw = False
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            'mudar a cor da coluna
            'For i = 1 To .Rows - 1
            '   .Row = i
            '   .Col = 6
            '   .CellBackColor = &HC0FFFF
            ' Next
            
            'ALINHAMENTO
            .ColAlignment(2) = 1
    
            .TextMatrix(.Rows - 1, 1) = rTabela("var_cod")
            .TextMatrix(.Rows - 1, 2) = rTabela("var_codbarra")
            .TextMatrix(.Rows - 1, 3) = rTabela("var_desc")
            .TextMatrix(.Rows - 1, 4) = ValidateNull(rTabela("var_med"))
            .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("var_prat"))
            .TextMatrix(.Rows - 1, 6) = ValidateNull(rTabela("var_quant"))
            .TextMatrix(.Rows - 1, 7) = Format(ValidateNull(rTabela("var_venda")), ocMONEY)
            
            rTabela.MoveNext
            .Rows = .Rows + 1
         Loop
      End If
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
End Sub
