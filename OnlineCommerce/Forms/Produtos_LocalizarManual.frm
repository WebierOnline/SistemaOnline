VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Produtos_LocalizarManual 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   60
      ScaleHeight     =   645
      ScaleWidth      =   10605
      TabIndex        =   8
      Top             =   60
      Width           =   10635
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Escolha abaixo um produto que represente esse produto de sua entrada."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   1380
         TabIndex        =   11
         Top             =   360
         Width           =   7740
      End
      Begin VB.Label lblProduto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUTO"
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
         Left            =   60
         TabIndex        =   9
         Top             =   60
         Width           =   10500
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4455
      Left            =   60
      TabIndex        =   5
      Top             =   1560
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   7858
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
      Left            =   60
      ScaleHeight     =   315
      ScaleWidth      =   6555
      TabIndex        =   1
      Top             =   720
      Width           =   6555
      Begin VB.OptionButton optDesc 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Descrição"
         Height          =   195
         Left            =   3060
         TabIndex        =   4
         Top             =   90
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.OptionButton optCodigo 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Código"
         Height          =   195
         Left            =   720
         TabIndex        =   3
         Top             =   90
         Width           =   855
      End
      Begin VB.OptionButton optCodBarra 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cód. de Barra"
         Height          =   195
         Left            =   1620
         TabIndex        =   2
         Top             =   90
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Filtros:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   10
         Top             =   60
         Width           =   795
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
      Top             =   1080
      Width           =   10635
   End
   Begin ChamaleonBtn.chameleonButton cmdRetornar 
      Height          =   315
      Left            =   7140
      TabIndex        =   6
      ToolTipText     =   "Adiciona"
      Top             =   6060
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Usar esse produto"
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
      MICON           =   "Produtos_LocalizarManual.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdCancelar 
      Height          =   315
      Left            =   9180
      TabIndex        =   7
      ToolTipText     =   "Adiciona"
      Top             =   6060
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Cancelar"
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
      MICON           =   "Produtos_LocalizarManual.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Shape Shape1 
      Height          =   6435
      Left            =   0
      Top             =   0
      Width           =   10755
   End
End
Attribute VB_Name = "Produtos_LocalizarManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sSQL As String
Dim r As ADODB.Recordset
Option Explicit

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdRetornar_Click()
Dim sSQL As String
Dim r As ADODB.Recordset

If Grid.Rows <= 1 Then Exit Sub

sSQL = "SELECT EAN, DESCRICAO, UNID_MEDIDA, QUANT_ESTOQUE, NCM, ICMSCST, CFOP, CODIGO,   " & _
             "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
             "FROM produtos WHERE (CODIGO = " & Grid.TextMatrix(Grid.Row, 1) & ")"
             
Set r = dbData.OpenRecordset(sSQL)
Entrada_Estoque.Show
Entrada_Estoque.lblProdutoCadastrado.Caption = "Produto Cadastrado"
Entrada_Estoque.lblProdutoCadastrado.ForeColor = &HC00000
Entrada_Estoque.txtCodProdExist = Format(r("CODIGO"), "@")
Entrada_Estoque.txtCodBarraExist = Format(r("EAN"), "@")
Entrada_Estoque.txtDescricaoExist = Format(r("DESCRICAO"), "@")
Entrada_Estoque.txtUnidMedExist = Format(r("UNID_MEDIDA"), "@")
Entrada_Estoque.txtQuantExist = Format(r("QUANT_ESTOQUE"), "@")
Entrada_Estoque.txtValorExist = Format(r("venda"), "##,##0.00")
Entrada_Estoque.txtNCMExist = Format(r("NCM"), "@")
Entrada_Estoque.txtCSTExist = Format(r("ICMSCST"), "@")
Entrada_Estoque.txtCFOPExist = Format(r("CFOP"), "@")
Entrada_Estoque.cmdAtualEAN.Enabled = True
Entrada_Estoque.cmdAtualDesc.Enabled = True
Entrada_Estoque.cmdAtualUnid.Enabled = True
Entrada_Estoque.cmdAtualNCM.Enabled = True
Entrada_Estoque.cmdAtualProd.Enabled = False
Entrada_Estoque.cmdCancelarEntrada.Enabled = True
Entrada_Estoque.cmdLocalizarManual.Enabled = False
txtDescricao.Text = ""
Me.Hide
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
   Save = 0
End Sub

Private Sub Form_Load()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   
   sSQL = "SELECT top(200) produtos.codigo AS var_cod, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, " & _
      "produtos.prateleira AS var_prat, produtos.unid_medida AS var_med, produtos.quant_estoque AS var_quant, produtos.fabricante as var_fab, " & _
      "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
      "FROM produtos " & _
      "WHERE (produtos.ativo = 1) ORDER BY produtos.descricao;"
   
   Set r = dbData.OpenRecordset(sSQL)
   
   Formatar_Grid r
   
   If r.State <> 0 Then r.Close
   Set r = Nothing
End Sub

Private Sub lblProduto_Change()
Dim vTextoSelecionado As Variant
Dim vConcatenado As String
vTextoSelecionado = Split(TirarEspaco(lblProduto.Caption), " ")
'vConcatenado = vTextoSelecionado(0) & " " & vTextoSelecionado(1)
vConcatenado = vTextoSelecionado(0)
txtDescricao.Text = vConcatenado
End Sub

Public Function TirarEspaco(ByVal Value As String) As String
Dim bRepete As Boolean
Value = Replace$(Value, "'", vbNullString)
Do
  Value = Replace$(Value, "  ", " ")
  bRepete = InStr(1, Value, "  ", vbTextCompare)
  Value = Trim(Value)
Loop Until Not bRepete

TirarEspaco = Value
End Function

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
   
   sSQL = "SELECT produtos.codigo AS var_cod, produtos.cod_barra AS var_codbarra, produtos.descricao AS var_desc, produtos.fabricante AS var_fab, " & _
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
            .TextMatrix(.Rows - 1, 2) = ValidateNull(rTabela("var_codbarra"))
            .TextMatrix(.Rows - 1, 3) = rTabela("var_desc") & " /  " & ValidateNull(rTabela("var_fab"))
            .TextMatrix(.Rows - 1, 4) = ValidateNull(rTabela("var_med"))
            .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("var_prat"))
            .TextMatrix(.Rows - 1, 6) = ValidateNull(rTabela("var_quant"))
            .TextMatrix(.Rows - 1, 7) = Format(ValidateNull(rTabela("venda")), ocMONEY)
            
            rTabela.MoveNext
            .Rows = .Rows + 1
         Loop
      End If
      
      .Rows = .Rows - 1
      .Redraw = True
   End With
End Sub
