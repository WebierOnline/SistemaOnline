VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form OS_Recapadora_Cons_Pecas 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12570
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   12570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Begin VB.OptionButton optRef 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Referência"
         Height          =   195
         Left            =   3540
         TabIndex        =   7
         Top             =   60
         Width           =   1395
      End
      Begin VB.OptionButton optDesc 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Descrição"
         Height          =   195
         Left            =   2340
         TabIndex        =   4
         Top             =   60
         Value           =   -1  'True
         Width           =   1155
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
         Width           =   1395
      End
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
      Width           =   12435
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4095
      Left            =   60
      TabIndex        =   5
      Top             =   960
      Width           =   12435
      _ExtentX        =   21934
      _ExtentY        =   7223
      _Version        =   393216
      BackColorBkg    =   16777215
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pressione [ENTER] para selecionar o produto."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7980
      TabIndex        =   6
      Top             =   5160
      Width           =   4575
   End
   Begin VB.Shape Shape1 
      Height          =   5415
      Left            =   0
      Top             =   0
      Width           =   12570
   End
End
Attribute VB_Name = "OS_Recapadora_Cons_Pecas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim VERIFICAR_QUANTIDADE As Boolean
Dim vTipoOS As String
Dim sSQL As String
Dim r As ADODB.Recordset
Private Sub Form_Activate()
txtDescricao.Text = ""
txtDescricao.SetFocus
End Sub

Private Sub Mostrar_Grid()

If vTipoOS = "Automóveis" Then

'            'Compartibilidade
'            Dim sSQL_Comp As String
'            Dim var_Comp As String
'            Dim rS2 As ADODB.Recordset
            
'            sSQL_Comp = "Select MODELO, ANO From PRODUTOS_COMP Where COD_PRODUTO = " & r("var_cod")
'            Set rS2 = dbData.OpenRecordset(sSQL_Comp)
            
'            Do While Not rS2.EOF
'            var_Comp = var_Comp & rS2!Modelo & "(" & rS2!Ano & "),  "
'            rS2.MoveNext
'            Loop
            
'            If Not IsNull(var_Comp) Then ItemLst.SubItems(3) = var_Comp
'            var_Comp = ""




    sSQL = "SELECT produtos.codigo AS var_cod, produtos.fabricante AS var_fab, " & _
       "produtos.descricao AS var_desc, " & _
       "produtos.quant_estoque AS var_quant, " & _
       "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
       "FROM produtos " & _
       "WHERE (produtos.ativo = 1) ORDER BY produtos.descricao;"
       Debug.Print sSQL
Else
    sSQL = "SELECT produtos.codigo AS var_cod, produtos.fabricante AS var_fab, produtos.fabricante AS var_ref, produtos.cod_barra AS var_codbarra, " & _
       "produtos.descricao AS var_desc, produtos.prateleira AS var_prat, produtos.unid_medida AS var_med, " & _
       "produtos.quant_estoque AS var_quant, " & _
       "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
       "FROM produtos " & _
       "WHERE (produtos.ativo = 1) ORDER BY produtos.descricao;"
End If

Set r = dbData.OpenRecordset(sSQL)

Formatar_Grid r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub

Private Sub Form_Load()
Set oCfg = sysConfig("TIPO_OS")
vTipoOS = oCfg.Value
Set oCfg = Nothing

Mostrar_Grid
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   Verifica_QuantEstoque
   If VERIFICAR_QUANTIDADE = True Then
      txtDescricao.SetFocus
      Exit Sub
   Else
      OS_Recapadora.txtCodPeca.Text = Grid.TextMatrix(Grid.Row, 1)
      OS_Recapadora.cboPecas.Text = Grid.TextMatrix(Grid.Row, 3)
      OS_Recapadora.txtValorPeca.Text = Grid.TextMatrix(Grid.Row, 8)
      Unload Me
      On Local Error Resume Next
      OS_Recapadora.txtQuantPeca.SetFocus
   End If
End If
End Sub

Private Sub Verifica_QuantEstoque()
   Dim sSQL As String
   Dim r As ADODB.Recordset
   Dim oCfg As ConfigItem
   Dim bEstNeg As Boolean
   
   'If txtCodProduto.Text = "" Then Exit Sub
   
   'mostrar o fundo do pdv
   'sSQL = "SELECT estoque_negativo, codigo FROM configuracao WHERE (codigo = 1);"
   'Set r = dbData.OpenRecordset(sSQL)
   
   Set oCfg = sysConfig("ESTOQUE_NEGATIVO")
   bEstNeg = CBool(oCfg.Value)
   Set oCfg = Nothing
   
   If bEstNeg = False Then
      sSQL = "SELECT codigo, quant_estoque FROM produtos WHERE (codigo = " & Grid.TextMatrix(Grid.Row, 1) & ");"
      Set r = dbData.OpenRecordset(sSQL)
      
      VERIFICAR_QUANTIDADE = False
      'If txtQuant.Text = "" Then txtQuant.Text = 0
      
      If r("quant_estoque") <= 0 Then
         ShowMsg "ESSA QUANTIDADE É INVÁLIDA!" & vbCrLf & "SEU ESTOQUE ATUAL É DE 0 (zero) PRODUTO", vbExclamation
         'LimparObjetos_Pedido
         'cmdAlterar.Enabled = False
         VERIFICAR_QUANTIDADE = True
         'txtCodBarra.Text = ""
      End If
   Else
      Exit Sub
   End If
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

Private Sub optRef_Click()
txtDescricao_Change
txtDescricao.SetFocus
End Sub


Private Sub txtDescricao_Change()
Dim sSQL As String
Dim r As ADODB.Recordset

If optCodigo.Value = True Then
sSQL = "SELECT produtos.codigo AS var_cod, produtos.fabricante AS var_fab, produtos.ref AS var_ref, produtos.cod_barra AS var_codbarra, " & _
   "produtos.descricao AS var_desc, produtos.prateleira AS var_prat, produtos.unid_medida AS var_med, " & _
   "produtos.quant_estoque AS var_quant, " & _
   "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
   "FROM produtos " & _
   "WHERE (produtos.codigo LIKE '" & txtDescricao & "%') AND (produtos.ativo = 1) ORDER BY produtos.descricao;"
   
ElseIf optCodBarra.Value = True Then
sSQL = "SELECT produtos.codigo AS var_cod, produtos.fabricante AS var_fab, produtos.ref AS var_ref, produtos.cod_barra AS var_codbarra, " & _
   "produtos.descricao AS var_desc, produtos.prateleira AS var_prat, produtos.unid_medida AS var_med, " & _
   "produtos.quant_estoque AS var_quant, " & _
   "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
   "FROM produtos " & _
   "WHERE (produtos.cod_barra LIKE '" & txtDescricao & "%') AND (produtos.ativo = 1) ORDER BY produtos.descricao;"
   
ElseIf optRef.Value = True Then
sSQL = "SELECT produtos.codigo AS var_cod, produtos.fabricante AS var_fab, produtos.ref AS var_ref, produtos.cod_barra AS var_codbarra, " & _
   "produtos.descricao AS var_desc, produtos.prateleira AS var_prat, produtos.unid_medida AS var_med, " & _
   "produtos.quant_estoque AS var_quant, " & _
   "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
   "FROM produtos " & _
   "WHERE (produtos.ref LIKE '" & txtDescricao & "%') AND (produtos.ativo = 1) ORDER BY produtos.descricao;"
   
ElseIf optDesc.Value = True Then
sSQL = "SELECT produtos.codigo AS var_cod, produtos.fabricante AS var_fab, produtos.ref AS var_ref, produtos.cod_barra AS var_codbarra, " & _
   "produtos.descricao AS var_desc, produtos.prateleira AS var_prat, produtos.unid_medida AS var_med, " & _
   "produtos.quant_estoque AS var_quant, " & _
   "(SELECT TOP 1 Produtos_Precos.VALOR_VV FROM Produtos_Precos Where produtos_precos.COD_PRODUTO = produtos.codigo order by CODIGO desc) AS venda " & _
   "FROM produtos " & _
   "WHERE (produtos.descricao LIKE '" & txtDescricao & "%') AND (produtos.ativo = 1) ORDER BY produtos.descricao;"
   
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
 '  Dim var_Comp As String     'Compartibilidade
   
   Dim sSQL As String
   Dim r2 As ADODB.Recordset
   
   With Grid
      '.Enabled = False
      .Clear
      .Cols = 9
      .Rows = 2
      
      .ColWidth(0) = 0
      
      If optCodigo.Value = True Then
         .ColWidth(1) = 1450
         .ColWidth(2) = 0
      ElseIf optCodBarra.Value = True Then
         .ColWidth(1) = 0
         .ColWidth(2) = 1450
      ElseIf optDesc.Value = True Then
         .ColWidth(1) = 0
         .ColWidth(2) = 0
      End If
      
    .ColWidth(1) = 1000
    .ColWidth(2) = 1700
    .ColWidth(3) = 4000
    .ColWidth(4) = 1400
    .ColWidth(5) = 800
    .ColWidth(6) = 700
    .ColWidth(7) = 800
    .ColWidth(8) = 1000
      
      .TextMatrix(0, 1) = "CÓDIGO"
      .TextMatrix(0, 2) = "CÓD.BARRA"
      .TextMatrix(0, 3) = "DESCRIÇÃO"
      .TextMatrix(0, 4) = "FABRICANTE"
      .TextMatrix(0, 5) = "UNID."
      .TextMatrix(0, 6) = "LOC."
      .TextMatrix(0, 7) = "ESTOQ."
      .TextMatrix(0, 8) = "PREÇO"
      
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
            'ALINHAMENTO
            .ColAlignment(1) = 1
            .ColAlignment(2) = 1
            
            .TextMatrix(.Rows - 1, 1) = rTabela("var_cod")
            .TextMatrix(.Rows - 1, 2) = rTabela("var_codbarra")
            .TextMatrix(.Rows - 1, 3) = rTabela("var_desc") & " /  " & rTabela("var_ref")
            .TextMatrix(.Rows - 1, 4) = ValidateNull(rTabela("var_fab"))
            .TextMatrix(.Rows - 1, 5) = rTabela("var_med")
            .TextMatrix(.Rows - 1, 6) = ValidateNull(rTabela("var_prat"))
            .TextMatrix(.Rows - 1, 7) = rTabela("var_quant")
            .TextMatrix(.Rows - 1, 8) = Format$(rTabela("venda"), ocMONEY)
            
            'sSQL = "SELECT modelo, ano FROM produtos_comp WHERE (cod_produto = " & .TextMatrix(.Rows - 1, 1) & ");"
            'Set r2 = dbData.OpenRecordset(sSQL)
            
            'Do While Not r2.EOF
            '   var_Comp = var_Comp & r2("modelo") & "(" & r2("ano") & "), "
            '   r2.MoveNext
            'Loop
            
            'If r2.State <> 0 Then r2.Close
            'Set r2 = Nothing
            
            'var_COMP = Mid(var_COMP, 1, Len(var_COMP) - 1) 'Tirar a virgula apos o ultimo
            '.TextMatrix(.Rows - 1, 5) = var_Comp
            'var_Comp = ""
            
            rTabela.MoveNext
            .Rows = .Rows + 1
         Loop
      End If
      
      .Rows = .Rows - 1
      .Redraw = True
      '.Enabled = True
   End With
End Sub
