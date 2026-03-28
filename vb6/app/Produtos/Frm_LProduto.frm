VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form Frm_LProduto 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de produtos"
   ClientHeight    =   5160
   ClientLeft      =   1650
   ClientTop       =   1755
   ClientWidth     =   8835
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8835
   Begin MSFlexGridLib.MSFlexGrid GridProdutos 
      Height          =   4200
      Left            =   60
      TabIndex        =   3
      Top             =   860
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   7408
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.PictureBox PicBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   8850
      TabIndex        =   0
      Top             =   0
      Width           =   8850
      Begin VB.Label dhgf 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de produtos cadastrados no sistema."
         Height          =   435
         Left            =   240
         TabIndex        =   2
         Top             =   570
         Width           =   9015
      End
      Begin VB.Label er 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Informações ao usuário"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   240
         TabIndex        =   1
         Top             =   180
         Width           =   2145
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   60
         X2              =   11880
         Y1              =   1170
         Y2              =   1170
      End
   End
End
Attribute VB_Name = "Frm_LProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim TbProduto As New ADODB.Recordset

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        SendK vbKeyTab
        KeyCode = 0
    ElseIf KeyCode = vbKeyDown Then
        SendK vbKeyTab
        KeyCode = 0
    End If
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
Dim sSQL As String
On Error GoTo erro
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    sSQL = "select *, " & _
           "(SELECT TOP 1 venda FROM produtos_entrada_itens WHERE (codigo_produto = produtos.CODIGO)) As PRVENDA " & _
           "from produtos where descricao like '" & Frm_ItemNF.Text3.Text & "%'"
    RsOpen TbProduto, sSQL
    LimparGridProdutos
    DoEvents
    FormatarGridNotas TbProduto
    Exit Sub

erro:
    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "EkklesiaSoft": Exit Sub
End Sub

Private Sub LimparGridProdutos()
   Dim i As Integer
   
   With GridProdutos
      .Visible = False
      .Redraw = False
      
      .Clear
      .Cols = 4
      .Rows = 2
      
      .ColWidth(0) = 200
      .ColWidth(1) = 1000
      .ColWidth(2) = 5000
      .ColWidth(3) = 1000
      
      'CodigoProduto, NomeProduto, Valor
      .TextMatrix(0, 1) = "CÓDIGO"
      .TextMatrix(0, 2) = "DESCRIÇÃO"
      .TextMatrix(0, 3) = "VENDA"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      'ALINHAMENTO
      .ColAlignment(0) = 1
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = flexAlignLeftCenter
            
      GridProdutos.Col = 0
      
      .Visible = True
      .Redraw = True
   End With
End Sub

Private Sub FormatarGridNotas(rTabela As ADODB.Recordset)
   Dim i As Integer, j As Integer
   
   With GridProdutos
      .Visible = False
      .Redraw = False
      
      .Clear
      .Cols = 4
      .Rows = 2
      
      .ColWidth(0) = 200
      .ColWidth(1) = 1000
      .ColWidth(2) = 5000
      .ColWidth(3) = 1000
      
      'CodigoProduto, NomeProduto, Valor
      .TextMatrix(0, 1) = "CÓDIGO"
      .TextMatrix(0, 2) = "DESCRIÇÃO"
      .TextMatrix(0, 3) = "VENDA"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next
      
      i = 1
      
      'ALINHAMENTO
      .ColAlignment(0) = 1
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = flexAlignLeftCenter
      
      'CodigoProduto, NomeProduto, CST, Unidade, Qtde, Valor, SubTotal
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = rTabela("CODIGO")
            .TextMatrix(.Rows - 1, 2) = rTabela("DESCRICAO")
            .TextMatrix(.Rows - 1, 3) = rTabela("PRVENDA")
            
            rTabela.MoveNext
            .Rows = .Rows + 1
            i = i + 1
         Loop
      End If
      
      .Rows = .Rows - 1
                   
     'GridProdutos.ColWidth(0) = 400
      'GridProdutos.Rows = 11
      GridProdutos.Col = 0
            
      .Visible = True
      .Redraw = True
   End With
End Sub

Private Sub GridProdutos_DblClick()
Dim sSQL As String
On Error GoTo erro
    sSQL = "select *, " & _
       "(SELECT TOP 1 venda FROM produtos_entrada_itens WHERE (codigo_produto = produtos.CODIGO)) As PRVENDA " & _
       "from produtos where codigo = " & GridProdutos.TextMatrix(GridProdutos.Row, 1)
    RsOpen TbProduto, sSQL
    
    If Not TbProduto.EOF And Not TbProduto.BOF Then
        If Not IsNull(TbProduto("codigo")) Then Frm_ItemNF.Text2.Text = TbProduto("codigo")
        If Not IsNull(TbProduto("descricao")) Then Frm_ItemNF.Text3.Text = TbProduto("descricao")
        If Not IsNull(TbProduto("unid_medida")) Then Frm_ItemNF.txtUnid.Text = TbProduto("unid_medida")
        If Not IsNull(TbProduto("CFOP")) Then Frm_ItemNF.txtCFOP.Text = TbProduto("CFOP")
        If Not IsNull(TbProduto("ICMSCST")) Then Frm_ItemNF.txtCST.Text = TbProduto("ICMSCST")
        If Not IsNull(TbProduto("PRVENDA")) Then Frm_ItemNF.txtValor.Text = Format(TbProduto("PRVENDA"), "##,##0.00")
        If Not Vazio(TbProduto("ICMSAliq")) Then Frm_ItemNF.txtICMS.Text = Format(TbProduto("ICMSAliq"), "##,##0.00")
        Unload Me
        Frm_ItemNF.txtICMS.SetFocus
    End If
    Exit Sub

erro:
    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "EkklesiaSoft": Exit Sub
End Sub

Private Sub GridProdutos_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sSQL As String
On Error GoTo erro
    If KeyCode <> 13 Then Exit Sub
    sSQL = "select *, " & _
       "(SELECT TOP 1 venda FROM produtos_entrada_itens WHERE (codigo_produto = produtos.CODIGO)) As PRVENDA " & _
       "from produtos where codigo = " & GridProdutos.TextMatrix(GridProdutos.Row, 1)
    RsOpen TbProduto, sSQL
    
    If Not TbProduto.EOF And Not TbProduto.BOF Then
        If Not IsNull(TbProduto("codigo")) Then Frm_ItemNF.Text2.Text = TbProduto("codigo")
        If Not IsNull(TbProduto("descricao")) Then Frm_ItemNF.Text3.Text = TbProduto("descricao")
        If Not IsNull(TbProduto("unid_medida")) Then Frm_ItemNF.txtUnid.Text = TbProduto("unid_medida")
        If Not IsNull(TbProduto("CFOP")) Then Frm_ItemNF.txtCFOP.Text = TbProduto("CFOP")
        If Not IsNull(TbProduto("ICMSCST")) Then Frm_ItemNF.txtCST.Text = TbProduto("ICMSCST")
        If Not IsNull(TbProduto("PRVENDA")) Then Frm_ItemNF.txtValor.Text = Format(TbProduto("PRVENDA"), "##,##0.00")
        If Not Vazio(TbProduto("ICMSAliq")) Then Frm_ItemNF.txtICMS.Text = Format(TbProduto("ICMSAliq"), "##,##0.00")
        Unload Me
        Frm_ItemNF.txtQuant.SetFocus
    End If
    Exit Sub

erro:
    MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "EkklesiaSoft": Exit Sub
End Sub
