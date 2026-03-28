VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PDV_Consulta 
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
         Value           =   -1  'True
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
      Width           =   10635
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4095
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   10575
      _ExtentX        =   18653
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
   Begin VB.Shape Shape1 
      Height          =   5115
      Left            =   0
      Top             =   0
      Width           =   10755
   End
End
Attribute VB_Name = "PDV_Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String
Dim RS As Recordset
Dim SQL_Conf As String
Dim RS_Conf As Recordset
Dim VERIFICAR_QUANTIDADE As Boolean




Private Sub Form_Activate()
txtDescricao.Text = ""
txtDescricao.SetFocus
End Sub

Private Sub Form_Load()
Call Abrir_BancodeDados
SQL = "SELECT (PRODUTOS.CODIGO) as var_COD, (PRODUTOS.COD_BARRA) as var_CodBarra,(PRODUTOS.DESCRICAO) as var_desc, (PRODUTOS.PRATELEIRA) as var_Prat, (PRODUTOS.UNID_MEDIDA) as var_Med, (PRODUTOS.QUANT_ESTOQUE) as var_Quant, IIF(ISNULL(PRODUTOS_ENTRADA_ITENS.VENDA),0 ,PRODUTOS_ENTRADA_ITENS.VENDA) AS VENDA FROM (PRODUTOS LEFT JOIN ULTIMAS_ENTRADAS ON PRODUTOS.CODIGO   = ULTIMAS_ENTRADAS.CODIGO_PRODUTO) LEFT JOIN PRODUTOS_ENTRADA_ITENS ON (ULTIMAS_ENTRADAS.CODIGO_PRODUTO = PRODUTOS_ENTRADA_ITENS.CODIGO_PRODUTO) AND (ULTIMAS_ENTRADAS.ULTENTRADA = PRODUTOS_ENTRADA_ITENS.CODIGO_ENTRADA) WHERE PRODUTOS.ATIVO = TRUE ORDER BY PRODUTOS.descricao;"
Set RS = BD.OpenRecordset(SQL, dbOpenSnapshot)

Formatar_Grid
End Sub
Private Sub Form_Unload(Cancel As Integer)
RS.Close
BD.Close
End Sub
Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Verifica_QuantEstoque
        If VERIFICAR_QUANTIDADE = True Then
            txtDescricao.SetFocus
            Exit Sub
        Else
            Me.Hide
            PDV.txtCodBarra.Text = Grid.TextMatrix(Grid.Row, 2)
            PDV.Show
        End If
End If
End Sub
Private Sub Verifica_QuantEstoque()
'If txtCodProduto.Text = "" Then Exit Sub
'mostrar o fundo do pdv
Call Abrir_BancodeDados
SQL_Conf = "SELECT ESTOQUE_NEGATIVO, CODIGO FROM CONFIGURACAO WHERE (CODIGO = 1)"
Set RS_Conf = BD.OpenRecordset(SQL_Conf, dbOpenSnapshot)


If RS_Conf!ESTOQUE_NEGATIVO = False Then
    Call Abrir_BancodeDados
    SQL = "SELECT CODIGO, Quant_Estoque FROM PRODUTOS WHERE (CODIGO = " & Grid.TextMatrix(Grid.Row, 1) & ")"
    Set RS = BD.OpenRecordset(SQL, dbOpenSnapshot)
    
    VERIFICAR_QUANTIDADE = False
    'If txtQuant.Text = "" Then txtQuant.Text = 0

        If RS!Quant_Estoque <= 0 Then
            MsgBox "ESSA QUANTIDADE É INVÁLIDA!" & vbCrLf & "SEU ESTOQUE ATUAL É DE 0 (zero) PRODUTO", vbExclamation, "Aviso do Sistema"
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
Private Sub txtDescricao_Change()
Call Abrir_BancodeDados
Dim SQL As String
Dim criterio As String
criterio = Chr$(39) & txtDescricao.Text & "*" & Chr(39)

If optCodigo.Value = True Then
    SQL = "SELECT (PRODUTOS.CODIGO) as var_COD, (PRODUTOS.COD_BARRA) as var_CodBarra,(PRODUTOS.DESCRICAO) as var_desc, (PRODUTOS.PRATELEIRA) as var_Prat, (PRODUTOS.UNID_MEDIDA) as var_Med, (PRODUTOS.QUANT_ESTOQUE) as var_Quant, IIF(ISNULL(PRODUTOS_ENTRADA_ITENS.VENDA),0 ,PRODUTOS_ENTRADA_ITENS.VENDA) AS VENDA FROM (PRODUTOS LEFT JOIN ULTIMAS_ENTRADAS ON PRODUTOS.CODIGO   = ULTIMAS_ENTRADAS.CODIGO_PRODUTO) LEFT JOIN PRODUTOS_ENTRADA_ITENS ON (ULTIMAS_ENTRADAS.CODIGO_PRODUTO = PRODUTOS_ENTRADA_ITENS.CODIGO_PRODUTO) AND (ULTIMAS_ENTRADAS.ULTENTRADA = PRODUTOS_ENTRADA_ITENS.CODIGO_ENTRADA) WHERE PRODUTOS.CODIGO LIKE " & criterio & " and PRODUTOS.ATIVO = TRUE ORDER BY PRODUTOS.descricao;"
    Set RS = BD.OpenRecordset(SQL, dbOpenSnapshot)
ElseIf optCodBarra.Value = True Then
    SQL = "SELECT (PRODUTOS.CODIGO) as var_COD, (PRODUTOS.COD_BARRA) as var_CodBarra,(PRODUTOS.DESCRICAO) as var_desc, (PRODUTOS.PRATELEIRA) as var_Prat, (PRODUTOS.UNID_MEDIDA) as var_Med, (PRODUTOS.QUANT_ESTOQUE) as var_Quant, IIF(ISNULL(PRODUTOS_ENTRADA_ITENS.VENDA),0 ,PRODUTOS_ENTRADA_ITENS.VENDA) AS VENDA FROM (PRODUTOS LEFT JOIN ULTIMAS_ENTRADAS ON PRODUTOS.CODIGO   = ULTIMAS_ENTRADAS.CODIGO_PRODUTO) LEFT JOIN PRODUTOS_ENTRADA_ITENS ON (ULTIMAS_ENTRADAS.CODIGO_PRODUTO = PRODUTOS_ENTRADA_ITENS.CODIGO_PRODUTO) AND (ULTIMAS_ENTRADAS.ULTENTRADA = PRODUTOS_ENTRADA_ITENS.CODIGO_ENTRADA) WHERE PRODUTOS.COD_BARRA LIKE " & criterio & " and PRODUTOS.ATIVO = TRUE ORDER BY PRODUTOS.descricao;"
    Set RS = BD.OpenRecordset(SQL, dbOpenSnapshot)
ElseIf optDesc.Value = True Then
    SQL = "SELECT (PRODUTOS.CODIGO) as var_COD, (PRODUTOS.COD_BARRA) as var_CodBarra,(PRODUTOS.DESCRICAO) as var_desc, (PRODUTOS.PRATELEIRA) as var_Prat, (PRODUTOS.UNID_MEDIDA) as var_Med, (PRODUTOS.QUANT_ESTOQUE) as var_Quant, IIF(ISNULL(PRODUTOS_ENTRADA_ITENS.VENDA),0 ,PRODUTOS_ENTRADA_ITENS.VENDA) AS VENDA FROM (PRODUTOS LEFT JOIN ULTIMAS_ENTRADAS ON PRODUTOS.CODIGO   = ULTIMAS_ENTRADAS.CODIGO_PRODUTO) LEFT JOIN PRODUTOS_ENTRADA_ITENS ON (ULTIMAS_ENTRADAS.CODIGO_PRODUTO = PRODUTOS_ENTRADA_ITENS.CODIGO_PRODUTO) AND (ULTIMAS_ENTRADAS.ULTENTRADA = PRODUTOS_ENTRADA_ITENS.CODIGO_ENTRADA) WHERE PRODUTOS.descricao LIKE " & criterio & " and PRODUTOS.ATIVO = TRUE ORDER BY PRODUTOS.descricao;"
    Set RS = BD.OpenRecordset(SQL, dbOpenSnapshot)
End If

Formatar_Grid
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
Private Sub Formatar_Grid()
With Grid
    '.Enabled = False
    .Clear
    .Cols = 8
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
    .ColWidth(2) = 1450
    End If
    
    .ColWidth(3) = 5100
    .ColWidth(4) = 850
    .ColWidth(5) = 850
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
    Dim X As Integer
    For X = 0 To .Cols - 1
    .Col = X
    .Row = 0
    .CellFontBold = True
    Next X
    
    'centralizar o titulo
    Dim f As Integer
    For f = 0 To .Cols - 1
    .Row = 0
    .Col = f
    .CellAlignment = flexAlignCenterCenter
    Next f
    
    .Redraw = False
    
    Do Until RS.EOF
   
    'ALINHAMENTO
    .ColAlignment(2) = 1
    
    If Not IsNull(RS!var_COD) Then .TextMatrix(.Rows - 1, 1) = RS!var_COD
    If Not IsNull(RS!var_CodBarra) Then .TextMatrix(.Rows - 1, 2) = RS!var_CodBarra
    If Not IsNull(RS!var_desc) Then .TextMatrix(.Rows - 1, 3) = RS!var_desc
    If Not IsNull(RS!var_Med) Then .TextMatrix(.Rows - 1, 4) = RS!var_Med
    If Not IsNull(RS!var_Prat) Then .TextMatrix(.Rows - 1, 5) = RS!var_Prat
    If Not IsNull(RS!var_Quant) Then .TextMatrix(.Rows - 1, 6) = RS!var_Quant
    If Not IsNull(RS!VENDA) Then .TextMatrix(.Rows - 1, 7) = Format(RS!VENDA, "##,##0.00")
    RS.MoveNext
    .Rows = .Rows + 1
        
    Loop
    
    .Rows = .Rows - 1
    .Redraw = True
    '.Enabled = True
End With
End Sub
