VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form Frm_ItemNF 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Itens da nota fiscal"
   ClientHeight    =   5775
   ClientLeft      =   1605
   ClientTop       =   2280
   ClientWidth     =   9480
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   9480
   Begin MSFlexGridLib.MSFlexGrid GridNotasItens 
      Height          =   2715
      Left            =   180
      TabIndex        =   21
      Top             =   2400
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4789
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox txtCFOP 
      Height          =   285
      Left            =   4881
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1200
      Width           =   450
   End
   Begin VB.TextBox txtICMS 
      Height          =   285
      Left            =   7819
      MaxLength       =   10
      TabIndex        =   7
      Top             =   1185
      Width           =   450
   End
   Begin VB.TextBox txtQuant 
      Height          =   285
      Left            =   6035
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1185
      Width           =   570
   End
   Begin VB.TextBox txtSubTotal 
      Height          =   285
      Left            =   8400
      MaxLength       =   8
      TabIndex        =   8
      Top             =   1185
      Width           =   960
   End
   Begin VB.TextBox txtValor 
      Height          =   285
      Left            =   6732
      MaxLength       =   8
      TabIndex        =   6
      Top             =   1185
      Width           =   960
   End
   Begin VB.TextBox txtUnid 
      Height          =   285
      Left            =   4304
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1185
      Width           =   450
   End
   Begin VB.TextBox txtCST 
      Height          =   285
      Left            =   5458
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1185
      Width           =   450
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   937
      MaxLength       =   60
      TabIndex        =   1
      Top             =   1185
      Width           =   3240
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   105
      MaxLength       =   13
      TabIndex        =   0
      Top             =   1185
      Width           =   705
   End
   Begin VB.PictureBox PicBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Index           =   0
      Left            =   -120
      ScaleHeight     =   825
      ScaleWidth      =   9720
      TabIndex        =   9
      Top             =   0
      Width           =   9720
      Begin VB.TextBox txtCodNota 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   7920
         MaxLength       =   10
         TabIndex        =   23
         Top             =   360
         Width           =   1530
      End
      Begin VB.Label lblNumeroNota 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código Nota"
         Height          =   195
         Left            =   8550
         TabIndex        =   22
         Top             =   120
         Width           =   885
      End
      Begin VB.Label dhgf 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista os itens a serem impressos na Nota Fiscal"
         Height          =   315
         Left            =   600
         TabIndex        =   11
         Top             =   480
         Width           =   8175
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
         TabIndex        =   10
         Top             =   120
         Width           =   2145
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CFOP"
      Height          =   195
      Left            =   4875
      TabIndex        =   20
      Top             =   960
      Width           =   420
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ICMS"
      Height          =   195
      Left            =   7819
      TabIndex        =   19
      Top             =   960
      Width           =   390
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SubTotal"
      Height          =   195
      Left            =   8400
      TabIndex        =   18
      Top             =   960
      Width           =   645
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
      Height          =   195
      Left            =   6732
      TabIndex        =   17
      Top             =   960
      Width           =   360
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qde"
      Height          =   195
      Left            =   6035
      TabIndex        =   16
      Top             =   960
      Width           =   300
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Und"
      Height          =   195
      Left            =   4304
      TabIndex        =   15
      Top             =   960
      Width           =   300
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CST"
      Height          =   195
      Left            =   5460
      TabIndex        =   14
      Top             =   960
      Width           =   315
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição do produto"
      Height          =   195
      Left            =   915
      TabIndex        =   13
      Top             =   960
      Width           =   1530
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      Height          =   195
      Left            =   105
      TabIndex        =   12
      Top             =   960
      Width           =   495
   End
End
Attribute VB_Name = "Frm_ItemNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tb As New ADODB.Recordset
Dim Titulo, Book As Variant, NomeTabela

Sub Load_Data_Itens()
Dim sSQL As String, seq As Integer
    sSQL = "SELECT MAX(Item) r FROM NotaFiscalItens WHERE CodigoNota = " & Val(Frm_NF.txtCodNota.Text)
    seq = SQLExecutaRetorno(sSQL, "r", 0) + 1
    Tb("CodigoNota") = Format(Frm_NF.txtCodNota.Text, "@")
    Tb("Item") = seq
    Tb("CodigoProduto") = Format(Text2, "@")
    Tb("NomeProduto") = UCase(Format(Text3, "@"))
    Tb("CFOP") = Format(txtCFOP, "@")
    Tb("CST") = Right(Format(txtCST, "@"), 3)
    Tb("UnidadeComercial") = UCase(Format(Text5, "@"))
    Tb("ValorUnitarioComercializacao") = CDbl(Format(txtValor, "@"))
    Tb("ValorTotalBruto") = CDbl(Format(txtSubTotal, "@"))
    Tb("QuantidadeComercial") = Format(txtQuant, "@")
    If txtICMS.Text <> "" Then Tb("pICMS") = CDbl(Format(txtICMS, "@"))
    If txtICMS.Text <> "" Then Tb("vBC") = CDbl(Format(txtSubTotal, "@"))
    If txtICMS.Text <> "" Then Tb("vICMS") = Round(Format(txtSubTotal, "@") * (Format(txtICMS, "@") / 100), 2)
End Sub

Public Sub Load_Controls()
    Long1 = Format(Tb("Item"), "@")
    txtCodNota = Format(Tb("CodigoNota"), "@")
    Text2 = Format(Tb("CodigoProduto"), "@")
    Text3 = Format(Tb("NomeProduto"), "@")
    txtCST = Format(Tb("CST"), "@")
    Text5 = Format(Tb("UnidadeComercial"), "@")
    txtValor = Format(Tb("ValorUnitarioComercializacao"), "@")
    txtSubTotal = Format(Tb("ValorTotalBruto"), "@")
    txtQuant = Format(Tb("QuantidadeComercial"), "@")
    txtICMS = Format(Tb("pICMS"), "@")
    txtCFOP = Format(Tb("CFOP"), "@")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendK vbKeyTab
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
Dim sSQL As String, enviada As Boolean
Dim totalRegistros As Long
    
    On Error GoTo ErrLoad
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    sSQL = "SELECT * FROM NotaFiscalItens WHERE CodigoNota = " & Val(Frm_NF.txtCodNota.Text)
    RsOpen Tb, sSQL
    
    If Tb.RecordCount > 0 Then totalRegistros = Tb.RecordCount
    
    txtCodNota.Text = Frm_NF.txtCodNota.Text
    
    enviada = SQLExecutaRetorno("SELECT Enviada FROM NotaFiscal WHERE CodigoNota = " & Val(Frm_NF.txtCodNota.Text), "Enviada", 0)
    
    If enviada Then
       Text2.Enabled = False
       Text3.Enabled = False
       txtCFOP.Enabled = False
       Text5.Enabled = False
       txtValor.Enabled = False
       txtSubTotal.Enabled = False
       txtQuant.Enabled = False
       txtICMS.Enabled = False
       txtCFOP.Enabled = False
    End If
    
    LimparGridItensNota
    DoEvents
    FormatarGridItensNota Tb
    Exit Sub
    
ErrLoad:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
   RsOpen Frm_NF.TbNotas, "SELECT *, " & _
                    "(CASE WHEN Denegada = 1 THEN 'Denegada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 0 THEN 'Enviada' ELSE (CASE WHEN Enviada = 1 AND Cancelada = 1 THEN 'Cancelada' ELSE 'Em Digitação' END) END) END) AS Status " & _
                    "FROM NotaFiscal"
    'Frm_NF.FormatarGridItensNota Frm_NF.TbNotas  'desativei ver depois
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sSQL As String
Dim TbProduto As New ADODB.Recordset, TbProdutoPreco As New ADODB.Recordset
On Error GoTo erro
If KeyCode = 13 And Not Vazio(Text2.Text) Then
    RsOpen TbProduto, "select * from produtos where codigo = " & Text2.Text
    If TbProduto.EOF And TbProduto.BOF Then
        Text2.SetFocus
    Else
        sSQL = "SELECT TOP 1 venda FROM produtos_entrada_itens WHERE (codigo_produto = " & TbProduto("codigo") & ") ORDER BY codigo DESC;"
        RsOpen TbProdutoPreco, sSQL
        If Not Vazio(TbProduto("codigo")) Then Text2.Text = TbProduto("codigo")
        If Not Vazio(TbProduto("descricao")) Then Text3.Text = TbProduto("descricao")
        If Not Vazio(TbProduto("unid_medida")) Then Text5.Text = TbProduto("unid_medida")
        If Not Vazio(TbProduto("CFOP")) Then txtCFOP.Text = TbProduto("CFOP")
        If Not Vazio(TbProduto("ICMSCST")) Then txtCFOP.Text = TbProduto("ICMSCST")
        If Not Vazio(TbProdutoPreco("venda")) Then txtValor.Text = Format(TbProdutoPreco("venda"), "##,##0.00")
        If Not Vazio(TbProduto("ICMSAliq")) Then txtICMS.Text = Format(TbProduto("ICMSAliq"), "##,##0.00")
        txtCFOP.SetFocus
    End If
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "SistemasNFe": Exit Sub
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Frm_LProduto.Show 1
End Sub

Private Sub Text3_LostFocus()
Text3.Text = UCase(Text3.Text)
End Sub

Private Sub txtUnid_LostFocus()
txtUnid.Text = UCase(txtUnid.Text)
End Sub

Private Sub txtValor_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo erro
If KeyCode = 13 Then
    If txtQuant.Text = "" Then MsgBox "O campo de quantidade esta vazio, confira.", vbCritical, "SistemasNFe": Exit Sub
    If txtValor.Text = "" Then MsgBox "O campo de valor unitário esta vazio, confira.", vbCritical, "SistemasNFe": Exit Sub
    txtSubTotal.Text = Format(CCur(txtQuant.Text) * CCur(txtValor.Text), "##,##0.00")
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "SistemasNFe": Exit Sub

End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii = Asc(".") Then KeyAscii = Asc(",")
ElseIf KeyAscii = Asc(",") Then KeyAscii = Asc(",")
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "SistemasNFe": Exit Sub

End Sub

Private Sub txtValor_LostFocus()
txtValor.Text = UCase(txtValor.Text)
End Sub

Private Sub txtSubTotal_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii = Asc(".") Then KeyAscii = Asc(",")
ElseIf KeyAscii = Asc(",") Then KeyAscii = Asc(",")
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "SistemasNFe": Exit Sub

End Sub

Private Sub txtQuant_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii = Asc(".") Then KeyAscii = Asc(",")
ElseIf KeyAscii = Asc(",") Then KeyAscii = Asc(",")
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "SistemasNFe": Exit Sub

End Sub

Private Sub txtICMS_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sSQL As String, vTotal As Double
On Error GoTo erro
If KeyCode = 13 Then
    'RsOpen Tb, "select * from NotaFiscalItens"
    vgDb.BeginTrans
    Tb.AddNew
    Load_Data_Itens
    Tb.Update
    vgDb.CommitTrans
    Limpa_Tudo Me
    sSQL = "SELECT ISNULL(SUM(ValorTotalBruto), 0) r FROM NotaFiscalItens WHERE CodigoNota = " & Val(Frm_NF.txtCodNota.Text)
    vTotal = SQLExecutaRetorno(sSQL, "r", 0)
    sSQL = "UPDATE NotaFiscal SET ValorProdutos = " & FSQL(vTotal, 2) & ", ValorNota = " & FSQL(vTotal, 2) & " WHERE CodigoNota = " & Val(Frm_NF.txtCodNota.Text)
    SQLExecuta sSQL
    sSQL = "SELECT * FROM NotaFiscalItens WHERE CodigoNota = " & Val(Frm_NF.txtCodNota.Text)
    RsOpen Tb, sSQL
    FormatarGridItensNota Tb
    KeyCode = 0
    Text2.SetFocus
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "SistemasNFe": Exit Sub
End Sub

Private Sub txtICMS_KeyPress(KeyAscii As Integer)
On Error GoTo erro
If KeyAscii = 8 Then
ElseIf KeyAscii = Asc(".") Then KeyAscii = Asc(",")
ElseIf KeyAscii = Asc(",") Then KeyAscii = Asc(",")
ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
Exit Sub
erro:
MsgBox "Erro no sistema: " & Err.Number & " - " & Err.Description, vbCritical, "SistemasNFe": Exit Sub
End Sub

Private Sub LimparGridItensNota()
   Dim i As Integer
   
   With GridNotasItens
      .Visible = False
      .Redraw = False
      
      .Clear
      .Cols = 8
      .Rows = 2
      
      .ColWidth(0) = 200
      .ColWidth(1) = 2000
      .ColWidth(2) = 1000
      .ColWidth(3) = 1000
      .ColWidth(4) = 1000
      .ColWidth(5) = 1200
      .ColWidth(6) = 1500
      .ColWidth(7) = 1500
      
      'CodigoProduto, NomeProduto, CST, Unidade, Qtde, Valor, SubTotal
      .TextMatrix(0, 1) = "CÓDIGO"
      .TextMatrix(0, 2) = "DESCRIÇÃO"
      .TextMatrix(0, 3) = "CST"
      .TextMatrix(0, 4) = "UND"
      .TextMatrix(0, 5) = "QTDE"
      .TextMatrix(0, 6) = "VALOR"
      .TextMatrix(0, 7) = "SUBTOTAL"
      
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
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 2
      .ColAlignment(6) = 2
      .ColAlignment(7) = 2
      .Rows = .Rows + 1
      
      i = i + 1
      .Rows = .Rows - 1
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 2
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 3
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
      
      'GridNotasItens.ColWidth(0) = 400
      'GridNotasItens.Rows = 11
      GridNotasItens.Col = 0
      
      .Visible = True
      .Redraw = True
   End With
End Sub

Private Sub FormatarGridItensNota(rTabela As ADODB.Recordset)
   Dim i As Integer, j As Integer
   
   With GridNotasItens
      .Visible = False
      .Redraw = False
      
      .Clear
      .Cols = 8
      .Rows = 2
      
      .ColWidth(0) = 200
      .ColWidth(1) = 800
      .ColWidth(2) = 2000
      .ColWidth(3) = 1000
      .ColWidth(4) = 1000
      .ColWidth(5) = 1200
      .ColWidth(6) = 1500
      .ColWidth(7) = 1500
      
      'CodigoProduto, NomeProduto, CST, Unidade, Qtde, Valor, SubTotal
      .TextMatrix(0, 1) = "CÓDIGO"
      .TextMatrix(0, 2) = "DESCRIÇÃO"
      .TextMatrix(0, 3) = "CST"
      .TextMatrix(0, 4) = "UND"
      .TextMatrix(0, 5) = "QTDE"
      .TextMatrix(0, 6) = "VALOR"
      .TextMatrix(0, 7) = "SUBTOTAL"
      
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
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 2
      .ColAlignment(6) = 2
      .ColAlignment(7) = 2
      
      'CodigoProduto, NomeProduto, CST, Unidade, Qtde, Valor, SubTotal
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.Rows - 1, 1) = rTabela("CodigoProduto")
            .TextMatrix(.Rows - 1, 2) = rTabela("NomeProduto")
            .TextMatrix(.Rows - 1, 3) = rTabela("CST")
            .TextMatrix(.Rows - 1, 4) = rTabela("UnidadeComercial")
            .TextMatrix(.Rows - 1, 5) = Format(rTabela("QuantidadeComercial"), ocMONEY)
            .TextMatrix(.Rows - 1, 6) = Format(rTabela("ValorUnitarioComercializacao"), ocMONEY)
            .TextMatrix(.Rows - 1, 7) = Format(rTabela("ValorTotalBruto"), ocMONEY)
            
            rTabela.MoveNext
            .Rows = .Rows + 1
            i = i + 1
         Loop
      End If
      
      .Rows = .Rows - 1
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 2
         .CellFontBold = True
      Next
      
      'MUDAR COR DE FONTE DA COLUNA
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 3
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next
              
     'GridNotasItens.ColWidth(0) = 400
      'GridNotasItens.Rows = 11
      GridNotasItens.Col = 0
            
      .Visible = True
      .Redraw = True
   End With
End Sub

