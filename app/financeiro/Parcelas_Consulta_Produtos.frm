VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Parcelas_Consulta_Produtos 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ITENS DO VENDA"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12060
   Icon            =   "Parcelas_Consulta_Produtos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   12060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3795
      Left            =   60
      TabIndex        =   9
      Top             =   1020
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   6694
      _Version        =   393216
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   11925
      TabIndex        =   7
      Top             =   60
      Width           =   11955
      Begin VB.TextBox lblOrdem2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   525
         Left            =   10440
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "000000"
         Top             =   420
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtCodPedido 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   525
         Left            =   10440
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "000000"
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label lblOrdem1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordem de Serviço:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         TabIndex        =   17
         Top             =   420
         Width           =   2130
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   9540
         TabIndex        =   11
         Top             =   60
         Width           =   900
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUTOS E SERVIÇOS"
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
         TabIndex        =   8
         Top             =   300
         Width           =   3765
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   240
         Picture         =   "Parcelas_Consulta_Produtos.frx":23D2
         Top             =   0
         Width           =   1140
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   9240
      ScaleHeight     =   1305
      ScaleWidth      =   2745
      TabIndex        =   0
      Top             =   4860
      Width           =   2775
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acréscimo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   285
         TabIndex        =   20
         Top             =   660
         Width           =   945
      End
      Begin VB.Label lblTotalAcresc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1320
         TabIndex        =   19
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   60
         Width           =   1335
      End
      Begin VB.Label lblTotalDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   510
         TabIndex        =   4
         Top             =   60
         Width           =   720
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   405
         TabIndex        =   3
         Top             =   360
         Width           =   825
      End
      Begin VB.Label lblTotalGeral 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   780
         TabIndex        =   1
         Top             =   960
         Width           =   450
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid_Parcelas 
      Height          =   1335
      Left            =   60
      TabIndex        =   12
      Top             =   4860
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   2355
      _Version        =   393216
      BackColor       =   12648447
      SelectionMode   =   1
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
   Begin VB.Label lblRecebedorTit 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recebedor(a):"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8220
      TabIndex        =   18
      Top             =   5460
      Width           =   945
   End
   Begin VB.Label lblRecebedor 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NENHUM"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   8520
      TabIndex        =   16
      Top             =   5700
      Width           =   630
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor(a):"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8280
      TabIndex        =   14
      Top             =   4860
      Width           =   870
   End
   Begin VB.Label lblFuncionario 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NENHUM"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   8520
      TabIndex        =   13
      Top             =   5100
      Width           =   630
   End
End
Attribute VB_Name = "Parcelas_Consulta_Produtos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cCfg As ConfigItem
Dim tipoEmpresa As Integer
Dim vTipoOS As String
Dim sSQL As String
Dim r As ADODB.Recordset
Public Sub loadPedidos(ByVal Pedido As Long, ByVal Tipo As String)
Set oCfg = sysConfig("TIPO_OS")
vTipoOS = oCfg.Value
Set oCfg = Nothing

Dim sSQL As String
Dim r As ADODB.Recordset
Dim r2 As ADODB.Recordset
Dim totalRegistros As Long

If Tipo = "OFICINA" Then Tipo = "OS"
'contar a quantidades de produtos na consulta, para saber se vai agrupar com outra tabela
sSQL = "SELECT 'PRODUTO' AS tipo_item, produtos.descricao as var_desc, tamanho as var_Tam, fabricante as var_Fab, quantidade, preco, pedidos_itens.total, produtos.codigo,  pedidos_itens.subtotal as var_Subtotal, pedidos_itens.desconto, '' as var_CodOS " & _
      "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
      "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
      "WHERE (pedidos_itens.cod_pedido = " & Pedido & ")"
Set r = dbData.OpenRecordset(sSQL, totalRegistros)

If totalRegistros >= 1 Then
    sSQL = "SELECT 'PRODUTO' AS tipo_item, produtos.descricao as var_desc, tamanho as var_Tam, fabricante as var_Fab, quantidade, preco, pedidos_itens.total, produtos.codigo,  pedidos_itens.subtotal as var_Subtotal, pedidos_itens.desconto, '' as var_CodOS " & _
          "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
          "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
          "WHERE (pedidos_itens.cod_pedido = " & Pedido & ")"
    'Debug.Print sSQL
    
    If UCase(Tipo) = "OS" Then  'If UCase(Tipo) = "OS" Then 'mudei e testar nas outras coisas se interfere
       'If vTipoOS = "Automóveis" Then
       If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Informática" Or vTipoOS = "Celular" Then
       sSQL = sSQL & " UNION "
       sSQL = sSQL & "SELECT 'SERVIÇO' AS tipo_item, descricao as var_desc, '' as var_Tam, '' as var_Fab, quantidade, preco, OS_Servicos_Auto.total, codigo,  OS_Servicos_Auto.subtotal as var_Subtotal, OS_Servicos_Auto.desconto, OS_Servicos_Auto.cod_os as var_CodOS " & _
              "FROM  OS_Servicos_Auto INNER JOIN OS ON OS_Servicos_Auto.cod_os = OS.COD_OS WHERE (OS.COD_PEDIDO = " & Pedido & ")"
       'Debug.Print sSQL
       Else
       End If
    End If
    
    If UCase(Tipo) = "RECEBER" Then
       sSQL = sSQL & " UNION "
     'sSQL = sSQL & "SELECT 'RECEBER' AS tipo_item, DESCRICAO as var_desc, '' as var_Tam, '' as var_Fab, quantidade, preco, ISNULL((quantidade * preco), 0) AS total, codigo, '', '', '' as var_CodOS " & _
             "FROM pedidos_itens WHERE (cod_pedido = " & Pedido & ")"
    'sSQL = "SELECT 'RECEBER' AS tipo_item, produtos.descricao as var_desc, tamanho as var_Tam, fabricante as var_Fab, quantidade, preco, pedidos_itens.total, produtos.codigo,  pedidos_itens.subtotal, pedidos_itens.desconto, '' as var_CodOS " & _
          "FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _
          "LEFT JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido " & _
          "WHERE (pedidos_itens.cod_pedido = " & Pedido & ")"
    
    sSQL = "SELECT 'RECEBER' AS tipo_item, DESCRICAO as var_desc, '' as var_Tam, '' as var_Fab, quantidade, preco, total, a_receber_itens.codigo, total as var_Subtotal, '' as desconto, '' as var_CodOS FROM a_receber_itens WHERE (cod_pedido = " & Pedido & ")" & _
       "UNION ALL "
    sSQL = sSQL & "SELECT 'RECEBER' AS tipo_item, produtos.DESCRICAO as var_desc, '' as var_Tam, '' as var_Fab, quantidade, pedidos_itens.PRECO, total, pedidos_itens.codigo, total as var_Subtotal, '' as desconto, '' as var_CodOS FROM pedidos_itens INNER JOIN produtos ON pedidos_itens.COD_PRODUTO = produtos.CODIGO WHERE (pedidos_itens.cod_pedido = " & Pedido & ")"
    
    
    End If
    
    If UCase(Tipo) = "ALUGUEL" Then
    'If varTipoConsulta = "ALUGUEL" Then
       sSQL = sSQL & " UNION "
     sSQL = sSQL & "SELECT 'ALUGUEL' AS tipo_item, Aluguel_Cadastro_Equipamento.DESCRICAO as var_desc, '' as var_Tam, '' as var_Fab, Aluguel_Cadastro_Itens.QUANT_ALUGADA AS quantidade, Aluguel_Cadastro_Itens.TOTAL_ALUGADA as preco, Aluguel_Cadastro_Itens.VALOR_FINAL AS total, Aluguel_Cadastro.codigo,  Aluguel_Cadastro_Itens.DESCONTO, Aluguel_Cadastro_Itens.SUBTOTAL as var_Subtotal, '' as var_CodOS " & _
             "FROM Aluguel_Cadastro_Itens INNER JOIN Aluguel_Cadastro_Equipamento ON Aluguel_Cadastro_Itens.COD_EQUIP = Aluguel_Cadastro_Equipamento.COD_EQUIP INNER JOIN Aluguel_Cadastro ON Aluguel_Cadastro_Itens.COD_LOCACAO = Aluguel_Cadastro.CODIGO WHERE (Aluguel_Cadastro.Cod_Pedido = " & Pedido & ")"
    End If
Else
    If UCase(Tipo) = "OS" Then
       If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Informática" Or vTipoOS = "Celular" Then
       sSQL = "SELECT 'SERVIÇO' AS tipo_item, descricao as var_desc, '' as var_Tam, '' as var_Fab, quantidade, preco, OS_Servicos_Auto.total, codigo,  OS_Servicos_Auto.subtotal as var_Subtotal, OS_Servicos_Auto.desconto, OS_Servicos_Auto.cod_os as var_CodOS " & _
              "FROM  OS_Servicos_Auto INNER JOIN OS ON OS_Servicos_Auto.cod_os = OS.COD_OS WHERE (OS.COD_PEDIDO = " & Pedido & ")"
       End If
    End If
    
    If UCase(Tipo) = "RECEBER" Then
        sSQL = "SELECT 'RECEBER' AS tipo_item, DESCRICAO as var_desc, '' as var_Tam, '' as var_Fab, quantidade, preco, total, a_receber_itens.codigo, total as var_Subtotal, '' as desconto, '' as var_CodOS FROM a_receber_itens WHERE (cod_pedido = " & Pedido & ")" & _
           "UNION ALL "
        sSQL = sSQL & "SELECT 'RECEBER' AS tipo_item, produtos.DESCRICAO as var_desc, '' as var_Tam, '' as var_Fab, quantidade, pedidos_itens.PRECO, total, pedidos_itens.codigo, total as var_Subtotal, '' as desconto, '' as var_CodOS FROM pedidos_itens INNER JOIN produtos ON pedidos_itens.COD_PRODUTO = produtos.CODIGO WHERE (pedidos_itens.cod_pedido = " & Pedido & ")"
    End If
    
    If UCase(Tipo) = "ALUGUEL" Then
     sSQL = "SELECT 'ALUGUEL' AS tipo_item, Aluguel_Cadastro_Equipamento.DESCRICAO as var_desc, '' as var_Tam, '' as var_Fab, Aluguel_Cadastro_Itens.QUANT_ALUGADA AS quantidade, Aluguel_Cadastro_Itens.TOTAL_ALUGADA as preco, Aluguel_Cadastro_Itens.VALOR_FINAL AS total, Aluguel_Cadastro.codigo,  Aluguel_Cadastro_Itens.DESCONTO, Aluguel_Cadastro_Itens.SUBTOTAL as var_Subtotal, '' as var_CodOS " & _
             "FROM Aluguel_Cadastro_Itens INNER JOIN Aluguel_Cadastro_Equipamento ON Aluguel_Cadastro_Itens.COD_EQUIP = Aluguel_Cadastro_Equipamento.COD_EQUIP INNER JOIN Aluguel_Cadastro ON Aluguel_Cadastro_Itens.COD_LOCACAO = Aluguel_Cadastro.CODIGO WHERE (Aluguel_Cadastro.Cod_Pedido = " & Pedido & ")"
    End If
End If
'Debug.Print sSQL
    Set r = dbData.OpenRecordset(sSQL)

FormatarGrid_Itens r

If r.State <> 0 Then r.Close
Set r = Nothing

If UCase(Tipo) = "OS" Then
    sSQL = "SELECT COD_OS " & _
          "FROM  OS WHERE (COD_PEDIDO = " & Pedido & ")"
    Set r = dbData.OpenRecordset(sSQL)
    'Debug.Print sSQL
    lblOrdem1.Visible = True
    lblOrdem2.Visible = True
    If Not r.BOF Then
        lblOrdem2.Text = Format(ValidateNull(r("COD_OS")), "000000")
        'lblOrdem2.Text = Format(r("var_CodOS"), "000000")
    End If
    If r.State <> 0 Then r.Close
    Set r = Nothing
Else
    lblOrdem1.Visible = False
    lblOrdem2.Visible = False
    lblOrdem2.Text = Format(0, "000000")
End If

Call MostrarParcelas
Call MostrarFuncionario

'pegar os totais
sSQL = "SELECT SUBTOTAL,TOTAL, ValorDescReal, ValorAcrescReal FROM pedidos WHERE (cod_pedido = " & Pedido & ");"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
   lblTotal.Caption = Format(r("subtotal"), ocMONEY)
   lblTotalGeral.Caption = Format(r("total"), ocMONEY)
   
   'If r("tipo_desc") = "R" Then
      lblTotalDesc.Caption = Format(r("ValorDescReal"), ocMONEY)
      lblTotalAcresc.Caption = Format(r("ValorAcrescReal"), ocMONEY)
   'Else
   '   lblTotalDesc.Caption = FormatNumber(r("valor_desc")) & "%"
   'End If
Else
    lblTotal.Caption = Format(0, ocMONEY)
    lblTotalGeral.Caption = Format(0, ocMONEY)
    lblTotalDesc.Caption = Format(0, ocMONEY)
    lblTotalAcresc.Caption = Format(0, ocMONEY)
End If

If r.State <> 0 Then r.Close
Set r = Nothing

txtCodPedido.Text = Format(Pedido, "000000")
End Sub

Private Sub MostrarFuncionario()
If txtCodPedido.Text = "" Then Exit Sub

sSQL = "SELECT funcionario.CODIGO, funcionario.NOME " & _
       "FROM funcionario LEFT OUTER JOIN pedidos ON funcionario.CODIGO = pedidos.COD_FUNCIONARIO WHERE (pedidos.COD_PEDIDO = " & txtCodPedido.Text & ");"

Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
   lblFuncionario.Caption = r("NOME")
Else
    lblFuncionario.Caption = "NENHUM"
End If

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub
Private Sub MostrarParcelas()
If txtCodPedido.Text = "" Then Exit Sub

sSQL = "SELECT DATA, PAGAMENTO, VALOR, JUROS, DESCONTO, VALOR_FINAL, FORMA_PGTO, CODCAIXA, CAIXA, " & _
        "CASE status WHEN 0 THEN 'Á PAGAR' ELSE 'PAGO' END AS varStatus,  " & _
        "(SELECT ISNULL(SUM(VALOR_HAVER), 0) FROM parcelas_haver WHERE (COD_PARCELA = parcelas.CODIGO)) AS varSomaHaveres " & _
        "FROM parcelas WHERE (cod_pedido = " & txtCodPedido.Text & ");"

Set r = dbData.OpenRecordset(sSQL)

FormatarGrid_Parcelas r

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub
Private Sub FormatarGrid_Parcelas(rTabela As ADODB.Recordset)
Dim i As Integer

With grid_Parcelas
   .Clear
   .Cols = 12
   .rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 800
   .ColWidth(2) = 900
   .ColWidth(3) = 800
   .ColWidth(4) = 800
   .ColWidth(5) = 800
   .ColWidth(6) = 900
   .ColWidth(7) = 850
   .ColWidth(8) = 800
   .ColWidth(9) = 1000
   .ColWidth(10) = 800
   .ColWidth(11) = 800
   
   'DATA, PAGAMENTO, VALOR_FINAL, STATUS, FORMA_PGTO, CODCAIXA, CAIXA
   
   .TextMatrix(0, 1) = "Venc."
   .TextMatrix(0, 2) = "Vlr Bruto"
   .TextMatrix(0, 3) = "Juros"
   .TextMatrix(0, 4) = "Desc."
   .TextMatrix(0, 5) = "Haver"
   .TextMatrix(0, 6) = "Vrl Liqu."
   .TextMatrix(0, 7) = "Status"
   .TextMatrix(0, 8) = "Pgto"
   .TextMatrix(0, 9) = "Forma."
   .TextMatrix(0, 10) = "Caixa"
   .TextMatrix(0, 11) = "Cód.Cx"
   
   'colocar os cabeçalho em negrito
   For i = 0 To .Cols - 1
      .Col = i
      .Row = 0
      .CellFontBold = True
   Next
   
   'ALINHAMENTO
   '.ColAlignment(2) = 1
   
   'centralizar o titulo
   For i = 0 To .Cols - 1
      .Row = 0
      .Col = i
      .CellAlignment = flexAlignCenterCenter
   Next i
   'VALOR, JUROS, DESCONTO, parcelas_haver
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         .TextMatrix(.rows - 1, 1) = Format(rTabela("DATA"), "dd/mm/yy")
         .TextMatrix(.rows - 1, 2) = FormatNumber(rTabela("VALOR"), 2)
         .TextMatrix(.rows - 1, 3) = FormatNumber(rTabela("JUROS"), 2)
         .TextMatrix(.rows - 1, 4) = FormatNumber(rTabela("DESCONTO"), 2)
         .TextMatrix(.rows - 1, 5) = FormatNumber(rTabela("varSomaHaveres"), 2)
         .TextMatrix(.rows - 1, 6) = FormatNumber(rTabela("VALOR_FINAL"), 2)
         .TextMatrix(.rows - 1, 7) = rTabela("varSTATUS")
         .TextMatrix(.rows - 1, 8) = Format(rTabela("PAGAMENTO"), "dd/mm/yy")
         .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("FORMA_PGTO"))
         .TextMatrix(.rows - 1, 10) = ValidateNull(rTabela("CAIXA"))
         .TextMatrix(.rows - 1, 11) = ValidateNull(rTabela("CODCAIXA"))
         rTabela.MoveNext
         .rows = .rows + 1
      Loop
   End If
   
   .rows = .rows - 1
End With
End Sub


Private Sub Form_Load()
Set cCfg = sysConfig("TIPO_EMPRESA")
tipoEmpresa = cCfg.Value
Set cCfg = Nothing
End Sub
Private Sub FormatarGrid_Itens(rTabela As ADODB.Recordset)
   Dim i As Integer
   
   With Grid
      .Clear
      .Cols = 8
      .rows = 2
      
      .ColWidth(0) = 0
      .ColWidth(1) = 950
      .ColWidth(2) = 5700
      .ColWidth(3) = 1000
      .ColWidth(4) = 900
      .ColWidth(5) = 1100
      .ColWidth(6) = 900
      .ColWidth(7) = 1000
      
      .TextMatrix(0, 1) = "TIPO"
      .TextMatrix(0, 2) = "DESCRIÇÃO"
      .TextMatrix(0, 3) = "PREÇO"
      .TextMatrix(0, 4) = "QUANT"
      .TextMatrix(0, 5) = "SUBTOTAL"
      .TextMatrix(0, 6) = "DESC"
      .TextMatrix(0, 7) = "TOTAL"
      
      'colocar os cabeçalho em negrito
      For i = 0 To .Cols - 1
         .Col = i
         .Row = 0
         .CellFontBold = True
      Next
      
      'ALINHAMENTO
      '.ColAlignment(2) = 1
      
      'centralizar o titulo
      For i = 0 To .Cols - 1
         .Row = 0
         .Col = i
         .CellAlignment = flexAlignCenterCenter
      Next i
      
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = rTabela("tipo_item")
            
            If tipoEmpresa = 4 Then
            .TextMatrix(.rows - 1, 2) = rTabela("var_desc") & " /  " & rTabela("var_tam") & " / " & rTabela("var_fab")
            Else
            .TextMatrix(.rows - 1, 2) = rTabela("var_desc") & " /  " & ValidateNull(rTabela("var_fab"))
            End If
            
            .TextMatrix(.rows - 1, 3) = Format(rTabela("preco"), ocMONEY)
            .TextMatrix(.rows - 1, 4) = rTabela("quantidade")
            .TextMatrix(.rows - 1, 5) = Format(rTabela("var_Subtotal"), ocMONEY)
            .TextMatrix(.rows - 1, 6) = Format(rTabela("desconto"), ocMONEY)
            .TextMatrix(.rows - 1, 7) = Format(rTabela("total"), ocMONEY)

            rTabela.MoveNext
            .rows = .rows + 1
         Loop
      End If
      
      .rows = .rows - 1
   End With
End Sub

Private Sub txtCodPedido_Change()
Call MostrarParcelas
Call MostrarRecebedor
Call MostrarFuncionario
End Sub
Private Sub MostrarRecebedor()
If txtCodPedido.Text = "" Then Exit Sub
sSQL = "SELECT * FROM Pedidos_Recebedor WHERE (cod_pedido = " & txtCodPedido.Text & ");"
Set r = dbData.OpenRecordset(sSQL)

If Not r.BOF Then
    lblRecebedor.Visible = True
    lblRecebedorTit.Visible = True
    lblRecebedor.Caption = UCase(r("recebedor"))
Else
    lblRecebedor.Visible = False
    lblRecebedorTit.Visible = False
    lblRecebedor.Caption = ""
End If

If r.State <> 0 Then r.Close
Set r = Nothing

End Sub

