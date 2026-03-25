VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Begin VB.Form NFCe_Consultar_Produtos 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   Caption         =   "ITENS DO PEDIDO"
   ClientHeight    =   6390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13185
   Icon            =   "NFCe_Consultar_Produtos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   13185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ChamaleonBtn.chameleonButton cmdFechar 
      Height          =   255
      Left            =   12900
      TabIndex        =   13
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "X"
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
      FCOL            =   128
      FCOLO           =   128
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "NFCe_Consultar_Produtos.frx":23D2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   5520
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   120
      ScaleHeight     =   885
      ScaleWidth      =   12945
      TabIndex        =   7
      Top             =   180
      Width           =   12975
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
         Left            =   8655
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "000000"
         Top             =   180
         Width           =   1455
      End
      Begin VB.TextBox txtCodNFCe 
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
         Left            =   11280
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "000000"
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido::"
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
         Index           =   1
         Left            =   7800
         TabIndex        =   12
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NFCe:"
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
         Left            =   10545
         TabIndex        =   10
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUTOS DA NFCE"
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
         Left            =   1635
         TabIndex        =   8
         Top             =   300
         Width           =   3255
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   240
         Picture         =   "NFCe_Consultar_Produtos.frx":23EE
         Top             =   0
         Width           =   1140
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   9900
      ScaleHeight     =   1185
      ScaleWidth      =   3165
      TabIndex        =   0
      Top             =   5100
      Width           =   3195
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   60
         Width           =   1755
      End
      Begin VB.Label lblTotalDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   420
         Width           =   1755
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUB-TOTAL:"
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
         Left            =   120
         TabIndex        =   4
         Top             =   60
         Width           =   1110
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCONTO:"
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
         Left            =   120
         TabIndex        =   3
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label lblTotalGeral 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   780
         Width           =   1755
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL:"
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
         Left            =   120
         TabIndex        =   1
         Top             =   780
         Width           =   675
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdCorrigirProduto 
      Height          =   315
      Left            =   120
      TabIndex        =   14
      Top             =   5400
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Atualizar Produto"
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
      MICON           =   "NFCe_Consultar_Produtos.frx":8C34
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3915
      Left            =   120
      TabIndex        =   16
      Top             =   1140
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   6906
      _Version        =   393216
      AllowBigSelection=   0   'False
      FocusRect       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ChamaleonBtn.chameleonButton cmdConsultarNCM 
      Height          =   315
      Left            =   1800
      TabIndex        =   18
      Top             =   5400
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Consultar NCM pela Descrição"
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
      MICON           =   "NFCe_Consultar_Produtos.frx":8C50
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdConsultaNCMean 
      Height          =   315
      Left            =   4380
      TabIndex        =   19
      Top             =   5400
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Consultar NCM pelo EAN"
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
      MICON           =   "NFCe_Consultar_Produtos.frx":8C6C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdRecalcular 
      Height          =   315
      Left            =   6960
      TabIndex        =   20
      Top             =   5400
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Recalcular Tributos"
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
      MICON           =   "NFCe_Consultar_Produtos.frx":8C88
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblEstornar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5640
      TabIndex        =   21
      Top             =   5940
      Width           =   2595
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "[NOVIDADE] Você pode alterar os produtos diretamente na grade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   120
      TabIndex        =   17
      Top             =   5100
      Width           =   8595
   End
End
Attribute VB_Name = "NFCe_Consultar_Produtos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim sSQL As String
Dim r As ADODB.Recordset
Dim r2 As ADODB.Recordset
Dim vPed As Long
Dim cCfg As ConfigItem
Dim tipoEmpresa As Integer
Private iRow As Long, iCol As Long, xCancelada As Boolean

'abrir site para consultar ncm
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long
Private Const conSwNormal = 1
Public Sub loadPedidos(ByVal Pedido As Long)
vPed = Pedido

'consultar da venda
sSQL = "SELECT IdNFProd, DescontoPromocional " & _
      "FROM TbNFCe " & _
      "WHERE (IdNFProd = " & Pedido & ")"
Set r = dbData.OpenRecordset(sSQL)

Dim vDescVenda As Currency
vDescVenda = r("DescontoPromocional")

'somar os descontos dos itens da venda
sSQL = "SELECT sum(Desconto) as varSomaDescItens " & _
      "FROM TbNFCe_Itens " & _
      "WHERE (IdNFProd = " & Pedido & ")"
Set r = dbData.OpenRecordset(sSQL)

Dim vDescItensVenda As Currency
vDescItensVenda = r("varSomaDescItens")


'calcular descontos dos produtos
If vDescVenda <> vDescItensVenda Then
    'adiciona em cada item do pedido o valor do desconto
    sSQL = "UPDATE pedidos_itens SET desconto = (subtotal * " & Replace(CDbl(vDescItensVenda), ",", ".") & " / 100), total = subtotal - (subtotal * " & Replace(CDbl(vDescItensVenda), ",", ".") & " / 100), data = '" & Format$(txtDataCompra, "yyyy-dd-MM") & "' where (cod_pedido = " & txtCodPedido.Text & ")"
    dbData.Execute sSQL
    
    'soma todos os descontos dos itens da venda em real
    sSQL = "SELECT SUM(Desconto) AS varSomaDescItens FROM pedidos_itens WHERE (cod_pedido = " & txtCodPedido.Text & ")"
    Set r = dbData.OpenRecordset(sSQL)
    
    'Dim vSomaDescItens As Currency
    If Not r.EOF Then
        vSomaDescItens = FormatNumber(ValidateNull(r("varSomaDescItens")), 2)
    End If
    
    'consulto quanto é para ser o valor do desconto em real
    sSQL = "SELECT ValorDescReal FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ")"
    Set r = dbData.OpenRecordset(sSQL)
    
    'Dim vValorDescVenda As Currency
    If Not r.EOF Then
        vValorDescVenda = FormatNumber(ValidateNull(r("ValorDescReal")), 2)
    End If
    
    'se o valor total do desconto for maior que a soma dos desconto dos itens da venda
    If vValorDescVenda < vSomaDescItens Then
        vValorSobraDesc = CCur(vSomaDescItens - vValorDescVenda)
        sSQL = "UPDATE pedidos_itens SET Desconto = Desconto - " & Replace(CCur(vValorSobraDesc), ",", ".") & ", Total = Total + " & Replace(CCur(vValorSobraDesc), ",", ".") & " " & _
                "WHERE (CODIGO = " & _
        "(SELECT MAX(CODIGO) FROM pedidos_itens WHERE (cod_pedido = " & txtCodPedido.Text & ")))"
        dbData.Execute sSQL
    ElseIf vValorDescVenda > vSomaDescItens Then
        vValorSobraDesc = CCur(vValorDescVenda - vSomaDescItens)
        sSQL = "UPDATE pedidos_itens SET Desconto = Desconto + " & Replace(CCur(vValorSobraDesc), ",", ".") & ", Total = Total - " & Replace(CCur(vValorSobraDesc), ",", ".") & " " & _
                "WHERE (CODIGO = " & _
        "(SELECT MAX(CODIGO) FROM pedidos_itens WHERE (cod_pedido = " & txtCodPedido.Text & ")))"
        dbData.Execute sSQL
    End If
Else
    If lblEstornar.Caption = "ESTORNO" Then 'desativei em 29/05/25 pq fui consultar uma NFCe e deu erro pq não existe esse objeto no
        sSQL = "UPDATE pedidos_itens SET Desconto = '0.00', Total = Subtotal " & _
                "WHERE (cod_pedido = " & txtCodPedido.Text & ")"
        dbData.Execute sSQL
    End If
End If

'consultar itens da venda
sSQL = "SELECT IdNFProd, IdNFProd_Item, IDProduto, DescricaoProduto, CodBarras, UN, CodNcm, CFOP, ICMSCST, Aliq_Icms, COFINSCST, PISCST, ValorUnit, QtdeMov, (ValorUnit * QtdeMov) AS vSubTotal, Desconto, Bc_Icms, Vlr_Icms, vlr_PIS, vlr_COFINS, Aliq_COFINS , Aliq_PIS " & _
      "FROM TbNFCe_Itens " & _
      "WHERE (IdNFProd = " & Pedido & ")"
Set r = dbData.OpenRecordset(sSQL)

FormatarGrid_Itens r

If r.State <> 0 Then r.Close
Set r = Nothing

sSQL = "SELECT DescontoPromocional, Valor_NF_Prod, Num_OS_VD_Origem  FROM TbNFCe WHERE (IdNFProd = " & Pedido & ");"
Set r = dbData.OpenRecordset(sSQL)

'Debug.Print sSQL

Dim varDesc As Currency
Dim varSubTotal As Currency
Dim varTotalGeral As Currency
varDesc = r("DescontoPromocional")
varSubTotal = r("Valor_NF_Prod")
varTotalGeral = varSubTotal - varDesc

If Not r.BOF Then
    lblTotal.Caption = Format(r("Valor_NF_Prod"), ocMONEY)
    lblTotalGeral.Caption = Format(varTotalGeral, ocMONEY)
    lblTotalDesc.Caption = Format(r("DescontoPromocional"), ocMONEY)
    txtCodPedido.Text = Format(r("Num_OS_VD_Origem"), "000000")
End If

txtCodNFCe.Text = Format(Pedido, "000000")

If r.State <> 0 Then r.Close
Set r = Nothing
End Sub



Private Sub FormatarGrid_Itens(rTabela As ADODB.Recordset)
Dim i As Integer

With Grid
   .Clear
   .Cols = 23
   .Rows = 2
   
   .ColWidth(0) = 0
   .ColWidth(1) = 550
   .ColWidth(2) = 3500
   .ColWidth(3) = 1350
   .ColWidth(4) = 450
   .ColWidth(5) = 850
   .ColWidth(6) = 600
   .ColWidth(7) = 700
   .ColWidth(8) = 600
   .ColWidth(9) = 700
   .ColWidth(10) = 800
   .ColWidth(11) = 600
   .ColWidth(12) = 700
   .ColWidth(13) = 700
    .ColWidth(14) = 600
    .ColWidth(15) = 700
    .ColWidth(16) = 700
   .ColWidth(17) = 700
   .ColWidth(18) = 700
   .ColWidth(19) = 700
   .ColWidth(20) = 700
   .ColWidth(16) = 700
   .ColWidth(21) = 0
   .ColWidth(22) = 0
   
   .TextMatrix(0, 1) = "CÓD."
   .TextMatrix(0, 2) = "PRODUTO"
   .TextMatrix(0, 3) = "EAN"
   .TextMatrix(0, 4) = "UN"
   .TextMatrix(0, 5) = "NCM"
   .TextMatrix(0, 6) = "CFOP"
   .TextMatrix(0, 7) = "ICMS"
   .TextMatrix(0, 8) = "ALIQ."
   .TextMatrix(0, 9) = "VLR"          '09
   .TextMatrix(0, 10) = "COFINS" '10
   .TextMatrix(0, 11) = "ALIQ"
   .TextMatrix(0, 12) = "VLR"
   .TextMatrix(0, 13) = "PIS"   '13
   .TextMatrix(0, 14) = "ALIQ"
   .TextMatrix(0, 15) = "VLR"
   .TextMatrix(0, 16) = "PREÇO"     '16
   .TextMatrix(0, 17) = "QTDA."     '17
   .TextMatrix(0, 18) = "SUBTOTAL"  '18
   .TextMatrix(0, 19) = "DESC."     '19
   .TextMatrix(0, 20) = "BC ICMS"   '20
   .TextMatrix(0, 21) = "CUPOM"     '21
   .TextMatrix(0, 22) = "ITEM"      '22
   
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
'IdNFProd, IdNFProd_Item, IDProduto, DescricaoProduto, UN, CodNcm, CFOP, ICMSCST, Aliq_Icms, COFINSCST, PISCST, ValorUnit, QtdeMov, (ValorUnit * QtdeMov) AS vSubTotal, Desconto, Bc_Icms, Vlr_Icms
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         .TextMatrix(.Rows - 1, 1) = rTabela("IDProduto")
         .TextMatrix(.Rows - 1, 2) = rTabela("DescricaoProduto")
         .TextMatrix(.Rows - 1, 3) = rTabela("CodBarras")
         .TextMatrix(.Rows - 1, 4) = rTabela("UN")
         .TextMatrix(.Rows - 1, 5) = rTabela("CodNcm")
         .TextMatrix(.Rows - 1, 6) = rTabela("CFOP")
         .TextMatrix(.Rows - 1, 7) = rTabela("ICMSCST")
         .TextMatrix(.Rows - 1, 8) = Format(rTabela("Aliq_Icms"), ocMONEY)
         .TextMatrix(.Rows - 1, 9) = Format(rTabela("Vlr_Icms"), ocMONEY)
         .TextMatrix(.Rows - 1, 10) = rTabela("COFINSCST")
         .TextMatrix(.Rows - 1, 11) = Format(rTabela("Aliq_COFINS"), ocMONEY)
         .TextMatrix(.Rows - 1, 12) = Format(rTabela("vlr_COFINS"), ocMONEY)
         .TextMatrix(.Rows - 1, 13) = rTabela("PISCST")
         .TextMatrix(.Rows - 1, 14) = Format(rTabela("Aliq_PIS"), ocMONEY)
         .TextMatrix(.Rows - 1, 15) = Format(rTabela("vlr_PIS"), ocMONEY)
         .TextMatrix(.Rows - 1, 16) = Format(rTabela("ValorUnit"), ocMONEY)
         .TextMatrix(.Rows - 1, 17) = rTabela("QtdeMov")
         .TextMatrix(.Rows - 1, 18) = Format(rTabela("vSubTotal"), ocMONEY)
         .TextMatrix(.Rows - 1, 19) = Format(rTabela("Desconto"), ocMONEY)
         .TextMatrix(.Rows - 1, 20) = Format(rTabela("Bc_Icms"), ocMONEY)
         .TextMatrix(.Rows - 1, 21) = rTabela("IdNFProd")
         .TextMatrix(.Rows - 1, 22) = rTabela("IdNFProd_Item")

         rTabela.MoveNext
         .Rows = .Rows + 1
      Loop
   End If
   
         'mudar a cor da coluna
      For i = 1 To .Rows - 1
         .Row = i
         .Col = 7:   .CellBackColor = &H8080FF
         .Col = 8:   .CellBackColor = &HC0C0FF
         '.Col = 9:   .CellBackColor = &HC0C0FF
         .Col = 10:   .CellBackColor = &HC0C0&
         .Col = 11:   .CellBackColor = &HC0FFFF
         '.Col = 12:   .CellBackColor = &HC0FFFF
         .Col = 13:   .CellBackColor = &H80FF80
         .Col = 14:   .CellBackColor = &HC0FFC0
         '.Col = 15:   .CellBackColor = &HC0FFC0
        
        
      Next
   

   
   .Rows = .Rows - 1
End With
End Sub

Private Sub Recalcular_Desconto()
If vDescItensVenda <> "0,00" Then
    'adiciona em cada item do pedido o valor do desconto
    sSQL = "UPDATE pedidos_itens SET desconto = (subtotal * " & Replace(CDbl(vDescItensVenda), ",", ".") & " / 100), total = subtotal - (subtotal * " & Replace(CDbl(vDescItensVenda), ",", ".") & " / 100), data = '" & Format$(txtDataCompra, "yyyy-dd-MM") & "' where (cod_pedido = " & txtCodPedido.Text & ")"
    dbData.Execute sSQL
    
    'soma todos os descontos dos itens da venda em real
    sSQL = "SELECT SUM(Desconto) AS varSomaDescItens FROM pedidos_itens WHERE (cod_pedido = " & txtCodPedido.Text & ")"
    Set r = dbData.OpenRecordset(sSQL)
    
    'Dim vSomaDescItens As Currency
    If Not r.EOF Then
        vSomaDescItens = FormatNumber(ValidateNull(r("varSomaDescItens")), 2)
    End If
    
    'consulto quanto é para ser o valor do desconto em real
    sSQL = "SELECT ValorDescReal FROM pedidos WHERE (cod_pedido = " & txtCodPedido.Text & ")"
    Set r = dbData.OpenRecordset(sSQL)
    
    'Dim vValorDescVenda As Currency
    If Not r.EOF Then
        vValorDescVenda = FormatNumber(ValidateNull(r("ValorDescReal")), 2)
    End If
    
    'se o valor total do desconto for maior que a soma dos desconto dos itens da venda
    If vValorDescVenda < vSomaDescItens Then
        vValorSobraDesc = CCur(vSomaDescItens - vValorDescVenda)
        sSQL = "UPDATE pedidos_itens SET Desconto = Desconto - " & Replace(CCur(vValorSobraDesc), ",", ".") & ", Total = Total + " & Replace(CCur(vValorSobraDesc), ",", ".") & " " & _
                "WHERE (CODIGO = " & _
        "(SELECT MAX(CODIGO) FROM pedidos_itens WHERE (cod_pedido = " & txtCodPedido.Text & ")))"
        dbData.Execute sSQL
    ElseIf vValorDescVenda > vSomaDescItens Then
        vValorSobraDesc = CCur(vValorDescVenda - vSomaDescItens)
        sSQL = "UPDATE pedidos_itens SET Desconto = Desconto + " & Replace(CCur(vValorSobraDesc), ",", ".") & ", Total = Total - " & Replace(CCur(vValorSobraDesc), ",", ".") & " " & _
                "WHERE (CODIGO = " & _
        "(SELECT MAX(CODIGO) FROM pedidos_itens WHERE (cod_pedido = " & txtCodPedido.Text & ")))"
        dbData.Execute sSQL
    End If
Else
    If lblEstornar.Caption = "ESTORNO" Then
        sSQL = "UPDATE pedidos_itens SET Desconto = '0.00', Total = Subtotal " & _
                "WHERE (cod_pedido = " & txtCodPedido.Text & ")"
        dbData.Execute sSQL
    End If
End If
End Sub

Private Sub cmdConsultaNCMean_Click()
Dim varNomeProduto As String
varNomeProduto = Grid.TextMatrix(Grid.Row, 3)
ShellExecute hwnd, "open", "https://cosmos.bluesoft.com.br/pesquisar?utf8=" + Chr(95) + "&q=" & varNomeProduto & "", vbNullString, vbNullString, conSwNo
End Sub

Private Sub cmdConsultarNCM_Click()
Dim varNomeProduto As String
varNomeProduto = Replace(Grid.TextMatrix(Grid.Row, 2), " ", "+")
ShellExecute hwnd, "open", "https://cosmos.bluesoft.com.br/pesquisar?utf8=" + Chr(95) + "&q=" & varNomeProduto & "", vbNullString, vbNullString, conSwNo
End Sub

Private Sub cmdCorrigirProduto_Click()
If Grid.Rows <= 1 Then
    MsgBox "Não existe nenhum pedido selecionado!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

Dim varCodProduto As String
varCodProduto = Grid.TextMatrix(Grid.Row, 1)

If ShowMsg("Deseja atualizar o produto " & Grid.TextMatrix(Grid.Row, 2) & " ?", vbInformation + vbYesNo) = vbYes Then

Load Produtos_Cadastro
Produtos_Cadastro.SSTab1.Tab = 0
Produtos_Cadastro.cmdNovo.Enabled = False
Produtos_Cadastro.cmdSalvar.Enabled = False
Produtos_Cadastro.cmdCancelar.Enabled = False
Produtos_Cadastro.cmdAlterar.Enabled = True
Produtos_Cadastro.cmdExcluir.Enabled = True
vTipoEdicao = "Edicao"
Produtos_Cadastro.txtCodigo.Text = varCodProduto
Produtos_Cadastro.Show 1

End If

'If Grid.TextMatrix(Grid.Row, 13) = "SIM" Then
'    MsgBox "Não é possivel abrir um pedido cancelado!", vbInformation, "Aviso do Sistema"
'    Exit Sub
'End If

'If Grid.TextMatrix(Grid.Row, 9) = "SIM" Then
'    MsgBox "Não é possivel abrir um pedido que já emitiu NFCE!", vbInformation, "Aviso do Sistema"
'    Exit Sub
'End If



'If ShowMsg("Tem certeza que deseja reabrir o pedido " & Grid.TextMatrix(Grid.Row, 1) & " ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'    PDV.frmAvancado.Visible = False
'    PDV.frmSenha.Visible = False
'    Unload Estonar
'    PDV.lblEstornar.Caption = "ESTORNO"
'    PDV.txtCodPedido.Text = varCodProduto
    
'End If
End Sub

Private Sub cmdRecalcular_Click()
loadPedidos (Val(txtCodNFCe.Text))
Recalcular_Desconto
End Sub

Private Sub Form_Activate()
'If vPed <> "" Then
    loadPedidos (vPed)
'End If
End Sub

Private Sub Form_Load()
Set cCfg = sysConfig("TIPO_EMPRESA")
tipoEmpresa = cCfg.Value
Set cCfg = Nothing
End Sub
Private Sub Grid_Click()
Dim i As Integer

For i = 3 To 14
   If Grid.ColSel = i Then
      txtEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth, Grid.CellHeight
      txtEdit.Text = Grid.TextMatrix(Grid.Row, Grid.Col)
      txtEdit.Visible = True
      txtEdit.SetFocus
      txtEdit.SelStart = 0
      txtEdit.SelLength = Len(txtEdit.Text)
      iRow = Grid.Row
      iCol = Grid.Col
   End If
Next
End Sub

Private Sub txtEdit_KeyUp(KeyCode As Integer, Shift As Integer)
'Exit Sub
If KeyCode = 38 Then
   If Grid.Row - 1 = 0 Then ShowMsg "VOCÊ JÁ ESTÁ NA PRIMEIRA LINHA !!!", vbExclamation: Exit Sub
   Grid.Row = iRow - 1
   Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
   Grid_Click

ElseIf KeyCode = 40 Then
   If Grid.Rows = Grid.Row + 1 Then ShowMsg "VOCÊ JÁ ESTÁ NA ULTIMA LINHA !!!", vbExclamation: Exit Sub
   Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
   Grid.Row = iRow + 1
   Grid_Click
End If
End Sub
Private Sub txtEdit_LostFocus()
Dim AtualizarProdutos As Boolean

AtualizarProdutos = False


If iCol = 3 Then
    'If txtEdit.Text <> "" Then
    '    If Len(txtEdit.Text) < 3 Or Len(txtEdit.Text) > 3 Then
    '        MsgBox "CST Inválido!", vbInformation, "Aviso do Sistema"
    '        Grid.TextMatrix(iRow, iCol) = 0
    '        AtualizarProdutos = False
    '        'Exit Sub
    '    Else
    '        Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
    '        AtualizarProdutos = True
    '    End If
    'Else
        'If txtEdit.Text <> "" Then
        Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", "", txtEdit.Text)
        AtualizarProdutos = True
    'End If
ElseIf iCol = 5 Then
    txtEdit.Text = Replace(txtEdit.Text, ".", "")
    'Grid.TextMatrix(iRow, 5) = Replace(Grid.TextMatrix(iRow, iCol), ".", "")
    If txtEdit.Text <> "" Then
        If Len(txtEdit.Text) < 8 Or Len(txtEdit.Text) > 8 Then
            MsgBox "NCM Inválido!", vbInformation, "Aviso do Sistema"
            Grid.TextMatrix(iRow, iCol) = 0
            AtualizarProdutos = False
        Else
            Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
            AtualizarProdutos = True
        End If
    Else
        Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
        AtualizarProdutos = True
    End If


ElseIf iCol = 6 Then
    If txtEdit.Text <> "" Then
           If Len(txtEdit.Text) < 4 Or Len(txtEdit.Text) > 4 Then
               MsgBox "CFOP Inválido!", vbInformation, "Aviso do Sistema"
               Grid.TextMatrix(iRow, iCol) = 0
               AtualizarProdutos = False
               'Exit Sub
           Else
               Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
               AtualizarProdutos = True
           End If
       Else
           Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
           AtualizarProdutos = True
       End If

ElseIf iCol = 7 Then
    If txtEdit.Text <> "" Then
        If Len(txtEdit.Text) < 3 Or Len(txtEdit.Text) > 3 Then
            MsgBox "ICMS CST Inválido!", vbInformation, "Aviso do Sistema"
            Grid.TextMatrix(iRow, iCol) = 0
            AtualizarProdutos = False
            'Exit Sub
        Else
            Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", "000", txtEdit.Text)
            AtualizarProdutos = True
        End If
    Else
        Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", "000", txtEdit.Text)
        AtualizarProdutos = True
    End If

ElseIf iCol = 8 Then
    If txtEdit.Text <> "" Then
        'If Len(txtEdit.Text) < 3 Or Len(txtEdit.Text) > 3 Then
        '    MsgBox "CST Inválido!", vbInformation, "Aviso do Sistema"
        '    Grid.TextMatrix(iRow, iCol) = 0
        '    AtualizarProdutos = False
        '    'Exit Sub
        'Else
            Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", "0,00", Format(txtEdit.Text, ocMONEY))
            AtualizarProdutos = True
        'End If
    Else
        Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", "0,00", Format(txtEdit.Text, ocMONEY))
        AtualizarProdutos = True
    End If

ElseIf iCol = 10 Then
    If txtEdit.Text <> "" Then
        If Len(txtEdit.Text) < 2 Or Len(txtEdit.Text) > 2 Then
            MsgBox "CONFINS CST Inválido!", vbInformation, "Aviso do Sistema"
            Grid.TextMatrix(iRow, iCol) = 0
            AtualizarProdutos = False
            'Exit Sub
        Else
            Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", "00", txtEdit.Text)
            AtualizarProdutos = True
        End If
    Else
        Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", "00", txtEdit.Text)
        AtualizarProdutos = True
    End If




ElseIf iCol = 11 Then
    If txtEdit.Text <> "" Then
        'If Len(txtEdit.Text) < 3 Or Len(txtEdit.Text) > 3 Then
        '    MsgBox "CST Inválido!", vbInformation, "Aviso do Sistema"
        '    Grid.TextMatrix(iRow, iCol) = 0
        '    AtualizarProdutos = False
        '    'Exit Sub
        'Else
            Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", "0,00", Format(txtEdit.Text, ocMONEY))
            AtualizarProdutos = True
        'End If
    Else
        Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", "0,00", Format(txtEdit.Text, ocMONEY))
        AtualizarProdutos = True
    End If
    
    
    
ElseIf iCol = 13 Then
    If txtEdit.Text <> "" Then
        If Len(txtEdit.Text) < 2 Or Len(txtEdit.Text) > 2 Then
            MsgBox "PIS Inválido!", vbInformation, "Aviso do Sistema"
            Grid.TextMatrix(iRow, iCol) = 0
            AtualizarProdutos = False
            'Exit Sub
        Else
            Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", "00", txtEdit.Text)
            AtualizarProdutos = True
        End If
    Else
        Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", "00", txtEdit.Text)
        AtualizarProdutos = True
    End If
'Else
'    Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
'    AtualizarProdutos = True
'End If






ElseIf iCol = 14 Then
    If txtEdit.Text <> "" Then
        'If Len(txtEdit.Text) < 3 Or Len(txtEdit.Text) > 3 Then
        '    MsgBox "CST Inválido!", vbInformation, "Aviso do Sistema"
        '    Grid.TextMatrix(iRow, iCol) = 0
        '    AtualizarProdutos = False
        '    'Exit Sub
        'Else
            Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", "0,00", Format(txtEdit.Text, ocMONEY))
            AtualizarProdutos = True
        'End If
    Else
        Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", "0,00", Format(txtEdit.Text, ocMONEY))
        AtualizarProdutos = True
    End If
    
    
Else
    Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)
    AtualizarProdutos = True
End If

txtEdit.Visible = False

If AtualizarProdutos = True Then
    AtualizarGrid_Itens
End If

End Sub

Private Sub AtualizarGrid_Itens()
Dim i As Integer
Dim sSQL As String
'Dim varCodBarra As String
   
For i = 1 To Grid.Rows - 1
'    If Grid.TextMatrix(i, 2) = "" Then
'        varCodBarra = 0
'    Else
'        varCodBarra = Grid.TextMatrix(i, 2)
'    End If

 
   If Grid.TextMatrix(i, 1) <> "" Then
      'dbData.Execute "UPDATE NotaFiscalItens SET CFOP = " & Grid.TextMatrix(i, 7) & ", CST = '" & Grid.TextMatrix(i, 6) & "', NCM = '" & Grid.TextMatrix(i, 5) & "'  WHERE CodigoNota = " & txtCodNota.Text & " AND ITEM = " & Grid.TextMatrix(i, 1) & ""
      dbData.Execute "UPDATE TbNFCe_Itens SET CodBarras = '" & Grid.TextMatrix(i, 3) & "', UN = '" & Grid.TextMatrix(i, 4) & "', CFOP = " & Grid.TextMatrix(i, 6) & ", ICMSCST = '" & Grid.TextMatrix(i, 7) & "', CodNcm = '" & Grid.TextMatrix(i, 5) & "', Aliq_Icms = " & fSQL(Grid.TextMatrix(i, 8), 2) & ", cofinsCST = '" & Grid.TextMatrix(i, 10) & "', pisCST = '" & Grid.TextMatrix(i, 13) & "', Aliq_COFINS  = " & fSQL(Grid.TextMatrix(i, 11), 2) & ", Aliq_PIS  = " & fSQL(Grid.TextMatrix(i, 14), 2) & ", Vlr_Icms = ((" & fSQL(Grid.TextMatrix(i, 20), 2) & " /100) * " & fSQL(Grid.TextMatrix(i, 8), 2) & "), vlr_COFINS = ((" & fSQL(Grid.TextMatrix(i, 20), 2) & " /100) * " & fSQL(Grid.TextMatrix(i, 11), 2) & "), vlr_PIS   = ((" & fSQL(Grid.TextMatrix(i, 20), 2) & " /100) * " & fSQL(Grid.TextMatrix(i, 14), 2) & ")  WHERE IdNFProd = " & Grid.TextMatrix(i, 21) & " AND  IdNFProd_Item = " & Grid.TextMatrix(i, 22) & ""
      dbData.Execute "UPDATE TbNFCe SET BaseCalc_ICMS = (SELECT ISNULL(SUM(Bc_Icms), 0) AS vTotalBCI FROM TbNFCe_Itens WHERE (IdNFProd = " & Grid.TextMatrix(i, 21) & ") AND (Aliq_Icms <> '0.00')), Valor_ICMS = (SELECT ISNULL(SUM(Vlr_Icms), 0) AS vValorICMS FROM TbNFCe_Itens AS TbNFCe_Itens_1 WHERE (IdNFProd = " & Grid.TextMatrix(i, 21) & ") AND (Aliq_Icms <> '0.00')) WHERE (IdNFProd = " & Grid.TextMatrix(i, 21) & ")"
      dbData.Execute "UPDATE Produtos SET EAN = '" & Grid.TextMatrix(i, 3) & "', UNID_MEDIDA = '" & Grid.TextMatrix(i, 4) & "', NCM = '" & Grid.TextMatrix(i, 5) & "', CFOP = " & Grid.TextMatrix(i, 6) & ", icmsCST = '" & Grid.TextMatrix(i, 7) & "', cofinsCST = '" & Grid.TextMatrix(i, 10) & "', pisCST = '" & Grid.TextMatrix(i, 13) & "'  WHERE CODIGO = " & Grid.TextMatrix(i, 1) & ""
   End If
Next
End Sub

Private Sub cmdFechar_Click()
Unload Me
End Sub

