VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Begin VB.Form NFCe_Consultar_Produtos 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ITENS DO PEDIDO"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13095
   Icon            =   "NFCe_Consultar_Produtos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   13095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkICMS 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ICMS"
      Height          =   195
      Left            =   60
      TabIndex        =   28
      Top             =   1080
      Width           =   750
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   5520
      TabIndex        =   14
      Top             =   2520
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      ScaleHeight     =   885
      ScaleWidth      =   12945
      TabIndex        =   7
      Top             =   60
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
         Picture         =   "NFCe_Consultar_Produtos.frx":23D2
         Top             =   0
         Width           =   1140
      End
      Begin VB.Label lblSomenteLeitura 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SOMENTE LEITURA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   5100
         TabIndex        =   29
         Top             =   330
         Visible         =   0   'False
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   9840
      ScaleHeight     =   1905
      ScaleWidth      =   3165
      TabIndex        =   0
      Top             =   5280
      Width           =   3195
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FRETE:"
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
         Left            =   120
         TabIndex        =   33
         Top             =   780
         Width           =   495
      End
      Begin VB.Label lblTotalFrete 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   32
         Top             =   780
         Width           =   1755
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ACRÉSCIMO:"
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
         Left            =   120
         TabIndex        =   31
         Top             =   420
         Width           =   915
      End
      Begin VB.Label lblTotalAcresc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   30
         Top             =   420
         Width           =   1755
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
            Name            =   "Arial Narrow"
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
         Top             =   1140
         Width           =   1755
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUB-TOTAL:"
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
         Left            =   120
         TabIndex        =   4
         Top             =   60
         Width           =   840
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCONTO:"
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
         Left            =   120
         TabIndex        =   3
         Top             =   1140
         Width           =   825
      End
      Begin VB.Label lblTotalGeral 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         Top             =   1500
         Width           =   1755
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL:"
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
         Left            =   120
         TabIndex        =   1
         Top             =   1500
         Width           =   495
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdCorrigirProduto 
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   5800
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
      MICON           =   "NFCe_Consultar_Produtos.frx":8C18
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox chkPis 
      BackColor       =   &H00C0FFC0&
      Caption         =   "PIS"
      Height          =   195
      Left            =   900
      TabIndex        =   21
      Top             =   1080
      Width           =   630
   End
   Begin VB.CheckBox chkCofins 
      BackColor       =   &H00C0FFC0&
      Caption         =   "COFINS"
      Height          =   195
      Left            =   1620
      TabIndex        =   22
      Top             =   1080
      Width           =   915
   End
   Begin VB.CheckBox chkFrete 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Frete"
      Height          =   195
      Left            =   2640
      TabIndex        =   23
      Top             =   1080
      Width           =   735
   End
   Begin VB.CheckBox chkSeguro 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Seguro"
      Height          =   195
      Left            =   3480
      TabIndex        =   24
      Top             =   1080
      Width           =   855
   End
   Begin VB.CheckBox chkOutros 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Outros"
      Height          =   195
      Left            =   4440
      TabIndex        =   25
      Top             =   1080
      Width           =   855
   End
   Begin VB.CheckBox chkReforma 
      BackColor       =   &H00C0FFC0&
      Caption         =   "CBS/IBS"
      Height          =   195
      Left            =   5400
      TabIndex        =   26
      Top             =   1080
      Width           =   975
   End
   Begin VB.CheckBox chkReformaIS 
      BackColor       =   &H00C0FFC0&
      Caption         =   "IS"
      Height          =   195
      Left            =   6480
      TabIndex        =   27
      Top             =   1080
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3915
      Left            =   60
      TabIndex        =   15
      Top             =   1320
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   6906
      _Version        =   393216
      AllowBigSelection=   0   'False
      FocusRect       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      TabIndex        =   17
      Top             =   5800
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
   Begin ChamaleonBtn.chameleonButton cmdConsultaNCMean 
      Height          =   315
      Left            =   4380
      TabIndex        =   18
      Top             =   5800
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
   Begin ChamaleonBtn.chameleonButton cmdRecalcular 
      Height          =   315
      Left            =   6960
      TabIndex        =   19
      Top             =   5800
      Visible         =   0   'False
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
   Begin VB.Label lblEstornar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5640
      TabIndex        =   20
      Top             =   6340
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
      TabIndex        =   16
      Top             =   5500
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
Private bSomenteLeitura As Boolean  'True quando a NFCe ja foi enviada, cancelada ou inutilizada
Dim vDescItensVenda As Currency  'compartilhada entre loadPedidos e Recalcular_Desconto

'abrir site para consultar ncm
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long
Private Const conSwNormal = 1
Public Sub loadPedidos(ByVal Pedido As Long)
vPed = Pedido

'NFCe ja consolidada na SEFAZ (enviada/cancelada/inutilizada) nao pode mais ter os dados
'fiscais alterados no grid - trava a edicao e avisa visualmente
Dim rStatus As ADODB.Recordset
sSQL = "SELECT NFCeEnviada, NFCeCancelada, Inutilizada FROM TbNFCe WHERE (IdNFProd = " & Pedido & ")"
Set rStatus = dbData.OpenRecordset(sSQL)
bSomenteLeitura = False
If Not rStatus.EOF Then
    bSomenteLeitura = ValidateNull(rStatus("NFCeEnviada")) Or ValidateNull(rStatus("NFCeCancelada")) Or ValidateNull(rStatus("Inutilizada"))
End If
If rStatus.State <> 0 Then rStatus.Close
Set rStatus = Nothing
lblSomenteLeitura.Visible = bSomenteLeitura


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
sSQL = "SELECT IdNFProd, IdNFProd_Item, IDProduto, DescricaoProduto, CodBarras, UN, CodNcm, CFOP, " & _
      "IBSCBS_CST, cClassTrib, IBS_vIBS, CBS_vCBS, IS_CST, cClassTrib_IS, IS_vIS, " & _
      "ValorUnit, QtdeMov, (ValorUnit * QtdeMov) AS vSubTotal, Valor_Frete, Valor_Seguro, ValorOutras, Desconto, " & _
      "ICMSCST, Aliq_Icms, Bc_Icms, Vlr_Icms, PISCST, Aliq_PIS, vlr_PIS, COFINSCST, Aliq_COFINS, vlr_COFINS " & _
      "FROM TbNFCe_Itens " & _
      "WHERE (IdNFProd = " & Pedido & ")"
Set r = dbData.OpenRecordset(sSQL)

FormatarGrid_Itens r
AplicarVisibilidadeGrid

If r.State <> 0 Then r.Close
Set r = Nothing

sSQL = "SELECT DescontoPromocional, OutrasDespesasAces, Valor_Frete, Valor_NF_Prod, Num_OS_VD_Origem  FROM TbNFCe WHERE (IdNFProd = " & Pedido & ");"
Set r = dbData.OpenRecordset(sSQL)

'Debug.Print sSQL

Dim varDesc As Currency
Dim varAcresc As Currency
Dim varFrete As Currency
Dim varSubTotal As Currency
Dim varTotalGeral As Currency
varDesc = r("DescontoPromocional")
varAcresc = ValidateNull(r("OutrasDespesasAces"))
varFrete = ValidateNull(r("Valor_Frete"))
varSubTotal = r("Valor_NF_Prod")
varTotalGeral = varSubTotal - varDesc + varAcresc + varFrete

If Not r.BOF Then
    lblTotal.Caption = Format(r("Valor_NF_Prod"), ocMONEY)
    lblTotalGeral.Caption = Format(varTotalGeral, ocMONEY)
    lblTotalDesc.Caption = Format(r("DescontoPromocional"), ocMONEY)
    lblTotalAcresc.Caption = Format(varAcresc, ocMONEY)
    lblTotalFrete.Caption = Format(varFrete, ocMONEY)
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
   .Cols = 33
   .Rows = 2

   .ColWidth(0) = 0
   .ColWidth(1) = 1150  'EAN
   .ColWidth(2) = 450   'COD.
   .ColWidth(3) = 3000  'PRODUTO
   .ColWidth(4) = 380   'UN
   .ColWidth(5) = 750   'NCM
   .ColWidth(6) = 550   'CFOP
   .ColWidth(7) = 850   'CST ICMS
   .ColWidth(8) = 0     'ALIQ. ICMS (chkICMS)
   .ColWidth(9) = 0     'BC ICMS (chkICMS)
   .ColWidth(10) = 0    'VLR ICMS (chkICMS)
   .ColWidth(11) = 0    'CST IBS (chkReforma)
   .ColWidth(12) = 0    'CLASS. IBS (chkReforma)
   .ColWidth(13) = 0    'V. IBS (chkReforma)
   .ColWidth(14) = 0    'V. CBS (chkReforma)
   .ColWidth(15) = 0    'CST IS (chkReformaIS)
   .ColWidth(16) = 0    'CLASS IS (chkReformaIS)
   .ColWidth(17) = 0    'V. IS (chkReformaIS)
   .ColWidth(18) = 700  'PRECO
   .ColWidth(19) = 650  'QTDE
   .ColWidth(20) = 0    'FRETE (chkFrete)
   .ColWidth(21) = 0    'SEGURO (chkSeguro)
   .ColWidth(22) = 0    'OUTROS (chkOutros)
   .ColWidth(23) = 650  'DESC.
   .ColWidth(24) = 850  'SUBTOTAL
   .ColWidth(25) = 0    'CST PIS (chkPis)
   .ColWidth(26) = 0    'ALIQ. PIS (chkPis)
   .ColWidth(27) = 0    'VLR PIS (chkPis)
   .ColWidth(28) = 0    'CST COFINS (chkCofins)
   .ColWidth(29) = 0    'ALIQ. COFINS (chkCofins)
   .ColWidth(30) = 0    'VLR COFINS (chkCofins)
   .ColWidth(31) = 0    'CUPOM (oculta)
   .ColWidth(32) = 0    'ITEM (oculta)

   .TextMatrix(0, 1) = "EAN"
   .TextMatrix(0, 2) = "CÓD."
   .TextMatrix(0, 3) = "PRODUTO"
   .TextMatrix(0, 4) = "UN"
   .TextMatrix(0, 5) = "NCM"
   .TextMatrix(0, 6) = "CFOP"
   .TextMatrix(0, 7) = "ICMS CST"
   .TextMatrix(0, 8) = "ICMS ALIQ."
   .TextMatrix(0, 9) = "ICMS BC"
   .TextMatrix(0, 10) = "ICMS VLR"
   .TextMatrix(0, 11) = "IBS CST"
   .TextMatrix(0, 12) = "IBS CLASS."
   .TextMatrix(0, 13) = "IBS VLR"
   .TextMatrix(0, 14) = "CBS VLR"
   .TextMatrix(0, 15) = "IS CST"
   .TextMatrix(0, 16) = "IS CLASS."
   .TextMatrix(0, 17) = "IS VLR"
   .TextMatrix(0, 18) = "PREÇO"
   .TextMatrix(0, 19) = "QTDE"
   .TextMatrix(0, 20) = "FRETE"
   .TextMatrix(0, 21) = "SEGURO"
   .TextMatrix(0, 22) = "OUTROS"
   .TextMatrix(0, 23) = "DESC."
   .TextMatrix(0, 24) = "SUBTOTAL"
   .TextMatrix(0, 25) = "PIS CST"
   .TextMatrix(0, 26) = "PIS ALIQ."
   .TextMatrix(0, 27) = "PIS  VLR"
   .TextMatrix(0, 28) = "COFINS CST"
   .TextMatrix(0, 29) = "COFINS ALIQ."
   .TextMatrix(0, 30) = "COFINS  VLR"
   .TextMatrix(0, 31) = "CUPOM"
   .TextMatrix(0, 32) = "ITEM"

   'cabecalho em negrito e centralizado
   For i = 0 To .Cols - 1
      .Col = i
      .Row = 0
      .CellFontBold = True
      .CellAlignment = flexAlignCenterCenter
   Next i

   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         .TextMatrix(.Rows - 1, 1) = rTabela("CodBarras")
         .TextMatrix(.Rows - 1, 2) = rTabela("IDProduto")
         .TextMatrix(.Rows - 1, 3) = rTabela("DescricaoProduto")
         .TextMatrix(.Rows - 1, 4) = rTabela("UN")
         .TextMatrix(.Rows - 1, 5) = rTabela("CodNcm")
         .TextMatrix(.Rows - 1, 6) = rTabela("CFOP")
         .TextMatrix(.Rows - 1, 7) = rTabela("ICMSCST")
         .TextMatrix(.Rows - 1, 8) = Format(rTabela("Aliq_Icms"), ocMONEY)
         .TextMatrix(.Rows - 1, 9) = Format(rTabela("Bc_Icms"), ocMONEY)
         .TextMatrix(.Rows - 1, 10) = Format(rTabela("Vlr_Icms"), ocMONEY)
         .TextMatrix(.Rows - 1, 11) = ValidateNull(rTabela("IBSCBS_CST"))
         .TextMatrix(.Rows - 1, 12) = ValidateNull(rTabela("cClassTrib"))
         .TextMatrix(.Rows - 1, 13) = FormatNumber(rTabela("IBS_vIBS"), 2)
         .TextMatrix(.Rows - 1, 14) = FormatNumber(rTabela("CBS_vCBS"), 2)
         .TextMatrix(.Rows - 1, 15) = ValidateNull(rTabela("IS_CST"))
         .TextMatrix(.Rows - 1, 16) = ValidateNull(rTabela("cClassTrib_IS"))
         .TextMatrix(.Rows - 1, 17) = FormatNumber(rTabela("IS_vIS"), 2)
         .TextMatrix(.Rows - 1, 18) = Format(rTabela("ValorUnit"), ocMONEY)
         .TextMatrix(.Rows - 1, 19) = rTabela("QtdeMov")
         .TextMatrix(.Rows - 1, 20) = FormatNumber(rTabela("Valor_Frete"), 2)
         .TextMatrix(.Rows - 1, 21) = FormatNumber(rTabela("Valor_Seguro"), 2)
         .TextMatrix(.Rows - 1, 22) = FormatNumber(rTabela("ValorOutras"), 2)
         .TextMatrix(.Rows - 1, 23) = Format(rTabela("Desconto"), ocMONEY)
         .TextMatrix(.Rows - 1, 24) = Format(rTabela("vSubTotal"), ocMONEY)
         .TextMatrix(.Rows - 1, 25) = rTabela("PISCST")
         .TextMatrix(.Rows - 1, 26) = Format(rTabela("Aliq_PIS"), ocMONEY)
         .TextMatrix(.Rows - 1, 27) = Format(rTabela("vlr_PIS"), ocMONEY)
         .TextMatrix(.Rows - 1, 28) = rTabela("COFINSCST")
         .TextMatrix(.Rows - 1, 29) = Format(rTabela("Aliq_COFINS"), ocMONEY)
         .TextMatrix(.Rows - 1, 30) = Format(rTabela("vlr_COFINS"), ocMONEY)
         .TextMatrix(.Rows - 1, 31) = rTabela("IdNFProd")
         .TextMatrix(.Rows - 1, 32) = rTabela("IdNFProd_Item")

         rTabela.MoveNext
         .Rows = .Rows + 1
      Loop
   End If

   'cor de fundo: colunas editaveis em amarelo claro (UN/NCM ficam sem cor especial; grupo ICMS inteiro fica amarelo)
   Dim colEdit As Variant
   For Each colEdit In Array(1, 6, 7, 8, 9, 10, 20, 21, 22, 25, 26, 27, 28, 29, 30)
      For i = 1 To .Rows - 1
         .Row = i: .Col = colEdit
         .CellBackColor = &HC8FFFF
      Next i
   Next colEdit

   'cor de fundo: colunas de reforma tributaria em azul claro
   Dim colRef As Variant
   For Each colRef In Array(11, 12, 13, 14, 15, 16, 17)
      For i = 1 To .Rows - 1
         .Row = i: .Col = colRef
         .CellBackColor = &HFFFFF0
      Next i
   Next colRef

   .Rows = .Rows - 1
   .Col = 0
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
varNomeProduto = Grid.TextMatrix(Grid.Row, 1)
ShellExecute hwnd, "open", "https://cosmos.bluesoft.com.br/pesquisar?utf8=" + Chr(95) + "&q=" & varNomeProduto & "", vbNullString, vbNullString, conSwNo
End Sub

Private Sub cmdConsultarNCM_Click()
Dim varNomeProduto As String
varNomeProduto = Replace(Grid.TextMatrix(Grid.Row, 3), " ", "+")
ShellExecute hwnd, "open", "https://cosmos.bluesoft.com.br/pesquisar?utf8=" + Chr(95) + "&q=" & varNomeProduto & "", vbNullString, vbNullString, conSwNo
End Sub

Private Sub cmdCorrigirProduto_Click()
If Grid.Rows <= 1 Then
    MsgBox "Não existe nenhum pedido selecionado!", vbInformation, "Aviso do Sistema"
    Exit Sub
End If

Dim varCodProduto As String
varCodProduto = Grid.TextMatrix(Grid.Row, 2)

If ShowMsg("Deseja atualizar o produto " & Grid.TextMatrix(Grid.Row, 3) & " ?", vbInformation + vbYesNo) = vbYes Then

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
Select Case Grid.Col
    Case 1, 4, 5, 6, 7, 8, 11, 12, 15, 16, 20, 21, 22, 25, 26, 28, 29
        If Grid.Row > 0 And Grid.TextMatrix(Grid.Row, 31) <> "" Then
            If bSomenteLeitura Then
                MsgBox "Esta NFCe já foi enviada, cancelada ou inutilizada." & vbCrLf & "Os dados fiscais não podem mais ser alterados.", vbExclamation, "Somente Leitura"
                Exit Sub
            End If
            txtEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth, Grid.CellHeight
            txtEdit.Text = Grid.TextMatrix(Grid.Row, Grid.Col)
            txtEdit.Visible = True
            txtEdit.SetFocus
            txtEdit.SelStart = 0
            txtEdit.SelLength = Len(txtEdit.Text)
            iRow = Grid.Row
            iCol = Grid.Col
        End If
End Select
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
Dim sVal As String
Dim sItemId As String
Dim sCodProd As String
Dim sChkCST As String
Dim sNewClassTrib As String
Dim dblPRedIBS As Double, dblPRedCBS As Double
Dim dblIBSvBC As Double, dblIBSUFpAliq As Double, dblIBSMunpAliq As Double
Dim dblCBSvBC As Double, dblCBSpAliq As Double
Dim curIBSvIBSUF As Currency, curIBSvIBSMun As Currency, curIBSvIBS As Currency, curCBSvCBS As Currency
Dim rsClassTrib As ADODB.Recordset
Dim rsIBSItem As ADODB.Recordset
Dim sISCSTAtual As String
Dim rsISTrib As ADODB.Recordset
Dim sISCST2 As String
Dim iTipoIS2 As Integer
Dim dISpAliq2 As Double, dISvUnid2 As Double
Dim sISTrib2uTrib As String
Dim dISvBC2 As Double, dISqUnid2 As Double, dCalcISvBC2 As Double
Dim curISvIS2 As Currency
Dim curBCICMS As Currency, curVICMS As Currency

txtEdit.Visible = False
sVal = Trim(txtEdit.Text)
sItemId = Grid.TextMatrix(iRow, 32)
sCodProd = Grid.TextMatrix(iRow, 2)

If sItemId = "" Then Exit Sub

Select Case iCol

    Case 1 ' EAN
        sVal = Trim(sVal)
        Do While InStr(sVal, "  ") > 0
            sVal = Replace(sVal, "  ", " ")
        Loop
        sVal = UCase(sVal)
        If sVal = "" Then
            sVal = "SEM GTIN"
        ElseIf sVal = "SEM GTIN" Then
            ' mantem como esta, ja normalizado
        Else
            If Not IsNumeric(sVal) Then
                MsgBox "EAN deve conter apenas dígitos, ou 'SEM GTIN' se o produto não tiver código de barras!", vbInformation, "Aviso"
                Exit Sub
            End If
            If Len(sVal) <> 8 And Len(sVal) <> 13 Then
                MsgBox "EAN deve ter 8 ou 13 dígitos!", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
        dbData.Execute "UPDATE TbNFCe_Itens SET CodBarras = '" & sVal & "' WHERE IdNFProd = " & vPed & " AND IdNFProd_Item = " & Val(sItemId)
        If sVal = "SEM GTIN" Then
            dbData.Execute "UPDATE Produtos SET EAN = '" & sVal & "' WHERE CODIGO = " & Val(sCodProd)
        Else
            dbData.Execute "UPDATE Produtos SET EAN = '" & sVal & "', COD_BARRA = '" & sVal & "' WHERE CODIGO = " & Val(sCodProd)
        End If
        Grid.TextMatrix(iRow, iCol) = sVal

    Case 4 ' UN
        If sVal = "" Then
            MsgBox "Unidade não pode ser vazia!", vbInformation, "Aviso"
            Exit Sub
        End If
        sVal = UCase(sVal)
        Dim sListaUND As String
        sListaUND = "|UN|PC|KG|CX|PA|PT|LT|ML|GR|MG|DZ|FD|RL|JG|KT|LA|GL|BD|SC|PR|M2|M3|CT|EX|BJ|DI|MET|"
        If InStr(sListaUND, "|" & sVal & "|") = 0 Then
            MsgBox "Unidade '" & sVal & "' inválida!" & vbCrLf & "Aceitas: UN PC KG CX PA PT LT ML GR MG DZ FD RL JG KT LA GL BD SC PR M2 M3 CT EX BJ DI MET", vbInformation, "Aviso"
            Exit Sub
        End If
        dbData.Execute "UPDATE TbNFCe_Itens SET UN = '" & sVal & "' WHERE IdNFProd = " & vPed & " AND IdNFProd_Item = " & Val(sItemId)
        dbData.Execute "UPDATE Produtos SET unid_medida = '" & sVal & "' WHERE CODIGO = " & Val(sCodProd)
        Grid.TextMatrix(iRow, iCol) = sVal

    Case 5 ' NCM
        sVal = Replace(sVal, ".", "")
        If sVal <> "" Then
            If Not IsNumeric(sVal) Or Len(sVal) <> 8 Then
                MsgBox "NCM Inválido!", vbInformation, "Aviso do Sistema"
                Exit Sub
            End If
        End If
        dbData.Execute "UPDATE TbNFCe_Itens SET CodNcm = '" & sVal & "' WHERE IdNFProd = " & vPed & " AND IdNFProd_Item = " & Val(sItemId)
        dbData.Execute "UPDATE Produtos SET NCM = '" & sVal & "' WHERE CODIGO = " & Val(sCodProd)
        Grid.TextMatrix(iRow, iCol) = sVal
        AtualizarValorTributosItem vPed, Val(sItemId), sVal, iRow

    Case 6 ' CFOP
        If sVal <> "" Then
            If Not IsNumeric(sVal) Or Len(sVal) <> 4 Then
                MsgBox "CFOP Inválido!", vbInformation, "Aviso do Sistema"
                Exit Sub
            End If
        End If
        dbData.Execute "UPDATE TbNFCe_Itens SET CFOP = " & Val(sVal) & " WHERE IdNFProd = " & vPed & " AND IdNFProd_Item = " & Val(sItemId)
        dbData.Execute "UPDATE Produtos SET CFOP = " & Val(sVal) & " WHERE CODIGO = " & Val(sCodProd)
        Grid.TextMatrix(iRow, iCol) = sVal

    Case 7 ' CST ICMS
        If sVal <> "" And Len(sVal) <> 3 Then
            MsgBox "ICMS CST Inválido!", vbInformation, "Aviso do Sistema"
            Exit Sub
        End If
        sVal = IIf(sVal = "", "000", sVal)
        dbData.Execute "UPDATE TbNFCe_Itens SET ICMSCST = '" & sVal & "' WHERE IdNFProd = " & vPed & " AND IdNFProd_Item = " & Val(sItemId)
        dbData.Execute "UPDATE Produtos SET icmsCST = '" & sVal & "' WHERE CODIGO = " & Val(sCodProd)
        Grid.TextMatrix(iRow, iCol) = sVal

    Case 8 ' ALIQ. ICMS
        sVal = Format(sVal, ocMONEY)
        curBCICMS = CCur(Val(Replace(Replace(Grid.TextMatrix(iRow, 9), ".", ""), ",", ".")))
        curVICMS = CCur(Format(curBCICMS * Val(Replace(Replace(sVal, ".", ""), ",", ".")) / 100, "0.00"))
        dbData.Execute "UPDATE TbNFCe_Itens SET Aliq_Icms = " & fSQL(sVal, 2) & ", Vlr_Icms = " & fSQL(curVICMS, 2) & " WHERE IdNFProd = " & vPed & " AND IdNFProd_Item = " & Val(sItemId)
        Grid.TextMatrix(iRow, iCol) = sVal
        Grid.TextMatrix(iRow, 10) = FormatNumber(curVICMS, 2)

    Case 11 ' CST IBS
        sVal = UCase(sVal)
        If sVal = "" Then
            MsgBox "CST IBS não pode ser vazio!", vbInformation, "Aviso"
            Exit Sub
        End If
        sChkCST = SQLExecutaRetorno("SELECT TOP 1 CST FROM TbIBSCBSClassTrib WHERE CST = '" & Replace(sVal, "'", "''") & "'", "CST", "")
        If sChkCST = "" Then
            MsgBox "CST IBS '" & sVal & "' não encontrado em TbIBSCBSClassTrib!", vbInformation, "Aviso"
            Exit Sub
        End If
        Set rsClassTrib = New ADODB.Recordset
        RsOpen rsClassTrib, "SELECT TOP 1 cClassTrib, pRedIBS, pRedCBS FROM TbIBSCBSClassTrib WHERE CST = '" & Replace(sVal, "'", "''") & "'"
        If rsClassTrib.EOF Then
            rsClassTrib.Close: Set rsClassTrib = Nothing
            Exit Sub
        End If
        sNewClassTrib = rsClassTrib!cClassTrib & ""
        dblPRedIBS = CDbl(rsClassTrib!pRedIBS)
        dblPRedCBS = CDbl(rsClassTrib!pRedCBS)
        rsClassTrib.Close: Set rsClassTrib = Nothing
        Set rsIBSItem = New ADODB.Recordset
        RsOpen rsIBSItem, "SELECT IBS_vBC, IBS_UFpAliq, IBS_MunpAliq, CBS_vBC, CBS_pAliq FROM TbNFCe_Itens WHERE IdNFProd = " & vPed & " AND IdNFProd_Item = " & Val(sItemId)
        If rsIBSItem.EOF Then
            rsIBSItem.Close: Set rsIBSItem = Nothing
            Exit Sub
        End If
        dblIBSvBC = CDbl(rsIBSItem!IBS_vBC)
        dblIBSUFpAliq = CDbl(rsIBSItem!IBS_UFpAliq)
        dblIBSMunpAliq = CDbl(rsIBSItem!IBS_MunpAliq)
        dblCBSvBC = CDbl(rsIBSItem!CBS_vBC)
        dblCBSpAliq = CDbl(rsIBSItem!CBS_pAliq)
        rsIBSItem.Close: Set rsIBSItem = Nothing
        curIBSvIBSUF = CCur(Format(dblIBSvBC * dblIBSUFpAliq * (1 - dblPRedIBS / 100) / 100, "0.00"))
        curIBSvIBSMun = CCur(Format(dblIBSvBC * dblIBSMunpAliq * (1 - dblPRedIBS / 100) / 100, "0.00"))
        curIBSvIBS = curIBSvIBSUF + curIBSvIBSMun
        curCBSvCBS = CCur(Format(dblCBSvBC * dblCBSpAliq * (1 - dblPRedCBS / 100) / 100, "0.00"))
        dbData.Execute "UPDATE TbNFCe_Itens SET IBSCBS_CST = '" & Replace(sVal, "'", "''") & "', cClassTrib = '" & Replace(sNewClassTrib, "'", "''") & "', IBS_pRed = " & fSQL(dblPRedIBS, 4) & ", CBS_pRed = " & fSQL(dblPRedCBS, 4) & ", IBS_vIBSUF = " & fSQL(curIBSvIBSUF, 2) & ", IBS_vIBSMun = " & fSQL(curIBSvIBSMun, 2) & ", IBS_vIBS = " & fSQL(curIBSvIBS, 2) & ", CBS_vCBS = " & fSQL(curCBSvCBS, 2) & " WHERE IdNFProd = " & vPed & " AND IdNFProd_Item = " & Val(sItemId)
        dbData.Execute "UPDATE Produtos SET IBSCBSCST = '" & Replace(sVal, "'", "''") & "', cClassTrib = '" & Replace(sNewClassTrib, "'", "''") & "' WHERE codigo = " & Val(sCodProd)
        Grid.TextMatrix(iRow, 11) = sVal
        Grid.TextMatrix(iRow, 12) = sNewClassTrib
        Grid.TextMatrix(iRow, 13) = FormatNumber(curIBSvIBS, 2)
        Grid.TextMatrix(iRow, 14) = FormatNumber(curCBSvCBS, 2)

    Case 12 ' CLASS. IBS
        sVal = UCase(sVal)
        If sVal = "" Then
            MsgBox "CLASS IBS não pode ser vazio!", vbInformation, "Aviso"
            Exit Sub
        End If
        sISCSTAtual = Trim(Grid.TextMatrix(iRow, 11))
        sChkCST = SQLExecutaRetorno("SELECT TOP 1 cClassTrib FROM TbIBSCBSClassTrib WHERE CST = '" & Replace(sISCSTAtual, "'", "''") & "' AND cClassTrib = '" & Replace(sVal, "'", "''") & "'", "cClassTrib", "")
        If sChkCST = "" Then
            MsgBox "CLASS IBS '" & sVal & "' não pertence ao CST '" & sISCSTAtual & "' em TbIBSCBSClassTrib!", vbInformation, "Aviso"
            Exit Sub
        End If
        Set rsClassTrib = New ADODB.Recordset
        RsOpen rsClassTrib, "SELECT pRedIBS, pRedCBS FROM TbIBSCBSClassTrib WHERE CST = '" & Replace(sISCSTAtual, "'", "''") & "' AND cClassTrib = '" & Replace(sVal, "'", "''") & "'"
        If rsClassTrib.EOF Then
            rsClassTrib.Close: Set rsClassTrib = Nothing
            Exit Sub
        End If
        dblPRedIBS = CDbl(rsClassTrib!pRedIBS)
        dblPRedCBS = CDbl(rsClassTrib!pRedCBS)
        rsClassTrib.Close: Set rsClassTrib = Nothing
        Set rsIBSItem = New ADODB.Recordset
        RsOpen rsIBSItem, "SELECT IBS_vBC, IBS_UFpAliq, IBS_MunpAliq, CBS_vBC, CBS_pAliq FROM TbNFCe_Itens WHERE IdNFProd = " & vPed & " AND IdNFProd_Item = " & Val(sItemId)
        If rsIBSItem.EOF Then
            rsIBSItem.Close: Set rsIBSItem = Nothing
            Exit Sub
        End If
        dblIBSvBC = CDbl(rsIBSItem!IBS_vBC)
        dblIBSUFpAliq = CDbl(rsIBSItem!IBS_UFpAliq)
        dblIBSMunpAliq = CDbl(rsIBSItem!IBS_MunpAliq)
        dblCBSvBC = CDbl(rsIBSItem!CBS_vBC)
        dblCBSpAliq = CDbl(rsIBSItem!CBS_pAliq)
        rsIBSItem.Close: Set rsIBSItem = Nothing
        curIBSvIBSUF = CCur(Format(dblIBSvBC * dblIBSUFpAliq * (1 - dblPRedIBS / 100) / 100, "0.00"))
        curIBSvIBSMun = CCur(Format(dblIBSvBC * dblIBSMunpAliq * (1 - dblPRedIBS / 100) / 100, "0.00"))
        curIBSvIBS = curIBSvIBSUF + curIBSvIBSMun
        curCBSvCBS = CCur(Format(dblCBSvBC * dblCBSpAliq * (1 - dblPRedCBS / 100) / 100, "0.00"))
        dbData.Execute "UPDATE TbNFCe_Itens SET cClassTrib = '" & Replace(sVal, "'", "''") & "', IBS_pRed = " & fSQL(dblPRedIBS, 4) & ", CBS_pRed = " & fSQL(dblPRedCBS, 4) & ", IBS_vIBSUF = " & fSQL(curIBSvIBSUF, 2) & ", IBS_vIBSMun = " & fSQL(curIBSvIBSMun, 2) & ", IBS_vIBS = " & fSQL(curIBSvIBS, 2) & ", CBS_vCBS = " & fSQL(curCBSvCBS, 2) & " WHERE IdNFProd = " & vPed & " AND IdNFProd_Item = " & Val(sItemId)
        dbData.Execute "UPDATE Produtos SET cClassTrib = '" & Replace(sVal, "'", "''") & "' WHERE codigo = " & Val(sCodProd)
        Grid.TextMatrix(iRow, iCol) = sVal
        Grid.TextMatrix(iRow, 13) = FormatNumber(curIBSvIBS, 2)
        Grid.TextMatrix(iRow, 14) = FormatNumber(curCBSvCBS, 2)

    Case 15 ' CST IS
        sVal = Trim(sVal)
        If sVal <> "" And sVal <> "00" And sVal <> "01" And sVal <> "99" Then
            MsgBox "CST IS inválido! Aceitos: vazio, 00, 01, 99", vbInformation, "Aviso"
            Exit Sub
        End If
        If sVal = "" Or sVal = "00" Or sVal = "99" Then
            dbData.Execute "UPDATE TbNFCe_Itens SET IS_CST = '" & sVal & "', cClassTrib_IS = NULL, IS_tipo_calculo = 1, IS_vBC = 0, IS_pAliq = 0, IS_qUnid = 0, IS_vUnid = 0, IS_vIS = 0, uTrib_IS = NULL WHERE IdNFProd = " & vPed & " AND IdNFProd_Item = " & Val(sItemId)
            Grid.TextMatrix(iRow, 15) = sVal
            Grid.TextMatrix(iRow, 16) = ""
            Grid.TextMatrix(iRow, 17) = FormatNumber(0, 2)
            dbData.Execute "UPDATE Produtos SET ISCST = '" & Replace(sVal, "'", "''") & "', cClassTrib_IS = NULL, tipo_calculo_is = 1 WHERE codigo = " & Val(sCodProd)
        Else
            dbData.Execute "UPDATE TbNFCe_Itens SET IS_CST = '01' WHERE IdNFProd = " & vPed & " AND IdNFProd_Item = " & Val(sItemId)
            Grid.TextMatrix(iRow, 15) = "01"
            dbData.Execute "UPDATE Produtos SET ISCST = '01' WHERE codigo = " & Val(sCodProd)
        End If

    Case 16 ' CLASS IS
        sVal = Trim(sVal)
        sISCSTAtual = Trim(Grid.TextMatrix(iRow, 15))
        If sISCSTAtual <> "01" Then
            MsgBox "CLASS IS só pode ser preenchido quando CST IS = '01'!", vbInformation, "Aviso"
            Exit Sub
        End If
        If sVal = "" Then
            dbData.Execute "UPDATE TbNFCe_Itens SET cClassTrib_IS = NULL, IS_tipo_calculo = 1, IS_vBC = 0, IS_pAliq = 0, IS_qUnid = 0, IS_vUnid = 0, IS_vIS = 0, uTrib_IS = NULL WHERE IdNFProd = " & vPed & " AND IdNFProd_Item = " & Val(sItemId)
            Grid.TextMatrix(iRow, iCol) = ""
            Grid.TextMatrix(iRow, 17) = FormatNumber(0, 2)
            dbData.Execute "UPDATE Produtos SET cClassTrib_IS = NULL, tipo_calculo_is = 1 WHERE codigo = " & Val(sCodProd)
        Else
            sChkCST = SQLExecutaRetorno("SELECT TOP 1 cClassTrib_IS FROM tbISClassTrib WHERE cClassTrib_IS = '" & Replace(sVal, "'", "''") & "'", "cClassTrib_IS", "")
            If sChkCST = "" Then
                MsgBox "CLASS IS '" & sVal & "' não encontrado em tbISClassTrib!", vbInformation, "Aviso"
                Exit Sub
            End If
            Set rsISTrib = New ADODB.Recordset
            RsOpen rsISTrib, "SELECT TOP 1 ISCST, tipo_calculo_is, ISpAliq, ISvUnid, uTrib_IS " & _
                             "FROM tbISClassTrib " & _
                             "WHERE cClassTrib_IS = '" & Replace(sVal, "'", "''") & "' " & _
                             "ORDER BY CASE WHEN GETDATE() BETWEEN dIniVig AND dFimVig THEN 0 ELSE 1 END ASC, dFimVig DESC"
            If rsISTrib.EOF Then
                rsISTrib.Close: Set rsISTrib = Nothing
                Exit Sub
            End If
            sISCST2 = IIf(IsNull(rsISTrib!ISCST), "", CStr(rsISTrib!ISCST))
            iTipoIS2 = CInt(IIf(IsNull(rsISTrib!tipo_calculo_is), 0, rsISTrib!tipo_calculo_is))
            dISpAliq2 = CDbl(IIf(IsNull(rsISTrib!ISpAliq), 0, rsISTrib!ISpAliq))
            dISvUnid2 = CDbl(IIf(IsNull(rsISTrib!ISvUnid), 0, rsISTrib!ISvUnid))
            sISTrib2uTrib = IIf(IsNull(rsISTrib!uTrib_IS), "", CStr(rsISTrib!uTrib_IS))
            rsISTrib.Close: Set rsISTrib = Nothing
            Set rsIBSItem = New ADODB.Recordset
            RsOpen rsIBSItem, "SELECT IS_vBC, IS_qUnid FROM TbNFCe_Itens WHERE IdNFProd = " & vPed & " AND IdNFProd_Item = " & Val(sItemId)
            dISvBC2 = 0: dISqUnid2 = 0
            If Not rsIBSItem.EOF Then
                dISvBC2 = CDbl(IIf(IsNull(rsIBSItem!IS_vBC), 0, rsIBSItem!IS_vBC))
                dISqUnid2 = CDbl(IIf(IsNull(rsIBSItem!IS_qUnid), 0, rsIBSItem!IS_qUnid))
            End If
            rsIBSItem.Close: Set rsIBSItem = Nothing
            dCalcISvBC2 = IIf(iTipoIS2 = 1 Or iTipoIS2 = 3, dISvBC2, 0)
            If iTipoIS2 = 0 Or iTipoIS2 = 1 Then dISqUnid2 = 0: dISvUnid2 = 0
            Select Case iTipoIS2
                Case 1: curISvIS2 = CCur(Format(dCalcISvBC2 * dISpAliq2 / 100, "0.00"))
                Case 2: curISvIS2 = CCur(Format(dISqUnid2 * dISvUnid2, "0.00"))
                Case 3: curISvIS2 = CCur(Format(dCalcISvBC2 * dISpAliq2 / 100 + dISqUnid2 * dISvUnid2, "0.00"))
                Case Else: curISvIS2 = 0
            End Select
            dbData.Execute "UPDATE TbNFCe_Itens SET IS_CST = '" & Replace(sISCST2, "'", "''") & "', cClassTrib_IS = '" & Replace(sVal, "'", "''") & "', IS_tipo_calculo = " & iTipoIS2 & ", IS_vBC = " & fSQL(dCalcISvBC2, 2) & ", IS_pAliq = " & fSQL(dISpAliq2, 4) & ", IS_qUnid = " & fSQL(dISqUnid2, 4) & ", IS_vUnid = " & fSQL(dISvUnid2, 4) & ", IS_vIS = " & fSQL(curISvIS2, 2) & ", uTrib_IS = '" & Replace(sISTrib2uTrib, "'", "''") & "' WHERE IdNFProd = " & vPed & " AND IdNFProd_Item = " & Val(sItemId)
            Grid.TextMatrix(iRow, 15) = sISCST2
            Grid.TextMatrix(iRow, 16) = sVal
            Grid.TextMatrix(iRow, 17) = FormatNumber(curISvIS2, 2)
            dbData.Execute "UPDATE Produtos SET ISCST = '" & Replace(sISCST2, "'", "''") & "', cClassTrib_IS = '" & Replace(sVal, "'", "''") & "', tipo_calculo_is = " & iTipoIS2 & " WHERE codigo = " & Val(sCodProd)
        End If

    Case 20 ' FRETE
        sVal = Format(sVal, ocMONEY)
        dbData.Execute "UPDATE TbNFCe_Itens SET Valor_Frete = " & fSQL(sVal, 2) & " WHERE IdNFProd = " & vPed & " AND IdNFProd_Item = " & Val(sItemId)
        Grid.TextMatrix(iRow, iCol) = sVal

    Case 21 ' SEGURO
        sVal = Format(sVal, ocMONEY)
        dbData.Execute "UPDATE TbNFCe_Itens SET Valor_Seguro = " & fSQL(sVal, 2) & " WHERE IdNFProd = " & vPed & " AND IdNFProd_Item = " & Val(sItemId)
        Grid.TextMatrix(iRow, iCol) = sVal

    Case 22 ' OUTROS
        sVal = Format(sVal, ocMONEY)
        dbData.Execute "UPDATE TbNFCe_Itens SET ValorOutras = " & fSQL(sVal, 2) & " WHERE IdNFProd = " & vPed & " AND IdNFProd_Item = " & Val(sItemId)
        Grid.TextMatrix(iRow, iCol) = sVal

    Case 25 ' CST PIS
        If sVal <> "" And Len(sVal) <> 2 Then
            MsgBox "PIS Inválido!", vbInformation, "Aviso do Sistema"
            Exit Sub
        End If
        sVal = IIf(sVal = "", "00", sVal)
        dbData.Execute "UPDATE TbNFCe_Itens SET PISCST = '" & sVal & "' WHERE IdNFProd = " & vPed & " AND IdNFProd_Item = " & Val(sItemId)
        dbData.Execute "UPDATE Produtos SET pisCST = '" & sVal & "' WHERE CODIGO = " & Val(sCodProd)
        Grid.TextMatrix(iRow, iCol) = sVal
        If CSTPisCofinsAliqZero(sVal) And Val(Replace(Replace(Grid.TextMatrix(iRow, 26), ".", ""), ",", ".")) <> 0 Then
            MsgBox "CST PIS '" & sVal & "' exige aliquota zero! A aliquota/valor de PIS desse item foram zerados.", vbExclamation, "Aviso"
            dbData.Execute "UPDATE TbNFCe_Itens SET Aliq_PIS = 0, vlr_PIS = 0 WHERE IdNFProd = " & vPed & " AND IdNFProd_Item = " & Val(sItemId)
            Grid.TextMatrix(iRow, 26) = "0,00"
            Grid.TextMatrix(iRow, 27) = "0,00"
        End If

    Case 26 ' ALIQ. PIS
        sVal = Format(sVal, ocMONEY)
        If CSTPisCofinsAliqZero(Trim(Grid.TextMatrix(iRow, 25))) And Val(Replace(Replace(sVal, ".", ""), ",", ".")) <> 0 Then
            MsgBox "CST PIS '" & Trim(Grid.TextMatrix(iRow, 25)) & "' exige aliquota zero! O valor informado foi descartado.", vbExclamation, "Aviso"
            sVal = "0,00"
        End If
        curBCICMS = CCur(Val(Replace(Replace(Grid.TextMatrix(iRow, 18), ".", ""), ",", ".")) * Val(Replace(Replace(Grid.TextMatrix(iRow, 19), ".", ""), ",", ".")))
        curVICMS = CCur(Format(curBCICMS * Val(Replace(Replace(sVal, ".", ""), ",", ".")) / 100, "0.00"))
        dbData.Execute "UPDATE TbNFCe_Itens SET Aliq_PIS = " & fSQL(sVal, 2) & ", vlr_PIS = " & fSQL(curVICMS, 2) & " WHERE IdNFProd = " & vPed & " AND IdNFProd_Item = " & Val(sItemId)
        Grid.TextMatrix(iRow, iCol) = sVal
        Grid.TextMatrix(iRow, 27) = FormatNumber(curVICMS, 2)

    Case 28 ' CST COFINS
        If sVal <> "" And Len(sVal) <> 2 Then
            MsgBox "COFINS Inválido!", vbInformation, "Aviso do Sistema"
            Exit Sub
        End If
        sVal = IIf(sVal = "", "00", sVal)
        dbData.Execute "UPDATE TbNFCe_Itens SET COFINSCST = '" & sVal & "' WHERE IdNFProd = " & vPed & " AND IdNFProd_Item = " & Val(sItemId)
        dbData.Execute "UPDATE Produtos SET cofinsCST = '" & sVal & "' WHERE CODIGO = " & Val(sCodProd)
        Grid.TextMatrix(iRow, iCol) = sVal
        If CSTPisCofinsAliqZero(sVal) And Val(Replace(Replace(Grid.TextMatrix(iRow, 29), ".", ""), ",", ".")) <> 0 Then
            MsgBox "CST COFINS '" & sVal & "' exige aliquota zero! A aliquota/valor de COFINS desse item foram zerados.", vbExclamation, "Aviso"
            dbData.Execute "UPDATE TbNFCe_Itens SET Aliq_COFINS = 0, vlr_COFINS = 0 WHERE IdNFProd = " & vPed & " AND IdNFProd_Item = " & Val(sItemId)
            Grid.TextMatrix(iRow, 29) = "0,00"
            Grid.TextMatrix(iRow, 30) = "0,00"
        End If

    Case 29 ' ALIQ. COFINS
        sVal = Format(sVal, ocMONEY)
        If CSTPisCofinsAliqZero(Trim(Grid.TextMatrix(iRow, 28))) And Val(Replace(Replace(sVal, ".", ""), ",", ".")) <> 0 Then
            MsgBox "CST COFINS '" & Trim(Grid.TextMatrix(iRow, 28)) & "' exige aliquota zero! O valor informado foi descartado.", vbExclamation, "Aviso"
            sVal = "0,00"
        End If
        curBCICMS = CCur(Val(Replace(Replace(Grid.TextMatrix(iRow, 18), ".", ""), ",", ".")) * Val(Replace(Replace(Grid.TextMatrix(iRow, 19), ".", ""), ",", ".")))
        curVICMS = CCur(Format(curBCICMS * Val(Replace(Replace(sVal, ".", ""), ",", ".")) / 100, "0.00"))
        dbData.Execute "UPDATE TbNFCe_Itens SET Aliq_COFINS = " & fSQL(sVal, 2) & ", vlr_COFINS = " & fSQL(curVICMS, 2) & " WHERE IdNFProd = " & vPed & " AND IdNFProd_Item = " & Val(sItemId)
        Grid.TextMatrix(iRow, iCol) = sVal
        Grid.TextMatrix(iRow, 30) = FormatNumber(curVICMS, 2)

End Select

AtualizarTotalICMSNFCe vPed
AtualizarTotaisNFCe vPed
AtualizarTotalFreteSeguroOutrosNFCe vPed
AtualizarTotalPisCofinsNFCe vPed

End Sub

Private Function CSTPisCofinsAliqZero(ByVal sCST As String) As Boolean
CSTPisCofinsAliqZero = (sCST = "04" Or sCST = "06" Or sCST = "07" Or sCST = "08" Or sCST = "09")
End Function

Private Sub AtualizarTotalICMSNFCe(ByVal vIdNFProd As Long)
dbData.Execute "UPDATE TbNFCe SET " & _
        "BaseCalc_ICMS = (SELECT ISNULL(SUM(Bc_Icms), 0) FROM TbNFCe_Itens WHERE IdNFProd = " & vIdNFProd & " AND Aliq_Icms <> '0.00'), " & _
        "Valor_ICMS = (SELECT ISNULL(SUM(Vlr_Icms), 0) FROM TbNFCe_Itens WHERE IdNFProd = " & vIdNFProd & " AND Aliq_Icms <> '0.00') " & _
        "WHERE IdNFProd = " & vIdNFProd
End Sub

'grava no cabecalho o total de Frete/Seguro/Outras despesas somado dos itens - esses 3 campos
'sao lidos direto pelo TransmitirNFCe (NFe!Valor_Frete/Valor_Seguro/OutrasDespesasAces) pra
'montar o total da nota na hora de gerar a XML, entao precisam ficar em dia com os itens
Private Sub AtualizarTotalFreteSeguroOutrosNFCe(ByVal vIdNFProd As Long)
dbData.Execute "UPDATE TbNFCe SET " & _
        "Valor_Frete = (SELECT ISNULL(SUM(Valor_Frete), 0) FROM TbNFCe_Itens WHERE IdNFProd = " & vIdNFProd & "), " & _
        "Valor_Seguro = (SELECT ISNULL(SUM(Valor_Seguro), 0) FROM TbNFCe_Itens WHERE IdNFProd = " & vIdNFProd & "), " & _
        "OutrasDespesasAces = (SELECT ISNULL(SUM(ValorOutras), 0) FROM TbNFCe_Itens WHERE IdNFProd = " & vIdNFProd & ") " & _
        "WHERE IdNFProd = " & vIdNFProd
End Sub

'grava no cabecalho o total de PIS/COFINS somado dos itens - mesma regra de exclusao de
'CST isento (04,06,07,08,09) que o TransmitirNFCe usa ao montar vlPIS/vlCOFINS
Private Sub AtualizarTotalPisCofinsNFCe(ByVal vIdNFProd As Long)
dbData.Execute "UPDATE TbNFCe SET " & _
        "vPIS = (SELECT ISNULL(SUM(vlr_PIS), 0) FROM TbNFCe_Itens WHERE IdNFProd = " & vIdNFProd & " AND ISNULL(PISCST, '07') NOT IN ('04','06','07','08','09')), " & _
        "vCOFINS = (SELECT ISNULL(SUM(vlr_COFINS), 0) FROM TbNFCe_Itens WHERE IdNFProd = " & vIdNFProd & " AND ISNULL(COFINSCST, '07') NOT IN ('04','06','07','08','09')) " & _
        "WHERE IdNFProd = " & vIdNFProd
End Sub

Private Sub AtualizarTotaisNFCe(ByVal vIdNFProd As Long)
Dim rTotais As ADODB.Recordset
Dim sSQLTot As String
sSQLTot = "SELECT " & _
        "ISNULL(SUM(CBS_vBC), 0) AS vBCCBS, " & _
        "ISNULL(SUM(IBS_vBC), 0) AS vBCIBS, " & _
        "ISNULL(SUM(IBS_vIBSUF), 0) AS vIBSUF, " & _
        "ISNULL(SUM(IBS_vIBSMun), 0) AS vIBSMun, " & _
        "ISNULL(SUM(IBS_vIBS), 0) AS vIBS, " & _
        "ISNULL(SUM(CBS_vCBS), 0) AS vCBS, " & _
        "ISNULL(SUM(IS_vBC), 0) AS vBCIS, " & _
        "ISNULL(SUM(IS_vIS), 0) AS vIS " & _
        "FROM TbNFCe_Itens WHERE IdNFProd = " & vIdNFProd
Set rTotais = dbData.OpenRecordset(sSQLTot)
If Not rTotais.EOF Then
    dbData.Execute "UPDATE TbNFCe SET " & _
            "vBCCBS = " & fSQL(rTotais("vBCCBS"), 2) & ", " & _
            "vBCIBS = " & fSQL(rTotais("vBCIBS"), 2) & ", " & _
            "vIBSUF = " & fSQL(rTotais("vIBSUF"), 2) & ", " & _
            "vIBSMun = " & fSQL(rTotais("vIBSMun"), 2) & ", " & _
            "vIBS = " & fSQL(rTotais("vIBS"), 2) & ", " & _
            "vCBS = " & fSQL(rTotais("vCBS"), 2) & ", " & _
            "vBCIS = " & fSQL(rTotais("vBCIS"), 2) & ", " & _
            "vIS = " & fSQL(rTotais("vIS"), 2) & _
            " WHERE IdNFProd = " & vIdNFProd
End If
If rTotais.State <> 0 Then rTotais.Close
Set rTotais = Nothing
End Sub

Private Sub AtualizarValorTributosItem(ByVal vIdNFProd As Long, ByVal vItem As Long, ByVal sNcm As String, ByVal iLinha As Long)
'recalcula o valor de tributos totais (Lei 12.741/2012) do item quando o NCM muda, via
'tbNCM (tabela do IBPT, sincronizada por frmImportarIBPT.frm) - mesma formula usada na
'NFCeIncluir: valor liquido do item x (federal+estadual+municipal)/100
Dim dPercTrib As Double
Dim curBase As Currency
Dim rNcm As ADODB.Recordset

If sNcm = "" Then
    dbData.Execute "UPDATE TbNFCe_Itens SET ValorTributos = 0 WHERE IdNFProd = " & vIdNFProd & " AND IdNFProd_Item = " & vItem
    Exit Sub
End If

Set rNcm = dbData.OpenRecordset("SELECT ISNULL(nacionalfederal, 0) + ISNULL(estadual, 0) + ISNULL(municipal, 0) AS varPerc FROM tbNCM WHERE NCM = '" & sNcm & "'")
If Not rNcm.EOF Then
    dPercTrib = rNcm("varPerc")
Else
    dPercTrib = 0
End If
If rNcm.State <> 0 Then rNcm.Close
Set rNcm = Nothing

curBase = CCur(Val(Replace(Replace(Grid.TextMatrix(iLinha, 18), ".", ""), ",", ".")) * Val(Replace(Replace(Grid.TextMatrix(iLinha, 19), ".", ""), ",", "."))) - CCur(Val(Replace(Replace(Grid.TextMatrix(iLinha, 23), ".", ""), ",", ".")))

dbData.Execute "UPDATE TbNFCe_Itens SET ValorTributos = " & fSQL(curBase * dPercTrib / 100, 2) & " WHERE IdNFProd = " & vIdNFProd & " AND IdNFProd_Item = " & vItem
End Sub

Private Sub AplicarVisibilidadeGrid()
If Grid.Cols < 33 Then Exit Sub
Grid.ColWidth(8) = IIf(chkICMS.Value = 1, 900, 0)       'ALIQ. ICMS
Grid.ColWidth(9) = IIf(chkICMS.Value = 1, 850, 0)        'BC ICMS
Grid.ColWidth(10) = IIf(chkICMS.Value = 1, 850, 0)       'VLR ICMS
Grid.ColWidth(11) = IIf(chkReforma.Value = 1, 700, 0)    'CST IBS
Grid.ColWidth(12) = IIf(chkReforma.Value = 1, 950, 0)   'CLASS. IBS
Grid.ColWidth(13) = IIf(chkReforma.Value = 1, 800, 0)    'V. IBS
Grid.ColWidth(14) = IIf(chkReforma.Value = 1, 800, 0)    'V. CBS
Grid.ColWidth(15) = IIf(chkReformaIS.Value = 1, 700, 0)  'CST IS
Grid.ColWidth(16) = IIf(chkReformaIS.Value = 1, 950, 0) 'CLASS IS
Grid.ColWidth(17) = IIf(chkReformaIS.Value = 1, 800, 0)  'V. IS
Grid.ColWidth(20) = IIf(chkFrete.Value = 1, 750, 0)      'FRETE
Grid.ColWidth(21) = IIf(chkSeguro.Value = 1, 850, 0)     'SEGURO
Grid.ColWidth(22) = IIf(chkOutros.Value = 1, 850, 0)     'OUTROS
Grid.ColWidth(25) = IIf(chkPis.Value = 1, 750, 0)        'CST PIS
Grid.ColWidth(26) = IIf(chkPis.Value = 1, 950, 0)        'ALIQ. PIS
Grid.ColWidth(27) = IIf(chkPis.Value = 1, 850, 0)        'VLR PIS
Grid.ColWidth(28) = IIf(chkCofins.Value = 1, 950, 0)     'CST COFINS
Grid.ColWidth(29) = IIf(chkCofins.Value = 1, 1050, 0)     'ALIQ. COFINS
Grid.ColWidth(30) = IIf(chkCofins.Value = 1, 1000, 0)     'VLR COFINS
End Sub

Private Sub chkICMS_Click()
AplicarVisibilidadeGrid
End Sub

Private Sub chkPis_Click()
AplicarVisibilidadeGrid
End Sub

Private Sub chkCofins_Click()
AplicarVisibilidadeGrid
End Sub

Private Sub chkFrete_Click()
AplicarVisibilidadeGrid
End Sub

Private Sub chkSeguro_Click()
AplicarVisibilidadeGrid
End Sub

Private Sub chkOutros_Click()
AplicarVisibilidadeGrid
End Sub

Private Sub chkReforma_Click()
AplicarVisibilidadeGrid
End Sub

Private Sub chkReformaIS_Click()
AplicarVisibilidadeGrid
End Sub

