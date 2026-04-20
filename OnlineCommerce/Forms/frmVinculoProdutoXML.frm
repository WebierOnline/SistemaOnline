VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmVinculoProdutoXML 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vincular Produtos da XML ao Cadastro"
   ClientHeight    =   9240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid lstItens 
      Height          =   1815
      Left            =   60
      TabIndex        =   16
      Top             =   780
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin VB.Frame fraFornecedor 
      Caption         =   "PRODUTO SELECIONADO DA NOTA FISCAL DE ENTRADA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   60
      TabIndex        =   1
      Top             =   2700
      Width           =   9960
      Begin VB.Label lblCProdVal 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   900
         TabIndex        =   2
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label lblEANVal 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   900
         TabIndex        =   3
         Top             =   540
         Width           =   1800
      End
      Begin VB.Label lblXProdVal 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   4
         Top             =   240
         Width           =   5940
      End
      Begin VB.Label lblUComVal 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   5
         Top             =   540
         Width           =   1680
      End
      Begin VB.Label lblVUnComVal 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7440
         TabIndex        =   6
         Top             =   540
         Width           =   2280
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "cProd:"
         Height          =   225
         Left            =   180
         TabIndex        =   20
         Top             =   240
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EAN:"
         Height          =   225
         Left            =   180
         TabIndex        =   21
         Top             =   540
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descricao:"
         Height          =   225
         Left            =   2880
         TabIndex        =   22
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unidade:"
         Height          =   225
         Left            =   2880
         TabIndex        =   23
         Top             =   540
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Unit. Fornecedor:"
         Height          =   195
         Left            =   5640
         TabIndex        =   24
         Top             =   540
         Width           =   1635
      End
   End
   Begin VB.Frame fraVinculo 
      Caption         =   "Localizar produto interno"
      Height          =   4560
      Left            =   60
      TabIndex        =   7
      Top             =   3960
      Width           =   9960
      Begin VB.TextBox txtBusca 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   6840
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   315
         Left            =   7020
         TabIndex        =   9
         Top             =   480
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid lstProdutos 
         Height          =   3330
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   9720
         _ExtentX        =   17145
         _ExtentY        =   5874
         _Version        =   393216
         Rows            =   1
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin VB.TextBox txtFracionamento 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         Text            =   "1"
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label lblBuscaHint 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Digite o ""Nome do Produto"" ou ""Código de Barra"" e clique Buscar:"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   4725
      End
      Begin VB.Label lblListaHint 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*Selecione o produto interno correspondente:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   225
         Left            =   6660
         TabIndex        =   26
         Top             =   4200
         Width           =   3165
      End
      Begin VB.Label lblFracLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fracionamento:"
         Height          =   225
         Left            =   120
         TabIndex        =   27
         Top             =   4200
         Width           =   1320
      End
      Begin VB.Label lblFracHint 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*unidades internas por embalagem do fornecedor"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   2100
         TabIndex        =   17
         Top             =   4200
         Width           =   3465
      End
   End
   Begin VB.CommandButton cmdVincular 
      Caption         =   "&Vincular"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   12
      Top             =   8580
      Width           =   2415
   End
   Begin VB.CommandButton cmdCadastrar 
      Caption         =   "C&adastrar da XML"
      Height          =   555
      Left            =   2640
      TabIndex        =   13
      Top             =   8580
      Width           =   2535
   End
   Begin VB.CommandButton cmdDesvincular 
      Caption         =   "&Desvincular"
      Height          =   555
      Left            =   5280
      TabIndex        =   19
      Top             =   8580
      Width           =   1440
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "&Alterar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   6840
      TabIndex        =   34
      Top             =   8580
      Width           =   1440
   End
   Begin VB.CommandButton cmdEncerrar 
      Caption         =   "&Encerrar"
      Height          =   555
      Left            =   8340
      TabIndex        =   14
      Top             =   8580
      Width           =   1635
   End
   Begin VB.Label Label73 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "    "
      Height          =   195
      Index           =   0
      Left            =   7500
      TabIndex        =   33
      Top             =   540
      Width           =   180
   End
   Begin VB.Label Label73 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "    "
      Height          =   195
      Index           =   2
      Left            =   8820
      TabIndex        =   32
      Top             =   540
      Width           =   180
   End
   Begin VB.Label Label74 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Legenda:"
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
      Index           =   0
      Left            =   6660
      TabIndex        =   31
      Top             =   540
      Width           =   810
   End
   Begin VB.Label Label74 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Com vinculo"
      Height          =   195
      Index           =   1
      Left            =   7800
      TabIndex        =   30
      Top             =   540
      Width           =   870
   End
   Begin VB.Label Label74 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sem vinculo"
      Height          =   195
      Index           =   3
      Left            =   9120
      TabIndex        =   29
      Top             =   540
      Width           =   870
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUTOS DO SEU ESTOQUE ATUAL:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   60
      TabIndex        =   28
      Top             =   3720
      Width           =   3195
   End
   Begin VB.Label lblTituloLista 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUTOS DA NOTA FISCAL DE ENTRADA (XML):"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   15
      Top             =   540
      Width           =   4140
   End
   Begin VB.Label lblContador 
      Alignment       =   2  'Center
      BackColor       =   &H00003399&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9960
   End
   Begin VB.Label lblCustoCalc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   180
      TabIndex        =   18
      Top             =   7020
      Width           =   9960
   End
End
Attribute VB_Name = "frmVinculoProdutoXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================
' frmVinculoProdutoXML
' Formulario de vinculacao de produtos da NF-e ao cadastro interno.
'==============================================================
Option Explicit

'--- Propriedades publicas ---
Public NumeroEntrada As Long
Public MostrarTodos  As Boolean   'True = mostra todos (vinculados + pendentes)
Public cProdParaSelecionar As String

'--- Tipo para armazenar dados dos itens pendentes ---
Private Type tItemXML
   cProd      As String
   sEAN       As String
   Nome       As String
   uCom       As String
   vUnCom     As Double
   vUnTrib    As Double
   qCom       As Double
   NCM        As String
   CEST       As String
   Vinculado       As Boolean
   IDProdVinculado As Long
   '--- Tributacao ICMS ---
   ICMSCST    As String
   ICMSAliq   As Double
   pRedBC     As Double
   modBC      As Integer
   '--- Substituicao Tributaria ---
   pMVAST     As Double
   pICMSST    As Double
   pRedBCST   As Double
   modBCST    As Integer
   '--- IPI / PIS / COFINS ---
   IPICST     As String
   IPIAliq    As Double
   PISCST     As String
   PISAliq    As Double
   COFINSCST  As String
   COFINSAliq As Double
   CFOP       As String
   '--- Reforma Tributaria (IBS/CBS/IS) ---
   IBSCBSCST  As String
   IBSUFpAliq As Double
   IBSMunpAliq As Double
   CBSpAliq   As Double
   ISCST      As String
   ISpIS      As Double
End Type

Private arrItens()   As tItemXML
Private iTotalItens  As Integer
Private iSelecionado As Integer   '-1 = nenhum item selecionado

'--- Recordsets ---
Private TbBusca As ADODB.Recordset

'--- Array paralelo ao lstProdutos ---
Private arrIDProduto() As Long

'==============================================================
Private Sub Form_Load()
   On Error GoTo ErrForm_Load

   iSelecionado = -1

   'Botoes inicialmente desabilitados ate o usuario selecionar itens
   cmdVincular.Enabled = False
   cmdDesvincular.Enabled = False
   cmdCadastrar.Enabled = False
   cmdAlterar.Enabled = False

   'Monta filtro: MostrarTodos carrega todos; caso contrario so pendentes
   Dim sWhere As String
   If MostrarTodos Then
      sWhere = "WHERE CodigoNota = " & NumeroEntrada
   Else
      sWhere = "WHERE CodigoNota = " & NumeroEntrada & _
               "  AND (CodigoProduto = 0 OR CodigoProduto IS NULL)"
   End If

   Dim rs As New ADODB.Recordset
   RsOpen rs, "SELECT Item, Referencia AS cProd, ISNULL(EAN,'') AS EAN, " & _
              "       NomeProduto, UnidadeComercial AS uCom, " & _
              "       ValorUnitarioComercializacao AS vUnCom, " & _
              "       ISNULL(QuantidadeComercial,0) AS qCom, " & _
              "       ISNULL(ValorUnitarioTributario,0) AS vUnTrib, " & _
              "       ISNULL(NCM,'') AS NCM, ISNULL(CEST,'') AS CEST, " & _
              "       ISNULL(CST,'') AS ICMSCST, ISNULL(pICMS,0) AS ICMSAliq, " & _
              "       ISNULL(pRedBC,0) AS pRedBC, ISNULL(modBC,3) AS modBC, " & _
              "       ISNULL(pMVAST,0) AS pMVAST, ISNULL(pICMSST,0) AS pICMSST, " & _
              "       ISNULL(pRedBCST,0) AS pRedBCST, ISNULL(modBCST,4) AS modBCST, " & _
              "       ISNULL(IPICST,'') AS IPICST, ISNULL(IPIpIPI,0) AS IPIAliq, " & _
              "       ISNULL(pisCST,'') AS PISCST, ISNULL(PISpPIS,0) AS PISAliq, " & _
              "       ISNULL(cofinsCST,'') AS COFINSCST, " & _
              "       ISNULL(COFINSpCOFINS,0) AS COFINSAliq, " & _
              "       ISNULL(CFOP,'') AS CFOP, " & _
              "       ISNULL(IBSCBSCST,'') AS IBSCBSCST, " & _
              "       ISNULL(IBSUFpAliq,0) AS IBSUFpAliq, ISNULL(IBSMunpAliq,0) AS IBSMunpAliq, " & _
              "       ISNULL(CBSpAliq,0) AS CBSpAliq, " & _
              "       ISNULL(ISCST,'') AS ISCST, ISNULL(ISpIS,0) AS ISpIS, " & _
              "       ISNULL(CodigoProduto, 0) AS IDProdVinculado, " & _
              "       CASE WHEN (CodigoProduto IS NULL OR CodigoProduto = 0) " & _
              "            THEN 0 ELSE 1 END AS jaVinculado " & _
              "FROM EntradaEstoqueItens " & sWhere & " ORDER BY Item"

   If rs.EOF Then
      If MostrarTodos Then
         MsgBox "Nenhum item encontrado para esta entrada.", vbInformation
      Else
         MsgBox "Todos os produtos desta entrada ja foram identificados.", vbInformation
      End If
      rs.Close
      Unload Me
      Exit Sub
   End If

   'Conta manualmente
   Dim n As Integer
   n = 0
   Do While Not rs.EOF
      n = n + 1
      rs.MoveNext
   Loop
   iTotalItens = n
   rs.MoveFirst

   ReDim arrItens(0 To iTotalItens - 1)

   'Popula array
   n = 0
   Do While Not rs.EOF
      With arrItens(n)
         .cProd = rs!cProd & ""
         .sEAN = rs!EAN & ""
         .Nome = rs!NomeProduto & ""
         .uCom = rs!uCom & ""
         .NCM = rs!NCM & ""
         .CEST = rs!CEST & ""
         .ICMSCST = rs!ICMSCST & ""
         .IPICST = rs!IPICST & ""
         .PISCST = rs!PISCST & ""
         .COFINSCST = rs!COFINSCST & ""
         .CFOP = rs!CFOP & ""
         .vUnCom = 0
         .vUnTrib = 0
         .qCom = 0
         .ICMSAliq = 0: .pRedBC = 0: .modBC = 3
         .pMVAST = 0: .pICMSST = 0: .pRedBCST = 0: .modBCST = 4
         .IPIAliq = 0: .PISAliq = 0: .COFINSAliq = 0
         .IBSUFpAliq = 0: .IBSMunpAliq = 0: .CBSpAliq = 0: .ISpIS = 0
         .IBSCBSCST = rs!IBSCBSCST & "": .ISCST = rs!ISCST & ""
         .Vinculado = (rs!jaVinculado <> 0)
         .IDProdVinculado = 0
         On Error Resume Next
         .IDProdVinculado = CLng(rs!IDProdVinculado)
         On Error GoTo ErrForm_Load
         On Error Resume Next
         .vUnCom = CDbl(rs!vUnCom)
         .vUnTrib = CDbl(rs!vUnTrib)
         .qCom = CDbl(rs!qCom)
         .ICMSAliq = CDbl(rs!ICMSAliq)
         .pRedBC = CDbl(rs!pRedBC): .modBC = CInt(rs!modBC)
         .pMVAST = CDbl(rs!pMVAST): .pICMSST = CDbl(rs!pICMSST)
         .pRedBCST = CDbl(rs!pRedBCST): .modBCST = CInt(rs!modBCST)
         .IPIAliq = CDbl(rs!IPIAliq)
         .PISAliq = CDbl(rs!PISAliq): .COFINSAliq = CDbl(rs!COFINSAliq)
         .IBSUFpAliq = CDbl(rs!IBSUFpAliq): .IBSMunpAliq = CDbl(rs!IBSMunpAliq)
         .CBSpAliq = CDbl(rs!CBSpAliq): .ISpIS = CDbl(rs!ISpIS)
         On Error GoTo ErrForm_Load
      End With
      n = n + 1
      rs.MoveNext
   Loop
   rs.Close

   'Inicializa FlexGrid
   lstItens.rows = iTotalItens
   lstItens.Cols = 5
   lstItens.ColWidth(0) = 1100
   lstItens.ColWidth(1) = 5200
   lstItens.ColWidth(2) = 700
   lstItens.ColWidth(3) = 1500
   lstItens.ColWidth(4) = 600
   lstItens.ColAlignment(0) = 6
   lstItens.ColAlignment(1) = 1
   lstItens.ColAlignment(2) = 1
   lstItens.ColAlignment(3) = 1
   lstItens.ColAlignment(4) = 1
   For n = 0 To iTotalItens - 1
      Dim jcI As Integer, cForeI As Long
      cForeI = IIf(arrItens(n).Vinculado, vbBlack, RGB(0, 0, 128))
      For jcI = 0 To 4
         lstItens.Row = n: lstItens.Col = jcI
         lstItens.CellBackColor = lstItens.BackColor
         lstItens.CellForeColor = cForeI
      Next jcI
      lstItens.Row = n
      lstItens.Col = 0: lstItens.Text = arrItens(n).cProd
      lstItens.Col = 1: lstItens.Text = arrItens(n).Nome
      lstItens.Col = 2: lstItens.Text = arrItens(n).uCom
      lstItens.Col = 3: lstItens.Text = arrItens(n).sEAN
      lstItens.Col = 4: lstItens.Text = IIf(arrItens(n).Vinculado, "[OK]", "")
   Next n

   'Inicializa colunas do lstProdutos
   lstProdutos.Cols = 4
   lstProdutos.ColWidth(0) = 1100
   lstProdutos.ColWidth(1) = 5200
   lstProdutos.ColWidth(2) = 700
   lstProdutos.ColWidth(3) = 1500
   lstProdutos.ColAlignment(0) = 6
   lstProdutos.ColAlignment(1) = 1
   lstProdutos.ColAlignment(2) = 1
   lstProdutos.ColAlignment(3) = 1

   Set TbBusca = New ADODB.Recordset
   ReDim arrIDProduto(0)

   'Seleciona o primeiro item pendente automaticamente
   Dim iPrimeiro As Integer
   iPrimeiro = -1
   For n = 0 To iTotalItens - 1
      If Not arrItens(n).Vinculado Then iPrimeiro = n: Exit For
   Next n
   If iPrimeiro < 0 Then iPrimeiro = 0   'se todos vinculados, vai para o primeiro
   'Se foi pedido para pre-selecionar um produto especifico
   If Trim(cProdParaSelecionar) <> "" Then
      Dim nBusca As Integer
      For nBusca = 0 To iTotalItens - 1
         If arrItens(nBusca).cProd = Trim(cProdParaSelecionar) Then
            iPrimeiro = nBusca
            Exit For
         End If
      Next nBusca
   End If
   lstItens.Row = iPrimeiro
   lstItens_Click

   Exit Sub

ErrForm_Load:
   MsgBox "Erro ao abrir formulario de vinculacao:" & vbCrLf & _
          "Numero: " & Err.Number & vbCrLf & _
          "Descricao: " & Err.Description, vbCritical, "Erro"
   Unload Me
End Sub


'Atualiza texto e cor de uma linha do FlexGrid
Private Sub AtualizarItemLista(idx As Integer)
   Dim jcA As Integer, cBackA As Long, cForeA As Long
   If idx = iSelecionado Then
      cBackA = RGB(0, 102, 204): cForeA = vbWhite
   ElseIf arrItens(idx).Vinculado Then
      cBackA = lstItens.BackColor: cForeA = vbBlack
   Else
      cBackA = lstItens.BackColor: cForeA = RGB(0, 0, 128)
   End If
   For jcA = 0 To 4
      lstItens.Row = idx: lstItens.Col = jcA
      lstItens.CellBackColor = cBackA
      lstItens.CellForeColor = cForeA
   Next jcA
   lstItens.Row = idx
   lstItens.Col = 0: lstItens.Text = arrItens(idx).cProd
   lstItens.Col = 1: lstItens.Text = arrItens(idx).Nome
   lstItens.Col = 2: lstItens.Text = arrItens(idx).uCom
   lstItens.Col = 3: lstItens.Text = arrItens(idx).sEAN
   lstItens.Col = 4: lstItens.Text = IIf(arrItens(idx).Vinculado, "[OK]", "")
End Sub

'==============================================================
Private Sub lstItens_Click()
   Dim idx As Integer
   Dim iPrev As Integer
   iPrev = iSelecionado   'salva selecao anterior antes de atualizar

   idx = lstItens.Row
   If idx < 0 Or idx >= iTotalItens Then Exit Sub

   'Restaura cor da linha anteriormente selecionada
   If iPrev >= 0 And iPrev <> idx Then
      Dim jcP As Integer, cForeP As Long
      cForeP = IIf(arrItens(iPrev).Vinculado, vbBlack, RGB(0, 0, 128))
      For jcP = 0 To 4
         lstItens.Row = iPrev: lstItens.Col = jcP
         lstItens.CellBackColor = lstItens.BackColor
         lstItens.CellForeColor = cForeP
      Next jcP
      lstItens.Row = idx
   End If

   iSelecionado = idx

   'Pinta linha selecionada em azul (todas as colunas)
   Dim jcS As Integer
   For jcS = 0 To 4
      lstItens.Col = jcS
      lstItens.CellBackColor = RGB(0, 102, 204)
      lstItens.CellForeColor = vbWhite
   Next jcS

   'Preenche painel do fornecedor
   lblContador.Caption = "Item " & (idx + 1) & " de " & iTotalItens & _
                         " - Entrada N" & Chr(186) & " " & NumeroEntrada & _
                         IIf(arrItens(idx).Vinculado, "  [JA VINCULADO - clique Vincular para alterar]", "")

   lblCProdVal.Caption = arrItens(idx).cProd
   lblEANVal.Caption = arrItens(idx).sEAN
   lblXProdVal.Caption = arrItens(idx).Nome
   lblUComVal.Caption = arrItens(idx).uCom
   lblVUnComVal.Caption = FormatNumber(arrItens(idx).vUnCom, 2)

   'Limpa busca
   LimparProdutos
   ReDim arrIDProduto(0)
   lblCustoCalc.Caption = ""
   txtFracionamento.Text = "1"

   'Coloca palavras acumuladas no txtBusca (minimo 5 chars) e busca automaticamente
   Dim sNome As String
   Dim sParts() As String
   Dim sTermo As String
   Dim j As Integer
   sNome = Trim(arrItens(idx).Nome)
   sParts = Split(sNome, " ")
   sTermo = ""
   For j = 0 To UBound(sParts)
      If Len(Trim(sParts(j))) > 0 Then
         If Len(sTermo) = 0 Then
            sTermo = Trim(sParts(j))
         Else
            sTermo = sTermo & " " & Trim(sParts(j))
         End If
         If Len(sTermo) >= 5 Then Exit For
      End If
   Next j
   txtBusca.Text = sTermo

   'Busca automaticamente ao selecionar item
   cmdBuscar_Click

   'Atualiza estado dos botoes conforme item selecionado
   AtualizarBotoes
End Sub

'==============================================================
Private Sub cmdBuscar_Click()
   If Vazio(txtBusca.Text) Then
      MsgBox "Digite um termo para buscar.", vbExclamation
      Exit Sub
   End If

   Dim sTermo As String
   sTermo = Replace(Trim(txtBusca.Text), "'", "''")

   If TbBusca.State = 1 Then TbBusca.Close
   RsOpen TbBusca, "SELECT Codigo, DESCRICAO, EAN, unid_medida " & _
                   "FROM Produtos " & _
                   "WHERE DESCRICAO LIKE '%" & sTermo & "%' " & _
                   "   OR EAN        LIKE '%" & sTermo & "%' " & _
                   "   OR CAST(Codigo AS VARCHAR) = '" & sTermo & "' " & _
                   "ORDER BY DESCRICAO"

   LimparProdutos
   ReDim arrIDProduto(0)

   If TbBusca.EOF Then
      lstProdutos.Row = 0: lstProdutos.Col = 0
      lstProdutos.Text = "  (nenhum produto encontrado)"
      Exit Sub
   End If

   Dim n As Integer
   n = 0
   Do While Not TbBusca.EOF
      n = n + 1
      ReDim Preserve arrIDProduto(n)
      arrIDProduto(n) = TbBusca!Codigo

      Dim sEAN As String
      sEAN = TbBusca!EAN & ""
      If Vazio(sEAN) Then sEAN = "(sem EAN)"

      AdicionarProduto CStr(TbBusca!Codigo), _
                       TbBusca!DESCRICAO & "", _
                       TbBusca!unid_medida & "", _
                       sEAN
      TbBusca.MoveNext
   Loop

   'Se o item selecionado ja esta vinculado, pinta o produto vinculado em vermelho
   If iSelecionado >= 0 Then
      If arrItens(iSelecionado).Vinculado Then
         Dim idVinc As Long
         idVinc = arrItens(iSelecionado).IDProdVinculado
         If idVinc > 0 Then
            Dim k As Integer
            For k = 1 To UBound(arrIDProduto)
               If arrIDProduto(k) = idVinc Then
                  Dim jcK As Integer
                  For jcK = 0 To 3
                     lstProdutos.Row = k - 1: lstProdutos.Col = jcK
                     lstProdutos.CellForeColor = RGB(180, 0, 0)
                     lstProdutos.CellBackColor = lstProdutos.BackColor
                  Next jcK
                  Exit For
               End If
            Next k
         End If
      End If
   End If
End Sub

'==============================================================
Private Sub txtBusca_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then cmdBuscar_Click
End Sub

'==============================================================
Private Sub txtFracionamento_Change()
   AtualizarCustoCalc
End Sub

Private Sub lstProdutos_Click()
   AtualizarCustoCalc
   AtualizarBotoes
End Sub

Private Sub AtualizarCustoCalc()
   Dim idxP As Integer
   idxP = lstProdutos.Row + 1
   If iSelecionado < 0 Or idxP > UBound(arrIDProduto) Then
      lblCustoCalc.Caption = ""
      Exit Sub
   End If
   If arrIDProduto(idxP) = 0 Then
      lblCustoCalc.Caption = ""
      Exit Sub
   End If
   Dim frac As Double
   frac = Val(txtFracionamento.Text)
   If frac <= 0 Then frac = 1
   Dim vUnCom As Double
   vUnCom = arrItens(iSelecionado).vUnCom
   Dim custo As Double
   custo = vUnCom / frac
   lblCustoCalc.Caption = "Custo unit" & Chr(225) & "rio calculado: R$ " & Format(custo, "##,##0.0000") & _
                          "   (R$ " & Format(vUnCom, "##,##0.0000") & " / " & CStr(frac) & " unidades)"
End Sub

'==============================================================
Private Sub cmdVincular_Click()
   If iSelecionado < 0 Then
      MsgBox "Selecione um item da lista.", vbExclamation
      Exit Sub
   End If
   Dim idxLista As Integer
   idxLista = lstProdutos.Row + 1
   If idxLista > UBound(arrIDProduto) Or arrIDProduto(idxLista) = 0 Then
      MsgBox "Selecione o produto interno na lista.", vbExclamation
      Exit Sub
   End If

   Dim frac As Double
   frac = Val(txtFracionamento.Text)
   If frac <= 0 Then
      MsgBox "O fracionamento deve ser maior que zero.", vbExclamation
      Exit Sub
   End If

   Dim IDProdSel As Long
   IDProdSel = arrIDProduto(idxLista)

   If Not ExecutarVinculo(IDProdSel, frac) Then Exit Sub

   'Marca item como vinculado na lista
   arrItens(iSelecionado).Vinculado = True
   arrItens(iSelecionado).IDProdVinculado = IDProdSel
   AtualizarItemLista iSelecionado

   'Avanca para o proximo item pendente automaticamente
   AvancarParaProximo

   VerificarConclusao
End Sub

'==============================================================
Private Sub cmdCadastrar_Click()
   If iSelecionado < 0 Then
      MsgBox "Selecione um item da lista.", vbExclamation
      Exit Sub
   End If

   Dim frac As Double
   frac = Val(txtFracionamento.Text)
   If frac <= 0 Then frac = 1

   Dim Item As tItemXML
   Item = arrItens(iSelecionado)

   ' Exibir form de escolha do modo de cadastro
   Load frmModoVenda
   frmModoVenda.SetNome Item.Nome
   frmModoVenda.Show vbModal, Me
   Dim nEscolha As Integer
   nEscolha = frmModoVenda.Escolha
   Unload frmModoVenda
   If nEscolha = 0 Then Exit Sub

   Dim bVarejo As Boolean
   bVarejo = (nEscolha = 2)

   'Unidade convertida para 2 chars maiusculo
   Dim sUnidade As String
   sUnidade = ConverterUnidade(Item.uCom, bVarejo)

   'Tributacao conforme regime da revenda
   Dim iRegime As Integer
   iRegime = var_RegimeEmpresa
   If iRegime = 0 Then
      iRegime = SQLExecutaRetorno("SELECT ISNULL(CRT,1) r FROM empresa", "r", 1)
   End If

   'Determina CFOP de saida pelo 2o digito do CFOP de entrada
   'Digito 4 (ex: 1405,2403,5405) -> 5405 (com ST)
   'Digito 1 (ex: 1102,5102) ou outros -> 5102
   Dim sCFOPSaida As String
   If Len(Item.CFOP) >= 2 And Mid(Item.CFOP, 2, 1) = "4" Then
      sCFOPSaida = "5405"
   Else
      sCFOPSaida = "5102"
   End If

   Dim sICMSCST    As String
   Dim dICMSAliq   As Double
   Dim dpRedBC     As Double
   Dim sIPICST     As String
   Dim dIPIAliq    As Double
   Dim sPISCST     As String
   Dim dPISAliq    As Double
   Dim sCOFINSCST  As String
   Dim dCOFINSAliq As Double

   If iRegime = 1 Or iRegime = 2 Then
      'Simples Nacional: ICMSCST depende do CFOP de saida
      If sCFOPSaida = "5405" Then
         sICMSCST = "500"   '5405 + SN = ICMS-ST retido pelo substituto
      Else
         sICMSCST = "102"   '5102 + SN = tributacao normal sem tributacao
      End If
      dICMSAliq = 0: dpRedBC = 0
      sIPICST = "99": dIPIAliq = 0
      sPISCST = "07": dPISAliq = 0
      sCOFINSCST = "07": dCOFINSAliq = 0
   Else
      'Regime Normal: usa valores da NF-e de entrada
      sICMSCST = Item.ICMSCST: dICMSAliq = Item.ICMSAliq
      dpRedBC = Item.pRedBC
      sIPICST = Item.IPICST: dIPIAliq = Item.IPIAliq
      sPISCST = Item.PISCST: dPISAliq = Item.PISAliq
      sCOFINSCST = Item.COFINSCST: dCOFINSAliq = Item.COFINSAliq
   End If

   Dim sDesc   As String
   Dim sEANCad As String

   If bVarejo Then
      'Varejo: exibe form para o usuario informar EAN da unidade
      Load frmCadProdXML
      frmCadProdXML.PubNome = Item.Nome
      frmCadProdXML.PubUnidade = sUnidade
      frmCadProdXML.PubNCM = Item.NCM
      frmCadProdXML.PubCEST = Item.CEST
      frmCadProdXML.PubICMSCST = sICMSCST
      frmCadProdXML.PubPISCST = sPISCST
      frmCadProdXML.PubCOFINSCST = sCOFINSCST
      frmCadProdXML.PubValorUnit = Item.vUnCom / frac
      frmCadProdXML.PubRegime = iRegime
      frmCadProdXML.InicializarUI
      frmCadProdXML.Show vbModal, Me
      If frmCadProdXML.Cancelado Then
         Unload frmCadProdXML
         Exit Sub
      End If
      sDesc = frmCadProdXML.ResDescricao
      sEANCad = frmCadProdXML.ResEAN
      sUnidade = frmCadProdXML.ResUnidade
      Unload frmCadProdXML
   Else
      'Atacado: usa nome e EAN da caixa conforme NF-e
      Dim sEANAtac As String
      sEANAtac = Trim(Item.sEAN)
      If nEscolha <> 3 And sEANAtac <> "" And UCase(sEANAtac) <> "SEM GTIN" Then
         Dim sEANAtacEsc As String
         sEANAtacEsc = Replace(sEANAtac, "'", "''")
         Dim lQtdDupAtac As Long
         lQtdDupAtac = SQLExecutaRetorno("SELECT COUNT(*) r FROM Produtos WHERE LTRIM(RTRIM(ISNULL(COD_BARRA,''))) = '" & sEANAtacEsc & "' OR LTRIM(RTRIM(ISNULL(EAN,''))) = '" & sEANAtacEsc & "'", "r", 0)
         If lQtdDupAtac > 0 Then
            Dim lCodDupAtac As Long
            lCodDupAtac = SQLExecutaRetorno("SELECT TOP 1 Codigo r FROM Produtos WHERE LTRIM(RTRIM(ISNULL(COD_BARRA,''))) = '" & sEANAtacEsc & "' OR LTRIM(RTRIM(ISNULL(EAN,''))) = '" & sEANAtacEsc & "'", "r", 0)
            Dim sDescAtac As String
            sDescAtac = SQLExecutaRetorno("SELECT ISNULL(DESCRICAO,'') r FROM Produtos WHERE Codigo = " & lCodDupAtac, "r", "")
            MsgBox "Ja existe um produto cadastrado com este codigo de barras:" & vbCrLf & vbCrLf & _
                   "EAN: " & sEANAtac & vbCrLf & _
                   "Descricao: " & sDescAtac & vbCrLf & vbCrLf & _
                   "Verifique se o produto ja esta cadastrado antes de criar um novo.", _
                   vbExclamation, "Codigo de Barras Duplicado"
            Exit Sub
         End If
      End If
      If nEscolha = 3 Then
         If MsgBox("Deseja fazer o cadastro do produto abaixo?" & vbCrLf & vbCrLf & _
                   "Descricao: " & Item.Nome & vbCrLf & _
                   "EAN      : " & Item.sEAN & vbCrLf & _
                   "Unidade  : " & sUnidade, _
                   vbQuestion + vbYesNo, "Cadastrar Produto - Manual") = vbNo Then Exit Sub
      Else
         If MsgBox("Criar produto (ATACADO) com os dados da XML?" & vbCrLf & vbCrLf & _
                   "Descricao: " & Item.Nome & vbCrLf & _
                   "Unidade  : " & sUnidade & vbCrLf & _
                   "NCM      : " & Item.NCM, _
                   vbQuestion + vbYesNo, "Cadastrar Produto - Atacado") = vbNo Then Exit Sub
      End If
      sDesc = Item.Nome
      sEANCad = Item.sEAN
   End If

   ' MANUAL: abre Produtos_Cadastro pre-preenchido com dados da XML
   If nEscolha = 3 Then
      Dim lMaxCodAnt As Long
      lMaxCodAnt = SQLExecutaRetorno("SELECT ISNULL(MAX(CODIGO),0) r FROM Produtos", "r", 0)

      Load Produtos_Cadastro
      Produtos_Cadastro.SSTab1.Tab = 0
      Produtos_Cadastro.CriarNovoProduto

      ' Dados basicos
      If Item.sEAN = "SEM GTIN" Or Item.sEAN = "" Then
          Produtos_Cadastro.txtCodBarra.Text = ""
          Produtos_Cadastro.txtEAN.Text = ""
      Else
          Produtos_Cadastro.txtCodBarra.Text = Item.sEAN
          Produtos_Cadastro.txtEAN.Text = Item.sEAN
      End If
      Produtos_Cadastro.txtDescricao.Text = Item.Nome
      Produtos_Cadastro.cboUnidMedida.Text = ConverterUnidade(Item.uCom, False)
      Produtos_Cadastro.txtNCM.Text = Item.NCM
      Produtos_Cadastro.txtCEST.Text = Item.CEST

      ' CFOP e CST ja convertidos pelo regime
      Produtos_Cadastro.cboCFOP.Text = sCFOPSaida
      Produtos_Cadastro.cboCST.Text = sICMSCST

      ' ICMS
      Produtos_Cadastro.txtICMSAliquota.Text = FormatNumber(dICMSAliq, 2)
      Produtos_Cadastro.txtRedBCAliquota.Text = FormatNumber(dpRedBC, 2)

      ' ST
      Produtos_Cadastro.txtMVA.Text = FormatNumber(Item.pMVAST, 2)
      Produtos_Cadastro.txtSTAliq.Text = FormatNumber(Item.pICMSST, 2)
      Produtos_Cadastro.txtRedBCST.Text = FormatNumber(Item.pRedBCST, 2)
      If Item.modBC >= 0 And Item.modBC <= 3 Then Produtos_Cadastro.cboModBC.ListIndex = Item.modBC
      If Item.modBCST >= 0 And Item.modBCST <= 6 Then Produtos_Cadastro.cboModBCST.ListIndex = Item.modBCST

      ' PIS / COFINS / IPI
      Produtos_Cadastro.txtPISCST.Text = sPISCST
      Produtos_Cadastro.txtPisAliquota.Text = FormatNumber(dPISAliq, 2)
      Produtos_Cadastro.txtCOFINSCST.Text = sCOFINSCST
      Produtos_Cadastro.txtCofinsAliquota.Text = FormatNumber(dCOFINSAliq, 2)
      Produtos_Cadastro.txtIPICST.Text = sIPICST
      Produtos_Cadastro.txtIPIAliquota.Text = FormatNumber(dIPIAliq, 2)

      ' Reforma Tributaria: selecionar combos pelo prefixo de 2 digitos
      Dim kM As Integer
      Dim sIBSM As String
      sIBSM = Left(Item.IBSCBSCST & "  ", 2)
      If Trim(sIBSM) = "" Then sIBSM = "01"
      For kM = 0 To Produtos_Cadastro.cboIBSCBSCST.ListCount - 1
          If Left(Produtos_Cadastro.cboIBSCBSCST.List(kM), 2) = sIBSM Then
              Produtos_Cadastro.cboIBSCBSCST.ListIndex = kM: Exit For
          End If
      Next kM
      Produtos_Cadastro.txtCBSpAliq.Text = FormatNumber(Item.CBSpAliq, 4)
      Produtos_Cadastro.txtIBSUFpAliq.Text = FormatNumber(Item.IBSUFpAliq, 4)
      Produtos_Cadastro.txtIBSMunpAliq.Text = FormatNumber(Item.IBSMunpAliq, 4)
      Dim sISM As String
      sISM = Left(Item.ISCST & "  ", 2)
      If Trim(sISM) = "" Then sISM = "00"
      For kM = 0 To Produtos_Cadastro.cboISCST.ListCount - 1
          If Left(Produtos_Cadastro.cboISCST.List(kM), 2) = sISM Then
              Produtos_Cadastro.cboISCST.ListIndex = kM: Exit For
          End If
      Next kM
      Produtos_Cadastro.txtISpIS.Text = FormatNumber(Item.ISpIS, 4)

      ' Custo e margens
      Produtos_Cadastro.txtCusto.Text = Format(Item.vUnCom / frac, "##,##0.00")
      Produtos_Cadastro.txtMargemVV.Text = "0,00%"
      Produtos_Cadastro.txtMargemVP.Text = "0,00%"
      Produtos_Cadastro.txtMargemAV.Text = "0,00%"
      Produtos_Cadastro.txtMargemAP.Text = "0,00%"
      Produtos_Cadastro.txtValorVV.Text = "0,00"
      Produtos_Cadastro.txtValorVP.Text = "0,00"
      Produtos_Cadastro.txtValorAV.Text = "0,00"
      Produtos_Cadastro.txtValorAP.Text = "0,00"

      Produtos_Cadastro.Show vbModal, Me

      ' Verificar se produto foi salvo (codigo maior que o anterior)
      Dim lNovoCod As Long
      lNovoCod = SQLExecutaRetorno("SELECT ISNULL(MAX(CODIGO),0) r FROM Produtos WHERE CODIGO > " & lMaxCodAnt, "r", 0)
      If lNovoCod > 0 Then
          If Not ExecutarVinculo(lNovoCod, frac) Then Exit Sub
          arrItens(iSelecionado).Vinculado = True
          arrItens(iSelecionado).IDProdVinculado = lNovoCod
          AtualizarItemLista iSelecionado
          AvancarParaProximo
          VerificarConclusao
      Else
          MsgBox "Produto nao salvo. Realize o vinculo manualmente apos o cadastro.", vbInformation
      End If
      Exit Sub
   End If

   'Sanitiza
   sDesc = Replace(Left(sDesc, 100), "'", "''")
   sEANCad = Left(Replace(sEANCad, "'", "''"), 14)

   'Obtem proximo codigo
   Dim novoCodigo As Long
   novoCodigo = SQLExecutaRetorno("SELECT ISNULL(MAX(Codigo),0)+1 r FROM Produtos", "r", 1)

   Dim sSQL As String
   sSQL = "INSERT INTO Produtos " & _
          "(codigo, ativo, destaque, USOCONSUMO, COMBUSTIVEL, MATERIAPRIMA, IMOBILIZADO, FRACIONADO, " & _
          " cod_barra, ean, descricao, fabricante, unid_medida, categoria, PRATELEIRA, " & _
          " quant_min, INF_ADICIONA, quant_estoque, ref, tamanho, " & _
          " ICMSCST, ICMSAliq, PISCST, PISALIQ, COFINSCST, COFINSALIQ, IPICST, IPIALIQ, pRedBc, " & _
          " NCM, CEST, CFOP, Alterado, PedirPeso, CODPROD_FRACAO, QUANT_FRACAO, " & _
          " pMVAST, pICMSST, pRedBCST, modBC, modBCST, " & _
          " IBSCBSCST, CBSpAliq, IBSUFpAliq, IBSMunpAliq, ISCST, ISpIS, ESTOQUE_FISCAL) " & _
          "VALUES (" & _
          novoCodigo & ", 1, 0, 0, 0, 0, 0, 0, " & _
          "'" & sEANCad & "', '" & sEANCad & "', '" & sDesc & "', '', '" & sUnidade & "', '', '', " & _
          "0, '', 0, '', '', " & _
          "'" & sICMSCST & "', " & FSQL(dICMSAliq) & ", " & _
          "'" & sPISCST & "', " & FSQL(dPISAliq) & ", " & _
          "'" & sCOFINSCST & "', " & FSQL(dCOFINSAliq) & ", " & _
          "'" & sIPICST & "', " & FSQL(dIPIAliq) & ", " & _
          FSQL(dpRedBC) & ", " & _
          "'" & Item.NCM & "', '" & Item.CEST & "', '" & sCFOPSaida & "', 0, 0, 0, 0, " & _
          FSQL(Item.pMVAST) & ", " & FSQL(Item.pICMSST) & ", " & FSQL(Item.pRedBCST) & ", " & _
          Item.modBC & ", " & Item.modBCST & ", " & _
          "'" & Item.IBSCBSCST & "', " & FSQL(Item.CBSpAliq) & ", " & _
          FSQL(Item.IBSUFpAliq) & ", " & FSQL(Item.IBSMunpAliq) & ", " & _
          "'" & Item.ISCST & "', " & FSQL(Item.ISpIS) & ", 0)"

   Dim msgErro As String
   msgErro = SQLExecuta(sSQL)
   If Not Vazio(msgErro) Then
      MsgBox "Erro ao cadastrar produto: " & msgErro, vbCritical
      Exit Sub
   End If

   MsgBox "Produto criado com codigo " & novoCodigo & "." & vbCrLf & _
          "Verifique e complete o cadastro depois.", vbInformation

   If Not ExecutarVinculo(novoCodigo, frac) Then Exit Sub

   arrItens(iSelecionado).Vinculado = True
   arrItens(iSelecionado).IDProdVinculado = novoCodigo
   AtualizarItemLista iSelecionado

   AvancarParaProximo
   VerificarConclusao
End Sub

'==============================================================
'Converte unidade comercial do fornecedor para 2 chars maiusculo padrao.
'Se varejo=True, retorna sempre "UN".
Private Function ConverterUnidade(sUCom As String, bVarejo As Boolean) As String
   If bVarejo Then
      ConverterUnidade = "UN"
      Exit Function
   End If
   Dim sU As String
   Dim i As Integer
   sU = ""
   For i = 1 To Len(sUCom)
      Dim c As String
      c = Mid(UCase(sUCom), i, 1)
      If c >= "A" And c <= "Z" Then
         sU = sU & c
      Else
         Exit For
      End If
   Next i
   If Len(sU) = 0 Then sU = UCase(Left(sUCom, 2))
   If Len(sU) > 2 Then sU = Left(sU, 2)
   ConverterUnidade = sU
End Function

'==============================================================
'Executa o vinculo: salva em VinculoXMLProduto e atualiza EntradaEstoqueItens
'Retorna True se ok
Private Function ExecutarVinculo(IDProdSel As Long, frac As Double) As Boolean
   ExecutarVinculo = False
   If iSelecionado < 0 Then Exit Function

   Dim Item As tItemXML
   Item = arrItens(iSelecionado)

   Dim custoUnit As Double
   custoUnit = Item.vUnTrib

   'EAN da embalagem
   Dim sEANEmb As String
   sEANEmb = Item.sEAN
   If Vazio(sEANEmb) Then
      On Error Resume Next
      Dim rsE As New ADODB.Recordset
      RsOpen rsE, "SELECT EAN FROM EntradaEstoqueItens " & _
                  "WHERE CodigoNota = " & NumeroEntrada & _
                  "  AND Referencia = '" & Replace(Item.cProd, "'", "''") & "'"
      If Err.Number = 0 Then
         If Not rsE.EOF Then sEANEmb = Trim(rsE!EAN & "")
      End If
      If rsE.State = 1 Then rsE.Close
      Err.Clear
      On Error GoTo 0
   End If

   'EAN do produto interno
   Dim sEANProd As String
   sEANProd = Trim(SQLExecutaRetorno("SELECT ISNULL(EAN,'') r FROM Produtos WHERE Codigo = " & IDProdSel, "r", ""))

   'Unidade de medida interna
   Dim sUnidMedida As String
   sUnidMedida = Trim(SQLExecutaRetorno("SELECT ISNULL(UNID_MEDIDA,'') r FROM Produtos WHERE Codigo = " & IDProdSel, "r", ""))

   'IDFornecedor
   Dim IDForn As Long
   IDForn = SQLExecutaRetorno("SELECT CodigoCorrentista FROM EntradaEstoque WHERE CodigoNota = " & NumeroEntrada, "CodigoCorrentista", 0)

   'Verifica se o produto interno ja esta vinculado a outro cProd deste fornecedor
   Dim sOutroCProd As String
   sOutroCProd = SQLExecutaRetorno("SELECT TOP 1 cProd + ' - ' + xProd r FROM VinculoXMLProduto " & _
                  "WHERE IDFornecedor = " & IDForn & _
                  "  AND IDProduto = " & IDProdSel & _
                  "  AND cProd <> '" & Replace(Item.cProd, "'", "''") & "'", "r", "")
   If Len(Trim(sOutroCProd)) > 0 Then
      MsgBox "Este produto interno ja esta vinculado a outro item da NF-e:" & vbCrLf & vbCrLf & _
             sOutroCProd & vbCrLf & vbCrLf & _
             "Nao e permitido vincular o mesmo produto a dois itens diferentes." & vbCrLf & _
             "Selecione outro produto na lista ou utilize 'Cadastrar da XML' para criar um novo cadastro.", _
             vbExclamation, "Vinculo Duplicado"
      Exit Function
   End If

   'INSERT ou UPDATE em VinculoXMLProduto
   Dim qtdExiste As Long
   qtdExiste = SQLExecutaRetorno("SELECT COUNT(*) r FROM VinculoXMLProduto " & _
                                 "WHERE IDFornecedor = " & IDForn & _
                                 "  AND cProd = '" & Replace(Item.cProd, "'", "''") & "'", "r", 0)
   Dim sSQL As String
   If qtdExiste = 0 Then
      sSQL = "INSERT INTO VinculoXMLProduto " & _
             "(IDFornecedor, cProd, EANEmbalagem, xProd, uCom, QuantidadeComercial, ValorUnitarioComercializacao, " & _
             "IDProduto, EANProduto, UNID_MEDIDA, Fracionamento, CustoUnitario, DataAtualizacao) " & _
             "VALUES (" & _
             IDForn & ", " & _
             "'" & Replace(Item.cProd, "'", "''") & "', " & _
             "'" & sEANEmb & "', " & _
             "'" & Replace(Item.Nome, "'", "''") & "', " & _
             "'" & Item.uCom & "', " & _
             FSQL(Item.qCom) & ", " & _
             FSQL(Item.vUnCom) & ", " & _
             IDProdSel & ", " & _
             "'" & sEANProd & "', " & _
             "'" & sUnidMedida & "', " & _
             FSQL(frac) & ", " & _
             FSQL(custoUnit) & ", GETDATE())"
   Else
      sSQL = "UPDATE VinculoXMLProduto SET " & _
             "EANEmbalagem = '" & sEANEmb & "', " & _
             "xProd = '" & Replace(Item.Nome, "'", "''") & "', " & _
             "uCom = '" & Item.uCom & "', " & _
             "QuantidadeComercial = " & FSQL(Item.qCom) & ", " & _
             "ValorUnitarioComercializacao = " & FSQL(Item.vUnCom) & ", " & _
             "IDProduto = " & IDProdSel & ", " & _
             "EANProduto = '" & sEANProd & "', " & _
             "UNID_MEDIDA = '" & sUnidMedida & "', " & _
             "Fracionamento = " & FSQL(frac) & ", " & _
             "CustoUnitario = " & FSQL(custoUnit) & ", " & _
             "DataAtualizacao = GETDATE() " & _
             "WHERE IDFornecedor = " & IDForn & _
             "  AND cProd = '" & Replace(Item.cProd, "'", "''") & "'"
   End If

   Dim msgErro As String
   msgErro = SQLExecuta(sSQL)
   If Not Vazio(msgErro) Then
      MsgBox "Erro ao salvar v" & Chr(237) & "nculo: " & msgErro, vbCritical
      Exit Function
   End If

   'Atualiza EANEmbalagem e Fracionamento em Produtos
   SQLExecuta "UPDATE Produtos SET EANEmbalagem = '" & sEANEmb & "', " & _
              "Fracionamento = " & FSQL(frac) & " WHERE Codigo = " & IDProdSel

   'Vincula CodigoProduto em EntradaEstoqueItens (dados brutos da XML preservados)
   sSQL = "UPDATE EntradaEstoqueItens SET " & _
          "CodigoProduto = " & IDProdSel & " " & _
          "WHERE CodigoNota = " & NumeroEntrada & _
          "  AND Referencia = '" & Replace(Item.cProd, "'", "''") & "' " & _
          "  AND (CodigoProduto = 0 OR CodigoProduto IS NULL)"
   msgErro = SQLExecuta(sSQL)
   If Not Vazio(msgErro) Then
      MsgBox "Erro ao atualizar item na entrada: " & msgErro, vbCritical
      Exit Function
   End If

   ExecutarVinculo = True
End Function

'==============================================================
Private Sub AvancarParaProximo()
   'Seleciona o proximo item pendente na lista
   Dim k As Integer
   For k = 0 To iTotalItens - 1
      If Not arrItens(k).Vinculado Then
         lstItens.Row = k
         lstItens_Click
         Exit Sub
      End If
   Next k
End Sub

'==============================================================
Private Sub VerificarConclusao()
   Dim pendentes As Integer
   Dim k As Integer
   For k = 0 To iTotalItens - 1
      If Not arrItens(k).Vinculado Then pendentes = pendentes + 1
   Next k
   If pendentes = 0 Then
      MsgBox "Todos os " & iTotalItens & " produto(s) foram vinculados!", vbInformation
   End If
End Sub

'==============================================================
Private Sub cmdDesvincular_Click()
   If iSelecionado < 0 Then
      MsgBox "Selecione um item da lista.", vbExclamation
      Exit Sub
   End If
   If Not arrItens(iSelecionado).Vinculado Then
      MsgBox "Este item nao possui vinculo para desfazer.", vbExclamation
      Exit Sub
   End If
   Dim lAdicionada As Long
   lAdicionada = SQLExecutaRetorno("SELECT COUNT(*) r FROM EntradaEstoqueItens " & _
                                  "WHERE CodigoNota = " & NumeroEntrada & _
                                  " AND Referencia = '" & Replace(arrItens(iSelecionado).cProd, "'", "''") & "'" & _
                                  " AND Adicionada = 1", "r", 0)
   If lAdicionada > 0 Then
      MsgBox "Esse produto ja foi adicionado ao seu estoque!" & vbCrLf & _
             "Para aceitar desvincular, voce precisa remove-lo da entrada de itens.", _
             vbExclamation, "Operacao nao permitida"
      Exit Sub
   End If

   Dim Item As tItemXML
   Item = arrItens(iSelecionado)

   Dim IDForn As Long
   IDForn = SQLExecutaRetorno("SELECT CodigoCorrentista FROM EntradaEstoque WHERE CodigoNota = " & NumeroEntrada, "CodigoCorrentista", 0)

   'Le fracionamento e produto vinculado antes de excluir
   Dim dFrac As Double
   Dim sDescVinc As String
   dFrac = SQLExecutaRetorno("SELECT ISNULL(Fracionamento,1) r FROM VinculoXMLProduto " & _
                             "WHERE IDFornecedor = " & IDForn & _
                             "  AND cProd = '" & Replace(Item.cProd, "'", "''") & "'", "r", 1)
   sDescVinc = SQLExecutaRetorno("SELECT TOP 1 CAST(IDProduto AS VARCHAR) + ' - ' + xProd r FROM VinculoXMLProduto " & _
                                 "WHERE IDFornecedor = " & IDForn & _
                                 "  AND cProd = '" & Replace(Item.cProd, "'", "''") & "'", "r", "")
   If dFrac <= 0 Then dFrac = 1

   If MsgBox("Desfazer vinculo do item:" & vbCrLf & vbCrLf & _
             "NF-e    : " & Item.cProd & " - " & Item.Nome & vbCrLf & _
             "Produto : " & sDescVinc & vbCrLf & vbCrLf & _
             "O CodigoProduto sera zerado na entrada e o produto sera desvinculado.", _
             vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar Desvinculo") = vbNo Then Exit Sub

   'Limpa campos do produto (mantem dados do XML)
   Dim msgErro As String
   msgErro = SQLExecuta("UPDATE VinculoXMLProduto SET " & _
                        "IDProduto = 0, EANProduto = '', Fracionamento = 1, " & _
                        "CustoUnitario = 0, DataAtualizacao = GETDATE() " & _
                        "WHERE IDFornecedor = " & IDForn & _
                        "  AND cProd = '" & Replace(Item.cProd, "'", "''") & "'")
   If Not Vazio(msgErro) Then
      MsgBox "Erro ao desvincular: " & msgErro, vbCritical
      Exit Sub
   End If

   'Reverte CodigoProduto e quantidades na entrada
   Dim sSQL As String
   If dFrac <> 1 Then
      sSQL = "UPDATE EntradaEstoqueItens SET " & _
             "CodigoProduto = 0, " & _
             "QuantidadeComercial = QuantidadeComercial / " & FSQL(dFrac) & ", " & _
             "ValorUnitarioComercializacao = ValorUnitarioComercializacao * " & FSQL(dFrac) & " " & _
             "WHERE CodigoNota = " & NumeroEntrada & _
             "  AND Referencia = '" & Replace(Item.cProd, "'", "''") & "'"
   Else
      sSQL = "UPDATE EntradaEstoqueItens SET CodigoProduto = 0 " & _
             "WHERE CodigoNota = " & NumeroEntrada & _
             "  AND Referencia = '" & Replace(Item.cProd, "'", "''") & "'"
   End If
   msgErro = SQLExecuta(sSQL)
   If Not Vazio(msgErro) Then
      MsgBox "Aviso ao reverter quantidades: " & msgErro, vbExclamation
   End If

   arrItens(iSelecionado).Vinculado = False
   arrItens(iSelecionado).IDProdVinculado = 0
   AtualizarItemLista iSelecionado
   cmdBuscar_Click
   MsgBox "Vinculo desfeito com sucesso.", vbInformation
   AtualizarBotoes
End Sub

Private Sub cmdAlterar_Click()
   Dim idxP As Integer
   idxP = lstProdutos.Row + 1
   If UBound(arrIDProduto) < idxP Then Exit Sub
   Dim lCodProd As Long
   lCodProd = arrIDProduto(idxP)
   If lCodProd <= 0 Then Exit Sub

   Load Produtos_Cadastro
   Produtos_Cadastro.SSTab1.Tab = 0
   Produtos_Cadastro.EditarProduto lCodProd
   Produtos_Cadastro.Show vbModal, Me
   Unload Produtos_Cadastro

   ' Recarrega lstProdutos para exibir dados atualizados
   cmdBuscar_Click
End Sub

'==============================================================
Private Sub cmdEncerrar_Click()
   Dim pendentes As Integer
   Dim k As Integer
   For k = 0 To iTotalItens - 1
      If Not arrItens(k).Vinculado Then pendentes = pendentes + 1
   Next k
   If pendentes > 0 Then
      If MsgBox(pendentes & " produto(s) ainda sem v" & Chr(237) & "nculo." & vbCrLf & _
                "Deseja encerrar mesmo assim?", vbQuestion + vbYesNo, "Pendentes") = vbNo Then
         Exit Sub
      End If
   End If
   Unload Me
End Sub

'==============================================================
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'Permite fechar por codigo (UnloadMode=1) sem restricao
   If UnloadMode = vbFormCode Then Exit Sub

   Dim pendentes As Integer
   Dim k As Integer
   For k = 0 To iTotalItens - 1
      If Not arrItens(k).Vinculado Then pendentes = pendentes + 1
   Next k
   If pendentes > 0 Then
      If MsgBox(pendentes & " produto(s) ainda sem v" & Chr(237) & "nculo." & vbCrLf & _
                "Deseja fechar mesmo assim?", vbQuestion + vbYesNo, "Pendentes") = vbNo Then
         Cancel = True
      End If
   End If
End Sub

'==============================================================
Private Sub Form_Unload(Cancel As Integer)
   If Not TbBusca Is Nothing Then
      If TbBusca.State = 1 Then TbBusca.Close
      Set TbBusca = Nothing
   End If
End Sub

'==============================================================
'Limpa o FlexGrid de produtos (equivalente a lstProdutos.Clear)
Private Sub LimparProdutos()
   lstProdutos.rows = 1
   lstProdutos.Row = 0
   Dim jcL As Integer
   For jcL = 0 To lstProdutos.Cols - 1
      lstProdutos.Col = jcL
      lstProdutos.Text = ""
      lstProdutos.CellForeColor = lstProdutos.ForeColor
      lstProdutos.CellBackColor = lstProdutos.BackColor
   Next jcL
End Sub

'Atualiza o Enabled dos botoes conforme o estado atual da selecao
'  cmdVincular   : item pendente  E produto selecionado em lstProdutos
'  cmdDesvincular: item vinculado E produto selecionado em lstProdutos
'  cmdCadastrar  : item pendente  (independe de lstProdutos)
Private Sub AtualizarBotoes()
   Dim bTemItem   As Boolean   'lstItens tem item valido selecionado
   Dim bVinculado As Boolean   'item selecionado ja possui vinculo
   Dim bTemProd   As Boolean   'lstProdutos tem produto valido selecionado

   bTemItem = (iSelecionado >= 0)
   If bTemItem Then bVinculado = arrItens(iSelecionado).Vinculado

   Dim idxP As Integer
   idxP = lstProdutos.Row + 1
   bTemProd = False
   If UBound(arrIDProduto) >= idxP Then
      bTemProd = (arrIDProduto(idxP) > 0)
   End If

   cmdVincular.Enabled = bTemItem And (Not bVinculado) And bTemProd
   cmdDesvincular.Enabled = bTemItem And bVinculado And bTemProd
   cmdCadastrar.Enabled = bTemItem And (Not bVinculado)
   cmdAlterar.Enabled = bTemProd
End Sub

'Adiciona uma linha ao FlexGrid de produtos (equivalente a lstProdutos.AddItem)
Private Sub AdicionarProduto(sCodigo As String, sNome As String, sUnid As String, sEAN As String)
   If lstProdutos.TextMatrix(0, 0) = "" Then
      lstProdutos.Row = 0
   Else
      lstProdutos.rows = lstProdutos.rows + 1
      lstProdutos.Row = lstProdutos.rows - 1
   End If
   Dim jcAP As Integer
   For jcAP = 0 To 3
      lstProdutos.Col = jcAP
      lstProdutos.CellForeColor = lstProdutos.ForeColor
      lstProdutos.CellBackColor = lstProdutos.BackColor
   Next jcAP
   lstProdutos.Col = 0: lstProdutos.Text = sCodigo
   lstProdutos.Col = 1: lstProdutos.Text = sNome
   lstProdutos.Col = 2: lstProdutos.Text = sUnid
   lstProdutos.Col = 3: lstProdutos.Text = sEAN
End Sub
