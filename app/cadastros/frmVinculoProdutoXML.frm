VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmVinculoProdutoXML 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vincular Produtos da XML ao Cadastro"
   ClientHeight    =   9240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   10110
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
      Caption         =   "Produto escolhido na Nota Fiscal de Entrada"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   60
      TabIndex        =   1
      Top             =   2640
      Width           =   9960
      Begin VB.Label lblCProdVal 
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
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   3600
      End
      Begin VB.Label lblEANVal 
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
         Height          =   255
         Left            =   6000
         TabIndex        =   3
         Top             =   360
         Width           =   3600
      End
      Begin VB.Label lblXProdVal 
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
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   720
         Width           =   8400
      End
      Begin VB.Label lblUComVal 
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
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   1080
         Width           =   1800
      End
      Begin VB.Label lblVUnComVal 
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
         Height          =   255
         Left            =   4800
         TabIndex        =   6
         Top             =   1080
         Width           =   2400
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "cProd:"
         Height          =   225
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EAN:"
         Height          =   225
         Left            =   5040
         TabIndex        =   21
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descricao:"
         Height          =   225
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unidade:"
         Height          =   225
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vl. Unit. Fornecedor:"
         Height          =   225
         Left            =   3120
         TabIndex        =   24
         Top             =   1080
         Width           =   1620
      End
   End
   Begin VB.Frame fraVinculo 
      Caption         =   "Localizar produto interno"
      Height          =   4020
      Left            =   120
      TabIndex        =   7
      Top             =   4380
      Width           =   9900
      Begin VB.TextBox txtBusca 
         Height          =   315
         Left            =   180
         TabIndex        =   8
         Top             =   480
         Width           =   6840
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   315
         Left            =   7140
         TabIndex        =   9
         Top             =   480
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid lstProdutos
         Height          =   2370
         Left            =   180
         TabIndex        =   10
         Top             =   840
         Width           =   9600
         _ExtentX        =   16933
         _ExtentY        =   4180
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
         Left            =   1620
         TabIndex        =   11
         Text            =   "1"
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label lblBuscaHint 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Digite o nome, EAN ou codigo do produto e clique Buscar:"
         ForeColor       =   &H00808080&
         Height          =   225
         Left            =   180
         TabIndex        =   25
         Top             =   240
         Width           =   6000
      End
      Begin VB.Label lblListaHint 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selecione o produto interno correspondente:"
         ForeColor       =   &H00808080&
         Height          =   225
         Left            =   180
         TabIndex        =   26
         Top             =   3300
         Width           =   4035
      End
      Begin VB.Label lblFracLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fracionamento:"
         Height          =   225
         Left            =   180
         TabIndex        =   27
         Top             =   3600
         Width           =   1320
      End
      Begin VB.Label lblFracHint 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "unidades internas por embalagem do fornecedor"
         ForeColor       =   &H00808080&
         Height          =   225
         Left            =   2460
         TabIndex        =   17
         Top             =   3660
         Width           =   4590
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
      Top             =   8520
      Width           =   2415
   End
   Begin VB.CommandButton cmdCadastrar 
      Caption         =   "C&adastrar da XML"
      Height          =   555
      Left            =   2640
      TabIndex        =   13
      Top             =   8520
      Width           =   2535
   End
   Begin VB.CommandButton cmdDesvincular 
      Caption         =   "&Desvincular"
      Height          =   555
      Left            =   5280
      TabIndex        =   19
      Top             =   8520
      Width           =   1905
   End
   Begin VB.CommandButton cmdEncerrar 
      Caption         =   "&Encerrar"
      Height          =   555
      Left            =   7260
      TabIndex        =   14
      Top             =   8520
      Width           =   2640
   End
   Begin VB.Label lblTituloLista 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Itens para vincular:"
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
      Left            =   120
      TabIndex        =   15
      Top             =   540
      Width           =   1800
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

'--- Tipo para armazenar dados dos itens pendentes ---
Private Type tItemXML
   cProd      As String
   sEAN       As String
   Nome       As String
   uCom       As String
   vUnCom     As Double
   NCM        As String
   CEST       As String
   Vinculado       As Boolean
   IDProdVinculado As Long      '0 = nao vinculado
   '--- Tributacao da NF-e (entrada) ---
   ICMSCST    As String
   ICMSAliq   As Double
   pRedBC     As Double
   IPICST     As String
   IPIAliq    As Double
   PISCST     As String
   PISAliq    As Double
   COFINSCST  As String
   COFINSAliq As Double
   CFOP       As String
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
   cmdVincular.Enabled    = False
   cmdDesvincular.Enabled = False
   cmdCadastrar.Enabled   = False

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
              "       ISNULL(NCM,'') AS NCM, ISNULL(CEST,'') AS CEST, " & _
              "       ISNULL(CST,'') AS ICMSCST, ISNULL(pICMS,0) AS ICMSAliq, " & _
              "       ISNULL(pRedBC,0) AS pRedBC, " & _
              "       ISNULL(IPICST,'') AS IPICST, ISNULL(IPIpIPI,0) AS IPIAliq, " & _
              "       ISNULL(pisCST,'') AS PISCST, ISNULL(PISpPIS,0) AS PISAliq, " & _
              "       ISNULL(cofinsCST,'') AS COFINSCST, " & _
              "       ISNULL(COFINSpCOFINS,0) AS COFINSAliq, " & _
              "       ISNULL(CFOP,'') AS CFOP, " & _
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
         .ICMSAliq = 0
         .pRedBC = 0
         .IPIAliq = 0
         .PISAliq = 0
         .COFINSAliq = 0
         .Vinculado = (rs!jaVinculado <> 0)
         .IDProdVinculado = 0
         On Error Resume Next
         .IDProdVinculado = CLng(rs!IDProdVinculado)
         On Error GoTo ErrForm_Load
         On Error Resume Next
         .vUnCom = CDbl(rs!vUnCom)
         .ICMSAliq = CDbl(rs!ICMSAliq)
         .pRedBC = CDbl(rs!pRedBC)
         .IPIAliq = CDbl(rs!IPIAliq)
         .PISAliq = CDbl(rs!PISAliq)
         .COFINSAliq = CDbl(rs!COFINSAliq)
         On Error GoTo ErrForm_Load
      End With
      n = n + 1
      rs.MoveNext
   Loop
   rs.Close

   'Inicializa FlexGrid
   lstItens.rows = iTotalItens
   lstItens.ColWidth(0) = lstItens.Width - 120
   For n = 0 To iTotalItens - 1
      lstItens.Row = n
      lstItens.Col = 0
      lstItens.Text = FormatItemLista(n)
      lstItens.CellBackColor = lstItens.BackColor
      If arrItens(n).Vinculado Then
         lstItens.CellForeColor = RGB(180, 0, 0)
      Else
         lstItens.CellForeColor = lstItens.ForeColor
      End If
   Next n

   'Inicializa coluna unica do lstProdutos com largura total
   lstProdutos.ColWidth(0) = lstProdutos.Width - 120

   Set TbBusca = New ADODB.Recordset
   ReDim arrIDProduto(0)

   'Seleciona o primeiro item pendente automaticamente
   Dim iPrimeiro As Integer
   iPrimeiro = -1
   For n = 0 To iTotalItens - 1
      If Not arrItens(n).Vinculado Then iPrimeiro = n: Exit For
   Next n
   If iPrimeiro < 0 Then iPrimeiro = 0   'se todos vinculados, vai para o primeiro
   lstItens.Row = iPrimeiro
   lstItens_Click

   Exit Sub

ErrForm_Load:
   MsgBox "Erro ao abrir formulario de vinculacao:" & vbCrLf & _
          "Numero: " & Err.Number & vbCrLf & _
          "Descricao: " & Err.Description, vbCritical, "Erro"
   Unload Me
End Sub

'==============================================================
Private Function FormatItemLista(idx As Integer) As String
   Dim sStatus As String
   If arrItens(idx).Vinculado Then
      sStatus = "[OK]"
   Else
      sStatus = "[  ]"
   End If
   FormatItemLista = sStatus & " | " & _
                     Left(arrItens(idx).cProd & Space(12), 12) & " | " & _
                     Left(arrItens(idx).Nome & Space(30), 30) & " | " & _
                     Left(arrItens(idx).uCom & Space(5), 5) & " | " & _
                     arrItens(idx).sEAN
End Function

'Atualiza texto e cor de uma linha do FlexGrid
Private Sub AtualizarItemLista(idx As Integer)
   lstItens.Row = idx
   lstItens.Col = 0
   lstItens.Text = FormatItemLista(idx)
   If idx = iSelecionado Then
      'Manter azul para o item que continua selecionado
      lstItens.CellBackColor = RGB(0, 102, 204)
      lstItens.CellForeColor = vbWhite
   ElseIf arrItens(idx).Vinculado Then
      lstItens.CellBackColor = lstItens.BackColor
      lstItens.CellForeColor = RGB(180, 0, 0)
   Else
      lstItens.CellBackColor = lstItens.BackColor
      lstItens.CellForeColor = lstItens.ForeColor
   End If
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
      lstItens.Row = iPrev
      lstItens.Col = 0
      lstItens.CellBackColor = lstItens.BackColor
      If arrItens(iPrev).Vinculado Then
         lstItens.CellForeColor = RGB(180, 0, 0)
      Else
         lstItens.CellForeColor = lstItens.ForeColor
      End If
      lstItens.Row = idx
   End If

   iSelecionado = idx

   'Pinta linha selecionada em azul
   lstItens.Col = 0
   lstItens.CellBackColor = RGB(0, 102, 204)
   lstItens.CellForeColor = vbWhite

   'Preenche painel do fornecedor
   lblContador.Caption = "Item " & (idx + 1) & " de " & iTotalItens & _
                         " - Entrada N" & Chr(186) & " " & NumeroEntrada & _
                         IIf(arrItens(idx).Vinculado, "  [JA VINCULADO - clique Vincular para alterar]", "")

   lblCProdVal.Caption = arrItens(idx).cProd
   lblEANVal.Caption = arrItens(idx).sEAN
   lblXProdVal.Caption = arrItens(idx).Nome
   lblUComVal.Caption = arrItens(idx).uCom
   lblVUnComVal.Caption = "R$ " & Format(arrItens(idx).vUnCom, "##,##0.0000")

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
      lstProdutos.Row = 0 : lstProdutos.Col = 0
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

      AdicionarProduto LPad(CStr(TbBusca!Codigo), 6, " ") & " | " & _
                       Left(TbBusca!DESCRICAO & Space(50), 50) & " | " & _
                       Left(TbBusca!unid_medida & Space(5), 5) & " | " & _
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
                  lstProdutos.Row = k - 1
                  lstProdutos.Col = 0
                  lstProdutos.CellForeColor = RGB(180, 0, 0)
                  lstProdutos.CellBackColor = lstProdutos.BackColor
                  Exit For
               End If
            Next k
         End If
      End If
   End If
End Sub

'==============================================================
Private Sub txtBusca_KeyPress(KeyAscii As Integer)
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

   Dim item As tItemXML
   item = arrItens(iSelecionado)

   'Pergunta modo de venda
   Dim resp As Integer
   resp = MsgBox("Como este produto sera vendido?" & vbCrLf & vbCrLf & _
                 "SIM  = Atacado (vende a caixa/embalagem do fornecedor)" & vbCrLf & _
                 "NAO  = Varejo  (vende a unidade individual)", _
                 vbQuestion + vbYesNoCancel, "Modo de Venda")
   If resp = vbCancel Then Exit Sub
   Dim bVarejo As Boolean
   bVarejo = (resp = vbNo)

   'Unidade convertida para 2 chars maiusculo
   Dim sUnidade As String
   sUnidade = ConverterUnidade(item.uCom, bVarejo)

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
   If Len(item.CFOP) >= 2 And Mid(item.CFOP, 2, 1) = "4" Then
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
      sICMSCST = item.ICMSCST: dICMSAliq = item.ICMSAliq
      dpRedBC = item.pRedBC
      sIPICST = item.IPICST: dIPIAliq = item.IPIAliq
      sPISCST = item.PISCST: dPISAliq = item.PISAliq
      sCOFINSCST = item.COFINSCST: dCOFINSAliq = item.COFINSAliq
   End If

   Dim sDesc   As String
   Dim sEANCad As String

   If bVarejo Then
      'Varejo: exibe form para o usuario informar EAN da unidade
      Load frmCadProdXML
      frmCadProdXML.PubNome = item.Nome
      frmCadProdXML.PubUnidade = sUnidade
      frmCadProdXML.PubNCM = item.NCM
      frmCadProdXML.PubCEST = item.CEST
      frmCadProdXML.PubICMSCST = sICMSCST
      frmCadProdXML.PubPISCST = sPISCST
      frmCadProdXML.PubCOFINSCST = sCOFINSCST
      frmCadProdXML.PubValorUnit = item.vUnCom / frac
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
      If MsgBox("Criar produto (ATACADO) com os dados da XML?" & vbCrLf & vbCrLf & _
                "Descricao: " & item.Nome & vbCrLf & _
                "Unidade  : " & sUnidade & vbCrLf & _
                "NCM      : " & item.NCM, _
                vbQuestion + vbYesNo, "Cadastrar Produto - Atacado") = vbNo Then Exit Sub
      sDesc = item.Nome
      sEANCad = item.sEAN
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
          " NCM, CEST, CFOP, Alterado, PedirPeso, CODPROD_FRACAO, QUANT_FRACAO) " & _
          "VALUES (" & _
          novoCodigo & ", 1, 0, 0, 0, 0, 0, 0, " & _
          "'" & sEANCad & "', '" & sEANCad & "', '" & sDesc & "', '', '" & sUnidade & "', '', '', " & _
          "0, '', 0, '', '', " & _
          "'" & sICMSCST & "', " & FSQL(dICMSAliq) & ", " & _
          "'" & sPISCST & "', " & FSQL(dPISAliq) & ", " & _
          "'" & sCOFINSCST & "', " & FSQL(dCOFINSAliq) & ", " & _
          "'" & sIPICST & "', " & FSQL(dIPIAliq) & ", " & _
          FSQL(dpRedBC) & ", " & _
          "'" & item.NCM & "', '" & item.CEST & "', '" & sCFOPSaida & "', 0, 0, 0, 0)"

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

   Dim item As tItemXML
   item = arrItens(iSelecionado)

   Dim custoUnit As Double
   custoUnit = IIf(frac > 0, item.vUnCom / frac, item.vUnCom)

   'EAN da embalagem
   Dim sEANEmb As String
   sEANEmb = item.sEAN
   If Vazio(sEANEmb) Then
      On Error Resume Next
      Dim rsE As New ADODB.Recordset
      RsOpen rsE, "SELECT EAN FROM EntradaEstoqueItens " & _
                  "WHERE CodigoNota = " & NumeroEntrada & _
                  "  AND Referencia = '" & Replace(item.cProd, "'", "''") & "'"
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

   'IDFornecedor
   Dim IDForn As Long
   IDForn = SQLExecutaRetorno("SELECT CodigoCorrentista FROM EntradaEstoque WHERE CodigoNota = " & NumeroEntrada, "CodigoCorrentista", 0)

   'Verifica se o produto interno ja esta vinculado a outro cProd deste fornecedor
   Dim sOutroCProd As String
   sOutroCProd = SQLExecutaRetorno("SELECT TOP 1 cProd + ' - ' + xProd r FROM VinculoXMLProduto " & _
                  "WHERE IDFornecedor = " & IDForn & _
                  "  AND IDProduto = " & IDProdSel & _
                  "  AND cProd <> '" & Replace(item.cProd, "'", "''") & "'", "r", "")
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
                                 "  AND cProd = '" & Replace(item.cProd, "'", "''") & "'", "r", 0)
   Dim sSQL As String
   If qtdExiste = 0 Then
      sSQL = "INSERT INTO VinculoXMLProduto " & _
             "(IDFornecedor, cProd, EANEmbalagem, xProd, uCom, IDProduto, EANProduto, Fracionamento, CustoUnitario, DataAtualizacao) " & _
             "VALUES (" & _
             IDForn & ", " & _
             "'" & Replace(item.cProd, "'", "''") & "', " & _
             "'" & sEANEmb & "', " & _
             "'" & Replace(item.Nome, "'", "''") & "', " & _
             "'" & item.uCom & "', " & _
             IDProdSel & ", " & _
             "'" & sEANProd & "', " & _
             FSQL(frac) & ", " & _
             FSQL(custoUnit) & ", GETDATE())"
   Else
      sSQL = "UPDATE VinculoXMLProduto SET " & _
             "IDProduto = " & IDProdSel & ", " & _
             "EANProduto = '" & sEANProd & "', " & _
             "EANEmbalagem = '" & sEANEmb & "', " & _
             "Fracionamento = " & FSQL(frac) & ", " & _
             "CustoUnitario = " & FSQL(custoUnit) & ", " & _
             "DataAtualizacao = GETDATE() " & _
             "WHERE IDFornecedor = " & IDForn & _
             "  AND cProd = '" & Replace(item.cProd, "'", "''") & "'"
   End If

   Dim msgErro As String
   msgErro = SQLExecuta(sSQL)
   If Not Vazio(msgErro) Then
      MsgBox "Erro ao salvar v" & Chr(237) & "nculo: " & msgErro, vbCritical
      Exit Function
   End If

   'Atualiza CodigoProduto + quantidades em EntradaEstoqueItens
   sSQL = "UPDATE EntradaEstoqueItens SET " & _
          "CodigoProduto = " & IDProdSel & ", " & _
          "QuantidadeComercial = QuantidadeComercial * " & FSQL(frac) & ", " & _
          "ValorUnitarioComercializacao = ValorUnitarioComercializacao / " & FSQL(frac) & " " & _
          "WHERE CodigoNota = " & NumeroEntrada & _
          "  AND Referencia = '" & Replace(item.cProd, "'", "''") & "' " & _
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

   Dim item As tItemXML
   item = arrItens(iSelecionado)

   Dim IDForn As Long
   IDForn = SQLExecutaRetorno("SELECT CodigoCorrentista FROM EntradaEstoque WHERE CodigoNota = " & NumeroEntrada, "CodigoCorrentista", 0)

   'Le fracionamento e produto vinculado antes de excluir
   Dim dFrac As Double
   Dim sDescVinc As String
   dFrac = SQLExecutaRetorno("SELECT ISNULL(Fracionamento,1) r FROM VinculoXMLProduto " & _
                             "WHERE IDFornecedor = " & IDForn & _
                             "  AND cProd = '" & Replace(item.cProd, "'", "''") & "'", "r", 1)
   sDescVinc = SQLExecutaRetorno("SELECT TOP 1 CAST(IDProduto AS VARCHAR) + ' - ' + xProd r FROM VinculoXMLProduto " & _
                                 "WHERE IDFornecedor = " & IDForn & _
                                 "  AND cProd = '" & Replace(item.cProd, "'", "''") & "'", "r", "")
   If dFrac <= 0 Then dFrac = 1

   If MsgBox("Desfazer vinculo do item:" & vbCrLf & vbCrLf & _
             "NF-e    : " & item.cProd & " - " & item.Nome & vbCrLf & _
             "Produto : " & sDescVinc & vbCrLf & vbCrLf & _
             "O CodigoProduto sera zerado na entrada e o vinculo sera removido.", _
             vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar Desvinculo") = vbNo Then Exit Sub

   'Exclui vinculo
   Dim msgErro As String
   msgErro = SQLExecuta("DELETE FROM VinculoXMLProduto " & _
                        "WHERE IDFornecedor = " & IDForn & _
                        "  AND cProd = '" & Replace(item.cProd, "'", "''") & "'")
   If Not Vazio(msgErro) Then
      MsgBox "Erro ao excluir vinculo: " & msgErro, vbCritical
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
             "  AND Referencia = '" & Replace(item.cProd, "'", "''") & "'"
   Else
      sSQL = "UPDATE EntradaEstoqueItens SET CodigoProduto = 0 " & _
             "WHERE CodigoNota = " & NumeroEntrada & _
             "  AND Referencia = '" & Replace(item.cProd, "'", "''") & "'"
   End If
   msgErro = SQLExecuta(sSQL)
   If Not Vazio(msgErro) Then
      MsgBox "Aviso ao reverter quantidades: " & msgErro, vbExclamation
   End If

   arrItens(iSelecionado).Vinculado = False
   arrItens(iSelecionado).IDProdVinculado = 0
   AtualizarItemLista iSelecionado
   MsgBox "Vinculo desfeito com sucesso.", vbInformation
   AtualizarBotoes
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
   lstProdutos.Rows = 1
   lstProdutos.Row = 0
   lstProdutos.Col = 0
   lstProdutos.Text = ""
   lstProdutos.CellForeColor = lstProdutos.ForeColor
   lstProdutos.CellBackColor = lstProdutos.BackColor
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

   cmdVincular.Enabled    = bTemItem And (Not bVinculado) And bTemProd
   cmdDesvincular.Enabled = bTemItem And bVinculado And bTemProd
   cmdCadastrar.Enabled   = bTemItem And (Not bVinculado)
End Sub

'Adiciona uma linha ao FlexGrid de produtos (equivalente a lstProdutos.AddItem)
Private Sub AdicionarProduto(sTexto As String)
   If lstProdutos.TextMatrix(0, 0) = "" Then
      lstProdutos.Row = 0
   Else
      lstProdutos.Rows = lstProdutos.Rows + 1
      lstProdutos.Row = lstProdutos.Rows - 1
   End If
   lstProdutos.Col = 0
   lstProdutos.Text = sTexto
   lstProdutos.CellForeColor = lstProdutos.ForeColor
   lstProdutos.CellBackColor = lstProdutos.BackColor
End Sub
