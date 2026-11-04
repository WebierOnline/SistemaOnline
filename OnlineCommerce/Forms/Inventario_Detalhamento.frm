VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Inventario_Detalhamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONSULTA DE INVENTÁRIO"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   Icon            =   "Inventario_Detalhamento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmCriterios 
      Caption         =   "Critérios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   915
      Left            =   60
      TabIndex        =   12
      Top             =   8460
      Width           =   1935
      Begin VB.ComboBox cboCriterios 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame frmOrdemComum 
      Caption         =   "Ordem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   915
      Left            =   2040
      TabIndex        =   11
      Top             =   8460
      Width           =   3015
      Begin VB.ComboBox cboOrdem2 
         Height          =   315
         Left            =   1800
         TabIndex        =   9
         Top             =   480
         Width           =   1100
      End
      Begin VB.ComboBox cboOrdem 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1635
      End
   End
   Begin VB.Frame frmFiltroComum 
      Caption         =   "Filtro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   915
      Left            =   5100
      TabIndex        =   5
      Top             =   8460
      Width           =   5700
      Begin VB.OptionButton optPorPalavra 
         Caption         =   "Palavra"
         Height          =   195
         Left            =   3180
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optPalavrasDuplas 
         Caption         =   "Palavras Duplas"
         Height          =   195
         Left            =   4140
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.OptionButton optPorIniciais 
         Caption         =   "Iniciais"
         Height          =   195
         Left            =   2340
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.OptionButton optCompleto 
         Caption         =   "Completo"
         Height          =   195
         Left            =   1320
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox cboConsProduto 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   4200
      End
      Begin ChamaleonBtn.chameleonButton cmdLocalizar 
         Height          =   315
         Left            =   4440
         TabIndex        =   8
         Top             =   480
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Consultar"
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
         MICON           =   "Inventario_Detalhamento.frx":1D82
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblNomeCombo 
         Caption         =   "Nome"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   1155
      End
   End
   Begin VB.PictureBox picAguarde 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   6300
      Picture         =   "Inventario_Detalhamento.frx":1D9E
      ScaleHeight     =   1095
      ScaleWidth      =   2895
      TabIndex        =   4
      Top             =   3540
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   2520
      TabIndex        =   3
      Top             =   2640
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
      ScaleWidth      =   15105
      TabIndex        =   0
      Top             =   60
      Width           =   15135
      Begin VB.Label lblValorTotalFiscal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "EXERCICIO XXXX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   11400
         TabIndex        =   19
         Top             =   480
         Width           =   3465
      End
      Begin VB.Label lblExercicio 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "EXERCICIO XXXX"
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
         Left            =   11400
         TabIndex        =   14
         Top             =   240
         Width           =   3465
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   480
         Picture         =   "Inventario_Detalhamento.frx":2DD6
         Top             =   60
         Width           =   690
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "INVENTÁRIO"
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
         Left            =   1365
         TabIndex        =   1
         Top             =   240
         Width           =   1860
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   7035
      Left            =   60
      TabIndex        =   2
      Top             =   1020
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   12409
      _Version        =   393216
      Cols            =   5
      AllowBigSelection=   0   'False
      HighLight       =   0
      ScrollBars      =   2
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdFechar 
      Height          =   315
      Left            =   13320
      TabIndex        =   20
      Top             =   8100
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "FECHAR INVENTÁRIO"
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
      MICON           =   "Inventario_Detalhamento.frx":423F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdEditarProduto 
      Height          =   315
      Left            =   10200
      TabIndex        =   21
      Top             =   8100
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "EDITAR PRODUTO"
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
      MICON           =   "Inventario_Detalhamento.frx":425B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdEditarCusto 
      Height          =   315
      Left            =   11880
      TabIndex        =   22
      Top             =   8100
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "EDITAR CUSTO"
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
      MICON           =   "Inventario_Detalhamento.frx":4277
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "Inventario_Detalhamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sSQL As String
Dim r As ADODB.Recordset
Dim printSQL As String
Private moCombo As cComboHelper
Dim mlIdInvent As Long          'inventario aberto (filtra a busca por produto)
Dim mbFechado As Boolean        'inventario fechado = so consulta; aberto = contagem fisica na coluna CONTADO
Dim mbEditando As Boolean       'editor da coluna CONTADO aberto
Dim mlLinhaEdit As Long         'linha que esta sendo editada (o usuario pode clicar em outra antes do editor perder o foco)

Public Sub LerInventario(ByVal vIDInvent As Long)
mlIdInvent = vIDInvent
mbFechado = (SQLExecutaRetorno("SELECT COUNT(*) r FROM TbInventarios WHERE (IdInventario = " & vIDInvent & ") AND (Fechado = 1)", "r", 0) > 0)
If mbFechado Then
   Me.Caption = "CONSULTA DE INVENTÁRIO - FECHADO (somente leitura)"
Else
   Me.Caption = "CONSULTA DE INVENTÁRIO - EM CONTAGEM: clique na coluna CONTAGEM"
End If
If SQLExecutaRetorno("SELECT COUNT(*) r FROM TbInventarios WHERE (IdInventario = " & vIDInvent & ") AND (PeriodoInventarioTipo = 'Mensal') AND (ISNUMERIC(PeriodoInventario) = 1)", "r", 0) > 0 Then
   lblExercicio.Caption = "EXERCÍCIO: " & NomeMes(Val(SQLExecutaRetorno("SELECT PeriodoInventario FROM TbInventarios WHERE (IdInventario = " & vIDInvent & ")", "PeriodoInventario", "0"))) & "/" & SQLExecutaRetorno("SELECT AnoInventario FROM TbInventarios WHERE (IdInventario = " & vIDInvent & ")", "AnoInventario", "")
Else
   lblExercicio.Caption = "EXERCÍCIO: " & SQLExecutaRetorno("SELECT AnoInventario FROM TbInventarios WHERE (IdInventario = " & vIDInvent & ")", "AnoInventario", "")
End If
AtualizarBotaoFechar

cboCriterios.Text = "TODOS"
cboOrdem.Text = "SEQ"
cboOrdem2.Text = "ASC"
cboConsProduto.Text = ""
AplicarCriterio False
cmdLocalizar_Click
End Sub


'Nome do mes por extenso (nao depende do idioma do Windows)
Private Function NomeMes(ByVal vMes As Integer) As String
Dim a As Variant
a = Array("JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO")
If vMes >= 1 And vMes <= 12 Then NomeMes = a(vMes - 1)
End Function

'Quantidade: produto vendido em KG com 3 casas; os demais inteiros (se vier fracao, mostra as casas)
Private Function FmtQtd(ByVal dQtd As Double, ByVal bKg As Boolean) As String
If bKg Then
   FmtQtd = Format(dQtd, ocPESO)
ElseIf dQtd = Int(dQtd) Then
   FmtQtd = Format(dQtd, "#,##0")
Else
   FmtQtd = Format(dQtd, ocPESO)
End If
End Function

Private Sub AtualizarBotaoFechar()
If mbFechado Then
   cmdFechar.Caption = "EDITAR INVENTÁRIO"
Else
   cmdFechar.Caption = "FECHAR INVENTÁRIO"
End If
End Sub

'Fecha o inventario (trava a contagem e vira base do estoque fiscal) ou reabre
Private Sub cmdFechar_Click()
If mlIdInvent = 0 Then Exit Sub

If mbFechado Then
   If ShowMsg("Editar o inventário nº " & mlIdInvent & "?" & vbCrLf & "Ele deixa de ser base do estoque fiscal e volta a aceitar contagem.", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
Else
   If ShowMsg("Fechar o inventário nº " & mlIdInvent & "?" & vbCrLf & "A contagem será travada e ele passa a ser a BASE do estoque fiscal (o saldo fiscal é recalculado a partir dele).", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

If SQLExecutaTratandoErro("EXEC InventarioFechar " & mlIdInvent & ", " & IIf(mbFechado, 0, 1)) Then Exit Sub

LerInventario mlIdInvent
End Sub

'Abre o cadastro do produto da linha selecionada (coluna 2 = CODIGO)
Private Sub cmdEditarProduto_Click()
Dim lCod As Long

If Grid.Row < 1 Then
   MsgBox "Selecione um produto na lista.", vbExclamation
   Exit Sub
End If

lCod = Val(Grid.TextMatrix(Grid.Row, 2))
If lCod <= 0 Then
   MsgBox "Selecione um produto na lista.", vbExclamation
   Exit Sub
End If

Load Produtos_Cadastro
Produtos_Cadastro.SSTab1.Tab = 0
Produtos_Cadastro.EditarProduto lCod
Produtos_Cadastro.Show vbModal, Me
Unload Produtos_Cadastro

'a descricao do item e uma copia gravada na geracao: com o inventario ABERTO acompanha o cadastro; FECHADO fica congelado
If Not mbFechado Then
   SQLExecutaTratandoErro "UPDATE i SET i.NomeProduto = ISNULL(p.DESCRICAO, '') FROM TbInventariosItens i INNER JOIN Produtos p ON (p.CODIGO = i.IDProduto) WHERE (i.IdInventario = " & mlIdInvent & ") AND (i.IDProduto = " & lCod & ")"
End If

'recarrega mantendo criterio/ordem (o cadastro pode ter mudado descricao, NCM, custo)
cmdLocalizar_Click
End Sub

'Altera o custo do produto da linha: ultimo registro de Produtos_Precos e custo do item do inventario
Private Sub cmdEditarCusto_Click()
Dim lCod As Long
Dim bSalvou As Boolean

If Grid.Row < 1 Then
   MsgBox "Selecione um produto na lista.", vbExclamation
   Exit Sub
End If

lCod = Val(Grid.TextMatrix(Grid.Row, 2))
If lCod <= 0 Then
   MsgBox "Selecione um produto na lista.", vbExclamation
   Exit Sub
End If

If mbFechado Then
   MsgBox "Inventário FECHADO: use EDITAR INVENTÁRIO antes de alterar o custo.", vbExclamation
   Exit Sub
End If

Load Produtos_AjustoPreco
If Not Produtos_AjustoPreco.PrepararSomenteCusto(lCod) Then
   Unload Produtos_AjustoPreco
   Exit Sub
End If
Produtos_AjustoPreco.Show vbModal, Me
bSalvou = Produtos_AjustoPreco.Salvou
Unload Produtos_AjustoPreco

If Not bSalvou Then Exit Sub

'o item do inventario passa a ter o custo gravado em Produtos_Precos e o total e refeito sobre a contagem
If SQLExecutaTratandoErro("UPDATE i SET i.VlrUnitInvent = c.CUSTO, i.TotalInvent = CONVERT(DECIMAL(12,2), i.SaldoLancado * c.CUSTO) " & _
   "FROM TbInventariosItens i INNER JOIN Produtos_Precos c ON (c.CODIGO = (SELECT MAX(CODIGO) FROM Produtos_Precos WHERE (COD_PRODUTO = i.IDProduto))) " & _
   "WHERE (i.IdInventario = " & mlIdInvent & ") AND (i.IDProduto = " & lCod & ")") Then Exit Sub

cmdLocalizar_Click
End Sub

Private Sub cboCriterios_GotFocus()
Dim var_Texto As String
var_Texto = cboCriterios.Text

cboCriterios.Clear
cboCriterios.AddItem "TODOS"
cboCriterios.AddItem "DESCRIÇÃO"
cboCriterios.AddItem "CÓD. BARRA"
cboCriterios.AddItem "CATEGORIA"
cboCriterios.AddItem "NCM"
moCombo.AttachTo cboCriterios

cboCriterios.Text = var_Texto
cboCriterios.SelStart = 0
cboCriterios.SelLength = Len(cboCriterios)
End Sub

Private Sub cboCriterios_Click()
AplicarCriterio True
End Sub

Private Sub cboCriterios_LostFocus()
AplicarCriterio True
End Sub

'Mostra so os objetos que o criterio escolhido usa
Private Sub AplicarCriterio(ByVal bFoco As Boolean)
Dim bDesc As Boolean

bDesc = (cboCriterios.Text = "DESCRIÇÃO")
optCompleto.Visible = bDesc
optPorIniciais.Visible = bDesc
optPorPalavra.Visible = bDesc
optPalavrasDuplas.Visible = bDesc

Select Case cboCriterios.Text
   Case "DESCRIÇÃO"
      lblNomeCombo.Caption = "Produto"
      optPorPalavra.Value = True
   Case "CÓD. BARRA"
      lblNomeCombo.Caption = "Cód. de Barra"
   Case "CATEGORIA"
      lblNomeCombo.Caption = "Categoria"
   Case "NCM"
      lblNomeCombo.Caption = "NCM"
End Select

cboConsProduto.Visible = (cboCriterios.Text = "DESCRIÇÃO" Or cboCriterios.Text = "CÓD. BARRA" Or cboCriterios.Text = "CATEGORIA" Or cboCriterios.Text = "NCM")
lblNomeCombo.Visible = cboConsProduto.Visible
If Not cboConsProduto.Visible Then cboConsProduto.Text = ""
If bFoco And cboConsProduto.Visible Then cboConsProduto.SetFocus
End Sub

Private Sub cboConsProduto_Change()
If cboCriterios.Text = "CÓD. BARRA" And Len(cboConsProduto) = 13 Then cmdLocalizar_Click
End Sub

Private Sub cboConsProduto_GotFocus()
Dim r2 As ADODB.Recordset

If cboCriterios.Text = "CÓD. BARRA" Then
   cboConsProduto.Clear

ElseIf cboCriterios.Text = "DESCRIÇÃO" Then
   cboConsProduto.Clear
   If optCompleto.Value = True Then
      Set r2 = dbData.OpenRecordset("SELECT DISTINCT NomeProduto FROM InventarioGerado WHERE (IdInventario = " & mlIdInvent & ") ORDER BY NomeProduto;")
      Do While Not r2.EOF
         cboConsProduto.AddItem ValidateNull(r2("NomeProduto"))
         r2.MoveNext
      Loop
   End If

ElseIf cboCriterios.Text = "CATEGORIA" Then
   cboConsProduto.Clear
   Set r2 = dbData.OpenRecordset("SELECT DISTINCT p.categoria FROM produtos p INNER JOIN InventarioGerado g ON (g.IDProduto = p.CODIGO) WHERE (g.IdInventario = " & mlIdInvent & ") ORDER BY p.categoria;")
   Do While Not r2.EOF
      cboConsProduto.AddItem ValidateNull(r2("categoria"))
      r2.MoveNext
   Loop
End If

SelectControl cboConsProduto
moCombo.AttachTo cboConsProduto
End Sub

Private Sub cboConsProduto_KeyPress(KeyAscii As Integer)
If cboCriterios.Text = "CÓD. BARRA" Or cboCriterios.Text = "NCM" Then
   If KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 48 Or KeyAscii > 57) Then KeyAscii = 0
ElseIf KeyAscii <> 13 Then
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
If KeyAscii = 13 Then
   KeyAscii = 0
   cmdLocalizar_Click
End If
End Sub

Private Sub cboOrdem_GotFocus()
Dim var_Texto As String
var_Texto = cboOrdem.Text

cboOrdem.Clear
cboOrdem.AddItem "SEQ"
cboOrdem.AddItem "DESCRIÇÃO"
cboOrdem.AddItem "CÓD. BARRA"
cboOrdem.AddItem "CÓDIGO"
cboOrdem.AddItem "QTDE ATUAL"
cboOrdem.AddItem "CONTAGEM"
cboOrdem.AddItem "DIFERENÇA"
cboOrdem.AddItem "CUSTO"
cboOrdem.AddItem "TOTAL"
moCombo.AttachTo cboOrdem

cboOrdem.Text = var_Texto
cboOrdem.SelStart = 0
cboOrdem.SelLength = Len(cboOrdem)
End Sub

Private Sub cboOrdem2_GotFocus()
Dim var_Texto As String
var_Texto = cboOrdem2.Text

cboOrdem2.Clear
cboOrdem2.AddItem "ASC"
cboOrdem2.AddItem "DESC"
moCombo.AttachTo cboOrdem2

cboOrdem2.Text = var_Texto
cboOrdem2.SelStart = 0
cboOrdem2.SelLength = Len(cboOrdem2)
End Sub

'Consulta os itens do inventario conforme Criterios / Filtro / Ordem
Private Sub cmdLocalizar_Click()
Dim sWhere As String
Dim sCons As String
Dim sNCM As String
Dim sPartes As String
Dim aPalavras() As String
Dim iPal As Integer
Dim iCh As Integer
Dim INDICE As String
Dim INDICE2 As String

sWhere = "(IdInventario = " & mlIdInvent & ")"
sCons = Replace(Trim(cboConsProduto.Text), "'", "''")

Select Case cboCriterios.Text
   Case "DESCRIÇÃO"
      If sCons <> "" Then
         If optCompleto.Value = True Then
            sWhere = sWhere & " AND (NomeProduto = '" & sCons & "')"
         ElseIf optPorIniciais.Value = True Then
            sWhere = sWhere & " AND (NomeProduto LIKE '" & sCons & "%')"
         ElseIf optPorPalavra.Value = True Then
            sWhere = sWhere & " AND (NomeProduto LIKE '%" & sCons & "%')"
         ElseIf optPalavrasDuplas.Value = True Then
            aPalavras = Split(sCons, " ")
            sPartes = ""
            For iPal = 0 To UBound(aPalavras)
               If Trim(aPalavras(iPal)) <> "" Then
                  If sPartes <> "" Then sPartes = sPartes & " AND "
                  sPartes = sPartes & "(NomeProduto LIKE '%" & Trim(aPalavras(iPal)) & "%')"
               End If
            Next iPal
            If sPartes <> "" Then sWhere = sWhere & " AND (" & sPartes & ")"
         End If
      End If

   Case "CÓD. BARRA"
      sCons = Replace(Replace(sCons, Chr(160), ""), " ", "")
      If sCons <> "" Then
         If Len(sCons) < 6 Then
            sWhere = sWhere & " AND ((EAN = '" & sCons & "') OR (EAN = '" & Format(sCons, "00000") & "'))"
         Else
            sWhere = sWhere & " AND (EAN = '" & sCons & "')"
         End If
      End If

   Case "CATEGORIA"
      If sCons <> "" Then sWhere = sWhere & " AND ((SELECT TOP 1 categoria FROM produtos WHERE (CODIGO = IDProduto)) = '" & sCons & "')"

   Case "NCM"
      sNCM = ""
      For iCh = 1 To Len(sCons)
         If Mid(sCons, iCh, 1) >= "0" And Mid(sCons, iCh, 1) <= "9" Then sNCM = sNCM & Mid(sCons, iCh, 1)
      Next iCh
      If sNCM <> "" Then sWhere = sWhere & " AND (REPLACE(NCM, '.', '') = '" & sNCM & "')"
End Select

Select Case cboOrdem.Text
   Case "DESCRIÇÃO":   INDICE = "NomeProduto"
   Case "CÓD. BARRA":   INDICE = "EAN"
   Case "CÓDIGO":       INDICE = "IDProduto"
   Case "QTDE ATUAL":   INDICE = "SaldoCalculado"
   Case "CONTAGEM":     INDICE = "SaldoLancado"
   Case "DIFERENÇA":    INDICE = "(SaldoLancado - SaldoCalculado)"
   Case "CUSTO":        INDICE = "VlrUnitInvent"
   Case "TOTAL":        INDICE = "TotalInvent"
   Case Else:           INDICE = "Seq"
End Select
If cboOrdem2.Text = "DESC" Then INDICE2 = "DESC" Else INDICE2 = "ASC"

sSQL = "SELECT IdInventario, Seq as vSeq, IDProduto as vCodProd, NomeProduto as vDesc, EAN as vEAN, NCM as vNCM, SaldoCalculado, VlrUnitInvent as vVlrUnit, TotalInvent as vTotal, MetaCalculado  as vMeta, SaldoLancado, TotalFisico, (SELECT TOP 1 UNID_MEDIDA FROM produtos WHERE (CODIGO = IDProduto)) as vUnid " & _
       "FROM InventarioGerado " & _
       "WHERE " & sWhere & " " & _
       "ORDER BY " & INDICE & " " & INDICE2 & IIf(INDICE = "Seq", "", ", Seq")
Set r = dbData.OpenRecordset(sSQL)

FormatarGrid r

If r.State <> 0 Then r.Close
Set r = Nothing

printSQL = sSQL
End Sub

Private Sub FormatarGrid(rTabela As ADODB.Recordset)
Dim i As Integer
Dim bKg As Boolean

With Grid
   .AllowBigSelection = False
   .SelectionMode = flexSelectionFree     'a selecao da linha toda e feita no Grid_Click (a coluna CONTAGEM seleciona so a celula)
   .HighLight = flexHighlightAlways
   .Clear
   .Cols = 12
   .rows = 2
   
   .ColWidth(0) = 0
   
   .ColWidth(1) = 525
   .ColWidth(2) = 680
   .ColWidth(3) = 1368
   .ColWidth(4) = 5216
   .ColWidth(5) = 900
   .ColWidth(6) = 1290
   .ColWidth(7) = 1173
   .ColWidth(8) = 1210
   .ColWidth(9) = 1100
   .ColWidth(10) = 1200
   .ColWidth(11) = 0       'unidade de medida (oculta): KG mostra 3 casas, demais inteiro
   
   .TextMatrix(0, 1) = "SEQ"
   .TextMatrix(0, 2) = "CÓD."
   .TextMatrix(0, 3) = "EAN"
   .TextMatrix(0, 4) = "DESCRIÇÃO"
   .TextMatrix(0, 5) = "NCM"
   .TextMatrix(0, 6) = "QTDE ATUAL"
   .TextMatrix(0, 7) = "CONTAGEM"
   .TextMatrix(0, 8) = "DIFERENÇA"
   .TextMatrix(0, 9) = "CUSTO"
   .TextMatrix(0, 10) = "TOTAL"

   .Redraw = False
   i = 1
   
   If Not rTabela Is Nothing Then
      Do While Not rTabela.EOF
         .TextMatrix(.rows - 1, 1) = rTabela("vSeq")
         .TextMatrix(.rows - 1, 2) = Format(ValidateNull(rTabela("vcodprod")), "00000")
         .TextMatrix(.rows - 1, 3) = rTabela("vEAN")
         .TextMatrix(.rows - 1, 4) = rTabela("vdesc")
         .TextMatrix(.rows - 1, 5) = rTabela("vNCM")
         bKg = (UCase(Trim(rTabela("vUnid") & "")) = "KG")
         .TextMatrix(.rows - 1, 11) = IIf(bKg, "KG", "")
         .TextMatrix(.rows - 1, 6) = FmtQtd(rTabela("SaldoCalculado"), bKg)
         .TextMatrix(.rows - 1, 7) = FmtQtd(rTabela("SaldoLancado"), bKg)
         .TextMatrix(.rows - 1, 8) = FmtQtd(rTabela("SaldoLancado") - rTabela("SaldoCalculado"), bKg)
         .TextMatrix(.rows - 1, 9) = Format(rTabela("vVlrUnit"), ocMONEY)
         .TextMatrix(.rows - 1, 10) = Format(rTabela("vtotal"), ocMONEY)
         If Not mbFechado Then
            .Row = .rows - 1
            .Col = 7
            .CellBackColor = &HC0FFFF
         End If
         If rTabela("SaldoLancado") <> rTabela("SaldoCalculado") Then
            .Row = .rows - 1
            .Col = 8
            .CellForeColor = &HC0&
         End If
         rTabela.MoveNext
         .rows = .rows + 1
         i = i + 1
      Loop
   End If

   For i = 0 To .Cols - 1
      .Col = i
      .Row = 0
      .CellFontBold = True
   Next
   
   
   'MUDAR COR DE FONTE DA COLUNA
   'For i = 1 To .Rows - 1
   '   .Row = i
   '   .Col = 9
   '   .CellForeColor = &HC0&
   '   .CellFontBold = True
   'Next
   
   .rows = .rows - 1
   .Redraw = True
End With

lblValorTotalFiscal.Caption = "TOTAL FISCAL: " & Format(SomaGrid(Grid, 10), ocMONEY)

'lblTotalEntrada.Caption = Format(SomaGrid(Grid, 7), ocMONEY)
'lblTotalSaida.Caption = Format(SomaGrid(Grid, 8), ocMONEY)
'lblTotal.Caption = Format(SomaGrid(Grid, 9), ocMONEY)
End Sub
Public Function SomaGrid(Grid As MSFlexGrid, Col As Integer) As Currency
   Dim i As Integer, Valor As Currency
   
   Valor = 0
   For i = 0 To Grid.rows - 1
      If IsNumeric(Grid.TextMatrix(i, Col)) Then
         Valor = Valor + CCur(Grid.TextMatrix(i, Col))
      End If
   Next
   
   SomaGrid = Valor
End Function

Private Sub Form_Load()
Set moCombo = New cComboHelper
End Sub

Private Function NumGrid(ByVal sTxt As String) As Double
NumGrid = Val(Replace(Replace(sTxt, ".", ""), ",", "."))
End Function

'Abre o editor na celula CONTADO da linha atual (inventario aberto)
Private Sub IniciarEdicao(Optional ByVal sInicial As String = "")
If Grid.Row < 1 Then Exit Sub
If Not IsNumeric(Grid.TextMatrix(Grid.Row, 2)) Then Exit Sub

If mbFechado Then
   ShowMsg "Inventário FECHADO: reabra na tela de inventários para alterar a contagem.", vbInformation
   Exit Sub
End If

Grid.Col = 7
Grid.ColSel = 7
Grid.RowSel = Grid.Row
mlLinhaEdit = Grid.Row
mbEditando = True
txtEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth, Grid.CellHeight
If sInicial = "" Then
   txtEdit.Text = Grid.TextMatrix(Grid.Row, 7)
Else
   txtEdit.Text = sInicial
End If
txtEdit.Visible = True
txtEdit.ZOrder 0
txtEdit.SetFocus
If sInicial = "" Then
   txtEdit.SelStart = 0
   txtEdit.SelLength = Len(txtEdit.Text)
Else
   txtEdit.SelStart = Len(txtEdit.Text)
End If
End Sub

Private Sub ConfirmarEdicao(Optional ByVal bProxima As Boolean = False)
Dim dQtd As Double
Dim dSistema As Double
Dim dUnit As Double
Dim bKg As Boolean
Dim lLinha As Long
Dim lRowAnt As Long
Dim lColAnt As Long

If Not mbEditando Then Exit Sub
mbEditando = False
lLinha = mlLinhaEdit
txtEdit.Visible = False

If Trim(txtEdit.Text) = "" Then Exit Sub
dQtd = NumGrid(txtEdit.Text)
If dQtd = NumGrid(Grid.TextMatrix(lLinha, 7)) Then GoTo fim     'nao mudou

'grava a contagem (SaldoLancado e Meta do inventario)
If SQLExecutaTratandoErro("EXEC InventarioItemContar " & mlIdInvent & ", " & Val(Grid.TextMatrix(lLinha, 2)) & ", " & Replace(CStr(dQtd), ",", ".")) Then Exit Sub

dSistema = NumGrid(Grid.TextMatrix(lLinha, 6))
dUnit = NumGrid(Grid.TextMatrix(lLinha, 9))
bKg = (Grid.TextMatrix(lLinha, 11) = "KG")
Grid.TextMatrix(lLinha, 7) = FmtQtd(dQtd, bKg)
Grid.TextMatrix(lLinha, 8) = FmtQtd(dQtd - dSistema, bKg)
Grid.TextMatrix(lLinha, 10) = Format(dQtd * dUnit, ocMONEY)

'cor da diferenca, sem tirar o usuario da celula onde ele clicou
lRowAnt = Grid.Row
lColAnt = Grid.Col
Grid.Row = lLinha
Grid.Col = 8
If dQtd <> dSistema Then Grid.CellForeColor = &HC0& Else Grid.CellForeColor = vbBlack
Grid.Row = lRowAnt
Grid.Col = lColAnt
lblValorTotalFiscal.Caption = "TOTAL FISCAL: " & Format(SomaGrid(Grid, 10), ocMONEY)

fim:
'Enter: segue para a proxima linha para continuar contando
If bProxima Then
   If lLinha < Grid.rows - 1 Then Grid.Row = lLinha + 1
   Grid.Col = 7
   Grid.SetFocus
End If
End Sub

Private Sub Grid_Click()
Dim lCol As Long

lCol = Grid.Col
If Grid.Row >= 1 And lCol <> 7 Then
   'linha toda selecionada (azul)
   Grid.Col = 1
   Grid.ColSel = Grid.Cols - 1
   Grid.RowSel = Grid.Row
Else
   'cabecalho ou coluna CONTAGEM: so a celula
   Grid.RowSel = Grid.Row
   Grid.ColSel = Grid.Col
End If

If lCol = 7 And Grid.Row >= 1 Then IniciarEdicao
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   IniciarEdicao
ElseIf KeyAscii >= 48 And KeyAscii <= 57 Then
   IniciarEdicao Chr$(KeyAscii)
End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
   Case 13
      KeyAscii = 0
      ConfirmarEdicao True
   Case 27
      KeyAscii = 0
      mbEditando = False
      txtEdit.Visible = False
      Grid.SetFocus
   Case 8, 48 To 57
      'ok
   Case 44
      'so uma virgula
      If InStr(txtEdit.Text, ",") > 0 And txtEdit.SelLength = 0 Then KeyAscii = 0
   Case Else
      KeyAscii = 0
End Select
End Sub

Private Sub txtEdit_LostFocus()
'perder o foco grava a contagem
ConfirmarEdicao
End Sub
