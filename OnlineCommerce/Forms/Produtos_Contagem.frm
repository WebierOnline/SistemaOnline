VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Produtos_Contagem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONTAGEM FÍSICA DE PRODUTOS"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   10920
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   10845
      TabIndex        =   14
      Top             =   0
      Width           =   10875
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CONTAGEM DE PRODUTOS"
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
         Left            =   1140
         TabIndex        =   15
         Top             =   240
         Width           =   4170
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   300
         Picture         =   "Produtos_Contagem.frx":0000
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.TextBox txtBusca 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   60
      TabIndex        =   0
      Top             =   1110
      Width           =   10800
   End
   Begin VB.ListBox lstOpcoes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   60
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   10800
   End
   Begin VB.TextBox txtQtd 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   60
      TabIndex        =   2
      Top             =   3030
      Width           =   2400
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   2580
      Left            =   60
      TabIndex        =   4
      Top             =   4140
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   4551
      _Version        =   393216
      Cols            =   6
      AllowBigSelection=   0   'False
      HighLight       =   0
      ScrollBars      =   2
      Appearance      =   0
   End
   Begin ChamaleonBtn.chameleonButton cmdSalva 
      Height          =   435
      Left            =   3180
      TabIndex        =   17
      ToolTipText     =   "Grava a contagem (Enter)"
      Top             =   3030
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   767
      BTYPE           =   3
      TX              =   "Salva"
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
      MICON           =   "Produtos_Contagem.frx":6370
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdCancela 
      Height          =   435
      Left            =   4500
      TabIndex        =   18
      ToolTipText     =   "Cancela (Esc)"
      Top             =   3030
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   767
      BTYPE           =   3
      TX              =   "Cancela"
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
      MICON           =   "Produtos_Contagem.frx":638C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdEditarProdutos 
      Height          =   435
      Left            =   5820
      TabIndex        =   3
      ToolTipText     =   "Abre o cadastro do produto"
      Top             =   3030
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   767
      BTYPE           =   3
      TX              =   "Editar Produto"
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
      MICON           =   "Produtos_Contagem.frx":63A8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdSair 
      Height          =   375
      Left            =   9300
      TabIndex        =   5
      ToolTipText     =   "Fecha a contagem"
      Top             =   6780
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Fechar"
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
      MICON           =   "Produtos_Contagem.frx":63C4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produto selecionado para contagem:"
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
      Left            =   60
      TabIndex        =   16
      Top             =   1680
      Width           =   3030
   End
   Begin VB.Label lblProgresso 
      AutoSize        =   -1  'True
      Caption         =   "CONTADOS: 0 DE 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   60
      TabIndex        =   6
      Top             =   6780
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Digite: CÓD. BARRAS ou CÓDIGO ou DESCRIÇÃO e pressione ENTER"
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
      Left            =   60
      TabIndex        =   7
      Top             =   840
      Width           =   5400
   End
   Begin VB.Label lblDescricao 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   360
      Left            =   60
      TabIndex        =   8
      Top             =   1920
      Width           =   10800
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   9
      Top             =   2340
      Width           =   10740
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Qtde. Contada (real):"
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
      Left            =   60
      TabIndex        =   10
      Top             =   2820
      Width           =   1800
   End
   Begin VB.Label lblUnid 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2520
      TabIndex        =   11
      Top             =   3010
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Enter grava o ajuste e volta ao código de barras  |  Esc cancela  |"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   60
      TabIndex        =   12
      Top             =   3480
      Width           =   4620
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "ÚLTIMOS PRODUTOS CONTADOS:"
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
      Left            =   60
      TabIndex        =   13
      Top             =   3900
      Width           =   3060
   End
End
Attribute VB_Name = "Produtos_Contagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Contagem fisica: o usuario conta o produto na prateleira e informa a quantidade que esta la. O sistema le o estoque atual,
'grava o ajuste (ADICAO ou REMOCAO) em Produtos_Quant com FORMA = 'CONTAGEM' (tinha x, ficou y), altera o estoque fisico
'(Produtos.QUANT_ESTOQUE) e marca o produto como contado. Tudo feito pela procedure ProdutoContagemGravar, numa transacao.
Option Explicit

Dim mlProduto As Long           'produto localizado, aguardando a quantidade
Dim msDescricao As String
Dim mbKg As Boolean
Dim mdAtual As Double           'estoque fisico atual no momento em que o produto foi localizado (sempre exibido)
Dim mbJaContado As Boolean
Dim mdtContagem As Date
Dim mbIniciado As Boolean

Private Sub Form_Load()
With Grid
   .Cols = 6
   .rows = 2                'cabecalho + 1 linha em branco (o grid exige pelo menos 1 linha alem das fixas)
   .FixedRows = 1
   .FixedCols = 0
   .ColWidth(0) = 900
   .ColWidth(1) = 800
   .ColWidth(2) = 4800
   .ColWidth(3) = 1300
   .ColWidth(4) = 1300
   .ColWidth(5) = 1300
   .TextMatrix(0, 0) = "HORA"
   .TextMatrix(0, 1) = "CÓD."
   .TextMatrix(0, 2) = "DESCRIÇÃO"
   .TextMatrix(0, 3) = "ANTES"
   .TextMatrix(0, 4) = "CONTADO"
   .TextMatrix(0, 5) = "AJUSTE"
   .ColAlignment(3) = flexAlignRightCenter
   .ColAlignment(4) = flexAlignRightCenter
   .ColAlignment(5) = flexAlignRightCenter
End With
AtualizarProgresso
AtualizarBotoes
End Sub

Private Sub Form_Activate()
If Not mbIniciado Then
   mbIniciado = True
   txtBusca.SetFocus
End If
End Sub

'Salva, Cancela e Editar Produto so fazem sentido com um produto em andamento
Private Sub AtualizarBotoes()
cmdSalva.Enabled = (mlProduto <> 0)
cmdCancela.Enabled = (mlProduto <> 0)
cmdEditarProdutos.Enabled = (mlProduto <> 0)
End Sub

Private Function NumTxt(ByVal sTxt As String) As Double
NumTxt = Val(Replace(Replace(sTxt, ".", ""), ",", "."))
End Function

Private Function FmtQtd(ByVal dQtd As Double, ByVal bKg As Boolean) As String
If bKg Then
   FmtQtd = Format(dQtd, ocPESO)
ElseIf dQtd = Int(dQtd) Then
   FmtQtd = Format(dQtd, "#,##0")
Else
   FmtQtd = Format(dQtd, ocPESO)
End If
End Function

Private Sub AtualizarProgresso()
Dim rP As ADODB.Recordset

Set rP = dbData.OpenRecordset("SELECT COUNT(*) AS total, SUM(CASE WHEN ISNULL(CONTADO, 0) = 1 THEN 1 ELSE 0 END) AS contados FROM Produtos WHERE (ATIVO = 1)")
If Not rP.EOF Then lblProgresso.Caption = "Produtos Contados: " & ValidateNull(rP("contados")) & " de " & ValidateNull(rP("total"))
If rP.State <> 0 Then rP.Close
Set rP = Nothing
End Sub

'Limpa o produto em andamento e volta para o campo de busca
Private Sub NovaBusca()
mlProduto = 0
mbJaContado = False
lblDescricao.Caption = ""
lblStatus.Caption = ""
lblUnid.Caption = ""
txtQtd.Text = ""
txtQtd.Enabled = False
lstOpcoes.Visible = False
lstOpcoes.Clear
txtBusca.Text = ""
AtualizarBotoes
txtBusca.SetFocus
End Sub

'Procura entre os produtos ativos por codigo de barras / EAN / codigo (so digitos) ou por palavras da descricao
Private Sub Localizar()
Dim sTxt As String
Dim sCons As String
Dim sWhere As String
Dim sPartes As String
Dim aPal() As String
Dim iPal As Integer
Dim iCh As Integer
Dim bNumero As Boolean
Dim rL As ADODB.Recordset

sTxt = Trim(txtBusca.Text)
If sTxt = "" Then Exit Sub

bNumero = True
For iCh = 1 To Len(sTxt)
   If Mid(sTxt, iCh, 1) < "0" Or Mid(sTxt, iCh, 1) > "9" Then bNumero = False
Next iCh

sCons = Replace(sTxt, "'", "''")
If bNumero Then
   sWhere = "(p.COD_BARRA = '" & sCons & "') OR (p.EAN = '" & sCons & "')"
   If Len(sCons) < 6 Then sWhere = sWhere & " OR (p.COD_BARRA = '" & Format(sCons, "00000") & "') OR (p.EAN = '" & Format(sCons, "00000") & "')"
   If Len(sCons) <= 6 Then sWhere = sWhere & " OR (p.CODIGO = " & Val(sCons) & ")"
   sWhere = "(" & sWhere & ")"
Else
   aPal = Split(sCons, " ")
   sPartes = ""
   For iPal = 0 To UBound(aPal)
      If Trim(aPal(iPal)) <> "" Then
         If sPartes <> "" Then sPartes = sPartes & " AND "
         sPartes = sPartes & "(p.DESCRICAO LIKE '%" & Trim(aPal(iPal)) & "%')"
      End If
   Next iPal
   If sPartes = "" Then Exit Sub
   sWhere = "(" & sPartes & ")"
End If

Set rL = dbData.OpenRecordset("SELECT TOP 50 p.CODIGO, p.DESCRICAO, ISNULL(p.COD_BARRA, '') AS COD_BARRA FROM Produtos p WHERE (p.ATIVO = 1) AND " & sWhere & " ORDER BY p.DESCRICAO;")

lstOpcoes.Clear
Do While Not rL.EOF
   lstOpcoes.AddItem Format(ValidateNull(rL("CODIGO")), "00000") & "  -  " & rL("DESCRICAO") & "  -  " & rL("COD_BARRA")
   lstOpcoes.ItemData(lstOpcoes.NewIndex) = rL("CODIGO")
   rL.MoveNext
Loop
If rL.State <> 0 Then rL.Close
Set rL = Nothing

Select Case lstOpcoes.ListCount
   Case 0
      Beep
      lblDescricao.Caption = ""
      lblStatus.ForeColor = &HC0&
      lblStatus.Caption = "Nenhum produto ativo encontrado."
      txtBusca.SelStart = 0
      txtBusca.SelLength = Len(txtBusca.Text)
   Case 1
      SelecionarProduto lstOpcoes.ItemData(0)
   Case Else
      lblStatus.ForeColor = vbBlack
      lblStatus.Caption = lstOpcoes.ListCount & " produtos encontrados: escolha com as setas e Enter (Esc cancela)."
      lstOpcoes.Visible = True
      lstOpcoes.ListIndex = 0
      lstOpcoes.SetFocus
End Select
End Sub

Private Sub SelecionarProduto(ByVal lProd As Long)
Dim rs As ADODB.Recordset
Dim sEan As String

lstOpcoes.Visible = False

Set rs = dbData.OpenRecordset("SELECT CODIGO, DESCRICAO, ISNULL(NULLIF(COD_BARRA, ''), ISNULL(EAN, '')) AS vEan, ISNULL(UNID_MEDIDA, '') AS vUnid, ISNULL(QUANT_ESTOQUE, 0) AS vAtual, ISNULL(CONTADO, 0) AS vContado, DATA_CONTAGEM FROM Produtos WHERE (ATIVO = 1) AND (CODIGO = " & lProd & ");")
If rs.EOF Then
   If rs.State <> 0 Then rs.Close
   Set rs = Nothing
   NovaBusca
   Exit Sub
End If

mlProduto = lProd
msDescricao = rs("DESCRICAO") & ""
sEan = Trim(rs("vEan") & "")
If sEan = "" Then sEan = Format(lProd, "00000")   'sem EAN/codigo de barras: mostra o codigo do produto
mbKg = (UCase(Trim(rs("vUnid") & "")) = "KG")
mdAtual = CDbl(ValidateNull(rs("vAtual")))
mbJaContado = (CLng(ValidateNull(rs("vContado"))) <> 0) And Not IsNull(rs("DATA_CONTAGEM"))
If mbJaContado Then mdtContagem = rs("DATA_CONTAGEM")
lblUnid.Caption = Trim(rs("vUnid") & "")
If rs.State <> 0 Then rs.Close
Set rs = Nothing

lblDescricao.Caption = sEan & " - " & msDescricao
MostrarStatus
AtualizarBotoes

txtQtd.Enabled = True
txtQtd.Text = ""
txtQtd.SetFocus
End Sub

'Linha de situacao do produto: ja contado / nao contado, e sempre o estoque fisico atual
Private Sub MostrarStatus()
Dim sTxt As String

If mlProduto = 0 Then Exit Sub

If mbJaContado Then
   lblStatus.ForeColor = &HC0&
   sTxt = "PRODUTO JÁ CONTADO em " & Format(mdtContagem, "dd/mm/yyyy hh:nn")
Else
   lblStatus.ForeColor = &H800000
   sTxt = "AINDA NÃO CONTADO"
End If
sTxt = sTxt & "     |     ESTOQUE ATUAL: " & FmtQtd(mdAtual, mbKg)
lblStatus.Caption = sTxt
End Sub

Private Sub Gravar()
Dim dQtd As Double
Dim dAnt As Double
Dim rG As ADODB.Recordset
Dim sErro As String
Dim sResumo As String

If mlProduto = 0 Then Exit Sub
If Trim(txtQtd.Text) = "" Then Exit Sub

dQtd = NumTxt(txtQtd.Text)

If mbJaContado Then
   If ShowMsg("Este produto já foi contado em " & Format(mdtContagem, "dd/mm/yyyy hh:nn") & "." & vbCrLf & "Contar de novo e ajustar para " & FmtQtd(dQtd, mbKg) & "?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
End If

On Error GoTo ErrGravar
Set rG = dbData.OpenRecordset("EXEC ProdutoContagemGravar " & mlProduto & ", " & Replace(CStr(dQtd), ",", ".") & ", " & vCodFunc)
If rG.EOF Then GoTo ErrGravar
dAnt = CDbl(ValidateNull(rG("QuantAnterior")))
If rG.State <> 0 Then rG.Close
Set rG = Nothing
On Error GoTo 0

AdicionarHistorico dAnt, dQtd

If dQtd > dAnt Then
   sResumo = "ADIÇÃO de " & FmtQtd(dQtd - dAnt, mbKg)
ElseIf dQtd < dAnt Then
   sResumo = "REMOÇÃO de " & FmtQtd(dAnt - dQtd, mbKg)
Else
   sResumo = "sem diferença"
End If

AtualizarProgresso
NovaBusca
lblStatus.ForeColor = &H8000&
lblStatus.Caption = "GRAVADO: tinha " & FmtQtd(dAnt, mbKg) & ", ficou " & FmtQtd(dQtd, mbKg) & " (" & sResumo & ")"
Exit Sub

ErrGravar:
sErro = Err.Description
On Error Resume Next
If Not rG Is Nothing Then
   If rG.State <> 0 Then rG.Close
End If
Set rG = Nothing
On Error GoTo 0
If sErro = "" Then sErro = "A procedure não devolveu o resultado da contagem."
MsgBox "Não foi possível gravar a contagem: " & sErro, vbCritical, "Contagem física"
End Sub

Private Sub AdicionarHistorico(ByVal dAnt As Double, ByVal dQtd As Double)
Dim sAjuste As String
Dim sLinha As String
Dim aCol() As String
Dim iCol As Integer

If dQtd > dAnt Then
   sAjuste = "+" & FmtQtd(dQtd - dAnt, mbKg)
ElseIf dQtd < dAnt Then
   sAjuste = "-" & FmtQtd(dAnt - dQtd, mbKg)
Else
   sAjuste = "0"
End If

sLinha = Format(Now, "hh:nn:ss") & vbTab & Format(mlProduto, "00000") & vbTab & msDescricao & vbTab & FmtQtd(dAnt, mbKg) & vbTab & FmtQtd(dQtd, mbKg) & vbTab & sAjuste

With Grid
   If .rows = 2 And .TextMatrix(1, 1) = "" Then
      'primeira contagem: usa a linha em branco
      aCol = Split(sLinha, vbTab)
      For iCol = 0 To UBound(aCol)
         .TextMatrix(1, iCol) = aCol(iCol)
      Next iCol
   Else
      .AddItem sLinha, 1
      If .rows > 201 Then .RemoveItem .rows - 1
   End If
End With
End Sub

Private Sub txtBusca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   Localizar
End If
End Sub

Private Sub txtQtd_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44      'ponto do teclado numerico vira virgula

Select Case KeyAscii
   Case 13
      KeyAscii = 0
      Gravar
   Case 27
      KeyAscii = 0
      NovaBusca
   Case 8, 48 To 57
      'ok
   Case 44
      'so uma virgula
      If InStr(txtQtd.Text, ",") > 0 And txtQtd.SelLength = 0 Then KeyAscii = 0
   Case Else
      KeyAscii = 0
End Select
End Sub

Private Sub lstOpcoes_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
   Case 13
      KeyAscii = 0
      If lstOpcoes.ListIndex >= 0 Then SelecionarProduto lstOpcoes.ItemData(lstOpcoes.ListIndex)
   Case 27
      KeyAscii = 0
      NovaBusca
End Select
End Sub

Private Sub lstOpcoes_DblClick()
If lstOpcoes.ListIndex >= 0 Then SelecionarProduto lstOpcoes.ItemData(lstOpcoes.ListIndex)
End Sub

'Selecionar o texto ao entrar: digitar uma nova busca substitui a anterior
Private Sub txtBusca_GotFocus()
txtBusca.SelStart = 0
txtBusca.SelLength = Len(txtBusca.Text)
End Sub

'Mesma funcao do Enter no campo da quantidade
Private Sub cmdSalva_Click()
Gravar
End Sub

'Mesma funcao do Esc no campo da quantidade
Private Sub cmdCancela_Click()
NovaBusca
End Sub

'Abre o cadastro do produto em andamento; ao fechar recarrega os dados dele (descricao, unidade, estoque...)
'mantendo a quantidade que ja tinha sido digitada
Private Sub cmdEditarProdutos_Click()
Dim lProd As Long
Dim sQtd As String

If mlProduto = 0 Then Exit Sub
lProd = mlProduto
sQtd = txtQtd.Text

Load Produtos_Cadastro
Produtos_Cadastro.SSTab1.Tab = 0
Produtos_Cadastro.EditarProduto lProd
Produtos_Cadastro.Show vbModal, Me
Unload Produtos_Cadastro

SelecionarProduto lProd
If mlProduto <> 0 Then
   txtQtd.Text = sQtd
   txtQtd.SelStart = Len(txtQtd.Text)
End If
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub
