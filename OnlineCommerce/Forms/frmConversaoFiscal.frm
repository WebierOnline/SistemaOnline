VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmConversaoFiscal
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conversão CFOP/CST - nota de entrada para a venda"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraLista
      Caption         =   "Regras"
      Height          =   4575
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   11295
      Begin VB.CheckBox chkSoRevisar
         Caption         =   "Só as marcadas para conferir com o contador"
         Height          =   255
         Left            =   3900
         TabIndex        =   3
         Top             =   300
         Width           =   4155
      End
      Begin VB.ComboBox cboFiltro
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   270
         Width           =   2835
      End
      Begin MSFlexGridLib.MSFlexGrid lstRegras
         Height          =   3615
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   6376
         _Version        =   393216
         FixedCols       =   0
         AllowUserResizing=   1
      End
      Begin VB.Label lblLegenda
         Caption         =   "* = qualquer.  Grupo ST = CFOP do fornecedor x401, x403 ou x405.  A regra mais específica vence (código + grupo, só o grupo, só o código, padrão)."
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   4290
         Width           =   11055
      End
      Begin VB.Label lblFiltro
         AutoSize        =   -1  'True
         Caption         =   "Regime:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   330
         Width           =   555
      End
   End
   Begin VB.Frame fraEdicao
      Caption         =   "Edição"
      Height          =   2175
      Left            =   60
      TabIndex        =   6
      Top             =   4680
      Width           =   11295
      Begin VB.CommandButton cmdExcluir
         Caption         =   "&Excluir"
         Height          =   375
         Left            =   3120
         TabIndex        =   26
         Top             =   1680
         Width           =   1440
      End
      Begin VB.CommandButton cmdSalvar
         Caption         =   "&Salvar"
         Height          =   375
         Left            =   1620
         TabIndex        =   25
         Top             =   1680
         Width           =   1440
      End
      Begin VB.CommandButton cmdNovo
         Caption         =   "&Nova regra"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Width           =   1440
      End
      Begin VB.CheckBox chkRevisar
         Caption         =   "Conferir com o contador"
         Height          =   255
         Left            =   8760
         TabIndex        =   23
         Top             =   1260
         Width           =   2415
      End
      Begin VB.TextBox txtObs
         Height          =   315
         Left            =   1380
         MaxLength       =   200
         TabIndex        =   22
         Top             =   1230
         Width           =   7215
      End
      Begin VB.TextBox txtCST
         Height          =   315
         Left            =   9960
         MaxLength       =   3
         TabIndex        =   20
         Top             =   780
         Width           =   735
      End
      Begin VB.TextBox txtCFOP
         Height          =   315
         Left            =   7680
         MaxLength       =   4
         TabIndex        =   18
         Top             =   780
         Width           =   855
      End
      Begin VB.ComboBox cboGrupo
         Height          =   315
         Left            =   4500
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   780
         Width           =   1455
      End
      Begin VB.TextBox txtCodigo
         Height          =   315
         Left            =   2460
         MaxLength       =   3
         TabIndex        =   14
         Top             =   780
         Width           =   735
      End
      Begin VB.ComboBox cboTipo
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   780
         Width           =   1035
      End
      Begin VB.ComboBox cboRegime
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   300
         Width           =   3255
      End
      Begin VB.Label lblSituacao
         Height          =   255
         Left            =   4800
         TabIndex        =   9
         Top             =   360
         Width           =   6375
      End
      Begin VB.Label lblObs
         AutoSize        =   -1  'True
         Caption         =   "Observação:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1290
         Width           =   885
      End
      Begin VB.Label lblCST
         AutoSize        =   -1  'True
         Caption         =   "CST/CSOSN:"
         Height          =   195
         Left            =   8880
         TabIndex        =   19
         Top             =   840
         Width           =   945
      End
      Begin VB.Label lblCFOP
         AutoSize        =   -1  'True
         Caption         =   "Venda: CFOP"
         Height          =   195
         Left            =   6600
         TabIndex        =   17
         Top             =   840
         Width           =   945
      End
      Begin VB.Label lblGrupo
         AutoSize        =   -1  'True
         Caption         =   "Grupo CFOP:"
         Height          =   195
         Left            =   3420
         TabIndex        =   15
         Top             =   840
         Width           =   945
      End
      Begin VB.Label lblTipo
         AutoSize        =   -1  'True
         Caption         =   "No XML:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   585
      End
      Begin VB.Label lblRegime
         AutoSize        =   -1  'True
         Caption         =   "Regime:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.Frame fraTeste
      Caption         =   "Testar (o que o sistema faria com um item da nota de entrada)"
      Height          =   1215
      Left            =   60
      TabIndex        =   27
      Top             =   6900
      Width           =   11295
      Begin VB.CommandButton cmdTestar
         Caption         =   "&Testar"
         Height          =   375
         Left            =   7800
         TabIndex        =   34
         Top             =   300
         Width           =   1440
      End
      Begin VB.TextBox txtTesteCFOP
         Height          =   315
         Left            =   6720
         MaxLength       =   4
         TabIndex        =   33
         Top             =   330
         Width           =   855
      End
      Begin VB.TextBox txtTesteCST
         Height          =   315
         Left            =   4920
         MaxLength       =   4
         TabIndex        =   31
         Top             =   330
         Width           =   735
      End
      Begin VB.ComboBox cboTesteRegime
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   330
         Width           =   2715
      End
      Begin VB.Label lblTeste
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   780
         Width           =   11055
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTesteCFOP
         AutoSize        =   -1  'True
         Caption         =   "CFOP:"
         Height          =   195
         Left            =   6180
         TabIndex        =   32
         Top             =   390
         Width           =   435
      End
      Begin VB.Label lblTesteCST
         AutoSize        =   -1  'True
         Caption         =   "CST/CSOSN do XML:"
         Height          =   195
         Left            =   3420
         TabIndex        =   30
         Top             =   390
         Width           =   1485
      End
      Begin VB.Label lblTesteRegime
         AutoSize        =   -1  'True
         Caption         =   "Empresa:"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   390
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdFechar
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   9900
      TabIndex        =   36
      Top             =   8190
      Width           =   1440
   End
End
Attribute VB_Name = "frmConversaoFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Regras de conversão do CFOP/CST da nota de entrada (XML do fornecedor) para o CFOP e o CST/CSOSN de venda do produto
' (tabela TbConversaoFiscal, script 172; função dbo.ConversaoFiscalVenda). SN = Simples Nacional e MEI (venda com
' CSOSN); RN = regime normal, Lucro Presumido/Real (venda com CST de 2 dígitos; a origem vem do XML).
' Toda regra criada ou alterada aqui fica Personalizada = 1: os scripts de atualização das regras padrão não mexem nela.
Private Const LISTA_CST As String = ",00,02,10,15,20,30,40,41,50,51,53,60,61,70,90,"
Private Const LISTA_CSOSN As String = ",101,102,103,201,202,203,300,400,500,900,"

Dim iId As Long
Dim nCRT As Integer
Dim bCarregando As Boolean

Private Sub Form_Load()
    On Error GoTo ErrL
    bCarregando = True
    nCRT = Val(SQLExecutaRetorno("SELECT TOP 1 ISNULL(CRT, 1) AS r FROM empresa", "r", 1))

    cboFiltro.AddItem "Todos os regimes"
    cboFiltro.AddItem "SN - Simples Nacional"
    cboFiltro.AddItem "RN - Regime normal"
    cboFiltro.ListIndex = IIf(EmpresaSimples, 1, 2)

    cboRegime.AddItem "SN - Simples Nacional (venda com CSOSN)"
    cboRegime.AddItem "RN - Regime normal (venda com CST)"
    cboTipo.AddItem "CST"
    cboTipo.AddItem "CSOSN"
    cboTipo.AddItem "*"
    cboGrupo.AddItem "*"
    cboGrupo.AddItem "ST"
    cboGrupo.AddItem "NORMAL"

    cboTesteRegime.AddItem "Simples Nacional (SN)"
    cboTesteRegime.AddItem "Regime normal (RN)"
    cboTesteRegime.ListIndex = IIf(EmpresaSimples, 0, 1)

    With lstRegras
        .Cols = 10
        .FixedCols = 0
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColWidth(0) = 700: .TextMatrix(0, 0) = "Regime"
        .ColWidth(1) = 700: .TextMatrix(0, 1) = "No XML"
        .ColWidth(2) = 700: .TextMatrix(0, 2) = "Código"
        .ColWidth(3) = 950: .TextMatrix(0, 3) = "Grupo CFOP"
        .ColWidth(4) = 800: .TextMatrix(0, 4) = "CFOP"
        .ColWidth(5) = 750: .TextMatrix(0, 5) = "CST"
        .ColWidth(6) = 850: .TextMatrix(0, 6) = "Conferir"
        .ColWidth(7) = 1150: .TextMatrix(0, 7) = "Personalizada"
        .ColWidth(8) = 4300: .TextMatrix(0, 8) = "Observação"
        .ColWidth(9) = 0     'Id
    End With
    bCarregando = False
    Carregar
    Novo
    Exit Sub
ErrL:
    bCarregando = False
    MsgBox "Não foi possível abrir as regras de conversão." & vbCrLf & vbCrLf & _
           "Verifique se o script CreateTable_TbConversaoFiscal (172) já foi aplicado no banco." & vbCrLf & _
           Err.Description, vbExclamation
End Sub

Private Function EmpresaSimples() As Boolean
    EmpresaSimples = (nCRT = 1 Or nCRT = 4)
End Function

Private Sub Carregar()
    Dim rRs As ADODB.Recordset
    Dim r As Long, c As Integer
    Dim sWhere As String

    sWhere = " WHERE Ativo = 1"
    If cboFiltro.ListIndex = 1 Then sWhere = sWhere & " AND Regime = 'SN'"
    If cboFiltro.ListIndex = 2 Then sWhere = sWhere & " AND Regime = 'RN'"
    If chkSoRevisar.Value = vbChecked Then sWhere = sWhere & " AND Revisar = 1"

    lstRegras.Rows = 1
    RsOpen rRs, "SELECT Id, Regime, TipoCodigo, CodigoEntrada, GrupoCFOP, CFOPVenda, CSTVenda, Revisar, Personalizada, ISNULL(Observacao, '') AS Observacao " & _
                "FROM TbConversaoFiscal" & sWhere & " ORDER BY Regime DESC, CASE TipoCodigo WHEN 'CST' THEN 1 WHEN 'CSOSN' THEN 2 ELSE 3 END, " & _
                "CodigoEntrada, CASE GrupoCFOP WHEN '*' THEN 2 ELSE 1 END"
    Do While Not rRs.EOF
        lstRegras.AddItem ""
        r = lstRegras.Rows - 1
        lstRegras.TextMatrix(r, 0) = rRs("Regime") & ""
        lstRegras.TextMatrix(r, 1) = rRs("TipoCodigo") & ""
        lstRegras.TextMatrix(r, 2) = rRs("CodigoEntrada") & ""
        lstRegras.TextMatrix(r, 3) = rRs("GrupoCFOP") & ""
        lstRegras.TextMatrix(r, 4) = rRs("CFOPVenda") & ""
        lstRegras.TextMatrix(r, 5) = rRs("CSTVenda") & ""
        lstRegras.TextMatrix(r, 6) = IIf(CBool(rRs("Revisar")), "Conferir", "")
        lstRegras.TextMatrix(r, 7) = IIf(CBool(rRs("Personalizada")), "Sim", "")
        lstRegras.TextMatrix(r, 8) = rRs("Observacao") & ""
        lstRegras.TextMatrix(r, 9) = rRs("Id")
        If CBool(rRs("Revisar")) Then
            lstRegras.Row = r
            For c = 0 To 8
                lstRegras.Col = c
                lstRegras.CellBackColor = RGB(255, 245, 200)
            Next
        End If
        rRs.MoveNext
    Loop
    If rRs.State <> 0 Then rRs.Close
End Sub

Private Sub cboFiltro_Click()
    If Not bCarregando Then Carregar
End Sub

Private Sub chkSoRevisar_Click()
    If Not bCarregando Then Carregar
End Sub

Private Sub lstRegras_Click()
    Dim r As Long
    r = lstRegras.Row
    If r < 1 Or lstRegras.Rows < 2 Then Exit Sub
    iId = Val(lstRegras.TextMatrix(r, 9))
    cboRegime.ListIndex = IIf(lstRegras.TextMatrix(r, 0) = "SN", 0, 1)
    SelecionarCombo cboTipo, lstRegras.TextMatrix(r, 1)
    txtCodigo.Text = lstRegras.TextMatrix(r, 2)
    SelecionarCombo cboGrupo, lstRegras.TextMatrix(r, 3)
    txtCFOP.Text = lstRegras.TextMatrix(r, 4)
    txtCST.Text = lstRegras.TextMatrix(r, 5)
    chkRevisar.Value = IIf(lstRegras.TextMatrix(r, 6) <> "", vbChecked, vbUnchecked)
    txtObs.Text = lstRegras.TextMatrix(r, 8)
    lblSituacao.Caption = "Editando a regra nº " & iId & IIf(lstRegras.TextMatrix(r, 7) <> "", " (personalizada)", " (padrão do sistema)")
    cmdExcluir.Enabled = True
End Sub

Private Sub SelecionarCombo(cbo As ComboBox, ByVal sTexto As String)
    Dim i As Integer
    For i = 0 To cbo.ListCount - 1
        If cbo.List(i) = sTexto Then cbo.ListIndex = i: Exit Sub
    Next
End Sub

Private Sub Novo()
    iId = 0
    cboRegime.ListIndex = IIf(EmpresaSimples, 0, 1)
    cboTipo.ListIndex = 0
    cboGrupo.ListIndex = 0
    txtCodigo.Text = ""
    txtCFOP.Text = ""
    txtCST.Text = ""
    txtObs.Text = ""
    chkRevisar.Value = vbUnchecked
    lblSituacao.Caption = "Nova regra"
    cmdExcluir.Enabled = False
End Sub

Private Sub cmdNovo_Click()
    Novo
    txtCodigo.SetFocus
End Sub

Private Sub cboTipo_Click()
    'tipo "*" = qualquer código
    If cboTipo.Text = "*" Then txtCodigo.Text = "*"
End Sub

Private Sub cmdSalvar_Click()
    Dim sRegime As String, sTipo As String, sCodigo As String, sGrupo As String, sCST As String, sObs As String
    Dim nCFOP As Long
    Dim sSQL As String, sErro As String

    sRegime = IIf(cboRegime.ListIndex = 0, "SN", "RN")
    sTipo = cboTipo.Text
    sGrupo = cboGrupo.Text
    sCodigo = Trim(txtCodigo.Text)
    sCST = Trim(txtCST.Text)
    sObs = Replace(Trim(txtObs.Text), "'", "''")
    nCFOP = Val(txtCFOP.Text)

    If sTipo = "*" Then sCodigo = "*"
    If sCodigo = "" Then sCodigo = "*"
    If sCodigo <> "*" Then
        If sTipo = "CST" And InStr(LISTA_CST, "," & sCodigo & ",") = 0 Then
            MsgBox "Código do XML inválido: o CST tem 2 dígitos (00, 10, 20, 30, 40, 41, 50, 51, 60, 70, 90...)." & vbCrLf & "Use * para qualquer CST.", vbExclamation
            txtCodigo.SetFocus
            Exit Sub
        End If
        If sTipo = "CSOSN" And InStr(LISTA_CSOSN, "," & sCodigo & ",") = 0 Then
            MsgBox "Código do XML inválido: o CSOSN tem 3 dígitos (101, 102, 103, 201, 202, 203, 300, 400, 500, 900)." & vbCrLf & "Use * para qualquer CSOSN.", vbExclamation
            txtCodigo.SetFocus
            Exit Sub
        End If
    End If
    If Len(txtCFOP.Text) <> 4 Or nCFOP < 5000 Or nCFOP > 7999 Then
        MsgBox "CFOP de venda inválido: 4 dígitos começando com 5 (dentro do estado), 6 ou 7.", vbExclamation
        txtCFOP.SetFocus
        Exit Sub
    End If
    If sRegime = "SN" And InStr(LISTA_CSOSN, "," & sCST & ",") = 0 Then
        MsgBox "No Simples Nacional a venda usa CSOSN (3 dígitos: 101, 102, 103, 201, 202, 203, 300, 400, 500, 900).", vbExclamation
        txtCST.SetFocus
        Exit Sub
    End If
    If sRegime = "RN" And InStr(LISTA_CST, "," & sCST & ",") = 0 Then
        MsgBox "No regime normal a venda usa CST de 2 dígitos (00, 10, 20, 40, 41, 60...)." & vbCrLf & "A origem da mercadoria vem do XML.", vbExclamation
        txtCST.SetFocus
        Exit Sub
    End If

    If iId = 0 Then
        sSQL = "INSERT INTO TbConversaoFiscal (Regime, TipoCodigo, CodigoEntrada, GrupoCFOP, CFOPVenda, CSTVenda, Observacao, Revisar, Personalizada) VALUES ('" & _
               sRegime & "', '" & sTipo & "', '" & sCodigo & "', '" & sGrupo & "', " & nCFOP & ", '" & sCST & "', '" & sObs & "', " & chkRevisar.Value & ", 1)"
    Else
        sSQL = "UPDATE TbConversaoFiscal SET Regime = '" & sRegime & "', TipoCodigo = '" & sTipo & "', CodigoEntrada = '" & sCodigo & "', GrupoCFOP = '" & sGrupo & _
               "', CFOPVenda = " & nCFOP & ", CSTVenda = '" & sCST & "', Observacao = '" & sObs & "', Revisar = " & chkRevisar.Value & ", Personalizada = 1 WHERE Id = " & iId
    End If
    sErro = SQLExecuta(sSQL)
    If sErro <> "" Then
        If InStr(1, sErro, "UX_TbConversaoFiscal", vbTextCompare) > 0 Then
            MsgBox "Já existe uma regra para esse regime, código do XML e grupo de CFOP." & vbCrLf & "Selecione a regra na lista para alterar.", vbExclamation
        Else
            MsgBox "Não foi possível gravar a regra: " & sErro, vbExclamation
        End If
        Exit Sub
    End If
    Carregar
    Novo
End Sub

Private Sub cmdExcluir_Click()
    Dim sErro As String
    If iId = 0 Then Exit Sub
    If SQLExecutaRetorno("SELECT COUNT(*) AS r FROM TbConversaoFiscal WHERE Id = " & iId & " AND TipoCodigo = '*' AND CodigoEntrada = '*' AND GrupoCFOP = '*'", "r", 0) > 0 Then
        MsgBox "Esta é a regra padrão do regime (usada quando nenhuma outra serve) e não pode ser excluída." & vbCrLf & "Altere o CFOP/CST dela, se precisar.", vbExclamation
        Exit Sub
    End If
    If MsgBox("Excluir a regra selecionada?" & vbCrLf & "Sem ela, vale a regra mais geral (só o grupo de CFOP ou a padrão do regime).", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    sErro = SQLExecuta("DELETE FROM TbConversaoFiscal WHERE Id = " & iId)
    If sErro <> "" Then
        MsgBox "Não foi possível excluir: " & sErro, vbExclamation
        Exit Sub
    End If
    Carregar
    Novo
End Sub

Private Sub cmdTestar_Click()
    Dim rRs As ADODB.Recordset
    Dim sCST As String

    sCST = Replace(Trim(txtTesteCST.Text), "'", "")
    If sCST = "" And Trim(txtTesteCFOP.Text) = "" Then
        MsgBox "Informe o CST ou CSOSN e o CFOP do item da nota do fornecedor (ex.: 060 e 6403, ou 0500 e 5405).", vbInformation
        txtTesteCST.SetFocus
        Exit Sub
    End If
    RsOpen rRs, "SELECT f.CFOPVenda, f.CSTVenda, f.Revisar, ISNULL(f.Observacao, '') AS Observacao, c.Regime, c.TipoCodigo, c.CodigoEntrada, c.GrupoCFOP " & _
                "FROM dbo.ConversaoFiscalVenda(" & IIf(cboTesteRegime.ListIndex = 0, 1, 3) & ", '" & sCST & "', " & Val(txtTesteCFOP.Text) & ") f " & _
                "JOIN TbConversaoFiscal c ON c.Id = f.Id"
    If rRs.EOF Then
        lblTeste.Caption = "Nenhuma regra serve para esse item (falta a regra padrão do regime)."
    Else
        lblTeste.Caption = "Venda: CFOP " & rRs("CFOPVenda") & " / " & IIf(rRs("Regime") = "SN", "CSOSN ", "CST ") & rRs("CSTVenda") & _
                           "   - regra " & rRs("Regime") & " / " & rRs("TipoCodigo") & " " & rRs("CodigoEntrada") & " / grupo " & rRs("GrupoCFOP") & _
                           ": " & rRs("Observacao") & IIf(CBool(rRs("Revisar")), "  (CONFERIR COM O CONTADOR)", "")
    End If
    If rRs.State <> 0 Then rRs.Close
End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub
