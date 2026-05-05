VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmNCM_Cadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de NCM"
   ClientHeight    =   7695
   ClientLeft      =   -30
   ClientTop       =   300
   ClientWidth     =   14700
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   14700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   495
      Left            =   13320
      TabIndex        =   32
      Top             =   7140
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4440
      TabIndex        =   31
      Top             =   7140
      Width           =   1335
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3000
      TabIndex        =   30
      Top             =   7140
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1560
      TabIndex        =   29
      Top             =   7140
      Width           =   1335
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "Novo"
      Height          =   495
      Left            =   120
      TabIndex        =   28
      Top             =   7140
      Width           =   1335
   End
   Begin VB.Frame fraFiltro 
      Caption         =   "Consulta"
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   14580
      Begin VB.ComboBox cboCriterio 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   210
         Width           =   1935
      End
      Begin VB.TextBox txtBusca 
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Top             =   210
         Width           =   9495
      End
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "Consultar"
         Height          =   315
         Left            =   11700
         TabIndex        =   2
         Top             =   210
         Width           =   1335
      End
      Begin VB.CommandButton cmdLimpar 
         Caption         =   "Todos"
         Height          =   315
         Left            =   13140
         TabIndex        =   3
         Top             =   210
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdNCM 
      Height          =   4215
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   14520
      _ExtentX        =   25612
      _ExtentY        =   7435
      _Version        =   393216
      Rows            =   1
      Cols            =   9
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin VB.Frame fraCarregando 
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Left            =   5160
      TabIndex        =   14
      Top             =   2550
      Visible         =   0   'False
      Width           =   5280
      Begin VB.Label lblCarregando 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "CARREGANDO..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   600
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   5040
      End
   End
   Begin VB.Frame fraEdicao 
      Caption         =   "Dados do NCM"
      Height          =   1635
      Left            =   120
      TabIndex        =   17
      Top             =   5400
      Width           =   14520
      Begin VB.TextBox txtNCM 
         Height          =   315
         Left            =   600
         MaxLength       =   10
         TabIndex        =   5
         Top             =   330
         Width           =   1455
      End
      Begin VB.TextBox txtDescricao 
         Height          =   315
         Left            =   3060
         TabIndex        =   6
         Top             =   330
         Width           =   11340
      End
      Begin VB.TextBox txtNacFed 
         Height          =   315
         Left            =   1380
         TabIndex        =   7
         Top             =   750
         Width           =   975
      End
      Begin VB.TextBox txtImpFed 
         Height          =   315
         Left            =   3720
         TabIndex        =   8
         Top             =   750
         Width           =   975
      End
      Begin VB.TextBox txtEstadual 
         Height          =   315
         Left            =   5760
         TabIndex        =   9
         Top             =   750
         Width           =   975
      End
      Begin VB.TextBox txtMunicipal 
         Height          =   315
         Left            =   7860
         TabIndex        =   10
         Top             =   750
         Width           =   975
      End
      Begin VB.ComboBox cboClassIBS 
         Height          =   315
         Left            =   1380
         TabIndex        =   11
         Top             =   1200
         Width           =   1935
      End
      Begin VB.ComboBox cboClassIS 
         Height          =   315
         Left            =   4560
         TabIndex        =   12
         Top             =   1200
         Width           =   2055
      End
      Begin VB.ComboBox cboTipoCalcIS 
         Height          =   315
         Left            =   8100
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   "NCM:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   405
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   2220
         TabIndex        =   19
         Top             =   360
         Width           =   765
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         Caption         =   "Nac. Federal %:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   750
         Width           =   1125
      End
      Begin VB.Label lbl4 
         AutoSize        =   -1  'True
         Caption         =   "Imp. Federal %:"
         Height          =   195
         Left            =   2520
         TabIndex        =   21
         Top             =   750
         Width           =   1080
      End
      Begin VB.Label lbl5 
         AutoSize        =   -1  'True
         Caption         =   "Estadual %:"
         Height          =   195
         Left            =   4860
         TabIndex        =   22
         Top             =   750
         Width           =   825
      End
      Begin VB.Label lbl6 
         AutoSize        =   -1  'True
         Caption         =   "Municipal %:"
         Height          =   195
         Left            =   6900
         TabIndex        =   23
         Top             =   750
         Width           =   885
      End
      Begin VB.Label lbl7 
         AutoSize        =   -1  'True
         Caption         =   "cClassTrib IBS:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   1080
      End
      Begin VB.Label lbl8 
         AutoSize        =   -1  'True
         Caption         =   "cClassTrib IS:"
         Height          =   195
         Left            =   3480
         TabIndex        =   25
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lbl9 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cálculo IS:"
         Height          =   195
         Left            =   6840
         TabIndex        =   26
         Top             =   1200
         Width           =   1125
      End
   End
   Begin VB.Label lblMsg 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "00"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   14415
      TabIndex        =   27
      Top             =   5040
      Width           =   180
   End
End
Attribute VB_Name = "frmNCM_Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sModoEdicao As String  ' "Novo" ou "Editar"

Private Sub Form_Load()
    cboCriterio.AddItem "NCM"
    cboCriterio.AddItem "Descrição"
    cboCriterio.ListIndex = 0

    cboTipoCalcIS.AddItem "0 - Não tem IS"
    cboTipoCalcIS.AddItem "1 - Ad Valorem (%)"
    cboTipoCalcIS.AddItem "2 - Ad Rem (R$)"
    cboTipoCalcIS.ListIndex = 0

    PreencherClassIBS
    PreencherClassIS
    ConfigurarGrid
    HabilitarEdicao False
    txtBusca.Text = "0101"
    cmdConsultar_Click
End Sub

Private Sub ConfigurarGrid()
    With grdNCM
        .TextMatrix(0, 0) = "NCM"
        .TextMatrix(0, 1) = "Descrição"
        .TextMatrix(0, 2) = "Nac. Fed %"
        .TextMatrix(0, 3) = "Imp. Fed %"
        .TextMatrix(0, 4) = "Estadual %"
        .TextMatrix(0, 5) = "Munic. %"
        .TextMatrix(0, 6) = "ClassTrib IBS"
        .TextMatrix(0, 7) = "ClassTrib IS"
        .TextMatrix(0, 8) = "Tipo IS"
        .ColWidth(0) = 1000
        .ColWidth(1) = 6000
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1100
        .ColWidth(7) = 1000
        .ColWidth(8) = 900
    End With
End Sub

Private Sub PreencherClassIBS()
    Dim rCls As ADODB.Recordset
    cboClassIBS.Clear
    cboClassIBS.AddItem "(nenhum)"
    RsOpen rCls, "SELECT DISTINCT cClassTrib FROM TbIBSCBSClassTrib ORDER BY cClassTrib"
    Do While Not rCls.EOF
        cboClassIBS.AddItem ValidateNull(rCls("cClassTrib"))
        rCls.MoveNext
    Loop
    If rCls.State <> 0 Then rCls.Close
    cboClassIBS.ListIndex = 0
End Sub

Private Sub PreencherClassIS()
    cboClassIS.Clear
    cboClassIS.AddItem "(nenhum)"
    cboClassIS.AddItem "900001 - Vinhos, Espumantes e Cachaças"
    cboClassIS.AddItem "900002 - Cervejas e Chopes"
    cboClassIS.AddItem "900003 - Uísque, Gin, Vodka (Destilados)"
    cboClassIS.AddItem "900010 - Refrigerantes e Sucos com Açúcar"
    cboClassIS.AddItem "900011 - Energéticos e Bebidas Esportivas"
    cboClassIS.AddItem "900020 - Veículos (Automóveis poluentes)"
    cboClassIS.AddItem "900040 - Cigarros e fumo"
    cboClassIS.ListIndex = 0
End Sub

Private Sub CarregarGrid(sFiltroSQL As String, Optional bMostrarAnim As Boolean = False)
    Const MAX_TODOS As Long = 1000
    Dim rG As ADODB.Recordset
    Dim sSQL As String
    Dim i As Long
    Dim nDots As Integer
    Dim bLimitado As Boolean
    bLimitado = False
    If bMostrarAnim Then
        fraCarregando.ZOrder 0
        lblCarregando.Caption = "CARREGANDO..."
        fraCarregando.Visible = True
        DoEvents
    End If
    If sFiltroSQL = "" Then
        sSQL = "SELECT TOP " & MAX_TODOS & " NCM, descricao, nacionalfederal, importadosfederal, " & _
               "estadual, municipal, cClassTrib_IBS, cClassTrib_IS, tipo_calculo_is " & _
               "FROM tbNCM ORDER BY NCM"
        bLimitado = True
    Else
        sSQL = "SELECT NCM, descricao, nacionalfederal, importadosfederal, " & _
               "estadual, municipal, cClassTrib_IBS, cClassTrib_IS, tipo_calculo_is " & _
               "FROM tbNCM WHERE " & sFiltroSQL & " ORDER BY NCM"
    End If
    RsOpen rG, sSQL
    grdNCM.rows = 1
    Do While Not rG.EOF
        grdNCM.rows = grdNCM.rows + 1
        i = grdNCM.rows - 1
        grdNCM.TextMatrix(i, 0) = ValidateNull(rG("NCM"))
        grdNCM.TextMatrix(i, 1) = ValidateNull(rG("descricao"))
        grdNCM.TextMatrix(i, 2) = FormatNumber(rG("nacionalfederal"), 2)
        grdNCM.TextMatrix(i, 3) = FormatNumber(rG("importadosfederal"), 2)
        grdNCM.TextMatrix(i, 4) = FormatNumber(rG("estadual"), 2)
        grdNCM.TextMatrix(i, 5) = FormatNumber(rG("municipal"), 2)
        grdNCM.TextMatrix(i, 6) = ValidateNull(rG("cClassTrib_IBS"))
        grdNCM.TextMatrix(i, 7) = ValidateNull(rG("cClassTrib_IS"))
        grdNCM.TextMatrix(i, 8) = ValidateNull(rG("tipo_calculo_is"))
        If bMostrarAnim And (i Mod 100 = 0) Then
            nDots = (nDots Mod 3) + 1
            lblCarregando.Caption = "CARREGANDO" & String(nDots, ".")
            DoEvents
        End If
        rG.MoveNext
    Loop
    If rG.State <> 0 Then rG.Close
    If bMostrarAnim Then fraCarregando.Visible = False
    If bLimitado Then
        lblMsg.Caption = "Exibindo os primeiros " & MAX_TODOS & " registros. Use a busca para filtrar."
    Else
        lblMsg.Caption = grdNCM.rows - 1 & " registro(s) encontrado(s)."
    End If
End Sub

Private Sub HabilitarEdicao(bHab As Boolean)
    txtNCM.Enabled = bHab
    txtDescricao.Enabled = bHab
    txtNacFed.Enabled = bHab
    txtImpFed.Enabled = bHab
    txtEstadual.Enabled = bHab
    txtMunicipal.Enabled = bHab
    cboClassIBS.Enabled = bHab
    cboClassIS.Enabled = bHab
    cboTipoCalcIS.Enabled = bHab
    cmdSalvar.Enabled = bHab
    cmdCancelar.Enabled = bHab
    txtNCM.Locked = (sModoEdicao = "Editar")
End Sub

Private Sub LimparCampos()
    txtNCM.Text = ""
    txtDescricao.Text = ""
    txtNacFed.Text = "0,00"
    txtImpFed.Text = "0,00"
    txtEstadual.Text = "0,00"
    txtMunicipal.Text = "0,00"
    cboClassIBS.ListIndex = 0
    cboClassIS.ListIndex = 0
    cboTipoCalcIS.ListIndex = 0
End Sub

Private Sub PreencherCampos(nRow As Integer)
    txtNCM.Text = grdNCM.TextMatrix(nRow, 0)
    txtDescricao.Text = grdNCM.TextMatrix(nRow, 1)
    txtNacFed.Text = grdNCM.TextMatrix(nRow, 2)
    txtImpFed.Text = grdNCM.TextMatrix(nRow, 3)
    txtEstadual.Text = grdNCM.TextMatrix(nRow, 4)
    txtMunicipal.Text = grdNCM.TextMatrix(nRow, 5)
    SelecionarNoCombo cboClassIBS, grdNCM.TextMatrix(nRow, 6)
    SelecionarNoCombo cboClassIS, Left(grdNCM.TextMatrix(nRow, 7), 6)
    cboTipoCalcIS.ListIndex = Val(grdNCM.TextMatrix(nRow, 8))
End Sub

Private Sub SelecionarNoCombo(cbo As ComboBox, sVal As String)
    Dim i As Integer
    cbo.ListIndex = 0
    If sVal = "" Then Exit Sub
    For i = 0 To cbo.ListCount - 1
        If Left(cbo.List(i), Len(sVal)) = sVal Then
            cbo.ListIndex = i
            Exit Sub
        End If
    Next i
End Sub

Private Sub grdNCM_Click()
    If grdNCM.Row < 1 Then Exit Sub
    sModoEdicao = "Editar"
    PreencherCampos grdNCM.Row
    HabilitarEdicao True
    cmdExcluir.Enabled = True
    lblMsg.Caption = ""
End Sub

Private Sub cmdConsultar_Click()
    Dim sFiltro As String
    sFiltro = ""
    If Trim(txtBusca.Text) <> "" Then
        If cboCriterio.Text = "NCM" Then
            If Len(Trim(txtBusca.Text)) >= 8 Then
                sFiltro = "NCM = '" & Trim(txtBusca.Text) & "'"
            Else
                sFiltro = "NCM LIKE '" & Trim(txtBusca.Text) & "%'"
            End If
        Else
            sFiltro = "descricao LIKE '%" & Trim(txtBusca.Text) & "%'"
        End If
    End If
    CarregarGrid sFiltro
End Sub

Private Sub cmdLimpar_Click()
    txtBusca.Text = ""
    CarregarGrid "", True
End Sub

Private Sub cmdNovo_Click()
    sModoEdicao = "Novo"
    LimparCampos
    HabilitarEdicao True
    cmdExcluir.Enabled = False
    txtNCM.SetFocus
End Sub

Private Function LerDecimal(sTxt As String) As Double
    LerDecimal = Val(Replace(Replace(Trim(sTxt), ".", ""), ",", "."))
End Function

Private Function SQLNum(d As Double) As String
    SQLNum = Replace(CStr(d), ",", ".")
End Function

Private Sub cmdSalvar_Click()
    If Trim(txtNCM.Text) = "" Then
        MsgBox "Informe o código NCM.", vbExclamation
        txtNCM.SetFocus: Exit Sub
    End If
    Dim sNCM    As String: sNCM = Trim(txtNCM.Text)
    Dim sDesc   As String: sDesc = Replace(Trim(txtDescricao.Text), "'", "''")
    Dim dNF     As Double: dNF = LerDecimal(txtNacFed.Text)
    Dim dIF     As Double: dIF = LerDecimal(txtImpFed.Text)
    Dim dEst    As Double: dEst = LerDecimal(txtEstadual.Text)
    Dim dMun    As Double: dMun = LerDecimal(txtMunicipal.Text)
    Dim sCIBS   As String
    Dim sCIS    As String
    Dim nTipo   As Integer
    sCIBS = IIf(cboClassIBS.ListIndex <= 0, "", Left(cboClassIBS.Text, 6))
    sCIS = IIf(cboClassIS.ListIndex <= 0, "", Left(cboClassIS.Text, 6))
    nTipo = cboTipoCalcIS.ListIndex
    Dim sSQL As String
    Dim nExiste As Long
    On Error GoTo ErrSalvar
    If sModoEdicao = "Novo" Then
        nExiste = SQLExecutaRetorno("SELECT COUNT(*) AS N FROM tbNCM WHERE NCM='" & sNCM & "'", "N", 0)
        If nExiste > 0 Then
            MsgBox "NCM " & sNCM & " já está cadastrado. Use o grid para editar.", vbExclamation
            Exit Sub
        End If
        sSQL = "INSERT INTO tbNCM (NCM, descricao, nacionalfederal, importadosfederal, " & _
               "estadual, municipal, cClassTrib_IBS, cClassTrib_IS, tipo_calculo_is) VALUES (" & _
               "'" & sNCM & "', '" & sDesc & "', " & _
               SQLNum(dNF) & ", " & SQLNum(dIF) & ", " & SQLNum(dEst) & ", " & SQLNum(dMun) & ", " & _
               "'" & sCIBS & "', '" & sCIS & "', " & nTipo & ")"
    Else
        sSQL = "UPDATE tbNCM SET " & _
               "descricao='" & sDesc & "', " & _
               "nacionalfederal=" & SQLNum(dNF) & ", importadosfederal=" & SQLNum(dIF) & ", " & _
               "estadual=" & SQLNum(dEst) & ", municipal=" & SQLNum(dMun) & ", " & _
               "cClassTrib_IBS='" & sCIBS & "', " & _
               "cClassTrib_IS='" & sCIS & "', " & _
               "tipo_calculo_is=" & nTipo & " " & _
               "WHERE NCM='" & sNCM & "'"
    End If
    dbData.Execute sSQL
    lblMsg.Caption = "Registro salvo com sucesso."
    cmdConsultar_Click
    HabilitarEdicao False
    cmdExcluir.Enabled = False
    sModoEdicao = ""
    Exit Sub
ErrSalvar:
    MsgBox "Erro ao salvar: " & Err.Description, vbCritical
End Sub

Private Sub cmdExcluir_Click()
    If Trim(txtNCM.Text) = "" Then Exit Sub
    If MsgBox("Excluir o NCM " & txtNCM.Text & "?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    On Error GoTo ErrExcluir
    dbData.Execute "DELETE FROM tbNCM WHERE NCM='" & Trim(txtNCM.Text) & "'"
    lblMsg.Caption = "Registro excluído."
    LimparCampos
    HabilitarEdicao False
    cmdExcluir.Enabled = False
    cmdConsultar_Click
    Exit Sub
ErrExcluir:
    MsgBox "Erro ao excluir: " & Err.Description, vbCritical
End Sub

Private Sub cmdCancelar_Click()
    LimparCampos
    HabilitarEdicao False
    cmdExcluir.Enabled = False
    sModoEdicao = ""
    lblMsg.Caption = ""
End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub

Private Sub txtBusca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdConsultar_Click
End Sub
