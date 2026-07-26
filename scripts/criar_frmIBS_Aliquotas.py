# -*- coding: utf-8 -*-
path = 'C:/Projeto/OnlineCommerce/Forms/frmIBS_Aliquotas.frm'

ti = [0]
def T(): v = ti[0]; ti[0] += 1; return v

def lbl(name, cap, left, top, width, height=240, bold=False, parent_indent=1):
    ind = '   ' * parent_indent
    s  = f'{ind}Begin VB.Label {name}\n'
    s += f'{ind}   Caption         =   "{cap}"\n'
    s += f'{ind}   Height          =   {height}\n'
    s += f'{ind}   Left            =   {left}\n'
    s += f'{ind}   TabIndex        =   {T()}\n'
    s += f'{ind}   Top             =   {top}\n'
    s += f'{ind}   Width           =   {width}\n'
    if bold: s += f'{ind}   FontBold        =   -1  \'True\n'
    s += f'{ind}End\n'
    return s

def txt(name, left, top, width, height=288, ro=False, parent_indent=1, maxlen=0):
    ind = '   ' * parent_indent
    s  = f'{ind}Begin VB.TextBox {name}\n'
    s += f'{ind}   Height          =   {height}\n'
    s += f'{ind}   Left            =   {left}\n'
    if maxlen: s += f'{ind}   MaxLength       =   {maxlen}\n'
    if ro: s += f'{ind}   Locked          =   -1  \'True\n'
    s += f'{ind}   TabIndex        =   {T()}\n'
    s += f'{ind}   Top             =   {top}\n'
    s += f'{ind}   Width           =   {width}\n'
    s += f'{ind}End\n'
    return s

def cbo(name, left, top, width, style=2, height=288, parent_indent=1):
    ind = '   ' * parent_indent
    s  = f'{ind}Begin VB.ComboBox {name}\n'
    s += f'{ind}   Height          =   {height}\n'
    s += f'{ind}   Left            =   {left}\n'
    s += f'{ind}   Style           =   {style}\n'
    s += f'{ind}   TabIndex        =   {T()}\n'
    s += f'{ind}   Top             =   {top}\n'
    s += f'{ind}   Width           =   {width}\n'
    s += f'{ind}End\n'
    return s

def cmd(name, cap, left, top, width=1680, height=360, parent_indent=1):
    ind = '   ' * parent_indent
    s  = f'{ind}Begin VB.CommandButton {name}\n'
    s += f'{ind}   Caption         =   "{cap}"\n'
    s += f'{ind}   Height          =   {height}\n'
    s += f'{ind}   Left            =   {left}\n'
    s += f'{ind}   TabIndex        =   {T()}\n'
    s += f'{ind}   Top             =   {top}\n'
    s += f'{ind}   Width           =   {width}\n'
    s += f'{ind}End\n'
    return s

def opt(name, cap, left, top, width, height=240, parent_indent=2):
    ind = '   ' * parent_indent
    s  = f'{ind}Begin VB.OptionButton {name}\n'
    s += f'{ind}   Caption         =   "{cap}"\n'
    s += f'{ind}   Height          =   {height}\n'
    s += f'{ind}   Left            =   {left}\n'
    s += f'{ind}   TabIndex        =   {T()}\n'
    s += f'{ind}   Top             =   {top}\n'
    s += f'{ind}   Width           =   {width}\n'
    s += f'{ind}End\n'
    return s

def grid(name, left, top, width, height, parent_indent=1):
    ind = '   ' * parent_indent
    ex = round(width * 2540 / 1440)
    ey = round(height * 2540 / 1440)
    s  = f'{ind}Begin MSFlexGridLib.MSFlexGrid {name}\n'
    s += f'{ind}   Height          =   {height}\n'
    s += f'{ind}   Left            =   {left}\n'
    s += f'{ind}   TabIndex        =   {T()}\n'
    s += f'{ind}   Top             =   {top}\n'
    s += f'{ind}   Width           =   {width}\n'
    s += f'{ind}   _ExtentX        =   {ex}\n'
    s += f'{ind}   _ExtentY        =   {ey}\n'
    s += f'{ind}   _Version        =   393216\n'
    s += f'{ind}End\n'
    return s

# ── FORM ──────────────────────────────────────────────────────────────────────
f  = 'VERSION 5.00\n'
f += 'Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSFLXGRD.OCX"\n'
f += 'Begin VB.Form frmIBS_Aliquotas\n'
f += '   Caption         =   "IBS - Al\xedquotas por Estado e Munic\xedpio"\n'
f += '   ClientHeight    =   12720\n'
f += '   ClientLeft      =   120\n'
f += '   ClientTop       =   432\n'
f += '   ClientWidth     =   14400\n'
f += '   LinkTopic       =   "Form1"\n'
f += '   ScaleHeight     =   12720\n'
f += '   ScaleWidth      =   14400\n'
f += '   StartUpPosition =   2  \'CenterScreen\n'

# ── FRAME ESTADO ──────────────────────────────────────────────────────────────
f += '   Begin VB.Frame fraEstado\n'
f += '      Caption         =   "Al\xedquota IBS por Estado"\n'
f += '      Height          =   4920\n'
f += '      Left            =   120\n'
f += f'      TabIndex        =   {T()}\n'
f += '      Top             =   120\n'
f += '      Width           =   14160\n'

# Filtro UF
f += lbl('lblFiltroUF','Filtrar UF:',120,360,840,parent_indent=2)
f += cbo('cboFiltroUF',1080,360,1440,parent_indent=2)
f += cmd('cmdFiltrarEstado','Filtrar',2640,360,1200,288,parent_indent=2)
f += cmd('cmdTodosEstado','Todos',3960,360,1200,288,parent_indent=2)

# Grid estado (cols: UF, Aliq%, Vig.Ini, Vig.Fim, Id[hidden])
f += grid('lstEstado',120,720,8160,2880,parent_indent=2)

# Edicao estado (direita do grid)
f += lbl('lblUFEditE','UF:',8400,720,480,parent_indent=2)
f += cbo('cboUFEdit',9000,720,1680,parent_indent=2)
f += lbl('lblAliqUFEdit','Al\xedq. %:',8400,1080,840,parent_indent=2)
f += txt('txtAliqUFEdit',9360,1080,1440,parent_indent=2)
f += lbl('lblDIniUFEdit','Vig. Ini:',8400,1440,840,parent_indent=2)
f += txt('txtDIniUFEdit',9360,1440,1440,parent_indent=2)
f += lbl('lblDFimUFEdit','Vig. Fim:',8400,1800,840,parent_indent=2)
f += txt('txtDFimUFEdit',9360,1800,1440,parent_indent=2)
f += cmd('cmdNovoEstado','Novo',8400,2400,1440,360,parent_indent=2)
f += cmd('cmdSalvarEstado','Salvar',10080,2400,1560,360,parent_indent=2)
f += cmd('cmdExcluirEstado','Excluir',11760,2400,1920,360,parent_indent=2)

# Sub-frame: aplicar todos estados
f += '      Begin VB.Frame fraAplicarTodosUF\n'
f += '         Caption         =   "Aplicar para TODOS os Estados"\n'
f += '         Height          =   960\n'
f += '         Left            =   120\n'
f += f'         TabIndex        =   {T()}\n'
f += '         Top             =   3840\n'
f += '         Width           =   14040\n'
f += lbl('lblAliqTodosUF','Al\xedq. %:',120,360,720,parent_indent=3)
f += txt('txtAliqTodosUF',960,360,1200,parent_indent=3)
f += lbl('lblDIniTodosUF','Vig. Ini:',2280,360,840,parent_indent=3)
f += txt('txtDIniTodosUF',3240,360,1440,parent_indent=3)
f += lbl('lblDFimTodosUF','Vig. Fim:',4800,360,840,parent_indent=3)
f += txt('txtDFimTodosUF',5760,360,1440,parent_indent=3)
f += cmd('cmdAplicarTodosUF','Aplicar para TODOS os Estados',7440,240,3000,480,parent_indent=3)
f += '      End\n'

f += '   End\n'  # fraEstado

# ── FRAME MUNICIPIO ───────────────────────────────────────────────────────────
f += '   Begin VB.Frame fraMunicipio\n'
f += '      Caption         =   "Al\xedquota IBS por Munic\xedpio"\n'
f += '      Height          =   7320\n'
f += '      Left            =   120\n'
f += f'      TabIndex        =   {T()}\n'
f += '      Top             =   5160\n'
f += '      Width           =   14160\n'

# Filtros municipio
f += lbl('lblFiltroUFMun','UF:',120,360,480,parent_indent=2)
f += cbo('cboFiltroUFMun',720,360,1200,parent_indent=2)
f += lbl('lblFiltroCidade','Cidade:',2040,360,720,parent_indent=2)
f += txt('txtFiltroCidade',2880,360,2880,height=288,parent_indent=2)
f += cmd('cmdFiltrarMun','Filtrar',5880,360,1200,288,parent_indent=2)
f += cmd('cmdLimparMun','Limpar',7200,360,1200,288,parent_indent=2)

# Grid municipio (cols: UF, CodMun, Nome, Aliq%, Vig.Ini, Vig.Fim, Id[hidden])
f += grid('lstMunicipio',120,720,14040,3000,parent_indent=2)

# Edicao municipio
f += lbl('lblCodMunEdit','Cod. Mun:',120,3840,960,parent_indent=2)
f += txt('txtCodMunEdit',1200,3840,960,ro=True,parent_indent=2)
f += lbl('lblNomeMunEdit','Nome:',2280,3840,600,parent_indent=2)
f += txt('txtNomeMunEdit',3000,3840,4200,ro=True,parent_indent=2)
f += lbl('lblAliqMunEdit','Al\xedq. %:',120,4200,840,parent_indent=2)
f += txt('txtAliqMunEdit',1080,4200,1440,parent_indent=2)
f += lbl('lblDIniMunEdit','Vig. Ini:',2640,4200,840,parent_indent=2)
f += txt('txtDIniMunEdit',3600,4200,1440,parent_indent=2)
f += lbl('lblDFimMunEdit','Vig. Fim:',5160,4200,840,parent_indent=2)
f += txt('txtDFimMunEdit',6120,4200,1440,parent_indent=2)
f += cmd('cmdNovoMun','Novo',120,4680,1440,360,parent_indent=2)
f += cmd('cmdSalvarMun','Salvar',1800,4680,1560,360,parent_indent=2)
f += cmd('cmdExcluirMun','Excluir',3600,4680,1680,360,parent_indent=2)

# Sub-frame: aplicar todos municipios
f += '      Begin VB.Frame fraAplicarTodosMun\n'
f += '         Caption         =   "Aplicar Al\xedquota em Massa"\n'
f += '         Height          =   1800\n'
f += '         Left            =   5400\n'
f += f'         TabIndex        =   {T()}\n'
f += '         Top             =   3720\n'
f += '         Width           =   8640\n'
f += opt('optTodosMun','Todos os munic\xedpios',120,300,3600,parent_indent=3)
f += opt('optEstadoFiltrado','Somente estado filtrado',3840,300,4680,parent_indent=3)
f += lbl('lblAliqTodosMun','Al\xedq. %:',120,600,720,parent_indent=3)
f += txt('txtAliqTodosMun',960,600,1200,parent_indent=3)
f += lbl('lblDIniTodosMun','Vig. Ini:',2280,600,840,parent_indent=3)
f += txt('txtDIniTodosMun',3240,600,1440,parent_indent=3)
f += lbl('lblDFimTodosMun','Vig. Fim:',4800,600,840,parent_indent=3)
f += txt('txtDFimTodosMun',5760,600,1440,parent_indent=3)
f += cmd('cmdAplicarTodosMun','Aplicar',120,1200,3600,480,parent_indent=3)
f += '      End\n'

f += '   End\n'  # fraMunicipio

# ── BARRA INFERIOR ────────────────────────────────────────────────────────────
f += cmd('cmdAtualizarCidade','Atualizar Cidade com Al\xedquotas Vigentes de Hoje',120,12600,5760,480)
f += lbl('lblStatusSync','',6120,12600,8160,480)
f += lbl('lblAvisoSync','',120,12360,14160,240)

# Label de carregamento (centralizado, visivel apenas durante carga)
f += (
    '   Begin VB.Label lblCarregando\n'
    '      Alignment       =   2  \'Center\n'
    '      BackColor       =   &H00FFFFFF&\n'
    '      BackStyle       =   1\n'
    '      Caption         =   "CARREGANDO..."\n'
    '      FontBold        =   -1  \'True\n'
    '      FontSize        =   16\n'
    '      ForeColor       =   &H000000FF&\n'
    f'      Height          =   600\n'
    '      Left            =   4320\n'
    f'      TabIndex        =   {T()}\n'
    '      Top             =   6360\n'
    '      Visible         =   0  \'False\n'
    '      Width           =   5760\n'
    '   End\n'
)

f += 'End\n'

# ── ATTRIBUTES ────────────────────────────────────────────────────────────────
f += 'Attribute VB_Name = "frmIBS_Aliquotas"\n'
f += 'Attribute VB_GlobalNameSpace = False\n'
f += 'Attribute VB_Creatable = False\n'
f += 'Attribute VB_PredeclaredId = True\n'
f += 'Attribute VB_Exposed = False\n'

# ── VB6 CODE ──────────────────────────────────────────────────────────────────
code = r"""Option Explicit
Dim bNovoEstado As Boolean
Dim bNovoMun    As Boolean
Dim iEstadoId   As Long
Dim iMunId      As Long
Dim sCodMunSel  As String

' ============================================================
' FORM LOAD
' ============================================================
Private Sub Form_Load()
    CarregarComboUF cboFiltroUF, True
    CarregarComboUF cboUFEdit, False
    CarregarComboUF cboFiltroUFMun, True
    optTodosMun.Value = True
    ' Carrega municipios so do primeiro estado por padrao
    If cboFiltroUFMun.ListCount > 1 Then cboFiltroUFMun.ListIndex = 1

    With lstEstado
        .Cols = 5
        .ColWidth(0) = 840:  .Row = 0: .Col = 0: .Text = "UF"
        .ColWidth(1) = 1440: .Row = 0: .Col = 1: .Text = "Al" & Chr(237) & "q. %"
        .ColWidth(2) = 1680: .Row = 0: .Col = 2: .Text = "Vig. Ini"
        .ColWidth(3) = 1680: .Row = 0: .Col = 3: .Text = "Vig. Fim"
        .ColWidth(4) = 0     ' Id oculto
        .AllowUserResizing = 1: .SelectionMode = 1
    End With

    With lstMunicipio
        .Cols = 7
        .ColWidth(0) = 720:  .Row = 0: .Col = 0: .Text = "UF"
        .ColWidth(1) = 960:  .Row = 0: .Col = 1: .Text = "Cod. Mun"
        .ColWidth(2) = 5040: .Row = 0: .Col = 2: .Text = "Nome"
        .ColWidth(3) = 1200: .Row = 0: .Col = 3: .Text = "Al" & Chr(237) & "q. %"
        .ColWidth(4) = 1560: .Row = 0: .Col = 4: .Text = "Vig. Ini"
        .ColWidth(5) = 1560: .Row = 0: .Col = 5: .Text = "Vig. Fim"
        .ColWidth(6) = 0     ' Id oculto
        .AllowUserResizing = 1: .SelectionMode = 1
    End With

    CarregarIBSEstado
    CarregarIBSMunicipio
    VerificarAliquotasPendentes
End Sub

Private Sub CarregarComboUF(cbo As ComboBox, comTodos As Boolean)
    Dim rRs As ADODB.Recordset
    cbo.Clear
    If comTodos Then cbo.AddItem "(Todos)"
    RsOpen rRs, "SELECT DISTINCT UF FROM Cidade WHERE UF IS NOT NULL ORDER BY UF"
    Do While Not rRs.EOF
        cbo.AddItem rRs("UF")
        rRs.MoveNext
    Loop
    If rRs.State <> 0 Then rRs.Close
    cbo.ListIndex = 0
End Sub

' ============================================================
' IBS ESTADO
' ============================================================
Private Sub CarregarIBSEstado()
    Dim rRs As ADODB.Recordset
    Dim sql As String, r As Long
    sql = "SELECT Id, UF, IBSUFpAliq, dIniVig, dFimVig FROM IBS_Estado"
    If cboFiltroUF.ListIndex > 0 Then
        sql = sql & " WHERE UF = '" & cboFiltroUF.Text & "'"
    End If
    sql = sql & " ORDER BY UF, dIniVig"
    lstEstado.Rows = 1
    RsOpen rRs, sql
    Do While Not rRs.EOF
        lstEstado.AddItem ""
        r = lstEstado.Rows - 1
        lstEstado.Row = r: lstEstado.Col = 0: lstEstado.Text = rRs("UF")
        lstEstado.Row = r: lstEstado.Col = 1: lstEstado.Text = Format(rRs("IBSUFpAliq"), "0.00")
        lstEstado.Row = r: lstEstado.Col = 2: lstEstado.Text = IIf(IsNull(rRs("dIniVig")), "", Format(rRs("dIniVig"), "dd/mm/yyyy"))
        lstEstado.Row = r: lstEstado.Col = 3: lstEstado.Text = IIf(IsNull(rRs("dFimVig")), "", Format(rRs("dFimVig"), "dd/mm/yyyy"))
        lstEstado.Row = r: lstEstado.Col = 4: lstEstado.Text = rRs("Id")
        rRs.MoveNext
    Loop
    If rRs.State <> 0 Then rRs.Close
End Sub

Private Sub lstEstado_Click()
    If lstEstado.Row = 0 Then Exit Sub
    iEstadoId = Val(lstEstado.TextMatrix(lstEstado.Row, 4))
    cboUFEdit.Text       = lstEstado.TextMatrix(lstEstado.Row, 0)
    txtAliqUFEdit.Text   = lstEstado.TextMatrix(lstEstado.Row, 1)
    txtDIniUFEdit.Text   = lstEstado.TextMatrix(lstEstado.Row, 2)
    txtDFimUFEdit.Text   = lstEstado.TextMatrix(lstEstado.Row, 3)
    bNovoEstado = False
End Sub

Private Sub cmdFiltrarEstado_Click(): CarregarIBSEstado: End Sub
Private Sub cmdTodosEstado_Click(): cboFiltroUF.ListIndex = 0: CarregarIBSEstado: End Sub

Private Sub cmdNovoEstado_Click()
    bNovoEstado = True: iEstadoId = 0
    cboUFEdit.ListIndex = 0: txtAliqUFEdit.Text = "": txtDIniUFEdit.Text = "": txtDFimUFEdit.Text = ""
    cboUFEdit.SetFocus
End Sub

Private Sub cmdSalvarEstado_Click()
    Dim sUF As String, sAliq As String, sDIni As String, sDFim As String
    sUF   = Trim(cboUFEdit.Text)
    sAliq = Replace(Trim(txtAliqUFEdit.Text), ",", ".")
    sDIni = Trim(txtDIniUFEdit.Text)
    sDFim = Trim(txtDFimUFEdit.Text)
    If sUF = "" Or sAliq = "" Or sDIni = "" Then
        MsgBox "UF, Al" & Chr(237) & "q e Vig. Ini s" & Chr(227) & "o obrigat" & Chr(243) & "rios.", vbExclamation: Exit Sub
    End If
    Dim sIdEstado As String
    sIdEstado = SQLExecutaRetorno("SELECT TOP 1 CAST(IdEstado AS VARCHAR) AS IdEstado FROM Cidade WHERE UF='" & sUF & "'", "IdEstado", "")
    If sIdEstado = "" Then MsgBox "UF n" & Chr(227) & "o encontrada em Cidade.", vbExclamation: Exit Sub
    Dim sDFimSQL As String
    sDFimSQL = IIf(sDFim = "", "NULL", "CONVERT(date,'" & sDFim & "',103)")
    Dim sql As String
    If bNovoEstado Then
        sql = "INSERT INTO IBS_Estado (IdEstado, UF, IBSUFpAliq, dIniVig, dFimVig) VALUES (" & sIdEstado & ",'" & sUF & "'," & sAliq & ",CONVERT(date,'" & sDIni & "',103)," & sDFimSQL & ")"
    Else
        sql = "UPDATE IBS_Estado SET UF='" & sUF & "',IBSUFpAliq=" & sAliq & ",dIniVig=CONVERT(date,'" & sDIni & "',103),dFimVig=" & sDFimSQL & " WHERE Id=" & iEstadoId
    End If
    On Error GoTo ErrE
    vgDb.Execute sql
    MsgBox "Salvo.", vbInformation
    bNovoEstado = False: CarregarIBSEstado
    Exit Sub
ErrE: MsgBox "Erro: " & Err.Description, vbCritical
End Sub

Private Sub cmdExcluirEstado_Click()
    If iEstadoId = 0 Then Exit Sub
    If MsgBox("Excluir este registro de " & cboUFEdit.Text & "?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    On Error GoTo ErrDE
    vgDb.Execute "DELETE FROM IBS_Estado WHERE Id=" & iEstadoId
    iEstadoId = 0: MsgBox "Exclu" & Chr(237) & "do.", vbInformation: CarregarIBSEstado
    Exit Sub
ErrDE: MsgBox "Erro: " & Err.Description, vbCritical
End Sub

Private Sub cmdAplicarTodosUF_Click()
    Dim sAliq As String, sDIni As String, sDFim As String
    sAliq = Replace(Trim(txtAliqTodosUF.Text), ",", ".")
    sDIni = Trim(txtDIniTodosUF.Text)
    sDFim = Trim(txtDFimTodosUF.Text)
    If sAliq = "" Or sDIni = "" Then MsgBox "Preencha Al" & Chr(237) & "q e Vig. Ini.", vbExclamation: Exit Sub
    If MsgBox("Criar al" & Chr(237) & "quota " & txtAliqTodosUF.Text & "% para TODOS os estados?" & Chr(13) & "Vig: " & sDIni & " a " & IIf(sDFim="","sem fim",sDFim), vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Dim sDFimSQL As String
    sDFimSQL = IIf(sDFim = "", "NULL", "CONVERT(date,'" & sDFim & "',103)")
    On Error GoTo ErrAUF
    vgDb.Execute "INSERT INTO IBS_Estado (IdEstado, UF, IBSUFpAliq, dIniVig, dFimVig) SELECT DISTINCT IdEstado, UF, " & sAliq & ", CONVERT(date,'" & sDIni & "',103), " & sDFimSQL & " FROM Cidade WHERE IdEstado IS NOT NULL"
    MsgBox "Al" & Chr(237) & "quotas criadas para todos os estados.", vbInformation
    CarregarIBSEstado
    Exit Sub
ErrAUF: MsgBox "Erro: " & Err.Description, vbCritical
End Sub

' ============================================================
' IBS MUNICIPIO
' ============================================================
Private Sub CarregarIBSMunicipio()
    lblCarregando.Visible = True
    lblCarregando.Caption = "CARREGANDO..."
    DoEvents
    Dim rRs As ADODB.Recordset
    Dim sql As String, r As Long, filtro As String, nDots As Integer
    sql = "SELECT M.Id, C.UF, M.CodigoMunicipio, C.Nome, M.IBSMunpAliq, M.dIniVig, M.dFimVig " & _
          "FROM IBS_Municipio M " & _
          "JOIN (SELECT DISTINCT CAST(CodigoMunicipio AS NVARCHAR(7)) AS CodigoMunicipio, Nome, UF FROM Cidade) C " & _
          "  ON C.CodigoMunicipio = M.CodigoMunicipio"
    If cboFiltroUFMun.ListIndex > 0 Then filtro = filtro & " AND C.UF='" & cboFiltroUFMun.Text & "'"
    If Trim(txtFiltroCidade.Text) <> "" Then filtro = filtro & " AND C.Nome LIKE '%" & Trim(txtFiltroCidade.Text) & "%'"
    If filtro <> "" Then sql = sql & " WHERE" & Mid(filtro, 5)
    sql = sql & " ORDER BY C.UF, C.Nome, M.dIniVig"
    lstMunicipio.Rows = 1
    RsOpen rRs, sql
    Do While Not rRs.EOF
        r = lstMunicipio.Rows - 1
        lstMunicipio.AddItem ""
        lstMunicipio.Row = r: lstMunicipio.Col = 0: lstMunicipio.Text = rRs("UF")
        lstMunicipio.Row = r: lstMunicipio.Col = 1: lstMunicipio.Text = rRs("CodigoMunicipio")
        lstMunicipio.Row = r: lstMunicipio.Col = 2: lstMunicipio.Text = rRs("Nome")
        lstMunicipio.Row = r: lstMunicipio.Col = 3: lstMunicipio.Text = Format(rRs("IBSMunpAliq"), "0.00")
        lstMunicipio.Row = r: lstMunicipio.Col = 4: lstMunicipio.Text = IIf(IsNull(rRs("dIniVig")), "", Format(rRs("dIniVig"), "dd/mm/yyyy"))
        lstMunicipio.Row = r: lstMunicipio.Col = 5: lstMunicipio.Text = IIf(IsNull(rRs("dFimVig")), "", Format(rRs("dFimVig"), "dd/mm/yyyy"))
        lstMunicipio.Row = r: lstMunicipio.Col = 6: lstMunicipio.Text = rRs("Id")
        If (lstMunicipio.Rows - 1) Mod 30 = 0 Then
            nDots = (nDots Mod 3) + 1
            lblCarregando.Caption = "CARREGANDO" & String(nDots, ".")
            DoEvents
        End If
        rRs.MoveNext
    Loop
    If rRs.State <> 0 Then rRs.Close
    lblCarregando.Visible = False
End Sub

Private Sub lstMunicipio_Click()
    If lstMunicipio.Row = 0 Then Exit Sub
    iMunId         = Val(lstMunicipio.TextMatrix(lstMunicipio.Row, 6))
    sCodMunSel     = lstMunicipio.TextMatrix(lstMunicipio.Row, 1)
    txtCodMunEdit.Text  = lstMunicipio.TextMatrix(lstMunicipio.Row, 1)
    txtNomeMunEdit.Text = lstMunicipio.TextMatrix(lstMunicipio.Row, 2)
    txtAliqMunEdit.Text = lstMunicipio.TextMatrix(lstMunicipio.Row, 3)
    txtDIniMunEdit.Text = lstMunicipio.TextMatrix(lstMunicipio.Row, 4)
    txtDFimMunEdit.Text = lstMunicipio.TextMatrix(lstMunicipio.Row, 5)
    bNovoMun = False
End Sub

Private Sub cmdFiltrarMun_Click(): CarregarIBSMunicipio: End Sub
Private Sub cmdLimparMun_Click()
    cboFiltroUFMun.ListIndex = 0: txtFiltroCidade.Text = "": CarregarIBSMunicipio
End Sub

Private Sub cmdNovoMun_Click()
    If sCodMunSel = "" Then MsgBox "Selecione um munic" & Chr(237) & "pio na lista primeiro.", vbExclamation: Exit Sub
    bNovoMun = True: iMunId = 0
    txtAliqMunEdit.Text = "": txtDIniMunEdit.Text = "": txtDFimMunEdit.Text = ""
    txtAliqMunEdit.SetFocus
End Sub

Private Sub cmdSalvarMun_Click()
    Dim sAliq As String, sDIni As String, sDFim As String, sCod As String
    sCod  = Trim(txtCodMunEdit.Text)
    sAliq = Replace(Trim(txtAliqMunEdit.Text), ",", ".")
    sDIni = Trim(txtDIniMunEdit.Text)
    sDFim = Trim(txtDFimMunEdit.Text)
    If sCod = "" Or sAliq = "" Or sDIni = "" Then
        MsgBox "Munic" & Chr(237) & "pio, Al" & Chr(237) & "q e Vig. Ini s" & Chr(227) & "o obrigat" & Chr(243) & "rios.", vbExclamation: Exit Sub
    End If
    Dim sDFimSQL As String
    sDFimSQL = IIf(sDFim = "", "NULL", "CONVERT(date,'" & sDFim & "',103)")
    Dim sql As String
    If bNovoMun Then
        sql = "INSERT INTO IBS_Municipio (CodigoMunicipio, IBSMunpAliq, dIniVig, dFimVig) VALUES ('" & sCod & "'," & sAliq & ",CONVERT(date,'" & sDIni & "',103)," & sDFimSQL & ")"
    Else
        sql = "UPDATE IBS_Municipio SET IBSMunpAliq=" & sAliq & ",dIniVig=CONVERT(date,'" & sDIni & "',103),dFimVig=" & sDFimSQL & " WHERE Id=" & iMunId
    End If
    On Error GoTo ErrM
    vgDb.Execute sql
    MsgBox "Salvo.", vbInformation
    bNovoMun = False: CarregarIBSMunicipio
    Exit Sub
ErrM: MsgBox "Erro: " & Err.Description, vbCritical
End Sub

Private Sub cmdExcluirMun_Click()
    If iMunId = 0 Then Exit Sub
    If MsgBox("Excluir este registro de " & txtNomeMunEdit.Text & "?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    On Error GoTo ErrDM
    vgDb.Execute "DELETE FROM IBS_Municipio WHERE Id=" & iMunId
    iMunId = 0: MsgBox "Exclu" & Chr(237) & "do.", vbInformation: CarregarIBSMunicipio
    Exit Sub
ErrDM: MsgBox "Erro: " & Err.Description, vbCritical
End Sub

Private Sub cmdAplicarTodosMun_Click()
    Dim sAliq As String, sDIni As String, sDFim As String
    sAliq = Replace(Trim(txtAliqTodosMun.Text), ",", ".")
    sDIni = Trim(txtDIniTodosMun.Text)
    sDFim = Trim(txtDFimTodosMun.Text)
    If sAliq = "" Or sDIni = "" Then MsgBox "Preencha Al" & Chr(237) & "q e Vig. Ini.", vbExclamation: Exit Sub
    Dim filtroWhere As String, descr As String
    If optEstadoFiltrado.Value And cboFiltroUFMun.ListIndex > 0 Then
        filtroWhere = " WHERE UF='" & cboFiltroUFMun.Text & "'"
        descr = "munic" & Chr(237) & "pios do estado " & cboFiltroUFMun.Text
    Else
        descr = "TODOS os munic" & Chr(237) & "pios"
    End If
    If MsgBox("Criar al" & Chr(237) & "quota " & txtAliqTodosMun.Text & "% para " & descr & "?" & Chr(13) & "Vig: " & sDIni & " a " & IIf(sDFim="","sem fim",sDFim), vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Dim sDFimSQL As String
    sDFimSQL = IIf(sDFim = "", "NULL", "CONVERT(date,'" & sDFim & "',103)")
    On Error GoTo ErrAM
    vgDb.Execute "INSERT INTO IBS_Municipio (CodigoMunicipio, IBSMunpAliq, dIniVig, dFimVig) SELECT DISTINCT CAST(CodigoMunicipio AS NVARCHAR(7))," & sAliq & ",CONVERT(date,'" & sDIni & "',103)," & sDFimSQL & " FROM Cidade" & filtroWhere
    MsgBox "Al" & Chr(237) & "quotas criadas.", vbInformation: CarregarIBSMunicipio
    Exit Sub
ErrAM: MsgBox "Erro: " & Err.Description, vbCritical
End Sub

' ============================================================
' SINCRONIZAR CIDADE
' ============================================================
Private Sub cmdAtualizarCidade_Click()
    Dim hoje As String
    hoje = Format(Now, "yyyy-mm-dd")
    On Error GoTo ErrSync
    vgDb.Execute "UPDATE C SET C.IBSUFpAliq=E.IBSUFpAliq FROM Cidade C JOIN IBS_Estado E ON C.IdEstado=E.IdEstado WHERE CONVERT(date,'" & hoje & "',23) BETWEEN E.dIniVig AND ISNULL(E.dFimVig,'9999-12-31')"
    vgDb.Execute "UPDATE C SET C.IBSMunpAliq=M.IBSMunpAliq FROM Cidade C JOIN IBS_Municipio M ON CAST(C.CodigoMunicipio AS NVARCHAR(7))=M.CodigoMunicipio WHERE CONVERT(date,'" & hoje & "',23) BETWEEN M.dIniVig AND ISNULL(M.dFimVig,'9999-12-31')"
    lblStatusSync.Caption = "Cidade atualizada em " & Format(Now, "dd/mm/yyyy hh:nn:ss")
    lblAvisoSync.Caption = ""
    MsgBox "Tabela Cidade atualizada com al" & Chr(237) & "quotas vigentes de hoje.", vbInformation
    Exit Sub
ErrSync: MsgBox "Erro ao atualizar Cidade: " & Err.Description, vbCritical
End Sub

Private Sub VerificarAliquotasPendentes()
    Dim n As String
    n = SQLExecutaRetorno("SELECT CAST(COUNT(*) AS VARCHAR) AS Total FROM IBS_Estado WHERE dIniVig <= GETDATE() AND NOT EXISTS (SELECT 1 FROM Cidade C JOIN IBS_Estado E2 ON C.IdEstado=E2.IdEstado WHERE E2.Id=IBS_Estado.Id AND C.IBSUFpAliq=IBS_Estado.IBSUFpAliq)", "Total", "0")
    If Val(n) > 0 Then
        lblAvisoSync.Caption = "AVISO: Ha al" & Chr(237) & "quotas vigentes que ainda n" & Chr(227) & "o foram aplicadas em Cidade. Clique em 'Atualizar Cidade'."
    End If
End Sub
"""

content = f + code

out = content.encode('windows-1252')
out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(path, 'wb').write(out)
print('frmIBS_Aliquotas.frm gerado com sucesso')
