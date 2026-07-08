"""
Patch frmISClassTrib.frm: adiciona campo ISCST
- Nova combo cboISCSTEdit + label no fraEdit
- Form_Load: Cols=9, col 8 header, popula cboISCSTEdit
- CarregarRegistros: busca ISCST, grava col 8
- lstRegistros_Click: popula cboISCSTEdit
- cmdSalvar_Click: salva ISCST no INSERT/UPDATE
- LimparCampos: reseta cboISCSTEdit
"""
FILE = r"C:\Projeto\OnlineCommerce\Forms\frmISClassTrib.frm"
with open(FILE, "rb") as f:
    raw = f.read()
data = raw.replace(b"\r\n", b"\n").replace(b"\r", b"\n")

errors = 0

def patch(old, new, label):
    global data, errors
    c = data.count(old)
    if c != 1:
        print(f"ERRO [{label}] ({c}x)")
        errors += 1
    else:
        data = data.replace(old, new, 1)
        print(f"OK: {label}")

# 1. Adicionar cboISCSTEdit e lblISCSTEdit dentro de fraEdit (antes do End do frame)
patch(
    b"       Begin VB.Label lblDFim \n"
    b"          AutoSize        =   -1  'True\n"
    b"          Caption         =   \"Vig\xeancia Final:\"\n"
    b"          Height          =   195\n"
    b"          Left            =   3060\n"
    b"          TabIndex        =   15\n"
    b"          Top             =   1280\n"
    b"          Width           =   1035\n"
    b"       End\n"
    b"    End\n",
    b"       Begin VB.Label lblDFim \n"
    b"          AutoSize        =   -1  'True\n"
    b"          Caption         =   \"Vig\xeancia Final:\"\n"
    b"          Height          =   195\n"
    b"          Left            =   3060\n"
    b"          TabIndex        =   15\n"
    b"          Top             =   1280\n"
    b"          Width           =   1035\n"
    b"       End\n"
    b"       Begin VB.ComboBox cboISCSTEdit \n"
    b"          Height          =   315\n"
    b"          Left            =   5520\n"
    b"          Style           =   2  'Dropdown List\n"
    b"          TabIndex        =   16\n"
    b"          Top             =   630\n"
    b"          Width           =   1800\n"
    b"       End\n"
    b"       Begin VB.Label lblISCSTEdit \n"
    b"          AutoSize        =   -1  'True\n"
    b"          Caption         =   \"CST IS:\"\n"
    b"          Height          =   195\n"
    b"          Left            =   4860\n"
    b"          TabIndex        =   17\n"
    b"          Top             =   636\n"
    b"          Width           =   555\n"
    b"       End\n"
    b"    End\n",
    "controles fraEdit"
)

# 2. Form_Load: Cols=9, header col 8, popula cboISCSTEdit
patch(
    b"    With lstRegistros\n"
    b"        .Cols = 8\n"
    b"        .ColWidth(0) = 0\n"
    b"        .ColWidth(1) = 1020\n"
    b"        .ColWidth(2) = 4500\n"
    b"        .ColWidth(3) = 1500\n"
    b"        .ColWidth(4) = 1080\n"
    b"        .ColWidth(5) = 1200\n"
    b"        .ColWidth(6) = 1500\n"
    b"        .ColWidth(7) = 1500\n"
    b"        .Row = 0: .Col = 0: .Text = \"Id\"\n"
    b"        .Row = 0: .Col = 1: .Text = \"C\xf3d. IS\"\n"
    b"        .Row = 0: .Col = 2: .Text = \"Descri\xe7\xe3o\"\n"
    b"        .Row = 0: .Col = 3: .Text = \"Tipo\"\n"
    b"        .Row = 0: .Col = 4: .Text = \"Al\xedq. %\"\n"
    b"        .Row = 0: .Col = 5: .Text = \"Val. Unid\"\n"
    b"        .Row = 0: .Col = 6: .Text = \"Vig. Ini\"\n"
    b"        .Row = 0: .Col = 7: .Text = \"Vig. Fim\"\n"
    b"        .AllowUserResizing = 1\n"
    b"        .SelectionMode = 1\n"
    b"    End With\n"
    b"    cboTipo.Clear\n"
    b"    cboTipo.AddItem \"1 - Ad Valorem (%)\"\n"
    b"    cboTipo.AddItem \"2 - Ad Rem (R$/unid)\"\n"
    b"    cboTipo.AddItem \"3 - Misto (% + R$/unid)\"\n",
    b"    With lstRegistros\n"
    b"        .Cols = 9\n"
    b"        .ColWidth(0) = 0\n"
    b"        .ColWidth(1) = 1020\n"
    b"        .ColWidth(2) = 4500\n"
    b"        .ColWidth(3) = 1500\n"
    b"        .ColWidth(4) = 1080\n"
    b"        .ColWidth(5) = 1200\n"
    b"        .ColWidth(6) = 1500\n"
    b"        .ColWidth(7) = 1500\n"
    b"        .ColWidth(8) = 0\n"
    b"        .Row = 0: .Col = 0: .Text = \"Id\"\n"
    b"        .Row = 0: .Col = 1: .Text = \"C\xf3d. IS\"\n"
    b"        .Row = 0: .Col = 2: .Text = \"Descri\xe7\xe3o\"\n"
    b"        .Row = 0: .Col = 3: .Text = \"Tipo\"\n"
    b"        .Row = 0: .Col = 4: .Text = \"Al\xedq. %\"\n"
    b"        .Row = 0: .Col = 5: .Text = \"Val. Unid\"\n"
    b"        .Row = 0: .Col = 6: .Text = \"Vig. Ini\"\n"
    b"        .Row = 0: .Col = 7: .Text = \"Vig. Fim\"\n"
    b"        .AllowUserResizing = 1\n"
    b"        .SelectionMode = 1\n"
    b"    End With\n"
    b"    cboTipo.Clear\n"
    b"    cboTipo.AddItem \"1 - Ad Valorem (%)\"\n"
    b"    cboTipo.AddItem \"2 - Ad Rem (R$/unid)\"\n"
    b"    cboTipo.AddItem \"3 - Misto (% + R$/unid)\"\n"
    b"    cboISCSTEdit.AddItem \"00 - Incid\xeancia na Origem\"\n"
    b"    cboISCSTEdit.AddItem \"01 - Sa\xedda Tributada\"\n"
    b"    cboISCSTEdit.AddItem \"99 - Outras Opera\xe7\xf5es\"\n",
    "Form_Load grid+combo"
)

# 3. CarregarRegistros: busca ISCST, grava col 8
patch(
    b"    sql = \"SELECT Id, cClassTrib_IS, Descricao, tipo_calculo_is, ISpAliq, ISvUnid, dIniVig, dFimVig FROM tbISClassTrib\"\n",
    b"    sql = \"SELECT Id, cClassTrib_IS, Descricao, tipo_calculo_is, ISpAliq, ISvUnid, dIniVig, dFimVig, ISNULL(ISCST,'01') AS ISCST FROM tbISClassTrib\"\n",
    "CarregarRegistros SQL"
)

patch(
    b"        lstRegistros.Row = r: lstRegistros.Col = 7: lstRegistros.Text = IIf(IsNull(rRs(\"dFimVig\")), \"\", Format(rRs(\"dFimVig\"), \"dd/mm/yyyy\"))\n",
    b"        lstRegistros.Row = r: lstRegistros.Col = 7: lstRegistros.Text = IIf(IsNull(rRs(\"dFimVig\")), \"\", Format(rRs(\"dFimVig\"), \"dd/mm/yyyy\"))\n"
    b"        lstRegistros.Row = r: lstRegistros.Col = 8: lstRegistros.Text = rRs(\"ISCST\")\n",
    "CarregarRegistros col8"
)

# 4. lstRegistros_Click: popula cboISCSTEdit a partir do col 8
patch(
    b"    txtDFim.Text = lstRegistros.TextMatrix(lstRegistros.Row, 7)\n"
    b"    DefinirEstado \"selecionado\"\n",
    b"    txtDFim.Text = lstRegistros.TextMatrix(lstRegistros.Row, 7)\n"
    b"    SelecionarISCST lstRegistros.TextMatrix(lstRegistros.Row, 8)\n"
    b"    DefinirEstado \"selecionado\"\n",
    "lstRegistros_Click cboISCSTEdit"
)

# 5. Adicionar helper SelecionarISCST antes de DefinirEstado
patch(
    b"Private Sub DefinirEstado(sEstado As String)\n",
    b"Private Sub SelecionarISCST(sCSTVal As String)\n"
    b"    Dim i As Integer\n"
    b"    For i = 0 To cboISCSTEdit.ListCount - 1\n"
    b"        If Left(cboISCSTEdit.List(i), 2) = sCSTVal Then\n"
    b"            cboISCSTEdit.ListIndex = i: Exit Sub\n"
    b"        End If\n"
    b"    Next i\n"
    b"    cboISCSTEdit.ListIndex = -1\n"
    b"End Sub\n\n"
    b"Private Sub DefinirEstado(sEstado As String)\n",
    "SelecionarISCST helper"
)

# 6. cmdSalvar_Click: capturar ISCST e incluir no SQL
patch(
    b"    Dim sCod As String, sDesc As String, iTipo As Integer\n"
    b"    Dim sAliq As String, sUnid As String, sDIni As String, sDFim As String\n",
    b"    Dim sCod As String, sDesc As String, iTipo As Integer\n"
    b"    Dim sAliq As String, sUnid As String, sDIni As String, sDFim As String\n"
    b"    Dim sISCST As String\n",
    "cmdSalvar Dim sISCST"
)

patch(
    b"    sDIni = Trim(txtDIni.Text)\n"
    b"    sDFim = Trim(txtDFim.Text)\n"
    b"    If sCod = \"\" Or sDesc = \"\" Or sDIni = \"\" Then\n",
    b"    sDIni = Trim(txtDIni.Text)\n"
    b"    sDFim = Trim(txtDFim.Text)\n"
    b"    sISCST = IIf(cboISCSTEdit.ListIndex >= 0, Left(cboISCSTEdit.Text, 2), \"01\")\n"
    b"    If sCod = \"\" Or sDesc = \"\" Or sDIni = \"\" Then\n",
    "cmdSalvar captura sISCST"
)

patch(
    b"        sql = \"INSERT INTO tbISClassTrib \" & _\n"
    b"              \"(cClassTrib_IS, Descricao, tipo_calculo_is, ISpAliq, ISvUnid, dIniVig, dFimVig) \" & _\n"
    b"              \"VALUES ('\" & sCod & \"','\" & Replace(sDesc, \"'\", \"''\") & \"',\" & _\n"
    b"              iTipo & \",\" & sAliqSQL & \",\" & sUnidSQL & \",\" & _\n"
    b"              \"CONVERT(date,'\" & sDIni & \"',103),\" & sDFimSQL & \")\"\n",
    b"        sql = \"INSERT INTO tbISClassTrib \" & _\n"
    b"              \"(cClassTrib_IS, Descricao, tipo_calculo_is, ISpAliq, ISvUnid, dIniVig, dFimVig, ISCST) \" & _\n"
    b"              \"VALUES ('\" & sCod & \"','\" & Replace(sDesc, \"'\", \"''\") & \"',\" & _\n"
    b"              iTipo & \",\" & sAliqSQL & \",\" & sUnidSQL & \",\" & _\n"
    b"              \"CONVERT(date,'\" & sDIni & \"',103),\" & sDFimSQL & \",'\" & sISCST & \"')\"\n",
    "cmdSalvar INSERT"
)

patch(
    b"        sql = \"UPDATE tbISClassTrib SET \" & _\n"
    b"              \"cClassTrib_IS='\" & sCod & \"',\" & _\n"
    b"              \"Descricao='\" & Replace(sDesc, \"'\", \"''\") & \"',\" & _\n"
    b"              \"tipo_calculo_is=\" & iTipo & \",\" & _\n"
    b"              \"ISpAliq=\" & sAliqSQL & \",\" & _\n"
    b"              \"ISvUnid=\" & sUnidSQL & \",\" & _\n"
    b"              \"dIniVig=CONVERT(date,'\" & sDIni & \"',103),\" & _\n"
    b"              \"dFimVig=\" & sDFimSQL & \" WHERE Id=\" & iId\n",
    b"        sql = \"UPDATE tbISClassTrib SET \" & _\n"
    b"              \"cClassTrib_IS='\" & sCod & \"',\" & _\n"
    b"              \"Descricao='\" & Replace(sDesc, \"'\", \"''\") & \"',\" & _\n"
    b"              \"tipo_calculo_is=\" & iTipo & \",\" & _\n"
    b"              \"ISpAliq=\" & sAliqSQL & \",\" & _\n"
    b"              \"ISvUnid=\" & sUnidSQL & \",\" & _\n"
    b"              \"dIniVig=CONVERT(date,'\" & sDIni & \"',103),\" & _\n"
    b"              \"dFimVig=\" & sDFimSQL & \",\" & _\n"
    b"              \"ISCST='\" & sISCST & \"' WHERE Id=\" & iId\n",
    "cmdSalvar UPDATE"
)

# 7. LimparCampos: resetar cboISCSTEdit
patch(
    b"    txtDFim.Text = \"\"\n"
    b"End Sub\n",
    b"    txtDFim.Text = \"\"\n"
    b"    cboISCSTEdit.ListIndex = -1\n"
    b"End Sub\n",
    "LimparCampos"
)

print(f"\nTotal erros: {errors}")
data = data.replace(b"\n", b"\r\n")
with open(FILE, "wb") as f:
    f.write(data)
print("Arquivo gravado")
