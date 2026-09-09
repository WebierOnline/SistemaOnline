import sys
p = r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm'
b = open(p, 'rb').read()

def rep(old, new, n=1):
    global b
    c = b.count(old)
    if c != n:
        print("ABORT ancora x%d (esperado %d): %r" % (c, n, old[:70])); sys.exit(1)
    b = b.replace(old, new, n)

# ---------- A) Form_Load: visibilidade do chkDadosVeiculo ----------
rep(b"optCodBarra.Value = True\r\n",
    b"optCodBarra.Value = True\r\n"
    b"chkDadosVeiculo.Visible = bUsaOS\r\n"
    b"chkDadosVeiculo.Value = 0\r\n")

# ---------- B) funcoes GetDadosVeiculoStr + ParDado (apos GetDadosBancariosStr) ----------
rep(b"    GetDadosBancariosStr = sResult\r\nEnd Function\r\n\r\n",
    b"    GetDadosBancariosStr = sResult\r\nEnd Function\r\n\r\n"
    b"Private Function ParDado(ByVal sRotulo As String, ByVal vVal As Variant) As String\r\n"
    b"    ' devolve  \", ROTULO: valor\"  se o valor nao for vazio/nulo; senao \"\" (campo pulado)\r\n"
    b"    Dim t As String\r\n"
    b"    t = Trim(IIf(IsNull(vVal), \"\", CStr(vVal)))\r\n"
    b"    If t = \"\" Then ParDado = \"\" Else ParDado = \", \" & sRotulo & \": \" & t\r\n"
    b"End Function\r\n"
    b"\r\n"
    b"Private Function GetDadosVeiculoStr() As String\r\n"
    b"    ' monta o texto \"Dados do Veiculo/Equipamento\" da OS ligada ao pedido de origem desta nota.\r\n"
    b"    Dim lPed As Long, lCodOS As Long, s As String\r\n"
    b"    Dim rE As ADODB.Recordset\r\n"
    b"    GetDadosVeiculoStr = \"\"\r\n"
    b"    If txtCodNota.Text = \"\" Then Exit Function\r\n"
    b"\r\n"
    b"    On Error Resume Next\r\n"
    b"    lPed = Val(SQLExecutaRetorno(\"SELECT TOP 1 ISNULL(Cod_Pedido,0) AS p FROM NotaFiscalItens WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ISNULL(Cod_Pedido,0) > 0\", \"p\", 0))\r\n"
    b"    On Error GoTo 0\r\n"
    b"    If lPed = 0 Then Exit Function\r\n"
    b"\r\n"
    b"    On Error Resume Next\r\n"
    b"    lCodOS = Val(SQLExecutaRetorno(\"SELECT TOP 1 COD_OS AS c FROM OS WHERE COD_PEDIDO = \" & lPed, \"c\", 0))\r\n"
    b"    On Error GoTo 0\r\n"
    b"    If lCodOS = 0 Then Exit Function\r\n"
    b"\r\n"
    b"    ' 1) veiculo: OS_Equipamento_Auto (Automoveis/Motocicletas/Recapadora)\r\n"
    b"    Set rE = Nothing\r\n"
    b"    On Error Resume Next\r\n"
    b"    Set rE = dbData.OpenRecordset(\"SELECT fabricante, modelo, ano, placa, km, cor, chassi FROM OS_Equipamento_Auto WHERE cod_os = \" & lCodOS)\r\n"
    b"    On Error GoTo 0\r\n"
    b"    If Not (rE Is Nothing) Then\r\n"
    b"        If Not rE.EOF Then\r\n"
    b"            s = \"Dados do Ve\" & Chr(237) & \"culo: COD_OS: \" & lCodOS\r\n"
    b"            s = s & ParDado(\"FABRICANTE\", rE(\"fabricante\"))\r\n"
    b"            s = s & ParDado(\"MODELO\", rE(\"modelo\"))\r\n"
    b"            s = s & ParDado(\"ANO\", rE(\"ano\"))\r\n"
    b"            s = s & ParDado(\"PLACA\", rE(\"placa\"))\r\n"
    b"            s = s & ParDado(\"KM\", rE(\"km\"))\r\n"
    b"            s = s & ParDado(\"COR\", rE(\"cor\"))\r\n"
    b"            s = s & ParDado(\"CHASSI\", rE(\"chassi\"))\r\n"
    b"            GetDadosVeiculoStr = s\r\n"
    b"            rE.Close: Set rE = Nothing\r\n"
    b"            Exit Function\r\n"
    b"        End If\r\n"
    b"        rE.Close: Set rE = Nothing\r\n"
    b"    End If\r\n"
    b"\r\n"
    b"    ' 2) equipamento: OS_Equipamento (Informatica/Celular/Climatizacao/etc.)\r\n"
    b"    Set rE = Nothing\r\n"
    b"    On Error Resume Next\r\n"
    b"    Set rE = dbData.OpenRecordset(\"SELECT fabricante, modelo, EQUIPAMENTO FROM OS_Equipamento WHERE cod_os = \" & lCodOS)\r\n"
    b"    On Error GoTo 0\r\n"
    b"    If Not (rE Is Nothing) Then\r\n"
    b"        If Not rE.EOF Then\r\n"
    b"            s = \"Dados do Equipamento: COD_OS: \" & lCodOS\r\n"
    b"            s = s & ParDado(\"EQUIPAMENTO\", rE(\"EQUIPAMENTO\"))\r\n"
    b"            s = s & ParDado(\"FABRICANTE\", rE(\"fabricante\"))\r\n"
    b"            s = s & ParDado(\"MODELO\", rE(\"modelo\"))\r\n"
    b"            GetDadosVeiculoStr = s\r\n"
    b"        End If\r\n"
    b"        rE.Close: Set rE = Nothing\r\n"
    b"    End If\r\n"
    b"End Function\r\n"
    b"\r\n")

# ---------- C) chkDadosVeiculo_Click (antes de cboFinalidade_Change) ----------
rep(b"    End If\r\nEnd Sub\r\n\r\nPrivate Sub cboFinalidade_Change()\r\n",
    b"    End If\r\nEnd Sub\r\n\r\n"
    b"Private Sub chkDadosVeiculo_Click()\r\n"
    b"    Static sVeicAdd As String\r\n"
    b"    Dim sVeic As String, sTmp As String, iPos As Integer\r\n"
    b"    If chkDadosVeiculo.Value = 1 Then\r\n"
    b"        sVeic = GetDadosVeiculoStr()\r\n"
    b"        If Len(sVeic) > 0 Then\r\n"
    b"            sVeicAdd = sVeic\r\n"
    b"            If Len(Trim(txtInfComple.Text)) > 0 Then\r\n"
    b"                txtInfComple.Text = txtInfComple.Text & vbCrLf & sVeic\r\n"
    b"            Else\r\n"
    b"                txtInfComple.Text = sVeic\r\n"
    b"            End If\r\n"
    b"        Else\r\n"
    b"            MsgBox \"N\" & Chr(227) & \"o encontrei OS ligada a esta nota (ou a OS n\" & Chr(227) & \"o tem dados de ve\" & Chr(237) & \"culo/equipamento cadastrados).\", vbInformation, \"Online Commerce\"\r\n"
    b"            chkDadosVeiculo.Value = 0\r\n"
    b"        End If\r\n"
    b"    Else\r\n"
    b"        If Len(sVeicAdd) > 0 Then\r\n"
    b"            sTmp = txtInfComple.Text\r\n"
    b"            iPos = InStr(sTmp, sVeicAdd)\r\n"
    b"            If iPos > 0 Then\r\n"
    b"                If iPos >= 3 And Mid(sTmp, iPos - 2, 2) = vbCrLf Then\r\n"
    b"                    sTmp = Left(sTmp, iPos - 3) & Mid(sTmp, iPos + Len(sVeicAdd))\r\n"
    b"                Else\r\n"
    b"                    sTmp = Left(sTmp, iPos - 1) & Mid(sTmp, iPos + Len(sVeicAdd))\r\n"
    b"                End If\r\n"
    b"                txtInfComple.Text = sTmp\r\n"
    b"            End If\r\n"
    b"            sVeicAdd = \"\"\r\n"
    b"        End If\r\n"
    b"    End If\r\n"
    b"End Sub\r\n"
    b"\r\n"
    b"Private Sub cboFinalidade_Change()\r\n")

# ---------- D) cmdConverterNFe_Click: auto-marca ao converter pedido com OS ----------
rep(b"    cmdRecalcular_Click\r\n    Frm_NF.Tab = 0\r\nElse\r\n",
    b"    cmdRecalcular_Click\r\n"
    b"    If chkDadosVeiculo.Visible Then\r\n"
    b"        chkDadosVeiculo.Value = 0\r\n"
    b"        If Len(GetDadosVeiculoStr()) > 0 Then chkDadosVeiculo.Value = 1\r\n"
    b"    End If\r\n"
    b"    Frm_NF.Tab = 0\r\nElse\r\n")

b = b.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(p, 'wb').write(b)
print("OK | ef bf bd:", b'\xef\xbf\xbd' in b, "| DESTINAT:", b'DESTINAT\xc1RIO' in b)
for tag in (b'Private Function GetDadosVeiculoStr', b'Private Sub chkDadosVeiculo_Click', b'chkDadosVeiculo.Visible = bUsaOS', b'If Len(GetDadosVeiculoStr()) > 0 Then chkDadosVeiculo.Value = 1'):
    print(tag.decode(), '->', b.count(tag))
