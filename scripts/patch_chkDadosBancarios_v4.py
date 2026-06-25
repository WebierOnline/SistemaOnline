"""
patch_chkDadosBancarios_v4.py

Corrige dois bugs no chkDadosBancarios:

1. AtualizarInfCompleCredSN (chamada pelo Salvar via AtualizarTotaisNota) sobrescrevia
   txtInfComple e apagava os dados bancarios. Fix: re-adiciona ao final se chk marcado.
   Extrai logica de leitura dos dados para GetDadosBancariosStr() reutilizavel.

2. LimparObjestosNotaOutros nao resetava chkDadosBancarios.
   Fix: adiciona chkDadosBancarios.Value = 0 ao final do sub.

3. Simplifica chkDadosBancarios_Click para usar GetDadosBancariosStr().
"""
import sys, shutil

FRM = r"C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm"
BAK = FRM + ".bak_chkDadosBancarios_v4"

shutil.copy2(FRM, BAK)
print(f"Backup: {BAK}")

with open(FRM, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

patches = []

# ---------------------------------------------------------------------------
# P1: AtualizarInfCompleCredSN — re-adiciona banco ao final + insere GetDadosBancariosStr
# ---------------------------------------------------------------------------
patches.append((
    b"    Else\r\n"
    b"        txtInfComple.Text = sBaseSimples\r\n"
    b"    End If\r\n"
    b"End Sub\r\n"
    b"\r\n"
    b"Private Sub CalcularICMSInterNota()\r\n",

    b"    Else\r\n"
    b"        txtInfComple.Text = sBaseSimples\r\n"
    b"    End If\r\n"
    b"    If chkDadosBancarios.Value = 1 Then\r\n"
    b"        Dim sBnk As String\r\n"
    b"        sBnk = GetDadosBancariosStr()\r\n"
    b"        If Len(sBnk) > 0 Then\r\n"
    b"            If Len(Trim(txtInfComple.Text)) > 0 Then\r\n"
    b"                txtInfComple.Text = txtInfComple.Text & vbCrLf & sBnk\r\n"
    b"            Else\r\n"
    b"                txtInfComple.Text = sBnk\r\n"
    b"            End If\r\n"
    b"        End If\r\n"
    b"    End If\r\n"
    b"End Sub\r\n"
    b"\r\n"
    b"Private Function GetDadosBancariosStr() As String\r\n"
    b"    Dim rBnk As ADODB.Recordset, sResult As String, sPrt As String\r\n"
    b"    RsOpen rBnk, \"SELECT Banco, Agencia, Conta, Tipo, Favorecido, Pix FROM Empresa\"\r\n"
    b"    If Not rBnk.EOF Then\r\n"
    b"        sResult = \"\"\r\n"
    b"        sPrt = Trim(IIf(IsNull(rBnk(\"Banco\")), \"\", rBnk(\"Banco\")))\r\n"
    b"        If Len(sPrt) > 0 Then sResult = sResult & IIf(Len(sResult) > 0, \", \", \"\") & \"Banco: \" & sPrt\r\n"
    b"        sPrt = Trim(IIf(IsNull(rBnk(\"Agencia\")), \"\", rBnk(\"Agencia\")))\r\n"
    b"        If Len(sPrt) > 0 Then sResult = sResult & IIf(Len(sResult) > 0, \", \", \"\") & \"Agencia: \" & sPrt\r\n"
    b"        sPrt = Trim(IIf(IsNull(rBnk(\"Conta\")), \"\", rBnk(\"Conta\")))\r\n"
    b"        If Len(sPrt) > 0 Then sResult = sResult & IIf(Len(sResult) > 0, \", \", \"\") & \"Conta: \" & sPrt\r\n"
    b"        sPrt = Trim(IIf(IsNull(rBnk(\"Tipo\")), \"\", rBnk(\"Tipo\")))\r\n"
    b"        If Len(sPrt) > 0 Then sResult = sResult & IIf(Len(sResult) > 0, \", \", \"\") & \"Tipo: \" & sPrt\r\n"
    b"        sPrt = Trim(IIf(IsNull(rBnk(\"Favorecido\")), \"\", rBnk(\"Favorecido\")))\r\n"
    b"        If Len(sPrt) > 0 Then sResult = sResult & IIf(Len(sResult) > 0, \", \", \"\") & \"Favorecido: \" & sPrt\r\n"
    b"        sPrt = Trim(IIf(IsNull(rBnk(\"Pix\")), \"\", rBnk(\"Pix\")))\r\n"
    b"        If Len(sPrt) > 0 Then sResult = sResult & IIf(Len(sResult) > 0, \", \", \"\") & \"Chave Pix: \" & sPrt\r\n"
    b"    End If\r\n"
    b"    rBnk.Close\r\n"
    b"    Set rBnk = Nothing\r\n"
    b"    GetDadosBancariosStr = sResult\r\n"
    b"End Function\r\n"
    b"\r\n"
    b"Private Sub CalcularICMSInterNota()\r\n",

    "P1: AtualizarInfCompleCredSN re-append + GetDadosBancariosStr"
))

# ---------------------------------------------------------------------------
# P2: chkDadosBancarios_Click — simplificar usando GetDadosBancariosStr
# ---------------------------------------------------------------------------
patches.append((
    b"Private Sub chkDadosBancarios_Click()\r\n"
    b"    Static sBancoAdicionado As String\r\n"
    b"    Dim rBanco As ADODB.Recordset\r\n"
    b"    Dim sBanco As String, sParte As String, sTmp As String, iPos As Integer\r\n"
    b"    If chkDadosBancarios.Value = 1 Then\r\n"
    b"        RsOpen rBanco, \"SELECT Banco, Agencia, Conta, Tipo, Favorecido, Pix FROM Empresa\"\r\n"
    b"        If Not rBanco.EOF Then\r\n"
    b"            sBanco = \"\"\r\n"
    b"            sParte = Trim(IIf(IsNull(rBanco(\"Banco\")), \"\", rBanco(\"Banco\")))\r\n"
    b"            If Len(sParte) > 0 Then sBanco = sBanco & IIf(Len(sBanco) > 0, \", \", \"\") & \"Banco: \" & sParte\r\n"
    b"            sParte = Trim(IIf(IsNull(rBanco(\"Agencia\")), \"\", rBanco(\"Agencia\")))\r\n"
    b"            If Len(sParte) > 0 Then sBanco = sBanco & IIf(Len(sBanco) > 0, \", \", \"\") & \"Agencia: \" & sParte\r\n"
    b"            sParte = Trim(IIf(IsNull(rBanco(\"Conta\")), \"\", rBanco(\"Conta\")))\r\n"
    b"            If Len(sParte) > 0 Then sBanco = sBanco & IIf(Len(sBanco) > 0, \", \", \"\") & \"Conta: \" & sParte\r\n"
    b"            sParte = Trim(IIf(IsNull(rBanco(\"Tipo\")), \"\", rBanco(\"Tipo\")))\r\n"
    b"            If Len(sParte) > 0 Then sBanco = sBanco & IIf(Len(sBanco) > 0, \", \", \"\") & \"Tipo: \" & sParte\r\n"
    b"            sParte = Trim(IIf(IsNull(rBanco(\"Favorecido\")), \"\", rBanco(\"Favorecido\")))\r\n"
    b"            If Len(sParte) > 0 Then sBanco = sBanco & IIf(Len(sBanco) > 0, \", \", \"\") & \"Favorecido: \" & sParte\r\n"
    b"            sParte = Trim(IIf(IsNull(rBanco(\"Pix\")), \"\", rBanco(\"Pix\")))\r\n"
    b"            If Len(sParte) > 0 Then sBanco = sBanco & IIf(Len(sBanco) > 0, \", \", \"\") & \"Chave Pix: \" & sParte\r\n"
    b"        End If\r\n"
    b"        rBanco.Close\r\n"
    b"        Set rBanco = Nothing\r\n"
    b"        If Len(sBanco) > 0 Then\r\n"
    b"            sBancoAdicionado = sBanco\r\n"
    b"            If Len(Trim(txtInfComple.Text)) > 0 Then\r\n"
    b"                txtInfComple.Text = txtInfComple.Text & vbCrLf & sBanco\r\n"
    b"            Else\r\n"
    b"                txtInfComple.Text = sBanco\r\n"
    b"            End If\r\n"
    b"        End If\r\n"
    b"    Else\r\n"
    b"        If Len(sBancoAdicionado) > 0 Then\r\n"
    b"            sTmp = txtInfComple.Text\r\n"
    b"            iPos = InStr(sTmp, sBancoAdicionado)\r\n"
    b"            If iPos > 0 Then\r\n"
    b"                If iPos >= 3 Then\r\n"
    b"                    If Mid(sTmp, iPos - 2, 2) = vbCrLf Then\r\n"
    b"                        txtInfComple.Text = Trim(Left(sTmp, iPos - 3))\r\n"
    b"                    Else\r\n"
    b"                        txtInfComple.Text = \"\"\r\n"
    b"                    End If\r\n"
    b"                Else\r\n"
    b"                    txtInfComple.Text = \"\"\r\n"
    b"                End If\r\n"
    b"            End If\r\n"
    b"            sBancoAdicionado = \"\"\r\n"
    b"        End If\r\n"
    b"    End If\r\n"
    b"End Sub\r\n",

    b"Private Sub chkDadosBancarios_Click()\r\n"
    b"    Static sBancoAdicionado As String\r\n"
    b"    Dim sBanco As String, sTmp As String, iPos As Integer\r\n"
    b"    If chkDadosBancarios.Value = 1 Then\r\n"
    b"        sBanco = GetDadosBancariosStr()\r\n"
    b"        If Len(sBanco) > 0 Then\r\n"
    b"            sBancoAdicionado = sBanco\r\n"
    b"            If Len(Trim(txtInfComple.Text)) > 0 Then\r\n"
    b"                txtInfComple.Text = txtInfComple.Text & vbCrLf & sBanco\r\n"
    b"            Else\r\n"
    b"                txtInfComple.Text = sBanco\r\n"
    b"            End If\r\n"
    b"        End If\r\n"
    b"    Else\r\n"
    b"        If Len(sBancoAdicionado) > 0 Then\r\n"
    b"            sTmp = txtInfComple.Text\r\n"
    b"            iPos = InStr(sTmp, sBancoAdicionado)\r\n"
    b"            If iPos > 0 Then\r\n"
    b"                If iPos >= 3 Then\r\n"
    b"                    If Mid(sTmp, iPos - 2, 2) = vbCrLf Then\r\n"
    b"                        txtInfComple.Text = Trim(Left(sTmp, iPos - 3))\r\n"
    b"                    Else\r\n"
    b"                        txtInfComple.Text = \"\"\r\n"
    b"                    End If\r\n"
    b"                Else\r\n"
    b"                    txtInfComple.Text = \"\"\r\n"
    b"                End If\r\n"
    b"            End If\r\n"
    b"            sBancoAdicionado = \"\"\r\n"
    b"        End If\r\n"
    b"    End If\r\n"
    b"End Sub\r\n",

    "P2: chkDadosBancarios_Click usa GetDadosBancariosStr"
))

# ---------------------------------------------------------------------------
# P3: LimparObjestosNotaOutros — resetar chkDadosBancarios
# ---------------------------------------------------------------------------
patches.append((
    b"txtChaveReferenciada.Text = \"\"\r\n"
    b"End Sub\r\n",

    b"txtChaveReferenciada.Text = \"\"\r\n"
    b"chkDadosBancarios.Value = 0\r\n"
    b"End Sub\r\n",

    "P3: LimparObjestosNotaOutros reset chkDadosBancarios"
))

# ---------------------------------------------------------------------------
errors = 0
for old, new, desc in patches:
    cnt = data.count(old)
    if cnt != 1:
        print(f"ERRO {desc}: count={cnt} (esperado 1)")
        errors += 1
    else:
        data = data.replace(old, new)
        print(f"OK   {desc}")

data = norm(data)

if errors:
    print(f"\n{errors} erro(s). Arquivo NAO foi salvo.")
    sys.exit(1)

with open(FRM, "wb") as f:
    f.write(data)
print("\nArquivo salvo com sucesso.")
