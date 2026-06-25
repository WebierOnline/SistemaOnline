"""
patch_chkDadosBancarios_v2.py

Corrige chkDadosBancarios_Click:
- Campos vazios/null sao omitidos (sem legenda e sem valor)
- Se todos estiverem vazios, nao adiciona nada
- Static sBancoAdicionado guarda o que foi inserido para remocao exata ao desmarcar
"""
import sys, shutil

FRM = r"C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm"
BAK = FRM + ".bak_chkDadosBancarios_v2"

shutil.copy2(FRM, BAK)
print(f"Backup: {BAK}")

with open(FRM, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

OLD = (
    b"Private Sub chkDadosBancarios_Click()\r\n"
    b"    Dim rBanco As ADODB.Recordset\r\n"
    b"    Dim sBanco As String, sTmp As String, iPos As Integer\r\n"
    b"    If chkDadosBancarios.Value = 1 Then\r\n"
    b"        RsOpen rBanco, \"SELECT Banco, Agencia, Conta, Tipo, Favorecido, Pix FROM Empresa\"\r\n"
    b"        If Not rBanco.EOF Then\r\n"
    b"            sBanco = \"Banco: \" & IIf(IsNull(rBanco(\"Banco\")), \"\", rBanco(\"Banco\")) & _\r\n"
    b"                     \", Agencia: \" & IIf(IsNull(rBanco(\"Agencia\")), \"\", rBanco(\"Agencia\")) & _\r\n"
    b"                     \", Conta: \" & IIf(IsNull(rBanco(\"Conta\")), \"\", rBanco(\"Conta\")) & _\r\n"
    b"                     \", Tipo: \" & IIf(IsNull(rBanco(\"Tipo\")), \"\", rBanco(\"Tipo\")) & _\r\n"
    b"                     \", Favorecido: \" & IIf(IsNull(rBanco(\"Favorecido\")), \"\", rBanco(\"Favorecido\")) & _\r\n"
    b"                     \", Chave Pix: \" & IIf(IsNull(rBanco(\"Pix\")), \"\", rBanco(\"Pix\"))\r\n"
    b"        End If\r\n"
    b"        rBanco.Close\r\n"
    b"        Set rBanco = Nothing\r\n"
    b"        If Len(sBanco) > 0 Then\r\n"
    b"            If Len(Trim(txtInfComple.Text)) > 0 Then\r\n"
    b"                txtInfComple.Text = txtInfComple.Text & vbCrLf & sBanco\r\n"
    b"            Else\r\n"
    b"                txtInfComple.Text = sBanco\r\n"
    b"            End If\r\n"
    b"        End If\r\n"
    b"    Else\r\n"
    b"        sTmp = txtInfComple.Text\r\n"
    b"        iPos = InStr(sTmp, \"Banco:\")\r\n"
    b"        If iPos > 0 Then\r\n"
    b"            If iPos >= 3 And Mid(sTmp, iPos - 2, 2) = vbCrLf Then\r\n"
    b"                txtInfComple.Text = Trim(Left(sTmp, iPos - 3))\r\n"
    b"            Else\r\n"
    b"                txtInfComple.Text = \"\"\r\n"
    b"            End If\r\n"
    b"        End If\r\n"
    b"    End If\r\n"
    b"End Sub\r\n"
)

NEW = (
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
    b"                If iPos >= 3 And Mid(sTmp, iPos - 2, 2) = vbCrLf Then\r\n"
    b"                    txtInfComple.Text = Trim(Left(sTmp, iPos - 3))\r\n"
    b"                Else\r\n"
    b"                    txtInfComple.Text = \"\"\r\n"
    b"                End If\r\n"
    b"            End If\r\n"
    b"            sBancoAdicionado = \"\"\r\n"
    b"        End If\r\n"
    b"    End If\r\n"
    b"End Sub\r\n"
)

cnt = data.count(OLD)
if cnt != 1:
    print(f"ERRO: trecho encontrado {cnt}x (esperado 1). Arquivo NAO alterado.")
    sys.exit(1)

data = data.replace(OLD, NEW)
data = norm(data)

with open(FRM, "wb") as f:
    f.write(data)

print("OK: chkDadosBancarios_Click atualizado (campos dinamicos + Static)")
print("Arquivo salvo com sucesso.")
