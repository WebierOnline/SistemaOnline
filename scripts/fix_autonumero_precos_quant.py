path = 'C:/projeto/OnlineCommerce/Forms/Entrada_Estoque.Frm'
data = open(path, 'rb').read().decode('windows-1252')

# ---- R1: adiciona autonumeros de precos e quant antes do BEGIN TRANSACTION ----
old1 = (
"Dim var_COD_ITENS As Long\r\n"
"\r\n"
"'AUTONUMERAaaO\r\n"
"sSQL = \"SELECT ISNULL(MAX(codigo), 0) AS cod_itens FROM produtos_entrada_itens;\"\r\n"
"Set r = dbData.OpenRecordset(sSQL)\r\n"
"\r\n"
"If Not r.BOF Then var_COD_ITENS = r(\"cod_itens\") + 1\r\n"
"\r\n"
"If r.State <> 0 Then r.Close\r\n"
"Set r = Nothing\r\n"
)

new1 = (
"Dim var_COD_ITENS As Long\r\n"
"Dim var_COD_PRECOS As Long\r\n"
"Dim var_COD_QUANT As Long\r\n"
"\r\n"
"'AUTONUMERAaaO\r\n"
"sSQL = \"SELECT ISNULL(MAX(codigo), 0) AS cod_itens FROM produtos_entrada_itens;\"\r\n"
"Set r = dbData.OpenRecordset(sSQL)\r\n"
"If Not r.BOF Then var_COD_ITENS = r(\"cod_itens\") + 1\r\n"
"If r.State <> 0 Then r.Close\r\n"
"Set r = Nothing\r\n"
"\r\n"
"sSQL = \"SELECT ISNULL(MAX(codigo), 0) AS cod_itens FROM produtos_precos;\"\r\n"
"Set r = dbData.OpenRecordset(sSQL)\r\n"
"If Not r.BOF Then var_COD_PRECOS = r(\"cod_itens\") + 1\r\n"
"If r.State <> 0 Then r.Close\r\n"
"Set r = Nothing\r\n"
"\r\n"
"sSQL = \"SELECT ISNULL(MAX(codigo), 0) AS cod_itens FROM produtos_quant;\"\r\n"
"Set r = dbData.OpenRecordset(sSQL)\r\n"
"If Not r.BOF Then var_COD_QUANT = r(\"cod_itens\") + 1\r\n"
"If r.State <> 0 Then r.Close\r\n"
"Set r = Nothing\r\n"
)

# ---- R2: passa os autonumeros como parametros nas chamadas ----
old2 = (
"Preco_Entrada\r\n"
"Quant_Entrada\r\n"
)

new2 = (
"Preco_Entrada var_COD_PRECOS\r\n"
"Quant_Entrada var_COD_QUANT\r\n"
)

# ---- R3: remove autonumero de dentro de Preco_Entrada e usa parametro ----
old3 = (
"Private Sub Preco_Entrada()\r\n"
"Dim sSQL As String\r\n"
"Dim r As ADODB.Recordset\r\n"
"\r\n"
"'ENTRADA DO PRODUTO\r\n"
"'If cmdSalvar.Enabled = True Then\r\n"
"   Dim AutoNumeracao As Long\r\n"
"   \r\n"
"   'AUTONUMERAaaO\r\n"
"   sSQL = \"SELECT ISNULL(MAX(codigo), 0) AS cod_itens FROM produtos_precos;\"\r\n"
"   Set r = dbData.OpenRecordset(sSQL)\r\n"
"   \r\n"
"   If Not r.BOF Then AutoNumeracao = r(\"cod_itens\") + 1\r\n"
"   If r.State <> 0 Then r.Close\r\n"
"   Set r = Nothing\r\n"
)

new3 = (
"Private Sub Preco_Entrada(ByVal AutoNumeracao As Long)\r\n"
"Dim sSQL As String\r\n"
"\r\n"
"'ENTRADA DO PRODUTO\r\n"
"'If cmdSalvar.Enabled = True Then\r\n"
)

# ---- R4: remove autonumero de dentro de Quant_Entrada e usa parametro ----
old4 = (
"Private Sub Quant_Entrada()\r\n"
"Dim sSQL As String\r\n"
"Dim r As ADODB.Recordset\r\n"
"\r\n"
"'ENTRADA DO PRODUTO\r\n"
"'If cmdSalvar.Enabled = True Then\r\n"
"Dim AutoNumeracao As Long\r\n"
"\r\n"
"'AUTONUMERAaaO\r\n"
"sSQL = \"SELECT ISNULL(MAX(codigo), 0) AS cod_itens FROM produtos_quant;\"\r\n"
"Set r = dbData.OpenRecordset(sSQL)\r\n"
"\r\n"
"If Not r.BOF Then AutoNumeracao = r(\"cod_itens\") + 1\r\n"
"If r.State <> 0 Then r.Close\r\n"
"Set r = Nothing\r\n"
)

new4 = (
"Private Sub Quant_Entrada(ByVal AutoNumeracao As Long)\r\n"
"Dim sSQL As String\r\n"
"\r\n"
"'ENTRADA DO PRODUTO\r\n"
"'If cmdSalvar.Enabled = True Then\r\n"
)

print('r1 found:', data.count(old1))
print('r2 found:', data.count(old2))
print('r3 found:', data.count(old3))
print('r4 found:', data.count(old4))

data2 = data
data2 = data2.replace(old1, new1, 1)
data2 = data2.replace(old2, new2, 1)
data2 = data2.replace(old3, new3, 1)
data2 = data2.replace(old4, new4, 1)

print('changed:', data2 != data)
raw = data2.encode('windows-1252')
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(path, 'wb').write(raw)
print('ok')
