path = 'C:/projeto/OnlineCommerce/Forms/Entrada_Estoque.Frm'
data = open(path, 'rb').read().decode('windows-1252')

# ---- R1: oculta coluna ITEMXML ----
old1 = "    .ColWidth(17) = 500\r\n"
new1 = "    .ColWidth(17) = 0\r\n"

# ---- R2: adiciona fundo cinza claro nas colunas 15 e 16 ----
old2 = (
"        'COLUNA EM NEGRITO\r\n"
"         For i = 1 To .rows - 1\r\n"
"            .Row = i\r\n"
"            .Col = 11\r\n"
"            .CellFontBold = True\r\n"
"         Next\r\n"
"   \r\n"
"   .Redraw = True\r\n"
)
new2 = (
"        'COLUNA EM NEGRITO\r\n"
"         For i = 1 To .rows - 1\r\n"
"            .Row = i\r\n"
"            .Col = 11\r\n"
"            .CellFontBold = True\r\n"
"         Next\r\n"
"   \r\n"
"        'FUNDO CINZA CLARO: QTDE TRIB e VALOR VV\r\n"
"         For i = 1 To .rows - 1\r\n"
"            .Row = i\r\n"
"            .Col = 15\r\n"
"            .CellBackColor = &HE0E0E0\r\n"
"         Next\r\n"
"         For i = 1 To .rows - 1\r\n"
"            .Row = i\r\n"
"            .Col = 16\r\n"
"            .CellBackColor = &HE0E0E0\r\n"
"         Next\r\n"
"   \r\n"
"   .Redraw = True\r\n"
)

print('r1 found:', data.count(old1))
print('r2 found:', data.count(old2))

data2 = data
data2 = data2.replace(old1, new1, 1)
data2 = data2.replace(old2, new2, 1)

print('changed:', data2 != data)
raw = data2.encode('windows-1252')
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(path, 'wb').write(raw)
print('ok')
