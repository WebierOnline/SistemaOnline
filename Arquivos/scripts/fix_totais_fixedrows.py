path = 'C:/projeto/OnlineCommerce/Forms/Entrada_Estoque.Frm'
data = open(path, 'rb').read().decode('windows-1252')

# R1: adiciona .FixedRows = 0 antes de .Rows = 1 e remove .FixedRows = 1 do bloco With
old1 = (
    "      .Redraw = False\r\n"
    "      .Rows = 1\r\n"
    "      .Cols = 3\r\n"
    "      .FixedRows = 1\r\n"
    "      .FixedCols = 0\r\n"
)
new1 = (
    "      .Redraw = False\r\n"
    "      .FixedRows = 0\r\n"
    "      .Rows = 1\r\n"
    "      .Cols = 3\r\n"
    "      .FixedCols = 0\r\n"
)

# R2: define FixedRows = 1 antes do Redraw = True final
old2 = (
    "   grdTotaisNota.Redraw = True\r\n"
    "End Sub\r\n"
)
new2 = (
    "   grdTotaisNota.FixedRows = 1\r\n"
    "   grdTotaisNota.Redraw = True\r\n"
    "End Sub\r\n"
)

print('r1 found:', data.count(old1))
print('r2 found:', data.count(old2))

data2 = data.replace(old1, new1, 1)
data2 = data2.replace(old2, new2, 1)

print('changed:', data2 != data)
raw = data2.encode('windows-1252')
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(path, 'wb').write(raw)
print('ok')
