data = open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'rb').read()
text = data.decode('windows-1252')

# ── Patch A: remover CellForeColor do Do While loop ──────────────────────────
old_a = (
    '            .TextMatrix(.rows - 1, 2) = rTabela("status")\r\n'
    '            .Row = .rows - 1: .Col = 2\r\n'
    '            If rTabela("status") = "SIM" Then\r\n'
    '                .CellForeColor = RGB(139, 0, 0)\r\n'
    '            Else\r\n'
    '                .CellForeColor = vbBlack\r\n'
    '            End If\r\n'
    '            .TextMatrix(.rows - 1, 3) = Format(rTabela("var_dtcompra"), "dd/mm/yy")'
)
new_a = (
    '            .TextMatrix(.rows - 1, 2) = rTabela("status")\r\n'
    '            .TextMatrix(.rows - 1, 3) = Format(rTabela("var_dtcompra"), "dd/mm/yy")'
)

# ── Patch B: adicionar CellForeColor no segundo For loop (pos-carregamento) ──
old_b = (
    '      For i = 1 To .rows - 1\r\n'
    '         .Row = i\r\n'
    '         .Col = 5\r\n'
    '         .CellForeColor = &HC0&\r\n'
    '         .CellFontBold = True\r\n'
    '      Next'
)
new_b = (
    '      For i = 1 To .rows - 1\r\n'
    '         .Row = i\r\n'
    '         .Col = 5\r\n'
    '         .CellForeColor = &HC0&\r\n'
    '         .CellFontBold = True\r\n'
    '         .Col = 2\r\n'
    '         If .TextMatrix(i, 2) = "SIM" Then\r\n'
    '             .CellForeColor = RGB(139, 0, 0)\r\n'
    '         Else\r\n'
    '             .CellForeColor = vbBlack\r\n'
    '         End If\r\n'
    '      Next'
)

ca = text.count(old_a)
cb = text.count(old_b)
print(f'Patch A: {ca}  Patch B: {cb}')

if ca == 1 and cb == 1:
    text = text.replace(old_a, new_a)
    text = text.replace(old_b, new_b)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'wb').write(out)
    print('OK — CellForeColor movido para o segundo For loop')
else:
    print('ABORTADO')
