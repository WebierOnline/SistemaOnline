path = r'C:\Projeto\Compartilhado\Forms\Produtos_Cadastro.frm'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

errors = []
R = '\r\n'

# caracteres especiais Windows-1252
DESC  = 'DESCRI\xc7\xc3O'    # DESCRICAO com C cedilha + A til
COD   = 'C\xd3D. BARRA'       # COD com O agudo
REF   = 'REFER\xcaNCIA'       # REFERENCIA com E circunflexo

def sub(label, old, new, c):
    n = c.count(old)
    if n != 1:
        errors.append(f'{label}: count={n}')
        return c
    print(f'{label} OK')
    return c.replace(old, new, 1)

# ── 1: cboCriterios_LostFocus — setar optSitTodos ou optHabilitado ────────────
content = sub('cboCriterios_LostFocus opt',
    '    cboConsProduto.SetFocus' + R +
    'End If' + R +
    'End Sub',

    '    cboConsProduto.SetFocus' + R +
    'End If' + R +
    'If cboCriterios.Text = "TODOS" Then' + R +
    '    optHabilitado.Value = True' + R +
    'Else' + R +
    '    optSitTodos.Value = True' + R +
    'End If' + R +
    'End Sub',
    content)

# ── 2: varProdutoHabilitado — estender filtro ativo para DESCRICAO/COD/CAT/REF ─
content = sub('varProdutoHabilitado extend',
    'If cboCriterios.Text = "TODOS" Then' + R +
    '    If optDesabilitados.Value = True Then' + R +
    '        varProdutoHabilitado = "produtos.ativo = 0 and "' + R +
    '    ElseIf optHabilitado.Value = True Then' + R +
    '        varProdutoHabilitado = "produtos.ativo = 1 and "' + R +
    '    End If' + R +
    'Else' + R +
    '        varProdutoHabilitado = "produtos.codigo > 1 and "' + R +
    'End If',

    'If cboCriterios.Text = "TODOS" Or cboCriterios.Text = "' + DESC + '" Or _' + R +
    '   cboCriterios.Text = "' + COD + '" Or cboCriterios.Text = "CATEGORIA" Or _' + R +
    '   cboCriterios.Text = "' + REF + '" Then' + R +
    '    If optDesabilitados.Value = True Then' + R +
    '        varProdutoHabilitado = "produtos.ativo = 0 and "' + R +
    '    ElseIf optHabilitado.Value = True Then' + R +
    '        varProdutoHabilitado = "produtos.ativo = 1 and "' + R +
    '    Else' + R +
    '        varProdutoHabilitado = "produtos.codigo > 1 and "' + R +
    '    End If' + R +
    'Else' + R +
    '        varProdutoHabilitado = "produtos.codigo > 1 and "' + R +
    'End If',
    content)

# ── 3: optSitTodos_Click — recarregar grid e ajustar caption ──────────────────
content = sub('optSitTodos_Click novo evento',
    'Private Sub optHabilitado_Click()' + R +
    'cmdExibir_Click' + R +
    'If optHabilitado.Value = True Then' + R +
    '    cmdDesativar.Caption = "Desativar"' + R +
    'ElseIf optDesabilitados.Value = True Then' + R +
    '    cmdDesativar.Caption = "Ativar"' + R +
    'Else' + R +
    '    cmdDesativar.Caption = "Desativar"' + R +
    'End If' + R +
    'End Sub',

    'Private Sub optHabilitado_Click()' + R +
    'cmdExibir_Click' + R +
    'If optHabilitado.Value = True Then' + R +
    '    cmdDesativar.Caption = "Desativar"' + R +
    'ElseIf optDesabilitados.Value = True Then' + R +
    '    cmdDesativar.Caption = "Ativar"' + R +
    'Else' + R +
    '    cmdDesativar.Caption = "Desativar"' + R +
    'End If' + R +
    'End Sub' + R +
    R +
    'Private Sub optSitTodos_Click()' + R +
    'cmdExibir_Click' + R +
    'cmdDesativar.Caption = "Desativar"' + R +
    'End Sub',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
