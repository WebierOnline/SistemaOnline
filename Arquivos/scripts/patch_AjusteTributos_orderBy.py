path = r'C:\Projeto\OnlineCommerce\Forms\Produtos_AjusteTributos.frm'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

errors = []

def sub(label, old, new, c):
    n = c.count(old)
    if n != 1:
        errors.append(f'{label}: count={n}')
        return c
    print(f'{label} OK')
    return c.replace(old, new, 1)

# ── 1: Remover definicoes ckkORDLinha e ckkORDDesc ────────────────────────
content = sub('Remove definicao ckkORD*',
    '         End\r\n'
    '         Begin VB.CheckBox ckkORDLinha \r\n'
    '            Caption         =   "Categoria"\r\n'
    '            Height          =   195\r\n'
    '            Left            =   120\r\n'
    '            TabIndex        =   47\r\n'
    '            Top             =   360\r\n'
    '            Width           =   1035\r\n'
    '         End\r\n'
    '         Begin VB.CheckBox ckkORDDesc \r\n'
    '            Caption         =   "Descri\xe7\xe3o"\r\n'
    '            Height          =   195\r\n'
    '            Left            =   120\r\n'
    '            TabIndex        =   46\r\n'
    '            Top             =   180\r\n'
    '            Value           =   1  \'Checked\r\n'
    '            Width           =   1035\r\n'
    '         End\r\n'
    '      End\r\n'
    '      Begin VB.Frame Frame2 ',

    '         End\r\n'
    '      End\r\n'
    '      Begin VB.Frame Frame2 ',
    content)

# ── 2: Substituir logica var_Indice em MostrarCriterios ───────────────────
content = sub('var_Indice nova logica optORD*',
    '   var_Indice = ""\r\n'
    '   var_Indice = var_Indice & IIf(ckkORDDesc.Value, IIf(var_Indice <> "", ", ", "") & "produtos.descricao", "")\r\n'
    '   var_Indice = var_Indice & IIf(ckkORDLinha.Value, IIf(var_Indice <> "", ", ", "") & "produtos.categoria", "")\r\n'
    '   \r\n'
    '   If var_Indice <> "" Then var_Indice = " ORDER BY " & var_Indice',

    '   If optORDDesc.Value Then\r\n'
    '      var_Indice = " ORDER BY produtos.descricao"\r\n'
    '   ElseIf optORDCategoria.Value Then\r\n'
    '      var_Indice = " ORDER BY produtos.categoria"\r\n'
    '   ElseIf optORDTags.Value Then\r\n'
    '      var_Indice = " ORDER BY produtos.TAGS"\r\n'
    '   ElseIf optORDNCM.Value Then\r\n'
    '      var_Indice = " ORDER BY produtos.NCM"\r\n'
    '   Else\r\n'
    '      var_Indice = ""\r\n'
    '   End If',
    content)

# ── 3: Substituir subs ckkORD* pelos novos optORD*_Click ─────────────────
content = sub('Substituir subs ckkORD* por optORD*',
    'Private Sub ckkORDDesc_Click()\r\n'
    '   MostrarCriterios\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub ckkORDLinha_Click()\r\n'
    '   MostrarCriterios\r\n'
    'End Sub',

    'Private Sub optORDDesc_Click()\r\n'
    '   MostrarCriterios\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optORDCategoria_Click()\r\n'
    '   MostrarCriterios\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optORDTags_Click()\r\n'
    '   MostrarCriterios\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optORDNCM_Click()\r\n'
    '   MostrarCriterios\r\n'
    'End Sub',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK — arquivo gravado.')
