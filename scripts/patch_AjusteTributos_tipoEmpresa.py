path = r'C:\Projeto\OnlineCommerce\Forms\Produtos_AjusteTributos.frm'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

errors = []

def sub(label, old, new, c, replace_all=False):
    n = c.count(old)
    if replace_all:
        if n == 0:
            errors.append(f'{label}: count=0')
            return c
        print(f'{label} OK ({n}x)')
        return c.replace(old, new)
    if n != 1:
        errors.append(f'{label}: count={n}')
        return c
    print(f'{label} OK')
    return c.replace(old, new, 1)

# ── 1: Declarar tipoEmpresa no nivel de modulo ────────────────────────────
content = sub('Declarar tipoEmpresa',
    'Private iRow As Long, iCol As Long\r\n',
    'Private iRow As Long, iCol As Long\r\n'
    'Private tipoEmpresa As Long\r\n',
    content)

# ── 2: Carregar tipoEmpresa no Form_Load ──────────────────────────────────
content = sub('Form_Load tipoEmpresa',
    'Private Sub Form_Load()\r\n'
    'Set moCombo = New cComboHelper\r\n'
    'End Sub',

    'Private Sub Form_Load()\r\n'
    'Set moCombo = New cComboHelper\r\n'
    'tipoEmpresa = CLng(sysConfig("TIPO_EMPRESA").Value)\r\n'
    'End Sub',
    content)

# ── 3: Filtrar SELECT Categorias por Tipo_Empresa (3 ocorrencias) ─────────
content = sub('SELECT Categorias filtro Tipo_Empresa',
    '"SELECT DISTINCT Categoria FROM Categorias ORDER BY Categoria;"',
    '"SELECT DISTINCT Categoria FROM Categorias WHERE Tipo_Empresa = " & tipoEmpresa & " ORDER BY Categoria;"',
    content, replace_all=True)

# ── 4: Filtrar SELECT Tags por c.Tipo_Empresa (2 ocorrencias) ────────────
content = sub('SELECT Tags filtro Tipo_Empresa',
    '"SELECT ct.Tags FROM Categorias_Tags ct INNER JOIN Categorias c ON ct.ID_Categoria = c.ID_Categoria ORDER BY c.Categoria, ct.Tags;"',
    '"SELECT ct.Tags FROM Categorias_Tags ct INNER JOIN Categorias c ON ct.ID_Categoria = c.ID_Categoria WHERE c.Tipo_Empresa = " & tipoEmpresa & " ORDER BY c.Categoria, ct.Tags;"',
    content, replace_all=True)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK — arquivo gravado.')
