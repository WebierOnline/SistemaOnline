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

# ── 1: Inserir bloco de validação antes do loop principal ─────────────────
validation = (
    'Dim iExact As Integer\r\n'
    'iExact = 0\r\n'
    'If optENCM.Value Then iExact = 8\r\n'
    'If optECFOP.Value Then iExact = 4\r\n'
    'If optEICMS.Value Then iExact = 3\r\n'
    'If optECest.Value Then iExact = 8\r\n'
    'If optECBS.Value Then iExact = 6\r\n'
    'If optEIS.Value Then iExact = 6\r\n'
    'If iExact > 0 Then\r\n'
    '   Dim k As Integer\r\n'
    '   Dim bNumOk As Boolean\r\n'
    '   bNumOk = (Len(sVal) = iExact)\r\n'
    '   If bNumOk Then\r\n'
    '      For k = 1 To Len(sVal)\r\n'
    '         If Mid(sVal, k, 1) < "0" Or Mid(sVal, k, 1) > "9" Then bNumOk = False: Exit For\r\n'
    '      Next k\r\n'
    '   End If\r\n'
    '   If Not bNumOk Then\r\n'
    '      MsgBox "Digite exatamente " & iExact & " d\xedgitos num\xe9ricos.", vbExclamation, "Edi\xe7\xe3o Coletiva"\r\n'
    '      txtEdicaoColetiva.SetFocus\r\n'
    '      Exit Sub\r\n'
    '   End If\r\n'
    'End If\r\n'
    '\r\n'
)

content = sub('Inserir validacao numerica',
    'If Len(Trim(sSet)) = 0 Then Exit Sub\r\n'
    '\r\n'
    'Dim bTemMarcadas As Boolean',

    'If Len(Trim(sSet)) = 0 Then Exit Sub\r\n'
    '\r\n'
    + validation +
    'Dim bTemMarcadas As Boolean',
    content)

# ── 2: Após ResetarMarcas, resetar optE* e ocultar controles ─────────────
reset_hide = (
    'optENCM.Value = False\r\n'
    'optECFOP.Value = False\r\n'
    'optEICMS.Value = False\r\n'
    'optECest.Value = False\r\n'
    'optECategoria.Value = False\r\n'
    'optETags.Value = False\r\n'
    'optECBS.Value = False\r\n'
    'optEIS.Value = False\r\n'
    'txtEdicaoColetiva.Visible = False\r\n'
    'cboEdicaoColetiva.Visible = False\r\n'
    'cmdEdicaoColetiva.Visible = False\r\n'
    'lblEdicaoColetiva.Visible = False\r\n'
)

content = sub('Reset optE* e ocultar apos edicao',
    'ProxCat:\r\n'
    '   Next j\r\n'
    'End If\r\n'
    '\r\n'
    'picAguarde.Visible = False\r\n'
    'ResetarMarcas\r\n'
    'End Sub',

    'ProxCat:\r\n'
    '   Next j\r\n'
    'End If\r\n'
    '\r\n'
    'picAguarde.Visible = False\r\n'
    'ResetarMarcas\r\n'
    + reset_hide +
    'End Sub',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK — arquivo gravado.')
