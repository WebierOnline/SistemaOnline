import sys

path = 'C:/Projeto/Compartilhado/Forms/Produtos_Cadastro.frm'
f = open(path, 'rb')
raw = f.read()
f.close()
text = raw.decode('windows-1252')

errors = []
def apply(text, old, new, tag):
    cnt = text.count(old)
    if cnt != 1:
        errors.append('ERRO {}: encontrado {}x'.format(tag, cnt))
        return text
    print('OK {}'.format(tag))
    return text.replace(old, new)

# 1. Style=2 no cboCategoria (dropdown list - nao permite digitar)
text = apply(text,
    'Begin VB.ComboBox cboCategoria \r\n            BackColor       =   &H00C0FFFF&\r\n            Height          =   315\r\n            Left            =   120\r\n            TabIndex        =   7\r\n            Top             =   1140\r\n            Width           =   1755\r\n         End',
    'Begin VB.ComboBox cboCategoria \r\n            BackColor       =   &H00C0FFFF&\r\n            Height          =   315\r\n            Left            =   120\r\n            Style           =   2  \'Dropdown List\r\n            TabIndex        =   7\r\n            Top             =   1140\r\n            Width           =   1755\r\n         End',
    '1-style-dropdown')

# 2. cboCategoria_GotFocus: trocar SELECT de produtos para Categorias filtrado por tipoEmpresa
old2 = ('Private Sub cboCategoria_GotFocus()\r\n'
        'Dim sSQL As String\r\n'
        'Dim r As ADODB.Recordset\r\n'
        '\r\n'
        '\'Limpa a lista atual\r\n'
        'cboCategoria.Clear\r\n'
        '\r\n'
        'sSQL = "SELECT DISTINCT categoria FROM produtos ORDER BY categoria;"\r\n'
        'Set r = dbData.OpenRecordset(sSQL)\r\n'
        '\r\n'
        'Do While Not r.EOF\r\n'
        '   cboCategoria.AddItem ValidateNull(r("categoria"))\r\n'
        '   r.MoveNext\r\n'
        'Loop\r\n'
        '\r\n'
        'moCombo.AttachTo cboCategoria\r\n'
        'End Sub')
new2 = ('Private Sub cboCategoria_GotFocus()\r\n'
        'Dim sSQL As String\r\n'
        'Dim r As ADODB.Recordset\r\n'
        'cboCategoria.Clear\r\n'
        'sSQL = "SELECT Categoria FROM Categorias WHERE Tipo_Empresa = " & tipoEmpresa & " ORDER BY Categoria"\r\n'
        'Set r = dbData.OpenRecordset(sSQL)\r\n'
        'Do While Not r.EOF\r\n'
        '   cboCategoria.AddItem ValidateNull(r("Categoria"))\r\n'
        '   r.MoveNext\r\n'
        'Loop\r\n'
        'If r.State <> 0 Then r.Close\r\n'
        'End Sub')
text = apply(text, old2, new2, '2-gotfocus-categorias')

# 3. cboCategoria_KeyPress: bloquear digitacao
old3 = ('Private Sub cboCategoria_KeyPress(KeyAscii As Integer)\r\n'
        'KeyAscii = Asc(UCase(Chr(KeyAscii)))\r\n'
        'End Sub')
new3 = ('Private Sub cboCategoria_KeyPress(KeyAscii As Integer)\r\n'
        'KeyAscii = 0\r\n'
        'End Sub')
text = apply(text, old3, new3, '3-keypress-block')

# 4. ValidarCampos: adicionar validacao de cboCategoria antes do CFOP
old4 = ('    \' --- ABA PRINCIPAL / FISCAL ATUAL ---\r\n'
        '    \' Verifica CFOP\r\n'
        '    If cboCFOP.ListIndex = -1 Then')
new4 = ('    \' --- ABA PRINCIPAL / FISCAL ATUAL ---\r\n'
        '    \' Verifica Categoria\r\n'
        '    If Trim(cboCategoria.Text) = "" Then\r\n'
        '        SSTab1.Tab = 0\r\n'
        '        MsgBox "Selecione a Categoria do produto!", vbExclamation, "Aten\xe7\xe3o"\r\n'
        '        cboCategoria.SetFocus\r\n'
        '        Exit Function\r\n'
        '    End If\r\n'
        '\r\n'
        '    \' Verifica CFOP\r\n'
        '    If cboCFOP.ListIndex = -1 Then')
text = apply(text, old4, new4, '4-validar-categoria')

if errors:
    for e in errors:
        print(e)
    sys.exit(1)

out = text.encode('windows-1252')
out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(path, 'wb').write(out)
print('Produtos_Cadastro.frm - CONCLUIDO')
