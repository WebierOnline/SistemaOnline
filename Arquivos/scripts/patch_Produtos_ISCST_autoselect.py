path = r'C:\Projeto\Compartilhado\Forms\Produtos_Cadastro.frm'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

R = '\r\n'

# Auto-seleciona "99" quando ISCST esta vazio no banco (produtos antigos pre-migracao)
old = "    SelecionarNoCombo cboISCST, ValidateNull(r(\"ISCST\")), True"

new = (
    "    Dim sISCSTLoad As String" + R +
    "    sISCSTLoad = Trim(ValidateNull(r(\"ISCST\")))" + R +
    "    If sISCSTLoad = \"\" Then sISCSTLoad = \"99\"" + R +
    "    SelecionarNoCombo cboISCST, sISCSTLoad, True"
)

n = content.count(old)
if n != 1:
    print(f'ERRO: count={n}')
else:
    content = content.replace(old, new, 1)
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
