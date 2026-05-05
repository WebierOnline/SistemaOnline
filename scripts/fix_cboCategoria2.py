import sys

path = 'C:/Projeto/Compartilhado/Forms/Produtos_Cadastro.frm'
f = open(path, 'rb')
raw = f.read()
f.close()
text = raw.decode('windows-1252')

old = ('    \' Combos de Texto Simples\r\n'
       '    cboFabricante.Text = ValidateNull(r("fabricante"))\r\n'
       '    cboCategoria.Text = ValidateNull(r("categoria"))\r\n'
       '    SelecionarNoCombo cboUnidMedida, ValidateNull(r("unid_medida"))')

new = ('    \' Combos de Texto Simples\r\n'
       '    cboFabricante.Text = ValidateNull(r("fabricante"))\r\n'
       '    \' Popula cboCategoria pela tabela Categorias e seleciona o valor salvo\r\n'
       '    Dim rCat As ADODB.Recordset\r\n'
       '    cboCategoria.Clear\r\n'
       '    Set rCat = dbData.OpenRecordset("SELECT Categoria FROM Categorias WHERE Tipo_Empresa = " & tipoEmpresa & " ORDER BY Categoria")\r\n'
       '    Do While Not rCat.EOF\r\n'
       '        cboCategoria.AddItem ValidateNull(rCat("Categoria"))\r\n'
       '        rCat.MoveNext\r\n'
       '    Loop\r\n'
       '    If rCat.State <> 0 Then rCat.Close\r\n'
       '    SelecionarNoCombo cboCategoria, ValidateNull(r("categoria"))\r\n'
       '    SelecionarNoCombo cboUnidMedida, ValidateNull(r("unid_medida"))')

cnt = text.count(old)
if cnt != 1:
    print('ERRO: encontrado {}x'.format(cnt))
    sys.exit(1)

text = text.replace(old, new)
print('OK')

out = text.encode('windows-1252')
out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(path, 'wb').write(out)
print('Produtos_Cadastro.frm - CONCLUIDO')
