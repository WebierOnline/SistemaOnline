path = r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

R = '\r\n'

old = (
    "    Tb(\"ValorUnitarioComercializacao\") = CDbl(Format(txtValor, \"@\"))" + R +
    "    If txtQuant.Text <> \"\" Then Tb(\"QuantidadeComercial\") = CDbl(Format(txtQuant, \"@\")) Else Tb(\"QuantidadeComercial\") = Format(1, \"@\")" + R +
    "    Tb(\"ValorFrete\") = CDbl(IIf(txtFrete.Text = \"\", 0, Format(txtFrete, \"@\")))"
)

new = (
    "    Tb(\"ValorUnitarioComercializacao\") = CDbl(Format(txtValor, \"@\"))" + R +
    "    If txtQuant.Text <> \"\" Then Tb(\"QuantidadeComercial\") = CDbl(Format(txtQuant, \"@\")) Else Tb(\"QuantidadeComercial\") = Format(1, \"@\")" + R +
    "    Tb(\"EANTrib\") = Format(vEAN, \"@\")" + R +
    "    Tb(\"UnidadeTributavel\") = UCase(Format(vUnid_medida, \"@\"))" + R +
    "    If txtQuant.Text <> \"\" Then Tb(\"QuantidadeTributavel\") = CDbl(Format(txtQuant, \"@\")) Else Tb(\"QuantidadeTributavel\") = Format(1, \"@\")" + R +
    "    Tb(\"ValorUnitarioTributario\") = CDbl(Format(txtValor, \"@\"))" + R +
    "    Tb(\"ValorFrete\") = CDbl(IIf(txtFrete.Text = \"\", 0, Format(txtFrete, \"@\")))"
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
