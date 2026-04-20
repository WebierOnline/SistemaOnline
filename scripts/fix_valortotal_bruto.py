data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

replacements = [
    # 1. CalcularTotalProdutos
    (
        b'sSQL = "SELECT SUM(ValorUnitarioComercializacao * QuantidadeComercial) as ValorProdutosItens FROM NotaFiscalItens WHERE CodigoNota = "',
        b'sSQL = "SELECT SUM(ValorTotalBruto) as ValorProdutosItens FROM NotaFiscalItens WHERE CodigoNota = "'
    ),
    # 2. AtualizarTotaisNota
    (
        b'"ISNULL(SUM(ValorUnitarioComercializacao * QuantidadeComercial), 0) AS ValorProdutos, "',
        b'"ISNULL(SUM(ValorTotalBruto), 0) AS ValorProdutos, "'
    ),
    # 3. DistribuirDesconto
    (
        b'"ISNULL(SUM(ValorUnitarioComercializacao * QuantidadeComercial), 0) AS TotalSubtotal "',
        b'"ISNULL(SUM(ValorTotalBruto), 0) AS TotalSubtotal "'
    ),
    # 4. MostrarValorProdutos
    (
        b'sSQL = "SELECT SUM(ValorUnitarioComercializacao * QuantidadeComercial) as ValorProdutos FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)',
        b'sSQL = "SELECT SUM(ValorTotalBruto) as ValorProdutos FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)'
    ),
    # 5. MostrarValorBaseICMS
    (
        b'sSQL = "SELECT SUM(ValorUnitarioComercializacao * QuantidadeComercial) as vValorProdutos FROM NotaFiscalItens WHERE (vICMS <> \'0.00\') AND CodigoNota = "',
        b'sSQL = "SELECT SUM(ValorTotalBruto) as vValorProdutos FROM NotaFiscalItens WHERE (vICMS <> \'0.00\') AND CodigoNota = "'
    ),
]

for old, new in replacements:
    c = data.count(old)
    label = new[:60].decode('latin-1')
    print(f'count={c} | {label}')
    if c == 1:
        data = data.replace(old, new)
    elif c == 0:
        print('  *** NAO ENCONTRADO')
    else:
        print(f'  *** MULTIPLAS OCORRENCIAS - nao substituido')

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))
