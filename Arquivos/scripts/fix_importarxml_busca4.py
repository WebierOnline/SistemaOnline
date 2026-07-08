path = 'C:/projeto/OnlineCommerce/Forms/Entrada_Estoque.Frm'
data = open(path, 'rb').read().decode('windows-1252')

old = (
    "          '--- Busca 3: Produtos.EAN por cEANTrib (vinculo pela unidade interna) ---\r\n"
    "          '   So vincula automaticamente se cEANTrib for diferente de cEAN.\r\n"
    "          '   Se cEANTrib = cEAN o fornecedor nao informou unidade separada: aguarda vinculo manual.\r\n"
    "          If CodProduto = 0 And Not Vazio(cEANTrib) And cEANTrib <> cEAN Then\r\n"
    "             CodProduto = SQLExecutaRetorno(\"SELECT Codigo FROM Produtos WHERE EAN = '\" & cEANTrib & \"'\", \"Codigo\", 0)\r\n"
    "          End If\r\n"
    "\r\n"
    "          '--- Upsert VinculoXMLProduto: todos os produtos da XML sao registrados ---\r\n"
)
new = (
    "          '--- Busca 3: Produtos.EAN por cEANTrib (vinculo pela unidade interna) ---\r\n"
    "          '   So vincula automaticamente se cEANTrib for diferente de cEAN.\r\n"
    "          '   Se cEANTrib = cEAN o fornecedor nao informou unidade separada: aguarda vinculo manual.\r\n"
    "          If CodProduto = 0 And Not Vazio(cEANTrib) And cEANTrib <> cEAN Then\r\n"
    "             CodProduto = SQLExecutaRetorno(\"SELECT Codigo FROM Produtos WHERE EAN = '\" & cEANTrib & \"'\", \"Codigo\", 0)\r\n"
    "          End If\r\n"
    "\r\n"
    "          '--- Busca 4: Produtos.EAN / Produtos.EANEmbalagem por cEAN ou cEANTrib ---\r\n"
    "          '   Fallback para quando cEAN = cEANTrib (produto sem embalagem separada)\r\n"
    "          '   e nenhuma busca anterior localizou o produto.\r\n"
    "          If CodProduto = 0 And Not Vazio(cEAN) And cEAN <> \"SEM GTIN\" Then\r\n"
    "             CodProduto = SQLExecutaRetorno( _\r\n"
    "                 \"SELECT TOP 1 Codigo FROM Produtos \" & _\r\n"
    "                 \"WHERE EAN = '\" & cEAN & \"' OR EANEmbalagem = '\" & cEAN & \"'\", \"Codigo\", 0)\r\n"
    "          End If\r\n"
    "          If CodProduto = 0 And Not Vazio(cEANTrib) And cEANTrib <> \"SEM GTIN\" And cEANTrib <> cEAN Then\r\n"
    "             CodProduto = SQLExecutaRetorno( _\r\n"
    "                 \"SELECT TOP 1 Codigo FROM Produtos \" & _\r\n"
    "                 \"WHERE EAN = '\" & cEANTrib & \"' OR EANEmbalagem = '\" & cEANTrib & \"'\", \"Codigo\", 0)\r\n"
    "          End If\r\n"
    "\r\n"
    "          '--- Upsert VinculoXMLProduto: todos os produtos da XML sao registrados ---\r\n"
)

print('found:', data.count(old))
data2 = data.replace(old, new, 1)
print('changed:', data2 != data)
raw = data2.encode('windows-1252')
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(path, 'wb').write(raw)
print('ok')
