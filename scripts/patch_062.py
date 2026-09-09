novo = (
"IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('pedidos_itens') AND name = 'Custo')\n"
"    ALTER TABLE pedidos_itens ADD Custo decimal(16,2) NULL DEFAULT 0;\n"
"GO\n"
"\n"
"IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('pedidos_itens') AND name = 'SUBTOTAL_CUSTO')\n"
"    ALTER TABLE pedidos_itens ADD SUBTOTAL_CUSTO decimal(16,2) NULL DEFAULT 0;\n"
"GO\n"
"\n"
"-- Preenche Custo/SUBTOTAL_CUSTO nos itens de pedido a partir do ultimo preco cadastrado.\n"
"-- Blindado contra dado corrompido (codigo de barras gravado no campo CUSTO de Produtos_Precos):\n"
"-- o valor que sai de Produtos_Precos passa por CONVERT(float,...) (nunca estoura) e so vira\n"
"-- decimal depois de cortado na faixa valida; item fora da faixa fica com custo 0.\n"
"\n"
"-- Passo 1: custo unitario\n"
"UPDATE PI\n"
"SET PI.Custo = CONVERT(decimal(12,2),\n"
"        CASE WHEN PP.v >= 0 AND PP.v < 1000000 THEN PP.v ELSE 0 END)\n"
"FROM pedidos_itens PI\n"
"CROSS APPLY (\n"
"    SELECT TOP 1 CONVERT(float, CUSTO) AS v\n"
"    FROM Produtos_Precos\n"
"    WHERE Produtos_Precos.COD_PRODUTO = PI.COD_PRODUTO\n"
"      AND Produtos_Precos.CUSTO >= 0\n"
"      AND Produtos_Precos.CUSTO < 1000000\n"
"    ORDER BY CODIGO DESC\n"
") PP\n"
"WHERE PI.Custo = 0 OR PI.Custo IS NULL;\n"
"\n"
"-- Passo 2: subtotal do custo = custo unitario * quantidade\n"
"UPDATE PI\n"
"SET PI.SUBTOTAL_CUSTO = CONVERT(decimal(16,2),\n"
"        CASE WHEN CALC.v >= 0 AND CALC.v < 10000000000000 THEN CALC.v ELSE 0 END)\n"
"FROM pedidos_itens PI\n"
"CROSS APPLY (\n"
"    SELECT CONVERT(float, ISNULL(PI.Custo,0)) * CONVERT(float, ISNULL(PI.QUANTIDADE,0)) AS v\n"
") CALC\n"
"WHERE PI.SUBTOTAL_CUSTO IS NULL OR PI.SUBTOTAL_CUSTO = 0;\n"
"\n"
"-- Verificacao imediata\n"
"SELECT COUNT(*) AS Restantes_Zerados FROM pedidos_itens WHERE Custo = 0;\n"
)

b = novo.encode('cp1252')
b = b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

alvos = [
    r'C:\Projeto\Arquivos\scripts\OS - Colocar o custo nos pedidos.sql',
    r'C:\scripts\OS - Colocar o custo nos pedidos.sql',
]
for path in alvos:
    with open(path, 'wb') as f:
        f.write(b)

# remove o teste (ja cumpriu o papel) das duas pastas
import os
for path in (r'C:\Projeto\Arquivos\scripts\teste_062.sql', r'C:\scripts\teste_062.sql'):
    if os.path.exists(path):
        os.remove(path)
        print("removido:", path)

print("identicos:", open(alvos[0], 'rb').read() == open(alvos[1], 'rb').read())
print("---")
print(b.decode('cp1252'))
