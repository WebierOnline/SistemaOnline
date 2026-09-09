novo = (
"SET NOCOUNT ON;\n"
"\n"
"PRINT '=== PARTE 0: apenas tocar as linhas (Custo = Custo + 0) ===';\n"
"UPDATE pedidos_itens SET Custo = Custo + CONVERT(decimal(12,2),0)\n"
"WHERE Custo = 0 OR Custo IS NULL;\n"
"PRINT '=== PARTE 0 OK ===';\n"
"\n"
"PRINT '=== PARTE 1: apenas PI.Custo a partir de Produtos_Precos ===';\n"
"UPDATE PI\n"
"SET PI.Custo = PP.CUSTO\n"
"FROM pedidos_itens PI\n"
"CROSS APPLY (\n"
"    SELECT TOP 1 CUSTO\n"
"    FROM Produtos_Precos\n"
"    WHERE Produtos_Precos.COD_PRODUTO = PI.COD_PRODUTO\n"
"      AND Produtos_Precos.CUSTO >= 0\n"
"      AND Produtos_Precos.CUSTO < 1000000\n"
"    ORDER BY CODIGO DESC\n"
") PP\n"
"WHERE (PI.Custo = 0 OR PI.Custo IS NULL)\n"
"  AND PI.QUANTIDADE >= 0;\n"
"PRINT '=== PARTE 1 OK ===';\n"
"\n"
"PRINT '=== PARTE 2: apenas SUBTOTAL_CUSTO (float -> corta -> decimal) ===';\n"
"UPDATE PI\n"
"SET PI.SUBTOTAL_CUSTO = CONVERT(decimal(16,2),\n"
"        CASE WHEN CALC.v >= 0 AND CALC.v < 10000000000000 THEN CALC.v ELSE 0 END)\n"
"FROM pedidos_itens PI\n"
"CROSS APPLY (\n"
"    SELECT CONVERT(float, ISNULL(PI.Custo,0)) * CONVERT(float, ISNULL(PI.QUANTIDADE,0)) AS v\n"
") CALC\n"
"WHERE PI.SUBTOTAL_CUSTO IS NULL OR PI.SUBTOTAL_CUSTO = 0;\n"
"PRINT '=== PARTE 2 OK ===';\n"
"\n"
"PRINT '=== FIM - todas as partes passaram ===';\n"
"SELECT COUNT(*) AS Restantes_Zerados FROM pedidos_itens WHERE Custo = 0;\n"
)

b = novo.encode('cp1252')
b = b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

for path in (r'C:\Projeto\Arquivos\scripts\teste_062.sql', r'C:\scripts\teste_062.sql'):
    with open(path, 'wb') as f:
        f.write(b)

print("identicos:", open(r'C:\Projeto\Arquivos\scripts\teste_062.sql','rb').read() == open(r'C:\scripts\teste_062.sql','rb').read())
print(b.decode('cp1252'))
