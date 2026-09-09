novo = (
"-- 1. LIMPEZA E REGRA GERAL (Aliquota Padrao)\n"
"-- Define todos como IBS Padrao (000001) e sem Imposto Seletivo\n"
"UPDATE tbNCM SET \n"
"    cClassTrib_IBS = '000001', \n"
"    cClassTrib_IS = '', \n"
"    tipo_calculo_is = 0;\n"
"\n"
"-- 2. CESTA BASICA NACIONAL (Art. 125, Anexo I da LC 214/2025 - reducao de 100% = aliquota zero)\n"
"-- cClassTrib 200003 = \"Vendas de produtos destinados a alimentacao humana (Anexo I)\", CST 200.\n"
"-- CORRIGIDO EM 2026-09-01: estava usando 400001 (CST 400 = isencao de transporte publico de\n"
"-- passageiros - nada a ver com alimento), o que a SEFAZ rejeitou em teste real (NCM 10063021 - arroz).\n"
"-- NCMs de Carnes (0201-0204), Peixes (0302-0305), Leite (0401), Feijao (0713), Arroz (1006)\n"
"UPDATE tbNCM SET \n"
"    cClassTrib_IBS = '200003' \n"
"WHERE NCM LIKE '0201%' OR NCM LIKE '0202%' OR NCM LIKE '0203%' OR NCM LIKE '0204%'\n"
"   OR NCM LIKE '0302%' OR NCM LIKE '0303%' OR NCM LIKE '0304%' OR NCM LIKE '0305%'\n"
"   OR NCM LIKE '0401%' OR NCM LIKE '0713%' OR NCM LIKE '1006%';\n"
"\n"
"-- 3. HIGIENE PESSOAL E LIMPEZA (Art. 136, Anexo VIII da LC 214/2025)\n"
"-- cClassTrib 200035 = \"Fornecimento dos produtos de higiene pessoal e limpeza relacionados no Anexo VIII\".\n"
"-- CORRIGIDO EM 2026-09-01: estava usando 200001 (que na verdade e transporte de bens ate zona de\n"
"-- processamento de exportacao - nao tem nada a ver com higiene/limpeza).\n"
"-- Higiene (3401), Limpeza (3402)\n"
"UPDATE tbNCM SET \n"
"    cClassTrib_IBS = '200035'\n"
"WHERE NCM LIKE '3401%' OR NCM LIKE '3402%';\n"
"\n"
"-- 3b. HORTIFRUTI FRESCO/CONGELADO NAO COZIDO (Art. 148, Anexo XV da LC 214/2025)\n"
"-- cClassTrib 200014 = horticolas, frutas e ovos, desde que nao cozidos. NCM 0710 e congelado mas\n"
"-- nao cozido na maioria dos casos (excecao: trufas, NCM 0710.80.00, fora do Anexo XV - raro nesse\n"
"-- ramo de negocio). CONFIANCA MENOR que os itens 2 e 3 acima - revisar se aparecer produto de NCM\n"
"-- 0710 dando problema na SEFAZ.\n"
"UPDATE tbNCM SET \n"
"    cClassTrib_IBS = '200014'\n"
"WHERE NCM LIKE '0710%';\n"
"\n"
"-- 4. IMPOSTO SELETIVO - BEBIDAS ALCOOLICAS (Ad Valorem %)\n"
"-- Cervejas (2203), Vinhos (2204), Aguardentes (2208)\n"
"UPDATE tbNCM SET \n"
"    cClassTrib_IS = '900001', \n"
"    tipo_calculo_is = 1 \n"
"WHERE NCM LIKE '2203%' OR NCM LIKE '2204%' OR NCM LIKE '2208%';\n"
"\n"
"-- 5. IMPOSTO SELETIVO - BEBIDAS ACUCARADAS (Ad Rem R$/Litro)\n"
"-- Refrigerantes, Refrescos, Isotonicos (2202)\n"
"UPDATE tbNCM SET \n"
"    cClassTrib_IS = '900010', \n"
"    tipo_calculo_is = 2 \n"
"WHERE NCM LIKE '2202%';\n"
"\n"
"-- 6. IMPOSTO SELETIVO - PRODUTOS DO FUMO (Misto: % + R$/unidade)\n"
"-- Cigarros e derivados do tabaco (capitulo 24 da NCM)\n"
"UPDATE tbNCM SET\n"
"    cClassTrib_IS = '900040',\n"
"    tipo_calculo_is = 3\n"
"WHERE NCM LIKE '24%';\n"
"\n"
"-- 7. COMBUSTIVEIS - REGIME MONOFASICO (Gas GLP)\n"
"-- NCM do GLP (27111910) - cClassTrib 620006 NAO CONFIRMADO em fonte oficial (busca nao achou\n"
"-- documentacao especifica). Mantido como estava; validar antes de confiar em cliente que vende GLP.\n"
"UPDATE tbNCM SET\n"
"    cClassTrib_IBS = '620006',\n"
"    cClassTrib_IS = '',\n"
"    tipo_calculo_is = 2\n"
"WHERE NCM = '27111910';\n"
)

b = novo.encode('cp1252')
b = b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

alvos = [
    r'C:\Projeto\Arquivos\scripts\tbNCM_preencher_IBSCBS.sql',
    r'C:\scripts\tbNCM_preencher_IBSCBS.sql',
]
for path in alvos:
    with open(path, 'wb') as f:
        f.write(b)

print("identicos:", open(alvos[0], 'rb').read() == open(alvos[1], 'rb').read())
print(b.decode('cp1252'))
