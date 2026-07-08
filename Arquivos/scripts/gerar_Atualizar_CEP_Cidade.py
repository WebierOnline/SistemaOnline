# -*- coding: utf-8 -*-
# Le CEP.sql e gera Atualizar_CEP_Cidade.sql
# Usa tabela temporaria para nao criar/alterar LOCALIDADES

import re

src  = open('C:/Projeto/CEP.sql', 'r', encoding='latin-1').read()
out  = open('C:/Projeto/scripts/Atualizar_CEP_Cidade.sql', 'w', encoding='utf-8')

out.write("""\
-- ============================================================
-- Preenche Cidade.CEP onde estiver vazio
-- Fonte: CEP.sql (tabela LOCALIDADES)
-- Chave: Cidade.CodigoMunicipio = COD_IBGE
-- ============================================================
SET NOCOUNT ON;
GO

-- Cria tabela temporaria com os CEPs
CREATE TABLE #CepRef (
    COD_IBGE  NVARCHAR(7),
    CEP       NVARCHAR(10)
);
GO

""")

# Extrai todos os VALUES do INSERT INTO LOCALIDADES
pattern = re.compile(
    r"INSERT INTO LOCALIDADES\s*\([^)]+\)\s*VALUES\s*\((\d+),\s*'[^']*',\s*'[^']*',\s*'([^']*)',\s*(\d+)\)",
    re.IGNORECASE
)

rows = []
for m in pattern.finditer(src):
    cep      = m.group(2).strip()
    cod_ibge = m.group(3).strip()
    rows.append(f"  ('{cod_ibge}', '{cep}')")

batch_size = 999
for i in range(0, len(rows), batch_size):
    lote = rows[i:i+batch_size]
    out.write('INSERT INTO #CepRef (COD_IBGE, CEP)\nVALUES\n')
    out.write(',\n'.join(lote))
    out.write(';\nGO\n\n')

out.write("""\
-- Atualiza Cidade.CEP onde estiver NULL ou vazio, e preenche aliquotas IBS
UPDATE C
SET    C.CEP        = RTRIM(R.CEP),
       C.IBSUFpAliq  = 0.10,
       C.IBSMunpAliq = 0.00
FROM   Cidade C
JOIN   #CepRef R ON CAST(C.CodigoMunicipio AS NVARCHAR(7)) = R.COD_IBGE
WHERE  (C.CEP IS NULL OR RTRIM(C.CEP) = '');
GO

PRINT 'Cidades atualizadas: ' + CAST(@@ROWCOUNT AS VARCHAR) + ' registros.';
GO

DROP TABLE #CepRef;
GO
""")

out.close()
print(f'Atualizar_CEP_Cidade.sql gerado ({len(rows)} CEPs).')
