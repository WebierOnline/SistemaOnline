-- Adiciona coluna cod_servico em OS_Servicos_Auto
-- e preenche com base em OS_Servicos_Auto.descricao = OS_Servicos.SERVICO

-- Passo 1: adicionar a coluna (nullable para permitir o UPDATE a seguir)
ALTER TABLE OS_Servicos_Auto
    ADD cod_servico INT NULL;
GO

-- Passo 2: popular cod_servico onde a descricao bate com o servico cadastrado
UPDATE a
SET    a.cod_servico = s.CODIGO
FROM   OS_Servicos_Auto a
INNER JOIN OS_Servicos s ON s.SERVICO = a.descricao;
GO

-- Resultado esperado: linhas sem match ficam com cod_servico = NULL
-- Para conferir registros sem match:
-- SELECT * FROM OS_Servicos_Auto WHERE cod_servico IS NULL;

-- Passo 3: adicionar coluna cod_mecanico (funcionario que executou o servico)
ALTER TABLE OS_Servicos_Auto
    ADD cod_mecanico INT NULL;
GO
