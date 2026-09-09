-- Adiciona cod_servico / cod_mecanico em OS_Servicos_Auto e popula cod_servico via descricao.
-- Tudo guardado contra as tabelas do modulo OS nao existirem ainda.
IF OBJECT_ID('OS_Servicos_Auto', 'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('OS_Servicos_Auto') AND name = 'cod_servico')
    ALTER TABLE OS_Servicos_Auto ADD cod_servico INT NULL;
GO

IF OBJECT_ID('OS_Servicos_Auto', 'U') IS NOT NULL AND OBJECT_ID('OS_Servicos', 'U') IS NOT NULL
    UPDATE a
    SET    a.cod_servico = s.CODIGO
    FROM   OS_Servicos_Auto a
    INNER JOIN OS_Servicos s ON s.SERVICO = a.descricao;
GO

IF OBJECT_ID('OS_Servicos_Auto', 'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('OS_Servicos_Auto') AND name = 'cod_mecanico')
    ALTER TABLE OS_Servicos_Auto ADD cod_mecanico INT NULL;
GO
