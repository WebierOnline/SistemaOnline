IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Comissao_Avista')
    EXEC sp_RENAME 'funcionario.Comissao_Avista' , 'Comissao_Avista1', 'COLUMN'
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Comissao_Prazo')
    EXEC sp_RENAME 'funcionario.Comissao_Prazo' , 'Comissao_Prazo1', 'COLUMN'
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Comissao_Recebido')
    EXEC sp_RENAME 'funcionario.Comissao_Recebido' , 'Comissao_Recebido1', 'COLUMN'
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Comissao_Servico')
    EXEC sp_RENAME 'funcionario.Comissao_Servico' , 'Comissao_Servico1', 'COLUMN'

/*** À VISTA ***/

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name IN ('Valor_Comissao1', 'Valor_ComissaoAV1'))
    ALTER TABLE funcionario ADD Valor_Comissao1 decimal(16,2) DEFAULT 0 NOT NULL
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name IN ('Valor_Comissao2', 'Valor_ComissaoAV2'))
    ALTER TABLE funcionario ADD Valor_Comissao2 decimal(16,2) DEFAULT 0 NOT NULL
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name IN ('Valor_Comissao3', 'Valor_ComissaoAV3'))
    ALTER TABLE funcionario ADD Valor_Comissao3 decimal(16,2) DEFAULT 0 NOT NULL
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Comissao_Avista2')
    ALTER TABLE funcionario ADD Comissao_Avista2 decimal(16,2) DEFAULT 0 NOT NULL
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Comissao_Avista3')
    ALTER TABLE funcionario ADD Comissao_Avista3 decimal(16,2) DEFAULT 0 NOT NULL
GO

/*** RECEBIDO ***/
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Valor_ComissaoRec1')
    ALTER TABLE funcionario ADD Valor_ComissaoRec1 decimal(16,2) DEFAULT 0 NOT NULL
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Valor_ComissaoRec2')
    ALTER TABLE funcionario ADD Valor_ComissaoRec2 decimal(16,2) DEFAULT 0 NOT NULL
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Valor_ComissaoRec3')
    ALTER TABLE funcionario ADD Valor_ComissaoRec3 decimal(16,2) DEFAULT 0 NOT NULL
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Comissao_Recebido2')
    ALTER TABLE funcionario ADD Comissao_Recebido2 decimal(16,2) DEFAULT 0 NOT NULL
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Comissao_Recebido3')
    ALTER TABLE funcionario ADD Comissao_Recebido3 decimal(16,2) DEFAULT 0 NOT NULL
GO

/*** RENOMEAR Valor_Comissao -> Valor_ComissaoAV ***/
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Valor_Comissao1')
   AND NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Valor_ComissaoAV1')
    EXEC sp_RENAME 'funcionario.Valor_Comissao1', 'Valor_ComissaoAV1', 'COLUMN'
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Valor_Comissao2')
   AND NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Valor_ComissaoAV2')
    EXEC sp_RENAME 'funcionario.Valor_Comissao2', 'Valor_ComissaoAV2', 'COLUMN'
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Valor_Comissao3')
   AND NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Valor_ComissaoAV3')
    EXEC sp_RENAME 'funcionario.Valor_Comissao3', 'Valor_ComissaoAV3', 'COLUMN'

/*** REMOVER campos legados ***/
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'comissaoservicos')
BEGIN
    DECLARE @dfName1 NVARCHAR(200)
    SELECT @dfName1 = dc.name
    FROM sys.default_constraints dc
    JOIN sys.columns c ON c.object_id = dc.parent_object_id AND c.column_id = dc.parent_column_id
    WHERE dc.parent_object_id = OBJECT_ID('funcionario') AND c.name = 'comissaoservicos'
    IF @dfName1 IS NOT NULL
        EXEC('ALTER TABLE funcionario DROP CONSTRAINT ' + @dfName1)
    ALTER TABLE funcionario DROP COLUMN comissaoservicos
END
GO

IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'comissaovendas')
BEGIN
    DECLARE @dfName2 NVARCHAR(200)
    SELECT @dfName2 = dc.name
    FROM sys.default_constraints dc
    JOIN sys.columns c ON c.object_id = dc.parent_object_id AND c.column_id = dc.parent_column_id
    WHERE dc.parent_object_id = OBJECT_ID('funcionario') AND c.name = 'comissaovendas'
    IF @dfName2 IS NOT NULL
        EXEC('ALTER TABLE funcionario DROP CONSTRAINT ' + @dfName2)
    ALTER TABLE funcionario DROP COLUMN comissaovendas
END
GO

/*** À PRAZO ***/
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Comissao_Prazo2')
    ALTER TABLE funcionario ADD Comissao_Prazo2 decimal(16,2) DEFAULT 0 NOT NULL
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Comissao_Prazo3')
    ALTER TABLE funcionario ADD Comissao_Prazo3 decimal(16,2) DEFAULT 0 NOT NULL
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Valor_ComissaoAP1')
    ALTER TABLE funcionario ADD Valor_ComissaoAP1 decimal(16,2) DEFAULT 0 NOT NULL
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Valor_ComissaoAP2')
    ALTER TABLE funcionario ADD Valor_ComissaoAP2 decimal(16,2) DEFAULT 0 NOT NULL
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Valor_ComissaoAP3')
    ALTER TABLE funcionario ADD Valor_ComissaoAP3 decimal(16,2) DEFAULT 0 NOT NULL
GO

/*** SERVIÇOS ***/
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Comissao_Servico2')
    ALTER TABLE funcionario ADD Comissao_Servico2 decimal(16,2) DEFAULT 0 NOT NULL
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Comissao_Servico3')
    ALTER TABLE funcionario ADD Comissao_Servico3 decimal(16,2) DEFAULT 0 NOT NULL
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Valor_ComissaoServ1')
    ALTER TABLE funcionario ADD Valor_ComissaoServ1 decimal(16,2) DEFAULT 0 NOT NULL
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Valor_ComissaoServ2')
    ALTER TABLE funcionario ADD Valor_ComissaoServ2 decimal(16,2) DEFAULT 0 NOT NULL
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('funcionario') AND name = 'Valor_ComissaoServ3')
    ALTER TABLE funcionario ADD Valor_ComissaoServ3 decimal(16,2) DEFAULT 0 NOT NULL
GO
