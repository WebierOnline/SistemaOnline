-- Reforma tributária: ajuste dos campos IBS/CBS/IS na tabela produtos
-- Remove: CBSpAliq, IBSUFpAliq, IBSMunpAliq, ISpIS
-- Adiciona: cClassTrib, cClassTrib_IS, tipo_calculo_is
-- Mantém: IBSCBSCST, ISCST (já existem)

IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos') AND name = 'CBSpAliq')
    ALTER TABLE [dbo].[produtos] DROP COLUMN [CBSpAliq];
GO

IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos') AND name = 'IBSUFpAliq')
    ALTER TABLE [dbo].[produtos] DROP COLUMN [IBSUFpAliq];
GO

IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos') AND name = 'IBSMunpAliq')
    ALTER TABLE [dbo].[produtos] DROP COLUMN [IBSMunpAliq];
GO

IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos') AND name = 'ISpIS')
    ALTER TABLE [dbo].[produtos] DROP COLUMN [ISpIS];
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos') AND name = 'cClassTrib')
    ALTER TABLE [dbo].[produtos] ADD [cClassTrib] VARCHAR(6) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos') AND name = 'cClassTrib_IS')
    ALTER TABLE [dbo].[produtos] ADD [cClassTrib_IS] VARCHAR(6) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos') AND name = 'fator_conversao_IS')
    ALTER TABLE [dbo].[produtos] ADD [fator_conversao_IS] decimal(9,3) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos') AND name = 'tipo_calculo_is')
    ALTER TABLE [dbo].[produtos] ADD [tipo_calculo_is] SMALLINT NULL;
GO

