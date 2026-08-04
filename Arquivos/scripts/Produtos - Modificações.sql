IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos') AND name = 'pRedBc')
    ALTER TABLE produtos ADD pRedBc decimal(8, 4) DEFAULT 0 NOT NULL
GO

IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos') AND name = 'OBSERVACAO')
   AND NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos') AND name = 'INF_ADICIONA')
    EXEC sp_rename 'dbo.produtos.OBSERVACAO', 'INF_ADICIONA', 'COLUMN';

-- Campos de Substituicao Tributaria no cadastro de produtos
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('produtos') AND name = 'pMVAST'
)
    ALTER TABLE produtos ADD pMVAST DECIMAL(7, 4) NULL;

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('produtos') AND name = 'pICMSST'
)
    ALTER TABLE produtos ADD pICMSST DECIMAL(7, 4) NULL;

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('produtos') AND name = 'pRedBCST'
)
    ALTER TABLE produtos ADD pRedBCST DECIMAL(7, 4) NULL;

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('produtos') AND name = 'modBC'
)
BEGIN
    ALTER TABLE produtos ADD modBC TINYINT NOT NULL DEFAULT 3;
END

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('produtos') AND name = 'modBCST'
)
BEGIN
    ALTER TABLE produtos ADD modBCST TINYINT NOT NULL DEFAULT 4;
END

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos') AND name = 'IBSCBSCST')
    ALTER TABLE produtos ADD IBSCBSCST VARCHAR(3);
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos') AND name = 'CBSpAliq')
    ALTER TABLE produtos ADD CBSpAliq DECIMAL(10,2);
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos') AND name = 'IBSUFpAliq')
    ALTER TABLE produtos ADD IBSUFpAliq DECIMAL(10,2);
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos') AND name = 'IBSMunpAliq')
    ALTER TABLE produtos ADD IBSMunpAliq DECIMAL(10,2);
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos') AND name = 'ISCST')
    ALTER TABLE produtos ADD ISCST VARCHAR(3);
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos') AND name = 'ISpIS')
    ALTER TABLE produtos ADD ISpIS DECIMAL(10,2);