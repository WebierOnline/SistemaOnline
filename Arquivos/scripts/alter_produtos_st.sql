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
