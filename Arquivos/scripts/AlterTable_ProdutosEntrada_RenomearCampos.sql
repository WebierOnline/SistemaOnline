-- Ajusta campos de produtos_entrada para ter nome/tipo igual aos campos correspondentes em EntradaEstoque
-- Execute cada bloco em ordem

-- 1. Renomear colunas
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'NOTAFISCAL')
   AND NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'NumeroNota')
    EXEC sp_rename 'produtos_entrada.NOTAFISCAL', 'NumeroNota', 'COLUMN';

IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'VALOR')
   AND NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'ValorNota')
    EXEC sp_rename 'produtos_entrada.VALOR', 'ValorNota', 'COLUMN';

IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'COD_FORNECEDOR')
   AND NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'CodigoCorrentista')
    EXEC sp_rename 'produtos_entrada.COD_FORNECEDOR', 'CodigoCorrentista', 'COLUMN';

IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'COD_TRANSPORTADORA')
   AND NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'TranspCodigo')
    EXEC sp_rename 'produtos_entrada.COD_TRANSPORTADORA', 'TranspCodigo', 'COLUMN';

IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'VALOR_FRETE')
   AND NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'ValorFrete')
    EXEC sp_rename 'produtos_entrada.VALOR_FRETE', 'ValorFrete', 'COLUMN';

IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'CHAVE')
   AND NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'ChavedeAcesso')
    EXEC sp_rename 'produtos_entrada.CHAVE', 'ChavedeAcesso', 'COLUMN';

IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'TIPO_FRETE')
   AND NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos_entrada') AND name = 'ModFrete')
    EXEC sp_rename 'produtos_entrada.TIPO_FRETE', 'ModFrete', 'COLUMN';

-- 2. Ajustar tipos (money -> decimal, nvarchar -> varchar, tamanho)
ALTER TABLE produtos_entrada ALTER COLUMN ValorNota      decimal(15, 2) NULL;
ALTER TABLE produtos_entrada ALTER COLUMN ValorFrete     decimal(15, 2) NULL;
ALTER TABLE produtos_entrada ALTER COLUMN ModFrete       varchar(38)    NULL;
