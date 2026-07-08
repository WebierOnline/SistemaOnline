-- Ajusta campos de produtos_entrada para ter nome/tipo igual aos campos correspondentes em EntradaEstoque
-- Execute cada bloco em ordem

-- 1. Renomear colunas
EXEC sp_rename 'produtos_entrada.NOTAFISCAL',        'NumeroNota',       'COLUMN';
EXEC sp_rename 'produtos_entrada.VALOR',             'ValorNota',        'COLUMN';
EXEC sp_rename 'produtos_entrada.COD_FORNECEDOR',    'CodigoCorrentista','COLUMN';
EXEC sp_rename 'produtos_entrada.COD_TRANSPORTADORA','TranspCodigo',     'COLUMN';
EXEC sp_rename 'produtos_entrada.VALOR_FRETE',       'ValorFrete',       'COLUMN';
EXEC sp_rename 'produtos_entrada.CHAVE',             'ChavedeAcesso',    'COLUMN';
EXEC sp_rename 'produtos_entrada.TIPO_FRETE',        'ModFrete',         'COLUMN';

-- 2. Ajustar tipos (money -> decimal, nvarchar -> varchar, tamanho)
ALTER TABLE produtos_entrada ALTER COLUMN ValorNota      decimal(15, 2) NULL;
ALTER TABLE produtos_entrada ALTER COLUMN ValorFrete     decimal(15, 2) NULL;
ALTER TABLE produtos_entrada ALTER COLUMN ModFrete       varchar(38)    NULL;
