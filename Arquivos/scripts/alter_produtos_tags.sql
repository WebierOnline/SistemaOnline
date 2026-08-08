-- Garantir que o campo não aceita Nulos
ALTER TABLE produtos 
ALTER COLUMN Codigo INT NOT NULL;

-- Criar a Chave Primária
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID('produtos') AND name = 'PK_produtos_Codigo')
ALTER TABLE produtos
ADD CONSTRAINT PK_produtos_Codigo PRIMARY KEY (Codigo);

-- Adiciona campo TAGS na tabela produtos
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('produtos') AND name = 'TAGS')
ALTER TABLE [dbo].[produtos]
    ADD [TAGS] NVARCHAR(50) NULL;
