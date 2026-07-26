-- Garantir que o campo não aceita Nulos
ALTER TABLE produtos 
ALTER COLUMN Codigo INT NOT NULL;

-- Criar a Chave Primária
ALTER TABLE produtos
ADD CONSTRAINT PK_produtos_Codigo PRIMARY KEY (Codigo);

-- Adiciona campo TAGS na tabela produtos
ALTER TABLE [dbo].[produtos]
    ADD [TAGS] NVARCHAR(50) NULL;
