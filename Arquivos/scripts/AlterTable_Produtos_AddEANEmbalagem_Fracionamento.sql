-- Adiciona EANCaixa, UnidMedCX e Fracionamento na tabela Produtos
-- EANCaixa: EAN da caixa/fardo fechado (unidade de atacado). O campo se chamava
--   EANEmbalagem antes - renomeado pra ficar mais claro (NAO confundir com
--   VinculoXMLProduto.EANEmbalagem, que e um campo de outra tabela e NAO muda).
-- UnidMedCX: unidade de medida da caixa (CX, FD, DZ, etc - diferente da UNID_MEDIDA
--   normal do produto, que e a unidade de venda/embalagem)
-- Fracionamento: quantas unidades vem dentro dessa caixa

IF COL_LENGTH('Produtos', 'EANCaixa') IS NULL
BEGIN
    IF COL_LENGTH('Produtos', 'EANEmbalagem') IS NOT NULL
        EXEC sp_rename 'Produtos.EANEmbalagem', 'EANCaixa', 'COLUMN';
    ELSE
        ALTER TABLE Produtos ADD EANCaixa varchar(14) NOT NULL DEFAULT ('');
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('Produtos') AND name = 'UnidMedCX')
    ALTER TABLE Produtos ADD UnidMedCX varchar(6) NOT NULL DEFAULT ('CX');
GO

-- Garante que todo produto ja cadastrado (coluna recem-criada ou nao) fique com
-- UnidMedCX preenchido - idempotente, so mexe em quem estiver vazio/nulo.
UPDATE Produtos SET UnidMedCX = 'CX' WHERE UnidMedCX IS NULL OR UnidMedCX = '';
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('Produtos') AND name = 'Fracionamento')
    ALTER TABLE Produtos ADD Fracionamento decimal(9,3) NOT NULL DEFAULT (1);
GO
