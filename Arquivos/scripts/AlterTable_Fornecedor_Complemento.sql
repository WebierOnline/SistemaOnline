-- Aumenta fornecedor.Complemento de nvarchar(4) para nvarchar(50).
-- Causa do erro "Dados de cadeia ou binarios seriam truncados" ao importar XML
-- (Entrada_Estoque.frm): coluna estava pequena demais para o campo de complemento
-- de endereco (xCpl) vindo da NF-e, ex: 'QD 20' (5 caracteres) ja estourava o limite de 4.
IF EXISTS (
    SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
    WHERE TABLE_NAME = 'fornecedor' AND COLUMN_NAME = 'Complemento'
      AND CHARACTER_MAXIMUM_LENGTH < 50
)
    ALTER TABLE fornecedor ALTER COLUMN Complemento NVARCHAR(50) NULL;
GO
