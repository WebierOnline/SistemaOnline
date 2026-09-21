-- VinculoXMLProduto.CST estava varchar(3), mas o codigo de importacao de XML (Entrada_Estoque.
-- Frm) grava orig+CST (Regime Normal, 3 chars: ex "060") OU orig+CSOSN (Simples Nacional,
-- 4 chars: ex "1102") nesse campo. Pra fornecedores do Simples Nacional isso estourava com
-- "Dados de cadeia ou binarios seriam truncados" ao importar XML (erro real, 2026-09-21).
-- Alarga pra varchar(4), suficiente pros dois casos.
IF EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('VinculoXMLProduto') AND name = 'CST' AND max_length < 4
)
    ALTER TABLE VinculoXMLProduto ALTER COLUMN CST VARCHAR(4);
GO
