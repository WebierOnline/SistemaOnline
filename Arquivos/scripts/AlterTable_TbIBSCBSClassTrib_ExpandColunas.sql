-- ============================================================
-- Expande colunas VARCHAR de TbIBSCBSClassTrib
-- Execute antes de rodar Preencher_TbIBSCBSClassTrib.sql
-- ============================================================

ALTER TABLE TbIBSCBSClassTrib ALTER COLUMN DescricaocClassTrib VARCHAR(1500);
ALTER TABLE TbIBSCBSClassTrib ALTER COLUMN LC_Redacao          VARCHAR(1500);
ALTER TABLE TbIBSCBSClassTrib ALTER COLUMN Credito_para        VARCHAR(500);
ALTER TABLE TbIBSCBSClassTrib ALTER COLUMN NomecClassTrib      VARCHAR(255);
GO

PRINT 'Colunas expandidas com sucesso.'
GO
