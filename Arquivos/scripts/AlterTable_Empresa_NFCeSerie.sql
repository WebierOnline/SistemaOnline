-- Empresa.NFCeSerie: serie da NFCe (1 a 9), escolhida em Empresa_Cadastro.cboNFCeSerie.
-- Antes a serie era fixa em 1 (constante gravada dentro da stored procedure NFCeIncluir).
-- Ver AlterProcedure_NFCeIncluir_SerieDinamica.sql (roda depois, le esse campo).
IF COL_LENGTH('Empresa', 'NFCeSerie') IS NULL
    ALTER TABLE Empresa ADD NFCeSerie SMALLINT NOT NULL DEFAULT 1;
GO

UPDATE Empresa SET NFCeSerie = 1 WHERE NFCeSerie IS NULL OR NFCeSerie < 1 OR NFCeSerie > 9;
GO
