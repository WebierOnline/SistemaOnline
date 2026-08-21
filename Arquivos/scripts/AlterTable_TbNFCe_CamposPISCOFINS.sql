-- Totais de PIS/COFINS no cabecalho da NFCe (TbNFCe), pra relatorio/apuracao sem
-- precisar somar TbNFCe_Itens toda vez. Nao exigido pela SEFAZ (a XML ja calcula
-- tudo direto dos itens na hora da transmissao, nunca depende de total salvo) -
-- conveniencia de relatorio, mesmo motivo dos totais de ICMS e da reforma.
-- NULL sem DEFAULT (metadata-only, seguro em SQL Server pre-2012/Express).

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe') AND name = 'vPIS')
    ALTER TABLE TbNFCe ADD vPIS DECIMAL(15, 2) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe') AND name = 'vCOFINS')
    ALTER TABLE TbNFCe ADD vCOFINS DECIMAL(15, 2) NULL;
GO
