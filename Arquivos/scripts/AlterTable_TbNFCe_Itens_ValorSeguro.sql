-- Adiciona o valor de seguro por item na NFCe (TbNFCe_Itens) - campo que nunca existiu
-- nesse modelo (diferente de Valor_Frete/ValorOutras, que ja existiam). O GerarItens
-- do NFCe sempre passou 0 fixo pro parametro valorSeguro por falta desse campo.
-- NULL sem DEFAULT (metadata-only, seguro em SQL Server pre-2012/Express).

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe_Itens') AND name = 'Valor_Seguro')
    ALTER TABLE TbNFCe_Itens ADD Valor_Seguro DECIMAL(15, 2) NULL;
GO
