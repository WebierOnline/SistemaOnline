-- Adiciona TbNFCe_Itens.TipoProduto - coluna usada pelas procedures NFCeDuplicar e NFCeIncluir
-- ('Produto' / 'Combustivel', vindo de CASE WHEN produtos.COMBUSTIVEL = 1). Existe nos bancos
-- antigos por schema acumulado, mas nao era criada por nenhum script do manifesto - cliente com
-- banco-base mais novo dava 'Nome de coluna TipoProduto invalido' (Msg 207) no script NFCeDuplicar.
-- NULL sem DEFAULT = metadata-only em qualquer versao do SQL Server (inclusive 2008 Express).
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbNFCe_Itens') AND name = 'TipoProduto')
    ALTER TABLE TbNFCe_Itens ADD TipoProduto varchar(20) NULL;
