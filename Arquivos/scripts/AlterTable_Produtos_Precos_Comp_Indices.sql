-- Adiciona indices que faltavam em Produtos_Precos e Produtos_Comp.
-- Produtos_Precos (3058 linhas) nao tinha NENHUM indice - toda consulta de preco do sistema
-- (SELECT TOP 1 ... WHERE COD_PRODUTO = produtos.codigo ORDER BY CODIGO DESC, usada em dezenas
-- de telas) fazia uma varredura completa da tabela por produto.
-- Produtos_Comp (compatibilidade veicular, usada no PDV para lojas de autopecas) esta vazia hoje
-- mas tem o mesmo problema - indexando preventivamente antes de comecar a ser usada.
IF NOT EXISTS (
    SELECT 1 FROM sys.indexes
    WHERE object_id = OBJECT_ID('Produtos_Precos') AND name = 'IX_Produtos_Precos_CodProduto'
)
    CREATE INDEX IX_Produtos_Precos_CodProduto ON Produtos_Precos (COD_PRODUTO, CODIGO DESC);
GO

IF NOT EXISTS (
    SELECT 1 FROM sys.indexes
    WHERE object_id = OBJECT_ID('Produtos_Comp') AND name = 'IX_Produtos_Comp_CodProduto'
)
    CREATE INDEX IX_Produtos_Comp_CodProduto ON Produtos_Comp (COD_PRODUTO);
GO
