-- Corrige a contaminacao causada pelo bug do vinculo automatico (Entrada_Estoque.frm):
-- itens de XML sem codigo de barras (cEAN = 'SEM GTIN') eram gravados com esse texto
-- literal na coluna EANEmbalagem. Como a busca de vinculo (Busca 2) nao excluia esse
-- valor, TODO produto sem GTIN em qualquer importacao futura passava a "casar" com o
-- primeiro produto que tivesse sido vinculado dessa forma (ver conversa: varios
-- parafusos diferentes vinculados automaticamente ao mesmo "FILTRO DE OLEO...").
-- O codigo ja foi corrigido para nao gravar mais 'SEM GTIN' como EAN; este script so
-- limpa o que ja estava gravado.

-- 1) Revisar antes de aplicar: mostra os vinculos afetados (produto que a XML descreve
--    x produto que foi vinculado por engano)
SELECT
    v.ID,
    v.IDFornecedor,
    v.cProd,
    v.xProd AS DescricaoNaXML,
    v.IDProduto,
    p.Descricao AS ProdutoVinculado,
    v.DataAtualizacao
FROM VinculoXMLProduto v
LEFT JOIN Produtos p ON p.Codigo = v.IDProduto
WHERE v.EANEmbalagem = 'SEM GTIN'
ORDER BY v.DataAtualizacao DESC;

-- 2) Limpa a coluna EANEmbalagem nos vinculos contaminados (para de "casar" produtos
--    diferentes entre si; o vinculo IDProduto em si NAO e mexido aqui - ver passo 3)
UPDATE VinculoXMLProduto
   SET EANEmbalagem = ''
 WHERE EANEmbalagem = 'SEM GTIN';

-- 3) Mesma contaminacao pode ter vazado para Produtos.EANEmbalagem (linha que roda
--    APOS encontrar um IDProduto por engano). Limpa so quando o valor gravado e o
--    texto literal 'SEM GTIN' (nunca seria um EAN real).
UPDATE Produtos
   SET EANEmbalagem = ''
 WHERE EANEmbalagem = 'SEM GTIN';

-- 4) Confirmado pelo resultado do passo 1: os 71 vinculos contaminados sao todos do
--    fornecedor 20 (ASA MULTIPECAS PIAUI), apontando por engano para o produto 1248
--    (FILTRO DE OLEO WOE912 FIAT GRAND SIENA, que pertence de verdade a um vinculo
--    correto do fornecedor 4). Desfaz so esses - devolve para "sem vinculo" (0),
--    para serem revinculados manualmente aos produtos certos (parafusos, porcas,
--    arruelas, pesos) pela tela de importacao/vinculos.
UPDATE VinculoXMLProduto
   SET IDProduto = 0,
       EANProduto = ''
 WHERE IDFornecedor = 20
   AND IDProduto = 1248;

-- Observacao: revise o resultado do passo 1 antes de rodar o passo 4 caso este script
-- seja reaproveitado depois com outro fornecedor/produto - o filtro acima (20 / 1248)
-- e especifico para o caso identificado nesta conversa.
