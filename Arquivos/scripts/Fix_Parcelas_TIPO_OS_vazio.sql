-- Corrige parcelas de Ordem de Servico (pedidos.tipo_pedido = 'OFICINA') que ficaram com
-- parcelas.tipo em branco/NULL apos a baixa, por bug em Parcelas.frm (vOrigem nao tratava
-- o caso "OFICINA", so "ALUGUEL"/"VENDA" - corrigido no codigo).
-- Efeito do bug: a parcela nao era contabilizada nos totais de "O.S." no fechamento de caixa.

-- 1) Conferir quantos registros serao afetados antes de rodar o UPDATE:
SELECT p.CODIGO, p.COD_PEDIDO, p.COD_OS, p.NUMERO, p.VALOR_FINAL, p.FORMA_PGTO, p.PAGAMENTO, p.STATUS, p.TIPO
FROM parcelas p
INNER JOIN pedidos ped ON ped.cod_pedido = p.cod_pedido
WHERE ped.tipo_pedido = 'OFICINA' AND (p.tipo IS NULL OR p.tipo = '');
GO

-- 2) Corrigir (so parcelas ja baixadas/status=1 - as em aberto (status=0) nao entram em
--    nenhum total do fechamento mesmo, tipo so importa no momento da baixa):
UPDATE p
SET p.TIPO = 'OS'
FROM parcelas p
INNER JOIN pedidos ped ON ped.cod_pedido = p.cod_pedido
WHERE ped.tipo_pedido = 'OFICINA' AND (p.tipo IS NULL OR p.tipo = '') AND p.status = 1;
GO
