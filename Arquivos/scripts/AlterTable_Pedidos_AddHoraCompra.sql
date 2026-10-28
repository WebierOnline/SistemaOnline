-- pedidos.DATA_COMPRA guarda so a data (hora sempre 00:00:00 - o PDV nunca grava hora
-- nesse campo, e varias telas de consulta ja dependem de comparacao exata com ele, entao
-- nao da pra simplesmente passar a gravar hora ali sem quebrar esses filtros).
-- HORA_COMPRA e um campo novo, independente, pra guardar o horario real do FECHAMENTO da
-- venda (nao da abertura/criacao automatica do pedido) - nada depende da precisao dele hoje.
IF COL_LENGTH('pedidos', 'HORA_COMPRA') IS NULL
    ALTER TABLE pedidos ADD HORA_COMPRA TIME NOT NULL DEFAULT ('00:00:00');
GO
