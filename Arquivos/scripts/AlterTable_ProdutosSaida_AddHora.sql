-- produtos_saida.data nao guarda hora nenhuma (so grava a data, sempre meia-noite,
-- veja Inserir_Dados/Produtos_Saida_Estoque.frm).
-- HORA e um campo novo, independente, pra guardar o horario real do registro da
-- saida (lancado no Salvar) - registros antigos ficam com 00:00:00 (tratado como
-- "sem hora" na tela de consulta).
IF COL_LENGTH('produtos_saida', 'HORA') IS NULL
    ALTER TABLE produtos_saida ADD HORA TIME NOT NULL DEFAULT ('00:00:00');
GO
