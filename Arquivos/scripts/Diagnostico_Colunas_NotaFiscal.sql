-- Diagnostico: mostra todas as colunas atuais da tabela NotaFiscal.
-- Objetivo: achar qual campo a Load_Data (NFe_Completa.frm) esta tentando gravar
-- que nao existe mais / nunca existiu na tabela (erro 3265 "item nao encontrado
-- na colecao correspondente ao nome ou ao ordinal solicitado").
-- So faz SELECT, nao altera nada.

SELECT
    c.COLUMN_NAME,
    c.DATA_TYPE,
    c.CHARACTER_MAXIMUM_LENGTH
FROM INFORMATION_SCHEMA.COLUMNS c
WHERE c.TABLE_NAME = 'NotaFiscal'
ORDER BY c.ORDINAL_POSITION;
