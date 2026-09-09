-- Empresa.IBPTUltimaImportacao: guarda o AAAAMM (ex: '202609') da ultima vez que a
-- rotina automatica de atualizacao da tabela IBPT rodou com sucesso naquele cliente.
-- Usado por Tela_Principal.VerificarAtualizacaoIBPT (Etapa 2) pra nao tentar de novo
-- no mesmo mes. VARCHAR(6) NULL - o codigo trata NULL como "nunca rodou".
IF COL_LENGTH('Empresa', 'IBPTUltimaImportacao') IS NULL
    ALTER TABLE Empresa ADD IBPTUltimaImportacao VARCHAR(6) NULL;
GO
