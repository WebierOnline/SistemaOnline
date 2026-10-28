-- Garante que exista o setor padrao 'GERAL' na tabela setor (usado como
-- valor inicial/fallback em varios forms - Caixa_Saida, Contas_Cadastro,
-- Caixa_Suprimento, etc). Idempotente.
SET NOCOUNT ON;
GO

IF NOT EXISTS (SELECT 1 FROM setor WHERE setor = 'GERAL')
BEGIN
    INSERT INTO setor (cod_setor, setor)
    VALUES ((SELECT ISNULL(MAX(cod_setor), 0) + 1 FROM setor), 'GERAL');
END
GO
