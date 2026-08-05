-- Adiciona o campo "anual" (empresas que pagam anualmente, nao aparecem
-- no relatorio mensal do GerenciaNet) na tabela empresas_desbloueio.
-- Rodar no banco cyber_baseFINANCEIRO.
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('empresas_desbloueio') AND name = 'anual'
)
BEGIN
    ALTER TABLE empresas_desbloueio ADD anual BIT NOT NULL DEFAULT 0;
END
