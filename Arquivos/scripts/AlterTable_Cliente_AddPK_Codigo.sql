-- Adiciona Primary Key em cliente.codigo
-- Motivo: Clientes_Cadastro.frm nao tinha protecao nenhuma no banco contra 2
-- terminais cadastrando clientes ao mesmo tempo com o mesmo codigo (o codigo
-- so era recalculado com lock no VB6 - essa PK e' a segunda camada de defesa).
-- Confirmado em 2026-08 que nao havia nenhum codigo duplicado na tabela.

-- 0) So roda se a tabela ainda nao tiver nenhuma Primary Key
IF NOT EXISTS (
    SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID('cliente') AND is_primary_key = 1
)
BEGIN
    -- 1) Checagem de seguranca - so roda o ALTER se realmente nao houver duplicado
    IF EXISTS (
        SELECT codigo FROM cliente GROUP BY codigo HAVING COUNT(*) > 1
    )
    BEGIN
        RAISERROR('Existem codigos duplicados em cliente - resolva antes de criar a PK.', 16, 1)
        RETURN
    END

    -- 2) Cria a Primary Key
    ALTER TABLE cliente
    ADD CONSTRAINT PK_cliente_Codigo PRIMARY KEY (codigo);
END
