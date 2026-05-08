-- Reforma tributária: ajuste dos campos IBS/CBS/IS na tabela produtos
-- Remove: CBSpAliq, IBSUFpAliq, IBSMunpAliq, ISpIS
-- Adiciona: cClassTrib, cClassTrib_IS, tipo_calculo_is
-- Mantém: IBSCBSCST, ISCST (já existem)

ALTER TABLE [dbo].[produtos]
    DROP COLUMN [CBSpAliq], [IBSUFpAliq], [IBSMunpAliq], [ISpIS];

ALTER TABLE [dbo].[produtos]
    ADD [cClassTrib]       VARCHAR(6)  NULL,
        [cClassTrib_IS]    VARCHAR(6)  NULL,
        [tipo_calculo_is]  SMALLINT    NULL;
