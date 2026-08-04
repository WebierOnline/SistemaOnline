-- Tabela de classificações tributárias do Imposto Seletivo (IS)
-- Permite múltiplos registros por cClassTrib_IS (diferentes períodos de vigência)

IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('[dbo].[tbISClassTrib]') AND type = 'U')
BEGIN
    CREATE TABLE [dbo].[tbISClassTrib] (
        [Id]             INT           IDENTITY(1,1) NOT NULL,
        [cClassTrib_IS]  VARCHAR(6)    NOT NULL,
        [Descricao]      VARCHAR(255)  NOT NULL,
        [tipo_calculo_is] SMALLINT     NOT NULL,   -- 1 = Ad Valorem (%), 2 = Ad Rem (R$/unid), 3 = Misto (% + R$/unid)
        [ISpAliq]        DECIMAL(9,4)  NULL,        -- alíquota % (usado quando tipo=1)
        [ISvUnid]        DECIMAL(15,4) NULL,        -- valor fixo por unidade (usado quando tipo=2)
        [dIniVig]        DATE          NOT NULL,
        [dFimVig]        DATE          NULL,
        [uTrib_IS]       VARCHAR(6)    NULL,        -- unidade tributável do IS (ex: 'UN', 'ML', 'KG')
        CONSTRAINT [PK_tbISClassTrib] PRIMARY KEY CLUSTERED ([Id] ASC)
    );

    -- Índice para buscas por código + vigência
    CREATE INDEX [IX_tbISClassTrib_Cod_Vig]
        ON [dbo].[tbISClassTrib] ([cClassTrib_IS], [dIniVig]);

    -- Dados iniciais (alíquotas de teste 2026 — ajustar conforme publicação oficial)
    INSERT INTO [dbo].[tbISClassTrib]
        ([cClassTrib_IS], [Descricao],                          [tipo_calculo_is], [ISpAliq], [ISvUnid], [dIniVig],    [dFimVig],    [uTrib_IS])
    VALUES
        ('900001', 'Vinhos, Espumantes e Cachaças',              1,  20.00, NULL,   '2026-01-01', '2026-12-31', 'UN'),
        ('900002', 'Cervejas e Chopes',                          1,  20.00, NULL,   '2026-01-01', '2026-12-31', 'UN'),
        ('900003', 'Destilados (Uísque, Gin, Vodka)',            1,  20.00, NULL,   '2026-01-01', '2026-12-31', 'UN'),
        ('900010', 'Refrigerantes e Sucos com Adição de Açúcar', 2,  NULL,  0.10,   '2026-01-01', '2026-12-31', 'UN'),
        ('900011', 'Energéticos e Bebidas Esportivas',           2,  NULL,  0.10,   '2026-01-01', '2026-12-31', 'UN'),
        ('900020', 'Veículos Autopropulsados (poluentes)',        1,   7.50, NULL,   '2026-01-01', '2026-12-31', 'UN'),
        ('900040', 'Produtos do Fumo (Cigarros e derivados)',    3,  100.00, 0.50,   '2026-01-01', '2026-12-31', 'UN');
END
