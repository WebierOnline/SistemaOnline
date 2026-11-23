-- Promocoes por WhatsApp: tipo "Texto + link". O link de verdade fica em Campanha.LinkUrl e a mensagem leva o link
-- curto do servidor de licencas (Campanha.LinkCurto = https://onlinecommerce.com.br/l/Ab3x9Q; cada destinatario
-- recebe o seu, com "-<numero do destinatario em base 36>" no fim). CampanhaDestinatario.Cliques = quantas vezes ele
-- clicou (lido do servidor ao abrir Acompanhar campanhas). Idempotente.
IF OBJECT_ID('dbo.Campanha', 'U') IS NOT NULL AND COL_LENGTH('dbo.Campanha', 'LinkUrl') IS NULL
    ALTER TABLE dbo.Campanha ADD LinkUrl nvarchar(2000) NULL
GO
IF OBJECT_ID('dbo.Campanha', 'U') IS NOT NULL AND COL_LENGTH('dbo.Campanha', 'LinkCurto') IS NULL
    ALTER TABLE dbo.Campanha ADD LinkCurto nvarchar(100) NULL
GO
IF OBJECT_ID('dbo.CampanhaDestinatario', 'U') IS NOT NULL AND COL_LENGTH('dbo.CampanhaDestinatario', 'Cliques') IS NULL
    ALTER TABLE dbo.CampanhaDestinatario ADD Cliques int NULL
GO
