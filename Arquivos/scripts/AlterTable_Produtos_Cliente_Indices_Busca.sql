-- Indices para as buscas "comeca com" (LIKE 'texto%') em Produtos_Cadastro (Iniciais) e Clientes_Cadastro (Nome)
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID('produtos') AND name = 'IX_Produtos_Descricao')
    CREATE INDEX IX_Produtos_Descricao ON produtos (descricao);
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID('cliente') AND name = 'IX_Cliente_Nome')
    CREATE INDEX IX_Cliente_Nome ON cliente (nome);
GO
