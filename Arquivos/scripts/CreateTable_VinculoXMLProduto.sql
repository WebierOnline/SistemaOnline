-- ============================================================
-- Script: CreateTable_VinculoXMLProduto.sql
-- Banco  : SQL Server 2014
-- Tabela : VinculoXMLProduto
-- Objetivo: Mapear o produto da NF-e (embalagem do fornecedor)
--           ao produto interno do sistema (unidade de venda),
--           armazenando fracionamento e custo unitario.
-- ============================================================

IF NOT EXISTS (
    SELECT 1 FROM sys.objects
    WHERE object_id = OBJECT_ID('VinculoXMLProduto') AND type = 'U'
)
BEGIN
    CREATE TABLE VinculoXMLProduto (
        -- Chave primaria
        ID              INT          IDENTITY(1,1) NOT NULL,

        -- Identificacao do fornecedor e produto na XML
        IDFornecedor    INT          NOT NULL,               -- Codigo do fornecedor (tabela Fornecedor)
        cProd           VARCHAR(60)  NOT NULL DEFAULT '',    -- <cProd>: codigo do produto no fornecedor
        EANEmbalagem    VARCHAR(14)  NOT NULL DEFAULT '',    -- <cEAN>:  EAN da embalagem comercial (caixa)
        xProd           VARCHAR(120) NOT NULL DEFAULT '',    -- <xProd>: descricao do produto conforme fornecedor
        uCom            VARCHAR(6)   NOT NULL DEFAULT '',    -- <uCom>:  unidade comercial do fornecedor (CX006, CX012, etc.)

        -- Vinculo com o produto interno
        IDProduto       INT          NOT NULL,               -- Codigo do produto interno (unidade de venda)
        EANProduto      VARCHAR(14)  NOT NULL DEFAULT '',    -- EAN da unidade interna (referencia, nao obrigatorio)

        -- Conversao de embalagem para unidade
        Fracionamento   DECIMAL(10,4) NOT NULL DEFAULT 1,   -- Qtd de unidades internas por embalagem do fornecedor
                                                            -- Ex: CX006 = 6, CX012 = 12, CX024 = 24

        -- Custo unitario (calculado na ultima entrada: vUnCom / Fracionamento)
        CustoUnitario   DECIMAL(15,4) NOT NULL DEFAULT 0,   -- Ultimo custo unitario registrado na entrada XML

        -- Auditoria
        DataAtualizacao DATETIME     NOT NULL DEFAULT GETDATE(), -- Data da ultima atualizacao do vinculo

        CONSTRAINT PK_VinculoXMLProduto PRIMARY KEY (ID)
    );

    -- Indice unico: um fornecedor nao pode ter o mesmo cProd mapeado duas vezes
    CREATE UNIQUE INDEX UQ_VinculoXMLProduto_Forn_cProd
        ON VinculoXMLProduto (IDFornecedor, cProd);

    -- Indice para busca por EAN da embalagem (pode nao ser unico entre fornecedores diferentes)
    CREATE INDEX IX_VinculoXMLProduto_EAN
        ON VinculoXMLProduto (EANEmbalagem)
        WHERE EANEmbalagem <> '';

    PRINT 'Tabela VinculoXMLProduto criada com sucesso.';
END
ELSE
BEGIN
    PRINT 'Tabela VinculoXMLProduto ja existe.';
END
GO

-- ============================================================
-- Verificacao
-- ============================================================
SELECT
    name        AS Campo,
    TYPE_NAME(user_type_id) AS Tipo,
    max_length  AS Tamanho,
    is_nullable AS Nulo,
    object_definition(default_object_id) AS ValorDefault
FROM sys.columns
WHERE object_id = OBJECT_ID('VinculoXMLProduto')
ORDER BY column_id;
GO
