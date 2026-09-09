-- ============================================================
-- Script : AlterDatabase_ARITHABORT_ON.sql
-- Banco  : SQL Server 2008/2014
-- Motivo : VinculoXMLProduto tinha um indice FILTRADO
--          (IX_VinculoXMLProduto_EAN ON (EANEmbalagem) WHERE EANEmbalagem <> '').
--          INSERT/UPDATE/DELETE numa tabela com indice filtrado exige ARITHABORT ON.
--          O VB6 conecta via Provider=SQLOLEDB.1, que forca ARITHABORT OFF na conexao
--          e NAO herda o default do banco. Resultado: os botoes Vincular / Desvincular
--          da tela "Vincular Produtos da XML ao Cadastro" quebravam com:
--          "Falha em INSERT/UPDATE porque as opcoes SET tem config. incorretas: 'ARITHABORT'".
-- Fix    : (2) recria o indice SEM o filtro -> indice nao-filtrado nao exige nenhuma opcao SET.
--          (1) tambem liga ARITHABORT no default do banco (ajuda sqlcmd e futuras views indexadas).
-- Idempotente.
-- ============================================================
SET NOCOUNT ON;
GO

-- (1) default do banco
IF EXISTS (SELECT 1 FROM sys.databases WHERE database_id = DB_ID() AND is_arithabort_on = 0)
BEGIN
    DECLARE @sql nvarchar(400) =
        N'ALTER DATABASE ' + QUOTENAME(DB_NAME()) + N' SET ARITHABORT ON;';
    EXEC (@sql);
    PRINT 'ARITHABORT ligado no default do banco ' + DB_NAME() + '.';
END
ELSE
    PRINT 'ARITHABORT ja estava ON no default do banco.';
GO

-- (2) FIX DEFINITIVO: de-filtra o indice de VinculoXMLProduto
IF OBJECT_ID('VinculoXMLProduto', 'U') IS NOT NULL
BEGIN
    IF EXISTS (SELECT 1 FROM sys.indexes
               WHERE object_id = OBJECT_ID('VinculoXMLProduto')
                 AND name = 'IX_VinculoXMLProduto_EAN' AND has_filter = 1)
    BEGIN
        DROP INDEX IX_VinculoXMLProduto_EAN ON VinculoXMLProduto;
        CREATE INDEX IX_VinculoXMLProduto_EAN ON VinculoXMLProduto (EANEmbalagem);
        PRINT 'IX_VinculoXMLProduto_EAN recriado SEM filtro.';
    END
    ELSE IF NOT EXISTS (SELECT 1 FROM sys.indexes
                        WHERE object_id = OBJECT_ID('VinculoXMLProduto')
                          AND name = 'IX_VinculoXMLProduto_EAN')
    BEGIN
        CREATE INDEX IX_VinculoXMLProduto_EAN ON VinculoXMLProduto (EANEmbalagem);
        PRINT 'IX_VinculoXMLProduto_EAN criado (nao existia).';
    END
    ELSE
        PRINT 'IX_VinculoXMLProduto_EAN ja esta sem filtro - nada a fazer.';
END
ELSE
    PRINT 'Tabela VinculoXMLProduto nao existe - nada a fazer.';
GO
