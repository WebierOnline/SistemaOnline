-- TbContabilista: dados do contador/contabilista usados na view Contador (SPED).
-- Idempotente: cria a tabela ja com o schema final (CNPJ/CPF separados) se nao existir;
-- se ja existir no schema antigo (so CPF_CNPJ, sem CPF), renomeia CPF_CNPJ->CNPJ e
-- acrescenta CPF. Baseado em "Contador - Cria Tabela.sql" + "Contador - Alteração Tabela.sql"
-- (pasta Arquivos\Contadores), unificados pra instalar certo em base nova ou existente.

IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('TbContabilista') AND type = 'U')
BEGIN
    CREATE TABLE [dbo].[TbContabilista](
        [NomeContabilista] [varchar](40) NOT NULL,
        [CNPJ] [varchar](18) NOT NULL,
        [CPF] [varchar](15) NOT NULL DEFAULT 0,
        [CRC] [varchar](20) NOT NULL,
        [Endereco] [varchar](150) NOT NULL,
        [Num] [varchar](8) NOT NULL,
        [Bairro] [varchar](50) NOT NULL,
        [Complemento] [varchar](100) NOT NULL,
        [CEP] [varchar](10) NOT NULL,
        [UF] [varchar](2) NOT NULL,
        [Cidade] [varchar](50) NOT NULL,
        [CodigoIBGE] [int] NOT NULL,
        [Fone] [varchar](14) NOT NULL,
        [Fax] [varchar](14) NOT NULL,
        [Email] [varchar](100) NOT NULL
    ) ON [PRIMARY];
END
GO

-- Schema antigo: CPF_CNPJ existe e CNPJ ainda nao -> renomeia a coluna
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbContabilista') AND name = 'CPF_CNPJ')
   AND NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbContabilista') AND name = 'CNPJ')
    EXEC sp_RENAME 'TbContabilista.CPF_CNPJ', 'CNPJ', 'COLUMN';
GO

-- Acrescenta CPF se ainda nao existir (tabela antiga so tinha CPF_CNPJ, sem CPF separado)
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('TbContabilista') AND name = 'CPF')
    ALTER TABLE TbContabilista ADD CPF varchar(15) NOT NULL DEFAULT 0;
GO
