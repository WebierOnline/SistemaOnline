-- NT 2026.005 (Simples Nacional na NF-e/NFC-e): campos de enquadramento e cAtivSN.
-- Idempotente. Nao altera nenhum comportamento sozinho: so cria as colunas.
--
-- Empresa.OptaRegimeRegularIBSCBS : 1 = optou pelo regime regular de IBS/CBS (fora do DAS)
-- Empresa.SimplesOpcaoPendente    : so CRT 3: 0 = nao, 1 = opcao pendente pelo CRT 1, 2 = pendente pelo CRT 2
-- Empresa.cAtivSNPadrao           : padrao do codigo de atividade do Simples (01,02,03,04,90) p/ produto novo
-- Produtos.cAtivSN                : codigo de atividade do item (NULL = usa o padrao da empresa)
-- NotaFiscalItens/TbNFCe_Itens.cAtivSN : valor usado na emissao (reemissao e devolucao fieis)

IF COL_LENGTH('Empresa', 'OptaRegimeRegularIBSCBS') IS NULL
    ALTER TABLE Empresa ADD OptaRegimeRegularIBSCBS BIT NOT NULL CONSTRAINT DF_Empresa_OptaRegimeRegularIBSCBS DEFAULT 0;
GO

IF COL_LENGTH('Empresa', 'SimplesOpcaoPendente') IS NULL
    ALTER TABLE Empresa ADD SimplesOpcaoPendente TINYINT NOT NULL CONSTRAINT DF_Empresa_SimplesOpcaoPendente DEFAULT 0;
GO

IF COL_LENGTH('Empresa', 'cAtivSNPadrao') IS NULL
    ALTER TABLE Empresa ADD cAtivSNPadrao VARCHAR(2) NOT NULL CONSTRAINT DF_Empresa_cAtivSNPadrao DEFAULT '01';
GO

IF COL_LENGTH('Produtos', 'cAtivSN') IS NULL
    ALTER TABLE Produtos ADD cAtivSN VARCHAR(2) NULL;
GO

IF COL_LENGTH('NotaFiscalItens', 'cAtivSN') IS NULL
    ALTER TABLE NotaFiscalItens ADD cAtivSN VARCHAR(2) NULL;
GO

IF OBJECT_ID('TbNFCe_Itens') IS NOT NULL AND COL_LENGTH('TbNFCe_Itens', 'cAtivSN') IS NULL
    ALTER TABLE TbNFCe_Itens ADD cAtivSN VARCHAR(2) NULL;
GO
