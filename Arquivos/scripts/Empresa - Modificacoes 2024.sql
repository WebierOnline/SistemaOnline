IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('empresa') AND name = 'ContigenciaNFe')
    ALTER TABLE empresa ADD ContigenciaNFe BIT DEFAULT 0 NOT NULL
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('empresa') AND name = 'ContigenciaNFCe')
    ALTER TABLE empresa ADD ContigenciaNFCe BIT DEFAULT 0 NOT NULL
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('empresa') AND name = 'Banco')
    ALTER TABLE empresa ADD Banco NVARCHAR(20) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('empresa') AND name = 'Agencia')
    ALTER TABLE empresa ADD Agencia NVARCHAR(10) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('empresa') AND name = 'Conta')
    ALTER TABLE empresa ADD Conta NVARCHAR(12) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('empresa') AND name = 'Tipo')
    ALTER TABLE empresa ADD Tipo NVARCHAR(9) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('empresa') AND name = 'Favorecido')
    ALTER TABLE empresa ADD Favorecido NVARCHAR(50) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('empresa') AND name = 'Pix')
    ALTER TABLE empresa ADD Pix NVARCHAR(35) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('empresa') AND name = 'WhatsAppApiKey')
    ALTER TABLE empresa ADD WhatsAppApiKey VARCHAR(255) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('empresa') AND name = 'NFCeOffline')
    ALTER TABLE empresa ADD NFCeOffline BIT DEFAULT 0 NOT NULL
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('empresa') AND name = 'VencimentoCert')
    ALTER TABLE empresa ADD VencimentoCert datetime NULL
GO

UPDATE empresa SET Perfil = 'B'

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('Empresa') AND name = 'RegimeTributario'
)
    ALTER TABLE Empresa ADD RegimeTributario TINYINT NULL;
GO

-- Preenche RegimeTributario com base no CRT para quem ainda nao tem valor definido.
-- CRT 1 = Simples Nacional                    -> RegimeTributario 1
-- CRT 2 = Simples Nacional (Excesso Sublimite) -> RegimeTributario 2
-- CRT 3 = Regime Normal                        -> assume 3 (Lucro Presumido) como padrao;
--         empresas que sao Lucro Real (4) ou Arbitrado (5) precisam ser corrigidas
--         manualmente depois, pois o CRT sozinho nao distingue Presumido/Real/Arbitrado
UPDATE Empresa
SET RegimeTributario = CASE CRT
        WHEN 1 THEN 1
        WHEN 2 THEN 2
        WHEN 3 THEN 3
        ELSE RegimeTributario
    END
WHERE RegimeTributario IS NULL;
GO


IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('Empresa') AND name = 'IPICompoeDIFAL'
)
    ALTER TABLE Empresa ADD IPICompoeDIFAL TINYINT NOT NULL DEFAULT 0;

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('empresa') AND name = 'IEMunicipal')
    ALTER TABLE empresa ADD IEMunicipal NVARCHAR(20) NULL;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('empresa') AND name = 'BackupDataHora')
    ALTER TABLE empresa ADD BackupDataHora DATETIME NULL;