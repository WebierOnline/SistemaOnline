-- Adiciona a coluna BackupDataHora na tabela empresa, ja usada pelo codigo
-- (cmdLogon/menulogoff_Click no OnlineCommerce, PDV.frm no PDV, BackupNuvem.vbs)
-- mas nunca criada na base. Nullable, mesmo padrao de DFeUltimaConsultaData/
-- DFeUltimaConsultaHora/VencimentoCert (datetime, NULL).

IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('empresa') AND name = 'BackupDataHora'
)
    ALTER TABLE empresa ADD BackupDataHora DATETIME NULL;
