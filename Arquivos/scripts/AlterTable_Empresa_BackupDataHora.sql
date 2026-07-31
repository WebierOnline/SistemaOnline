-- Adiciona a coluna BackupDataHora na tabela empresa, ja usada pelo codigo
-- (cmdLogon/menulogoff_Click no OnlineCommerce, PDV.frm no PDV, BackupNuvem.vbs)
-- mas nunca criada na base. Nullable, mesmo padrao de DFeUltimaConsultaData/
-- DFeUltimaConsultaHora/VencimentoCert (datetime, NULL).

ALTER TABLE empresa ADD BackupDataHora DATETIME NULL;
