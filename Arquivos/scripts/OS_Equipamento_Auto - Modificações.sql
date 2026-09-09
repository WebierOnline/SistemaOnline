-- Adiciona CHASSI em OS_Equipamento_Auto. Guarda tambem contra a TABELA nao existir
-- (cliente que usa OS mas cujo banco-base ainda nao tem as tabelas do modulo OS) -
-- nesse caso o script apenas nao faz nada, em vez de estourar Msg 4902.
IF OBJECT_ID('OS_Equipamento_Auto', 'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('OS_Equipamento_Auto') AND name = 'CHASSI')
    ALTER TABLE OS_Equipamento_Auto ADD CHASSI nvarchar(20) NULL;
GO
