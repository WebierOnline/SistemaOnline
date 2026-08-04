IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('OS_Equipamento_Auto') AND name = 'CHASSI')
    ALTER TABLE OS_Equipamento_Auto ADD CHASSI nvarchar(20) NULL
GO