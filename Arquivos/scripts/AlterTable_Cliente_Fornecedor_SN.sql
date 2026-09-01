-- Campo SN (endereço "Sem Número") em cliente e fornecedor. Quando SN = 1, o número
-- do endereço não é usado (o campo numero fica "0", travado na tela) e a NFe/NFCe
-- emitida pra esse cliente/fornecedor usa "S/N" no lugar do número do endereço.

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('cliente') AND name = 'SN')
    ALTER TABLE [dbo].[cliente] ADD [SN] BIT NOT NULL DEFAULT 0;
GO

IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('fornecedor') AND name = 'SN')
    ALTER TABLE [dbo].[fornecedor] ADD [SN] BIT NOT NULL DEFAULT 0;
GO
