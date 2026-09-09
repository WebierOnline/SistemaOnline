SET NOCOUNT ON;
DECLARE @CodigoNota INT = 275;

PRINT '=== Empresa (emitente): regime / CRT ===';
SELECT TOP 1 CRT, RegimeTributario
FROM Empresa;

PRINT '=== NotaFiscal: marcacao de destino/consumidor (nomes conforme modNFe.bas) ===';
SELECT CodigoNota, NumeroNota, TipoCliente, CodigoCorrentista,
       IdentificadorDestino, ConsumidorFinal, IndicadorIEDestinatario
FROM NotaFiscal WHERE CodigoNota = @CodigoNota;

PRINT '=== NotaFiscalItens: campos DIFAL (os que modNFe.bas passa pro GerarItensImpostoUFDest) ===';
SELECT ITEM, CFOP, CST,
       pICMSInter, pICMSInterPart, pICMSUFDest,
       pFCPUFDest, vBCFCPUFDest, vBCUFDest, vFCPUFDest, vICMSUFDest, vICMSUFRemet
FROM NotaFiscalItens WHERE CodigoNota = @CodigoNota ORDER BY ITEM;

PRINT '=== Colunas da tabela cliente (pra achar IE / contribuinte) ===';
SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('cliente') ORDER BY column_id;

PRINT '=== Destinatario (linha inteira) ===';
DECLARE @cc INT = (SELECT CodigoCorrentista FROM NotaFiscal WHERE CodigoNota = @CodigoNota);
DECLARE @tc VARCHAR(20) = (SELECT TipoCliente FROM NotaFiscal WHERE CodigoNota = @CodigoNota);
IF @tc = 'FORNECEDOR'
    SELECT * FROM fornecedor WHERE CODIGO = @cc;
ELSE
    SELECT * FROM cliente WHERE CODIGO = @cc;
