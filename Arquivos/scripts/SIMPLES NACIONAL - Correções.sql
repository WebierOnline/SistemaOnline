
/****** 
Update empresa set EMAIL = 'seu_email@gmail.com' where EMAIL = '' 
******/
IF EXISTS (SELECT 1 FROM empresa WHERE CRT <> 3)
BEGIN
	UPDATE produtos SET ESTOQUE_FISCAL = ABS(QUANT_ESTOQUE) WHERE QUANT_ESTOQUE <= 0
	UPDATE produtos SET ESTOQUE_FISCAL = ABS(QUANT_ESTOQUE) WHERE QUANT_ESTOQUE > 0
	UPDATE produtos SET ESTOQUE_FISCAL = 0 WHERE ESTOQUE_FISCAL is null
	UPDATE produtos SET QUANT_ESTOQUE = 0 WHERE QUANT_ESTOQUE <= 0
	UPDATE produtos SET QUANT_ESTOQUE = 0 WHERE QUANT_ESTOQUE is null

	UPDATE Produtos SET DESCRICAO = UPPER(DESCRICAO)
	Update Produtos Set DESCRICAO = REPLACE(DESCRICAO, '|',' ')
	Update Produtos Set DESCRICAO = REPLACE(DESCRICAO, '*','')
	Update Produtos Set DESCRICAO = REPLACE(DESCRICAO, '#','')
	UPDATE Produtos SET DESCRICAO = REPLACE(DESCRICAO, 'Ç', 'C')
	UPDATE Produtos SET DESCRICAO = REPLACE(REPLACE(DESCRICAO, '''', ''), '  ', ' ')
	Update Produtos Set DESCRICAO = LTRIM(RTRIM(DESCRICAO))
	Update Produtos Set DESCRICAO = REPLACE(DESCRICAO, '  ',' ')

	Update Produtos Set NCM = LTRIM(RTRIM(NCM))
	Update Produtos Set NCM = '00000000' WHERE (LEN(NCM) > 8)
	Update Produtos Set NCM = '00000000' WHERE (LEN(NCM) < 8)
	Update Produtos Set NCM = '00000000' WHERE (ISNUMERIC(NCM) = 0)
	Update Produtos Set CEST = LTRIM(RTRIM(CEST))
	Update Produtos Set CEST = '0' WHERE (LEN(CEST) > 8)
	Update Produtos Set CEST = '0' WHERE (LEN(CEST) < 8)

	UPDATE produtos SET cClassTrib = '000001', IBSCBSCST = '000', ISCST = '99' WHERE cClassTrib IS NULL OR len(cClassTrib) < 6;

	Update Produtos Set ICMSCST = 102 WHERE CFOP = 5102
	Update Produtos Set ICMSCST = 500 WHERE CFOP = 5405
	Update Produtos Set ICMSCST = 102, CFOP = 5102 WHERE CFOP <> 5102 AND CFOP <> 5405

	Update Produtos Set PISCST = '07', COFINSCST = '07' 

		/* PIS E COFINS */
	-- 1. Produtos Normais (CFOP 5102)
	UPDATE produtos 
	SET PISCST = '49', PISAliq = 0, COFINSCST = '49', COFINSAliq = 0
	WHERE CFOP = '5102';

	-- 2. Produtos Monofásicos/ST (CFOP 5405)
	UPDATE produtos 
	SET PISCST = '04', PISAliq = 0, COFINSCST = '04', COFINSAliq = 0
	WHERE CFOP = '5405';

	-- 3. Outros casos (CFOP 5403 ou Devoluções)
	UPDATE produtos 
	SET PISCST = '49', PISAliq = 0, COFINSCST = '49', COFINSAliq = 0
	WHERE CFOP NOT IN ('5102', '5405');

	Update Produtos Set IPICST = 99
	Update Produtos_Precos Set CUSTO = VALOR_VV WHERE CUSTO = '0.00'

	Update Produtos Set EAN = LTRIM(RTRIM(EAN))
	Update Produtos Set EAN = COD_BARRA WHERE LEN(COD_BARRA) = 13
	UPDATE Produtos SET EAN = RIGHT(EAN, 13) WHERE LEN(EAN) = 14
	Update Produtos Set EAN = CODIGO WHERE LEN(EAN) > 14
	Update Produtos Set EAN = 'SEM GTIN' WHERE LEN(EAN) < 8	

	Update Produtos Set COD_BARRA = LTRIM(RTRIM(COD_BARRA))
	UPDATE Produtos SET COD_BARRA = RIGHT(COD_BARRA, 13) WHERE LEN(COD_BARRA) = 14
	Update Produtos Set COD_BARRA = CODIGO WHERE LEN(COD_BARRA) > 14

	Update Produtos set UNID_MEDIDA = LEFT(UNID_MEDIDA, 2) WHERE (LEN(UNID_MEDIDA) > 2)
	Update Produtos set UNID_MEDIDA = 'UN' WHERE (UNID_MEDIDA NOT IN ('CT', 'UN', 'CX', 'KG', 'PO', 'SC', 'PA', 'EX', 'BJ', 'DZ', 'PC', 'DI', 'FD', 'PT'))
	
	Update Produtos_Precos Set CUSTO = VALOR_VV WHERE CUSTO <= 0
	
	Update NotaFiscalItens Set EAN = 'SEM GTIN' WHERE LEN(EAN) < 8
	Update NotaFiscalItens Set PISCST = '08', COFINSCST = '08' 

	UPDATE TbNFCe SET BaseCalc_ICMS = '0.00'

	Update TbNFCe_Itens Set PISCST = '08', COFINSCST = '08' 
	Update TbNFCe_Itens Set CodBarras = 'SEM GTIN' WHERE LEN(CodBarras) < 8
	Update TbNFCe_Itens Set CodBarras = 'SEM GTIN' WHERE LEN(CodBarras) < 8
	Update TbNFCe_Itens Set ICMSCST = 102 WHERE CFOP = 5102
	Update TbNFCe_Itens Set ICMSCST = 500 WHERE CFOP = 2405
	UPDATE TbNFCe_Itens SET Bc_Icms = '0.00'
	UPDATE TbNFCe_Itens SET Vlr_Icms = '0.00'
    UPDATE TbNFCe_Itens SET Aliq_Icms = '22.50'
END
ELSE
BEGIN
  -- --- LUCRO PRESUMIDO (CRT 3) ---
    
    -- 1. Normal (CFOP 5102)
    UPDATE produtos SET PISCST = '01', PISAliq = 0.65, COFINSCST = '01', COFINSAliq = 3.00 WHERE CFOP = '5102';
    
    -- 2. Monofásico (CFOP 5405)
    UPDATE produtos SET PISCST = '04', PISAliq = 0, COFINSCST = '04', COFINSAliq = 0 WHERE CFOP = '5405';
    
    -- 3. Isentos (CST 040, 041)
    UPDATE produtos SET PISCST = '07', PISAliq = 0, COFINSCST = '07', COFINSAliq = 0 WHERE ICMSCST IN ('040', '041');
END



 