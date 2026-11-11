/***************************
DESCRICAO.: Conversao do CFOP/CST da nota de entrada (XML do fornecedor) para o CFOP e o CST/CSOSN de VENDA do produto,
            conforme o regime da empresa:
              Regime 'SN' = Simples Nacional e MEI (empresa.CRT 1 e 4): CST de venda = CSOSN (3 digitos);
              Regime 'RN' = regime normal - Lucro Presumido/Real e Simples com excesso de sublimite (CRT 2 e 3):
                            CST de venda = 2 digitos; a origem (1o digito) vem do XML.
            TipoCodigo/CodigoEntrada = o que veio no XML: CST (2 digitos, fornecedor do regime normal) ou CSOSN
            (3 digitos, fornecedor do Simples); '*' = qualquer.
            GrupoCFOP = 'ST' quando o CFOP do fornecedor e x401, x403 ou x405 (mercadoria com ICMS-ST), 'NORMAL' nos
            demais; '*' = qualquer.
            A regra mais especifica vence: codigo + grupo, depois so o grupo, depois so o codigo, depois '*'/'*'.
            Usada pela entrada de estoque (importacao do XML), pelo cadastro de produto e pela conferencia do SPED
            (funcao dbo.ConversaoFiscalVenda). As regras sao ajustadas na tela Fiscal > Conversao CFOP/CST (Revisar = 1 =
            conferir com a contabilidade; Personalizada = 1 = ajustada no cliente).
DATA......: 2026-11
***************************/
IF OBJECT_ID('dbo.TbConversaoFiscal') IS NULL
BEGIN
    CREATE TABLE dbo.TbConversaoFiscal (
        Id            INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_TbConversaoFiscal PRIMARY KEY,
        Regime        CHAR(2)       NOT NULL,
        TipoCodigo    VARCHAR(5)    NOT NULL,
        CodigoEntrada VARCHAR(3)    NOT NULL,
        GrupoCFOP     VARCHAR(6)    NOT NULL,
        CFOPVenda     INT           NOT NULL,
        CSTVenda      VARCHAR(3)    NOT NULL,
        Observacao    VARCHAR(200)  NULL,
        Revisar       BIT           NOT NULL CONSTRAINT DF_TbConversaoFiscal_Revisar DEFAULT (0),
        Personalizada BIT           NOT NULL CONSTRAINT DF_TbConversaoFiscal_Personalizada DEFAULT (0),
        Ativo         BIT           NOT NULL CONSTRAINT DF_TbConversaoFiscal_Ativo DEFAULT (1)
    )
    CREATE UNIQUE INDEX UX_TbConversaoFiscal ON dbo.TbConversaoFiscal (Regime, TipoCodigo, CodigoEntrada, GrupoCFOP)
END
GO

-- Personalizada = 1: regra criada ou alterada na tela Fiscal > Conversao CFOP/CST. Scripts de atualizacao das regras
-- padrao devem usar WHERE Personalizada = 0 para nao apagar o ajuste feito no cliente.
IF COL_LENGTH('dbo.TbConversaoFiscal', 'Personalizada') IS NULL
    ALTER TABLE dbo.TbConversaoFiscal ADD Personalizada BIT NOT NULL CONSTRAINT DF_TbConversaoFiscal_Personalizada DEFAULT (0)
GO

-- Regras iniciais (so quando a tabela esta vazia: as ajustadas pela contabilidade nao sao sobrescritas)
IF NOT EXISTS (SELECT 1 FROM dbo.TbConversaoFiscal)
BEGIN
    INSERT INTO dbo.TbConversaoFiscal (Regime, TipoCodigo, CodigoEntrada, GrupoCFOP, CFOPVenda, CSTVenda, Observacao, Revisar) VALUES
    -- ===== SIMPLES NACIONAL (venda com CSOSN) =====
    -- fornecedor do regime normal (CST)
    ('SN', 'CST',   '00', '*',      5102, '102', 'Tributada integralmente -> revenda tributada pelo Simples', 0),
    ('SN', 'CST',   '10', '*',      5405, '500', 'Tributada com ST -> ICMS cobrado anteriormente por ST', 0),
    ('SN', 'CST',   '20', '*',      5102, '102', 'Reducao de base -> revenda tributada pelo Simples', 0),
    ('SN', 'CST',   '30', '*',      5405, '500', 'Isenta/nao tributada com ST -> ICMS cobrado anteriormente por ST', 0),
    ('SN', 'CST',   '40', '*',      5102, '300', 'Isenta -> CSOSN 300 (mantido do sistema; conferir: 300 = imune, 400 = nao tributada)', 1),
    ('SN', 'CST',   '41', '*',      5102, '400', 'Nao tributada -> nao tributada pelo Simples', 0),
    ('SN', 'CST',   '50', '*',      5102, '400', 'Suspensao -> conferir com a contabilidade', 1),
    ('SN', 'CST',   '51', '*',      5102, '102', 'Diferimento do fornecedor nao vale na revenda -> tributada pelo Simples (conferir)', 1),
    ('SN', 'CST',   '60', '*',      5405, '500', 'ICMS cobrado anteriormente por ST', 0),
    ('SN', 'CST',   '70', '*',      5405, '500', 'Reducao de base com ST -> ICMS cobrado anteriormente por ST', 0),
    ('SN', 'CST',   '90', '*',      5102, '900', 'Outras -> CSOSN 900 (mantido do sistema; conferir)', 1),
    -- fornecedor do Simples (CSOSN)
    ('SN', 'CSOSN', '101', '*',     5102, '102', 'Fornecedor do Simples com credito -> revenda tributada pelo Simples', 0),
    ('SN', 'CSOSN', '102', '*',     5102, '102', 'Fornecedor do Simples sem credito -> revenda tributada pelo Simples', 0),
    ('SN', 'CSOSN', '103', '*',     5102, '102', 'Isencao por faixa de receita do fornecedor nao vale na revenda', 0),
    ('SN', 'CSOSN', '201', '*',     5405, '500', 'Fornecedor do Simples com ST -> ICMS cobrado anteriormente por ST', 0),
    ('SN', 'CSOSN', '202', '*',     5405, '500', 'Fornecedor do Simples com ST -> ICMS cobrado anteriormente por ST', 0),
    ('SN', 'CSOSN', '203', '*',     5405, '500', 'Fornecedor do Simples com ST -> ICMS cobrado anteriormente por ST', 0),
    ('SN', 'CSOSN', '300', '*',     5102, '300', 'Imune', 0),
    ('SN', 'CSOSN', '400', '*',     5102, '400', 'Nao tributada pelo Simples (conferir)', 1),
    ('SN', 'CSOSN', '500', '*',     5405, '500', 'ICMS cobrado anteriormente por ST', 0),
    ('SN', 'CSOSN', '900', '*',     5102, '900', 'Outras (conferir)', 1),
    -- CFOP de ST (x401/x403/x405) vence o codigo: CFOP e CST de venda sempre juntos
    ('SN', '*',     '*',  'ST',     5405, '500', 'CFOP do fornecedor com ST -> ICMS cobrado anteriormente por ST', 0),
    ('SN', '*',     '*',  '*',      5102, '102', 'Padrao: revenda tributada pelo Simples', 0),

    -- ===== REGIME NORMAL - Lucro Presumido / Real (venda com CST; origem do XML) =====
    -- fornecedor do regime normal (CST)
    ('RN', 'CST',   '00', '*',      5102, '00', 'Tributada integralmente', 0),
    ('RN', 'CST',   '10', '*',      5405, '60', 'Tributada com ST na compra -> revenda com ICMS cobrado anteriormente por ST', 0),
    ('RN', 'CST',   '20', '*',      5102, '20', 'Reducao de base: so se a reducao tambem vale na saida do seu estado (conferir)', 1),
    ('RN', 'CST',   '30', '*',      5405, '60', 'Isenta com ST na compra -> ICMS cobrado anteriormente por ST (conferir)', 1),
    ('RN', 'CST',   '40', '*',      5102, '40', 'Isenta (conferir se a isencao vale na saida)', 1),
    ('RN', 'CST',   '41', '*',      5102, '41', 'Nao tributada', 0),
    ('RN', 'CST',   '50', '*',      5102, '00', 'Suspensao do fornecedor nao vale na revenda -> tributada (conferir)', 1),
    ('RN', 'CST',   '51', '*',      5102, '00', 'Diferimento do fornecedor nao vale na revenda -> tributada (conferir)', 1),
    ('RN', 'CST',   '60', '*',      5405, '60', 'ICMS cobrado anteriormente por ST', 0),
    ('RN', 'CST',   '70', '*',      5405, '60', 'Reducao de base com ST -> ICMS cobrado anteriormente por ST', 0),
    ('RN', 'CST',   '90', '*',      5102, '00', 'Outras -> tributada (conferir)', 1),
    -- fornecedor do Simples (CSOSN)
    ('RN', 'CSOSN', '101', '*',     5102, '00', 'Compra de fornecedor do Simples -> revenda tributada', 0),
    ('RN', 'CSOSN', '102', '*',     5102, '00', 'Compra de fornecedor do Simples -> revenda tributada', 0),
    ('RN', 'CSOSN', '103', '*',     5102, '00', 'Compra de fornecedor do Simples -> revenda tributada', 0),
    ('RN', 'CSOSN', '201', '*',     5405, '60', 'Fornecedor do Simples com ST -> ICMS cobrado anteriormente por ST', 0),
    ('RN', 'CSOSN', '202', '*',     5405, '60', 'Fornecedor do Simples com ST -> ICMS cobrado anteriormente por ST', 0),
    ('RN', 'CSOSN', '203', '*',     5405, '60', 'Fornecedor do Simples com ST -> ICMS cobrado anteriormente por ST', 0),
    ('RN', 'CSOSN', '300', '*',     5102, '41', 'Imune -> nao tributada (conferir)', 1),
    ('RN', 'CSOSN', '400', '*',     5102, '00', 'Nao tributada pelo Simples do fornecedor -> revenda tributada (conferir)', 1),
    ('RN', 'CSOSN', '500', '*',     5405, '60', 'ICMS cobrado anteriormente por ST', 0),
    ('RN', 'CSOSN', '900', '*',     5102, '00', 'Outras -> tributada (conferir)', 1),
    ('RN', '*',     '*',  'ST',     5405, '60', 'CFOP do fornecedor com ST -> ICMS cobrado anteriormente por ST', 0),
    ('RN', '*',     '*',  '*',      5102, '00', 'Padrao: revenda tributada', 0)
END
GO

-- Conversao para a venda: @crt = empresa.CRT; @cstxml = CST gravado da nota de entrada (origem + CST, 3 digitos, ou
-- origem + CSOSN, 4 digitos; sem origem tambem serve); @cfopxml = CFOP do item no XML do fornecedor.
-- Devolve CFOPVenda e CSTVenda prontos para o cadastro do produto (regime normal ja com a origem do XML na frente).
IF OBJECT_ID('dbo.ConversaoFiscalVenda') IS NOT NULL
    DROP FUNCTION dbo.ConversaoFiscalVenda
GO
CREATE FUNCTION dbo.ConversaoFiscalVenda(@crt INT, @cstxml VARCHAR(4), @cfopxml INT)
RETURNS TABLE
AS
RETURN
SELECT TOP 1 c.Id, c.CFOPVenda,
       CAST(CASE WHEN c.Regime = 'RN' THEN x.Origem + c.CSTVenda ELSE c.CSTVenda END AS VARCHAR(3)) AS CSTVenda,
       c.Observacao, c.Revisar
FROM dbo.TbConversaoFiscal c
CROSS JOIN (SELECT y.Tipo AS TipoCodigo,
                   CASE WHEN LEN(y.s) = 4 THEN RIGHT(y.s, 3) WHEN y.Tipo = 'CSOSN' THEN y.s ELSE RIGHT(y.s, 2) END AS Codigo,
                   CASE WHEN LEN(y.s) = 4 OR (LEN(y.s) = 3 AND y.Tipo = 'CST') THEN LEFT(y.s, 1) ELSE '0' END AS Origem,
                   CASE WHEN ISNULL(@cfopxml, 0) % 1000 IN (401, 403, 405) THEN 'ST' ELSE 'NORMAL' END AS Grupo
            FROM (SELECT z.s,
                         -- 4 digitos = origem + CSOSN (como a importacao grava); 3 digitos = origem + CST, salvo quando os 2
                         -- ultimos nao sao CST (101, 102, 103, 201...) ou a origem seria 9 (900): ai e CSOSN sem origem
                         CASE WHEN LEN(z.s) = 4 THEN 'CSOSN'
                              WHEN LEN(z.s) = 3 AND (RIGHT(z.s, 2) NOT IN ('00','10','20','30','40','41','50','51','60','70','90') OR LEFT(z.s, 1) = '9') THEN 'CSOSN'
                              ELSE 'CST' END AS Tipo
                  FROM (SELECT LTRIM(RTRIM(ISNULL(@cstxml, ''))) AS s) z) y) x
WHERE c.Ativo = 1
  AND c.Regime = CASE WHEN ISNULL(@crt, 0) IN (2, 3) THEN 'RN' ELSE 'SN' END
  AND (c.TipoCodigo = x.TipoCodigo OR c.TipoCodigo = '*')
  AND (c.CodigoEntrada = x.Codigo OR c.CodigoEntrada = '*')
  AND (c.GrupoCFOP = x.Grupo OR c.GrupoCFOP = '*')
ORDER BY CASE WHEN c.CodigoEntrada <> '*' AND c.GrupoCFOP <> '*' THEN 1
              WHEN c.GrupoCFOP <> '*' THEN 2
              WHEN c.CodigoEntrada <> '*' THEN 3
              ELSE 4 END
GO

-- '' = o CFOP e o CST/CSOSN de VENDA do produto combinam com o regime da empresa (@crt = empresa.CRT; 2 e 3 = regime
-- normal, os demais = Simples); senao, o motivo. Usada no cadastro de produto (aviso), na entrada de estoque (so troca
-- o CFOP/CST de produto existente quando ele esta incoerente) e na conferencia do SPED.
IF OBJECT_ID('dbo.CFOPCSTCoerente') IS NOT NULL
    DROP FUNCTION dbo.CFOPCSTCoerente
GO
CREATE FUNCTION dbo.CFOPCSTCoerente(@crt INT, @cfop INT, @cst VARCHAR(4))
RETURNS VARCHAR(200)
AS
BEGIN
    DECLARE @c VARCHAR(4), @cod VARCHAR(3), @st BIT, @cfopST BIT, @sCFOP VARCHAR(10)
    SET @c = LTRIM(RTRIM(ISNULL(@cst, '')))
    SET @cfop = ISNULL(@cfop, 0)
    SET @sCFOP = CAST(@cfop AS VARCHAR(10))
    IF @cfop = 0 RETURN 'CFOP de venda não informado.'
    IF @cfop < 5000 OR @cfop > 7999 RETURN 'CFOP ' + @sCFOP + ' é de entrada (compra), não de venda.'
    IF @c = '' RETURN 'CST/CSOSN não informado.'
    SET @cfopST = CASE WHEN @cfop % 1000 IN (401, 403, 405) THEN 1 ELSE 0 END

    IF ISNULL(@crt, 0) NOT IN (2, 3)
    BEGIN
        IF @c NOT IN ('101', '102', '103', '201', '202', '203', '300', '400', '500', '900')
            RETURN 'No Simples Nacional o produto usa CSOSN (101, 102, 500...), não o CST ' + @c + '.'
        SET @st = CASE WHEN @c IN ('201', '202', '203', '500') THEN 1 ELSE 0 END
        IF @st = 1 AND @cfopST = 0 RETURN 'CSOSN ' + @c + ' (com ST) pede CFOP 5405, não ' + @sCFOP + '.'
        IF @st = 0 AND @cfopST = 1 RETURN 'CFOP ' + @sCFOP + ' (com ST) pede CSOSN 500, não ' + @c + '.'
    END
    ELSE
    BEGIN
        SET @cod = RIGHT(@c, 2)
        IF LEN(@c) NOT IN (2, 3) OR @cod NOT IN ('00', '02', '10', '15', '20', '30', '40', '41', '50', '51', '53', '60', '61', '70', '90')
           OR (LEN(@c) = 3 AND LEFT(@c, 1) NOT LIKE '[0-8]')
            RETURN 'No regime normal o produto usa CST (origem + 2 dígitos, ex.: 000, 060), não ' + @c + '.'
        SET @st = CASE WHEN @cod IN ('10', '30', '60', '70') THEN 1 ELSE 0 END
        IF @st = 1 AND @cfopST = 0 RETURN 'CST ' + @c + ' (com ST) pede CFOP 5405, não ' + @sCFOP + '.'
        IF @st = 0 AND @cfopST = 1 RETURN 'CFOP ' + @sCFOP + ' (com ST) pede CST 60, não ' + @c + '.'
    END
    RETURN ''
END
GO
