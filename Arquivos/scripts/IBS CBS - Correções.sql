/* SCRIPT DE ATUALIZAÇÃO GLOBAL - REFORMA TRIBUTÁRIA 
   Este script mapeia os CSOSN/CST atuais para a nova estrutura IBS/CBS.
*/

-- 1. CENÁRIO PARA EMPRESAS DO SIMPLES NACIONAL (CRT <> 3)
IF EXISTS (SELECT 1 FROM empresa WHERE CRT <> 3)
BEGIN
    PRINT 'Atualizando produtos para SIMPLES NACIONAL...';

    -- Regra para Produtos Tributados no Simples (Ex: CSOSN 101, 102, 103, 400, 900)
    UPDATE produtos 
    SET IBSCBSCST = '01', -- Tributado
        CBSpAliq = 0.00,
        IBSUFpAliq = 0.00,
        IBSMunpAliq = 0.00,
        ISCST = '00',     -- Não incidente
        ISpIS = 0.00
    WHERE ICMSCST IN ('101', '102', '103', '400', '900');

    -- Regra para Produtos Monofásicos/Substituição Tributária (Ex: CSOSN 500, 201, 202, 203)
    UPDATE produtos 
    SET IBSCBSCST = '02', -- Monofásico (Equivalente à ST na reforma)
        CBSpAliq = 0.00,
        IBSUFpAliq = 0.00,
        IBSMunpAliq = 0.00,
        ISCST = '00',     -- Padrão não incidente (usuário altera manuamente se for bebida/fumo)
        ISpIS = 0.00
    WHERE ICMSCST IN ('500', '201', '202', '203');
END

-- 2. CENÁRIO PARA EMPRESAS DO REGIME NORMAL (CRT = 3)
ELSE
BEGIN
    PRINT 'Atualizando produtos para REGIME NORMAL...';

    -- Regra para Produtos Tributados Integralmente ou com Redução (CST 00, 20, 90)
    UPDATE produtos 
    SET IBSCBSCST = '01', 
        CBSpAliq = 8.80,   -- Alíquota estimada CBS
        IBSUFpAliq = 17.70, -- Alíquota estimada IBS (Estado + Município)
        IBSMunpAliq = 0.00, -- Geralmente o IBS é gerido pelo estado/conselho
        ISCST = '00',
        ISpIS = 0.00
    WHERE ICMSCST IN ('000', '020', '090');

    -- Regra para Produtos com ST ou Cobrados Anteriormente (CST 10, 60, 70)
    UPDATE produtos 
    SET IBSCBSCST = '02', -- Monofásico
        CBSpAliq = 0.00,
        IBSUFpAliq = 0.00,
        IBSMunpAliq = 0.00,
        ISCST = '00',
        ISpIS = 0.00
    WHERE ICMSCST IN ('010', '060', '070');
END

-- 3. AJUSTE PARA O IMPOSTO SELETIVO (Opcional - Baseado em NCM ou Categoria)
-- Se você quiser automatizar Refrigerantes e Cervejas (Exemplos de NCM)
UPDATE produtos 
SET ISCST = '01', -- Incidência de Imposto Seletivo
    ISpIS = 0.00  -- Alíquota deixamos 0 para revenda
WHERE NCM LIKE '2202%' -- Águas e Refrigerantes
   OR NCM LIKE '2203%'; -- Cervejas