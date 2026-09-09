-- Recarga segura das categorias-semente.
-- Deploy antigo (sqlcmd -f 65001) gravou os acentos como '?' (ex.: EL?TRICA, HIDR?ULICA).
-- Nenhuma categoria digitada na tela de cadastro tem '?', entao da pra mira-las com seguranca.
-- 1) apaga o vinculo de tag das corrompidas (FK Categorias_Tags.ID_Categoria -> Categorias);
--    numa base nova nao ha tag nenhuma. 2) apaga as categorias corrompidas.
-- 3) insere as sementes que faltarem, sem duplicar e sem tocar em categoria propria do cliente.
-- Idempotente.
SET NOCOUNT ON;
GO

-- Categorias_Tags so existe a partir do script 065; numa instalacao do zero ainda nao existe.
IF OBJECT_ID('Categorias_Tags', 'U') IS NOT NULL
    DELETE FROM Categorias_Tags
    WHERE ID_Categoria IN (SELECT ID_Categoria FROM Categorias WHERE Categoria LIKE '%?%');
GO

DELETE FROM Categorias WHERE Categoria LIKE '%?%';
GO

INSERT INTO Categorias (Categoria, Tipo_Empresa)
SELECT s.Categoria, s.Tipo_Empresa
FROM (VALUES
  ('ALIMENTOS', 1),
  ('ALIMENTOS (CESTA BÁSICA)', 1),
  ('HORTIFRÚTI', 1),
  ('CARNES', 1),
  ('FRIOS E LATICÍNIOS', 1),
  ('PADARIA E CONFEITARIA', 1),
  ('LIMPEZA', 1),
  ('HIGIENE E PERFUMARIA', 1),
  ('BEBIDAS', 1),
  ('BEBIDAS (ALCOÓLICAS)', 1),
  ('BEBIDAS (AÇUCARADAS)', 1),
  ('CONGELADOS', 1),
  ('LATICÍNIOS', 1),
  ('TABACARIA', 1),
  ('PET SHOP', 1),
  ('BAZAR E UTILIDADES', 1),
  ('ÉTICOS', 2),
  ('GENÉRICOS', 2),
  ('SIMILARES', 2),
  ('PERFUMARIA E COSMÉTICOS', 2),
  ('HIGIENE', 2),
  ('CUIDADOS INFANTIS', 2),
  ('DERMOCOSMÉTICOS', 2),
  ('SUPLEMENTOS E VITAMINAS', 2),
  ('PRIMEIROS SOCORROS', 2),
  ('CONVENIÊNCIA', 2),
  ('PRATOS EXECUTIVOS', 3),
  ('LANCHES E SANDUÍCHES', 3),
  ('PORÇÕES E PETISCOS', 3),
  ('PIZZAS', 3),
  ('SALGADOS', 3),
  ('SOBREMESAS', 3),
  ('BEBIDAS QUENTES', 3),
  ('ENTRADAS', 3),
  ('BEBIDAS', 3),
  ('BEBIDAS (ALCOÓLICAS)', 3),
  ('BEBIDAS (AÇUCARADAS)', 3),
  ('TABACARIA', 3),
  ('CALÇADOS MASCULINOS', 4),
  ('CALÇADOS FEMININOS', 4),
  ('CALÇADOS INFANTIS', 4),
  ('MODA MASCULINA', 4),
  ('MODA FEMININA', 4),
  ('MODA ÍNTIMA', 4),
  ('MODA PRAIA', 4),
  ('ACESSÓRIOS E BOLSAS', 4),
  ('CINTOS E CARTEIRAS', 4),
  ('MEIAS', 4),
  ('MOTOR E TRANSMISSÃO', 5),
  ('SUSPENSÃO E AMORTECEDORES', 5),
  ('FREIOS', 5),
  ('SISTEMA ELÉTRICO', 5),
  ('ILUMINAÇÃO', 5),
  ('LATARIA E ACESSÓRIOS', 5),
  ('PNEUS E RODAS', 5),
  ('ÓLEOS E LUBRIFICANTES', 5),
  ('ESCAPAMENTOS', 5),
  ('FERRAMENTAS AUTOMOTIVAS', 5),
  ('ALVENARIA E ESTRUTURA', 6),
  ('HIDRÁULICA', 6),
  ('ELÉTRICA', 6),
  ('PISOS E REVESTIMENTOS', 6),
  ('TINTAS E ACESSÓRIOS', 6),
  ('FERRAGENS', 6),
  ('FERRAMENTAS', 6),
  ('ILUMINAÇÃO/LUMINÁRIAS', 6),
  ('LOUÇAS E METAIS', 6),
  ('MADEIRAS E TELHADOS', 6),
  ('BEBIDAS', 7),
  ('BEBIDAS (ALCOÓLICAS)', 7),
  ('BEBIDAS (AÇUCARADAS)', 7),
  ('CONVENIÊNCIA', 7),
  ('ALIMENTOS', 7),
  ('GELO E CARVÃO', 7),
  ('PETISCOS/SNACKS', 7),
  ('TABACARIA', 7),
  ('COMBUSTÍVEIS', 8),
  ('VASILHAMES', 8),
  ('SERVIÇOS', 8),
  ('PRODUTOS GERAIS', 8),
  ('BRINDES', 8),
  ('MANUTENÇÃO', 8),
  ('OUTROS', 8),
  ('GÁS REFRIGERANTE', 9),
  ('COMPRESSORES', 9),
  ('CAPACITORES E RELÉS', 9),
  ('PLACAS ELETRÔNICAS', 9),
  ('CONTROLES REMOTOS', 9),
  ('FILTROS', 9),
  ('SUPORTES E MÃO FRANCESA', 9),
  ('TUBULAÇÃO E ISOLAMENTO', 9),
  ('DRENOS E BOMBAS', 9),
  ('MATERIAL ELÉTRICO', 9),
  ('FERRAMENTAS (VÁCUO/MANIFOLD)', 9),
  ('EPI', 9),
  ('ELETRODOMÉSTICOS', 9),
  ('SERVIÇOS', 9),
  ('PEÇAS GERAIS', 9),
  ('OUTROS', 9),
  ('DECORAÇÃO', 10),
  ('ARTIGOS PARA FESTA', 10),
  ('BRINQUEDOS', 10),
  ('PAPELARIA', 10),
  ('BOLSAS E ACESSÓRIOS', 10),
  ('CANECAS E COPOS', 10),
  ('QUADROS E MOLDURAS', 10),
  ('VELAS E AROMATIZADORES', 10),
  ('PELÚCIAS', 10),
  ('BIJUTERIAS', 10),
  ('CARTÕES E EMBALAGENS PARA PRESENTE', 10),
  ('UTILIDADES DOMÉSTICAS', 10),
  ('ARTIGOS RELIGIOSOS', 10),
  ('ENFEITES SAZONAIS (NATAL/PÁSCOA)', 10),
  ('OUTROS', 10)
) AS s (Categoria, Tipo_Empresa)
WHERE NOT EXISTS (
    SELECT 1 FROM Categorias c
    WHERE c.Categoria = s.Categoria AND c.Tipo_Empresa = s.Tipo_Empresa
);
GO
