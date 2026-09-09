-- Fix_Categorias_Acento_Interrogacao.sql
-- Deploy antigo (sqlcmd -f 65001 sobre arquivo cp1252) gravou os acentos das categorias-semente como '?'.
-- Corrige a tabela Categorias no lugar (preserva ID_Categoria, entao Categorias_Tags nao quebra)
-- e tambem Produtos.categoria (texto livre, sem FK). Idempotente: 2a execucao nao acha mais '?'.
SET NOCOUNT ON;
GO

IF OBJECT_ID('tempdb..#mapa') IS NOT NULL DROP TABLE #mapa;
CREATE TABLE #mapa (corrompido VARCHAR(50) NOT NULL, correto VARCHAR(50) NOT NULL);
INSERT INTO #mapa (corrompido, correto) VALUES
  ('ACESS?RIOS E BOLSAS', 'ACESSÓRIOS E BOLSAS'),
  ('ALIMENTOS (CESTA B?SICA)', 'ALIMENTOS (CESTA BÁSICA)'),
  ('BEBIDAS (ALCO?LICAS)', 'BEBIDAS (ALCOÓLICAS)'),
  ('BEBIDAS (A?UCARADAS)', 'BEBIDAS (AÇUCARADAS)'),
  ('CAPACITORES E REL?S', 'CAPACITORES E RELÉS'),
  ('COMBUST?VEIS', 'COMBUSTÍVEIS'),
  ('CONVENI?NCIA', 'CONVENIÊNCIA'),
  ('DERMOCOSM?TICOS', 'DERMOCOSMÉTICOS'),
  ('ELETRODOM?STICOS', 'ELETRODOMÉSTICOS'),
  ('EL?TRICA', 'ELÉTRICA'),
  ('FERRAMENTAS (V?CUO/MANIFOLD)', 'FERRAMENTAS (VÁCUO/MANIFOLD)'),
  ('FRIOS E LATIC?NIOS', 'FRIOS E LATICÍNIOS'),
  ('GELO E CARV?O', 'GELO E CARVÃO'),
  ('GEN?RICOS', 'GENÉRICOS'),
  ('G?S REFRIGERANTE', 'GÁS REFRIGERANTE'),
  ('HIDR?ULICA', 'HIDRÁULICA'),
  ('HORTIFR?TI', 'HORTIFRÚTI'),
  ('ILUMINA??O', 'ILUMINAÇÃO'),
  ('ILUMINA??O/LUMIN?RIAS', 'ILUMINAÇÃO/LUMINÁRIAS'),
  ('LANCHES E SANDU?CHES', 'LANCHES E SANDUÍCHES'),
  ('LATARIA E ACESS?RIOS', 'LATARIA E ACESSÓRIOS'),
  ('LATIC?NIOS', 'LATICÍNIOS'),
  ('LOU?AS E METAIS', 'LOUÇAS E METAIS'),
  ('MANUTEN??O', 'MANUTENÇÃO'),
  ('MATERIAL EL?TRICO', 'MATERIAL ELÉTRICO'),
  ('MODA ?NTIMA', 'MODA ÍNTIMA'),
  ('MOTOR E TRANSMISS?O', 'MOTOR E TRANSMISSÃO'),
  ('PERFUMARIA E COSM?TICOS', 'PERFUMARIA E COSMÉTICOS'),
  ('PE?AS GERAIS', 'PEÇAS GERAIS'),
  ('PLACAS ELETR?NICAS', 'PLACAS ELETRÔNICAS'),
  ('POR??ES E PETISCOS', 'PORÇÕES E PETISCOS'),
  ('SERVI?OS', 'SERVIÇOS'),
  ('SISTEMA EL?TRICO', 'SISTEMA ELÉTRICO'),
  ('SUPORTES E M?O FRANCESA', 'SUPORTES E MÃO FRANCESA'),
  ('SUSPENS?O E AMORTECEDORES', 'SUSPENSÃO E AMORTECEDORES'),
  ('TINTAS E ACESS?RIOS', 'TINTAS E ACESSÓRIOS'),
  ('TUBULA??O E ISOLAMENTO', 'TUBULAÇÃO E ISOLAMENTO'),
  ('?LEOS E LUBRIFICANTES', 'ÓLEOS E LUBRIFICANTES'),
  ('?TICOS', 'ÓTICOS');
GO

-- 1) Categorias: se a versao correta ainda NAO existe pro mesmo Tipo_Empresa -> corrige no lugar
UPDATE c SET c.Categoria = m.correto
FROM Categorias c
INNER JOIN #mapa m ON c.Categoria = m.corrompido
WHERE NOT EXISTS (SELECT 1 FROM Categorias c2
                  WHERE c2.Categoria = m.correto AND c2.Tipo_Empresa = c.Tipo_Empresa);
GO

-- 2) Categorias: se a correta JA existe (dup), religa tags da corrompida pra correta e apaga a corrompida
IF OBJECT_ID('Categorias_Tags', 'U') IS NOT NULL
BEGIN
    UPDATE t SET t.ID_Categoria = cok.ID_Categoria
    FROM Categorias_Tags t
    INNER JOIN Categorias cbad ON t.ID_Categoria = cbad.ID_Categoria
    INNER JOIN #mapa m ON cbad.Categoria = m.corrompido
    INNER JOIN Categorias cok ON cok.Categoria = m.correto AND cok.Tipo_Empresa = cbad.Tipo_Empresa
    WHERE NOT EXISTS (SELECT 1 FROM Categorias_Tags t2
                      WHERE t2.ID_Categoria = cok.ID_Categoria AND t2.Tags = t.Tags);

    DELETE t FROM Categorias_Tags t
    INNER JOIN Categorias cbad ON t.ID_Categoria = cbad.ID_Categoria
    INNER JOIN #mapa m ON cbad.Categoria = m.corrompido;
END
GO

DELETE c FROM Categorias c
INNER JOIN #mapa m ON c.Categoria = m.corrompido
WHERE EXISTS (SELECT 1 FROM Categorias c2
             WHERE c2.Categoria = m.correto AND c2.Tipo_Empresa = c.Tipo_Empresa);
GO

-- 3) Produtos.categoria (texto livre): alinha com o nome correto
IF COL_LENGTH('Produtos', 'categoria') IS NOT NULL
    UPDATE p SET p.categoria = m.correto
    FROM Produtos p
    INNER JOIN #mapa m ON p.categoria = m.corrompido;
GO

DROP TABLE #mapa;
GO
