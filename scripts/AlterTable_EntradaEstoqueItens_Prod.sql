-- ============================================================
-- Script: AlterTable_EntradaEstoqueItens_Prod.sql
-- Banco  : SQL Server 2014
-- Tabela : EntradaEstoqueItens
-- Objetivo: Adicionar os campos do grupo <prod> da NF-e que
--           nao existiam na tabela:
--           CEST, indEscala, indTot, nFCI
-- ============================================================

-- CEST: Codigo Especificador da Substituicao Tributaria.
-- Formato: 7 digitos numericos (ex: 0100100).
-- Obrigatorio para produtos sujeitos ao regime de ST ou antecipacao tributaria.
-- Tag XML: <CEST> dentro de <prod>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'CEST'
)
    ALTER TABLE EntradaEstoqueItens ADD CEST VARCHAR(7) NOT NULL DEFAULT '';
GO

-- indEscala: Indicador de Producao em Escala Relevante.
-- Valores: S = Producao em escala relevante (industria de grande porte)
--          N = Producao nao em escala relevante (industrias menores / artesanal)
-- Obrigatorio a partir de 2018 para os segmentos contemplados pelo Convenio ICMS 52/2017.
-- Tag XML: <indEscala> dentro de <prod>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'indEscala'
)
    ALTER TABLE EntradaEstoqueItens ADD indEscala VARCHAR(1) NOT NULL DEFAULT 'S';
GO

-- indTot: Indicador de composicao do valor total da NF-e.
-- Valores: 0 = Item NAO compoe o valor total da NF-e (ex: mercadoria bonificada)
--          1 = Item compoe o valor total da NF-e (caso mais comum)
-- Tag XML: <indTot> dentro de <prod>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'indTot'
)
    ALTER TABLE EntradaEstoqueItens ADD indTot TINYINT NOT NULL DEFAULT 1;
GO

-- nFCI: Numero da Ficha de Conteudo de Importacao (FCI).
-- Formato: GUID (36 caracteres) gerado pela SEFAZ para produtos importados
--          que sao objeto de operacoes interestaduais e sujeitos ao diferencial de ST.
-- Campo opcional - presente apenas para mercadorias importadas com ST interestadual.
-- Tag XML: <nFCI> dentro de <prod>
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('EntradaEstoqueItens') AND name = 'nFCI'
)
    ALTER TABLE EntradaEstoqueItens ADD nFCI VARCHAR(36) NOT NULL DEFAULT '';
GO

-- ============================================================
-- Verificacao final
-- ============================================================
SELECT
    name        AS Campo,
    TYPE_NAME(user_type_id) AS Tipo,
    max_length  AS Tamanho,
    is_nullable AS Nulo,
    object_definition(default_object_id) AS ValorDefault
FROM sys.columns
WHERE object_id = OBJECT_ID('EntradaEstoqueItens')
  AND name IN ('CEST', 'indEscala', 'indTot', 'nFCI')
ORDER BY name;
GO
