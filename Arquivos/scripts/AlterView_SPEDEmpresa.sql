/***************************
PROJETO...: Online.Info
ANALISTA..: Leandro Santos
DESCRICAO.: View de Empresa usada para SPED
DATA......: 2022-03-14
ALTERACAO.: 
***************************/
-- Idempotente: cria um placeholder se o objeto ainda nao existir e sempre roda o ALTER com a definicao final.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.SPEDEmpresa') IS NULL
    EXEC('CREATE VIEW dbo.SPEDEmpresa AS SELECT 1 AS Placeholder')
GO

ALTER VIEW dbo.SPEDEmpresa AS

SELECT 1 AS IdFilial, RAZAO AS RazaoSocial, Fantasia, CNPJ, IE AS InscEst, Endereco, NUmero AS Num, Bairro, Cidade, ESTADO AS UF, CEP, CodigoIBGE, Telefone AS Fone, '' AS Fax, Email, CRT, 
       Perfil AS IndicadorPerfil, 1 AS IndicadorAtividade, 0 AS SPEDBlocoK
FROM empresa
GO
