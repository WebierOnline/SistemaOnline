/***************************
PROJETO...: Online.Info
ANALISTA..: Leandro
DESCRICAO.: Procedure de Processamento de Participantes SPED
DATA......: 2022-03-15
ALTERACAO.: 2026-11 - clientes das NF-e: CNPJ/CPF e codigo do municipio invalidos na nota saem do cadastro do cliente.
***************************/
-- Idempotente: cria um placeholder se o objeto ainda nao existir e sempre roda o ALTER com a definicao final.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.SPEDParticipantes') IS NULL
    EXEC('CREATE PROCEDURE dbo.SPEDParticipantes AS SELECT 1')
GO

ALTER PROCEDURE dbo.SPEDParticipantes(@dataini DATETIME, @datafim DATETIME, @tipoescrituracao INT, @somenteinventario INT) AS
DECLARE @msg VARCHAR(200), @saldoEst DECIMAL(10,3), @saldoMeta DECIMAL(10,3), @ExisteProd INT

--Desabilita a contagem recuperacao/alteracao de dados
SET NOCOUNT ON

--Valor padrão
SET @msg = ''

--Opa, tem mensagem, entao retorna em formato de erro  
IF (@msg <> '')
 BEGIN
	RAISERROR(@msg, 16, 1)
	RETURN
 END

IF NOT (object_id('vSPEDParticipantes') is null)
BEGIN
    DROP TABLE vSPEDParticipantes
END


--Tipo Escrituração
--1 - EFD Fiscal
--2 - EFD Contribuições
IF (@tipoescrituracao = 1)
	BEGIN

		SELECT IDParticipante, CNPJCPF, RazaoSocial, Fantasia, Endereco, Numero, Bairro, Cidade, UF, CodigoIBGE, CEP, Fone, Fax, IERG, Email, CodigoPais, NomePais
		INTO vSPEDParticipantes
		FROM
		(
		SELECT Fornecedor.CPF AS IDParticipante, Fornecedor.CPF AS CNPJCPF, Fornecedor.RAZAO AS RazaoSocial, Fornecedor.FANTASIA AS Fantasia, Fornecedor.Endereco, Fornecedor.Numero, Fornecedor.Bairro, 
		Fornecedor.Cidade, Fornecedor.Estado AS UF, Fornecedor.CodigoIBGE, Fornecedor.CEP, Fornecedor.Telefone AS Fone, Fornecedor.Telefone AS Fax, Fornecedor.IE AS IERG, Fornecedor.Correio_eletronico AS Email, 1058 AS CodigoPais, 'BRASIL' AS NomePais
		FROM Fornecedor
		WHERE CODIGO in (SELECT CodigoCorrentista FROM EntradaEstoque WHERE EntradaEstoque.ModeloNF = '55' AND DataSaida BETWEEN @dataini AND @datafim GROUP BY CodigoCorrentista) AND 1 = (CASE WHEN @somenteinventario = 1 THEN 0 ELSE 1 END)

		UNION ALL

		--dados gravados na NF-e; CNPJ/CPF e codigo do municipio invalidos na nota (notas antigas) saem do cadastro do cliente
		SELECT nf.Doc AS IDParticipante, nf.Doc AS CNPJCPF, nf.RazaoSocial, nf.RazaoSocial AS Fantasia, nf.Logradouro AS Endereco, nf.Numero,
		nf.Bairro, nf.Municipio AS Cidade, nf.UF AS UF, nf.CodigoIBGE, nf.CEP AS CEP,
		nf.Telefone AS Fone, nf.Telefone AS Fax, nf.InscricaoEstadual AS IERG, '' AS Email, 1058 AS CodigoPais, 'BRASIL' AS NomePais
		FROM (SELECT CASE WHEN LEN(REPLACE(REPLACE(REPLACE(ISNULL(n.CNPJ_CPF, ''), '.', ''), '/', ''), '-', '')) IN (11, 14) THEN CAST(n.CNPJ_CPF AS NVARCHAR(20))
		                  ELSE ISNULL(c.CPF, CAST(n.CNPJ_CPF AS NVARCHAR(20))) END AS Doc,
		             CASE WHEN n.CodigoIBGE BETWEEN 1000000 AND 9999999 THEN n.CodigoIBGE ELSE ISNULL(c.CodigoIBGE, n.CodigoIBGE) END AS CodigoIBGE,
		             n.CodigoCorrentista, n.RazaoSocial, n.Logradouro, n.Numero, n.Bairro, n.Municipio, n.UF, n.CEP, n.Telefone, n.InscricaoEstadual
		      FROM NotaFiscal n LEFT JOIN cliente c ON c.CODIGO = n.CodigoCorrentista
		      WHERE n.Enviada <> 0 AND n.Cancelada = 0 AND n.DataEmissao BETWEEN @dataini AND @datafim AND 1 = (CASE WHEN @somenteinventario = 1 THEN 0 ELSE 1 END)) nf
		GROUP BY nf.CodigoCorrentista, nf.Doc, nf.RazaoSocial, nf.Logradouro, nf.Bairro, nf.Municipio, nf.UF,
		nf.CodigoIBGE, nf.CEP, nf.Telefone, nf.InscricaoEstadual, nf.Numero
		) tb

	END

ELSE

	BEGIN

		SELECT IDParticipante, CNPJCPF, RazaoSocial, Fantasia, Endereco, Numero, Bairro, Cidade, UF, CodigoIBGE, CEP, Fone, Fax, IERG, Email, CodigoPais, NomePais
		INTO vSPEDParticipantes
		FROM
		(
		--dados gravados na NF-e; CNPJ/CPF e codigo do municipio invalidos na nota (notas antigas) saem do cadastro do cliente
		SELECT nf.Doc AS IDParticipante, nf.Doc AS CNPJCPF, nf.RazaoSocial, nf.RazaoSocial AS Fantasia, nf.Logradouro AS Endereco, nf.Numero,
		nf.Bairro, nf.Municipio AS Cidade, nf.UF AS UF, nf.CodigoIBGE, nf.CEP AS CEP,
		nf.Telefone AS Fone, nf.Telefone AS Fax, nf.InscricaoEstadual AS IERG, '' AS Email, 1058 AS CodigoPais, 'BRASIL' AS NomePais
		FROM (SELECT CASE WHEN LEN(REPLACE(REPLACE(REPLACE(ISNULL(n.CNPJ_CPF, ''), '.', ''), '/', ''), '-', '')) IN (11, 14) THEN CAST(n.CNPJ_CPF AS NVARCHAR(20))
		                  ELSE ISNULL(c.CPF, CAST(n.CNPJ_CPF AS NVARCHAR(20))) END AS Doc,
		             CASE WHEN n.CodigoIBGE BETWEEN 1000000 AND 9999999 THEN n.CodigoIBGE ELSE ISNULL(c.CodigoIBGE, n.CodigoIBGE) END AS CodigoIBGE,
		             n.CodigoCorrentista, n.RazaoSocial, n.Logradouro, n.Numero, n.Bairro, n.Municipio, n.UF, n.CEP, n.Telefone, n.InscricaoEstadual
		      FROM NotaFiscal n LEFT JOIN cliente c ON c.CODIGO = n.CodigoCorrentista
		      WHERE n.Enviada <> 0 AND n.Cancelada = 0 AND n.DataEmissao BETWEEN @dataini AND @datafim AND 1 = (CASE WHEN @somenteinventario = 1 THEN 0 ELSE 1 END)) nf
		GROUP BY nf.CodigoCorrentista, nf.Doc, nf.RazaoSocial, nf.Logradouro, nf.Bairro, nf.Municipio, nf.UF,
		nf.CodigoIBGE, nf.CEP, nf.Telefone, nf.InscricaoEstadual, nf.Numero
		) tb

	END
GO
