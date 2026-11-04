/***************************
PROJETO...: Online.Info
ANALISTA..: Leandro
DESCRICAO.: Procedure de Processamento de Participantes SPED
DATA......: 2022-03-15
ALTERACAO.: 
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

		SELECT NotaFiscal.CNPJ_CPF AS IDParticipante, NotaFiscal.CNPJ_CPF AS CNPJCPF, NotaFiscal.RazaoSocial, NotaFiscal.RazaoSocial AS Fantasia, NotaFiscal.Logradouro AS Endereco, NotaFiscal.Numero, 
		NotaFiscal.Bairro, NotaFiscal.Municipio AS Cidade, NotaFiscal.UF AS UF, NotaFiscal.CodigoIBGE, NotaFiscal.CEP AS CEP, 
		NotaFiscal.Telefone AS Fone, NotaFiscal.Telefone AS Fax, NotaFiscal.InscricaoEstadual AS IERG, '' AS Email, 1058 AS CodigoPais, 'BRASIL' AS NomePais
		FROM NotaFiscal 
		WHERE Enviada <> 0 AND Cancelada = 0 AND DataEmissao BETWEEN @dataini AND @datafim AND 1 = (CASE WHEN @somenteinventario = 1 THEN 0 ELSE 1 END)
		GROUP BY NotaFiscal.CodigoCorrentista, NotaFiscal.CNPJ_CPF, NotaFiscal.RazaoSocial, 
		NotaFiscal.RazaoSocial, NotaFiscal.Logradouro, NotaFiscal.Bairro, NotaFiscal.Municipio, NotaFiscal.UF, 
		NotaFiscal.CodigoIBGE, NotaFiscal.CEP, NotaFiscal.Telefone, NotaFiscal.InscricaoEstadual, NotaFiscal.Numero
		) tb

	END

ELSE

	BEGIN

		SELECT IDParticipante, CNPJCPF, RazaoSocial, Fantasia, Endereco, Numero, Bairro, Cidade, UF, CodigoIBGE, CEP, Fone, Fax, IERG, Email, CodigoPais, NomePais
		INTO vSPEDParticipantes
		FROM
		(
		SELECT NotaFiscal.CNPJ_CPF AS IDParticipante, NotaFiscal.CNPJ_CPF AS CNPJCPF, NotaFiscal.RazaoSocial, NotaFiscal.RazaoSocial AS Fantasia, NotaFiscal.Logradouro AS Endereco, NotaFiscal.Numero, 
		NotaFiscal.Bairro, NotaFiscal.Municipio AS Cidade, NotaFiscal.UF AS UF, NotaFiscal.CodigoIBGE, NotaFiscal.CEP AS CEP, 
		NotaFiscal.Telefone AS Fone, NotaFiscal.Telefone AS Fax, NotaFiscal.InscricaoEstadual AS IERG, '' AS Email, 1058 AS CodigoPais, 'BRASIL' AS NomePais
		FROM NotaFiscal 
		WHERE Enviada <> 0 AND Cancelada = 0 AND DataEmissao BETWEEN @dataini AND @datafim AND 1 = (CASE WHEN @somenteinventario = 1 THEN 0 ELSE 1 END)
		GROUP BY NotaFiscal.CodigoCorrentista, NotaFiscal.CNPJ_CPF, NotaFiscal.RazaoSocial, 
		NotaFiscal.RazaoSocial, NotaFiscal.Logradouro, NotaFiscal.Bairro, NotaFiscal.Municipio, NotaFiscal.UF, 
		NotaFiscal.CodigoIBGE, NotaFiscal.CEP, NotaFiscal.Telefone, NotaFiscal.InscricaoEstadual, NotaFiscal.Numero
		) tb

	END
GO
