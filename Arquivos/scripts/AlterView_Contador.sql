/***************************
PROJETO...: Online.Info
ANALISTA..: Leandro Santos
DESCRICAO.: View de Contador usada para SPED
DATA......: 2022-03-14
ALTERACAO.: 2023-01-27
***************************/

-- Idempotente: cria um placeholder se a view ainda nao existir, depois sempre roda o
-- ALTER com a definicao final - funciona tanto na 1a instalacao (view nao existe) quanto
-- em bases que ja tinham a view (so atualiza a definicao). Depende de TbContabilista.CNPJ/
-- CPF ja existirem - rodar CreateTable_TbContabilista.sql (112) antes deste.
IF NOT EXISTS (SELECT 1 FROM sys.views WHERE name = 'Contador')
    EXEC('CREATE VIEW Contador AS SELECT 1 AS Placeholder');
GO

ALTER VIEW Contador AS

SELECT NomeContabilista AS ContadorNome, CNPJ AS CPF_CNPJ, CNPJ AS ContadorCNPJ, CPF AS ContadorCPF, CRC AS ContadorCRC, Endereco AS ContadorEndereco, Num AS ContadorEndNumero,
       Bairro AS ContadorBairro, Complemento AS ContadorEndComplemento, CEP AS ContadorCEP, UF AS ContadorUF, Cidade AS ContadorCidade, CodigoIBGE AS ContadorCodigoIBGE,
       Fone AS ContadorTelefone, FAX AS ContadorFax, Email AS ContadorEmail
FROM TbContabilista
GO
