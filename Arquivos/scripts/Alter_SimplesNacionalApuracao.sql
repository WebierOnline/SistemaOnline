-- NT 2026.005: RBT12 e aliquotas do Simples por competencia + anexo do produto.
-- Idempotente. Nao altera nenhum comportamento sozinho: so cria tabelas e colunas.
--
-- SimplesNacionalApuracao        : RBT12 oficial da empresa por competencia (AAAAMM), digitado ou sugerido pelo sistema.
--                                  RBT12 = receita bruta dos 12 meses anteriores ao periodo de apuracao (Res. CGSN 140).
-- SimplesNacionalApuracaoAnexo   : (opcional) pIBSSN/pCBSSN informados pelo contador (PGDAS-D) por anexo e competencia.
--                                  Quando existir, tem prioridade sobre o calculo.
-- Empresa.AnexoSNPadrao          : anexo (I a V) padrao da empresa
-- Produtos.AnexoSN               : anexo do produto/servico (NULL = padrao da empresa)

IF OBJECT_ID('SimplesNacionalApuracao') IS NULL
BEGIN
    CREATE TABLE SimplesNacionalApuracao (
        Competencia   CHAR(6)       NOT NULL,
        RBT12         DECIMAL(15,2) NOT NULL,
        Origem        VARCHAR(10)   NOT NULL CONSTRAINT DF_SNApuracao_Origem DEFAULT 'MANUAL',
        DataGravacao  DATETIME      NOT NULL CONSTRAINT DF_SNApuracao_Data DEFAULT GETDATE(),
        CONSTRAINT PK_SimplesNacionalApuracao PRIMARY KEY (Competencia)
    );
END
GO

IF OBJECT_ID('SimplesNacionalApuracaoAnexo') IS NULL
BEGIN
    CREATE TABLE SimplesNacionalApuracaoAnexo (
        Competencia   CHAR(6)       NOT NULL,
        Anexo         VARCHAR(3)    NOT NULL,
        pIBSSN        DECIMAL(7,4)  NULL,
        pCBSSN        DECIMAL(7,4)  NULL,
        CONSTRAINT PK_SimplesNacionalApuracaoAnexo PRIMARY KEY (Competencia, Anexo)
    );
END
GO

IF COL_LENGTH('Empresa', 'AnexoSNPadrao') IS NULL
    ALTER TABLE Empresa ADD AnexoSNPadrao VARCHAR(3) NOT NULL CONSTRAINT DF_Empresa_AnexoSNPadrao DEFAULT 'I';
GO

IF COL_LENGTH('Produtos', 'AnexoSN') IS NULL
    ALTER TABLE Produtos ADD AnexoSN VARCHAR(3) NULL;
GO
