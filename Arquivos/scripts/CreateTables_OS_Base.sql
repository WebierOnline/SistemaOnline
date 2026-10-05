-- Cria do zero as tabelas base do modulo de Ordem de Servico (necessarias quando um
-- cliente novo usa OS=SIM mas partiu de uma base sem essas tabelas). Gerado a partir do
-- "Generate Scripts" de um cliente saudavel (cyber_base) e conferido contra o codigo ativo
-- de OrdemServico.frm/OS_Recapadora.frm (ver OrdemServico.vbp para a lista de forms em uso).
-- Idempotente: so cria o que ainda nao existir, nao mexe em tabela/coluna ja existente.
--
-- OS_ControleOleo fica de fora de proposito - ja e criada pelo script 080
-- (CreateTable_OS_ControleOleo.sql), nao duplicar aqui.
--
-- ATENCAO: a tabela gerada pelo SSMS como "OS_Pneus" foi RENOMEADA aqui para
-- "OS_Recapadora_Pneus". No banco de referencia ela tem a constraint PK_OS_Recapadora_Pneus
-- (nome que sobra de um sp_rename antigo so NAQUELE cliente) - o codigo ativo
-- (OS_Recapadora_Pneus.frm, OS_CAD_Servicos_Recapadora.frm) consulta literalmente
-- "OS_Recapadora_Pneus", nunca "OS_Pneus". Criar com o nome errado deixaria essa parte
-- (cadastro de pneus da recapadora) quebrada de novo, do mesmo jeito que a falta da
-- tabela OS quebrou o form inteiro.

IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('OS') AND type = 'U')
BEGIN
    CREATE TABLE OS (
        COD_OS          INT NOT NULL PRIMARY KEY,
        STATUS_OS       SMALLINT NULL,
        STATUS          NVARCHAR(35) NULL,
        COD_RESPONSAVEL INT NULL,
        TIPO_OS         NVARCHAR(20) NULL,
        DATA_ENTRADA    DATETIME NULL,
        HORA_ENTRADA    DATETIME NULL,
        DATA_TERMINO    DATETIME NULL,
        HORA_TERMINO    DATETIME NULL,
        COD_FUNCIONARIO INT NULL,
        COD_CLIENTE     INT NULL,
        SUBTOTAL        DECIMAL(16, 2) NULL,
        TIPO_DESC       NVARCHAR(1) NULL,
        VALOR_DESC      DECIMAL(16, 2) NULL,
        TOTAL           DECIMAL(16, 2) NULL,
        TIPO_PAGAMENTO  NVARCHAR(50) NULL,
        PAGAMENTO       NVARCHAR(50) NULL,
        ValorDescReal   DECIMAL(16, 2) NULL,
        COD_PEDIDO      INT NULL,
        OBS             NTEXT NULL,
        ENTRADA         DECIMAL(16, 2) NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('OS_acessorios') AND type = 'U')
BEGIN
    CREATE TABLE OS_acessorios (
        CODIGO    INT NOT NULL,
        ACESSORIO NVARCHAR(40) NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('OS_acessorios_Auto') AND type = 'U')
BEGIN
    CREATE TABLE OS_acessorios_Auto (
        CODIGO        INT NOT NULL,
        COD_OS        INT NULL,
        COD_ACESSORIO INT NULL,
        ACESSORIO     NVARCHAR(40) NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('OS_Equipamento') AND type = 'U')
BEGIN
    CREATE TABLE OS_Equipamento (
        CODIGO          INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        COD_OS          INT NOT NULL,
        EQUIPAMENTO     NVARCHAR(25) NULL,
        FABRICANTE      NVARCHAR(25) NULL,
        MODELO          NVARCHAR(30) NULL,
        PARECER_CLIENTE NTEXT NULL,
        PARECER_TECNICO NTEXT NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('OS_Equipamento_Auto') AND type = 'U')
BEGIN
    CREATE TABLE OS_Equipamento_Auto (
        CODIGO          INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        COD_OS          INT NOT NULL,
        FABRICANTE      NVARCHAR(50) NULL,
        MODELO          NVARCHAR(30) NULL,
        ANO             NVARCHAR(10) NULL,
        PLACA           NVARCHAR(10) NULL,
        KM              NVARCHAR(10) NULL,
        COR             NVARCHAR(10) NULL,
        TANQUE          NVARCHAR(10) NULL,
        PARECER_CLIENTE NTEXT NULL,
        PARECER_TECNICO NTEXT NULL,
        CHASSI          NVARCHAR(20) NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('OS_Fabricante_Caminhao') AND type = 'U')
BEGIN
    CREATE TABLE OS_Fabricante_Caminhao (
        CODIGO     INT NOT NULL,
        FABRICANTE NVARCHAR(15) NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('OS_Fabricante_Moto') AND type = 'U')
BEGIN
    CREATE TABLE OS_Fabricante_Moto (
        CODIGO     INT NOT NULL,
        FABRICANTE NVARCHAR(15) NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('OS_Fabricantes_Carro') AND type = 'U')
BEGIN
    CREATE TABLE OS_Fabricantes_Carro (
        CODIGO     INT NOT NULL,
        FABRICANTE NVARCHAR(15) NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('OS_Modelo_Caminhao') AND type = 'U')
BEGIN
    CREATE TABLE OS_Modelo_Caminhao (
        CODIGO INT NOT NULL,
        MODELO NVARCHAR(15) NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('OS_Modelo_Carro') AND type = 'U')
BEGIN
    CREATE TABLE OS_Modelo_Carro (
        CODIGO INT NOT NULL,
        MODELO NVARCHAR(15) NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('OS_Modelo_Moto') AND type = 'U')
BEGIN
    CREATE TABLE OS_Modelo_Moto (
        CODIGO INT NOT NULL,
        MODELO NVARCHAR(15) NULL
    );
END
GO

-- Nome corrigido (ver nota do cabecalho) - codigo ativo usa "OS_Recapadora_Pneus", nao "OS_Pneus"
IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('OS_Recapadora_Pneus') AND type = 'U')
BEGIN
    CREATE TABLE OS_Recapadora_Pneus (
        CODIGO INT NOT NULL PRIMARY KEY,
        TIPO   NVARCHAR(25) NULL,
        MEDIDA NVARCHAR(15) NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('OS_Servicos') AND type = 'U')
BEGIN
    CREATE TABLE OS_Servicos (
        CODIGO  INT NOT NULL PRIMARY KEY,
        SERVICO NVARCHAR(80) NULL,
        VALOR   DECIMAL(19, 4) NULL,
        TIPO    NVARCHAR(20) NULL,
        MEDIDA  NVARCHAR(10) NULL,
        ARO     NVARCHAR(10) NULL,
        BANDA   NVARCHAR(15) NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('OS_Servicos_Auto') AND type = 'U')
BEGIN
    CREATE TABLE OS_Servicos_Auto (
        codigo       INT NOT NULL PRIMARY KEY,
        cod_os       INT NULL,
        descricao    NVARCHAR(40) NULL,
        preco        DECIMAL(16, 2) NULL,
        quantidade   FLOAT NULL,
        subtotal     DECIMAL(16, 2) NULL,
        desconto     DECIMAL(16, 2) NULL,
        total        DECIMAL(16, 2) NULL,
        data         DATETIME NULL,
        cod_servico  INT NULL,
        cod_mecanico INT NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('OS_Servicos_Comunicacao') AND type = 'U')
BEGIN
    CREATE TABLE OS_Servicos_Comunicacao (
        codigo     INT NOT NULL PRIMARY KEY,
        cod_os     INT NULL,
        descricao  NVARCHAR(40) NULL,
        preco      DECIMAL(16, 2) NULL,
        quantidade DECIMAL(16, 3) NULL,
        subtotal   DECIMAL(16, 2) NULL,
        desconto   DECIMAL(16, 2) NULL,
        total      DECIMAL(16, 2) NULL,
        data       DATETIME NULL,
        Obs        NTEXT NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('OS_Servicos_recapadora') AND type = 'U')
BEGIN
    CREATE TABLE OS_Servicos_recapadora (
        codigo     INT NOT NULL PRIMARY KEY,
        cod_os     INT NULL,
        descricao  NVARCHAR(40) NULL,
        preco      DECIMAL(16, 2) NULL,
        quantidade FLOAT NULL,
        total      DECIMAL(16, 2) NULL,
        data       DATETIME NULL,
        TIPO       NVARCHAR(25) NULL,
        SERIE      NVARCHAR(20) NULL,
        FOGO       NVARCHAR(20) NULL,
        FABRICANTE NVARCHAR(20) NULL,
        MEDIDA     NVARCHAR(10) NULL,
        ARO        NVARCHAR(10) NULL,
        BANDA      NVARCHAR(15) NULL,
        subtotal   DECIMAL(16, 2) NULL,
        desconto   DECIMAL(16, 2) NULL,
        DOTE       NVARCHAR(10) NULL,
        MARCA      NVARCHAR(30) NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('OS_Situacao') AND type = 'U')
BEGIN
    CREATE TABLE OS_Situacao (
        CODIGO   INT NULL,
        SITUACAO NVARCHAR(50) NULL
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('OS_Situacao_Auto') AND type = 'U')
BEGIN
    CREATE TABLE OS_Situacao_Auto (
        CODIGO       INT NULL,
        COD_OS       INT NULL,
        SITUACAO     NVARCHAR(50) NULL,
        COD_SITUACAO INT NULL
    );
END
GO
