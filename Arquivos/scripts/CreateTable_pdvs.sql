-- Tabela usada por ListarPDV.frm (botao cmdAvanVendaTransferir do PDV.frm) para listar os terminais
-- disponiveis na transferencia de venda entre caixas. "descricao" e gravado direto em pedidos.maquina.

-- Se existir uma 'pdvs' de OUTRA origem (tem coluna 'id', que nao faz parte do nosso
-- esquema - o nosso usa 'codigo' IDENTITY), apaga pra deixar recriar do zero certinho.
-- ATENCAO: isso descarta qualquer dado que essa tabela alheia tivesse.
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('pdvs') AND name = 'id')
BEGIN
    DROP TABLE pdvs;
END
GO

IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('pdvs') AND type = 'U')
BEGIN
    CREATE TABLE pdvs (
        codigo    INT IDENTITY(1,1) PRIMARY KEY,
        descricao VARCHAR(50) NOT NULL,
        ativo     BIT NOT NULL DEFAULT 1
    );
END
GO

INSERT INTO pdvs (descricao, ativo)
SELECT v.descricao, 1
FROM (VALUES ('PDV01'), ('PDV02')) AS v(descricao)
WHERE NOT EXISTS (SELECT 1 FROM pdvs WHERE pdvs.descricao = v.descricao);
GO
