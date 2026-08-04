IF NOT EXISTS (
    SELECT 1 FROM sys.objects
    WHERE object_id = OBJECT_ID('TribRegraDifalUF') AND type = 'U'
)
BEGIN
    CREATE TABLE TribRegraDifalUF (
        ID INT IDENTITY(1,1) PRIMARY KEY,
        UF_Destino CHAR(2) NOT NULL,
        AliquotaInterna DECIMAL(5,2) NOT NULL, -- Ex: 20.00
        AliquotaFCP DECIMAL(5,2) DEFAULT 0.00, -- Ex: 2.00
        DataInicioVigencia DATE NOT NULL,
        DataFimVigencia DATE NULL,

        -- TipoCalculo: 1 para Base Unica (Simples), 2 para Base Dupla (Por Dentro)
        TipoCalculo TINYINT NOT NULL DEFAULT 2,

        -- Indica se o FCP deve ser somado a aliquota interna no calculo da Base Dupla
        FCPCompoeBase BIT NOT NULL DEFAULT 1,

        -- Observacoes para o suporte/faturamento
        Observacao VARCHAR(255)
    );

    -- indice para busca rapida por UF e Data (Performance)
    CREATE INDEX IX_RegraDifal_UF_Vigencia ON TribRegraDifalUF (UF_Destino, DataInicioVigencia);
END
