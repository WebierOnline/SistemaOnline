CREATE TABLE TribRegraDifalUF (
    ID INT IDENTITY(1,1) PRIMARY KEY,
    UF_Destino CHAR(2) NOT NULL,
    AliquotaInterna DECIMAL(5,2) NOT NULL, -- Ex: 20.00
    AliquotaFCP DECIMAL(5,2) DEFAULT 0.00, -- Ex: 2.00
    DataInicioVigencia DATE NOT NULL,
    DataFimVigencia DATE NULL,
    
    -- TipoCalculo: 1 para Base Única (Simples), 2 para Base Dupla (Por Dentro)
    TipoCalculo TINYINT NOT NULL DEFAULT 2, 
    
    -- Indica se o FCP deve ser somado à alíquota interna no cálculo da Base Dupla
    FCPCompoeBase BIT NOT NULL DEFAULT 1,
    
    -- Observações para o suporte/faturamento
    Observacao VARCHAR(255)
);

-- Índice para busca rápida por UF e Data (Performance)
CREATE INDEX IX_RegraDifal_UF_Vigencia ON TribRegraDifalUF (UF_Destino, DataInicioVigencia);
