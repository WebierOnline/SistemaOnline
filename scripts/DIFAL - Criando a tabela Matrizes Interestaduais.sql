CREATE TABLE TribMatrizInterestadual (
    UF_Origem CHAR(2) NOT NULL,
    UF_Destino CHAR(2) NOT NULL,
    AliquotaInterestadual DECIMAL(5,2) NOT NULL,
    PRIMARY KEY (UF_Origem, UF_Destino)
);
