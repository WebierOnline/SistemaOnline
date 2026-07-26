-- Tabela de controle de troca de oleo por produto (limite de KM e de prazo em dias)
CREATE TABLE OS_ControleOleo (
    ID           INT IDENTITY(1,1) PRIMARY KEY,
    COD_PRODUTO  INT NOT NULL UNIQUE,
    LIMITE_KM    INT NULL,
    LIMITE_PRAZO INT NULL, -- dias entre trocas
    CONSTRAINT FK_OS_ControleOleo_Produto
        FOREIGN KEY (COD_PRODUTO) REFERENCES Produtos(CODIGO)
);
