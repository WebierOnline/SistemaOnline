-- Adiciona EANEmbalagem e Fracionamento na tabela Produtos
-- EANEmbalagem: EAN da caixa/fardo (unidade de atacado)
-- Fracionamento: quantas unidades vem dentro dessa embalagem

ALTER TABLE Produtos ADD
    EANEmbalagem  varchar(14)    NOT NULL DEFAULT (''),
    Fracionamento decimal(9,3)   NOT NULL DEFAULT (1);
