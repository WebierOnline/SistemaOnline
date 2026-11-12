/***************************
DESCRICAO.: Controle dos scripts aplicados pelo Atualizador.exe (atualizacao automatica do sistema).
            O atualizador roda, no PC servidor, os scripts do _manifesto.txt que ainda nao estao aqui OU cujo
            conteudo mudou (Hash = SHA-256 do arquivo): os scripts sao idempotentes e alguns sao reeditados no mesmo
            numero. Ao terminar sem erro ele grava em configuracao VERSAO_SISTEMA (versao do pacote instalado) e
            VERSAO_MINIMA_EXE (os terminais abaixo dela sao obrigados a atualizar).
            O proprio atualizador cria a tabela se este script ainda nao rodou; mesma estrutura.
            Marco zero: tabela vazia = na primeira atualizacao automatica todos os scripts rodam de novo (idempotentes).
DATA......: 2026-11
***************************/
IF OBJECT_ID('dbo.ScriptsAplicados') IS NULL
CREATE TABLE dbo.ScriptsAplicados (
    Numero        INT          NOT NULL PRIMARY KEY,
    Arquivo       VARCHAR(200) NOT NULL,
    Hash          VARCHAR(64)  NOT NULL,
    DataAplicacao DATETIME     NOT NULL,
    Versao        VARCHAR(20)  NULL,
    Maquina       VARCHAR(60)  NULL
)
GO