-- Popup de avisos no login do Tela_Principal (modAvisos.bas + frmAvisoPopup.frm
-- + Meus_Avisos.frm). Usuario_Avisos guarda, por usuario, quais avisos ele
-- marcou "Me avise" (tela Meus_Avisos). So aparece pra quem tiver a permissao
-- 'EXIBIR POPUP' (ver script 127) E tiver o aviso especifico marcado aqui.
-- Idempotente: so cria se a tabela ainda nao existir.
IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'Usuario_Avisos')
BEGIN
    CREATE TABLE Usuario_Avisos (
        Cod_Usuario INT NOT NULL,
        Aviso VARCHAR(30) NOT NULL,
        Ativo BIT NOT NULL DEFAULT 0,
        CONSTRAINT PK_Usuario_Avisos PRIMARY KEY (Cod_Usuario, Aviso)
    );
END
GO
