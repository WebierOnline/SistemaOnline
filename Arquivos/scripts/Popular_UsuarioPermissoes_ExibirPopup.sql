-- Usuarios.Frm: o ListView1 de permissoes e alimentado direto da tabela
-- Usuario_permissoes (ListarAcesso: "SELECT codigo, permissao FROM
-- usuario_permissoes ORDER BY permissao") - nao ha nada a mudar no form,
-- so falta a LINHA da nova permissao 'EXIBIR POPUP' pra ela aparecer na lista.
-- Usuarios de nivel 1 (Administrador) recebem automaticamente TODAS as
-- permissoes dessa tabela ao salvar (cboNivel_Validate: INSERT...SELECT
-- Codigo FROM Usuario_permissoes), entao essa nova permissao ja fica
-- liberada pra admin sem passo extra.
-- Codigo NAO e identity (confirmado: Msg 515 ao tentar inserir sem valor) -
-- gera manualmente o proximo numero via MAX(Codigo)+1.
-- Idempotente: so insere se ainda nao existir essa permissao.
IF NOT EXISTS (SELECT 1 FROM Usuario_permissoes WHERE permissao = 'EXIBIR POPUP')
BEGIN
    DECLARE @ProximoCodigo INT;
    SELECT @ProximoCodigo = ISNULL(MAX(Codigo), 0) + 1 FROM Usuario_permissoes;

    INSERT INTO Usuario_permissoes (Codigo, permissao) VALUES (@ProximoCodigo, 'EXIBIR POPUP');
END
GO
