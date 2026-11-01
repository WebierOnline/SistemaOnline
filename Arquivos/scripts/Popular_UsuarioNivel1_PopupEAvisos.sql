-- Usuarios ja existentes com nivel = 1 (Administrador) nunca passam pelo
-- cboNivel_LostFocus do Usuarios.Frm depois dessa atualizacao (esse codigo
-- so roda quando alguem RESSALVA o nivel na tela) - sem esse script, os
-- administradores dos clientes que ja tinham usuario cadastrado nunca
-- ganhariam a permissao EXIBIR POPUP nem os 3 avisos marcados, mesmo sendo
-- "nivel Total".
-- Idempotente: so insere o que ainda nao existe (nao sobrescreve nada que
-- algum administrador ja tenha desmarcado manualmente depois).

-- 1) Permissao EXIBIR POPUP pra todo usuario nivel 1
INSERT INTO Usuario_Acessos (Cod_Usuario, Cod_Permissao)
SELECT u.codigo, p.Codigo
FROM Usuario u
CROSS JOIN Usuario_permissoes p
WHERE u.nivel = '1'
  AND p.permissao = 'EXIBIR POPUP'
  AND NOT EXISTS (
      SELECT 1 FROM Usuario_Acessos ua
      WHERE ua.Cod_Usuario = u.codigo AND ua.Cod_Permissao = p.Codigo
  );
GO

-- 2) Os 3 avisos (Ativo = 1) pra todo usuario nivel 1
INSERT INTO Usuario_Avisos (Cod_Usuario, Aviso, Ativo)
SELECT u.codigo, av.Aviso, 1
FROM Usuario u
CROSS JOIN (VALUES ('CONTAS_PAGAR'), ('CONTAS_RECEBER'), ('ESTOQUE_BAIXO')) AS av(Aviso)
WHERE u.nivel = '1'
  AND NOT EXISTS (
      SELECT 1 FROM Usuario_Avisos ua
      WHERE ua.Cod_Usuario = u.codigo AND ua.Aviso = av.Aviso
  );
GO
