-- Aviso ALERTAS_GERAIS (4o popup do login: contabilidade sem e-mail, meta fiscal, certificado digital).
-- Usuarios nivel 1 (Administrador) ja existentes nao passam pelo cboNivel_LostFocus do Usuarios.Frm, entao
-- ganham o aviso marcado aqui (os 3 avisos antigos vieram do script 129).
-- Idempotente: so insere o que ainda nao existe (nao sobrescreve quem ja desmarcou).

INSERT INTO Usuario_Avisos (Cod_Usuario, Aviso, Ativo)
SELECT u.codigo, 'ALERTAS_GERAIS', 1
FROM Usuario u
WHERE u.nivel = '1'
  AND NOT EXISTS (
      SELECT 1 FROM Usuario_Avisos ua
      WHERE ua.Cod_Usuario = u.codigo AND ua.Aviso = 'ALERTAS_GERAIS'
  );
GO
