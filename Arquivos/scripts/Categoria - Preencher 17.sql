-- Categorias padrao do tipo de empresa 17 (Vidracaria, Marmoraria e Esquadrias): ramo unificado.
-- Idempotente: so insere a categoria que ainda nao existe para o tipo.
INSERT INTO Categorias (Categoria, Tipo_Empresa)
SELECT s.Categoria, s.Tipo_Empresa
FROM (VALUES
  ('VIDROS COMUNS (FLOAT)', 17),
  ('VIDROS TEMPERADOS', 17),
  ('VIDROS LAMINADOS', 17),
  ('VIDROS INSULADOS', 17),
  ('VIDROS ESPECIAIS (ARAMADO/IMPRESSO)', 17),
  ('ESPELHOS', 17),
  ('BOX PARA BANHEIRO', 17),
  ('GUARDA-CORPO E FECHAMENTOS', 17),
  ('FERRAGENS PARA VIDRO', 17),
  ('MÁRMORES', 17),
  ('GRANITOS', 17),
  ('QUARTZO E PEDRAS SINTÉTICAS', 17),
  ('BANCADAS, PIAS E LAVATÓRIOS', 17),
  ('SOLEIRAS, PEITORIS E RODAPÉS', 17),
  ('REVESTIMENTOS EM PEDRA', 17),
  ('CUBAS E ACESSÓRIOS', 17),
  ('PERFIS DE ALUMÍNIO', 17),
  ('JANELAS E PORTAS DE ALUMÍNIO', 17),
  ('ESQUADRIAS DE PVC', 17),
  ('PORTÕES E GRADES', 17),
  ('TELAS MOSQUITEIRAS', 17),
  ('FERRAGENS E ACESSÓRIOS PARA ESQUADRIAS', 17),
  ('VEDAÇÕES, BORRACHAS E ESCOVAS', 17),
  ('SILICONES, COLAS E VEDANTES', 17),
  ('ABRASIVOS E DISCOS DE CORTE', 17),
  ('FERRAMENTAS E ACESSÓRIOS', 17),
  ('SERVIÇOS (MEDIÇÃO/INSTALAÇÃO)', 17),
  ('OUTROS', 17)
) AS s (Categoria, Tipo_Empresa)
WHERE NOT EXISTS (
    SELECT 1 FROM Categorias c
    WHERE c.Categoria = s.Categoria AND c.Tipo_Empresa = s.Tipo_Empresa
);
GO
