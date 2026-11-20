-- Categorias padrao dos tipos de empresa 15 (Papelaria e Armarinhos) e 16 (Agropecuaria e Agroveterinaria).
-- Idempotente: so insere a categoria que ainda nao existe para o tipo.
INSERT INTO Categorias (Categoria, Tipo_Empresa)
SELECT s.Categoria, s.Tipo_Empresa
FROM (VALUES
  ('PAPELARIA ESCOLAR', 15),
  ('MATERIAL DE ESCRITÓRIO', 15),
  ('CADERNOS E BLOCOS', 15),
  ('CANETAS, LÁPIS E MARCADORES', 15),
  ('PAPÉIS E ENVELOPES', 15),
  ('COLAS, FITAS E ADESIVOS', 15),
  ('ARTIGOS PARA ARTESANATO', 15),
  ('TINTAS, PINCÉIS E TELAS', 15),
  ('ARMARINHO (LINHAS E AGULHAS)', 15),
  ('AVIAMENTOS E BOTÕES', 15),
  ('FITAS, RENDAS E LAÇOS', 15),
  ('TECIDOS E ENCHIMENTOS', 15),
  ('LÃS E MALHAS', 15),
  ('MOCHILAS E ESTOJOS', 15),
  ('LIVROS E REVISTAS', 15),
  ('BRINQUEDOS', 15),
  ('ARTIGOS PARA FESTA', 15),
  ('SUPRIMENTOS DE INFORMÁTICA', 15),
  ('PRESENTES E LEMBRANCINHAS', 15),
  ('SERVIÇOS (XEROX/IMPRESSÃO)', 15),
  ('OUTROS', 15),
  ('RAÇÕES', 16),
  ('SUPLEMENTOS E SAL MINERAL', 16),
  ('MEDICAMENTOS VETERINÁRIOS', 16),
  ('VACINAS', 16),
  ('ANTIPARASITÁRIOS', 16),
  ('DEFENSIVOS AGRÍCOLAS', 16),
  ('FERTILIZANTES E ADUBOS', 16),
  ('SEMENTES E MUDAS', 16),
  ('FERRAMENTAS AGRÍCOLAS', 16),
  ('IRRIGAÇÃO E MANGUEIRAS', 16),
  ('CERCAS E ARAMES', 16),
  ('EQUIPAMENTOS DE ORDENHA', 16),
  ('SELARIA E ACESSÓRIOS', 16),
  ('PET SHOP', 16),
  ('AVES E CRIAÇÃO', 16),
  ('AQUICULTURA', 16),
  ('EPI E VESTUÁRIO', 16),
  ('MATERIAL HIDRÁULICO', 16),
  ('LIMPEZA E DESINFECÇÃO', 16),
  ('OUTROS', 16)
) AS s (Categoria, Tipo_Empresa)
WHERE NOT EXISTS (
    SELECT 1 FROM Categorias c
    WHERE c.Categoria = s.Categoria AND c.Tipo_Empresa = s.Tipo_Empresa
);
GO
