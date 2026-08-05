-- Lista todos os bancos de dados do servidor cujo nome contenha
-- "cyber" ou "financeiro", para confirmar o nome exato (e se existe)
-- do banco usado pelo projeto Financeiro.
-- Rode conectado no servidor .\SQLEXPRESS2008 (aba de conexao do SSMS,
-- pode conectar direto no banco "master").
SELECT name AS NomeBanco, create_date, state_desc
FROM sys.databases
WHERE name LIKE '%cyber%' OR name LIKE '%financeiro%'
ORDER BY name;
