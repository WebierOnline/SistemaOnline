-- TbContabilista fica vazia numa base nova - a view Contador (usada no SPED) espera
-- exatamente 1 linha nela. Esse script popula com dados ficticios/placeholder (todas as
-- colunas sao NOT NULL) so pra nao quebrar a geracao do SPED enquanto o dado real nao e
-- preenchido. No processo de atualizacao do cliente, pede-se o e-mail do contador e
-- atualiza-se manualmente (os demais campos tambem, se o cliente informar).
-- Idempotente: so insere se a tabela estiver vazia (ela deve ter sempre 1 linha so).
--
-- CRC e Cidade precisam vir preenchidos (nao so NomeContabilista) - o
-- Contador_Cadastro.frm (cmdSalvar_Click/cmdAlterar_Click/cmdExcluir_Click) bloqueia
-- "FORMULARIO INCOMPLETO" se txtRazao, txtCRC ou cboCidade estiverem em branco, entao
-- deixa-los vazios impedia ate de salvar so o e-mail depois.
IF NOT EXISTS (SELECT 1 FROM TbContabilista)
    INSERT INTO TbContabilista
        (NomeContabilista, CNPJ, CPF, CRC, Endereco, Num, Bairro, Complemento, CEP, UF, Cidade, CodigoIBGE, Fone, Fax, Email)
    VALUES
        ('A DEFINIR', '', '0', 'A DEFINIR', '', '', '', '', '', '', 'A DEFINIR', 0, '', '', '');
GO

-- corrige quem ja rodou a versao anterior do script (CRC/Cidade em branco, bloqueando o
-- cmdSalvar_Click mesmo so pra atualizar o e-mail)
UPDATE TbContabilista SET CRC = 'A DEFINIR' WHERE LTRIM(RTRIM(ISNULL(CRC, ''))) = '';
UPDATE TbContabilista SET Cidade = 'A DEFINIR' WHERE LTRIM(RTRIM(ISNULL(Cidade, ''))) = '';
GO
