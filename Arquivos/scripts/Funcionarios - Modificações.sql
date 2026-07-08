EXEC sp_RENAME 'funcionario.Comissao_Avista' , 'Comissao_Avista1', 'COLUMN'
EXEC sp_RENAME 'funcionario.Comissao_Prazo' , 'Comissao_Prazo1', 'COLUMN'
EXEC sp_RENAME 'funcionario.Comissao_Recebido' , 'Comissao_Recebido1', 'COLUMN'
EXEC sp_RENAME 'funcionario.Comissao_Servico' , 'Comissao_Servico1', 'COLUMN'

/*** À VISTA ***/

ALTER TABLE funcionario
ADD Valor_Comissao1 decimal(16,2) DEFAULT 0 NOT NULL
GO

ALTER TABLE funcionario
ADD Valor_Comissao2 decimal(16,2) DEFAULT 0 NOT NULL
GO

ALTER TABLE funcionario
ADD Valor_Comissao3 decimal(16,2) DEFAULT 0 NOT NULL
GO

ALTER TABLE funcionario
ADD Comissao_Avista2 decimal(16,2) DEFAULT 0 NOT NULL
GO

ALTER TABLE funcionario
ADD Comissao_Avista3 decimal(16,2) DEFAULT 0 NOT NULL
GO

/*** RECEBIDO ***/
ALTER TABLE funcionario
ADD Valor_ComissaoRec1 decimal(16,2) DEFAULT 0 NOT NULL
GO

ALTER TABLE funcionario
ADD Valor_ComissaoRec2 decimal(16,2) DEFAULT 0 NOT NULL
GO

ALTER TABLE funcionario
ADD Valor_ComissaoRec3 decimal(16,2) DEFAULT 0 NOT NULL
GO

ALTER TABLE funcionario
ADD Comissao_Recebido2 decimal(16,2) DEFAULT 0 NOT NULL
GO

ALTER TABLE funcionario
ADD Comissao_Recebido3 decimal(16,2) DEFAULT 0 NOT NULL
GO

/*** RENOMEAR Valor_Comissao -> Valor_ComissaoAV ***/
EXEC sp_RENAME 'funcionario.Valor_Comissao1', 'Valor_ComissaoAV1', 'COLUMN'
EXEC sp_RENAME 'funcionario.Valor_Comissao2', 'Valor_ComissaoAV2', 'COLUMN'
EXEC sp_RENAME 'funcionario.Valor_Comissao3', 'Valor_ComissaoAV3', 'COLUMN'

/*** REMOVER campos legados ***/
ALTER TABLE funcionario DROP COLUMN comissaoservicos
GO

ALTER TABLE funcionario DROP COLUMN comissaovendas
GO

/*** À PRAZO ***/
ALTER TABLE funcionario
ADD Comissao_Prazo2 decimal(16,2) DEFAULT 0 NOT NULL
GO

ALTER TABLE funcionario
ADD Comissao_Prazo3 decimal(16,2) DEFAULT 0 NOT NULL
GO

ALTER TABLE funcionario
ADD Valor_ComissaoAP1 decimal(16,2) DEFAULT 0 NOT NULL
GO

ALTER TABLE funcionario
ADD Valor_ComissaoAP2 decimal(16,2) DEFAULT 0 NOT NULL
GO

ALTER TABLE funcionario
ADD Valor_ComissaoAP3 decimal(16,2) DEFAULT 0 NOT NULL
GO

/*** SERVIÇOS ***/
ALTER TABLE funcionario
ADD Comissao_Servico2 decimal(16,2) DEFAULT 0 NOT NULL
GO

ALTER TABLE funcionario
ADD Comissao_Servico3 decimal(16,2) DEFAULT 0 NOT NULL
GO

ALTER TABLE funcionario
ADD Valor_ComissaoServ1 decimal(16,2) DEFAULT 0 NOT NULL
GO

ALTER TABLE funcionario
ADD Valor_ComissaoServ2 decimal(16,2) DEFAULT 0 NOT NULL
GO

ALTER TABLE funcionario
ADD Valor_ComissaoServ3 decimal(16,2) DEFAULT 0 NOT NULL
GO
