-- NFS-e: pedido de analise fiscal para cancelamento (evento 101103 do Sistema Nacional).
-- Quando o municipio nao deixa o contribuinte cancelar direto (prazo E0822, valor E0823, tomador E0824), o
-- NFSeEmissor (acao=analisefiscal) pede o cancelamento a prefeitura. A nota continua AUTORIZADA (Situacao 1) ate
-- o deferimento (evento 105104 -> Situacao 3); indeferida (105105) volta a coluna para NULL.
-- AnaliseFiscalSolicitada: data/hora do pedido ainda sem resposta. Idempotente.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

IF OBJECT_ID('dbo.NFSe', 'U') IS NOT NULL AND COL_LENGTH('dbo.NFSe', 'AnaliseFiscalSolicitada') IS NULL
    ALTER TABLE dbo.NFSe ADD AnaliseFiscalSolicitada datetime NULL
GO
