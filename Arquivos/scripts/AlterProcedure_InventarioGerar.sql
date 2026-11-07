/***************************
PROJETO...: Online.Info
ANALISTA..: Leandro Santos
DESCRICAO.: Procedure para Gerar Inventário de Estoque
DATA......: 2022-03-16
ALTERACAO.: 2026-11 - transacao (cabecalho + itens ou nada), ESTOQUE_FISCAL nulo vira 0, checagem de
            duplicidade que funciona (data nula), responsavel (@fu) e filial (@filial) separados,
            sem tabela fixa vEstoqueConsultaData (leitura direta de Produtos / Produtos_Precos).
            2026-11 - saldo = EstoqueFiscalNaData(data do inventario): saldo-base do ultimo inventario fechado (Abertura)
            + movimento fiscal depois dele ate a data. Sem nenhum inventario fechado ate a data (cliente ainda sem
            Abertura, ou data anterior a ela) usa o ESTOQUE_FISCAL atual, como antes.
            2026-11 - so entram produtos ATIVOS com saldo positivo (zerados e negativos ficam fora do inventario).
***************************/
-- Idempotente: cria um placeholder se o objeto ainda nao existir e sempre roda o ALTER com a definicao final.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.InventarioGerar') IS NULL
    EXEC('CREATE PROCEDURE dbo.InventarioGerar AS SELECT 1')
GO

ALTER PROCEDURE dbo.InventarioGerar(@ano INT, @fu INT, @periodotipo VARCHAR(10), @periodo VARCHAR(2), @datainventario DATETIME, @filial INT = 1) AS
DECLARE @idInv INT, @dt DATETIME, @qt INT, @trans INT, @temBase BIT, @dtDia DATE
DECLARE @msg VARCHAR(200)

--Desabilita a contagem recuperacao/alteracao de dados
SET NOCOUNT ON

--Valor padrão
SET @msg = ''
SET @periodo = ISNULL(@periodo, '')
SET @filial = ISNULL(@filial, 1)

--Anual|Semestral|Trimestral|Bimestral|Mensal|Dia
IF (@periodotipo = 'Anual')
   BEGIN
     SET @dt = CONVERT(DATETIME, CONVERT(VARCHAR(4), @ano) + '-12-31', 102)
   END
ELSE IF (@periodotipo = 'Semestral')    
   BEGIN
     IF (@periodo = 1)
       BEGIN
         SET @dt = CONVERT(DATETIME, CONVERT(VARCHAR(4), @ano) + '-06-30', 102)
       END
     ELSE  
       BEGIN
         SET @dt = CONVERT(DATETIME, CONVERT(VARCHAR(4), @ano) + '-12-31', 102)
       END
   END
ELSE IF (@periodotipo = 'Trimestral')    
   BEGIN
     IF (@periodo = 1)
       BEGIN
         SET @dt = CONVERT(DATETIME, CONVERT(VARCHAR(4), @ano) + '-03-31', 102)
       END
     ELSE IF (@periodo = 2) 
       BEGIN
         SET @dt = CONVERT(DATETIME, CONVERT(VARCHAR(4), @ano) + '-06-30', 102)
       END
     ELSE IF (@periodo = 3) 
       BEGIN
         SET @dt = CONVERT(DATETIME, CONVERT(VARCHAR(4), @ano) + '-09-30', 102)
       END
     ELSE IF (@periodo = 4) 
       BEGIN
         SET @dt = CONVERT(DATETIME, CONVERT(VARCHAR(4), @ano) + '-12-31', 102)
       END
   END
ELSE IF (@periodotipo = 'Bimestral')    
   BEGIN
     IF (@periodo = 1)
       BEGIN
         SET @dt = (CONVERT(DATETIME, CONVERT(VARCHAR(4), @ano) + '-03-01', 102)) - 1
       END
     ELSE IF (@periodo = 2) 
       BEGIN
         SET @dt = CONVERT(DATETIME, CONVERT(VARCHAR(4), @ano) + '-04-30', 102)
       END
     ELSE IF (@periodo = 3) 
       BEGIN
         SET @dt = CONVERT(DATETIME, CONVERT(VARCHAR(4), @ano) + '-06-30', 102)
       END
     ELSE IF (@periodo = 4) 
       BEGIN
         SET @dt = CONVERT(DATETIME, CONVERT(VARCHAR(4), @ano) + '-08-31', 102)
       END
     ELSE IF (@periodo = 5) 
       BEGIN
         SET @dt = CONVERT(DATETIME, CONVERT(VARCHAR(4), @ano) + '-10-31', 102)
       END
     ELSE IF (@periodo = 6) 
       BEGIN
         SET @dt = CONVERT(DATETIME, CONVERT(VARCHAR(4), @ano) + '-12-31', 102)
       END
   END
ELSE IF (@periodotipo = 'Mensal')
   BEGIN
     SELECT @dt = CAST(DATEADD(SECOND, -1, DATEADD(MONTH, DATEDIFF(MONTH, '2000-01-01', CONVERT(DATETIME, CONVERT(VARCHAR(4), @ano) + '-' + @periodo + '-15', 102)) + 1, '2000-01-01')) AS DATE) 
   END   
ELSE
   BEGIN
     SET @dt = @datainventario
   END   

IF (@dt IS NULL)
 BEGIN
   SET @msg = 'Data do Inventário inválida!'
 END

--Ja existe inventario desse tipo/ano/periodo na filial?
IF (@msg = '')
 BEGIN
   SELECT @qt = COUNT(IdInventario) FROM TbInventarios
   WHERE AnoInventario = CONVERT(VARCHAR(4), @ano) AND PeriodoInventarioTipo = @periodotipo AND PeriodoInventario = @periodo
         AND IdFilial = @filial AND (@periodotipo <> 'Dia' OR DataInventario = @dt)

   IF (@qt > 0)
    BEGIN
      SET @msg = 'Já existe um Inventário para o Ano Fiscal informado!'
    END
 END

--Opa, tem mensagem, entao retorna em formato de erro
IF (@msg <> '')
 BEGIN
	RAISERROR(@msg, 16, 1)
	RETURN
 END

--funcao de tabela recebe variavel (expressao como argumento da erro de sintaxe em banco com compatibilidade antiga)
SET @dtDia = CAST(@dt AS DATE)

--Existe inventario fechado (base do estoque fiscal) ate a data do inventario?
SET @temBase = 0
IF EXISTS (SELECT 1 FROM TbInventarios WHERE Fechado = 1 AND DataInventario IS NOT NULL AND CAST(DataInventario AS DATE) <= CAST(@dt AS DATE))
   SET @temBase = 1

--Cabecalho + itens na mesma transacao: se qualquer parte falhar nao sobra inventario vazio
SET @trans = @@TRANCOUNT
BEGIN TRY
   IF (@trans = 0) BEGIN TRANSACTION

   SELECT @idInv = ISNULL(MAX(IdInventario), 0) + 1 FROM TbInventarios WITH (UPDLOCK, HOLDLOCK)

   INSERT INTO TbInventarios(IdInventario, Data, IDFuncionario, AnoInventario, PeriodoInventarioTipo, PeriodoInventario, DataInventario, Fechado, Nada, IdFilial)
   VALUES(@idInv, GETDATE(), @fu, @ano, @periodotipo, @periodo, @dt, 0, '', @filial)

   INSERT INTO TbInventariosItens
              (IdInventario, Seq, IDProduto, NomeProduto, SaldoCalculado, SaldoLancado, MetaCalculado, VlrUnitInvent, TotalInvent)
   SELECT @idInv, ROW_NUMBER() OVER(ORDER BY x.DESCRICAO, x.CODIGO), x.CODIGO, x.DESCRICAO, x.Saldo, x.Saldo, x.Saldo,
          x.Custo, CONVERT(DECIMAL(12,2), x.Saldo * x.Custo)
   FROM (SELECT p.CODIGO, ISNULL(p.DESCRICAO, '') AS DESCRICAO, p.ATIVO,
                CASE WHEN @temBase = 1 THEN ISNULL(s.Saldo, 0) ELSE ISNULL(p.ESTOQUE_FISCAL, 0) END AS Saldo,
                ISNULL((SELECT TOP 1 pr.CUSTO FROM Produtos_Precos pr WHERE pr.COD_PRODUTO = p.CODIGO ORDER BY pr.CODIGO DESC), 0) AS Custo
         FROM Produtos p
         LEFT JOIN dbo.EstoqueFiscalNaData(@dtDia) s ON s.IdProduto = p.CODIGO) x
   WHERE (x.ATIVO = 1) AND (x.Saldo > 0)
   ORDER BY x.DESCRICAO, x.CODIGO

   IF (@trans = 0) COMMIT TRANSACTION
END TRY
BEGIN CATCH
   SET @msg = ERROR_MESSAGE()
   IF (@trans = 0 AND @@TRANCOUNT > 0) ROLLBACK TRANSACTION
   RAISERROR(@msg, 16, 1)
END CATCH
GO
