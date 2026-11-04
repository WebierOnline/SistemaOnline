/***************************
PROJETO...: Online.Info
ANALISTA..: Leandro
DESCRICAO.: Procedure de Processamento de Unidades SPED
DATA......: 2022-03-15
ALTERACAO.: 
***************************/
-- Idempotente: cria um placeholder se o objeto ainda nao existir e sempre roda o ALTER com a definicao final.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF OBJECT_ID('dbo.SPEDUnidades') IS NULL
    EXEC('CREATE PROCEDURE dbo.SPEDUnidades AS SELECT 1')
GO

ALTER PROCEDURE dbo.SPEDUnidades(@filial INT, @dataini DATETIME, @datafim DATETIME, @inventario INT, @usoconsumo BIT, @imobilizado BIT, @tipoescrituracao INT, @somenteinventario INT, @blococ170saidas INT) AS
DECLARE @msg VARCHAR(200), @saldoEst DECIMAL(10,3), @saldoMeta DECIMAL(10,3), @ExisteProd INT

--Desabilita a contagem recuperacao/alteracao de dados
SET NOCOUNT ON

--Valor padrão
SET @msg = ''

--Opa, tem mensagem, entao retorna em formato de erro  
IF (@msg <> '')
 BEGIN
	RAISERROR(@msg, 16, 1)
	RETURN
 END

IF NOT (object_id('vSPEDUnidades') is null)
BEGIN
    DROP TABLE vSPEDUnidades
END

IF (@somenteinventario = 0)
	BEGIN
	IF (@inventario = 0)
	  BEGIN
		SELECT UND
		INTO vSPEDUnidades
		FROM
		(
		SELECT UND
		FROM NFeEntradasItens
		WHERE UND <> '' AND IDLancamento in (SELECT CodigoNota FROM EntradaEstoque WHERE EntradaEstoque.ModeloNF = '55' AND EntradaEstoque.DataSaida BETWEEN @dataini AND @datafim)
		) tb
		GROUP BY UND
	  END
	ELSE
	  BEGIN
		IF (@usoconsumo = 0 AND @imobilizado = 0)
		  BEGIN
			SELECT UND
			INTO vSPEDUnidades
			FROM
			(
			SELECT UND
			FROM NFeEntradasItens
			WHERE UND <> '' AND IDLancamento in (SELECT CodigoNota FROM EntradaEstoque WHERE EntradaEstoque.ModeloNF = '55' AND EntradaEstoque.DataSaida BETWEEN @dataini AND @datafim)
			UNION ALL
			SELECT UN AS UND
			FROM InventarioGerado
			WHERE MetaCalculado > 0 AND IdInventario = @inventario AND Imobilizado = 0 AND UsoConsumo = 0
			) tb
			GROUP BY UND
		  END 
		ELSE IF (@usoconsumo <> 0 AND @imobilizado = 0)
		  BEGIN
			SELECT UND
			INTO vSPEDUnidades
			FROM
			(
			SELECT UND
			FROM NFeEntradasItens
			WHERE UND <> '' AND IDLancamento in (SELECT CodigoNota FROM EntradaEstoque WHERE EntradaEstoque.ModeloNF = '55' AND EntradaEstoque.DataSaida BETWEEN @dataini AND @datafim)
			UNION ALL
			SELECT UN AS UND
			FROM InventarioGerado
			WHERE MetaCalculado > 0 AND IdInventario = @inventario AND Imobilizado = 0
			) tb
			GROUP BY UND
		  END 
		ELSE IF (@usoconsumo = 0 AND @imobilizado <> 0)
		  BEGIN
			SELECT UND
			INTO vSPEDUnidades
			FROM
			(
			SELECT UND
			FROM NFeEntradasItens
			WHERE UND <> '' AND IDLancamento in (SELECT CodigoNota FROM EntradaEstoque WHERE EntradaEstoque.ModeloNF = '55' AND EntradaEstoque.DataSaida BETWEEN @dataini AND @datafim)
			UNION ALL
			SELECT UN AS UND
			FROM InventarioGerado
			WHERE MetaCalculado > 0 AND IdInventario = @inventario AND UsoConsumo = 0
			) tb
			GROUP BY UND
		  END 
	  END
	END
ELSE
	BEGIN
		IF (@usoconsumo = 0 AND @imobilizado = 0)
		  BEGIN
			SELECT UND
			INTO vSPEDUnidades
			FROM
			(
			SELECT UN AS UND
			FROM InventarioGerado
			WHERE MetaCalculado > 0 AND IdInventario = @inventario AND Imobilizado = 0 AND UsoConsumo = 0
			) tb
			GROUP BY UND
		  END 
		ELSE IF (@usoconsumo <> 0 AND @imobilizado = 0)
		  BEGIN
			SELECT UND
			INTO vSPEDUnidades
			FROM
			(
			SELECT UN AS UND
			FROM InventarioGerado
			WHERE MetaCalculado > 0 AND IdInventario = @inventario AND Imobilizado = 0
			) tb
			GROUP BY UND
		  END 
		ELSE IF (@usoconsumo = 0 AND @imobilizado <> 0)
		  BEGIN
			SELECT UND
			INTO vSPEDUnidades
			FROM
			(
			SELECT UN AS UND
			FROM InventarioGerado
			WHERE MetaCalculado > 0 AND IdInventario = @inventario AND UsoConsumo = 0
			) tb
			GROUP BY UND
		  END 
	END
GO
