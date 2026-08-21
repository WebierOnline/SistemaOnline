-- Corrige NFCeIncluir pra nao gerar incoerencia quando o pedido tem acrescimo
-- (optAscrescRS/optAscrescPorc/txtAcresc do PDV.frm):
--   1) DescontoPromocional usava (SUBTOTAL - pedidos.TOTAL); com acrescimo e sem desconto,
--      TOTAL fica MAIOR que SUBTOTAL e isso gerava DescontoPromocional NEGATIVO na NFCe -
--      incoerencia que travava a transmissao (por isso existia um bloqueio no PDV pra nao
--      gerar NFCe de venda com acrescimo). Agora usa pedidos.ValorDescReal direto (sempre
--      o valor real do desconto em R$, nunca negativo).
--   2) Acrescenta OutrasDespesasAces (cabecalho TbNFCe) <- pedidos.ValorAcrescReal.
--   3) Acrescenta ValorOutras (item TbNFCe_Itens) <- pedidos_itens.ValorAcrescimo, rateado
--      por RatearAcrescimoItensPedido no PDV.frm (mesmo padrao de rateio/arredondamento
--      que RatearDescontoItensPedido ja usa pro Desconto por item).
-- Resto da procedure identico ao que ja estava rodando (mesma logica/campos que nao mudam).

IF OBJECT_ID('dbo.NFCeIncluir') IS NULL
    EXEC('CREATE PROCEDURE dbo.NFCeIncluir AS BEGIN SET NOCOUNT ON; END')
GO

ALTER PROCEDURE [dbo].[NFCeIncluir](@codpedido INT) AS
DECLARE @msg VARCHAR(200), @id INT, @seqNF as INT

--Desabilita a contagem recuperacao/alteracao de dados
SET NOCOUNT ON

--Opa, tem mensagem, entao retorna em formato de erro
IF (@msg <> '')
 BEGIN
	RAISERROR(@msg, 16, 1)
	RETURN
 END

SELECT @id = ISNULL(MAX(CAST(IdNFProd AS BIGINT)), 0) + 1 FROM TbNFCe

SELECT @seqNF = ISNULL(MAX(CAST(NumeNota AS BIGINT)), 0) + 1 FROM TbNFCe

INSERT INTO [TbNFCe]
           ([IdNFProd]
           ,[Origem]
           ,[Num_OS_VD_Origem]
           ,[TipoNF]
           ,[NaturezaOperacao]
           ,[CFOP]
           ,[NomeRazSocial]
           ,[CPF_CNPJ]
           ,[Municipio]
           ,[UF]
           ,[DataEmissao]
           ,[DataSaidaEntrada]
           ,[HoraSaida]
		   ,[Valor_NF_Prod]
           ,[Frete_Por_Conta]
           ,[NumeNota]
           ,[SerieNF]
           ,[Usuario]
           ,[DescontoPromocional]
           ,[OutrasDespesasAces]
           ,[NFeIndicadorFormaPagto]
           ,[NFeFormatoImpressaoDANFe]
           ,[NFeTipoEmissao]
           ,[NFeCodigoNota]
           ,[IDCliente]
           ,[NFeFinalidadeEmissao]
           ,[CodigoPais]
           ,[NomePais]
           ,[CRT]
           ,[OrigemDestinatario]
           ,[NFeIdentificadorDestino]
           ,[NFeConsumidorFinal]
           ,[NFeIndicadorPresencaComprador]
           ,[NFeIndicadorIEDestinatario]
		   ,[Linha1]
		   ,[NFCeDataHoraContingencia]
		   ,[NFCeJustificativaContingencia])

	 SELECT
	        @id    --<IdNFProd, int,>
           ,'VD'   --<Origem, varchar(2),>
           ,@codpedido  --<Num_OS_VD_Origem, int,>
           ,'S'         --<TipoNF, varchar(1),>
           ,'VENDA DE MERCADORIA'   --<NaturezaOperacao, varchar(60),>
           ,5102        --<CFOP, smallint,>
           ,cliente.nome     --<NomeRazSocial, varchar(70),>
           ,cliente.CPF          --<CPF_CNPJ, varchar(18),>
           ,empresa.CIDADE   --<Municipio, varchar(60),>
           ,empresa.ESTADO   --<UF, varchar(9),>
           ,pedidos.DATA_COMPRA   --<DataEmissao, datetime,>
           ,pedidos.DATA_COMPRA   --<DataSaidaEntrada, datetime,>
           ,getdate()   --<HoraSaida, datetime,>
           ,SUBTOTAL    --<Valor_NF_Prod, decimal(11,2),>
           ,'9 - Sem Frete'   --<Frete_Por_Conta, varchar(38),>
           ,@seqNF    --<NumeNota, int,>
           ,1         --<SerieNF, smallint,>
           ,MAQUINA   --<Usuario, varchar(100),>
           ,ISNULL(pedidos.ValorDescReal, 0)  --<DescontoPromocional, decimal(11,2),>
           ,ISNULL(pedidos.ValorAcrescReal, 0)  --<OutrasDespesasAces, decimal(15,2),>
           ,(CASE WHEN pedidos.TIPO_PAGAMENTO = 'À Vista' THEN '0 - Pagamento à vista' ELSE '1 - Pagamento a prazo' END) --<NFeIndicadorFormaPagto, varchar(21),>
           ,'4 - DANFE NFC-e'                 --<NFeFormatoImpressaoDANFe, varchar(38),>
           ,'1 - Normal'                      --<NFeTipoEmissao, varchar(34),>
           ,Right(CONVERT(DECIMAL(12,8), RAND()*(10-5)+5), 8)  --<NFeCodigoNota, int,>
           ,cliente.codigo                    --<IDCliente, int,>
           ,''   --<NFeFinalidadeEmissao, varchar(27),>
           ,1058   --<CodigoPais, smallint,>
           ,'BRASIL'  --<NomePais, varchar(60),>
           ,empresa.CRT --<CRT, varchar(58),>
           ,'C'   --<OrigemDestinatario, varchar(1),>
           ,'1' --<NFeIdentificadorDestino, varchar(26),>
           ,1   --<NFeConsumidorFinal, bit,>
           ,'1'    --<NFeIndicadorPresencaComprador, varchar(45),>
           ,'9'    --<NFeIndicadorIEDestinatario, varchar(101),>
		   ,''
		   ,(CASE WHEN empresa.ContigenciaNFCe = 0 THEN '' ELSE CONVERT(VARCHAR(19), GETDATE(), 127) END)
		   ,(CASE WHEN empresa.ContigenciaNFCe = 0 THEN '' ELSE 'Contingência off-line da NFC-e, servidor fora do ar' END)
	FROM cliente INNER JOIN pedidos ON cliente.codigo = pedidos.cod_cliente
	     CROSS JOIN empresa
	WHERE pedidos.cod_pedido = @codpedido

		   --Importar Itens
INSERT INTO TbNFCe_Itens (
							IDProduto,
							DescricaoProduto,
							CodBarras,
							CFOP,
							CodNcm,
							ICMSCST,
							UN,
							ValorUnit,
							Desconto,
							QtdeMov,
							Aliq_Icms,
							Bc_Icms,
							Vlr_Icms,
							IdNFProd_Item,
							COFINSCST,
							PISCST,
							IdNFProd, Aliq_PIS, Aliq_COFINS, vlr_COFINS, vlr_PIS, TipoProduto, ValorOutras
    				     )
        SELECT pedidos_itens.cod_produto, produtos.descricao, produtos.EAN, produtos.cfop, produtos.ncm, produtos.icmscst, produtos.unid_medida, pedidos_itens.preco, pedidos_itens.desconto, pedidos_itens.quantidade, produtos.ICMSAliq, (pedidos_itens.Total) as varVBC, ((pedidos_itens.total /100) * produtos.ICMSAliq), pedidos_itens.item, produtos.COFINSCST, produtos.PISCST, @id,
        produtos.PISAliq, produtos.COFINSAliq, ((pedidos_itens.total /100) * produtos.COFINSAliq), ((pedidos_itens.total /100) * produtos.PISAliq), (CASE WHEN produtos.COMBUSTIVEL = 1 THEN 'Combustível' ELSE 'Produto' END), ISNULL(pedidos_itens.ValorAcrescimo, 0)
		FROM produtos
		     INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto
			 INNER JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido
        WHERE pedidos_itens.COD_PEDIDO = @codpedido

--Calcular o ICMS dos produtos
UPDATE    TbNFCe
SET              BaseCalc_ICMS =
                          (SELECT     ISNULL(SUM(Bc_Icms), 0) AS vTotalBCI
                            FROM          TbNFCe_Itens
                            WHERE      (IdNFProd = @id) AND (Aliq_Icms <> '0.00')), Valor_ICMS =
                          (SELECT     ISNULL(SUM(Vlr_Icms), 0) AS vValorICMS
                            FROM          TbNFCe_Itens AS TbNFCe_Itens_1
                            WHERE      (IdNFProd = @id) AND (Aliq_Icms <> '0.00'))
WHERE     (IdNFProd = @id)
GO
