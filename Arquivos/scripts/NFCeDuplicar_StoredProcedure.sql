-- Stored procedure NFCeDuplicar: clona uma NFCe (cabecalho + itens) numa nota nova,
-- trocando so a data/hora de emissao para agora e resetando os campos de status/
-- transmissao (chave, protocolo, enviada/cancelada/inutilizada, etc). Usada pelo botao
-- cmdDuplicar do Consultar_NFCe.frm para os casos: NFCe cancelada, inutilizada, ou em
-- digitacao mas fora do prazo de transmissao da SEFAZ (dia de emissao != hoje).
-- Mesmo padrao de MAX(IdNFProd)/MAX(NumeNota)+1 e geracao de NFeCodigoNota que a
-- NFCeIncluir ja usa hoje (sem lock explicito - mesma premissa de uso mono-usuario
-- por terminal que o resto do fluxo de NFCe ja assume).

IF OBJECT_ID('dbo.NFCeDuplicar') IS NULL
    EXEC('CREATE PROCEDURE dbo.NFCeDuplicar AS BEGIN SET NOCOUNT ON; END')
GO

ALTER PROCEDURE [dbo].[NFCeDuplicar]
    (@IdNFProdOrigem INT, @Usuario VARCHAR(100))
AS
BEGIN
    SET NOCOUNT ON

    IF NOT EXISTS (SELECT 1 FROM TbNFCe WHERE IdNFProd = @IdNFProdOrigem)
    BEGIN
        RAISERROR('NFCeDuplicar: NFCe de origem (IdNFProd = %d) nao encontrada.', 16, 1, @IdNFProdOrigem)
        RETURN
    END

    DECLARE @id INT, @seqNF VARCHAR(7), @hoje DATETIME

    SELECT @id = ISNULL(MAX(CAST(IdNFProd AS BIGINT)), 0) + 1 FROM TbNFCe
    SELECT @seqNF = CAST(ISNULL(MAX(CAST(NumeNota AS BIGINT)), 0) + 1 AS VARCHAR(7)) FROM TbNFCe
    -- DataEmissao/DataSaidaEntrada sao sempre gravadas sem hora (00:00:00) em todo o
    -- sistema (NFCeIncluir usa pedidos.DATA_COMPRA, que tambem e sempre meia-noite) - a
    -- tela Consultar_NFCe filtra por DATA com igualdade exata (DataEmissao = data escolhida,
    -- sem hora), entao gravar GETDATE() com hora quebraria esse filtro
    SELECT @hoje = DATEADD(DAY, DATEDIFF(DAY, 0, GETDATE()), 0)

    BEGIN TRAN

    BEGIN TRY
        INSERT INTO TbNFCe (
            IdNFProd, Origem, Num_OS_VD_Origem, TipoNF, NaturezaOperacao, CFOP, InscEstSubstTrib,
            NomeRazSocial, CPF_CNPJ, Endereco, Num, Bairro, CEP, Municipio, fone, UF, InscEst,
            DataEmissao, DataSaidaEntrada, HoraSaida, BaseCalc_ICMS, BaseCalc_ICSM_Subst, Valor_ICMS,
            Valor_ICMS_Subst, Valor_Frete, Valor_Seguro, OutrasDespesasAces, Valor_IPI, Valor_NF_Prod,
            NomeTrasnportador, Frete_Por_Conta, Placa_Veiculo, UF_Mot_Transp, CPF_CNPJ_Transp,
            Endereco_Transp, Cidade_Transp, UF_Trasnportador, InscEst_Trasnp, Qtde_Trasnp,
            Especie_Transp, Marca_Trasnp, Num_Transp, PesoBruto_Transp, PesoLiq_Transp,
            Linha1, Linha2, Linha3, Linha4, Linha5, NumeNota, SerieNF, Cancelada, DataCancelamento,
            Usuario, DescontoPromocional, NFeIndicadorFormaPagto, NFeFormatoImpressaoDANFe,
            NFeTipoEmissao, NFeCodigoNota, NFCeChaveAcesso, NFCeChaveAcessoAdicional, NFCeProtocolo,
            NFCeProtocoloDataHora, NFCeRecibo, NFCeCancelada, NFCeCanceladaProtocolo,
            NFCeCanceladaJustificativa, NFCeEnviada, NFCeDataHoraContingencia,
            NFCeJustificativaContingencia, IDCliente, NFeFinalidadeEmissao, NFCeChaveAcessoReferenciada,
            ValorImpostoImportacao, CodigoPais, NomePais, CRT, OrigemDestinatario,
            NFeIdentificadorDestino, NFeConsumidorFinal, NFeIndicadorPresencaComprador,
            NFeIndicadorIEDestinatario, XMLAutorizado, URLQRCode, MonitorStatus, UltimocStat,
            UltimoRetornoWS, UltimoRetornoXML, Inutilizada, CodigoInutilizacao,
            vIBSUF, vIBSMun, vIBS, vCBS, vBCIS, vIS, vBCCBS, vBCIBS, vPIS, vCOFINS
        )
        SELECT
            @id, Origem, Num_OS_VD_Origem, TipoNF, NaturezaOperacao, CFOP, InscEstSubstTrib,
            NomeRazSocial, CPF_CNPJ, Endereco, Num, Bairro, CEP, Municipio, fone, UF, InscEst,
            @hoje, @hoje, GETDATE(), BaseCalc_ICMS, BaseCalc_ICSM_Subst, Valor_ICMS,
            Valor_ICMS_Subst, Valor_Frete, Valor_Seguro, OutrasDespesasAces, Valor_IPI, Valor_NF_Prod,
            NomeTrasnportador, Frete_Por_Conta, Placa_Veiculo, UF_Mot_Transp, CPF_CNPJ_Transp,
            Endereco_Transp, Cidade_Transp, UF_Trasnportador, InscEst_Trasnp, Qtde_Trasnp,
            Especie_Transp, Marca_Trasnp, Num_Transp, PesoBruto_Transp, PesoLiq_Transp,
            Linha1, Linha2, Linha3, Linha4, Linha5, @seqNF, SerieNF, 0, NULL,
            @Usuario, DescontoPromocional, NFeIndicadorFormaPagto, NFeFormatoImpressaoDANFe,
            NFeTipoEmissao, RIGHT(CONVERT(DECIMAL(12, 8), RAND() * (10 - 5) + 5), 8), '', '', 0,
            '', 0, 0, 0,
            '', 0,
            (SELECT CASE WHEN ContigenciaNFCe = 0 THEN '' ELSE CONVERT(VARCHAR(19), GETDATE(), 127) END FROM empresa),
            (SELECT CASE WHEN ContigenciaNFCe = 0 THEN '' ELSE 'Conting�ncia off-line da NFC-e, servidor fora do ar' END FROM empresa),
            IDCliente, NFeFinalidadeEmissao, '',
            ValorImpostoImportacao, CodigoPais, NomePais, CRT, OrigemDestinatario,
            NFeIdentificadorDestino, NFeConsumidorFinal, NFeIndicadorPresencaComprador,
            NFeIndicadorIEDestinatario, '', '', '', 0,
            '', '', 0, 0,
            vIBSUF, vIBSMun, vIBS, vCBS, vBCIS, vIS, vBCCBS, vBCIBS, vPIS, vCOFINS
        FROM TbNFCe
        WHERE IdNFProd = @IdNFProdOrigem

        INSERT INTO TbNFCe_Itens (
            IdNFProd, IdNFProd_Item, CodProduto, IDProduto, CodBarras, DescricaoProduto, CodNcm,
            QtdeMov, ValorUnit, Desconto, DescProdRural, CFOP, IdEstoque, Bc_Icms, Bc_AliquotaReducao,
            Vlr_Icms, Vlr_IPI, Aliq_Icms, Aliq_IPI, ICMSCST, IdLeiDecreto, IPICST, ProdInfAdicional,
            COFINSCST, PISCST, UN, BDNome, BCSTRet, ICMSSTRet, BCImpostoImportacao,
            DespesasAduaneiras, ValorImpostoImportacao, ValorIOF, Valor_Frete, ValorTributos,
            ValorOutras, Nada, vlr_PIS, vlr_COFINS, Aliq_COFINS, Aliq_PIS, TipoProduto,
            cClassTrib, IBSCBS_CST, IBS_vBC, IBS_UFpAliq, IBS_MunpAliq, IBS_pRed, IBS_vIBSUF,
            IBS_vIBSMun, IBS_vIBS, CBS_vBC, CBS_pAliq, CBS_pRed, CBS_vCBS, cClassTrib_IS, IS_CST,
            IS_tipo_calculo, IS_vBC, IS_pAliq, IS_qUnid, IS_vUnid, IS_vIS, uTrib_IS, Valor_Seguro
        )
        SELECT
            @id, IdNFProd_Item, CodProduto, IDProduto, CodBarras, DescricaoProduto, CodNcm,
            QtdeMov, ValorUnit, Desconto, DescProdRural, CFOP, IdEstoque, Bc_Icms, Bc_AliquotaReducao,
            Vlr_Icms, Vlr_IPI, Aliq_Icms, Aliq_IPI, ICMSCST, IdLeiDecreto, IPICST, ProdInfAdicional,
            COFINSCST, PISCST, UN, BDNome, BCSTRet, ICMSSTRet, BCImpostoImportacao,
            DespesasAduaneiras, ValorImpostoImportacao, ValorIOF, Valor_Frete, ValorTributos,
            ValorOutras, Nada, vlr_PIS, vlr_COFINS, Aliq_COFINS, Aliq_PIS, TipoProduto,
            cClassTrib, IBSCBS_CST, IBS_vBC, IBS_UFpAliq, IBS_MunpAliq, IBS_pRed, IBS_vIBSUF,
            IBS_vIBSMun, IBS_vIBS, CBS_vBC, CBS_pAliq, CBS_pRed, CBS_vCBS, cClassTrib_IS, IS_CST,
            IS_tipo_calculo, IS_vBC, IS_pAliq, IS_qUnid, IS_vUnid, IS_vIS, uTrib_IS, Valor_Seguro
        FROM TbNFCe_Itens
        WHERE IdNFProd = @IdNFProdOrigem

        COMMIT TRAN
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRAN
        DECLARE @errmsg NVARCHAR(4000) = ERROR_MESSAGE()
        RAISERROR(@errmsg, 16, 1)
    END CATCH
END
GO
