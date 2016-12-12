update ItemNFcompra set CI_ValorICMS = DESENV.CI_ValorICMS, CI_ValorICMSST = DESENV.CI_ValorICMSST, CI_PrecoUnitario = DESENV.CI_PrecoUnitario,
CI_Quantidade = DESENV.CI_Quantidade,
CI_CodigoBarra = DESENV.CI_CodigoBarra,
CI_DescricaoFornecedor = DESENV.CI_DescricaoFornecedor,
CI_BaseICMSST = DESENV.CI_BaseICMSST, CI_BaseICMS = DESENV.CI_BaseICMS,
CI_ValorIPI = DESENV.CI_ValorIPI, CI_ValorIVAST = DESENV.CI_ValorIVAST from 
ItemNFcompraTEMP as desenv, ItemNFcompra as dmac where 
desenv.CI_NotaFiscal = dmac.CI_NotaFiscal and 
desenv.CI_Serie = dmac.CI_Serie and 
desenv.CI_Item = dmac.CI_Item AND
dmac.ci_dataentrada > '2015/01/01' and 
dmac.CI_Serie = 'NE' 
AND DMAC.CI_NotaFiscal = 489517


	update ItemNFCompra set CI_Referencia = prb_referencia from itemNFcompra, produtobarras
	where ci_serie = 'NE' and CI_CodigoBarra = prb_codigoBarras and prb_tipocodigo='B'
	AND CI_DataEntrada > '2015/01/01'








SELECT * FROM ItemNFcompraTEMP WHERE ci_notafiscal = '489440' order by ci_item
SELECT * FROM ItemNFcompra WHERE ci_notafiscal = '489440' order by ci_item
--SELECT * INTO ItemNFcompraTEMP from dmac353.desenv_cdm.dbo.ItemNFcompra


select * from  ItemNFCompra where ci_notafiscal = '489440' ORDER BY CI_ITEM