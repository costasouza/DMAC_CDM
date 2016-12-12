SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO












/*

	Procedure para Verificar critica nos itens de entrada de Compras

*/
--select * from itemNFCompra where CI_NotaFiscal = 13702 AND  CI_dataEntrada > '2013/01/01'



ALTER            Procedure SP_ComparaChekinNFCompras
                 @Loja         char(05),   
		 @NotaFiscal   Int,
		 @Serie	       Char(2),
		 @Fornecedor   SmallInt

As

	Declare 	@Frete			Float,
			@Encargos		Float,
			@Embalagem		Float,
			@FreteCalculado		Float,
			@EncargosCalculado	Float,
			@EmbalagemCalculado	Float,
			@ValorMercadorias	Float,
			@PrecoUnitario		Float,
			@NovoCusto		Float,
			@FretePedido		Float,
			@EncargosPedido		Float,
			@EmbalagemPedido	Float,
			@NovoFrete		Float,
			@NovoEncargos		Float,
			@NovoEmbalagem		Float,
			@Critica		VarChar(255),
			@DiferencaQuantidade	Int,
			@Referencia             Char(7),
			@PrecoPedido		Float,
			@PrecoVenda1		Float,
			@FreteNota		Float,
			@EmbalagemNota		Float,
			@DespFinanNota		Float,
			@QuantidadePedido	Int,
			@QuantidadeNota		Int,
			@AliquotaIPI		Float,
			@IPISobreFrete		Char(1),
			@Pedido			Int,
			@DataEntrada		DateTime,
			@SaldoPedido		Int,
			@ContaSit		Int,
                        @DiferencaPreco         Float,
                        @MontaCalc              VarChar(255),
			@IvaAjustado		Float


Begin
	
	Begin Transaction
           update ChekinMercadoria set CHM_chekinOK='N'
                   where chm_loja= @loja and chm_notafiscal = @NotaFiscal and chm_serie = @Serie 
                         and chm_fornecedor = @fornecedor 

           update ChekinMercadoria set CHM_QuantidadeNF = ci_quantidade from ChekinMercadoria,itemnfcompra
                  where CHM_Loja = ci_loja and CHM_Serie = ci_serie and CHM_Fornecedor = ci_fornecedor and 
                   CHM_Referencia = ci_referencia and 
		   chm_loja= @loja and chm_notafiscal = @NotaFiscal and 
		   chm_serie = @Serie and chm_fornecedor = @fornecedor 

           update ChekinMercadoria set CHM_chekinOK='S'
                   where chm_quantidadechekin = chm_quantidadeNF and
                         chm_loja= @loja and chm_notafiscal = @NotaFiscal and chm_serie = @Serie 
                         and chm_fornecedor = @fornecedor 


Insert into chekinmercadoria (chm_loja, chm_fornecedor, chm_notafiscal, chm_serie, chm_item, chm_codigobarras, 
chm_referencia , CHM_QuantidadeNF, CHM_QuantidadeChekin, chm_tipopedido, chm_situacao) 
select @loja, @fornecedor, @notafiscal, @serie, 0, 0, CI_Referencia, CI_Quantidade, 0, '', 'A' from 
itemnfcompra where CI_fornecedor = @fornecedor and CI_notaFiscal = @notafiscal and CI_Serie= @serie and 
not exists (select * from chekinmercadoria where CI_fornecedor = CHM_fornecedor and 
CI_notaFiscal = CHM_notafiscal and CI_Serie= CHM_serie and CI_referencia = CHM_referencia) 

update chekinMercadoria set CHM_CodigoBarras = prb_codigoBarras from produtobarras, chekinmercadoria
where prb_referencia = chm_referencia and prb_tipoCodigo = 'B'
and chm_loja= @loja and chm_notafiscal = @NotaFiscal and chm_serie = @Serie and chm_fornecedor = @fornecedor

      --select * from produtobarras

      End
       
	Commit Transaction

	Return(0)

Desfaz:

	Rollback Transaction

	Return(1)




/*

	select * from chekinMercadoria
	TRUNCATE TABLE chekinMercadoria

*/











GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

