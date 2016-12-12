SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





/*
	Procedure para cria‡Æo de Notas Fiscais de Transferˆncia
        SET QUOTED_IDENTIFIER OFF
*/

ALTER      Procedure SP_Cria_NF_Transferencia
	          @RomaneioImpressao    numeric,
	          @DataEmissao	        Char(10)
As
Declare		@AliquotaICMS		Float,
		@Destino		Char(5),
		@NotaFiscal		Int,
		@NumeroItem		SmallInt,
		@AliquotaICMSAnt	Float,
		@DestinoAnt		Char(5),
		@Sequencia		Int,
		@Romaneio		Int,
		@RomaneioAnt		Int,
		@HoraString		VarChar(4),
		@Erro			Int,
		@Serie                  Char(2),
                @CNPJDestino            char(15),
                @RItens_Sequencia       int,
		@RItens_NotaFiscal	Int,
		@RItens_NumeroItem	SmallInt,
		@RItens_Referencia	Char(7),
		@RItens_Reserva		Int,
		@RItens_Quantidade	Int,
		@RItens_PrecoUnitario	Float,
		@RItens_PrecoLista	Float,
		@RItens_ValorMercadoria	Float,
		@RItens_AliquotaICMS	Float,
		@RItens_Peso		Float,
		@RItens_PercentualICMS	Float,
		@RItens_Destino		Char(5),
		@RItens_Romaneio	Int,
                @RItens_RomaneioImp     Int,
		@RItens_Origem		Char(5),
                @Soma_PesoBruto         float,
                @Soma_PesoLiquido       float,
                @Soma_TotalNota         float,
                @Soma_ValorMercadorias  float,
                @Soma_BaseIcms          float,
                @Soma_ValorIcms         float		

	--Begin Transaction

        Select 	@HoraString = Convert(VarChar(20), GetDate(), 14)

	Select 	@HoraString = SubString(@HoraString, 1, 2) + SubString(@HoraString, 4, 2)

        Create table #RomaneioImpressao (
                     RIMP_Romaneio    int Not Null,
                     RIMP_RomaneioImp int Not Null,
                     RIMP_LojaO       char(05) Not Null,
                     RIMP_LojaD       char(05) Not Null,
                     RIMP_Referencia  char(07) Not Null,
                     RIMP_Qtde        int Not Null,
                     RIMP_Sequencia   int IDENTITY(1,1) NOT FOR REPLICATION NOT NULL)

	Insert Into #RomaneioImpressao(
		     RIMP_Romaneio,
                     RIMP_RomaneioImp,
                     RIMP_LojaO,
                     RIMP_LojaD,
                     RIMP_Referencia,
                     RIMP_Qtde)
	Select       RO_NumeroRomaneio,
                     RO_RomaneioImpressao,
		     RO_LojaOrigem,
                     RO_LojaDestino,
                     RO_Referencia,
		     Sum(RO_QuantidadeEnviada)
	From 	 Romaneio Where RO_RomaneioImpressao = @RomaneioImpressao and RO_Situacao = 'A'
	Group by RO_NumeroRomaneio,
                 RO_RomaneioImpressao,
		 RO_LojaOrigem,
                 RO_LojaDestino,
                 RO_Referencia,
                 RO_Tipo


        Declare curRomaneioItens Insensitive Cursor For
	Select 	RIMP_Sequencia,
		0,
		0,
		RIMP_Referencia,
		RIMP_Qtde,
		PR_PrecoCusto1,
		PR_PrecoCusto1,
		(PR_PrecoCusto1 * RIMP_Qtde),
                PR_ICMSSaida,
		PR_Peso * RIMP_Qtde,
        	(Case When IE_CodigoReducaoICMS Is Null Then 0 
                When PR_CodigoReducaoICMS = 0 Then 0 Else ((IE_PercentualReducao)) End), 
		RIMP_LojaD,
		RIMP_Romaneio,
                RIMP_RomaneioImp,
		RIMP_LojaO 
	From 	#RomaneioImpressao,
		Produto,
		ICMSEstado 
	Where 	RIMP_Referencia = PR_Referencia and 
		PR_CodigoReducaoICMS *= IE_CodigoReducaoICMS and
		PR_ICMSSaida *= IE_PercentualICMS and
		PR_Situacao = 'A' and
		IE_Estado = 'SP' and
		IE_TipoPessoa = 'J' and
		RIMP_RomaneioImp = @RomaneioImpressao 
       Order By RIMP_LojaO,
		RIMP_LojaD,
		RIMP_Referencia

       select @DestinoAnt = '999',
              @Soma_PesoBruto=0,@Soma_PesoLiquido=0,
              @Soma_TotalNota=0,@Soma_ValorMercadorias=0,
              @Soma_BaseIcms=0,@Soma_ValorIcms=0

       Select @Serie= (select CS_SerieNotaFiscal From ControleCDM)

       Open curRomaneioItens

       Fetch Next From curRomaneioItens Into
             @RItens_Sequencia,@RItens_NotaFiscal,@RItens_NumeroItem,
	     @RItens_Referencia,@RItens_Quantidade,@RItens_PrecoUnitario,
	     @RItens_PrecoLista,@RItens_ValorMercadoria,@RItens_AliquotaICMS,
	     @RItens_Peso,@RItens_PercentualICMS,@RItens_Destino,
	     @RItens_Romaneio,@RItens_RomaneioImp,@RItens_Origem

             While @@Fetch_Status = 0
                  Begin
                    Select @NumeroItem = @NumeroItem + 1
                    IF @DestinoAnt <> @RItens_Destino or @NumeroItem > 10
                    Begin
                        IF @DestinoAnt <> '999'
                           Begin 
                             If @NumeroItem > 10
                                Select @NumeroItem = 10

                             Update CapaNFVenda set VC_PesoBruto=@Soma_PesoBruto,VC_PesoLiquido=@Soma_PesoLiquido,
                                    VC_TotalNota=@Soma_TotalNota,VC_ValorMercadorias=@Soma_ValorMercadorias,
                                    VC_BaseIcms=@Soma_BaseIcms,VC_ValorIcms=@Soma_ValorIcms,VC_Itens=@NumeroItem

                             Select @Soma_PesoBruto=0,@Soma_PesoLiquido=0,
                                    @Soma_TotalNota=0,@Soma_ValorMercadorias=0,
                                    @Soma_BaseIcms=0,@Soma_ValorIcms=0
                          End

                        Update ControleCDM Set
                               CS_NumeroNotaFiscal = (CS_NumeroNotaFiscal + 1)
                        --If @@Error <> 0 Goto Desfaz
		        Select @NotaFiscal = (Select CS_NumeroNotaFiscal From ControleCDM)
                        Select @NumeroItem =  1 
                        Select @DestinoAnt = @RItens_Destino
                        Select @CNPJDestino =(Select LO_CGC from Loja Where LO_Loja = @RItens_Destino) 
-->Capa
                       Insert Into CapaNFVenda (
			       VC_NotaFiscal,VC_Serie,VC_LojaOrigem,VC_CGCLojaDestino,VC_HoraEmissao,VC_AV,
			       VC_Cliente,VC_CondicaoPagamento,VC_CodigoVendedor,VC_LojaDestino,VC_DataEmissao,
			       VC_NumeroPedido,VC_CodigoOperacao,VC_CodigoOperacaoNovo,VC_PesoLiquido,VC_PesoBruto,
			       VC_TotalNota,VC_ValorMercadorias,VC_BaseIcms,VC_AliquotaIcms,VC_ValorIcms,VC_Situacao,
			       VC_EnderecoCliente,VC_TipoNota,VC_LojaVenda,VC_VendedorLojaVenda,VC_ECF,VC_Itens,
                               VC_AVistaReceber,VC_Faturada,VC_Financiada,VC_NotaCredito,VC_Deposito,VC_Cartao,
			       VC_ChequePre,VC_Dinheiro,VC_Cheque,VC_PedidoCliente,VC_ValorIPI,VC_ValorFreteCobrado,
			       VC_ValorFrete,VC_Desconto,VC_PagamentoEntrada,VC_Observacao,
                               VC_DataProcessamento,VC_RomaneioImpressao)
                        Values(@NotaFiscal,@Serie,@RItens_Origem,@CNPJDestino,@HoraString,0,0,0,0,@RItens_Destino,
                               @DataEmissao,0,5152,5152,0,0,0,0,0,0,0,'D',@RItens_RomaneioImp,'T',@RItens_Origem,
                               0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,'0',@DataEmissao,@RomaneioImpressao)

                    end

-->item
		Insert 	Into ItemNFVenda (
			VI_NotaFiscal,VI_Serie,VI_LojaOrigem,VI_lojaDestino,VI_PesoBruto,VI_PesoLiquido,VI_DataEmissao,
			VI_NumeroItem,VI_Referencia,VI_Quantidade,VI_PrecoUnitario,VI_PrecoLista,
			VI_ValorMercadoria,VI_AliquotaIcms,VI_ValorIPI,VI_Situacao,VI_TipoNota,
			VI_Reserva,VI_ValorICMS,VI_Reducao,VI_BaseIcms,VI_AliquotaIPI,VI_CustoMedioUnit,
			VI_PrecoCustoUnit,VI_Desconto)
		Values (@NotaFiscal,@Serie,@RItens_Origem,@RItens_Destino,@RItens_Peso,@RItens_Peso,@DataEmissao,@NumeroItem,
                        @RItens_Referencia,@RItens_Quantidade,
			@RItens_PrecoUnitario,@RItens_PrecoUnitario,@RItens_ValorMercadoria,
                        @RItens_AliquotaICMS,0,'D','T',0,		        Convert(Decimal(14,2), (@RItens_ValorMercadoria 
                        -(@RItens_ValorMercadoria * @RItens_PercentualICMS)/100) 
                        * (@RItens_AliquotaICMS/100)),
		        (@RItens_PercentualICMS),
                        (Case @RItens_AliquotaICMS When 0 Then 0 Else (Convert(Decimal(14,2), 
                        (@RItens_ValorMercadoria - ((@RItens_ValorMercadoria * (@RItens_PercentualICMS)) /100)))) End),
                	0,0,0,0)

                Select @Soma_PesoBruto=(@Soma_PesoBruto+@RItens_Peso)
                Select @Soma_PesoLiquido=(@Soma_PesoLiquido+@RItens_Peso)
                Select @Soma_BaseIcms=(@Soma_BaseIcms+(Case @RItens_AliquotaICMS
                       When 0 Then 0 Else (Convert(Decimal(14,2),(@RItens_ValorMercadoria - ((@RItens_ValorMercadoria 
                       * (@RItens_PercentualICMS)) /100)))) End))
                Select @Soma_ValorIcms=(@Soma_ValorIcms+Convert(Decimal(14,2), (@RItens_ValorMercadoria 
                        -(@RItens_ValorMercadoria * @RItens_PercentualICMS)/100)) 
                               * (@RItens_AliquotaICMS/100))
                Select @Soma_ValorMercadorias=(@Soma_ValorMercadorias+@RItens_ValorMercadoria)
                Select @Soma_TotalNota=@Soma_ValorMercadorias

	
       Fetch Next From curRomaneioItens Into
             @RItens_Sequencia,@RItens_NotaFiscal,@RItens_NumeroItem,
	     @RItens_Referencia,@RItens_Quantidade,@RItens_PrecoUnitario,
	     @RItens_PrecoLista,@RItens_ValorMercadoria,@RItens_AliquotaICMS,
	     @RItens_Peso,@RItens_PercentualICMS,@RItens_Destino,
	     @RItens_Romaneio,@RItens_RomaneioImp,@RItens_Origem
       end  

	Close curRomaneioItens
	Deallocate curRomaneioItens

--------------------------------------------------------------------------------------------------------------------
/*
Desfaz:
	Close curRomaneioItens
	Deallocate curRomaneioItens
        Rollback Transaction
	Return(1)
*/

/*
  select * from romaneio
  exec SP_Cria_NF_Transferencia 2,'2013/07/01'

select top 1 vc_situacao,* from capanfvenda  order by  Vc_NotaFiscal 
select top 1 vc_situacao,* from demeo..capanfvenda  
where vc_tiponota='T' and vc_dataemissao='2013/06/28' and vc_lojaorigem='cd' and vc_serie <>'CT' 
order by  Vc_NotaFiscal

select VI_NumeroItem,* from itemnfvenda order by  Vi_NotaFiscal,VI_NumeroItem  
select * from controlecdm
insert into controlecdm (CS_NumeroPedido,CS_NumeroFormula,CS_NumeroRomaneio,CS_NumeroNotaFiscal,
                         CS_SerieNotaFiscal,CS_Empresa,CS_Versao,CS_UF) values
                         (0,0,0,0,'S8','DE MEO FERRAMENTAS','1.1.01','SP') 
truncate table capanfvenda
truncate table itemnfvenda
select * from controlecdm

*/





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

