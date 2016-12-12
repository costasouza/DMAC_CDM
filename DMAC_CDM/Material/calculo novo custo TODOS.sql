


declare @notafiscal varchar(10)
declare @serie varchar(10)
declare @fornecedor varchar(10)
declare @acao varchar(1)

drop table #tempCAPANFCOMPRA
SELECT * INTO #tempCAPANFCOMPRA FROM CAPANFCOMPRA WHERE  CC_Serie = 'NE' and CC_DATAentrada >= '2015/01/01' ORDER BY CC_DataEntrada

while (select count(*) from #tempCAPANFCOMPRA) > 0
begin
	
	select top 1 @notafiscal = CC_NotaFiscal, @serie = CC_Serie, @fornecedor = CC_Fornecedor, @acao = CC_AcaoEntrada from #tempCAPANFCOMPRA WHERE  CC_Serie = 'NE' and CC_DATAentrada >= '2015/01/01' ORDER BY CC_DataEntrada

	
	PRINT ('SP_ATUALIZA_CUSTOS_ENTRADA_COMPRAS ''' + @notafiscal + ''',''FONR'',''SERI'',''ACAO''')
	--print @notafiscal
	
	exec SP_ATUALIZA_CUSTOS_ENTRADA_COMPRAS @notafiscal, @fornecedor, @serie, @acao
	
	--EXEC SP_calculo_custo_novo_entrada_CDM @notafiscal, @serie, @fornecedor
		
	delete #tempCAPANFCOMPRA where CC_NotaFiscal = @notafiscal and CC_Serie = @serie and CC_Fornecedor = @fornecedor
	
end
