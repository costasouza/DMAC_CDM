alter Procedure SP_AtualizaNumeroPedido
		 @notafiscal		int,
		 @Serie			char(2),
		 @fornecedor		smallint,
		 @loja			char(5),
		 @numeroPedido		int
As

Begin
	
	--select PI_numeroPedido from itempedido as P, itemNFcompra as N 
	--where p.pi_referencia = n.ci_referencia and p.pi_saldoPedido >= n.ci_quantidade  and p.pi_situacao = 'A'
	--and N.ci_notafiscal = '697' and N.ci_serie = 'NE' and N.ci_fornecedor = '277' and N.ci_loja = 'CD'

	update itemNFcompra set ci_nossoPedido = PI_numeroPedido from itempedido, itemNFcompra
	where pi_referencia = ci_referencia and pi_saldoPedido >= ci_quantidade  and pi_situacao = 'A' and pi_numeroPedido = @numeroPedido
	and ci_notafiscal = @notafiscal and ci_serie = @Serie and ci_fornecedor = @fornecedor and ci_loja = @loja and ci_nossoPedido = 0

End


--exec SP_AtualizaNumeroPedido '4464E','97','CD','70803'
--select * from itemNFcompra where ci_notafiscal = '233521' and ci_serie = 'NE' and ci_fornecedor = '97' and ci_loja = 'CD'
--select * from itempedido where pi_referencia = '0970110' and pi_situacao = 'A' order by pi_saldoPedido
--update itemNFcompra set ci_nossopedido = '0' where ci_notafiscal = '446426' and ci_serie = 'NE' and ci_fornecedor = '45' and ci_loja = 'CD'