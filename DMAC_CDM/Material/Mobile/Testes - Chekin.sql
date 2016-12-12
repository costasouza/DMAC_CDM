Select max(chm_item) as item from chekinmercadoria where chm_loja = 'CD' and chm_fornecedor = '45' and chm_notafiscal = '446426' and chm_serie = 'NE'

Select * from chekinmercadoria where chm_loja = 'CD' and chm_fornecedor = '46' and chm_notafiscal = '446426' and chm_serie = 'NE'
--delete chekinmercadoria

exec SP_ComparaChekinNFCompras 'CD','446426','NE','46'

------------------------------------------

select * from itemnfcompra where ci_quantidade = '20' and ci_fornecedor = '46'
--delete itemnfcompra where ci_quantidade = '20' and ci_fornecedor = '46'