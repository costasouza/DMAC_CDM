SELECT * FROM capaPedido as a, capaPedido2016 as b 
where a.PC_NumeroPedido = b.PC_NumeroPedido and
a.PC_Situacao <> b.pc_situacao



SELECT a.PI_Situacao,b.pi_situacao,a.PI_NumeroPedido, 
a.PI_Situacao,a.PI_QuantidadePedida,a.PI_QuantidadeEntrada,a.PI_SaldoPedido,'',
b.PI_Situacao,b.PI_QuantidadePedida,b.PI_QuantidadeEntrada,b.PI_SaldoPedido FROM itempedido as a, itempedido2016 as b 
where a.Pi_NumeroPedido = b.Pi_NumeroPedido 
and a.PI_Referencia = b.pi_referencia and
a.Pi_Situacao <> b.pi_situacao
and a.PI_SaldoPedido < 0


select * from capapedido where PC_NumeroPedido = 82899
select * from itempedido where Pi_NumeroPedido = 82899

select CI_Situacao, * from itemnfcompra where CI_Referencia = 0402237





select itempedido.PI_Situacao,  itemnfcompra.ci_quantidade, itempedido.PI_QuantidadeEntrada, itempedido.PI_SaldoPedido,itempedido.* from itemnfcompra,itempedido, capanfcompra
where CI_NossoPedido = Pi_NumeroPedido
and CC_NotaFiscal = CI_NotaFiscal and CC_Serie = CI_Serie and CC_Fornecedor = Ci_Fornecedor
and PI_Referencia = CI_Referencia
AND CI_NOSSOPEDIDO = 82927

SELECT CI_NossoPedido, * FROM itemnfcompra WHERE CI_NossoPedido = 82927 AND CI_Serie = 'NE'
SELECT * FROM itempedido WHERE PI_NumeroPedido = 82927 AND PI_Situacao ='A'

SELECT * FROM ITEMNFCOMPRA WHERE CI_Referencia = '0139637' AND CI_SERIE = 'NE'




select PC_Situacao, * from capaPedido  where PC_NumeroPedido = 82947
select Pi_Situacao, * from itempedido  where Pi_NumeroPedido = 82947
select CI_Situacao,CI_NossoPedido, * from itemnfcompra where CI_NossoPedido = 82947


select  itemnfcompra.ci_quantidade, itempedido.PI_QuantidadeEntrada, itempedido.PI_SaldoPedido,itempedido.* from itemnfcompra,itempedido, capanfcompra
where CI_Situacao in ('T','L') and CI_NossoPedido = Pi_NumeroPedido AND PI_Situacao = 'A' and Cc_DataEntrada > '2016/01/01'
and CC_NotaFiscal = CI_NotaFiscal and CC_Serie = CI_Serie and CC_Fornecedor = Ci_Fornecedor
and PI_Referencia = CI_Referencia


select PI_Situacao,PI_QuantidadePedida,PI_QuantidadeEntrada,PI_SaldoPedido, * from itempedido where PI_QuantidadePedida = PI_QuantidadeEntrada and PI_SaldoPedido < 0

select PI_Situacao,PI_QuantidadePedida,PI_QuantidadeEntrada,PI_SaldoPedido, * from itempedido 
where PI_QuantidadePedida = PI_QuantidadeEntrada and PI_SaldoPedido = 0 and PI_Situacao = 'A'

