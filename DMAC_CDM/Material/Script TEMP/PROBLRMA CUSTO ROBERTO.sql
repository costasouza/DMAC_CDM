select PR_PrecoCusto1,PR_PrecoCusto2,PR_PrecoCusto3, * from produto where pr_referencia = '0618081'


select CI_DescricaoFornecedor, CI_NovoCusto, CI_PrecoUnitario, CI_ValorIPI,CI_ValorICMSST, * from itemnfcompra where CI_Referencia = '0618081' and ci_serie = 'NE'



--select PRB_Referencia,count(PRB_CodigoBarras) from produtoBarras group by PRB_Referencia having count(PRB_CodigoBarras) > 2
--select * from produtoBarras where PRB_Referencia = '0139563'



--7650003


select * from produtoBarras WHERE PRB_Referencia = '0618081'
select * from dmac353.dmac_cdm.dbo.produtoBarras where PRB_CodigoBarras = '885911273855'
select * from produtoBarras where PRB_CodigoBarras = '885911273855'