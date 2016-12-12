use DMAC_CDM

SELECT vc_situacao,* FROM CAPANFVENDA WHERE vC_LOJAORIGEM = 'cd' and vc_dataemissao = '2013/08/08'

select * from romaneio WHERE RO_romaneioImpressao = '2'

select * from romaneio where ro_lojaDestino in('','85','316','364') and ro_numeroromaneio in ('')


select top 1 cs_romaneioImpressao from controlecdm
select * from controlecdm

update controlecdm set cs_r

select * from capanfvenda where vc_serie = 'S3' and VC_LojaOrigem = 'CD' and vc_romaneioImpressao = 38
select * from itemnfvenda where vi_serie = 'S3' and Vi_LojaOrigem = 'CD' and vi_notaFiscal = 257
--delete itemnfvenda
select * from itemnfvenda
--update itemnfvenda set vi_lojaOrigem = 'CD' where vi_notaFiscal = 138

SELECT * FROM empresa

----------------------------------------------------------------------------------

Select * FROM controlecdm
Select * FROM paramentrosSistema
--update controlecdm set cs_serieNotaFiscal = 'S3'

-----------------------------------------------------------

update capanfvenda set vc_situacao = 'I' where vc_notafiscal = 163 and vc_serie = 'S3' and VC_LojaOrigem = 'CD' and vc_romaneioImpressao = 25
update capanfvenda set vc_situacao = 'I' where vc_notafiscal = 164 and vc_serie = 'S3' and VC_LojaOrigem = 'CD' and vc_romaneioImpressao = 25
update capanfvenda set vc_situacao = 'I' where vc_notafiscal = 165 and vc_serie = 'S3' and VC_LojaOrigem = 'CD' and vc_romaneioImpressao = 25
update capanfvenda set vc_situacao = 'I' where vc_notafiscal = 166 and vc_serie = 'S3' and VC_LojaOrigem = 'CD' and vc_romaneioImpressao = 25

update capanfvenda set vc_situacao = 'I' where vc_notafiscal in ( ) and vc_serie = 'S3' and VC_LojaOrigem = 'CD' and vc_romaneioImpressao = 26