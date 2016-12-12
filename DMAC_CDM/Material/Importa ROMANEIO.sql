update svdmac.dmac.dbo.Romaneio set ro_dataSolicitacao = '2015/10/08' where RO_NumeroRomaneio = 512517

INSERT INTO Romaneio(RO_NumeroRomaneio,RO_LojaOrigem,RO_LojaDestino,RO_Referencia,RO_DataSolicitacao,RO_DataSaida,RO_QuantidadePedida,RO_QuantidadeEnviada,RO_Tipo,RO_Situacao,RO_Motivo,RO_NotaFiscal,RO_Requisicao,RO_RomaneioImpressao,Ro_RomaneioNf,Ro_impressao,ro_dataprocesso,ro_conexao) 
select RO_NumeroRomaneio,RO_LojaOrigem,RO_LojaDestino,RO_Referencia,RO_DataSolicitacao,RO_DataSaida,RO_QuantidadePedida,RO_QuantidadeEnviada,RO_Tipo,RO_Situacao,RO_Motivo,RO_NotaFiscal,RO_Requisicao,RO_RomaneioImpressao,Ro_RomaneioNf,Ro_impressao,ro_dataprocesso,ro_conexao from svdmac.dmac.dbo.Romaneio where RO_Situacao = 'A'
AND ro_datasolicitacao = '2015/10/08' and RO_NumeroRomaneio = 512517


--update Romaneio set Ro_impressao = 0 where Ro_impressao is null and  RO_Situacao = 'A'
--UPDATE CONTROLESUP SET CS_NumeroRomaneio = (SELECT CS_NumeroRomaneio FROM DMAC353.DMAC_CDM.DBO.CONTROLECDM)



--select * from Romaneio where RO_NumeroRomaneio = 512511
--select * from svdmac.dmac.dbo.Romaneio where RO_NumeroRomaneio = 512511

--TRUNCATE TABLE Romaneio

--sp_help Romaneio