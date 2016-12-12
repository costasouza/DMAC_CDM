SELECT * FROM Romaneio WHERE ro_numeroromaneio in (521398,521408,521384,521374,521377,521372,521359,521413)


	select pr_precocusto1,NUMEROPED, * from nfitens, produto where dataemi = '2016/04/19' and VLUNIT = 0 and pr_referencia = referencia



	update nfitens set vlunit = pr_precocusto1, vlunit2 = pr_precocusto1 * qtde, vltotitem = pr_precocusto1 * qtde, plista = pr_precocusto1 from produto	 where VLUNIT = 0 and referencia = pr_referencia and 
	numeroped in (4370,4371,4372,4374,4375,4376,4370,4371,4372,4374,4375,4376,4370,4371,4372,4374,4375,4376,4370,4372,4374,4375,4376,4370,4371,4372,4374,4375,4376,4370,4372,4374,4370,4372,4374,4375,4376,4374,4375,4376,4370,4371,4372,4374,4375,4376,4375,4376,4370,4371,4372,4374,4375,4376,4370,4371,4372,4374,4375,4376)

	select nf, * from nfcapa where numeroped in (4370,4371,4372,4374,4375,4376,4370,4371,4372,4374,4375,4376,4370,4371,4372,4374,4375,4376,4370,4372,4374,4375,4376,4370,4371,4372,4374,4375,4376,4370,4372,4374,4370,4372,4374,4375,4376,4374,4375,4376,4370,4371,4372,4374,4375,4376,4375,4376,4370,4371,4372,4374,4375,4376,4370,4371,4372,4374,4375,4376)

	update nfcapa set vlrmercadoria = (select sum(vltotitem) from nfitens where NF = 9387) from nfcapa as capa where capa.NF = 9387

	update nfcapa set totalnota = vlrmercadoria, subtotal = vlrmercadoria  from nfcapa as capa where capa.NF in (9387)

	exec SP_VDA_Cria_NFe 'CD','9387','NE',''

	9381
9382
9383
9385
9386
9387

	DELETE NFE_IDE
