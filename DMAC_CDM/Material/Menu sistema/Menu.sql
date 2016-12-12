--------------------------------------------------------------------------------------------------------------
-------------------------------------------   MENU SISTEMA   -------------------------------------------------
--------------------------------------------------------------------------------------------------------------

select * from GLB_MenuSistema 
order by msi_COdigo

--sp_help GLB_MenuSistema

--update GLB_MenuSistema set msi_codigo = '000010' where msi_codigo = '000000' and msi_descricao = 'Agenda de Recebimentos'
--update GLB_MenuSistema set msi_controle = '0000' where msi_codigo = '020000'
--insert into GLB_MenuSistema(msi_codigo, msi_descricao, msi_nomeForm) values ('021300','Cria Romaneio','frmCriaRomaneio')
--delete GLB_MenuSistema where msi_codigo = '021300' and msi_descricao = 'Cria Romaneio' and msi_nomeForm = 'frmCriaRomaneio'
 
select * from glb_menusistema where msi_codigo like '02__00' and substring (msi_codigo, 3, 2) > '07'
select * from glb_menusistema where msi_codigo like '02__00' and substring (msi_codigo, 3, 2) > '06'
select * from glb_menusistema where msi_codigo like '02__00' and substring (msi_codigo, 3, 2) > '00'

Select substring(msi_codigo, 5,6) as codigo from glb_menusistema where msi_descricao =  'Manutenção de Romaneio'

select * from GLB_MenuSistema where msi_controle = '0200' order by msi_COdigo

/* INSERIR NOVO MENU

insert into GLB_MenuSistema values ('000000','Agenda de Recebimentos','frmAgendaDeRecebimento',null,null)

*/

--------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------

select * from GLB_AcessoSistema ORDER BY as_codigoTela
--insert into GLB_AcessoSistema values ('19','ReleminfEsp')
--delete GLB_MenuSistema where MSI_Codigo = 20800 and msi_descricao = 'Manutenção de Romaneio'                         

--update GLB_AcessoSistema set as_nomeTela = 'RelRoman' where as_codigoTela = 23

--delete GLB_AcessoSistema where as_codigoTela in (3,7,8,9,16) 


--------------------------------------------------------------------------------------------------------------
-------------------------------------   PERMISSAO USUARIO   --------------------------------------------------
--------------------------------------------------------------------------------------------------------------

select * from GLB_PermissaoSistema 
where ps_codigoUsuario = 11

--update GLB_PermissaoSistema set ps_nometela = 'RelEminf' where ps_nomeTela = 'frmBaixaTitulo'
--delete GLB_PermissaoSistema where ps_nomeTela = 'código'


--------------------------------------------------------------------------------------------------------------
--------------------------------------   USUARIO SISTEMA   ---------------------------------------------------
--------------------------------------------------------------------------------------------------------------

select * from GLB_UsuariosSistema

--update GLB_UsuariosSistema set us_senha = 'jeda36' where us_codigo > 10


--------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------

select count(*) from GLB_MenuSistema,GLB_PermissaoSistema,GLB_UsuariosSistema 
where ps_nomeTela = msi_codigo and ps_codigoUsuario = 11 and msi_nomeForm = 'frmAgendaDeRecebimento' or us_nivelAcesso = 'A' and us_codigo = 11


select count(*) administrador from GLB_UsuariosSistema where us_nivelAcesso = 'A' and us_codigo = '1'
--select * from glb_menusistema where msi_nomeForm = 'frmAgendaDeRecebimento'