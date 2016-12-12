Attribute VB_Name = "modInicial"
Option Explicit

Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Global subMenu(2) As String
Global msnon As String
Global Glb_AlteraResolucao As Boolean


Global NroPedido As Long
Global NroNota As Long
Global wPedido As String
Global GLB_Impressora00 As String


Type Natureza
    CFO As Long
    Descricao As String
    TipoNatureza As String
End Type

Public matNatureza() As Natureza

Public Type DefCab
    Cabecalho As String * 40
    Tamanho As Double
End Type

Public CabMat() As DefCab

Type CFOs
    CFO As Long
    CFOAux As Long
    Descricao As String
End Type

Public matCFO() As CFOs

Private comandoSQLanterior As String '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim rdoAgenda As rdoResultset

'---ADO Chekin---------------------------------------


Global adoDemeoChekin As New ADODB.Recordset
Global adoDemeoChekinChekin2  As New ADODB.Recordset
Global adoDemeoChekin3 As New ADODB.Recordset
Global adoDemeoChekin4 As New ADODB.Recordset
Global adoDemeoChekin5 As New ADODB.Recordset
Global adoDemeoChekin6 As New ADODB.Recordset



'---ADO Connection------------------------------------
Global ADO_Cn_CDA As New ADODB.Connection
Global ADO_Cn_CD As New ADODB.Connection
Global ADO_Cn_CDLocal As New ADODB.Connection


'---ADO RecordSet-------------------------------------
Global rs As New ADODB.Recordset
Global chamarTimer As String
Global ADO_Cn_rsCd As New ADODB.Recordset
Global rdorsNatOper As New ADODB.Recordset
Global adorsExtra1 As New ADODB.Recordset
Global adorsExtra2 As New ADODB.Recordset
Global adordoExtra2 As New ADODB.Recordset
Global adorsExtra3 As New ADODB.Recordset
Global adorsExtra6 As New ADODB.Recordset
Global adoMenuAnterior As New ADODB.Recordset
Global adoMenu As New ADODB.Recordset
Global adoAcessoSistema  As New ADODB.Recordset
Global RsdadosItens As New ADODB.Recordset
Global adoXML As New ADODB.Recordset
'---Variaveis----------------------------------------
Global telaChamou As String
Global sql As String
Global finalizaOutrasOperacoes As Boolean

Global LojaINI As String
Global Nomeservidor As String
Global BancoDeDados As String
Global NomeservidorLocal As String
Global BancoDeDadosLocal As String
Global wCont
Global NomeUsuario As String
Global SenhaUsuario As String

Global nomeBotao As String
Global codigoTela As String
Global auxAux As String
Global j As String

Global wAnterior As String
Global wControle As String
Global wContAux As String
Global wAuxMenuControle As String
Global wAuxMenuControle2 As String
Global telaControle As String

'Global wAux As String
'Global permissao As String
Global GLB_USU_Nome As String
Global GLB_modoOffline As Boolean
Global GLB_USU_Codigo As Integer
Global GLB_USU_NivelAcesso As String * 1
'Global nomeTelas As String

Global chamar As String
Global Grupo As String

Global WNovoCodigoOperacao As Integer
Global WCodigoOperacaoVelho As Integer

Global wLojaMCE85ouCD As String * 5
Global LojaOrigem  As String
Global nomeImpressora As Printer
Global wSerieImpressao As String
Global wControlaQuebraDaPagina As Integer
Public NomeEmpresa As String



Public DesabilitaStatus As Boolean

Public ClausulaWhere As String

Dim wImpressora As String

Global wCFO1 As String * 6
Global wCFO2 As String * 6
Global wCFO3 As String * 1

Global wTelaOperacaoEspecial As Boolean

Global wSerie As String
Global wRazao As String
Global wendereco As String
Global wbairro As String
Global WCGC As String
Global WIest As String
Global wMunicipio As String
Global westadop As String
Global WCep As String
Global wFone As String
Global wDDDLoja As String
Global WFax As String
Global wLoja As String
Global wSenhaLiberacao As String
Global wNovaRazao As String
Global westado As String
Global GLB_Loja As String
Global glb_LojaCNPJ As String
    
Global chamou As String
    
Global pedido As Integer
    
Global NroNotaFiscal As Long
Global wNroItens As Integer
Global wTipoNota As String
Global cliente As String
Global wTotNota As Double
Global wVlrMercadoria As Double
Global wCFOP As String
Global wCodigoProduto As String
Global wQtde As Integer
Global wItemVenda As Integer
Global wVlTotItem As Double
Global wICMS As Double
Global wPLISTA As Double

Global wQuantdadeTotalItem As Double
Global wQuantItensCapaNF As Double
Global wQuantItensNF As Double
Global wChaveICMS As Long
Global wIE_icmsAplicado As Double
Global wRecebeCarimboAnexo As String
Global wQuant As Long
Global GLB_ECF As Integer

Global wAnexo As String
Global wAnexo1 As String
Global wAnexo2 As String
Global wTotalPed As Double
Global wDesconto As Double
Global wCodigo As String
Global GLB_TotalICMSCalculado As Double
Global GLB_ValorCalculadoICMS As Double
Global GLB_BasedeCalculoICMS As Double
Global wReemissaoNotaFiscal As Boolean
Global GLB_AliquotaAplicadaICMS As Double
Global GLB_AliquotaICMS As Double
Global WNomeCliente As String
Global GLB_BaseTotalICMS As Double
Global GLB_Tributacao As String * 3
Global wCFOItem As Double

Global wUltimoItem As Long
Global wComissaoVenda As Double
Global wSomaVenda As Double
Global wSomaMargem As Double

Global wCarimbo5 As String * 132
Global wCarimbo2 As String * 132
Global wPessoa As Double

Global GLB_CFOP As String
Global wTM As String
Global wST20 As String
Global wST60 As String
Global notaRe As String

Global wStrI As String

Global i As Integer
Global rsCarimbo2 As New ADODB.Recordset
Global rsPegaData As New ADODB.Recordset
Global rsPegaLoja As New ADODB.Recordset

Global SQL2 As String

Global wReemissao As Boolean
Global Wsm As Boolean
Global wPegaDescricaoAlternativa As String
Global wEndCliente As String
Global wCgcCliente As String
Global WVendedor As String
Global wPegaSequenciaCO As Double
Global WNfTransferencia As String
Global WNF As String

Global wNotaDoDia As Boolean
Global wImpressoraNota As String
Global wDetalheImpressao As String
Global wIE_BasedeReducao  As Double

Global wRomaneio As Boolean

Global Wnatureza As String
Global wPagina As Integer
Global NroItens As Long
Global wAnexoIten As Integer

Global wSubstituicaoTributaria As Double

Global Usuario As String

Global wChaveICMSitem As Double
Global WAnexoAux As String * 20
Global wIE_Cfo  As Integer

Global Wcondicao As String

Global wNomeservidor As String
Global wNomeBanco As String
Global wUsuario As String
Global wSenha As String

Global GLB_ImpCotacao As String
Global wNumeroCupom As String * 6
Global GLB_ConectouOK As Boolean
Global conectADO As Boolean
Global lsDSN As String
Global ContaImpressora As Integer
Global wDescricao As String
Global wIE_Tributacao As Double
Global wIE_icmsdestino As Double
Global wNotaTransferencia As Boolean

Global adoProd As New ADODB.Recordset

Global quantBotaoControleCD As Integer
Global botaoInicioControleCD As Integer


    Global wAliqICMSInterEstadual As String

   Global Wav As String
        
Global wConta As Long

       Global Glb_NfDevolucao As Boolean


       Global wReferenciaEspecial As String
       Global wCarimbo4 As String * 132
       Global wStr0, wStr1, wStr2, wStr3, wStr4, wStr5, wStr6, wStr7 As String
Global wStr8, wStr9, wStr10, wStr11, wStr12, wStr13, wStr15, wStr16, wStr17, wStr18, wStr19, wStr20, wStr21 As String

Global NroBanner As Integer
Global GLB_logoPedido As String
    
Public adotipo As New ADODB.Recordset
Public adorsLojas As New ADODB.Recordset
Public adorsCtsup As New ADODB.Recordset
Public adoControle As New ADODB.Recordset
Public adoContaItens As New ADODB.Recordset
Public adoSerie As New ADODB.Recordset
Public adoItemNota As New ADODB.Recordset
Public adoCapaNF As New ADODB.Recordset
Global adoItensNf As New ADODB.Recordset
Global RsICMSIntER As New ADODB.Recordset
Global rsDados As New ADODB.Recordset
Global adoConPag As New ADODB.Recordset
Global RsPegaItensEspeciais As New ADODB.Recordset
    Global wLojaVenda As String
    Global wVendedorLojaVenda As String
    Global Wentrada As String
    
'-----Constantes-------------------------------
Public Const Almoxarifado = "CD"
Global Const enderecoBancoINI = "c:\sistemas\DMAC_CDM.mdb"


Public pastaRecebido As String
Public pastaLido As String
Public pastaInvalido As String
Public buscaAutomaticaXML As String
Public servidor, banco As String
Public Loja, Tempo As String

Sub Main()
    Dim lsDSN As String
   
    Call verificaAppExecucao
   
On Error GoTo TrataErro
   
     lsDSN = "Driver={Microsoft Access Driver (*.mdb)};" & _
             "Dbq=" & enderecoBancoINI & ";" & _
             "Uid=Admin; Pwd=astap36"
    
    ADO_Cn_CDA.Open lsDSN
    
    sql = "Select count(*) as cdm_codigo from conexaoCDM"
    
    ADO_Cn_rsCd.CursorLocation = adUseClient
    ADO_Cn_rsCd.Open sql, ADO_Cn_CDA, adOpenForwardOnly, adLockPessimistic
    
         If Not ADO_Cn_rsCd.EOF Then
               ADO_Cn_rsCd.Close
               
                      sql = "Select top 1 * from conexaoCDM order by cdm_codigo"
                      ADO_Cn_rsCd.CursorLocation = adUseClient
                      ADO_Cn_rsCd.Open sql, ADO_Cn_CDA, adOpenForwardOnly, adLockPessimistic
                      
                      If Not ADO_Cn_rsCd.EOF Then
                      
                         LojaINI = Trim(ADO_Cn_rsCd("CDM_Loja"))
                         Nomeservidor = Trim(ADO_Cn_rsCd("CDM_Servidor"))
                         BancoDeDados = Trim(ADO_Cn_rsCd("CDM_Banco"))
                        
                         NomeservidorLocal = Trim(ADO_Cn_rsCd("CDM_ServidorLocal"))
                         BancoDeDadosLocal = Trim(ADO_Cn_rsCd("CDM_BancoLocal"))
                        
                         'NomeUsuario = Trim(ADO_Cn_rsCd("CDM_Usuario"))
                         'SenhaUsuario = Trim(ADO_Cn_rsCd("CDM_Senha"))
                         
                         ADO_Cn_rsCd.Close
                         
                        sql = "Select top 1 * from parametroSistema"
                        ADO_Cn_rsCd.CursorLocation = adUseClient
                        ADO_Cn_rsCd.Open sql, ADO_Cn_CDA, adOpenForwardOnly, adLockPessimistic
                        
                        wImpressoraNota = Trim(ADO_Cn_rsCd("GLB_ImpNFE"))
                        
                        If Not ADO_Cn_rsCd.EOF Then
                            Glb_AlteraResolucao = Trim(ADO_Cn_rsCd("GLB_AlteraResolucao"))
                        Else
                            Glb_AlteraResolucao = False
                        End If
                        
                         ADO_Cn_rsCd.Close
                         
                      Else
                         MsgBox "Problemas no banco de dados de inicializacao", vbCritical, "Aviso"
                         End
                         Exit Sub
                      End If
    
         Else
            MsgBox "Banco de dados de inicializacao Vazio", vbCritical, "Aviso"
            End
            Exit Sub
         End If
    
    
    Call ConectaADOLocal
    Call DadosLoja
    

    CarregarDBIniXML
    
    resolucaoOriginal.Colunas = resolucaoTela.Colunas
    resolucaoOriginal.Linhas = resolucaoTela.Linhas
    Call AlterarResolucao(1024, 768)
    
    If buscaAutomaticaXML Then
        frmBandeja.Show
    Else
        frmLogin.Show
    End If
    
    'CarregaNaturezaOperacao
     
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then
        Select Case Err.Number
            Case -2147467259
                MsgBox "Não foi possível localizar ou conectar ao ini:" & vbNewLine & _
                enderecoBancoINI & ", vbCritical, Err.Source"
            Case -2147217865
                MsgBox "Não foi possível ler a tabela do ini:" & vbNewLine & _
                enderecoBancoINI & ", vbCritical, Err.Source"
            Case Else
                MsgBox "Ocorreu um erro desconhecido no modulo de conexão", _
                vbCritical, Err.Source
        End Select
    End If
End Sub

Function CarregarDBIniXML()
    
  On Error GoTo erroConexaoBancoINI
    
  'Dim ado_cn_dmac  As New ADODB.Connection
  Dim ADO_Cn_DmacA  As New ADODB.Connection
  Dim ADO_Cn_rsDmac As New ADODB.Recordset
    
  ADO_Cn_DmacA.Open "Driver={Microsoft Access Driver (*.mdb)};" & _
                    "Dbq=c:\sistemas\DMAC_CDM.mdb;" & _
                    "Uid=Admin; Pwd=astap36"
 
  sql = "Select * from ConexaoLeituraXML"

  ADO_Cn_rsDmac.CursorLocation = adUseClient
  ADO_Cn_rsDmac.Open sql, ADO_Cn_DmacA, adOpenForwardOnly, adLockPessimistic
 
  With ADO_Cn_rsDmac
  
        If .BOF And .EOF Then
            MsgBox "Problemas no banco de dados de inicialização", vbCritical, "Erro"
            End
        Else
                        servidor = .Fields("CLX_Servidor")
                        banco = .Fields("CLX_Banco")
                        'Usuario = .Fields("CLX_Usuario")
                        'Senha = .Fields("CLX_Senha")
                        Tempo = .Fields("CLX_TempoBusca")
                        Loja = .Fields("CLX_Loja")
                        pastaInvalido = .Fields("CLX_Invalidos")
                        pastaLido = .Fields("CLX_lidos")
                        pastaRecebido = .Fields("CLX_Recebidos")
                        buscaAutomaticaXML = .Fields("CLX_BuscaAutomaticaXML")
                        'tempo = 10000
        End If
        
    End With
     
    ADO_Cn_DmacA.Close
    Exit Function
    
erroConexaoBancoINI:
    Select Case Err.Number
        Case -2147467259
            MsgBox "Não foi possível localizar ou conectar ao ini:" & vbNewLine _
            & "c:\sistemas\DMAC_CDM.mdb", vbCritical, "Erro"
        Case Else
            MsgBox "Não foi possível localizar ou conectar ao Banco de Dados" & vbCritical, "Erro"
        End
    End Select
     
End Function

Function ConectaADO() As Boolean   'conexao retaguarda


    If ConexaoDLLaDO.abrirConexaoADO(ADO_Cn_CD, Nomeservidor, BancoDeDados) Then
        ConectaADO = True
        Exit Function
    Else
        MsgBox "Erro ao se conectar no banco de dados da Retaguarda", vbCritical, "DMAC CDM " & _
        App.Major & "." & App.Minor & "." & App.Revision
    End If
    
End Function


Function ConectaADOLocal() As Boolean

    'On Error GoTo TrataErro
    
    If ConexaoDLLaDO.abrirConexaoADO(ADO_Cn_CDLocal, NomeservidorLocal, BancoDeDadosLocal) Then
        ConectaADOLocal = True
        Exit Function
    Else
        MsgBox "Erro ao se conectar no banco de dados Local", vbCritical, "DMAC CDM " & _
        App.Major & "." & App.Minor & "." & App.Revision
        End
    End If

'TrataErro:
    
 '       erroBancoDeDados Err
    
End Function


Function CadastraTela(nomeTelas)

    sql = "Exec SP_GLB_Ler_Acesso_Sistema_Por_Parametro '" & nomeTelas & "'"
    
        adoAcessoSistema.CursorLocation = adUseClient
        adoAcessoSistema.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
        If adoAcessoSistema.EOF Then
             ADO_Cn_CDLocal.BeginTrans
                sql = "Exec SP_GLB_Grava_Acesso_Sistema '" & nomeTelas & "'"
      
            'ADO_Cn_CDLocal.Execute (SQL)
            'ADO_Cn_CDLocal.CommitTrans
           
        End If
        
    adoAcessoSistema.Close
    
End Function

Function verificaPermissao(ByRef nomeTela As String) As Boolean
    Screen.MousePointer = 11
    Dim adopermissaoSistema  As New ADODB.Recordset

    sql = "select count(*) administrador from GLB_UsuariosSistema " & _
    "where us_nivelAcesso = 'A' and us_codigo = '" & GLB_USU_Codigo & "'"
    adopermissaoSistema.CursorLocation = adUseClient
    adopermissaoSistema.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If adopermissaoSistema("administrador") = 1 Then
        verificaPermissao = True
    Else
        adopermissaoSistema.Close
        sql = "select count(*) autorizado from GLB_MenuSistema,GLB_PermissaoSistema " & _
        "where ps_nomeTela = msi_codigo and ps_codigoUsuario = " & GLB_USU_Codigo & _
        " and msi_nomeForm = '" & nomeTela & "'"
        adopermissaoSistema.CursorLocation = adUseClient
        adopermissaoSistema.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
        If adopermissaoSistema("autorizado") = 1 Then
            verificaPermissao = True
        Else
            verificaPermissao = False
        End If
    End If

    adopermissaoSistema.Close
    
    Screen.MousePointer = 0
End Function


Function menuProximo(menuControle As String, proximoMenus As Boolean)
    Dim auxiliar, comandoSQLMenu As String
    Dim i As Byte
    
    If Not proximoMenus Then
        wAnterior = "00"
        comandoSQLanterior = menuControle
    Else
        menuControle = comandoSQLanterior
    End If
    
    'quantBotaoControleCD = 0
    
    subMenu(0) = Mid(menuControle, 1, 2)
    subMenu(1) = Mid(menuControle, 3, 2)
    subMenu(2) = Mid(menuControle, 5, 2)

        'comandoSQLMenu = "select distinct top " & quantBotaoControleCD & " msi_codigo, msi_descricao, msi_nomeForm " & _
        '"from glb_menusistema,GLB_PermissaoSistema where "

        comandoSQLMenu = "select distinct top " & 12 & " msi_codigo, msi_descricao, msi_nomeForm " & _
        "from glb_menusistema,GLB_PermissaoSistema where "
        
        comandoSQLMenu = comandoSQLMenu & "ps_codigoUsuario = '" & GLB_USU_Codigo & "'"
        
        If Mid(GLB_USU_NivelAcesso, 1, 1) <> "A" Then
            comandoSQLMenu = comandoSQLMenu & " and msi_codigo = ps_nometela"
        End If

        comandoSQLMenu = comandoSQLMenu & " and msi_codigo like '"

        If subMenu(0) = "00" Then
            comandoSQLMenu = comandoSQLMenu & "__0000' and substring (msi_codigo, 1, 2) > '" & wAnterior & "'"
        ElseIf subMenu(1) = "00" Then
            comandoSQLMenu = comandoSQLMenu & subMenu(0) & "__00' and substring (msi_codigo, 3, 2) > '" & wAnterior & "'"
        ElseIf subMenu(2) = "00" Then
            comandoSQLMenu = comandoSQLMenu & subMenu(0) & subMenu(1) & "__' and substring (msi_codigo, 5, 2) > '" & wAnterior & "'"
        Else
            Exit Function
        End If
    
        If adoMenu.State = 1 Then
            adoMenu.Close
        End If
        
        adoMenu.CursorLocation = adUseClient
        adoMenu.Open comandoSQLMenu, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
       
        If adoMenu.EOF Then
            adoMenu.Close
            Exit Function
        End If
        
        If adoMenu.RecordCount > 12 Then
            quantBotaoControleCD = 5
            botaoInicioControleCD = 2
            frmControleCD.cmdAvanca.Visible = True
            frmControleCD.cmdVolta.Visible = True
        Else
            quantBotaoControleCD = 11
            botaoInicioControleCD = 1
            frmControleCD.cmdAvanca.Visible = False
            frmControleCD.cmdVolta.Visible = False
        End If
        
    menuBotoes
    
End Function


Function menuAnterior(menuControle As String)
'        Dim comandoSQLMenu As String
'
'        If auxiliarMenu > "05" Then
'            auxiliarMenu = Format((Mid(comandoSQLanterior, 98, 2) - 7), "0#")
'            comandoSQLMenu = Mid(comandoSQLanterior, 1, 97) & auxiliarMenu & "'"
'
'            If adoMenu.State = 1 Then
'                adoMenu.Close
'            End If
'            adoMenu.CursorLocation = adUseClient
'            adoMenu.Open comandoSQLMenu, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
'
'            If adoMenu.EOF Then
'                adoMenu.Close
'                Exit Function
'            End If
'        End If
'
'        menuBotoes
'        comandoSQLanterior = comandoSQLMenu
    MsgBox "Erro"
End Function

Private Sub menuBotoes()

    'Revisado por Felip3FL
    'Não Completo
    'Dia 08/06/2013
            
    Dim i As Byte
            
    If Not adoMenu.EOF Then
        
        For i = botaoInicioControleCD To quantBotaoControleCD
            frmControleCD.cmdBotao(i).Visible = False
        Next
        wCont = botaoInicioControleCD
        Do While Not adoMenu.EOF
            For i = botaoInicioControleCD To quantBotaoControleCD
                If wCont = i Then
                    frmControleCD.cmdBotao(i).ToolTipText = Trim(adoMenu("msi_descricao"))
                    frmControleCD.cmdBotao(i).Tag = Trim(adoMenu("msi_descricao"))
                    frmControleCD.cmdBotao(i).Visible = True
                    frmControleCD.cmdBotao(i).Picture = LoadPicture(endIMGBotao(Trim(adoMenu("msi_nomeForm"))))
                    'frmControleCD.cmdBotao(i).Height = 200
                    'LoadPicture (endIMGBotao(Trim(adoMenu("msi_codigo"))))
                    'wCont = wCont + 1
                    
                    
                        If Mid(adoMenu("msi_codigo"), 5, 2) <> "00" Then
                            wAnterior = Mid(adoMenu("msi_codigo"), 5, 2)
                        ElseIf Mid(adoMenu("msi_codigo"), 3, 2) <> "00" Then
                            wAnterior = Mid(adoMenu("msi_codigo"), 3, 2)
                        ElseIf Mid(adoMenu("msi_codigo"), 1, 2) <> "00" Then
                            wAnterior = Mid(adoMenu("msi_codigo"), 1, 2)
                        End If
                    
                End If
            Next
            wCont = wCont + 1
            
            adoMenu.MoveNext
        Loop
    End If
    ' chamarTimer = "ok"
    
End Sub

Function ChamaTelaMenu(nomeBotao As String)

    Screen.MousePointer = 11
    
    Dim adoMenu As New ADODB.Recordset
    Dim sql As String
    
    On Error GoTo TrataErro
    
    sql = "select msi_codigo, msi_descricao, msi_nomeForm from GLB_MenuSistema where msi_descricao = '" & nomeBotao & "'"
    adoMenu.CursorLocation = adUseClient
    adoMenu.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Trim(adoMenu("msi_nomeForm")) Like "MENU*" Then
         wCont = 1
         menuProximo adoMenu("msi_codigo"), False
    Else
            Screen.MousePointer = 0
            frmControleCD.lblNomeTelas.Caption = nomeBotao
            frmControleCD.lblNomeTelas.Visible = True
            Forms.Add(Trim(adoMenu("msi_nomeForm"))).Show 1
            frmControleCD.lblNomeTelas.Visible = False
            Err.Number = 0
    End If
    'frmControleCD.Image1.Picture = LoadPicture("C:\Sistemas\cd\Imagens\t3.JPG")
        
    'CadastraTela nomeTelas
    
TrataErro:
    If Err.Number <> 0 Then
        erroBancoDeDados Err
        Select Case Err.Number
            Case 424 'Formulario não encontrado
            MsgBox "Não foi possível encontrar " & Chr(34) & Trim(adoMenu("msi_descricao")) & Chr(34) & " no sistema" & vbNewLine & _
            "Verifique se o menu foi cadastrado corretamente" _
            , vbCritical, "Erro no formulário"
            Case Else 'Formulario não encontrado
            MsgBox "Erro ao monta ou abrir o Menu" & vbNewLine & Err.Number & " - " & Err.Description, vbCritical, "Erro no formulário"
        End Select
    End If
    
    Screen.MousePointer = 0
End Function


Sub PreencheCombo(ByRef NomeCombo As ComboBox, ByRef SeuRecordset, ByVal CampoCodigo As String, ByVal CampoTexto As String)

    NomeCombo.Clear
    Do While Not SeuRecordset.EOF
        If CampoTexto <> "" Then
            If Not IsNull(SeuRecordset(CampoTexto)) Then
                NomeCombo.AddItem SeuRecordset(CampoCodigo) & " - " & SeuRecordset(CampoTexto)
            Else
                NomeCombo.AddItem SeuRecordset(CampoCodigo) & " - "
            End If
        Else
            NomeCombo.AddItem SeuRecordset(CampoCodigo)
        End If
        NomeCombo.ItemData(NomeCombo.NewIndex) = Val(SeuRecordset(CampoCodigo))
        SeuRecordset.MoveNext
    Loop

End Sub

Function GetIndice(ByVal TextoProcurado As String, ByRef SeuCombo As ComboBox) As Long

    Dim Indice As Long
    Dim Tamanho As Long
    
    Tamanho = Len(TextoProcurado)
    
    If Tamanho > 0 Then
        For Indice = 0 To SeuCombo.ListCount - 1 Step 1
            If UCase(TextoProcurado) = UCase(left(SeuCombo.List(Indice), Tamanho)) Then
                Exit For
            End If
        Next Indice
        
        If Indice >= SeuCombo.ListCount Then
            Indice = 0
        End If
    Else
        Indice = -1
    End If
    
    GetIndice = Indice

End Function

Sub MontaComboNatureza(ByVal CFO As Long, ByRef cmbNaturezaOperacao As ComboBox)
    
    Dim Ponteiro As Long
    Dim Maximo As Long
    
    cmbNaturezaOperacao.Clear
    Ponteiro = 0
    Maximo = UBound(matNatureza)
    'CFO = WCodigoOperacaoVelho
    
    If CFO > 0 Then
        Do While CFO <> matNatureza(Ponteiro).CFO
            Ponteiro = Ponteiro + 1
            If Ponteiro > Maximo Then
                Exit Sub
            End If
        Loop
        
        Do While CFO = matNatureza(Ponteiro).CFO
            cmbNaturezaOperacao.AddItem matNatureza(Ponteiro).Descricao
            Ponteiro = Ponteiro + 1
            If Ponteiro > Maximo Then
                Exit Sub
            End If
        Loop
    End If
    
End Sub


Function Trimar(ByVal Texto As String) As String

    Dim Char As Integer
    Dim Maximo As Integer
    Dim Charlido As String * 1
    Dim Retorno As String
    
    Maximo = Len(Texto)
    
    Retorno = ""
    For Char = 1 To Maximo Step 1
        Charlido = UCase(Mid(Texto, Char, 1))
        If IsNumeric(Charlido) Or Charlido Like "[A-Z]" Then
            Retorno = Retorno & Charlido
        End If
    Next Char
    
    Trimar = Retorno

End Function

Function ConverteVirgula(ByVal Expressao) As String
    Dim ContPad As String
    Dim flgpad As Integer
    
    If Len(Expressao) <> 0 Then
        ContPad = CStr(Expressao)
        flgpad = InStr(ContPad, ",")
        Do While flgpad <> 0
            Mid(ContPad, flgpad, 1) = "."
            flgpad = InStr(ContPad, ",")
        Loop
    Else
        ContPad = 0
    End If
    ConverteVirgula = ContPad
End Function

Function Numeros(ByVal Texto As String) As String

    Dim Maximo As Integer
    Dim Char As Integer
    Dim Charlido As String * 1
    Dim Retorno As String
    
    Maximo = Len(Texto)
    
    Retorno = ""
    For Char = 1 To Maximo Step 1
        Charlido = Mid(Texto, Char, 1)
        If IsNumeric(Charlido) Then
            Retorno = Retorno & Charlido
        End If
    Next Char
    
    Texto = Retorno
    
    Numeros = Texto

End Function

Sub VerTecla(ByRef Tecla As Integer)
    
    If Not IsNumeric(Chr(Tecla)) And Tecla <> vbKeyBack Then
        Tecla = 0
    End If

End Sub

Function DefineGrade(ByRef GradeUsu, Optional Linha)

    If IsMissing(Linha) Then
        Linha = 0
    End If

    Dim wCntDefGrid As Integer
    With GradeUsu
        .Rows = 2
        .Cols = UBound(CabMat)
        .row = Linha
        For wCntDefGrid = 0 To (.Cols - 1)
            .col = wCntDefGrid
            .Text = CabMat(wCntDefGrid).Cabecalho
            .CellAlignment = 9
            .ColWidth(wCntDefGrid) = CabMat(wCntDefGrid).Tamanho
        Next wCntDefGrid
    End With
End Function


Sub MostraErro()

    Dim Erro As rdoError

    For Each Erro In rdoErrors
        MsgBox Erro.Number & ": " & Erro.Description, vbCritical, "Descrição de Erro"
    Next Erro

End Sub

Sub PreencheComboUF(ByRef cmbEstado As ComboBox)
   cmbEstado.AddItem "AC"
   cmbEstado.AddItem "AL"
   cmbEstado.AddItem "AM"
   cmbEstado.AddItem "AP"
   cmbEstado.AddItem "BA"
   cmbEstado.AddItem "CE"
   cmbEstado.AddItem "DF"
   cmbEstado.AddItem "ES"
   cmbEstado.AddItem "GO"
   cmbEstado.AddItem "MG"
   cmbEstado.AddItem "MS"
   cmbEstado.AddItem "MT"
   cmbEstado.AddItem "PA"
   cmbEstado.AddItem "PB"
   cmbEstado.AddItem "PE"
   cmbEstado.AddItem "PI"
   cmbEstado.AddItem "PR"
   cmbEstado.AddItem "SC"
   cmbEstado.AddItem "SE"
   cmbEstado.AddItem "SP"
   cmbEstado.AddItem "RJ"
   cmbEstado.AddItem "RN"
   cmbEstado.AddItem "RO"
   cmbEstado.AddItem "RR"
   cmbEstado.AddItem "RS"
   cmbEstado.AddItem "TO"
End Sub

Sub CarregaNaturezaOperacao()
    
    Dim adoCombos As New ADODB.Recordset
    Dim Ponteiro As Long
    
    
    sql = "Select NO_CodigoOperacao, NO_CodigoNatureza, NO_Descricao, NO_TipoNatureza from NaturezaOperacao order by NO_CodigoOperacao"
    
    adoCombos.CursorLocation = adUseClient
    adoCombos.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    
    ReDim matNatureza(0) As Natureza
    
    Ponteiro = 0
    
    Do While Not adoCombos.EOF
        ReDim Preserve matNatureza(Ponteiro) As Natureza
        
        matNatureza(Ponteiro).CFO = adoCombos("NO_CodigoOperacao")
        matNatureza(Ponteiro).Descricao = adoCombos("NO_CodigoNatureza") & " - " & adoCombos("NO_Descricao")
        matNatureza(Ponteiro).TipoNatureza = adoCombos("NO_TipoNatureza")
        
        Ponteiro = Ponteiro + 1
        
        adoCombos.MoveNext
    Loop
    adoCombos.Close
    
    sql = "Select CF_CodigoOperacao, CF_CodigoOperacaoAux, CF_Descricao from CodigoOperacao"
    
   adoCombos.CursorLocation = adUseClient
    adoCombos.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    ReDim matCFO(0) As CFOs
    
    Ponteiro = 0
    
    Do While Not adoCombos.EOF
        ReDim Preserve matCFO(Ponteiro) As CFOs
        
        matCFO(Ponteiro).CFO = adoCombos("CF_CodigoOperacao")
        matCFO(Ponteiro).CFOAux = adoCombos("CF_CodigoOperacaoAux")
        matCFO(Ponteiro).Descricao = adoCombos("CF_Descricao")
        
        Ponteiro = Ponteiro + 1
        
        adoCombos.MoveNext
    Loop
    
    adoCombos.Close
    
End Sub

Private Function LocalizaTipoNatureza(ByVal codigo As Long) As String

    Dim MaximoNatureza As Long
    Dim Indice As Long
    Dim TipoNatureza As String
    
    MaximoNatureza = UBound(matNatureza)
    
    For Indice = 0 To MaximoNatureza Step 1
        If matNatureza(Indice).CFO = codigo Then
            Exit For
        End If
    Next Indice
    
    If Indice <= MaximoNatureza Then
        LocalizaTipoNatureza = matNatureza(Indice).TipoNatureza
    Else
        LocalizaTipoNatureza = ""
    End If

End Function

Function DefineImpressora(ByVal NotaFiscal As Long) As Boolean

sql = "Select vc_notafiscal,vc_baseicms,vc_enderecocliente," _
        & "vc_serie,vc_lojaorigem,vc_lojadestino,vc_dataemissao,Vc_LojaVenda," _
        & "vc_codigooperacao,vc_totalnota,vc_valormercadorias,vc_codigooperacaoNovo," _
        & "vc_aliquotaicms,vc_valoricms,vc_situacao,vi_notafiscal," _
        & "vi_serie,vi_referencia,vi_quantidade,vi_precounitario," _
        & "vi_reserva,vi_valormercadoria,vi_valoripi,vi_aliquotaicms, " _
        & "VC_Observacao,PR_substituicaotributaria,vc_nomecliente, vc_cgccliente, vc_dataemissao, vc_enderecocliente, " _
        & " vc_bairrocliente, vc_cepcliente, vc_municipiocliente, vc_ufcliente, vc_InscEstCliente, VI_CustoMedioLiquidoUnitario" _
        & " From capanfvenda,itemnfvenda, produto " _
        & "where " _
        & " vc_notafiscal = vi_notafiscal and " _
        & " vc_serie = vi_serie and " _
        & " vc_lojaorigem = vi_lojaorigem and" _
        & " vi_referencia=pr_referencia and" _
        & " vc_notafiscal= " & NotaFiscal & " And " _
        & " vc_serie =  '" & wSerieImpressao & "' and " _
        & " vc_lojaorigem = '" & LojaOrigem & "'" _
        & "order by PR_CodigoFornecedor, PR_Descricao"
        
       adorsExtra1.CursorLocation = adUseClient
        adorsExtra1.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
       
        If Not adorsExtra1.EOF Then
         If adorsExtra1("vc_serie") = "CT" Then
            wImpressora = "ROMANEIO"
         Else
            wImpressora = "NOTA FISCAL"
         End If
   
         For Each nomeImpressora In Printers
             If UCase(nomeImpressora.DeviceName) = UCase(wImpressora) Then
                Set Printer = nomeImpressora
                Exit For
             End If
         Next
       
        If adorsExtra1("vc_serie") = "CT" Then
           Printer.ScaleMode = vbMillimeters
           Printer.ForeColor = "0"
           Printer.FontSize = 8
           Printer.FontName = "draft 20cpi"
           Printer.FontSize = 8
           Printer.FontBold = False
           Printer.DrawWidth = 3
           Printer.FontName = "COURIER NEW"
           Printer.FontSize = 7.3
        Else
            Printer.ScaleMode = vbMillimeters
            Printer.ForeColor = "0"
            Printer.FontSize = 8
            Printer.FontName = "draft 20cpi"
            Printer.FontSize = 8
            Printer.FontBold = False
            Printer.DrawWidth = 3
            Printer.FontName = "COURIER NEW"
            Printer.FontSize = 8#
        End If
    End If
    adorsExtra1.Close

End Function

Sub status(ByVal wMens)

    If Not DesabilitaStatus Then
        'Forms(0).stbMDI.Panels.item(1).Text = Trim(wMens)
        'Forms(0).stbMDI.Refresh
    End If
    
End Sub


Function ImprimirNota(ByVal NotaFiscal As Long) As Boolean
   
Dim wConta As Long
Dim wChave As Long
Dim flg As Long
Dim i As Long
Dim wReduz As Long
Dim wAliquotaZero As Boolean
Dim wRes As Long
Dim wpag As Long
Dim wlin As Long
Dim tmporient As Long
Dim Ind As Long
Dim Impressora As Printer
Dim serie As String
Dim wnome As String
Dim wendereco As String
Dim wbairro As String
Dim westado As String
Dim wvar As String
Dim wRodape As String
Dim wCodIPI As String
Dim wCodTri As String
Dim wStr1 As String
Dim wStr2 As String
Dim wStr3 As String
Dim wStr4 As String
Dim wStr5 As String
Dim wStr6 As String
Dim wStr7 As String
Dim wStr8 As String
Dim wStr9 As String
Dim wStr10 As String
Dim wStr11 As String
Dim wStr12 As String
Dim wStr13 As String
Dim wStr14 As String
Dim wStr15 As String
Dim wStr16 As String
Dim wStr17 As String
Dim wStr18 As String
Dim wStr19 As String
Dim wStr20 As String
Dim wEspaco As String
Dim wLO_inscricaoestadual As String
Dim wDescricao As String
Dim Wnatureza As String
Dim ParaImpr As Boolean
Dim WcodigoOperacao As String
Dim rdoTransportadora As New ADODB.Recordset


Dim wWhere As String

wTelaOperacaoEspecial = True

 sql = "Select * from transportadora"
 
     rdoTransportadora.CursorLocation = adUseClient
     rdoTransportadora.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
  
        
    If Not rdoTransportadora.EOF Then
       wnome = Trim(rdoTransportadora("tr_nome"))
       wendereco = Trim(rdoTransportadora("tr_endereco"))
       wbairro = Trim(rdoTransportadora("tr_bairro"))
       westado = Trim(rdoTransportadora("tr_estado"))
    End If
    
    rdoTransportadora.Close

   
  
   
   sql = "Select vc_notafiscal,vc_baseicms,vc_enderecocliente," _
        & "vc_serie,vc_lojaorigem,vc_lojadestino,vc_dataemissao,Vc_LojaVenda," _
        & "vc_codigooperacao,vc_totalnota,vc_valormercadorias,vc_codigooperacaoNovo," _
        & "vc_aliquotaicms,vc_valoricms,vc_situacao,vi_notafiscal," _
        & "vi_serie,vi_referencia,vi_quantidade,vi_precounitario," _
        & "vi_reserva,vi_valormercadoria,vi_valoripi,vi_aliquotaicms, " _
        & "VC_Observacao,PR_substituicaotributaria,vc_nomecliente, vc_cgccliente, vc_dataemissao, vc_enderecocliente, " _
        & " vc_bairrocliente, vc_cepcliente, vc_municipiocliente, vc_ufcliente, vc_InscEstCliente, VI_CustoMedioLiquidoUnitario" _
        & " From capanfvenda,itemnfvenda, produto " _
        & "where " _
        & " vc_notafiscal = vi_notafiscal and " _
        & " vc_serie = vi_serie and " _
        & " vc_lojaorigem = vi_lojaorigem and" _
        & " vi_referencia=pr_referencia and" _
        & " vc_notafiscal= " & NotaFiscal & " And " _
        & " vc_serie = 'S2'  and " _
        & " vc_lojaorigem = '" & LojaOrigem & "'" _
        & "order by PR_CodigoFornecedor, PR_Descricao"
 
        adorsExtra1.CursorLocation = adUseClient
        adorsExtra1.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
     

        adorsExtra3.CursorLocation = adUseClient
        adorsExtra3.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
       
       
        If adorsExtra1.EOF Then
          MsgBox "Nota Fiscal não Encontrada"
          Exit Function
       End If
       
       serie = adorsExtra1("vc_serie")
 
 
       
      
       
       
        wCFO1 = " "
        wCFO2 = " "
        wCFO3 = " "
       
       If Not adorsExtra3.EOF Then
          Do While Not adorsExtra3.EOF
         
             
             If Trim(adorsExtra3("Vc_CodigoOperacaoNovo")) = "5152" Then
                If adorsExtra3("PR_substituicaotributaria") = "S" Then
                   wCFO2 = "5409"
                Else
                   wCFO1 = "5152"
                End If
                WcodigoOperacao = Trim(wCFO1) & wCFO3 & Trim(wCFO2)
             Else
                WcodigoOperacao = Trim(adorsExtra3("Vc_CodigoOperacaoNovo"))
             End If
                
          adorsExtra3.MoveNext
          Loop
 
           ADO_Cn_CDLocal.BeginTrans
            sql = "Update capanfvenda " _
                  & "Set Vc_CodigoOperacaoNovo = '" & Trim(WcodigoOperacao) & "'" _
                  & " where  vc_notafiscal= " & NotaFiscal & " and vc_serie='" & serie & "' and vc_lojaorigem = '" & LojaOrigem & "'"
           ADO_Cn_CDLocal.Execute (sql)
            ADO_Cn_CDLocal.CommitTrans
         
           
        End If
    
       
    
        
       'ROTINA PRINCIPAL
     '   tmporient = Printer.Orientation
        wConta = 0
        wChave = 0
        wReduz = 0
        wAliquotaZero = False
        wStr15 = ""
        wStr16 = ""
        wRes = 0
        Wnatureza = ""
        
        
        
        Do While Not adorsExtra1.EOF
           flg = flg + 1
          
           DoEvents
           If ParaImpr Then
              Printer.Print "***** INTERROMPIDO PELO USUÁRIO *****"
              Printer.EndDoc
              Exit Function
           End If
                 
           If wChave = 0 Then
             
              sql = "select lo_endereco,lo_bairro,lo_municipio,lo_uf," _
                  & "lo_cep,lo_cgc,lo_inscricaoestadual,lo_fax,lo_telefone " _
                  & " from loja " _
                  & " where lo_loja = '" & adorsExtra1("vc_lojaorigem") & "' "
               
              
              adorsExtra2.CursorLocation = adUseClient
            adorsExtra2.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
              
              If Not adorsExtra2.EOF Then
                 
             
                 If wTelaOperacaoEspecial = True Then
                    wStr17 = ""
                 Else
                    wStr17 = adorsExtra1("vc_enderecocliente")
                 End If
                 
                 wStr1 = Space(2) & left(Format(wStr17) & Space(34), 34) & left(Format(Trim(adorsExtra2("lo_endereco")), ">") & Space(34), 34) & left(Format(Trim(adorsExtra2("lo_bairro")), ">") & Space(11), 11) & Space(5) & "X" & Space(26) & left(Format(adorsExtra1("vc_notafiscal"), "######"), 7)
                 
                 If wTelaOperacaoEspecial = True Then
                    wStr18 = ""
                 Else
                    wStr18 = IIf(IsNull(adorsExtra1("VC_Observacao")), "", adorsExtra1("VC_Observacao"))
                 End If
                 wStr2 = Space(2) & left(Format(wStr18) & Space(34), 34) & left(Format(Trim(adorsExtra2("lo_municipio"))) & Space(15), 15) & Space(24) & left$(Trim(adorsExtra2("lo_uf")), 2)
                 wStr3 = Space(2) & left$(Format(wStr19) & Space(34), 34) & "(011)" & left$(Trim(Format(adorsExtra2("lo_telefone"), "####-####")), 9) & "/(011)" & left$(Format(adorsExtra2("lo_fax"), "####-####"), 9) & Space(5) & left$(Format(adorsExtra2("lo_cep"), "00###-###"), 9)
                 wStr20 = ""
                 wStr4 = Space(2) & left(Format(wStr20) & Space(40), 40) & Space(46) & left$(Trim(Format(adorsExtra2("lo_cgc"), "###,###,###")), 10) & "/" & right$(Format(adorsExtra2("lo_cgc"), "####-##"), 7)
                 wEspaco = ""
                 wLO_inscricaoestadual = adorsExtra2("lo_inscricaoestadual")
                 
        
                 
                 If wTelaOperacaoEspecial = True Then
                    sql = "Select * from codigooperacaonovo " _
                        & "where CN_CodigoOperacaoNovo = " & Trim(adorsExtra1("Vc_CodigoOperacaoNovo")) & ""
                          
                      
                    rdorsNatOper.CursorLocation = adUseClient
                    rdorsNatOper.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
            
                    
                    If rdorsNatOper.EOF = False Then
                       wDescricao = rdorsNatOper("CN_DescricaoOperacao")
                    Else
                       wDescricao = ""
                    End If
                    wStr5 = Space(2) & left$(wEspaco & Space(31), 31) & left$(WcodigoOperacao, 10) & " - " & Format(Trim(wDescricao), ">") & Space(1) & Space(31) & left$(Trim(Format((adorsExtra2("lo_inscricaoestadual")), "###,###,###,###")), 15)
                 Else
                    wDescricao = "TRANSFERENCIA"
                    wStr5 = Space(2) & left$(wEspaco & Space(32), 32) & Format(Trim(wDescricao), ">") & Space(9) & left$(WcodigoOperacao, 10) & Space(31) & left$(Trim(Format((adorsExtra2("lo_inscricaoestadual")), "###,###,###,###")), 15)
                 End If
                
              End If
            
            If wTelaOperacaoEspecial = True Then
             
               wStr6 = left(Trim(wEspaco) & Space(31), 31) & left$(Format(Trim(wEspaco)) & Space(0), 0) & left$(Format(Trim(adorsExtra1("vc_nomecliente")), ">") & Space(56), 56) & left$(Trim(Format(adorsExtra1("vc_cgccliente"), "###,###,###")), 10) & "/" & right$(Format(adorsExtra1("vc_cgccliente"), "####-##"), 7) & Space(7) & left$(Format(adorsExtra1("vc_dataemissao"), "dd/mm/yy") & Space(12), 12)
               wStr7 = "               "
               wStr7 = Space(2) & left(wStr7 & Space(29), 29) & left$(Format(Trim(adorsExtra1("vc_enderecocliente")), ">") & Space(42), 42) & left$(Format(Trim(adorsExtra1("vc_bairrocliente")), ">") & Space(21), 21) & right$(Space(11) & Format(adorsExtra1("vc_cepcliente"), "00###-###"), 11) & Space(7) & left$(Format(adorsExtra1("vc_dataemissao"), "dd/mm/yy"), 12)
               wEspaco = ""
               wStr8 = Space(31) & left$(Format(Trim(adorsExtra1("vc_municipiocliente")), ">") & Space(15), 15) & Space(19) & left$(Format(Trim(wEspaco)) & Space(15), 15) & left$(Trim(adorsExtra1("vc_ufcliente")), 2) & Space(5) & left$(Trim(Format(adorsExtra1("vc_InscEstCliente"), "###,###,###,###")), 15)
             
            Else
                 sql = "select em_descricao,lo_endereco,lo_bairro," _
                     & "lo_municipio,lo_uf,lo_cep,lo_cgc," _
                     & "lo_inscricaoestadual,lo_fax,lo_telefone " _
                     & " from loja, empresa " _
                     & " where lo_empresa=em_codigoempresa " _
                     & " and lo_loja = '" & adorsExtra1("vc_lojadestino") & "' "
                     
                     adordoExtra2.CursorLocation = adUseClient
                    adordoExtra2.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
                    
                 If Not adordoExtra2.EOF Then
                    
                    wStr6 = left(Trim(wEspaco) & Space(31), 31) & left$(Format(Trim(wEspaco)) & Space(0), 0) & left$(Format(Trim(adordoExtra2("em_descricao")), ">") & Space(56), 56) & left$(Trim(Format(adordoExtra2("lo_cgc"), "###,###,###")), 10) & "/" & right$(Format(adordoExtra2("lo_cgc"), "####-##"), 7) & Space(7) & left$(Format(adorsExtra1("vc_dataemissao"), "dd/mm/yy") & Space(12), 12)
                    wStr7 = "               "
                    wStr7 = Space(2) & left(wStr7 & Space(29), 29) & left$(Format(Trim(adordoExtra2("lo_endereco")), ">") & Space(42), 42) & left$(Format(Trim(adordoExtra2("lo_bairro")), ">") & Space(21), 21) & right$(Space(11) & Format(adordoExtra2("lo_cep"), "00###-###"), 11) & Space(7) & left$(Format(adorsExtra1("vc_dataemissao"), "dd/mm/yy"), 12)
                    wEspaco = ""
                    wStr8 = Space(31) & left$(Format(Trim(adordoExtra2("lo_municipio")), ">") & Space(15), 15) & Space(19) & left$(Format(Trim(wEspaco)) & Space(15), 15) & left$(Trim(adordoExtra2("lo_uf")), 2) & Space(5) & left$(Trim(Format(adordoExtra2("lo_inscricaoestadual"), "###,###,###,###")), 15)
                 End If
            End If
               
                 
                 adorsExtra2.Close
               '  adordoExtra2.Close
                            
              wStr9 = right$(Space(9) & Format(adorsExtra1("vc_baseicms"), "######0.00"), 9) & right$(Space(9) & Format(adorsExtra1("vc_valoricms"), "######0.00"), 9) & Space(35) & right$(Space(10) & Format(adorsExtra1("vc_valormercadorias"), "######0.00"), 10)
              wStr10 = right(Space(9) & Format(Space(9) & wEspaco, "######0.00"), 9) & Space(44) & right(Space(10) & Format(adorsExtra1("vc_totalnota"), "######0.00"), 10)
              
              If wTelaOperacaoEspecial = True Then
                 wStr11 = ""
                 wStr12 = ""
              Else
                 wStr11 = Space(0) & wnome
                 wStr12 = Space(0) & wendereco & "              " & wbairro & "            " & westado
              End If
              
            
              wStr13 = Space(74) & "LOJA " & LojaOrigem & Space(10) & right$(Space(7) & Format(adorsExtra1("vc_notafiscal"), "###,###"), 7)
            
                    Printer.Print
                    Printer.Print
                    If wTelaOperacaoEspecial = True Then
                       Printer.Print
                    Else
                       Printer.Print "  ROMANEIO:"
                    End If
                    Printer.Print wStr1
                    Printer.Print wStr2
                    Printer.Print wStr3
                    Printer.Print wStr4
                    Printer.Print
                    Printer.Print wStr5
                    Printer.Print
                    Printer.CurrentY = Printer.CurrentY + 2
                    Printer.Print wStr6
                    Printer.Print
                    Printer.CurrentY = Printer.CurrentY - 2
                    Printer.Print
                    Printer.Print wStr7
                    Printer.Print
                    Printer.Print wStr8
                    Printer.Print
                    Printer.Print
                    
             
         End If
           wConta = wConta + 1
           wStr1 = ""
          
           
              sql = "select pr_codigoipi,pr_codigoreducaoicms,pr_descricao," _
                  & "pr_classefiscal,pr_unidade " _
                  & "from produto " _
                  & "where pr_referencia = '" & adorsExtra1("vi_referencia") & "' "
            
               adorsExtra2.CursorLocation = adUseClient
               adorsExtra2.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
                    
            
              wCodIPI = 0
              If adorsExtra2("pr_codigoipi") = 4 Then
                 wCodIPI = 1
              End If
              If adorsExtra2("pr_codigoipi") = 5 Then
                 wCodIPI = 2
              End If
              If adorsExtra2("pr_codigoreducaoicms") <> 0 Then
                 wCodTri = 2
              Else
                 wCodTri = 0
              End If
              If adorsExtra2("pr_codigoreducaoicms") <> 0 Then
                 wReduz = 1
                 If wStr15 = "" Then
                     wStr15 = wStr15 & wConta
                  Else
                     wStr15 = wStr15 & "," & wConta
                  End If
              End If
              
                            
              If adorsExtra1("vi_reserva") <> "" And adorsExtra1("vi_reserva") <> 0 Then
                 wRes = 1
                 If wStr16 = "" Then
                    wStr16 = wStr16 & wConta
                 Else
                    wStr16 = wStr16 & "," & wConta
                 End If
              End If
              
              wEspaco = wCodIPI
              wEspaco = wEspaco & wCodTri
              
              If adorsExtra1("vi_aliquotaicms") = "0" Then
                 wAliquotaZero = True
              End If
              
           
              
                  wStr1 = ""
                  wStr1 = left$(adorsExtra1("vi_referencia") & Space(7), 7) _
                         & Space(1) & left$(Format(Trim(adorsExtra2("pr_descricao")), ">") & Space(38), 38) _
                         & Space(16) & left$(Format(Trim(adorsExtra2("pr_classefiscal")), ">") _
                         & Space(12), 12) & left$(Trim(wEspaco) & Space(3), 3) _
                         & "" & Space(3) & left$(Trim(adorsExtra2("pr_unidade")) & Space(2), 2) _
                         & right$(Space(6) & Format(adorsExtra1("vi_quantidade"), "##0"), 6) _
                         & right$(Space(12) & Format(adorsExtra1("vi_precounitario"), "#####0.00"), 14) _
                         & right$(Space(15) & Format(adorsExtra1("vi_valormercadoria"), "#####0.00"), 15) & Space(1) _
                         & right$(Space(2) & Format(adorsExtra1("vi_aliquotaicms"), "#0"), 2)

              
              Printer.Print wStr1
            
              wStr1 = ""
              wChave = 1
              adorsExtra2.Close
              adorsExtra1.MoveNext
        Loop
            
           Do While wConta < 10
              wConta = wConta + 1
              Printer.Print
           Loop
     
     
  If wTelaOperacaoEspecial = True Then
      

           
     sql = "Select on_carimbo1,on_carimbo2 " _
          & "from observacaonotafiscal " _
          & "where ON_NumeroNotaFiscal= " & UCase(NotaFiscal) & " " _
          & "and ON_Serie in ('" & serie & "') " _
          & "and ON_Loja='" & LojaOrigem & "'"
        
   
          adorsExtra6.CursorLocation = adUseClient
          adorsExtra6.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
               
         
         
          If Not adorsExtra6.EOF Then
             Printer.Print Trim(adorsExtra6("on_carimbo1"))
             Printer.Print Trim(adorsExtra6("on_carimbo2"))
             Printer.Print
          Else
             Printer.Print
             Printer.Print
             Printer.Print
          End If
   
  Else
     If wReduz = 1 Then
        Printer.Print Space(4) & "   BASE CALC REDUZ CONF.ART.51. ANEXOS I E II ART. 12 - I,II,III E IV DECR.45.490 --> ITENS " & wStr15
       
     Else
        Printer.Print
      
     End If
  
     If wRes = 1 Then
        Printer.Print Space(4) & "   ITENS C/RESERVA " & wStr16
     Else
        Printer.Print
     End If
    
     If wAliquotaZero = True Then
        Printer.Print Space(4) & "   ALIQ.ICMS = 0  IMPOSTO RECOLHIDO POR SUBSTITUICAO - ART 313-Z3,Z11,Z17 e Z19 DO RICMS "
     Else
        Printer.Print
     End If
  End If
   

           Printer.Print
           Printer.Print
           Printer.Print wStr9
           Printer.Print
           Printer.Print wStr10
           Printer.Print
           Printer.Print
           Printer.Print wStr11
           Printer.Print
           Printer.Print wStr12
           Printer.CurrentY = Printer.CurrentY + 1
           Printer.Print
           Printer.Print
           Printer.Print
           Printer.Print
           Printer.Print
           
                    
           Printer.Print wStr13
           Printer.CurrentY = Printer.CurrentY - 1
          
           '------------------------------------------------------------------------------
                 'Acerto emissao de nota com mais de um formulario
                       ' Printer.EndDoc
                        
                         Printer.Print ""
                         Printer.Print ""
                         Printer.Print ""
                         Printer.Print ""
                         Printer.Print ""
                         
                         wControlaQuebraDaPagina = wControlaQuebraDaPagina + 1
                         If wControlaQuebraDaPagina = 3 Then
                            Printer.Print ""
                            wControlaQuebraDaPagina = 0
                         End If
           '----------------------------------------------------------------------------------
          ' Printer.Orientation = 1
           adorsExtra1.Close
           adorsExtra3.Close
         '  rdorsNatOper.Close
         '  adorsExtra6.Close
           wTelaOperacaoEspecial = False
           
           Printer.EndDoc
End Function


Public Sub ImprimeBarcode(ByVal bc_string As String, Xref As Double, Yref As Double)

    Dim xpos As Double
    Dim Y1 As Double
    Dim Y2 As Double
    Dim dw As Double
    Dim th As Double
    Dim tw As Double
    Dim new_string As String
    Dim RefPixelX As Double
    Dim RefPixelY As Double
    Dim SalvaEscala As Long
    Dim n As Integer
    Dim c As Integer
    Dim i As Integer
    Dim bc_pattern$

    
    If Trim(bc_string) = "" Then
        Exit Sub
    End If
    
    'define barcode patterns
    Dim bc(90) As String
    
    bc(1) = "1 1221"            'pre-amble
    bc(2) = "1 1221"            'post-amble
    bc(48) = "11 221"           'digits
    bc(49) = "21 112"
    bc(50) = "12 112"
    bc(51) = "22 111"
    bc(52) = "11 212"
    bc(53) = "21 211"
    bc(54) = "12 211"
    bc(55) = "11 122"
    bc(56) = "21 121"
    bc(57) = "12 121"
                                'capital letters
    bc(65) = "211 12"           'A
    bc(66) = "121 12"           'B
    bc(67) = "221 11"           'C
    bc(68) = "112 12"           'D
    bc(69) = "212 11"           'E
    bc(70) = "122 11"           'F
    bc(71) = "111 22"           'G
    bc(72) = "211 21"           'H
    bc(73) = "121 21"           'I
    bc(74) = "112 21"           'J
    bc(75) = "2111 2"           'K
    bc(76) = "1211 2"           'L
    bc(77) = "2211 1"           'M
    bc(78) = "1121 2"           'N
    bc(79) = "2121 1"           'O
    bc(80) = "1221 1"           'P
    bc(81) = "1112 2"           'Q
    bc(82) = "2112 1"           'R
    bc(83) = "1212 1"           'S
    bc(84) = "1122 1"           'T
    bc(85) = "2 1112"           'U
    bc(86) = "1 2112"           'V
    bc(87) = "2 2111"           'W
    bc(88) = "1 1212"           'X
    bc(89) = "2 1211"           'Y
    bc(90) = "1 2211"           'Z
                                'Misc
    bc(32) = "1 2121"           'space
    bc(35) = ""                 '# cannot do!
    bc(36) = "1 1 1 11"         '$
    bc(37) = "11 1 1 1"         '%
    bc(43) = "1 11 1 1"         '+
    bc(45) = "1 1122"           '-
    bc(47) = "1 1 11 1"         '/
    bc(46) = "2 1121"           '.
    bc(64) = ""                 '@ cannot do!
    bc(65) = "1 1221"           '*
    
    bc_string = UCase(bc_string)
    
    With Printer
        SalvaEscala = .ScaleMode
        
        .CurrentX = Xref
        .CurrentY = Yref

'        RefPixelX = Abs(Int(-Printer.ScaleX(Xref, vbMillimeters, vbPixels)))
'        RefPixelY = Abs(Int(-Printer.ScaleY(Yref, vbMillimeters, vbPixels)))

        RefPixelX = Printer.ScaleX(Xref, vbMillimeters, vbPixels)
        RefPixelY = Printer.ScaleY(Yref, vbMillimeters, vbPixels)
        
        'dimensions
        .ScaleMode = vbPixels
        
'        .CurrentX = RefPixelX
'        .CurrentY = RefPixelY
        
        dw = 7      'CInt(.ScaleHeight / 40)                    'space between bars
        'If dw < 1 Then dw = 1
        th = .TextHeight(bc_string)                     'text height
        tw = .TextWidth(bc_string)                      'text width
        
        new_string = Chr$(1) & bc_string & Chr$(2)      'add pre-amble, post-amble
        
        Y1 = .CurrentY '.ScaleTop
        Y2 = .CurrentY + 90 '.ScaleTop + .ScaleHeight - 1.5 * th
'        .Width = 1.1 * Len(new_string) * (15 * dw) * .Width / .ScaleWidth
        
        
        'draw each character in barcode string
        xpos = RefPixelX
        For n = 1 To Len(new_string)
            c = Asc(Mid$(new_string, n, 1))
            If c > 90 Then c = 0
            bc_pattern$ = bc(c)
            
            'draw each bar
            For i = 1 To Len(bc_pattern$)
                Select Case Mid$(bc_pattern$, i, 1)
                    Case " "
                        'space
                        Printer.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
                        xpos = xpos + dw
                        
                    Case "1"
                        'space
                        Printer.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
                        xpos = xpos + dw
                        'line
                        Printer.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &H0&, BF
                        xpos = xpos + dw
                    
                    Case "2"
                        'space
                        Printer.Line (xpos, Y1)-(xpos + 2 * dw, Y2), &HFFFFFF, BF
                        xpos = xpos + dw
                        'wide line
                        Printer.Line (xpos, Y1)-(xpos + 2 * dw, Y2), &H0&, BF
                        xpos = xpos + 2 * dw
                End Select
            Next
        Next
        
        '1 more space
        Printer.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
        xpos = xpos + dw
        
        .CurrentX = RefPixelX
        .CurrentY = Y2 + 0.25 * th
        
        Printer.FontName = "Arial"
        Printer.FontSize = 8
        
        Printer.Print bc_string & "  ";
        
        .ScaleMode = SalvaEscala
    End With

End Sub

Function MesesTendencia() As Variant

    Dim Vezes As Long
    Dim Retorno As String
    Dim mes As String
    Dim Data As Date
    Dim Meses(6) As String
    
    
    Data = Date
    Retorno = ""
    For Vezes = 0 To 6 Step 1
        mes = TraduzMes(Month(DateAdd("m", -(Vezes), Data)))
        Meses(Vezes) = mes
    Next Vezes
    
    MesesTendencia = Array(Meses(0), Meses(1), Meses(2), Meses(3), Meses(4), Meses(5), Meses(6))

End Function

Function TraduzMes(ByVal mes As Integer) As String

    Select Case mes
        Case 1, -11: TraduzMes = "Janeiro"
        Case 2, -10: TraduzMes = "Fevereiro"
        Case 3, -9: TraduzMes = "Março"
        Case 4, -8: TraduzMes = "Abril"
        Case 5, -7: TraduzMes = "Maio"
        Case 6, -6: TraduzMes = "Junho"
        Case 7, -5: TraduzMes = "Julho"
        Case 8, -4: TraduzMes = "Agosto"
        Case 9, -3: TraduzMes = "Setembro"
        Case 10, -2: TraduzMes = "Outubro"
        Case 11, -1: TraduzMes = "Novembro"
        Case 12, 0: TraduzMes = "Dezembro"
    End Select
    
End Function

Function CalculaPeriodoMesSQL(ByVal Data As String) As String

    Dim Periodo As String
    Dim mes As String
    
    Periodo = ""
    
    mes = Format(Data, "mm/yyyy")
    Periodo = " '" & left(mes, 2) & "/01/" & right(mes, 4) & "' and '"
    
    If IsDate("31/" & mes) Then
        Periodo = Periodo & left(mes, 2) & "/31/" & right(mes, 4) & "' "
    ElseIf IsDate("30/" & mes) Then
        Periodo = Periodo & left(mes, 2) & "/30/" & right(mes, 4) & "' "
    ElseIf IsDate("29/" & mes) Then
        Periodo = Periodo & left(mes, 2) & "/29/" & right(mes, 4) & "' "
    Else
        Periodo = Periodo & left(mes, 2) & "/28/" & right(mes, 4) & "' "
    End If
    
    CalculaPeriodoMesSQL = Periodo

End Function


Public Function DadosLoja()
    Dim rsdadosLoja As New ADODB.Recordset
 
    sql = ""
    sql = "Select CTS_SerieNota,CTS_Loja,CTS_SenhaLiberacao,CTS_LogoPedido,Loja.* from loja,ControleSistema where lo_loja=CTS_Loja"
 
    rsdadosLoja.CursorLocation = adUseClient
    rsdadosLoja.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
 
    If Not rsdadosLoja.EOF Then
       wSerie = Trim(rsdadosLoja("ctS_serienota"))
       wRazao = Trim(rsdadosLoja("lo_nomeloja"))
       wendereco = rsdadosLoja("lo_ENDERECO")
       wbairro = rsdadosLoja("lo_bairro")
       WCGC = rsdadosLoja("lo_CGC")
       WIest = rsdadosLoja("lo_INSCRICAOESTADUAL")
       wMunicipio = rsdadosLoja("lo_MUNICIPIO")
       westado = rsdadosLoja("lo_UF")
       WCep = rsdadosLoja("lo_CEP")
       wFone = rsdadosLoja("lo_TELEFONE")
     '  wDDDLoja = rsdadosLoja("LO_DDD")
       WFax = rsdadosLoja("lo_Fax")
       wLoja = rsdadosLoja("CTS_Loja")
       GLB_Loja = rsdadosLoja("CTs_Loja")
       GLB_logoPedido = Trim(rsdadosLoja("CTS_LogoPedido"))
      ' wNovaRazao = IIf(IsNull(rsdadosLoja("lo_NovaRazao")), "0", rsdadosLoja("lo_NovaRazao"))
    
    End If
    
    rsdadosLoja.Close
 
End Function

Function fDataExt(ByVal wDatConv)
    Dim wDExtConv As String, wMExtConv As String
    Select Case Weekday(wDatConv)
        Case 1: wDExtConv = "Domingo"
        Case 2: wDExtConv = "Segunda-Feira"
        Case 3: wDExtConv = "Terça-Feira"
        Case 4: wDExtConv = "Quarta-Feira"
        Case 5: wDExtConv = "Quinta-Feira"
        Case 6: wDExtConv = "Sexta-Feira"
        Case 7: wDExtConv = "Sábado"
    End Select
    Select Case Month(wDatConv)
        Case 1: wMExtConv = "Janeiro"
        Case 2: wMExtConv = "Fevereiro"
        Case 3: wMExtConv = "Março"
        Case 4: wMExtConv = "Abril"
        Case 5: wMExtConv = "Maio"
        Case 6: wMExtConv = "Junho"
        Case 7: wMExtConv = "Julho"
        Case 8: wMExtConv = "Agosto"
        Case 9: wMExtConv = "Setembro"
        Case 10: wMExtConv = "Outubro"
        Case 11: wMExtConv = "Novembro"
        Case 12: wMExtConv = "Dezembro"
    End Select
    fDataExt = wDExtConv & ", " & Format$(Date, "dd") & " de " & wMExtConv & " de " & Format$(Date, "yyyy")
End Function


Sub PegaNumeroPedido()
 Screen.MousePointer = 11
 
    sql = "Select * from Controle "
    
    adoControle.CursorLocation = adUseClient
    adoControle.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    NroPedido = adoControle("CT_NumPed")
    pedido = NroPedido
    adoControle.Close
    
    ADO_Cn_CDLocal.BeginTrans
       
    sql = "Update Controle set CT_NumPed = " & NroPedido & " + 1"
   ADO_Cn_CDLocal.Execute (sql)
    ADO_Cn_CDLocal.CommitTrans
    
  '  CriaCapaPedido NroPedido

' GravaItensPedido NroPedido, 11, 725
Screen.MousePointer = vbNormal
End Sub

Public Function ExtraiSeqNotaControle00() As Double
     Dim WnovaSeqNota As Long
     Dim rsDados As New ADODB.Recordset
     sql = ""
     sql = "Select CTS_Numero00 + 1 as NumNota from ControleSistema"
   
     rsDados.CursorLocation = adUseClient
     rsDados.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
     
     If Not rsDados.EOF Then
        
        ExtraiSeqNotaControle00 = rsDados("NumNota")
        sql = "update ControleSistema set CTS_Numero00 = " & rsDados("NumNota") & ""
        ADO_Cn_CDLocal.Execute (sql)
     End If
     rsDados.Close
End Function

Public Function ExtraiSeqNotaControle() As Double
     Dim WnovaSeqNota As Long
     Dim rsDados As New ADODB.Recordset
     sql = ""
     sql = "Select CTS_NumeroNF + 1 as NumNota from ControleSistema"
   
     rsDados.CursorLocation = adUseClient
     rsDados.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
     
     If Not rsDados.EOF Then
        
        ExtraiSeqNotaControle = rsDados("NumNota")
        sql = "update ControleSistema set CTS_NumeroNF= " & rsDados("NumNota") & ""
        ADO_Cn_CDLocal.Execute (sql)
     End If
     rsDados.Close
End Function

Public Function ExtraiSeqNotaControleNE() As Double
     Dim WnovaSeqNota As Long
     Dim rsDados As New ADODB.Recordset
     sql = ""
     sql = "Select CTS_NumeroNE + 1 as NumNota from ControleSistema"
   
     rsDados.CursorLocation = adUseClient
     rsDados.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
     
     If Not rsDados.EOF Then
        
        ExtraiSeqNotaControleNE = rsDados("NumNota")
        sql = "update ControleSistema set CTS_NumeroNE = " & rsDados("NumNota") & ""
        ADO_Cn_CDLocal.Execute (sql)
     End If
     rsDados.Close
End Function

'Function GravaItensPedido(ByVal NumeroPedido As Double, ByVal Vendedor As Integer)
'
'    sql = ""
'      ADO_Cn_CDLocal.BeginTrans
'
'  '   SQL = "Insert into NfItens (nf, NUMEROPED,Serie, DATAEMI, REFERENCIA, QTDE, VLUNIT, " _
'        & "VLTOTITEM, ICMS, PLISTA,  " _
'        & "LOJAORIGEM,  TIPONOTA,  Item, cfop, desconto, ReferenciaAlternativa) " _
'        & "Values (" & NroNotaFiscal & "," & NroPedido & ", '" & wSerie & "', '" & Format(Date, "mm/dd/yyyy") & "', '" _
'        & wCodigoProduto & "', '" & wQtde & "', " _
'        & "" & ConverteVirgula(Format(wItemVenda, "0.00")) & ", " _
'        & ConverteVirgula(Format(wVlTotItem, "0.00")) & ", " & ConverteVirgula(Format(wICMS, "0.00")) & ", " _
'        & "  " & ConverteVirgula(Format(wPLISTA, "0.00")) & ",  '" _
'        & GLB_Loja & "', '" & wTipoNota & "', '" & wNroItens & "', '" & wCFOP & "', 0, 0)"
'
'    sql = "Insert into ItemNFVenda(VI_NotaFiscal, VI_Serie, VI_LojaOrigem, VI_DataEmissao, VI_NumeroItem, " _
'            & "VI_Referencia, VI_Quantidade, VI_PrecoUnitario, VI_ValorMercadoria, VI_AliquotaICMS," _
'            & "VI_AliquotaIPI, VI_ValorIPI,VI_ValorICMS,VI_BaseICMS,VI_PrecoLista, VI_PesoBruto, VI_PesoLiquido," _
'            & "VI_TipoNota,VI_Situacao, vi_customediounit, vi_precocustounit, vi_desconto, vi_reserva) " _
'            & "Values(" & NroNotaFiscal & ",'" & wSerie & "','" _
'            & lojaorigem & "','" & Format(Date, "yyyy/mm/dd") & "'," & frmEncerraNFOutrasOperacoes.grdItens.Rows - 1 & ",'" _
'            & wCodigoProduto & "', " & wQtde & ", " _
'            & wItemVenda & ", " _
'            & ConverteVirgula(Format(frmEncerraNFOutrasOperacoes.grdItens.TextMatrix(frmEncerraNFOutrasOperacoes.grdItens.Row, 13), "0.00")) & "," _
'            & ConverteVirgula(Format(frmEncerraNFOutrasOperacoes.grdItens.TextMatrix(frmEncerraNFOutrasOperacoes.grdItens.Row, 4), "0.00")) & ",0,0," _
'            & ConverteVirgula(Format(frmEncerraNFOutrasOperacoes.grdItens.TextMatrix(frmEncerraNFOutrasOperacoes.grdItens.Row, 10), "0.00")) & "," _
'            & ConverteVirgula(Format(frmEncerraNFOutrasOperacoes.grdItens.TextMatrix(frmEncerraNFOutrasOperacoes.grdItens.Row, 9), "0.00")) & "," _
'            & ConverteVirgula(Format(frmEncerraNFOutrasOperacoes.grdItens.TextMatrix(frmEncerraNFOutrasOperacoes.grdItens.Row, 3), "0.00")) _
'            & ",0,0,'" & wTipoNota & "','P',0,0,0,0)"
'
'    ADO_Cn_CDLocal.Execute (sql)
'    ADO_Cn_CDLocal.CommitTrans
'
'End Function

Function CriaCapaPedido(ByVal NumeroPedido As Double, ByVal Vendedor As Integer)
      
      
      sql = ""
      sql = "Select count(vi_referencia) as NumeroItem from itemnfvenda " _
          & "where vi_notafiscal = '" & NroNotaFiscal & "' and vi_dataemissao = '" & Format(Date, "yyyy/mm/dd") & "'"
          
          adoContaItens.CursorLocation = adUseClient
          adoContaItens.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
  
    
    sql = ""
     ADO_Cn_CDLocal.BeginTrans
   ' SQL = "Insert into NfCapa (NUMEROPED,Serie, DATAEMI, " _
        & "LOJAORIGEM, TIPONOTA, Vendedor, DATAPED, HORA, " _
        & " VendedorLojaVenda, LojaVenda,TM,ECF,nf,qtditem, cliente, totalnota, vlrmercadoria, ecfnf, condpag) " _
        & "Values (" & NroPedido & ", '" & wSerie & "', '" & Format(Date, "mm/dd/yyyy") & "', " _
        & "'" & GLB_Loja & "', '" & wTipoNota & "', '" & Vendedor & "', '" _
        & Format(Date, "mm/dd/yyyy") & "', '" & Format(time, "hh:mm:ss") & "', '" _
        & Vendedor & "', '" & GLB_Loja & "',0," & GLB_ECF & "," & NroNotaFiscal & ", '" & wNroItens & "', '" & _
        cliente & "', " & ConverteVirgula(wTotNota) & "," & ConverteVirgula(wVlrMercadoria) & ", 0, '01')"
        
    sql = "Insert into CapaNFVenda (VC_NotaFiscal, VC_Serie, VC_LojaOrigem, " _
             & "VC_LojaDestino, VC_DataEmissao, VC_CGCLojaDestino, VC_ValorMercadorias, VC_BaseICMS, " _
             & "VC_AliquotaICMS, VC_ValorICMS, VC_ValorIPI, VC_EncargosFinanceiros, VC_TotalNota," _
             & "VC_TipoNota, VC_SituacaoComunicacao, VC_Situacao, VC_Pesobruto, VC_PesoLiquido," _
             & " VC_NumeroPedido, VC_Cliente,VC_NomeCliente,VC_EnderecoCliente,VC_BairroCliente," _
             & "VC_MunicipioCliente,VC_UFCliente,VC_CEPCliente,VC_TelefoneCliente,VC_EndEntregaCliente," _
             & "VC_CGCCliente,VC_InscEstCliente," _
             & "VC_HoraEmissao, VC_CodigoOperacaoNovo, VC_DataProcessamento,VC_LojaVenda, " _
             & "vc_avistareceber, vc_financiada, vc_faturada, vc_notacredito, vc_deposito, " _
             & "vc_cartao, vc_chequepre, vc_cheque, vc_dinheiro, vc_pedidocliente, vc_valorfretecobrado, " _
             & "vc_valorfrete, vc_desconto, vc_pagamentoentrada, vc_av, vc_codigooperacao, vc_itens)" _
             & " values (" & NroNotaFiscal & ", ' " & wSerie & "', " _
             & "'" & LojaOrigem & "','0','" & Format(Date, "yyyy/mm/dd") & "'," _
             & "'" & frmEncerraNFOutrasOperacoes.txtCNPJDestinatario.Text & "', " & ConverteVirgula(Format(frmEncerraNFOutrasOperacoes.txtValormercadoria.Text, "0.00")) & ", " _
             & ConverteVirgula(Format(frmEncerraNFOutrasOperacoes.txtBaseICMS.Text, "0.00")) & ",0," _
             & ConverteVirgula(Format(frmEncerraNFOutrasOperacoes.txtValorICMS.Text, "0.00")) & ",0,0," _
             & ConverteVirgula(Format(frmEncerraNFOutrasOperacoes.txtTotalNF.Text, "0.00")) & ", 'S', " _
             & "'P', 'A',0,0," & frmEncerraNFOutrasOperacoes.txtPedido.Text & "," & frmEncerraNFOutrasOperacoes.txtCodigoDestinatario.Text & ",'" _
             & frmEncerraNFOutrasOperacoes.txtDestinatario.Text & "','" & frmEncerraNFOutrasOperacoes.txtEnderecoDestinatario.Text & "','" _
             & frmEncerraNFOutrasOperacoes.txtBairroDestinatario.Text & "','" & frmEncerraNFOutrasOperacoes.txtMunicipioDestinatario.Text & "','" _
             & frmEncerraNFOutrasOperacoes.txtUFDestinatario.Text & "', '" & frmEncerraNFOutrasOperacoes.txtCepEmitente.Text & "', '" _
             & Mid(frmEncerraNFOutrasOperacoes.txtFoneFaxDestinatario.Text, 1, 8) & "','" & frmEncerraNFOutrasOperacoes.txtEnderecoDestinatario.Text & "','" _
             & frmEncerraNFOutrasOperacoes.txtCNPJDestinatario.Text & "','" & frmEncerraNFOutrasOperacoes.txtInscricaoDestinatario.Text & "',0, " _
             & "'" & Trim(Mid(frmEncerraNFOutrasOperacoes.cmbCFOP.Text, 1, 4)) & "', '" & Format(Date, "yyyy/mm/dd") & "', '" & frmEncerraNFOutrasOperacoes.cmbLojaOrigem.Text & "', '', '', " _
             & "'','', '', '', '', '', '', 0,0, 0, 0, 0, '', 0, " & "" & ")"
     ADO_Cn_CDLocal.Execute (sql)
    ADO_Cn_CDLocal.CommitTrans
     
     
     adoContaItens.Close
    
     
End Function

Function EncerraVenda(ByVal NumeroDocumento As Double, ByVal SerieDocumento As String, ByVal TipoAtualizacaoEstoque As Double) As Boolean
    
    Dim SerieProd As String
    
        wQuantdadeTotalItem = 0
        wAnexo = ""
        wAnexo1 = ""
        wAnexo2 = ""
        wQuantItensCapaNF = 0
        wCFO2 = " "
        wCFO1 = " "
        wChaveICMS = 0
        GLB_TotalICMSCalculado = 0
        GLB_ValorCalculadoICMS = 0
        GLB_BasedeCalculoICMS = 0
        GLB_AliquotaAplicadaICMS = 0
        GLB_AliquotaICMS = 0
        GLB_BaseTotalICMS = 0
        GLB_Tributacao = 0
        wCFOItem = 0
        wUltimoItem = 0
        wComissaoVenda = 0
        wSomaVenda = 0
        wSomaMargem = 0
        wCarimbo5 = ""
        wCarimbo2 = ""
        EncerraVenda = True
        SerieProd = ""
        wRecebeCarimboAnexo = ""

        If ConsistenciaNota(NumeroDocumento, SerieDocumento) = False Then
            EncerraVenda = False
            Exit Function
        End If
        
        If frmEncerraNFOutrasOperacoes.optCliente.Value = True Then

            sql = "Select capanfvenda.*, Estados.*, cliente.* from capanfvenda, Estados, cliente where capanfvenda.vc_numeropedido = " & _
                NroPedido & " and capanfvenda.vc_nomecliente = cliente.ce_razao " & _
                "And cliente.ce_estado = Estados.UF_Estado"
             
             
          Else
            sql = "Select capanfvenda.*, Estados.*, fornecedor.* from capanfvenda, Estados, fornecedor where capanfvenda.vc_numeropedido = " & _
                NroPedido & " and capanfvenda.vc_nomecliente = fornecedor.fo_razaosocial " & _
                "And fornecedor.fo_estado = Estados.UF_Estado"
          
             End If
             
       adoCapaNF.CursorLocation = adUseClient
      adoCapaNF.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
        If Not adoCapaNF.EOF Then
           If frmEncerraNFOutrasOperacoes.optCliente.Value = True Then
                If adoCapaNF("ce_Tipopessoa") = "F" Or adoCapaNF("ce_Tipopessoa") = "U" Then
                    wPessoa = 2
                Else
                    wPessoa = 1
                End If
            Else
                If adoCapaNF("fo_Tipofornecedor") = "F" Or adoCapaNF("fo_Tipofornecedor") = "D" Then
                    wPessoa = 2
                Else
                    wPessoa = 1
                End If
           End If

           
           
           wChaveICMS = adoCapaNF("UF_Regiao") & wPessoa
           If adoCapaNF("vc_Serie") <> "S1" And adoCapaNF("vc_Serie") <> "D1" Then
              wSerie = ""
           Else
              wSerie = IIf(IsNull(adoCapaNF("Serie")), "", adoCapaNF("Serie"))
           End If
        Else
            MsgBox "Nota não encontrada", vbInformation, "Atenção"
            Exit Function
        End If
                    
        sql = "Select * from produto,itemnfvenda " _
              & "where vi_notafiscal = " & NroNotaFiscal & "" _
              & " and pr_referencia = vi_referencia and vi_lojaorigem = '" & frmEncerraNFOutrasOperacoes.cmbLojaOrigem.Text & "' "
              
              adoItensNf.CursorLocation = adUseClient
              adoItensNf.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
           
          
          If Not adoItensNf.EOF Then
          

     wChaveICMSitem = wChaveICMS
     If adoItensNf("PR_IcmsSaida") = 0 And adoItensNf("PR_SubstituicaoTributaria") = "N" Then
        wST20 = "S"
     End If
     
     If adoItensNf("PR_SubstituicaoTributaria") = "S" Then
        wSubstituicaoTributaria = 1
        wST60 = "S"
     End If
     
     wChaveICMSitem = wChaveICMSitem & adoItensNf("PR_IcmsSaida") & adoItensNf("PR_CodigoReducaoIcms") & wSubstituicaoTributaria
     If AcharICMSInterEstadual(adoItensNf("PR_Referencia"), wChaveICMSitem) = False Then

        Exit Function
     Else

     End If
 
        GLB_ValorCalculadoICMS = Format((((adoItensNf("vi_valormercadoria") - adoItensNf("vi_Desconto")) _
                                 * GLB_AliquotaAplicadaICMS) / 100), "0.00")
                                 
        GLB_TotalICMSCalculado = (GLB_TotalICMSCalculado + GLB_ValorCalculadoICMS)
        
        If GLB_TotalICMSCalculado > 0 Then
          If RsICMSIntER("IE_BasedeReducao") = 0 Then
            If GLB_AliquotaAplicadaICMS = 0 Then
               GLB_BasedeCalculoICMS = 0
            Else
               GLB_BasedeCalculoICMS = (adoItensNf("vltotitem") - adoItensNf("Desconto"))
            End If
          Else
            GLB_BasedeCalculoICMS = Format((adoItensNf("vi_valormercadoria") - adoItensNf("vi_Desconto")) - _
                                    (((adoItensNf("vi_valormercadoria") - adoItensNf("vi_Desconto")) * _
                                    RsICMSIntER("IE_BasedeReducao")) / 100), "0.00")
          End If
            GLB_BaseTotalICMS = (GLB_BaseTotalICMS + GLB_BasedeCalculoICMS)
        End If
        
        
        If wTipoNota = "T" Then
           If adoItensNf("PR_substituicaotributaria") = "S" Then
              wCFO2 = 5409
           Else
              wCFO1 = 5152 & " "
           End If
        End If


            Do While Not adoItensNf.EOF

' -------------------------------------- CALCULO DA MARGEM DE VENDA ---------------------------------------------------
'
'               wSomaVenda = wSomaVenda + (adoItensNF("vltotitem") - adoItensNF("desconto"))
'               wSomaMargem = wSomaMargem + ((adoItensNF("vltotitem") - adoItensNF("desconto")) - (adoItensNF("pr_customedio1") * adoItensNF("qtde")))
 
 
'
' -------------------------------------- ATUALIZA ITENS DE VENDA --------------------------------------------------
'

                    wQuantItensCapaNF = adoCapaNF("vc_itens")
                    wQuantItensNF = adoItensNf("vi_numeroitem")
                    wQuantdadeTotalItem = wQuantdadeTotalItem + 1
                    If adoItensNf("vi_TipoNota") <> "E" Then
                        wQuant = (wQuantItensNF Mod 8)
                        If wQuant <> 0 Then
                            wDetalheImpressao = "D"
                        Else
                            wDetalheImpressao = "C"
                            If wQuantItensCapaNF > wQuantItensNF Then
                                wUltimoItem = wUltimoItem + 1
                            End If
                        End If
                        
                        If wQuantItensCapaNF = wQuantItensNF Then
                            wDetalheImpressao = "T"
                            wUltimoItem = wUltimoItem + 1
                        ElseIf wQuantItensCapaNF = wQuantdadeTotalItem Then
                            wDetalheImpressao = "T"
                            wUltimoItem = wUltimoItem + 1
                        End If
                    Else
                        
                        wQuant = (wQuantItensNF Mod 6)
                        If wQuant <> 0 Then
                            wDetalheImpressao = "D"
                        Else
                            wDetalheImpressao = "C"
                            wUltimoItem = wUltimoItem + 1
                        End If
                                        
                        If wQuantItensCapaNF = wQuantItensNF Then
                            wDetalheImpressao = "T"
                            wUltimoItem = wUltimoItem + 1
                        ElseIf wQuantItensCapaNF = wQuantdadeTotalItem Then
                            wDetalheImpressao = "T"
                            wUltimoItem = wUltimoItem + 1
                        End If
                    End If

                    
                    If wRomaneio = True Then
                       GLB_BasedeCalculoICMS = 0
                       GLB_ValorCalculadoICMS = 0
                    End If
                    
                    sql = "UPDATE itemnfvenda set vi_baseicms = " & ConverteVirgula(GLB_BasedeCalculoICMS) & ", " _
                    & "vi_Valoricms = " & ConverteVirgula(GLB_ValorCalculadoICMS) & "" _
                    & " where itemnfvenda.vi_notafiscal = " & NroNotaFiscal & " and vi_lojaorigem = '" & frmEncerraNFOutrasOperacoes.cmbLojaOrigem.Text & "'" _
                    & " and vi_Referencia = '" & adoItensNf("PR_Referencia") & "' and vi_numeroItem=" & adoItensNf("vi_numeroItem") & " and vi_serie = '" & wSerie & "'"
                    ADO_Cn_CDLocal.Execute (sql)
                
                
                adoItensNf.MoveNext
             Loop
        End If
        If wRomaneio = True Then
           wRomaneio = False
        End If
        
        adoItensNf.Close
'
' -------------------------------------- ATUALIZA CAPA DE VENDA --------------------------------------------------
 
             sql = "UPDATE capanfvenda set " _
                & "vc_ECF  = " & GLB_ECF & " " _
                & "where capanfvenda.vc_notafiscal = " & NroNotaFiscal & " and vc_lojaorigem = '" & LojaOrigem & _
                "' and vc_dataemissao = '" & Format(Date, "yyyy/mm/dd") & "'"
                ADO_Cn_CDLocal.Execute (sql)
 
 
 
   adoCapaNF.Close
            
    
End Function

Function PegaSerieNota() As String
    
    
    sql = ""
    sql = "Select CTS_SerieNota from ControleSistema"
    
    
    adoSerie.CursorLocation = adUseClient
    adoSerie.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic

    If Not adoSerie.EOF Then
        PegaSerieNota = adoSerie("CTS_SerieNota")
    End If
    adoSerie.Close

End Function

Public Function EmiteNotafiscal(ByVal nota As Double, ByVal serie As String)
 
wControlaQuebraDaPagina = 0
 
    For Each nomeImpressora In Printers
      
        If Trim(nomeImpressora.DeviceName) = UCase(wImpressoraNota) Then
         
            Set Printer = nomeImpressora
            Exit For
        End If
    Next
     wSerie = serie
    
    wNotaTransferencia = False
    wPagina = 1
    
    Call DadosLoja
            
    sql = ""
''    SQL = "Select NFCAPA.FreteCobr,NFCAPA.Carimbo5,NFCAPA.PedCli,NFCAPA.LojaVenda,NFCAPA.VendedorLojaVenda,NFCAPA.AV,NFCAPA.Carimbo3,NFCAPA.Carimbo2,NFCAPA.CFOAUX,NFCAPA.NF,NFCAPA.BASEICMS,NFCAPA.SERIE,NFCAPA.PAGINANF, " _
''        & "NFCAPA.CLIENTE,NFCAPA.FONECLI,NFCAPA.NUMEROPED,NFCAPA.VENDEDOR,NFCAPA.PGENTRA," _
''        & "NFCAPA.LOJAORIGEM,NFCAPA.DATAEMI,NFCAPA.SUBTOTAL,Nfcapa.nf,Nfcapa.Carimbo1,NfCapa.Desconto," _
''        & "NFCAPA.CODOPER,NFCAPA.TOTALNOTA,NFCAPA.VlrMercadoria,Nfcapa.cfoaux,Nfcapa.lojaOrigem,Nfcapa.Carimbo4," _
''        & "NFCAPA.ALIQICMS,NFCAPA.VLRICMS,NFCAPA.TIPONOTA,NFCAPA.NOMCLI,NFCAPA.CGCCLI,NFCAPA.CONDPAG, " _
''        & "NFCAPA.ENDCLI,NFCAPA.MUNICIPIOCLI,NFCAPA.BAIRROCLI,NFCAPA.CEPCLI,NFCAPA.INSCRICLI,NfCapa.CondPag,NfCapa.DataPag," _
''        & "NFCAPA.UFCLIENTE,NFCAPA.TOTALNOTAALTERNATIVA,NFCAPA.VALORTOTALCODIGOZERO,NFITENS.REFERENCIA,NFITENS.QTDE,NFITENS.VLUNIT," _
''        & "NFITENS.VLTOTITEM,NFITENS.ICMS,NfItens.TipoNota,NfCapa.EmiteDataSaida " _
''        & "From NFCAPA,NFITENS " _
''        & "Where NfCapa.nf= " & Nota & " and NfCapa.Serie in ('" & Serie & "') " _
''        & "and NfCapa.lojaorigem='" & Trim(wLoja) & "' " _
''        & "and NfItens.LojaOrigem=NfCapa.LojaOrigem " _
''        & "and NfItens.Serie=NfCapa.Serie " _
''        & "and NfItens.Nf=NfCapa.NF"
 
    sql = "Select NFITENS.cfop, NFITENS.detalheimpressao,NFCAPA.FreteCobr,NFCAPA.PedCli,NFCAPA.LojaVenda,NFCAPA.VendedorLojaVenda, " & _
          "NFCAPA.AV,NFCAPA.NF,NFCAPA.BASEICMS,NFCAPA.SERIE,NFCAPA.PAGINANF, " & _
          "NFCAPA.CLIENTE,cliente.CE_Telefone,NFCAPA.NUMEROPED,NFCAPA.VENDEDOR,NFCAPA.PGENTRA, " & _
          "NFCAPA.LOJAORIGEM,NFCAPA.DATAEMI,Nfcapa.nf,NfCapa.Desconto,NFCAPA.CODOPER,NFCAPA.TOTALNOTA, " & _
          "NFCAPA.VlrMercadoria,Nfcapa.lojaOrigem,NFCAPA.ALIQICMS,NFCAPA.VLRICMS,NFCAPA.TIPONOTA,cliente.ce_razao, " & _
          "cliente.ce_cgc,NFCAPA.CONDPAG,cliente.ce_Endereco,cliente.ce_municipio,cliente.ce_Bairro,cliente.CE_Cep," & _
          "cliente.ce_inscricaoEstadual,NfCapa.CondPag,NfCapa.DataPag,cliente.ce_Estado,NFCAPA.TOTALNOTAALTERNATIVA, " & _
          "NFCAPA.VALORTOTALCODIGOZERO, NfItens.Referencia , NfItens.QTDE, NfItens.VLUNIT, NfItens.VLTOTITEM, " & _
          "NfItens.ICMS, NfItens.TipoNota, NFCAPA.EmiteDataSaida " & _
          "From NFCAPA,NFITENS,cliente Where NfCapa.nf= " & nota & " and NfCapa.Serie in ('" & wSerie & "') and " & _
          "NfCapa.lojaorigem='" & Trim(wLoja) & "' and NfItens.LojaOrigem=NfCapa.LojaOrigem  and " & _
          "NfItens.Serie = NfCapa.Serie And NfItens.nf = NfCapa.nf And cliente.ce_CodigoCliente = NfCapa.Cliente"
        
    rsDados.CursorLocation = adUseClient
    rsDados.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not rsDados.EOF Then
           
      ' Cabecalho RsDados("ICMS")
      Cabecalho rsDados("tiponota")
    '*******************************************************************
      sql = "Select nfitens.cfop, nfitens.detalheimpressao,produto.pr_referencia,produto.pr_descricao, " _
          & "produto.pr_classefiscal,produto.pr_unidade, " _
          & "produto.pr_icmssaida,nfitens.referencia,nfitens.qtde,NfItens.TipoNota, " _
          & "nfitens.vlunit,nfitens.vltotitem,nfitens.icms,nfitens.icms,nfitens.detalheImpressao,nfitens.ReferenciaAlternativa, " _
          & "nfitens.PrecoUnitAlternativa,nfitens.DescricaoAlternativa " _
          & "from produto,nfitens " _
          & "where produto.pr_referencia=nfitens.referencia " _
          & "and nfitens.nf = " & nota & " and Serie='" & wSerie & "' order by nfitens.item"
 
      
      RsdadosItens.CursorLocation = adUseClient
      RsdadosItens.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
 
      If Not RsdadosItens.EOF Then
         wConta = 0
         Do While Not RsdadosItens.EOF
  '          wPegaDescricaoAlternativa = "0"
            wDescricao = ""
            wReferenciaEspecial = RsdadosItens("PR_Referencia")
                              
'*                   wPegaDescricaoAlternativa = IIf(IsNull(RsdadosItens("DescricaoAlternativa")), RsdadosItens("PR_Descricao"), RsdadosItens("DescricaoAlternativa"))
'*                   If wPegaDescricaoAlternativa = "" Then
'*                        wPegaDescricaoAlternativa = "0"
'*                   End If
'*                   If wPegaDescricaoAlternativa <> "0" Then
'*                        wDescricao = wPegaDescricaoAlternativa
'*                  Else
'*                        wDescricao = Trim(RsdadosItens("pr_descricao"))
'*                  End If
 
                   If RsdadosItens("ReferenciaAlternativa") = 0 Then
                       wDescricao = Trim(RsdadosItens("pr_descricao"))
                   Else
                       wDescricao = Trim(RsdadosItens("DescricaoAlternativa"))
                   End If
                   
                                   
'                 If RsDados("ce_Estado") = "SP" Then
                    wAliqICMSInterEstadual = RsdadosItens("icms")
'                 Else
'                    wAliqICMSInterEstadual = RsdadosItens("icmpdv")
'                 End If
                 
                 
                   
                   wStr16 = ""
'                   wStr16 = Left$(RsdadosItens("pr_referencia") & Space(7), 7) _
'                         & Space(1) & Left$(Format(Trim(wDescricao), ">") & Space(38), 38) _
'                         & Space(16) & Left$(Format(Trim(RsdadosItens("pr_classefiscal")), ">") _
'                         & Space(12), 12) & Left$(Trim(RsdadosItens("Tributacao")) & Space(3), 3) _
'                         & "" & Space(3) & Left$(Trim(RsdadosItens("pr_unidade")) & Space(2), 2) _
'                         & Right$(Space(6) & Format(RsdadosItens("QTDE"), "##0"), 6) _
'                         & Right$(Space(12) & Format(RsdadosItens("vlunit"), "#####0.00"), 14) _
'                         & Right$(Space(15) & Format(RsdadosItens("VlTotItem"), "#####0.00"), 15) & Space(1) _
'                         & Right$(Space(2) & Format(wAliqICMSInterEstadual, "#0"), 2)
 
                    
                    wStr16 = left$(wReferenciaEspecial & Space(7), 7) _
                         & Space(1) & left$(Format(Trim(wDescricao), ">") & Space(38), 38) _
                         & Space(16) & left$(Format(Trim(RsdadosItens("pr_classefiscal")), ">") _
                         & Space(12), 12) & left$(Trim(wIE_Tributacao) & Space(3), 3) _
                         & "" & Space(3) & left$(Trim(RsdadosItens("pr_unidade")) & Space(2), 2) _
                         & right$(Space(6) & Format(RsdadosItens("QTDE"), "##0"), 6) _
                         & right$(Space(12) & Format(RsdadosItens("vlunit"), "#####0.00"), 14) _
                         & right$(Space(15) & Format(RsdadosItens("VlTotItem"), "#####0.00"), 15) & Space(1) _
                         & right$(Space(2) & Format(wAliqICMSInterEstadual, "#0"), 2)
                            
                     
                      Printer.Print wStr16
                      
                      
                      If RsdadosItens("DetalheImpressao") = "D" Then
                         wConta = wConta + 1
                         RsdadosItens.MoveNext
                      ElseIf RsdadosItens("DetalheImpressao") = "C" Then
                         Do While wConta < 28
                            wConta = wConta + 1
                            Printer.Print ""
                         Loop
                         RsdadosItens.MoveNext
                         
                         wStr13 = Space(78) & "Lj " & rsDados("LojaOrigem") & Space(3) & right$(Space(7) & Format(rsDados("Nf"), "###,###"), 7)
                         Printer.Print wStr13
                         
                         wConta = 0
                         wPagina = wPagina + 1
                         
'------------------------------------------------------------------------------
                 'Acerto emissao de nota com mais de um formulario
                       ' Printer.EndDoc
                        
                         Printer.Print ""
                         Printer.Print ""
                         Printer.Print ""
                         Printer.Print ""
                         
                         wControlaQuebraDaPagina = wControlaQuebraDaPagina + 1
                         If wControlaQuebraDaPagina = 3 Then
                            Printer.Print ""
                            wControlaQuebraDaPagina = 0
                         End If
'----------------------------------------------------------------------------------
                        ' Cabecalho RsdadosItens("ICMS")
                       ' Cabecalho rsdados("ICMS")
                       Cabecalho rsDados("tiponota")
                      ElseIf RsdadosItens("DetalheImpressao") = "T" Then
                         wConta = wConta + 1
                         RsdadosItens.MoveNext
                         Call FinalizaNota
                      Else
                         wConta = wConta + 1
                         RsdadosItens.MoveNext
                      End If
            
            Loop
         Else
            
            MsgBox "Produto não encontrado", vbInformation, "Aviso"
         End If
        
      
    Else
        MsgBox "Nota Não Pode ser impressa", vbInformation, "Aviso"
    End If
 
 
RsdadosItens.Close
rsDados.Close

 
 
End Function

Function ConsistenciaNota(ByVal pedido As Double, ByVal serie As String) As Boolean
    
    
    sql = ""
    sql = "Select count(itemnfvenda.vi_Referencia) as QuantRef, capanfvenda.vc_itens from capanfvenda,itemnfvenda " _
        & "where capanfvenda.vc_notafiscal=" & NroNotaFiscal & " " _
        & "and itemnfvenda.vi_notafiscal=capanfvenda.vc_notafiscal " _
        & "and vc_dataemissao = '" & Format(Date, "yyyy/mm/dd") & "' and vc_dataemissao = vi_dataemissao " _
        & "group by capanfvenda.vc_itens " _
        & "having Count(itemnfvenda.vi_Referencia) = capanfvenda.vc_itens"
    
    adoItemNota.CursorLocation = adUseClient
    adoItemNota.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not adoItemNota.EOF Then
        ConsistenciaNota = True
    Else
        MsgBox "A nota não pode ser impressa porque exite um erro com a quantidade de itens ", vbCritical, "Atenção"
        ConsistenciaNota = False
    End If
    adoItemNota.Close
End Function

Function AcharICMSInterEstadual(ByVal referencia As String, ByVal ChaveIcms As Double) As Boolean
    
    wIE_icmsAplicado = 0
    wIE_Tributacao = 0
    wIE_Cfo = 0
    wIE_BasedeReducao = 0
    wIE_icmsdestino = 0
    
    sql = "SELECT * from IcmsInterEstadual where IE_Codigo = " & ChaveIcms
       
    RsICMSIntER.CursorLocation = adUseClient
    RsICMSIntER.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
           
    If RsICMSIntER.EOF Then
        AcharICMSInterEstadual = False
        'MsgBox "ICMS inter estadual da referencia " & referencia & " não encontrado" & Chr(10) & "A nota não pode ser impressa", vbCritical, "Aviso"
        RsICMSIntER.Close
        Exit Function
    Else
        AcharICMSInterEstadual = True
    End If
    
    
    wIE_icmsAplicado = RsICMSIntER("IE_icmsAplicado")
    wIE_Tributacao = RsICMSIntER("IE_CST")
    wIE_Cfo = RsICMSIntER("IE_Cfop")
    wIE_BasedeReducao = RsICMSIntER("IE_BasedeReducao")
    wIE_icmsdestino = RsICMSIntER("IE_icmsdestino")
    
    RsICMSIntER.Close
    
End Function

Function Cabecalho(ByVal TipoNota As String)
    Dim wCgcCliente As String
    Dim impri As Long
    
    impri = Printer.Orientation
       
    
    Printer.ScaleMode = vbMillimeters
    Printer.ForeColor = "0"
    Printer.FontSize = 8
    Printer.FontName = "draft 20cpi"
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.DrawWidth = 3
    
    
    Printer.FontName = "COURIER NEW"
    Printer.FontSize = 8#
    
    
            
    Wcondicao = "            "
    Wav = "          "
    If rsDados("CondPag") = 85 Then
        wCarimbo4 = Format(rsDados("DataPag"), "dd/mm/yyyy")
    Else
        sql = ""
        sql = "Select CP_condicao from CondicaoPagamento " _
            & "where CP_Codigo=" & rsDados("CondPag") & ""
             
             adoConPag.CursorLocation = adUseClient
             adoConPag.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
        
        If Not adoConPag.EOF Then
            wCarimbo4 = adoConPag("CP_condicao")
        End If
    End If

    wLojaVenda = "            "
    wVendedorLojaVenda = "            "
    wLojaVenda = IIf(IsNull(rsDados("LojaVenda")), rsDados("LojaOrigem"), rsDados("LojaVenda"))
    wVendedorLojaVenda = IIf(IsNull(rsDados("VendedorLojaVenda")), 0, rsDados("VendedorLojaVenda"))
    Wentrada = 0
    Wcondicao = "            "
    wStr20 = ""
    wStr19 = "               "
    wStr7 = "               "
    If Val(rsDados("CONDPAG")) = 1 Then
       Wcondicao = "Avista"
    ElseIf Val(rsDados("CONDPAG")) = 3 Then
       Wcondicao = "Financiada"
    ElseIf Val(rsDados("CONDPAG")) > 3 Then
       
       Wcondicao = wCarimbo4
    End If
    
    
  '  If chamou = "frmEncerraNFOutrasOperacoes" Then
     SQL2 = "Select * from cfopentradasaida, nfitens where lojaorigem = '" & GLB_Loja & _
        "' and cfo_codigo = cfop and nf = '" & NroNotaFiscal & "'"
        adotipo.CursorLocation = adUseClient
        adotipo.Open SQL2, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    Wnatureza = Trim(adotipo("cfo_codigo")) & " - " & Trim(adotipo("cfo_DescricaoOperacao"))
   
  '  Else
   
    ' If UCase(TipoNota) = "T" Then
    '     WNatureza = "TRANSFERENCIA"
    ' ElseIf UCase(TipoNota) = "V" Then
    '     WNatureza = "VENDA"
    ' ElseIf UCase(TipoNota) = "E" Then
   '     WNatureza = "DEVOLUCAO"
    ' ElseIf UCase(TipoNota) = "S" And (RsDados("CFOAUX") = "5949" Or RsDados("CFOAUX") = "6949") Then
    '     WNatureza = "OUTRAS OPER Ñ ESPEC."
    ' End If
   ' End If
    
    If Trim(wLojaVenda) > 0 Then
        If Trim(wLojaVenda) <> Trim(rsDados("LojaOrigem")) Then
            wStr6 = "VENDA OUTRA LOJA " & wLojaVenda & " " & wVendedorLojaVenda
        Else
            wStr6 = ""
        End If
    Else
        wStr6 = ""
    End If
    If Trim(rsDados("AV")) > 1 Then
        If Mid(Wcondicao, 1, 9) = "Faturada " Then
            Wav = "AV            : " & Trim(rsDados("AV"))
        End If
    End If
    
    If Trim(Wnatureza) = "TRANSFERENCIA" Then
        Wcondicao = "            "
    ElseIf Trim(Wnatureza) = "DEVOLUCAO" Then
        Wcondicao = "            "
    End If
    
   ' wStr17 = "Pedido        : " & rsDados("NUMEROPED")
   ' wStr18 = "Vendedor      : " & rsDados("VENDEDOR")
    If chamou = "frmencerranfoutrasoperacoes" Then
        If Trim(Wcondicao) <> "" Then
            wStr19 = "Cond Pagto : " & Trim(Wcondicao)
        ElseIf Trim(rsDados("Carimbo3")) <> "" Then
            wStr19 = "Transporte    : " & left(Format(Trim(rsDados("Carimbo3"))) & Space(10), 10)
        Else
            Wcondicao = "            "
        End If
    End If
    

    If rsDados("Pgentra") <> 0 Then
       Wentrada = Format(rsDados("Pgentra"), "#####0.00")
       wStr20 = "Entrada       : " & Format(Wentrada, "0.00")
    End If
    If (IIf(IsNull(rsDados("PedCli")), 0, rsDados("PedCli"))) <> 0 Then
        wStr7 = "Ped. Cliente    : " & Trim(rsDados("PedCli"))
    End If
    
    
    
   
    If wPagina = 1 Then
        WCGC = right(String(14, "0") & WCGC, 14)
        WCGC = Format(Mid(WCGC, 1, Len(WCGC) - 6), "###,###,###") & "/" & Mid(WCGC, Len(WCGC) - 5, Len(WCGC) - 10) & "-" & Mid(WCGC, 13, Len(WCGC))
        WCGC = right(String(18, "0") & WCGC, 18)
    End If
    wStr0 = Space(105) & wPagina & "/" & rsDados("PAGINANF") 'Inicio Impressão
    Printer.Print wStr0
    
    Printer.ScaleMode = vbMillimeters
    Printer.ForeColor = "0"
    Printer.FontSize = 6
    Printer.FontName = "draft 20cpi"
    Printer.FontSize = 6
    Printer.FontBold = False
    Printer.DrawWidth = 3
    Printer.FontName = "COURIER NEW"
    Printer.FontSize = 6#
    
    If wNovaRazao <> "0" Then
        wStr1 = Space(64) & wNovaRazao
        Printer.Print wStr1
        Printer.Print ""
    Else
        Printer.Print ""
    End If
    Printer.ScaleMode = vbMillimeters
    Printer.ForeColor = "0"
    Printer.FontSize = 8
    Printer.FontName = "draft 20cpi"
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.DrawWidth = 3
    Printer.FontName = "COURIER NEW"
    Printer.FontSize = 8#
    
    
   If Glb_NfDevolucao = True Then
       ' WNatureza = "DEVOLUCAO"
        wStr1 = Space(2) & left(Format(wStr17) & Space(34), 34) & left(Format(Trim(wendereco), ">") & Space(34), 34) & left(Format(Trim(wbairro), ">") & Space(11), 11) & Space(15) & "X" & Space(16) & left(Format(rsDados("nf"), "######"), 7)
    Else
        wStr1 = Space(2) & left(Format(wStr17) & Space(34), 34) & left(Format(Trim(wendereco), ">") & Space(34), 34) & left(Format(Trim(wbairro), ">") & Space(11), 11) & Space(5) & "X" & Space(26) & left(Format(rsDados("nf"), "######"), 7)
    End If
    Printer.Print wStr1
    wStr2 = Space(2) & left(Format(wStr18) & Space(34), 34) & left(Format(Trim(wMunicipio)) & Space(15), 15) & Space(24) & left$(Trim(westado), 2)
    Printer.Print wStr2
    If wSerie = "CT" Then
        wStr3 = Space(2) & left$(Format(wStr19) & Space(34), 34) & Space(29) & "(" & wDDDLoja & ")" & left$(Trim(Format(wFone, "###-####")), 9) & "/(" & wDDDLoja & ")" & left$(Format(WFax, "###-####"), 9) & Space(5) & left$(Format((WCep), "####-##'"), 9)
    Else
        wStr3 = Space(2) & left$(Format(wStr19) & Space(34), 34) & "(" & wDDDLoja & ")" & left$(Trim(Format(wFone, "###-####")), 9) & "/(" & wDDDLoja & ")" & left$(Format(WFax, "###-####"), 9) & Space(5) & left$(Format((WCep), "####-###"), 9)
    End If
    Printer.Print wStr3
    If wSerie = "CT" Then
        wStr4 = ""
    Else
        wStr4 = Space(2) & left(Format(wStr20) & Space(40), 40) & Space(46) & left(Trim(Format(WCGC, "###,###,###")), 19)
    End If
    Printer.Print wStr4
    Printer.Print ""
    
'*?*    If Wserie = "CT" Then
'        If Trim(WNatureza) = "TRANSFERENCIA" Then
            wStr5 = Space(36) & Format(Trim(Wnatureza), ">")
'        End If
'    Else
'
'        If Trim(Wav) <> "" Then
'            wStr5 = Space(2) & Left$(Wav & Space(32), 32) & Format(Trim(WNatureza), ">") & Space(27) & Left$(RsDados("CFOAUX"), 10) & Space(25) & Left$(Trim(Format((WIest), "###,###,###,###")), 15)
'        Else
'            wStr5 = Space(31) & Left(Trim(WNatureza) & Space(26), 26) & Left$(RsDados("CFOAUX"), 10) & Space(28) & Left$(Trim(Format((WIest), "###,###,###,###")), 15)
'        End If
'    End If
    Printer.Print wStr5
    
    Printer.Print ""
    Printer.Print ""
  
        wCgcCliente = right(String(14, "0") & Trim(rsDados("ce_cgc")), 14)
        wCgcCliente = Format(Mid(wCgcCliente, 1, Len(wCgcCliente) - 6), "###,###,###") & "/" & Mid(wCgcCliente, Len(wCgcCliente) - 5, Len(wCgcCliente) - 10) & "-" & Mid(wCgcCliente, 13, Len(wCgcCliente))
        wCgcCliente = right(String(18, "0") & Trim(wCgcCliente), 18)
   
    If wSerie = "CT" Then
        If wStr6 <> "" Then
            wStr6 = Space(2) & wStr6 & Space(8) & left$(Format(Trim(rsDados("CLIENTE"))) & Space(7), 7) & Space(1) & " - " & left$(Format(Trim(rsDados("ce_razao")), ">") & Space(50), 50) & Space(6) & left$(Format(rsDados("Dataemi"), "dd/mm/yyyy"), 12)
        Else
            wStr6 = Space(36) & left$(Format(Trim(rsDados("CLIENTE"))) & Space(7), 7) & Space(1) & " - " & left$(Format(Trim(rsDados("ce_razao")), ">") & Space(45), 45) & left$(Format(rsDados("Dataemi"), "dd/mm/yyyy"), 12)
        End If
    Else
        wStr6 = left(Trim(wStr6) & Space(31), 31) & left$(Format(Trim(rsDados("CLIENTE"))) & Space(7), 7) & Space(1) & " - " & left$(Format(Trim(rsDados("ce_razao")), ">") & Space(45), 45) & left$(Trim(wCgcCliente) & Space(24), 24) & Space(1) & left$(Format(rsDados("Dataemi"), "dd/mm/yy") & Space(12), 12)
    End If
    
    Printer.Print wStr6
    If rsDados("EmiteDataSaida") = "S" Then
        If wSerie = "CT" Then
            wStr7 = Space(2) & left(wStr7 & Space(29), 29) & left$(Format(Trim(rsDados("ce_endereco")), ">") & Space(42), 42) & Space(14) & left$(Format(rsDados("Dataemi"), "dd/mm/yyyy"), 12)
        Else
            wStr7 = Space(2) & left(wStr7 & Space(29), 29) & left$(Format(Trim(rsDados("ce_endereco")), ">") & Space(42), 42) & left$(Format(Trim(rsDados("ce_bairro")), ">") & Space(21), 21) & right$(Space(11) & rsDados("ce_cep"), 11) & Space(7) & left$(Format(rsDados("Dataemi"), "dd/mm/yy"), 12)
        End If
    Else
        If wSerie = "CT" Then
            wStr7 = Space(2) & left(wStr7 & Space(29), 29) & left$(Format(Trim(rsDados("ce_endereco")), ">") & Space(42), 42) '& Space(14) & Left$(Format(RsDados("Dataemi"), "dd/mm/yyyy"), 12)
        Else
            wStr7 = Space(2) & left(wStr7 & Space(29), 29) & left$(Format(Trim(rsDados("ce_endereco")), ">") & Space(42), 42) & left$(Format(Trim(rsDados("ce_bairro")), ">") & Space(21), 21) & right$(Space(11) & rsDados("ce_cep"), 11) '& Space(7) & Left$(Format(RsDados("Dataemi"), "dd/mm/yy"), 12)
        End If
    End If
    
    Printer.Print ""
    Printer.Print wStr7
    If wSerie = "CT" Then
        wStr8 = ""
    Else
        
        wStr8 = Space(31) & left$(Format(Trim(rsDados("ce_municipio")), ">") & Space(30), 30) & Space(4) & left$(Format(Trim(rsDados("ce_telefone"))) & Space(15), 15) & left$(Trim(rsDados("ce_estado")), 2) & Space(5) & left$(Trim(Format(rsDados("ce_inscricaoEstadual"), "###,###,###,###")), 15)
    End If
    Printer.Print ""
    Printer.Print wStr8
    
    Printer.Print ""
    Printer.Print ""


   adoConPag.Close
  adotipo.Close
End Function

Private Sub FinalizaNota()
        If wNotaTransferencia = False Then

'*******************************************************************************************************************
               Do While wConta < 7
                   wConta = wConta + 1
                   Printer.Print ""
                Loop
                          
                sql = ""
                sql = "Select * from carimbosEspeciais,CarimboNotaFiscal where CE_referencia = CNF_Carimbo And " & _
                      "CNF_TipoCarimbo = 'S' And CNF_NumeroPed = " & NroPedido
 
                RsPegaItensEspeciais.CursorLocation = adUseClient
                RsPegaItensEspeciais.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
                
                If Not RsPegaItensEspeciais.EOF Then
                    Do While Not RsPegaItensEspeciais.EOF
                    
                      If RsPegaItensEspeciais("CE_Linha12") <> "" Then
                         Printer.Print Space(8) & right(RsPegaItensEspeciais("CE_Linha12"), 90)
                      End If
                      
                      If RsPegaItensEspeciais("CE_Linha1") <> "" Then
                        wConta = wConta + 7
                        If Trim(RsPegaItensEspeciais("CE_Linha5")) = "" Then
                             Printer.Print Space(7) & "______________________________________________________________"
                             Printer.Print Space(8) & right(RsPegaItensEspeciais("CE_Linha2"), 60)
                             Printer.Print Space(8) & right(RsPegaItensEspeciais("CE_Linha3"), 60)
                             Printer.Print Space(8) & right(RsPegaItensEspeciais("CE_Linha4"), 60)
                             Printer.Print Space(9) & "___________________________________     ____/____/______   "
                             Printer.Print Space(9) & "            Assinatura                        Data         "
                        Else
                             Printer.Print Space(7) & "______________________________________________________________"
                             Printer.Print Space(8) & right(RsPegaItensEspeciais("CE_Linha2"), 60)
                             Printer.Print Space(8) & right(RsPegaItensEspeciais("CE_Linha3"), 60)
                             Printer.Print Space(8) & right(RsPegaItensEspeciais("CE_Linha4"), 60)
                             Printer.Print Space(8) & right(RsPegaItensEspeciais("CE_Linha5"), 60)
                             Printer.Print Space(9) & "___________________________________     ____/____/______   "
                             Printer.Print Space(9) & "            Assinatura                        Data         "
                        End If
                      End If
                      RsPegaItensEspeciais.MoveNext
                    Loop
                End If
                RsPegaItensEspeciais.Close
                
               
               sql = ""
             '  SQL = "Select * from CarimboNotaFiscal where " & _
                     "CNF_NumeroPed = " & NroPedido & " and " & _
                     "CNF_TipoCarimbo = 'I'"
                     
                If chamou = "reemissao" Then
                    sql = "Select * from nfcapa, CarimboNotaFiscal where " & _
                     "CNF_NumeroPed = numeroped and nf = '" & NroNotaFiscal & "'"
                     
                    RsPegaItensEspeciais.CursorLocation = adUseClient
                    RsPegaItensEspeciais.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
                Else
                
                     sql = "Select * from CarimboNotaFiscal where " & _
                     "CNF_NumeroPed = " & NroPedido & ""

                RsPegaItensEspeciais.CursorLocation = adUseClient
                RsPegaItensEspeciais.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
                End If
                
             If chamou = "frmEncerraNFOutrasOperacoes" Or chamou = "reemissao" Then
                If Not RsPegaItensEspeciais.EOF Then
                    Do While Not RsPegaItensEspeciais.EOF
                      
                        wStrI = Space(8) & Trim((RsPegaItensEspeciais("CNF_Carimbo")))
                       Printer.Print wStrI
                       RsPegaItensEspeciais.MoveNext
                    Loop
                End If
                RsPegaItensEspeciais.Close
            
            Else
                j = 1
                If Not RsPegaItensEspeciais.EOF Then
                    
                    Do While Not RsPegaItensEspeciais.EOF
                        wStrI = Space(8) & Trim((RsPegaItensEspeciais("CNF_Carimbo")))
                        If j = 2 Then
                            Printer.Print "Por motivos de " & wStrI
                        Else
                            Printer.Print wStrI
                            
                        End If
                        RsPegaItensEspeciais.MoveNext
                    j = j + 1
                    Loop
                End If
                RsPegaItensEspeciais.Close
            End If
            Do While wConta < 14
        wConta = wConta + 1
        Printer.Print ""
     Loop

End If
'********************************************************
     
 '   End If
    
     wStr9 = right$(Space(9) & Format(rsDados("BaseICMS"), "######0.00"), 9) & right$(Space(9) & Format(rsDados("VLRICMS"), "######0.00"), 9) & Space(35) & right$(Space(10) & Format(rsDados("VlrMercadoria"), "######0.00"), 10)
     Printer.Print wStr9
     Printer.Print ""
     wStr10 = right(Space(9) & Format(Space(9) & rsDados("FreteCobr"), "######0.00"), 9) & Space(44) & right(Space(10) & Format(rsDados("TotalNota"), "######0.00"), 10)
     Printer.Print wStr10
         
     wStr11 = Space(2) & "                          "
     Printer.Print wStr11
     wStr12 = Space(2) & "                                                     "
     Printer.Print wStr12
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     Printer.Print ""
     wStr13 = Space(78) & "Lj " & rsDados("LojaOrigem") & Space(4) & right$(Space(7) & Format(rsDados("Nf"), "###,###"), 7)
     Printer.Print wStr13
     Printer.Print ""
     Printer.Print ""
     
     Printer.EndDoc
    
       
End Sub


Sub TrocaBannerTopo1()
    
    If NroBanner = 1 Then
       'frmControleCD.Image1.Picture = LoadPicture("C:\sistemas\CD\Imagens\t3.JPG")
       NroBanner = 2
    ElseIf NroBanner = 2 Then
       'frmControleCD.Image1.Picture = LoadPicture("C:\sistemas\CD\Imagens\fundoSistema.JPG")
       NroBanner = 1
   ' ElseIf NroBanner = 3 Then
   '    frmPedido.webInternet2.Picture = LoadPicture("C:\sistemas\DMAC Balcao\Imagens\BannerTopo1\BannerTopo1d.swf")
   '    NroBanner = 4
   ' ElseIf NroBanner = 4 Then
   '    frmPedido.webInternet2.Picture = LoadPicture("C:\sistemas\DMAC Balcao\Imagens\BannerTopo1\BannerTopo1a.swf")
   '    NroBanner = 1
    End If
 
    EsperarTrocaBanner 10

  
 
End Sub

Sub EsperarTrocaBanner(ByVal Tempo As Long)
    
    Dim StartTime As Long
    StartTime = Timer
    Do While Timer < StartTime + Tempo
        DoEvents
    Loop
 
End Sub

Function ConvertePonto(ByVal Expressao) As String
    Dim ContPad As String
    Dim flgpad As Integer
    
    If Len(Expressao) <> 0 Then
        ContPad = CStr(Expressao)
        flgpad = InStr(ContPad, ".")
        Do While flgpad <> 0
            Mid(ContPad, flgpad, 1) = ","
            flgpad = InStr(ContPad, ".")
        Loop
    Else
        ContPad = 0
    End If
    ConvertePonto = ContPad
End Function

Public Function PegaNumPedido()
    
    Screen.MousePointer = 11
    
    sql = "Select * from ControleSistema "

    adoControle.CursorLocation = adUseClient
    adoControle.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    NroPedido = adoControle("CTS_NumeroPedido")
    PegaNumPedido = NroPedido
    adoControle.Close
    
   
       
    sql = "Update ControleSistema set CTS_NumeroPedido = " & NroPedido & " + 1"
    adoControle.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
  
    
Screen.MousePointer = vbNormal

End Function

Function PegaCliente()

    Dim nroCli As String
    
    Screen.MousePointer = 11
    
    sql = "select * from ControleSistema "
    adoControle.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    nroCli = adoControle("CTS_sequenciaCliente")
    PegaCliente = nroCli
    adoControle.Close
    
    sql = "Update ControleSistema set CTS_sequenciaCliente = " & NroPedido & " + 1"
    adoControle.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
End Function


Public Sub carimboImposto(NotaFiscal As String, serie As String, LojaOrigem As String)

    Dim resCarimboImpostos As New ADODB.Recordset
    Dim sql As String

    'FELIPE
    sql = "select NUMEROPED, CSTICMS from nfitens where nf = " & NotaFiscal & " and serie = '" & serie & "' and lojaorigem = '" & LojaOrigem & "' group by NUMEROPED,CSTICMS"
    resCarimboImpostos.CursorLocation = adUseClient
    resCarimboImpostos.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        Do While Not resCarimboImpostos.EOF
            If Val(resCarimboImpostos("CSTICMS")) = 60 Then
                sql = "Insert into CarimboNotaFiscal(CNF_NumeroPed,CNF_Loja,CNF_serie,CNF_NF,CNF_Sequencia,CNF_Carimbo,CNF_TipoCarimbo,CNF_SituacaoProcesso)" _
                  & "Values ( " & resCarimboImpostos("NUMEROPED") & ",'" & LojaOrigem & "','" _
                  & serie & "'," & "0" & "," & 1 & ",'" & "ST 060 " & Chr(34) & "IMPOSTO RECOLHIDO POR SUBSTITUICAO - ART 313-Z3,Z11,Z17 e Z19 DO RICMS" & Chr(34) & "','S','A')"
                ADO_Cn_CDLocal.Execute (sql)
            ElseIf Val(resCarimboImpostos("CSTICMS")) = 20 Then
                sql = "Insert into CarimboNotaFiscal(CNF_NumeroPed,CNF_Loja,CNF_serie,CNF_NF,CNF_Sequencia,CNF_Carimbo,CNF_TipoCarimbo,CNF_SituacaoProcesso)" _
                  & "Values ( " & resCarimboImpostos("NUMEROPED") & ",'" & LojaOrigem & "','" _
                  & serie & "'," & "0" & "," & 2 & ",'" & "ST 020 BASE CALC REDUZ ART.51 ANEXOS I.II ART. 12 I.II E IV DECR.45.490" & "','S','A')"
                ADO_Cn_CDLocal.Execute (sql)
            End If
            resCarimboImpostos.MoveNext
        Loop
    resCarimboImpostos.Close
    '''''

End Sub

