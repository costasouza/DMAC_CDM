Attribute VB_Name = "ModImpressao"
Option Explicit

Global nomeImpressora As Printer

Global lojaOrigemImpressao As String
Global serieImpressao As String

Public Sub impressoraPadraoSistema(NOME As String)
    Dim wImpressora As String
    wImpressora = "PDF2"
    For Each nomeImpressora In Printers
    If UCase(nomeImpressora.DeviceName) = UCase(NOME) Then
       Set Printer = nomeImpressora
    Exit For
    End If
    Next
End Sub

Public Function ImprimirNotaFiscal(ByVal NotaFiscal As Long) As Boolean
   
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
Dim rdorsExtra1 As New ADODB.Recordset
Dim rdorsExtra2 As New ADODB.Recordset
Dim rdorsExtra3 As New ADODB.Recordset
Dim rdorsExtra6 As New ADODB.Recordset
Dim rdorsNatOper As New ADODB.Recordset

Dim wWhere As String

ImprimirNotaFiscal = False

 sql = "Select tr_nome,tr_endereco,tr_bairro,tr_estado from transportadora"
            
    rdoTransportadora.CursorLocation = adUseClient
    rdoTransportadora.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
    If Not rdoTransportadora.EOF Then
       wnome = Trim(rdoTransportadora("tr_nome"))
       wendereco = Trim(rdoTransportadora("tr_endereco"))
       wbairro = Trim(rdoTransportadora("tr_bairro"))
       westado = Trim(rdoTransportadora("tr_estado"))
    End If
    rdoTransportadora.Close

  
  wSerieImpressao = serieImpressao
  'LojaOrigem '= lojaOrigemImpressao
   
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
        & " vc_serie = '" & wSerieImpressao & "'  and " _
        & " vc_lojaorigem = '" & LojaOrigem & "'" _
        & " order by PR_CodigoFornecedor, PR_Descricao"
 

        
       'Set rdorsExtra1 = rdoCnSupBatch.OpenResultset(sql, Options:=rdExecDirect)
            rdorsExtra1.CursorLocation = adUseClient
            rdorsExtra1.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
       
       'Set rdorsExtra3 = rdoCnSupBatch.OpenResultset(sql, Options:=rdExecDirect)
            rdorsExtra3.CursorLocation = adUseClient
            rdorsExtra3.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
       
       
       'Serie = rdorsExtra1("vc_serie")
       serie = wSerieImpressao
       'SerieParaRomaneio = rdorsExtra1("vc_serie")
 
       
       If rdorsExtra1.EOF Then
          'MsgBox "Nota Fiscal " & NotaFiscal & " não Encontrada", vbCritical, "Erro"
          Exit Function
       End If
       
       
        wCFO1 = " "
        wCFO2 = " "
        wCFO3 = " "
       
       If Not rdorsExtra3.EOF Then
          Do While Not rdorsExtra3.EOF
         
             
             If Trim(rdorsExtra3("Vc_CodigoOperacaoNovo")) = "5152" Then
                If rdorsExtra3("PR_substituicaotributaria") = "S" Then
                   wCFO2 = "5409"
                Else
                   wCFO1 = "5152"
                End If
                WcodigoOperacao = Trim(wCFO1) & wCFO3 & Trim(wCFO2)
             Else
                WcodigoOperacao = Trim(rdorsExtra3("Vc_CodigoOperacaoNovo"))
             End If
                
          rdorsExtra3.MoveNext
          Loop
 
          
          sql = "Update capanfvenda " _
                  & "Set Vc_CodigoOperacaoNovo = '" & Trim(WcodigoOperacao) & "'" _
                  & " where  vc_notafiscal= " & NotaFiscal & " and vc_serie='" & serie & "' and vc_lojaorigem = '" & LojaOrigem & "'"
                  ADO_Cn_CDLocal.Execute (sql)
           
        End If
    
       
    
        
        wConta = 0
        wChave = 0
        wReduz = 0
        wAliquotaZero = False
        wStr15 = ""
        wStr16 = ""
        wRes = 0
        Wnatureza = ""
        
        
        
        Do While Not rdorsExtra1.EOF
           flg = flg + 1 '???????????????????????
          
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
                  & " where lo_loja = '" & rdorsExtra1("vc_lojaorigem") & "' "
               
              'Set rdorsExtra2 = rdoCnSupBatch.OpenResultset(sql, Options:=rdExecDirect)
              rdorsExtra2.CursorLocation = adUseClient
              rdorsExtra2.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
              
              If Not rdorsExtra2.EOF Then
                 
               
                 If wTelaOperacaoEspecial = True Then
                    wStr17 = ""
                 Else
                    wStr17 = rdorsExtra1("vc_enderecocliente")
                 End If
                 
                 wStr1 = Space(2) & left(Format(wStr17) & Space(34), 34) & left(Format(Trim(rdorsExtra2("lo_endereco")), ">") & Space(34), 34) & left(Format(Trim(rdorsExtra2("lo_bairro")), ">") & Space(11), 11) & Space(5) & "X" & Space(26) & left(Format(rdorsExtra1("vc_notafiscal"), "######"), 7)
                 
                 If wTelaOperacaoEspecial = True Then
                    wStr18 = ""
                 Else
                    wStr18 = IIf(IsNull(rdorsExtra1("VC_Observacao")), "", rdorsExtra1("VC_Observacao"))
                 End If
                 wStr2 = Space(2) & left(Format(wStr18) & Space(34), 34) & left(Format(Trim(rdorsExtra2("lo_municipio"))) & Space(15), 15) & Space(24) & left$(Trim(rdorsExtra2("lo_uf")), 2)
                 wStr3 = Space(2) & left$(Format(wStr19) & Space(34), 34) & "(011)" & left$(Trim(Format(rdorsExtra2("lo_telefone"), "####-####")), 9) & "/(011)" & left$(Format(rdorsExtra2("lo_fax"), "####-####"), 9) & Space(5) & left$(Format(rdorsExtra2("lo_cep"), "00###-###"), 9)
                 wStr20 = ""
                 wStr4 = Space(2) & left(Format(wStr20) & Space(40), 40) & Space(46) & left$(Trim(Format(rdorsExtra2("lo_cgc"), "###,###,###")), 10) & "/" & right$(Format(rdorsExtra2("lo_cgc"), "####-##"), 7)
                 wEspaco = ""
                 wLO_inscricaoestadual = rdorsExtra2("lo_inscricaoestadual")
                 
           '    Set rdoCombos = rdoCnSup.OpenResultset("Select NO_CodigoOperacao, NO_CodigoNatureza, NO_Descricao, NO_TipoNatureza from NaturezaOperacao order by NO_CodigoOperacao", Options:=rdExecDirect)
      
                 
                 If wTelaOperacaoEspecial = True Then
                    sql = "Select * from codigooperacaonovo " _
                        & "where CN_CodigoOperacaoNovo = " & Trim(rdorsExtra1("Vc_CodigoOperacaoNovo")) & ""
                          
                    'Set rdorsNatOper = rdoCnSupBatch.OpenResultset(sql, Options:=rdExecDirect)
                    rdorsNatOper.CursorLocation = adUseClient
                    rdorsNatOper.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
                    
                    
                    If rdorsNatOper.EOF = False Then
                       wDescricao = rdorsNatOper("CN_DescricaoOperacao")
                    Else
                       wDescricao = ""
                    End If
                    wStr5 = Space(2) & left$(wEspaco & Space(31), 31) & Format(Trim(wDescricao), ">") & Space(1) & left$(WcodigoOperacao, 10) & Space(31) & left$(Trim(Format((rdorsExtra2("lo_inscricaoestadual")), "###,###,###,###")), 15)
                 Else
                    wDescricao = "TRANSFERENCIA"
                    wStr5 = Space(2) & left$(wEspaco & Space(32), 32) & Format(Trim(wDescricao), ">") & Space(9) & left$(WcodigoOperacao, 10) & Space(31) & left$(Trim(Format((rdorsExtra2("lo_inscricaoestadual")), "###,###,###,###")), 15)
                 End If
                
                ' wStr5 = Space(2) & left$(wEspaco & Space(32), 32) & Format(Trim(wDescricao), ">") & Space(9) & left$(WcodigoOperacao, 10) & Space(31) & left$(Trim(Format((rdorsExtra2("lo_inscricaoestadual")), "###,###,###,###")), 15)
              End If
            
            If wTelaOperacaoEspecial = True Then
             
               wStr6 = left(Trim(wEspaco) & Space(31), 31) & left$(Format(Trim(wEspaco)) & Space(0), 0) & left$(Format(Trim(rdorsExtra1("vc_nomecliente")), ">") & Space(56), 56) & left$(Trim(Format(rdorsExtra1("vc_cgccliente"), "###,###,###")), 10) & "/" & right$(Format(rdorsExtra1("vc_cgccliente"), "####-##"), 7) & Space(7) & left$(Format(rdorsExtra1("vc_dataemissao"), "dd/mm/yy") & Space(12), 12)
               wStr7 = "               "
               wStr7 = Space(2) & left(wStr7 & Space(29), 29) & left$(Format(Trim(rdorsExtra1("vc_enderecocliente")), ">") & Space(42), 42) & left$(Format(Trim(rdorsExtra1("vc_bairrocliente")), ">") & Space(21), 21) & right$(Space(11) & Format(rdorsExtra1("vc_cepcliente"), "00###-###"), 11) & Space(7) & left$(Format(rdorsExtra1("vc_dataemissao"), "dd/mm/yy"), 12)
               wEspaco = ""
               wStr8 = Space(31) & left$(Format(Trim(rdorsExtra1("vc_municipiocliente")), ">") & Space(15), 15) & Space(19) & left$(Format(Trim(wEspaco)) & Space(15), 15) & left$(Trim(rdorsExtra1("vc_ufcliente")), 2) & Space(5) & left$(Trim(Format(rdorsExtra1("vc_InscEstCliente"), "###,###,###,###")), 15)
             
            Else
                 sql = "select em_descricao,lo_endereco,lo_bairro," _
                     & "lo_municipio,lo_uf,lo_cep,lo_cgc," _
                     & "lo_inscricaoestadual,lo_fax,lo_telefone " _
                     & " from loja, empresa " _
                     & " where lo_empresa=em_codigoempresa " _
                     & " and lo_loja = '" & rdorsExtra1("vc_lojadestino") & "' "
                 'Set rdorsExtra2 = rdoCnSupBatch.OpenResultset(sql, Options:=rdExecDirect)
                 rdorsExtra2.Close
                 rdorsExtra2.CursorLocation = adUseClient
                 rdorsExtra2.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
                    
                 If Not rdorsExtra2.EOF Then
                    
                    wStr6 = left(Trim(wEspaco) & Space(31), 31) & left$(Format(Trim(wEspaco)) & Space(0), 0) & left$(Format(Trim(rdorsExtra2("em_descricao")), ">") & Space(56), 56) & left$(Trim(Format(rdorsExtra2("lo_cgc"), "###,###,###")), 10) & "/" & right$(Format(rdorsExtra2("lo_cgc"), "####-##"), 7) & Space(7) & left$(Format(rdorsExtra1("vc_dataemissao"), "dd/mm/yy") & Space(12), 12)
                    wStr7 = "               "
                    wStr7 = Space(2) & left(wStr7 & Space(29), 29) & left$(Format(Trim(rdorsExtra2("lo_endereco")), ">") & Space(42), 42) & left$(Format(Trim(rdorsExtra2("lo_bairro")), ">") & Space(21), 21) & right$(Space(11) & Format(rdorsExtra2("lo_cep"), "00###-###"), 11) & Space(7) & left$(Format(rdorsExtra1("vc_dataemissao"), "dd/mm/yy"), 12)
                    wEspaco = ""
                    wStr8 = Space(31) & left$(Format(Trim(rdorsExtra2("lo_municipio")), ">") & Space(15), 15) & Space(19) & left$(Format(Trim(wEspaco)) & Space(15), 15) & left$(Trim(rdorsExtra2("lo_uf")), 2) & Space(5) & left$(Trim(Format(rdorsExtra2("lo_inscricaoestadual"), "###,###,###,###")), 15)
                 End If
            End If
               
                 
                 rdorsExtra2.Close
                            
              wStr9 = right$(Space(9) & Format(rdorsExtra1("vc_baseicms"), "######0.00"), 9) & right$(Space(9) & Format(rdorsExtra1("vc_valoricms"), "######0.00"), 9) & Space(35) & right$(Space(10) & Format(rdorsExtra1("vc_valormercadorias"), "######0.00"), 10)
              wStr10 = right(Space(9) & Format(Space(9) & wEspaco, "######0.00"), 9) & Space(44) & right(Space(10) & Format(rdorsExtra1("vc_totalnota"), "######0.00"), 10)
              
              If wTelaOperacaoEspecial = True Then
                 wStr11 = ""
                 wStr12 = ""
              Else
                 wStr11 = Space(0) & wnome
                 wStr12 = Space(0) & wendereco & "              " & wbairro & "            " & westado
              End If
              
            ' wStr13 = Space(74) & "LOJA  CD "  & Space(11) & right$(Space(7) & Format(rdorsExtra1("vc_notafiscal"), "###,###"), 7)
            ' wStr13 = Space(74) & "LOJA MCE85" & Space(10) & right$(Space(7) & Format(rdorsExtra1("vc_notafiscal"), "###,###"), 7)
              wStr13 = Space(74) & "LOJA " & wLojaMCE85ouCD & Space(10) & right$(Space(7) & Format(rdorsExtra1("vc_notafiscal"), "###,###"), 7)
            
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
                  & "where pr_referencia = '" & rdorsExtra1("vi_referencia") & "' "
            
              'Set rdorsExtra2 = rdoCnSupBatch.OpenResultset(sql, Options:=rdExecDirect)
              rdorsExtra2.CursorLocation = adUseClient
              rdorsExtra2.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
            
              wCodIPI = 0
              If rdorsExtra2("pr_codigoipi") = 4 Then
                 wCodIPI = 1
              End If
              If rdorsExtra2("pr_codigoipi") = 5 Then
                 wCodIPI = 2
              End If
              If rdorsExtra2("pr_codigoreducaoicms") <> 0 Then
                 wCodTri = 2
              Else
                 wCodTri = 0
              End If
              If rdorsExtra2("pr_codigoreducaoicms") <> 0 Then
                 wReduz = 1
                 If wStr15 = "" Then
                     wStr15 = wStr15 & wConta
                  Else
                     wStr15 = wStr15 & "," & wConta
                  End If
              End If
              
                            
              If rdorsExtra1("vi_reserva") <> "" And rdorsExtra1("vi_reserva") <> 0 Then
                 wRes = 1
                 If wStr16 = "" Then
                    wStr16 = wStr16 & wConta
                 Else
                    wStr16 = wStr16 & "," & wConta
                 End If
              End If
              
              wEspaco = wCodIPI
              wEspaco = wEspaco & wCodTri
              
              If rdorsExtra1("vi_aliquotaicms") = "0" Then
                 wAliquotaZero = True
              End If

              
                  wStr1 = ""
                  wStr1 = left$(rdorsExtra1("vi_referencia") & Space(7), 7) _
                         & Space(1) & left$(Format(Trim(rdorsExtra2("pr_descricao")), ">") & Space(38), 38) _
                         & Space(16) & left$(Format(Trim(rdorsExtra2("pr_classefiscal")), ">") _
                         & Space(12), 12) & left$(Trim(wEspaco) & Space(3), 3) _
                         & "" & Space(3) & left$(Trim(rdorsExtra2("pr_unidade")) & Space(2), 2) _
                         & right$(Space(6) & Format(rdorsExtra1("vi_quantidade"), "##0"), 6) _
                         & right$(Space(12) & Format(rdorsExtra1("vi_precounitario"), "#####0.00"), 14) _
                         & right$(Space(15) & Format(rdorsExtra1("vi_valormercadoria"), "#####0.00"), 15) & Space(1) _
                         & right$(Space(2) & Format(rdorsExtra1("vi_aliquotaicms"), "#0"), 2)

              
              Printer.Print wStr1
             ' wConta = wConta + 1
              wStr1 = ""
              wChave = 1
              rdorsExtra2.Close
              rdorsExtra1.MoveNext
        Loop
            
           Do While wConta < 10
              wConta = wConta + 1
              Printer.Print
           Loop
     
     
  If wTelaOperacaoEspecial = True Then

           
     sql = "Select on_carimbo1,on_carimbo2 " _
          & "from observacaonotafiscal " _
          & "where ON_NumeroNotaFiscal= " & UCase(NotaFiscal) & " " _
          & "and ON_Serie in ('" & wSerieImpressao & "') " _
          & "and ON_Loja='" & LojaOrigem & "'"
        
          
          'Set rdorsExtra6 = rdoCnSupBatch.OpenResultset(sql, Options:=rdExecDirect)
          rdorsExtra6.CursorLocation = adUseClient
          rdorsExtra6.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
         
          If Not rdorsExtra6.EOF Then
             Printer.Print Trim(rdorsExtra6("on_carimbo1"))
             Printer.Print Trim(rdorsExtra6("on_carimbo2"))
             Printer.Print
          Else
             Printer.Print
             Printer.Print
             Printer.Print
          End If
   
  Else
     If wReduz = 1 Then
        Printer.Print Space(4) & "   BASE CALC REDUZ CONF.ART.51. ANEXOS I E II ART. 12 - I,II,III E IV DECR.45.490 --> ITENS " & wStr15
        'Printer.Print Space(4) & "   ITENS " & wStr15
     Else
        Printer.Print
      ' Printer.Print
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
           'Printer.Print
          ' Printer.Print
           
           'Printer.EndDoc
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
           rdorsExtra1.Close
           wTelaOperacaoEspecial = False
           
ImprimirNotaFiscal = True

End Function




Public Sub ImprimeTransferencia00(ByRef nf As String, ByVal serie As String, ByRef LojaOrigem As String)

    Dim ValorlItem As Double
    Dim ValorDesconto As Double
    Dim SubTotal As Double
    Dim TotalVenda As Double
    Dim nomeImpressora As Printer
    
    Dim RsDadosCapa As New ADODB.Recordset
    Dim rsDados As New ADODB.Recordset
    Dim rdoProduto As New ADODB.Recordset
    
    'Open GLB_Impressora00 For Output As #1
    
  'For Each NomeImpressora In Printers
        'If Trim(UCase(NomeImpressora.DeviceName)) = Trim(UCase("CDM")) Then
 '        If Trim(NomeImpressora.Orientation) = Trim(GLB_ImpressoraNota) Then

 '          ' Seta impressora no sistema
            'Set Printer = NomeImpressora
            'Exit For
        'End If
    'Next
    
'    Set Printer = "CDM"
    Printer.ScaleMode = vbMillimeters
    Printer.FontName = "Romam"
    Printer.FontSize = 9
 
    Screen.MousePointer = 11
   
    ValorlItem = 0
    ValorDesconto = 0
    SubTotal = 0


    sql = "Select * from Loja Where LO_Loja= '" & LojaOrigem & "'"

    rsDados.CursorLocation = adUseClient
    rsDados.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
   Printer.Print Tab(10); rsDados("LO_Razao")
   Printer.Print ; "CNPJ: " & rsDados("LO_CGC") & " I.E.: " & rsDados("LO_InscricaoEstadual")
   Printer.Print ; UCase(rsDados("LO_Endereco")) & ", " & rsDados("LO_numero")
   Printer.Print ; "TELEFONE: "; rsDados("LO_Telefone")
   Printer.Print ; Format(Date, "dd/mm/yyyy") & " " & Format(time, "HH:MM:SS") & Space(16) & "Nota Fiscal: "; Format(nf, "000000")
   Printer.Print "======================================"
   
   rsDados.Close
   
    sql = "Select Lojaorigem,lojat as cliente,totalnota" & vbNewLine & _
          "from nfcapa " & vbNewLine & _
          "Where nf = " & nf & vbNewLine & _
          "and serie = '" & serie & "'" & vbNewLine & _
          "and lojaorigem = '" & LojaOrigem & "'"
                 
    RsDadosCapa.CursorLocation = adUseClient
    RsDadosCapa.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
   
   Printer.Print "TRANSFERENCIA PARA LOJA " & Trim(RsDadosCapa("cliente"))
   Printer.Print "________________________________________"
   Printer.Print ""
   Printer.Print "DESCRICAO DO PRODUTO                    "
   Printer.Print "CODIGO  PRODUTO  QTDxUNIT.   VALOR TOTAL"
   Printer.Print "________________________________________"
'   rsDados.Close

    
   
    sql = "Select *" & vbNewLine & _
          "from NFITENS " & vbNewLine & _
          "Where nf = " & nf & vbNewLine & _
          "and serie = '" & serie & "'" & vbNewLine & _
          "and lojaorigem = '" & LojaOrigem & "'"
       
      
       rsDados.CursorLocation = adUseClient
       rsDados.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic

       
       If Not rsDados.EOF Then
          Do While Not rsDados.EOF
             sql = "Select PR_Descricao from Produto Where PR_Referencia ='" & rsDados("Referencia") & "'"
             rdoProduto.CursorLocation = adUseClient
             rdoProduto.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic

             
             ValorlItem = (rsDados("vlunit") * rsDados("Qtde"))
             SubTotal = (SubTotal + ValorlItem)
           
             Printer.Print Trim(rdoProduto("PR_Descricao"))
             Printer.Print rsDados("referencia") _
             & Space(3) & right(Space(4) & Format(rsDados("Qtde"), "###0"), 4) & "x" _
             & Format(rsDados("vlunit"), "###,###,###.00") & Space(5) _
             & right(Space(10) & Format(ValorlItem, "###,###,###.00"), 14)
             rdoProduto.Close
             rsDados.MoveNext
          Loop
       End If
       
       'rsDados.Close
       
         
'       totalvenda = (SubTotal - ValorDesconto)
       Printer.Print ""
'       Printer.Print "SUB TOTAL " & Space(16) & Right(Space(10) & Format(RsDadosCapa("vlrMercadoria"), "###,###,##0.00"), 14)
'       Printer.Print ""
'       Printer.Print "DESCONTO  " & Space(16) & Right(Space(10) & Format(RsDadosCapa("desconto"), "###,###,##0.00"), 14)
'       Printer.Print " "
'       Printer.Print "FRETE     " & Space(16) & Right(Space(10) & Format(RsDadosCapa("fretecobr"), "###,###,##0.00"), 14)
'       Printer.Print " "
       Printer.Print "TOTAL     " & Space(16) & right(Space(10) & Format(RsDadosCapa("totalnota"), "###,###,##0.00"), 14)
       Printer.Print ""
       Printer.Print "________________________________________"
       Printer.Print "Nota Fiscal: "; Format(nf, "000000")
       Printer.Print "======================================"
       Printer.Print " "
       Printer.Print " "
       Printer.Print " "
       Printer.Print " "
       Printer.Print " "
       Printer.Print " "
       Printer.Print " "
       Printer.Print " "
       
     RsDadosCapa.Close
     
     Printer.EndDoc
      Screen.MousePointer = 0

End Sub
