VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form RelRoman 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Emissão de Romaneio"
   ClientHeight    =   4695
   ClientLeft      =   5925
   ClientTop       =   2490
   ClientWidth     =   6990
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   300
      ScaleHeight     =   45
      ScaleWidth      =   6405
      TabIndex        =   9
      Top             =   3765
      Width           =   6405
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   2460
      Left            =   1710
      TabIndex        =   3
      Top             =   330
      Width           =   3600
      Begin VB.OptionButton optLoja 
         BackColor       =   &H00404040&
         Caption         =   "CMCE"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   1
         Left            =   1950
         TabIndex        =   8
         Top             =   285
         Width           =   1170
      End
      Begin VB.OptionButton optLoja 
         BackColor       =   &H00404040&
         Caption         =   "CD"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   0
         Left            =   765
         TabIndex        =   7
         Top             =   285
         Width           =   1170
      End
      Begin VB.ListBox lstForne 
         BackColor       =   &H00A3A3A3&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   645
         Left            =   1320
         MultiSelect     =   1  'Simple
         TabIndex        =   1
         Top             =   1155
         Width           =   1665
      End
      Begin VB.ComboBox cmbLoja 
         BackColor       =   &H00A3A3A3&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Text            =   " "
         Top             =   690
         Width           =   1665
      End
      Begin VB.TextBox mskRefer 
         BackColor       =   &H00A3A3A3&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1305
         TabIndex        =   2
         Text            =   " "
         Top             =   1920
         Width           =   1665
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Loja"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   870
         TabIndex        =   6
         Top             =   795
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Fornecedor"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   390
         TabIndex        =   5
         Top             =   1215
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Referência"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   405
         TabIndex        =   4
         Top             =   2025
         Width           =   780
      End
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   300
      TabIndex        =   10
      Top             =   3165
      Visible         =   0   'False
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin CentroDeDistribuicao.chameleonButton cmdOk 
      Height          =   510
      Left            =   2460
      TabIndex        =   11
      Top             =   3930
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "&Imprime"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   5263440
      BCOLO           =   0
      FCOL            =   16423203
      FCOLO           =   16423203
      MCOL            =   5263440
      MPTR            =   1
      MICON           =   "RelRoman.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdConfig 
      Height          =   510
      Left            =   3885
      TabIndex        =   12
      Top             =   3930
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "&Configura"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   5263440
      BCOLO           =   0
      FCOL            =   16423203
      FCOLO           =   16423203
      MCOL            =   5263440
      MPTR            =   1
      MICON           =   "RelRoman.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdRet 
      Height          =   510
      Left            =   5310
      TabIndex        =   13
      Top             =   3930
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "&Retorna"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   5263440
      BCOLO           =   0
      FCOL            =   16423203
      FCOLO           =   16423203
      MCOL            =   5263440
      MPTR            =   1
      MICON           =   "RelRoman.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "RelRoman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ParaImpr As Boolean

Dim adoRsRomaneio As New ADODB.Recordset

Dim cont As Long
Dim numero As Long
Dim TotReg As Long
Dim wpag As Long
Dim wlin As Long
Dim wcalc As Long
Dim wCont As Long
Dim Ponteiro As Long
Dim Maximo As Long
Dim flg As Long
Dim Ind As Long

Dim wQbLoja As String
Dim wvar As String
Dim wvar1 As String
Dim wvar2 As String
Dim wCond As String
Dim wRodape As String

Dim wVal As Currency
Dim wVal1 As Currency
Dim wVal2 As Currency
Dim wUnit As Currency
Dim wAliq As Currency

Dim rsCodBarras As New ADODB.Recordset
Dim CodBarras As String

Dim wMat As Variant

Dim lojaorigem As String
Dim Impressora As Printer


Private Sub cmbLoja_Click()

    lstForne.Clear

End Sub

Private Sub cmdConfIG_Click()
    
    'MDISup.cdlMDI.ShowPrinter

End Sub

Private Sub Form_Load()
      
    carregarPosicaoTela Me
        
    optloja(0).Value = True
        
    sql = "Select ro_lojadestino " _
        & "From romaneio " _
        & "where ( ro_numeroromaneio = 0 ) " _
        & "and ( ro_situacao = 'A') " _
        & "and ro_lojaorigem in ('CMCE','ALM01','CD') " _
        & "and ro_lojadestino <> '797' " _
        & "group by ro_lojadestino"
        
    adorsLojas.CursorLocation = adUseClient
    adorsLojas.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    
        
    If Not adorsLojas.EOF Then
       Do While Not adorsLojas.EOF
          cmbLoja.AddItem UCase(adorsLojas("ro_lojadestino"))
          adorsLojas.MoveNext
       Loop
    Else
       adorsLojas.Close
       MsgBox "Não existem ítens em romaneio no momento", vbInformation, "Atenção"
       Exit Sub
    End If
    
    adorsLojas.Close
    
    wQbLoja = " "
    cmbLoja.AddItem "CONSO"
    Call GetAsyncKeyState(vbKeyTab)

End Sub

Private Sub CmbLoja_LostFocus()

    If optloja(0).Value = True Then
        lojaorigem = "ALM01', 'CD"
    Else
        lojaorigem = "CMCE"
    End If
    
    If GetAsyncKeyState(vbKeyTab) <> 0 Then
       wCond = ""
       If cmbLoja.Text <> "CONSO" Then
          wCond = " and ro_lojadestino= '" & cmbLoja.Text & "' "
       Else
          wCond = " and ro_lojadestino <> '797' "
       End If
       lstForne.Clear
       
       sql = "Select ro_numeroromaneio,fo_nomefantasia,fo_codigofornecedor " _
        & "From romaneio,fornecedor,produto " _
        & "where ( ro_numeroromaneio = 0 ) " _
        & "and ( ro_situacao = 'A') " _
        & "and ro_lojaorigem in ('" & lojaorigem & "') " _
        & wCond _
        & "and ro_referencia=pr_referencia " _
        & "and pr_codigofornecedor=fo_codigofornecedor " _
        & "group by ro_numeroromaneio,fo_nomefantasia,fo_codigofornecedor"
        
        adoRsRomaneio.CursorLocation = adUseClient
        adoRsRomaneio.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
       
            
        If Not adoRsRomaneio.EOF Then
        
           Do While Not adoRsRomaneio.EOF
              lstForne.AddItem Format(adoRsRomaneio("fo_codigofornecedor"), "000") & " - " & adoRsRomaneio("fo_nomefantasia")
              adoRsRomaneio.MoveNext
           Loop
        Else
           adoRsRomaneio.Close
           MsgBox "Não existem ítens em romaneio no momento", vbInformation, "Atenção"
           Exit Sub
        End If
        
        adoRsRomaneio.Close
    End If

End Sub

Private Sub cmdOK_Click()
    
    Screen.MousePointer = 11
    Frame1.Visible = False
    Label1.Visible = True
    ProgressBar1.Visible = True
    cmdRet.Caption = "&Cancelar"
    wlin = 99
    Me.Refresh
    
  '  Status "Processando dados..."
        
    cmdOK.Visible = True
    cmdConfig.Visible = True
    
    If optloja(0).Value = True Then
        lojaorigem = "ALM01', 'CD"
    Else
        lojaorigem = "CMCE"
    End If

   Dim wImpressora As String
   wImpressora = "ROMANEIO"
   For Each Impressora In Printers
       If UCase(Impressora.DeviceName) = UCase(wImpressora) Then
          Set Printer = Impressora
           Exit For
       End If
   Next
           
    
    wCond = ""
    If Trim(cmbLoja.Text) <> "CONSO" Then
       wCond = " and (RO_LOJADESTINO ='" & left(cmbLoja.Text, 5) & "') "
    Else
       wCond = " and ( ro_lojadestino <> '797 ' ) "
    End If
    
    numero = 0
    For cont = 0 To lstForne.ListCount - 1
       If lstForne.Selected(cont) = True Then
          If numero = 0 Then
             wCond = wCond & " and ((substring(ro_referencia,1,3)='" & left(lstForne.List(cont), 3) & "') "
             numero = 1
          Else
             wCond = wCond & " or (substring(ro_referencia,1,3)= '" & left(lstForne.List(cont), 3) & "') "
          End If
       End If
    Next
    If numero <> 0 Then
       wCond = wCond & ") "
    End If
        
    If Trim(mskRefer.Text) <> "" Then
       wCond = wCond & "and (RO_referencia= '" & left(mskRefer.Text, 7) & "') "
    End If
        
    On Error Resume Next
        
    sql = "Select Count(ro_lojaorigem) As TotReg " _
        & "From romaneio " _
        & "where ( ro_numeroromaneio = 0 ) " _
        & "and ( ro_situacao = 'A') " _
        & wCond _
        & "and ro_lojaorigem in ('" & lojaorigem & "')"
            
    adoRsRomaneio.CursorLocation = adUseClient
    adoRsRomaneio.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    
        
    If adoRsRomaneio.EOF Then
       MsgBox "Problemas no arquivo de romaneios", vbCritical, "Atenção"
       GoTo Finaliza
    Else
       If adoRsRomaneio("totreg") <> 0 Then
          ProgressBar1.Max = adoRsRomaneio("TOTREG")
          TotReg = adoRsRomaneio("totreg")
       Else
          MsgBox "Não encontrou nenhum ítem para esta opção", vbExclamation, "Atenção"
          GoTo Finaliza
       End If
    End If
        
    'Monta consulta
   ' Status ("Imprimindo...")
    Ind = -1
        
     sql = "Select ro_numeroromaneio,ro_lojaorigem,ro_lojadestino," _
        & "ro_referencia,ro_datasolicitacao,ro_datasaida," _
        & "ro_quantidadeenviada,ro_quantidadepedida,ro_situacao," _
        & "pr_descricao, pr_aliquotaipi, pr_precovenda1, pr_precocusto1, ro_sequencia " _
        & "From romaneio, produto " _
        & "where ro_referencia = pr_referencia " _
        & "and (ro_numeroromaneio=0) " _
        & "and (ro_situacao = 'A') " _
        & wCond _
        & "and ro_lojaorigem in ('" & lojaorigem & "') " _
        & "order by ro_lojadestino, ro_referencia"
            
'    SQL = "Select ro_numeroromaneio,ro_lojaorigem,ro_lojadestino," _
        & "ro_referencia,ro_datasolicitacao,ro_datasaida," _
        & "ro_quantidadeenviada,ro_quantidadepedida,ro_situacao," _
        & "pr_descricao, pr_aliquotaipi, pr_precovenda1, pr_precocusto1, ro_sequencia " _
        & "From romaneio, produto " _
        & "where ro_referencia = pr_referencia " _
        & "and (ro_numeroromaneio=0) " _
        & "and (ro_situacao = 'A') " _
        & wCond _
        & "and ro_lojaorigem='ALM01' " _
        & "order by pr_Codigofornecedor, pr_descricao, ro_lojadestino, ro_referencia"
'=======================================
    
'     SQL = "Select ro_numeroromaneio,ro_lojaorigem,ro_lojadestino," _
        & "ro_referencia,ro_datasolicitacao,ro_datasaida," _
        & "ro_quantidadeenviada,ro_quantidadepedida,ro_situacao," _
        & "pr_descricao, pr_aliquotaipi, pr_precovenda1, pr_precocusto1, ro_sequencia " _
        & "From romaneio, produto " _
        & "where ro_referencia = pr_referencia " _
        & "and (ro_numeroromaneio=0) " _
        & "and (ro_situacao = 'A') " _
        & wCond _
        & "and ro_lojaorigem='ALM01' " _
        & "order by ro_lojadestino, pr_Codigofornecedor, pr_descricao"
    
   ' adoRsRomaneio.CursorLocation = adUseClient
   ' adoRsRomaneio.Open SQL, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
   
     
    adorsExtra1.CursorLocation = adUseClient
    adorsExtra1.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    
        
    If adorsExtra1.EOF Then
       MsgBox "Problemas no arquivo de romaneio", vbCritical, "Atenção"
       GoTo Finaliza
    End If
        
    wMat = adorsExtra1.GetRows(TotReg)
    
    adoRsRomaneio.Close
    adorsExtra1.Close
    Maximo = UBound(wMat, 2)
       
    ADO_Cn_CD.BeginTrans
        
    'ROTINA PRINCIPAL
    For Ponteiro = 0 To Maximo Step 1
        flg = flg + 1
        ProgressBar1.Value = flg
        
        DoEvents
        If ParaImpr Then
           Printer.Print "***** INTERROMPIDO PELO USUÁRIO *****"
           Printer.EndDoc
           GoTo Finaliza
        End If
        
        If wQbLoja = " " Then
           wQbLoja = wMat(2, Ponteiro)
        End If
        
        If wQbLoja <> wMat(2, Ponteiro) Then
           Call TotGeral
           Call Numeracao
           wQbLoja = wMat(2, Ponteiro)
        End If
        
        If wCont = 0 Then
           Call Numeracao
        Else
           wCont = wCont + 1
        End If
        
        If wCont > 10 Then
           Call TotGeral
           Call Numeracao
        End If
        
        wMat(0, Ponteiro) = wcalc
        Printer.ScaleMode = vbMillimeters
        'Printer.FontName = "draft 10cpi"
       ' Printer.FontName = "ARIAL"
        Printer.FontName = "COURIER NEW"
        Printer.FontSize = 8
        Printer.ForeColor = "0"
        Printer.FontBold = False
        
        
        sql = ""
        sql = "Select prb_referencia, prb_codigobarras " _
            & " from produtobarras where prb_Referencia='" & wMat(3, Ponteiro) & "'" _
            & " and prb_tipocodigo='B' "
        
        rsCodBarras.CursorLocation = adUseClient
        rsCodBarras.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
        
        ' If Not rsCodBarras.EOF Then
       '     CodBarras = rsCodBarras("prb_codigobarras")
       ' Else
       '     CodBarras = ""
       ' End If
         CodBarras = ""
       If Not rsCodBarras.EOF Then
          
          Do While Not rsCodBarras.EOF
             CodBarras = CodBarras & rsCodBarras("prb_codigobarras") & " "
             rsCodBarras.MoveNext
          
          Loop
       End If
        
        rsCodBarras.Close
        
                
        If wMat(10, Ponteiro) = 2 Or wMat(10, Ponteiro) = 4 Then
           wAliq = (wMat(10, Ponteiro) / 100) + 1
           wUnit = (wMat(11, Ponteiro) / wAliq)
        Else
           wUnit = wMat(12, Ponteiro)
        End If
        
        wVal = wUnit * wMat(7, Ponteiro)
        wVal1 = wVal1 + wVal
        wvar = right$(Space(7) & wMat(3, Ponteiro), 7) & "/" & CodBarras
        
        wvar2 = (left$(Format(wMat(9, Ponteiro), ">") & Space(38), 38)) & "         "
        wvar2 = wvar2 & right$(Space(12) & Format(wVal, "####,###0.00"), 12) & " "
        wvar2 = wvar2 & right$(Space(6) & wMat(7, Ponteiro), 6) & " "
        wvar2 = wvar2 & Space(5) & "________" & "  "
        
        Printer.Print
        Printer.Print wvar
        Printer.Print wvar2
        'Printer.Print
        wlin = wlin + 2
    Next
    
    Call TotGeral
    Call Atualizacao
    
    If Err.Number = 0 Then
       ADO_Cn_CD.CommitTrans
    Else
       
       ADO_Cn_CD.RollbackTrans
       Dim Erro As rdoError
       For Each Erro In rdoErrors
           MsgBox Erro.Number & ": " & Erro.Description, vbCritical, "Erro"
       Next
       rdoErrors.Clear
    End If
    
    Printer.EndDoc

Finaliza:
    flg = 0
    wlin = 99
    wcalc = 0
    wpag = 0
    wCont = 0
    wVal1 = 0
    wQbLoja = " "
    cmbLoja.Text = " "
    lstForne.Text = " "
    mskRefer.Text = "   "
    mskRefer.Enabled = True
    lstForne.Clear
    cmbLoja.Clear
    Screen.MousePointer = 0
    Status ("Pronto")
    ProgressBar1.Value = 0
    cmdOK.Visible = True
    cmdConfig.Visible = True
    Frame1.Visible = True
    Label1.Visible = True
    ProgressBar1.Visible = True
    cmdRet.Caption = "&Retorna"
    
    sql = "Select ro_lojadestino " _
        & "From romaneio " _
        & "where ( ro_numeroromaneio = 0 ) " _
        & "and ( ro_situacao = 'A') " _
        & "and ro_lojaorigem in ('" & lojaorigem & "') " _
        & "and ro_lojadestino <> '797' " _
        & "group by ro_lojadestino"
        
    adorsLojas.CursorLocation = adUseClient
    adorsLojas.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    
        
    If Not adorsLojas.EOF Then
    
       Do While Not adorsLojas.EOF
          cmbLoja.AddItem UCase(adorsLojas("ro_lojadestino"))
          adorsLojas.MoveNext
          
       Loop
    Else
       adorsLojas.Close
       
       MsgBox "Não existem ítens em romaneio no momento", vbInformation, "Atenção"
       Exit Sub
    End If
    
    adorsLojas.Close
    cmbLoja.AddItem "CONSO"
    cmbLoja.SetFocus
    Me.Refresh

End Sub

Private Sub Cabecalho()

    wlin = 7
    
    If wpag > 0 Then
       Printer.NewPage
    End If
    wpag = 1
    
    Printer.ScaleMode = vbMillimeters
   ' Printer.FontName = "draft 10cpi"
    'Printer.FontName = "ARIAL"
    Printer.FontName = "COURIER NEW"
    Printer.FontSize = 12
    Printer.ForeColor = "0"
    Printer.FontBold = False
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    
    Printer.Print "RELATÓRIO DE ROMANEIO DE TRANSFERÊNCIA" & Space(25) & "PAGINA:" & Space(4) & right$(Space(6) & Format(wpag), 6)
    Printer.Print
    Printer.Print "NR.ROMANEIO:" & Space(3) & right$(Space(6) & Format(wcalc), 6) & Space(5) & wMat(1, Ponteiro) & Space(5) & "- LOJA " & wMat(2, Ponteiro) & Space(6) & "IMPR:" & Space(2) & Now
    Printer.Print
    Printer.Print "REFERENCIA/" & "CÓDIGO DE BARRAS"
    Printer.Print "DESCRICAO                                          " & "   VALOR" & "   QTDE" & Space(6) & "AJUSTE"
    'Printer.Print
    'Printer.CurrentY = Printer.CurrentY + 3
    Printer.CurrentX = 0
    
    Dim i As Long
    
    For i = 1 To 8
        Printer.Print "----------";
    Next
    
    Printer.Print
    'Printer.Print

End Sub

Private Sub cmdRet_Click()

    If cmdRet.Caption = "&Retorna" Then
        frmControleCD.lblNomeTelas.Caption = ""
       unload Me
    Else
       Screen.MousePointer = 0
       
       Select Case MsgBox("Cancelar Impressão?", vbQuestion + vbYesNo, "Atenção")
          Case vbYes
               ParaImpr = True
          Case vbNo
               ParaImpr = False
       End Select
       
       Screen.MousePointer = 11
    End If

    unload Me

End Sub

Sub TotGeral()
   
    If wlin > 60 Then
       Call Cabecalho
    End If
    
    If wCont > 0 Then
       wvar = Space(10) & " TOTAL GERAL    " & Space(15) & "ROMANEIO   " & Space(15) & "LOJA DESTINO"
       wvar1 = Space(7) & right$(Space(15) & Format(wVal1, "###,###,###0.00"), 15) & Space(18) & right$(Space(6) & Format(wcalc), 6) & Space(21) & wQbLoja
       Printer.Print
       Printer.Print wvar
       Printer.Print wvar1
    End If

    'Printer.Print Tab(52); "ROMANEIO:"; Tab(70); " LOJA DESTINO "
    'Printer.Print Tab(52); Space(3) & Right$(Space(6) & Format(wcalc), 6); Tab(70); wQbLoja
    
End Sub

Sub Numeracao()
        
    sql = "select cs_numeroromaneio from controlesup"
        
    adorsCtsup.CursorLocation = adUseClient
    adorsCtsup.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    
        
    wcalc = adorsCtsup("cs_numeroromaneio") + 1
    
    ADO_Cn_CD.BeginTrans
        sql = " update controlesup" _
        & " set cs_numeroromaneio = " & wcalc & ""
    ADO_Cn_CD.Execute (sql)
    ADO_Cn_CD.CommitTrans
        
    
    
    adorsCtsup.Close
    
    Call Cabecalho
    
    wVal1 = 0
    wCont = 1
    
End Sub

Sub Atualizacao()

    sql = ""
    
      ADO_Cn_CD.BeginTrans
       
    For Ponteiro = 0 To Maximo Step 1
        sql = sql & "Update Romaneio Set RO_NumeroRomaneio = " & wMat(0, Ponteiro) _
            & ", RO_QuantidadeEnviada = RO_QuantidadePedida " _
            & "Where RO_Sequencia = " & wMat(13, Ponteiro) & Chr(vbKeyReturn)
    Next Ponteiro
    
    ADO_Cn_CD.Execute (sql)
    ADO_Cn_CD.CommitTrans

End Sub

Sub lstForne_LostFocus()

  If GetAsyncKeyState(vbKeyTab) <> 0 Then
     numero = 0
     For cont = 0 To lstForne.ListCount - 1
         If lstForne.Selected(cont) = True Then
            mskRefer.Enabled = False
            cmdOK.SetFocus
            Exit For
         Else
            mskRefer.Enabled = True
            mskRefer.SetFocus
         End If
     Next
  End If

End Sub

Sub lstForne_Click()

     For cont = 0 To lstForne.ListCount - 1
         If lstForne.Selected(cont) = True Then
            mskRefer.Enabled = False
            Exit For
         Else
            mskRefer.Enabled = True
         End If
     Next

End Sub

Private Sub OptLoja_Click(Index As Integer)

    lstForne.Clear

End Sub
