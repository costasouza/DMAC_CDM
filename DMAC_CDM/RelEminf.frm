VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form RelEminf 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Emissão de Nota Fiscal de Transferência"
   ClientHeight    =   5190
   ClientLeft      =   9525
   ClientTop       =   4185
   ClientWidth     =   5190
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   30
      Left            =   150
      TabIndex        =   14
      Top             =   3465
      Width           =   3990
      _Version        =   65536
      _ExtentX        =   7038
      _ExtentY        =   53
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
   End
   Begin VB.Frame FraEmissaoNota 
      ForeColor       =   &H8000000D&
      Height          =   2820
      Left            =   165
      TabIndex        =   7
      Top             =   60
      Width           =   3960
      Begin VB.OptionButton optLojaOrigem 
         Caption         =   "CMCE"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   1
         Left            =   1875
         TabIndex        =   17
         Top             =   240
         Width           =   1200
      End
      Begin VB.OptionButton optLojaOrigem 
         Caption         =   "CD"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   0
         Left            =   720
         TabIndex        =   16
         Top             =   240
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.ComboBox CmbLojas 
         BackColor       =   &H80000004&
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   1455
         Width           =   1605
      End
      Begin VB.TextBox txtUltimoNumero 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   345
         Left            =   1890
         TabIndex        =   13
         Top             =   615
         Width           =   1605
      End
      Begin VB.ListBox mskRoman 
         BackColor       =   &H8000000A&
         Height          =   840
         Left            =   1920
         MultiSelect     =   1  'Simple
         TabIndex        =   11
         Top             =   1815
         Width           =   1605
      End
      Begin MSMask.MaskEdBox mskData 
         Height          =   315
         Left            =   1905
         TabIndex        =   10
         Top             =   1080
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483638
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Caption         =   "Loja"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   750
         TabIndex        =   15
         Top             =   1530
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "Nr.Última NF"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   735
         TabIndex        =   12
         Top             =   765
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Romaneio/Loja"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   735
         TabIndex        =   9
         Top             =   1875
         Width           =   1170
      End
      Begin VB.Label Label2 
         Caption         =   " Data Opcional"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   690
         TabIndex        =   8
         Top             =   1125
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Imprime"
      Height          =   330
      Left            =   495
      TabIndex        =   2
      Top             =   3585
      Width           =   1200
   End
   Begin VB.CommandButton cmdConf 
      Caption         =   "&Configura"
      Height          =   330
      Left            =   1710
      TabIndex        =   3
      Top             =   3585
      Width           =   1215
   End
   Begin VB.CommandButton cmdRetorna 
      Cancel          =   -1  'True
      Caption         =   "&Retorna"
      Height          =   330
      Left            =   2940
      TabIndex        =   4
      Top             =   3585
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1680
      Picture         =   "RelEminf.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   480
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   165
      TabIndex        =   5
      Top             =   3060
      Visible         =   0   'False
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin CentroDeDistribuicao.chameleonButton cmdEncerraNF 
      Height          =   510
      Left            =   240
      TabIndex        =   18
      Top             =   4200
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Encerra NF"
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
      MICON           =   "RelEminf.frx":0442
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "Aguarde..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   1320
      TabIndex        =   6
      Top             =   1065
      Visible         =   0   'False
      Width           =   1320
   End
End
Attribute VB_Name = "RelEminf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ParaImpr As Boolean

Dim wConta As Long
Dim wChave As Long
Dim flg As Long
Dim i As Long
Dim wReduz As Long
Dim wRes As Long
Dim wpag As Long
Dim wlin As Long
Dim tmporient As Long
Dim Ind As Long

Dim Impressora As Printer
'Dim nomeImpressora As Object

Dim serie As String
Dim ImpressoraPadrao As String
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
Dim MatLojas() As String

Dim Data As String

Dim adotransportadora As New ADODB.Recordset
Dim adoExtra1 As New ADODB.Recordset
Dim adoDados As New ADODB.Recordset
Dim adoCodOper As New ADODB.Recordset
Dim adoControle As New ADODB.Recordset
Dim adoEntraNota As New ADODB.Recordset
Dim sql As String

Private Sub cmbLojas_LostFocus()

    PreencheTela

End Sub

Private Sub cmdConf_Click()
    
   ' MDISup.cdlMDI.ShowPrinter

End Sub

Private Sub cmdEncerraNF_Click()
    sairFormulario Me
End Sub

Private Sub Form_Load()
    
    Screen.MousePointer = 11
    
    left = (Screen.Width - Width) / 2
    top = (Screen.Height - Height) / 2
    
    optLojaOrigem(0).Value = True
    
    sql = "Select * from transportadora"
            
    adotransportadora.CursorLocation = adUseClient
    adotransportadora.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
      
        
    If Not adotransportadora.EOF Then
       wnome = adotransportadora("tr_nome")
       wendereco = adotransportadora("tr_endereco")
       wbairro = adotransportadora("tr_bairro")
       westado = adotransportadora("tr_estado")
    End If
    
    adotransportadora.Close
    
    ObtemLojas
    
    CarregaLojas
    
    
    sql = "select cs_numeronotafiscal from controlesup"
    
    adoControle.CursorLocation = adUseClient
    adoControle.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    
    
    txtUltimoNumero = adoControle("cs_numeronotafiscal")
        
    adoControle.Close
    
    'mskData.Text = Format(Date, "dd/mm/yyyy")
    MostraData
   'If adoEntraNota("CV_ProcessandoFecMes") = "S" Then
   '       MsgBox "Fechamento Mensal Processando.", vbInformation, "Atenção"
   '       FraEmissaoNota.Enabled = False
   '       cmdOk.Enabled = False
   '       cmdConf.Enabled = False
   '       adoEntraNota.Close
       '   Unload Me
   ' End If
    
    Screen.MousePointer = 0
    
End Sub

Private Sub cmdOK_Click()
    
    Dim StringRomaneios As String
    Dim rdoNotas As New ADODB.Recordset
    Dim adoPegaLoja As New ADODB.Recordset
    
    On Error Resume Next
    
    If Not DataOK Then
        MsgBox "Data inválida!", vbExclamation, "Atenção"
    
        'mskData.SetFocus
        
        Exit Sub
    End If
    
     wLojaMCE85ouCD = "CD"
    If optLojaOrigem(0).Value = True Then
       lojaorigem = "CD"
      
    Else
       lojaorigem = "CMCE"
        
    End If
    
   
    
    StringRomaneios = ObtemRomaneios
    
    If StringRomaneios <> "" Then
        Screen.MousePointer = 11
        
        Err.Clear
        rdoErrors.Clear
        
        sql = "Select * from romaneio where ro_numeroromaneio in (" & StringRomaneios & ")"
        
        adoPegaLoja.CursorLocation = adUseClient
        adoPegaLoja.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
        
       
        
        If Not adoPegaLoja.EOF Then
           serie = "S2"
           ImpressoraPadrao = "NotaFis"
          
        '   Dim wImpressora As String
        '   wImpressora = "NOTA FISCAL"
        '   For Each Impressora In Printers
        '   If UCase(Impressora.DeviceName) = UCase(wImpressora) Then
        '      Set Printer = Impressora
        '      Exit For
        '   End If
        '   Next
           
            Dim wImpressora As String
            wImpressora = "NOTA FISCAL"
            For Each nomeImpressora In Printers
            If UCase(nomeImpressora.DeviceName) = UCase(wImpressora) Then
               Set Printer = nomeImpressora
            Exit For
            End If
            Next
           
                             
           Do While Not adoPegaLoja.EOF
         '  ADO_Cn_CD.BeginTrans
               
         sql = "CriaNotaTransferencia '" & StringRomaneios & "', '" & Format(mskData.Text, "mm/dd/yyyy") & "'"
             
        'ADO_Cn_CD.Execute (SQL)
        'ADO_Cn_CD.CommitTrans
              
             
              
                 
              rdoNotas.CursorLocation = adUseClient
              rdoNotas.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
        
        
              If Err.Number = 0 Then
                 FraEmissaoNota.Visible = False
                 ProgressBar1.Visible = True
                 wlin = 99
                 Me.Refresh
                 cmdOk.Visible = True
                 cmdConf.Visible = True
                  
                 
                 If Not rdoNotas.EOF Then
                    wSerieImpressao = serie
                    DefineImpressora rdoNotas("Nota")
                 End If
                 
                 Do While Not rdoNotas.EOF
                    
                    Status "Imprimindo Nota Fiscal número " & rdoNotas("Nota") & "..."
                    wControlaQuebraDaPagina = 0
                    
                    wSerieImpressao = serie
                    ImprimirNota rdoNotas("Nota")
                    rdoNotas.MoveNext
                 Loop
              
              Else
                 MsgBox "Problemas na criação da nota de transferencia, Processo não realizado", vbCritical, Me.Caption
                 Screen.MousePointer = 0
                 Exit Sub
              End If
              
              adoPegaLoja.MoveNext
           Loop
        End If
        
        cmdOk.Visible = True
        cmdConf.Visible = True
        FraEmissaoNota.Visible = True
        ProgressBar1.Visible = False
        Me.Refresh
        
        rdoNotas.Close
        adoPegaLoja.Close
        
        PreencheTela
        
        Screen.MousePointer = 0
    Else
        MsgBox "Você deve selecionar, pelo menos, um item na lista.", vbInformation, "Informação"
    End If
    
    Status "Pronto."
    Printer.EndDoc
End Sub



'

'
'
Sub MostraData()
    
    
    
    sql = "Select CV_UltimoDiaMes, CV_ProcessandoFecMes  from ControleFec"
    
    adoEntraNota.CursorLocation = adUseClient
    adoEntraNota.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    
    
    
    If Not adoEntraNota.EOF Then
       If adoEntraNota("CV_ProcessandoFecMes") = "S" Then
          MsgBox "Fechamento Mensal Processando.", vbInformation, "Atenção"
          FraEmissaoNota.Enabled = False
          cmdOk.Enabled = False
          cmdConf.Enabled = False
          Exit Sub
       End If
       If IsNull(adoEntraNota("CV_UltimoDiaMes")) Then
          Data = Format(Date, "dd/mm/yyyy")
       Else
          Data = Format(adoEntraNota("CV_UltimoDiaMes"), "dd/mm/yyyy")
       End If
    End If
    
    adoEntraNota.Close
    mskData = Data
    
End Sub

Function ObtemRomaneios() As String

    Dim Indice As Long
    Dim Maximo As Long
    Dim Retorno As String
    
    Maximo = mskRoman.ListCount - 1

    Retorno = ""

    For Indice = 0 To Maximo Step 1
        If mskRoman.Selected(Indice) Then
            Retorno = Retorno & "," & Val(mskRoman.List(Indice))
        End If
    Next Indice
    
    If Retorno <> "" Then
        Retorno = Mid(Retorno, 2)
    End If

    ObtemRomaneios = Retorno

End Function

Function Imprimir(ByVal NotaFiscal As Long) As Boolean
        
        
    Dim WcodigoOperacao As Long
    Dim lojaorigem As String
  
    
    Printer.ScaleMode = vbMillimeters
    Printer.ForeColor = "0"
    Printer.FontSize = 8
    Printer.FontName = "draft 20cpi"
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.DrawWidth = 3
    
    
    Printer.FontName = "COURIER NEW"
    Printer.FontSize = 8#
    
       
    
    
    If optLojaOrigem(0).Value = True Then
        lojaorigem = "CD"
    Else
        lojaorigem = "CMCE"
    End If
        
    sql = "Select Count(vi_notafiscal) As TotReg " _
        & "From itemnfvenda " _
        & "where vi_notafiscal= " & NotaFiscal & " " _
        & "and vi_serie='" & serie & "' and vi_lojaorigem = '" & lojaorigem & "' "
            
    adoExtra1.CursorLocation = adUseClient
    adoExtra1.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    
        
    ProgressBar1.Max = adoExtra1("TOTREG")
        
    sql = "Select vc_notafiscal,vc_baseicms,vc_enderecocliente," _
        & "vc_serie,vc_lojaorigem,vc_lojadestino,vc_dataemissao," _
        & "vc_codigooperacao,vc_totalnota,vc_valormercadorias,vc_codigooperacaoNovo," _
        & "vc_aliquotaicms,vc_valoricms,vc_situacao,vi_notafiscal," _
        & "vi_serie,vi_referencia,vi_quantidade,vi_precounitario," _
        & "vi_reserva,vi_valormercadoria,vi_valoripi,vi_aliquotaicms, " _
        & "VC_Observacao From capanfvenda,itemnfvenda, produto " _
        & "where vc_notafiscal = vi_notafiscal and vc_serie = vi_serie and vc_lojaorigem = vi_lojaorigem " _
        & "and vi_referencia=pr_referencia and vc_notafiscal= " & NotaFiscal & " and vc_serie='" & serie & "' and vc_lojaorigem = '" & lojaorigem & "'" _
        & "order by PR_CodigoFornecedor, PR_Descricao"
            
    adoExtra1.CursorLocation = adUseClient
    adoExtra1.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
        
    tmporient = Printer.Orientation
    
    wConta = 0
    wChave = 0
    wReduz = 0
    wStr15 = ""
    wStr16 = ""
    wRes = 0
    flg = -1
    
    Do While Not adoExtra1.EOF
       flg = flg + 1
       ProgressBar1.Value = flg
       
       If wChave = 0 Then
          wChave = 1
       
          sql = "select lo_endereco,lo_bairro,lo_municipio,lo_uf," _
              & "lo_cep,lo_cgc,lo_inscricaoestadual,lo_fax,lo_telefone " _
              & "from loja " _
              & "where lo_loja = '" & adoExtra1("vc_lojaorigem") & "' "
       
          adorsExtra2.CursorLocation = adUseClient
          adorsExtra2.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
       
          If Not adorsExtra2.EOF Then
          
              
'             SQL = "Select CN_CODIGOOPERACAONOVO from CodigoOperacaoNovo where CN_CodigoOperacaoAntigo=" & adoExtra1("vc_codigooperacao") & ""
'             Set adoCodOper = rdoCnSup.OpenResultset(SQL)
'
'             If Not adoCodOper.EOF Then
'                WcodigoOperacao = adoCodOper("cn_codigooperacaoNovo")
'             Else
'                WcodigoOperacao = adoExtra1("vc_codigooperacao")
'             End If
          
             wStr1 = Space(5) & left$(Format(Trim(adoExtra1("vc_enderecocliente")), ">") & Space(23), 23) & Space(20) & left$(Format(Trim(adorsExtra2("lo_endereco")), ">") & Space(30), 30) & Space(7) & left$(Format(Trim(adorsExtra2("lo_bairro")), ">") & Space(18), 15) & Space(4) & "X" & Space(33) & left$(Format(adoExtra1("vc_notafiscal"), "###,###"), 7)
             wStr2 = Space(5) & left(IIf(IsNull(adoExtra1("VC_Observacao")), "", adoExtra1("VC_Observacao")) & Space(43), 43) & left$(Format(Trim(adorsExtra2("lo_municipio")), ">") & Space(15), 15) & Space(29) & left$(Trim(adorsExtra2("lo_uf")), 2)
             wStr3 = Space(48) & "(011)" & left$(Trim(Format(adorsExtra2("lo_telefone"), "###-####")), 8) & "/(011)" & left$(Format(adorsExtra2("lo_fax"), "###-####"), 8) & Space(11) & left$(Format(adorsExtra2("lo_cep"), "#####-###"), 8)
             wStr4 = Space(111) & left$(Trim(Format(adorsExtra2("lo_cgc"), "###,###,###")), 10) & "/" & right$(Format(adorsExtra2("lo_cgc"), "####-##"), 7)
             wStr5 = Space(40) & "TRANSFERENCIA" & Space(18) & left$(adoExtra1("Vc_CodigoOperacaoNovo"), 4) & Space(40) & left$(Trim(Format(adorsExtra2("lo_inscricaoestadual"), "###,###,###,###")), 15)
          End If
           
          adorsExtra2.Close
    
          sql = "select em_descricao,lo_endereco,lo_bairro,lo_municipio," _
              & "lo_uf,lo_cep,lo_cgc,lo_inscricaoestadual,lo_fax," _
              & "lo_telefone " _
              & "from loja, empresa " _
              & "where lo_empresa=em_codigoempresa " _
              & "and lo_loja = '" & adoExtra1("vc_lojadestino") & "' "
       
           adorsExtra2.CursorLocation = adUseClient
          adorsExtra2.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
       
          If Not adorsExtra2.EOF Then
             wStr6 = Space(40) & left$(Format(Trim(adorsExtra2("em_descricao")), ">") & Space(50), 50) & Space(21) & left$(Trim(Format(adorsExtra2("lo_cgc"), "###,###,###")), 10) & "/" & right$(Format(adorsExtra2("lo_cgc"), "####-##"), 7) & Space(5) & left$(Format(adoExtra1("vc_dataemissao"), "dd/mm/yyyy"), 12)
             wStr7 = Space(40) & left$(Format(Trim(adorsExtra2("lo_endereco")), ">") & Space(40), 40) & Space(7) & left$(Format(Trim(adorsExtra2("lo_bairro")), ">") & Space(15), 15) & Space(32) & left$(Format(adoExtra1("vc_dataemissao"), "dd/mm/yyyy"), 12)
             wStr8 = Space(40) & left$(Format(Trim(adorsExtra2("lo_municipio")), ">") & Space(15), 15) & Space(43) & left$(Trim(adorsExtra2("lo_uf")), 9) & Space(14) & left$(Trim(Format(adorsExtra2("lo_inscricaoestadual"), "###,###,###,###")), 15)
          End If
        
          wStr9 = Space(4) & right$(Space(12) & Format(adoExtra1("vc_baseicms"), "#,###,##0.00"), 12) & Space(1) & right$(Space(12) & Format(adoExtra1("vc_valoricms"), "#,###,##0.00"), 12) & Space(38) & right$(Space(15) & Format(adoExtra1("vc_valormercadorias"), "#,###,##0.00"), 12)
          wStr10 = Space(67) & right(Space(12) & Format(adoExtra1("vc_totalnota"), "#,###,##0.00"), 12)
          wStr11 = Space(2) & wnome
          wStr12 = Space(2) & wendereco & "                    " & wbairro & "          " & westado
          wStr13 = Space(99) & "LOJA  CD " & Space(11) & right$(Space(7) & Format(adoExtra1("vc_notafiscal"), "###,###"), 7)
          
          adorsExtra2.Close
       
          Printer.ScaleMode = vbMillimeters
          Printer.ForeColor = "0"
          Printer.FontSize = 8
          Printer.FontName = "draft 10cpi"
          Printer.FontSize = 8
          Printer.FontBold = False
          Printer.DrawWidth = 3
       
          Printer.Print
          Printer.Print
          Printer.Print
          Printer.Print "  ROMANEIO:"
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
          & "where pr_referencia = '" & adoExtra1("vi_referencia") & "' order by PR_CodigoFornecedor, PR_Descricao"
     
       adorsExtra2.CursorLocation = adUseClient
          adorsExtra2.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
     
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
      If adoExtra1("vi_reserva") <> "" Then
         wRes = 1
         If wStr16 = "" Then
            wStr16 = wStr16 & wConta
         Else
            wStr16 = wStr16 & "," & wConta
         End If
      End If
      wStr1 = Space(6) & left$(adoExtra1("vi_referencia") & Space(7), 7) & Space(2) & left$(Format(Trim(adorsExtra2("pr_descricao")), ">") & Space(38), 38) & Space(25) & left$(Format(Trim(adorsExtra2("pr_classefiscal")), ">") & Space(10), 10) & Space(2) & left$(Trim(wCodIPI), 1) & left$(Trim(wCodTri), 1) & "  " & Space(2) & left$(Trim(adorsExtra2("pr_unidade")) & Space(2), 2) & Space(5) & right$(Space(6) & Format(adoExtra1("vi_quantidade"), "#####0"), 6) & Space(2) & right$(Space(12) & Format(adoExtra1("vi_precounitario"), "#,###,##0.00"), 12) & Space(2) & right$(Space(12) & Format(adoExtra1("vi_valormercadoria"), "#,###,##0.00"), 15) & Space(2) & right$(Space(2) & Format(adoExtra1("vi_aliquotaicms"), "#0"), 2)
      Printer.Print wStr1
        
      adorsExtra2.Close
      adoExtra1.MoveNext
    Loop
        
    Do While wConta < 10
       wConta = wConta + 1
       Printer.Print
    Loop
    
    If wReduz = 1 Then
'       Printer.Print Space(4) & "                 BASE CALC.REDUZ.CONF. ART.1 DECR. N.34185 DE 18.11.91"
       Printer.Print Space(4) & "                 BASE CALC REDUZ CONF.ART.51. ANEXOS I E II ART. 12 - I,II,III E IV DECR.45.490"
       Printer.Print Space(4) & "                 ITENS " & wStr15
    Else
       Printer.Print
       Printer.Print
    End If
    
    If wRes = 1 Then
       Printer.Print Space(4) & "                 ITENS C/RESERVA " & wStr16
    Else
       Printer.Print
    End If
    
    Printer.Print wStr9
    Printer.Print
    Printer.Print wStr10
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
    Printer.Print
    Printer.CurrentY = Printer.CurrentY - 1
    Printer.Print
   ' Printer.Print
    
'    Printer.EndDoc
    
    adoExtra1.Close
    
Finaliza:
    flg = 0
    wlin = 99
    ProgressBar1.Value = 0

End Function

Private Sub cmdRetorna_Click()
   
    Unload Me

End Sub

Sub PreencheTela()
    Dim lojaorigem As String
    
    mskRoman.Clear
    
   If optLojaOrigem(0).Value = True Then
      lojaorigem = "CD"
   ElseIf optLojaOrigem(1).Value = True Then
      lojaorigem = "CMCE"
   Else
      MsgBox "Uma loja de origem deve ser selecionada.", vbInformation, "Atenção"
      Exit Sub
   End If
      
   If Trim(CmbLojas.Text) = "Todas" Then
    
       sql = "select ro_numeroromaneio,ro_lojadestino " _
         & "from romaneio " _
         & "Where ro_situacao='A' " _
         & "and ro_numeroromaneio <> 0 and ro_lojaorigem = '" & lojaorigem & "' " _
         & "group by ro_numeroromaneio,ro_lojadestino"
       
          adoExtra1.CursorLocation = adUseClient
          adoExtra1.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
          
      
         Do While Not adoExtra1.EOF
             mskRoman.AddItem adoExtra1("ro_numeroromaneio") & " / " & adoExtra1("ro_lojadestino")
             adoExtra1.MoveNext
         Loop
      
   Else
    
        sql = "select ro_numeroromaneio,ro_lojadestino " _
        & "from romaneio " _
        & "where ro_lojadestino='" & CmbLojas.Text & "' " _
        & "and ro_situacao='A' and ro_lojaorigem = '" & lojaorigem & "' " _
        & "and ro_numeroromaneio <> 0 " _
        & "group by ro_numeroromaneio,ro_lojadestino " _
        & "order by ro_numeroromaneio"
    
          adoExtra1.CursorLocation = adUseClient
          adoExtra1.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
      
         Do While Not adoExtra1.EOF
             mskRoman.AddItem adoExtra1("ro_numeroromaneio")
             adoExtra1.MoveNext
         Loop
    
    End If
    
    
    adoExtra1.Close
    
    If mskRoman.ListCount <> 0 Then
       mskRoman.Selected(0) = True
    End If
    
    

End Sub

Private Sub ObtemLojas()

    Dim Conta As Long

    sql = "Select LO_Loja from Loja where LO_Situacao = 'A' and LO_MostraEstoque = 'S' and LO_Loja not in ('182', '183', '184','CMC', 'CONSO','ALM01','CD')"
    
    adoDados.CursorLocation = adUseClient
    adoDados.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
 
    
    Conta = 0
    ReDim MatLojas(Conta) As String
    
    Do While Not adoDados.EOF
        ReDim Preserve MatLojas(Conta) As String
    
        MatLojas(Conta) = adoDados("LO_Loja")
        
        Conta = Conta + 1
        adoDados.MoveNext
    Loop
    
       ReDim Preserve MatLojas(Conta) As String
    
       MatLojas(Conta) = "Todas"

End Sub

Private Sub CarregaLojas()

    Dim i As Long
    Dim Maximo As Long
    
    Maximo = UBound(MatLojas)

    CmbLojas.Clear
    
    For i = 0 To Maximo Step 1
        CmbLojas.AddItem MatLojas(i)
    Next i
    
    CmbLojas.ListIndex = 0

End Sub


Private Sub mskData_LostFocus()

    If mskData.Text <> "__/__/____" Then
        If Not DataOK Then
            mskData.SetFocus
        End If
    End If

End Sub

Private Function DataOK() As Boolean

    DataOK = False

    If InStr(mskData.Text, mskData.PromptChar) = 0 Then
        If IsDate(mskData.Text) Then
            DataOK = True
        End If
    End If

End Function

Sub mskRoman_DblClick()

    Dim cont As Long
    Dim numero As Long
    
    numero = mskRoman.ListIndex
    
    For cont = 0 To mskRoman.ListCount - 1
        If mskRoman.Selected(numero) = True Then
           mskRoman.Selected(cont) = True
        Else
           mskRoman.Selected(cont) = False
        End If
    Next

End Sub

Private Sub optLojaOrigem_Click(Index As Integer)

    mskRoman.Clear

End Sub
