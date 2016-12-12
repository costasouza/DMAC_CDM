VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form ReleminfEsp 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Emissão de N.Fiscal Operações Especiais"
   ClientHeight    =   3750
   ClientLeft      =   3675
   ClientTop       =   5250
   ClientWidth     =   7020
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   300
      ScaleHeight     =   45
      ScaleWidth      =   6405
      TabIndex        =   10
      Top             =   2790
      Width           =   6405
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   1545
      TabIndex        =   5
      Top             =   420
      Width           =   3975
      Begin VB.TextBox TxtLoja 
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
         Height          =   330
         Left            =   3135
         TabIndex        =   4
         Top             =   705
         Width           =   705
      End
      Begin VB.TextBox TxtSerie 
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
         Height          =   330
         Left            =   2055
         TabIndex        =   3
         Top             =   705
         Width           =   510
      End
      Begin VB.OptionButton optnf 
         BackColor       =   &H00404040&
         Caption         =   "Impressão"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   0
         Top             =   300
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.OptionButton optnf 
         BackColor       =   &H00404040&
         Caption         =   "Cancelamento"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   1
         Left            =   2205
         TabIndex        =   1
         Top             =   300
         Width           =   1320
      End
      Begin MSMask.MaskEdBox msknumnf 
         Height          =   315
         Left            =   570
         TabIndex        =   2
         Top             =   720
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   10724259
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Loja:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   2715
         TabIndex        =   9
         Top             =   825
         Width           =   345
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Serie :"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   1575
         TabIndex        =   8
         Top             =   840
         Width           =   450
      End
      Begin VB.Label lbl1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "N. F. :"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   105
         TabIndex        =   6
         Top             =   810
         Width           =   435
      End
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   300
      TabIndex        =   11
      Top             =   2190
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
      TabIndex        =   12
      Top             =   2955
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "&Processar"
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
      MICON           =   "ReleminfEsp.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdConf 
      Height          =   510
      Left            =   3885
      TabIndex        =   13
      Top             =   2955
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
      MICON           =   "ReleminfEsp.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdRetorna 
      Height          =   510
      Left            =   5310
      TabIndex        =   14
      Top             =   2955
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
      MICON           =   "ReleminfEsp.frx":0038
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
      BackColor       =   &H00505050&
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
      ForeColor       =   &H00FA9923&
      Height          =   330
      Left            =   2505
      TabIndex        =   7
      Top             =   810
      Visible         =   0   'False
      Width           =   1320
   End
End
Attribute VB_Name = "ReleminfEsp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ParaImpr As Boolean

Dim wpag As Long
Dim wlin As Long
Dim wConta As Long
Dim wChave As Long
Dim flg As Long
Dim i As Long
Dim wReduz As Long
Dim tmporient As Long
Dim Ind As Long

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
Dim RdodadosItens As New ADODB.Recordset
Dim adoAuxExtra1 As New ADODB.Recordset


Private Sub cmdConf_Click()
    
   ' MDISup.cdlMDI.ShowPrinter

End Sub

Private Sub Form_Load()
    
carregarPosicaoTela Me

End Sub

Private Sub cmdOK_Click()
    
    Screen.MousePointer = 11
    Frame1.Visible = False
    Label1.Visible = True
    ProgressBar1.Visible = True
    cmdRetorna.Caption = "&Cancelar"
    wlin = 99
    Me.Refresh
    
  '  Status ("Processando dados...")
        
    cmdOk.Visible = True
    cmdConf.Visible = True
        
    SQL = "Select Count(vi_notafiscal) As TotReg " _
        & "From itemnfvenda " _
        & "where vi_notafiscal= " & UCase(msknumnf.Text) & " " _
        & "and vi_serie='" & txtSerie.Text & "' " _
        & "and vi_lojaorigem='" & txtLoja.Text & " ' " _
        & "and vi_tiponota <> 'C'"
            
    
    adoAuxExtra1.CursorLocation = adUseClient
    adoAuxExtra1.Open SQL, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    
        
    If adoAuxExtra1.EOF Then
       MsgBox "Nota não existe ou ja foi cancelada", vbCritical, Me.Caption
       adoAuxExtra1.Close
       Call Finaliza
       Exit Sub
    Else
       If adoAuxExtra1("TOTREG") <> 0 Then
          ProgressBar1.Max = adoAuxExtra1("TOTREG")
       Else
          MsgBox "Nota Fiscal Não Encontrada", vbInformation, "Atenção"
          adoAuxExtra1.Close
          Call Finaliza
          Exit Sub
       End If
    End If
        
    
    If optnf(0).Value = True Then
        '  Status ("Imprimindo...")
       '   Call Impressao
          
          If UCase(txtLoja.Text) = "MC85" Or UCase(txtLoja.Text) = "MC85S" Then
             wLojaMCE85ouCD = "MC85"
          Else
             wLojaMCE85ouCD = "CD "
          End If
                    
           wTelaOperacaoEspecial = True
       '    wLojaMCE85ouCD = txtLjOrigem.Text
           LojaOrigem = txtLoja.Text
           wControlaQuebraDaPagina = 0
           
            Dim wImpressora As String
            wImpressora = "NOTA FISCAL"
            For Each NomeImpressora In Printers
            If UCase(NomeImpressora.DeviceName) = UCase(wImpressora) Then
               Set Printer = NomeImpressora
            Exit For
            End If
            Next
           
           
           
           wSerieImpressao = txtSerie.Text
           
           DefineImpressora msknumnf.Text
           ImprimirNota msknumnf.Text
           Printer.EndDoc
    
    ElseIf optnf(1).Value = True Then
       If MsgBox("Confirma o cancelamento na Nota Fiscal Especial? ", vbYesNo + vbQuestion, "Deleção de Nota Fiscal de Transferência") = vbYes Then
         ' Status ("Deletando ...")
          Call Delecao
       Else
          optnf(0).Value = True
          optnf(1).Value = False
          Exit Sub
       End If
    End If
    adoAuxExtra1.Close
    Call Finaliza

End Sub
    
Private Sub Finaliza()

    flg = 0
    wlin = 99
    msknumnf.Text = ""
    txtSerie.Text = ""
    txtLoja.Text = ""
    Screen.MousePointer = 0
   ' Status ("Pronto")
    ProgressBar1.Value = 0
    cmdOk.Visible = True
    cmdConf.Visible = True
    Frame1.Visible = True
    Label1.Visible = True
    ProgressBar1.Visible = False
    cmdRetorna.Caption = "&Retorna"
    msknumnf.SetFocus
    Me.Refresh

End Sub
    
    
Private Sub Impressao()
    Dim adoCarimbo As New ADODB.Recordset
    
    SQL = "Select vc_notafiscal,vc_baseicms,vc_nomecliente,vc_enderecocliente, " _
        & "vc_bairrocliente,vc_municipiocliente,vc_ufcliente,vc_cepcliente,vc_cgccliente, " _
        & "vc_inscestcliente,vc_serie,vc_lojaorigem,vc_lojadestino,vc_dataemissao," _
        & "vc_codigooperacao,vc_totalnota,vc_valormercadorias," _
        & "vc_aliquotaicms,vc_valoricms,vc_situacao,vi_notafiscal," _
        & "vi_serie,vi_referencia,vi_quantidade,vi_precounitario," _
        & "vi_valormercadoria,vi_valoripi,vi_aliquotaicms " _
        & "From capanfvenda,itemnfvenda " _
        & "where vc_notafiscal = vi_notafiscal " _
        & "And vc_serie=vi_serie " _
        & "And vc_dataemissao=vi_dataemissao " _
        & "And vc_notafiscal= " & UCase(msknumnf.Text) & "  " _
        & "And vc_serie='" & txtSerie.Text & "' " _
        & "And vc_lojaorigem= '" & txtLoja.Text & "' " _
        & "And vc_tiponota <> 'C' "
        
    adoAuxExtra1.CursorLocation = adUseClient
    adoAuxExtra1.Open SQL, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    

    If adoAuxExtra1.EOF Then
       MsgBox "Nota Fiscal não existe ou esta cancelada", vbCritical, Me.Caption
       Call Finaliza
       Exit Sub
    End If
    
    'ROTINA PRINCIPAL
    tmporient = Printer.Orientation
    wConta = 0
    wChave = 0
    wReduz = 0
    wStr15 = ""
    
    Do While Not adoAuxExtra1.EOF
        flg = flg + 1
        ProgressBar1.Value = flg
        DoEvents
        If ParaImpr Then
           Printer.Print "***** INTERROMPIDO PELO USUÁRIO *****"
           Printer.EndDoc
           Call Finaliza
           Exit Sub
        End If
        
        If optnf(0).Value = True Then
           If wChave = 0 Then
              wChave = 1
           
              SQL = "select lo_endereco,lo_bairro,lo_municipio,lo_uf," _
                  & "lo_cep,lo_cgc,lo_inscricaoestadual,lo_fax,lo_telefone " _
                  & " from loja " _
                  & " where lo_loja = '" & adoAuxExtra1("vc_lojaorigem") & "' "
              
                  rdorsExtra2.CursorLocation = adUseClient
                  rdorsExtra2.Open SQL, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    
           
              If Not rdorsExtra2.EOF Then
                 wStr1 = Space(48) & left$(Format(Trim(rdorsExtra2("lo_endereco")), ">") & Space(30), 30) & Space(7) & left$(Format(Trim(rdorsExtra2("lo_bairro")), ">") & Space(18), 15) & Space(4) & "X" & Space(33) & left$(Format(adoAuxExtra1("vc_notafiscal"), "###,###"), 7)
                 wStr2 = Space(48) & left$(Format(Trim(rdorsExtra2("lo_municipio")), ">") & Space(15), 15) & Space(29) & left$(Trim(rdorsExtra2("lo_uf")), 2)
                 wStr3 = Space(48) & "(011)" & left$(Trim(Format(rdorsExtra2("lo_telefone"), "###-####")), 8) & "/(011)" & left$(Format(rdorsExtra2("lo_fax"), "###-####"), 8) & Space(11) & left$(Format(rdorsExtra2("lo_cep"), "#####-###"), 8)
                 wStr4 = Space(111) & left$(Trim(Format(rdorsExtra2("lo_cgc"), "###,###,###")), 10) & "/" & right$(Format(rdorsExtra2("lo_cgc"), "####-##"), 7)
                 wStr5 = Space(40) & "Remessa para Conserto" & Space(10) & left$(adoAuxExtra1("vc_codigooperacao"), 3) & Space(40) & left$(Trim(Format(rdorsExtra2("lo_inscricaoestadual"), "###,###,###,###")), 15)
                 'wStr5 = Space(40) & "             " & Space(18) & Left$(adoauxextra1("vc_codigooperacao"), 3) & Space(40) & Left$(Trim(Format(rdorsExtra2("lo_inscricaoestadual"), "###,###,###,###")), 15)
              End If
              
              rdorsExtra2.Close
        
            
              wStr6 = Space(40) & left$(Format(Trim(adoAuxExtra1("Vc_NomeCliente")), ">") & Space(50), 50) & Space(21) & left$(Trim(Format(adoAuxExtra1("Vc_CgcCliente"), "###,###,###")), 10) & "/" & right$(Format(adoAuxExtra1("Vc_CgcCliente"), "####-##"), 7) & Space(5) & left$(Format(adoAuxExtra1("vc_dataemissao"), "dd/mm/yyyy"), 12)
              wStr7 = Space(40) & left$(Format(Trim(adoAuxExtra1("Vc_EnderecoCliente")), ">") & Space(40), 40) & Space(7) & left$(Format(Trim(adoAuxExtra1("Vc_BairroCliente")), ">") & Space(15), 15) & Space(32) & left$(Format(adoAuxExtra1("vc_dataemissao"), "dd/mm/yyyy"), 12)
              wStr8 = Space(40) & left$(Format(Trim(adoAuxExtra1("Vc_MunicipioCliente")), ">") & Space(15), 15) & Space(43) & left$(Trim(adoAuxExtra1("Vc_UFCliente")), 9) & Space(14) & left$(Trim(Format(adoAuxExtra1("Vc_InscEstCliente"), "###,###,###,###")), 15)
           
              wStr9 = Space(4) & right$(Space(12) & Format(adoAuxExtra1("vc_baseicms"), "########0.00"), 12) & Space(1) & right$(Space(12) & Format(adoAuxExtra1("vc_valoricms"), "########0.00"), 12) & Space(38) & right$(Space(15) & Format(adoAuxExtra1("vc_valormercadorias"), "########0.00"), 12)
              wStr10 = Space(67) & right(Space(12) & Format(adoAuxExtra1("vc_totalnota"), "########0.00"), 12)
              wStr11 = Space(2) & "                          "
              wStr12 = Space(2) & "                                                     "
              wStr13 = Space(99) & "LOJA  CD " & Space(11) & right$(Space(7) & Format(adoAuxExtra1("vc_notafiscal"), "###,###"), 7)
           
              Printer.ScaleMode = vbMillimeters
              Printer.ForeColor = "0"
              Printer.FontSize = 8
              Printer.FontName = "draft 10cpi"
              Printer.FontSize = 8
              Printer.FontBold = False
              Printer.DrawWidth = 3
           
              Printer.Print
              Printer.Print
              Printer.Print "           "
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
        
           SQL = "select pr_codigoipi,pr_codigoreducaoicms,pr_descricao," _
               & "pr_classefiscal,pr_unidade " _
               & " from produto " _
               & " where pr_referencia = '" & adoAuxExtra1("vi_referencia") & "' "
        
                  rdorsExtra2.CursorLocation = adUseClient
                  rdorsExtra2.Open SQL, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
        
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
              wStr15 = wStr15 & wConta & " "
           End If
           wStr1 = Space(6) & left$(adoAuxExtra1("vi_referencia") & Space(7), 7) & Space(2) & left$(Format(Trim(rdorsExtra2("pr_descricao")), ">") & Space(38), 38) & Space(25) & left$(Format(Trim(rdorsExtra2("pr_classefiscal")), ">") & Space(10), 10) & Space(2) & left$(Trim(wCodIPI), 1) & left$(Trim(wCodTri), 1) & "  " & Space(2) & left$(Trim(rdorsExtra2("pr_unidade")) & Space(2), 2) & Space(5) & right$(Space(6) & Format(adoAuxExtra1("vi_quantidade"), "#####0"), 6) & Space(2) & right$(Space(12) & Format(adoAuxExtra1("vi_precounitario"), "########0.00"), 12) & Space(2) & right$(Space(12) & Format(adoAuxExtra1("vi_valormercadoria"), "########0.00"), 15) & Space(2) & right$(Space(2) & Format(adoAuxExtra1("vi_aliquotaicms"), "#0"), 2)
           Printer.Print wStr1
           
           rdorsExtra2.Close
           adoAuxExtra1.MoveNext
        End If
        
        If optnf(1).Value = True Then
           Call Delecao
           
           adoAuxExtra1.MoveNext
        End If
        
    Loop
        
    If optnf(0).Value = True Then
       Do While wConta < 9
          wConta = wConta + 1
          Printer.Print
       Loop
       If wReduz = 1 Then
          Printer.Print Space(4) & "                 BASE CALC.REDUZ.CONF. ART.1 DECR. N.34185 DE 18.11.91"
          Printer.Print Space(4) & "                 ITENS " & wStr15
       Else
          SQL = "Select on_carimbo1,on_carimbo2 " _
            & "from observacaonotafiscal " _
            & "where ON_NumeroNotaFiscal= " & UCase(msknumnf.Text) & " " _
            & "and ON_Serie in ('S1','S2','S3') " _
            & "and ON_Loja='" & txtLoja.Text & "'"
              
              
          adoCarimbo.CursorLocation = adUseClient
          adoCarimbo.Open SQL, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
                  
                
            
          If Not adoCarimbo.EOF Then
             Printer.Print Trim(adoCarimbo("on_carimbo1"))
             Printer.Print Trim(adoCarimbo("on_carimbo2"))
          Else
             Printer.Print
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
       Printer.Print
       Printer.Print
       Printer.EndDoc
       Printer.Orientation = 1
       
       adoAuxExtra1.Close
    End If
    
End Sub

Private Sub cmdRetorna_Click()
   
    If cmdRetorna.Caption = "&Retorna" Then
        frmControleCD.lblNomeTelas.Caption = ""
       Unload Me
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
    
    sairFormulario Me

End Sub

Private Sub Delecao()
  
    Screen.MousePointer = 11
   
    On Error Resume Next
    
    rdoCnSupBatch.BeginTrans
         
    SQL = "update capanfvenda " _
        & "set vc_tiponota= 'C' " _
        & "where vc_notafiscal = " & UCase(msknumnf.Text) & " " _
        & "And vc_serie='" & txtSerie.Text & "' " _
        & "And vc_lojaorigem='" & txtLoja.Text & "'"
    Set RdodadosItens = rdoCnSupBatch.OpenResultset(SQL)
      
      
    SQL = "update itemnfvenda " _
        & "Set vi_tiponota= 'C' " _
        & "Where vi_notafiscal = " & UCase(msknumnf.Text) & " " _
        & "And vi_serie='" & txtSerie.Text & "' " _
        & "And vi_lojaorigem='" & txtLoja.Text & "'"
    Set RdodadosItens = rdoCnSupBatch.OpenResultset(SQL)
         
    
    If Err.Number = 0 Then
       rdoCnSupBatch.CommitTrans
    Else
       rdoCnSupBatch.RollbackTrans
       Screen.MousePointer = 0
       MsgBox "Erro ao cancelar os registros das notas", vbCritical, Me.Caption
       Exit Sub
    End If
    
    
    
    rdoCnSupBatch.BeginTrans
    SQL = "Select Vc_LojaOrigem,Vc_CodigoOperacao,Vi_Referencia,Vi_Quantidade " _
        & "From CapaNfVenda,ItemNfVenda " _
        & "Where Vc_Notafiscal=Vi_Notafiscal " _
        & "And Vc_LojaOrigem=Vi_LojaOrigem " _
        & "And Vc_Serie=Vi_Serie " _
        & "And Vc_Notafiscal=" & msknumnf.Text & " " _
        & "And Vc_LojaOrigem='" & txtLoja.Text & "' " _
        & "And Vc_Serie='" & txtSerie.Text & "'"
      
     RdodadosItens.CursorLocation = adUseClient
     RdodadosItens.Open SQL, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
          
        
      
    If Not RdodadosItens.EOF Then
       Do While Not RdodadosItens.EOF
          If Trim(UCase(RdodadosItens("Vc_LojaOrigem"))) = "CMC" Then
             SQL = "Update Estoque " _
                 & "Set Es_Estoque=Es_Estoque + " & RdodadosItens("Vi_Quantidade") & " " _
                 & "Where Es_Referencia='" & RdodadosItens("Vi_Referencia") & "' " _
                 & "And Es_Loja= '" & RdodadosItens("Vc_LojaOrigem") & "'"
             
            rdoCnSupBatch.Execute (SQL)
                   
            If Err.Number = 0 Then
               rdoCnSupBatch.CommitTrans
            Else
               rdoCnSupBatch.RollbackTrans
               Screen.MousePointer = 0
               MsgBox "Erro ao Atualizar estoque dos Itens", vbCritical, Me.Caption
               Exit Sub
            End If
             
             SQL = "Update Estoque " _
                 & "Set Es_Estoque=Es_Estoque - " & RdodadosItens("Vi_Quantidade") & " " _
                 & "Where Es_Referencia='" & RdodadosItens("Vi_Referencia") & "' " _
                 & "And Es_Loja= 'CMCS'"
             
          Else
             If RdodadosItens("Vc_CodigoOpercao") > 5999 Then
                SQL = "Update Estoque " _
                 & "Set Es_Estoque=Es_Estoque + " & RdodadosItens("Vi_Quantidade") & " " _
                 & "Where Es_Referencia='" & RdodadosItens("Vi_Referencia") & "' " _
                 & "And Es_Loja='" & RdodadosItens("Vc_LojaOrigem") & "'"
                
             Else
                SQL = "Update Estoque " _
                 & "Set Es_Estoque=Es_Estoque - " & RdodadosItens("Vi_Quantidade") & " " _
                 & "Where Es_Referencia='" & RdodadosItens("Vi_Referencia") & "' " _
                 & "And Es_Loja= '" & RdodadosItens("Vc_LojaOrigem") & "'"
                
             End If
          End If
                 
          rdoCnSupBatch.Execute (SQL)
                 
          If Err.Number = 0 Then
             rdoCnSupBatch.CommitTrans
          Else
             rdoCnSupBatch.RollbackTrans
             Screen.MousePointer = 0
             MsgBox "Erro ao Atualizar estoque dos Itens", vbCritical, Me.Caption
             Exit Sub
          End If
       RdodadosItens.MoveNext
       Loop
    Else
       Screen.MousePointer = 0
       MsgBox "Registros para o cancelamento não encontrados", vbCritical, Me.Caption
       Exit Sub
    End If
    
End Sub
