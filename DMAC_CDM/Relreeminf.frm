VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form RelReemiNF 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Reimpressão/Cancelamento de NF "
   ClientHeight    =   3780
   ClientLeft      =   2820
   ClientTop       =   4500
   ClientWidth     =   6975
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   285
      ScaleHeight     =   45
      ScaleWidth      =   6405
      TabIndex        =   12
      Top             =   2880
      Width           =   6405
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1320
      Left            =   330
      TabIndex        =   8
      Top             =   330
      Width           =   1680
      Begin VB.OptionButton optLojaOrigem 
         BackColor       =   &H00404040&
         Caption         =   "CD"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   10
         Top             =   465
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.OptionButton optLojaOrigem 
         BackColor       =   &H00404040&
         Caption         =   "CMCE"
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   9
         Top             =   825
         Width           =   1215
      End
      Begin VB.Label lblOpcoes 
         BackColor       =   &H00404040&
         Caption         =   "Opções"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FA9923&
         Height          =   195
         Left            =   60
         TabIndex        =   11
         Top             =   60
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1320
      Left            =   2085
      TabIndex        =   4
      Top             =   315
      Width           =   4620
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
         Height          =   315
         Left            =   2715
         TabIndex        =   3
         Top             =   825
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.OptionButton optnf 
         BackColor       =   &H00404040&
         Caption         =   "Cancelamento"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   1
         Left            =   1515
         TabIndex        =   1
         Top             =   435
         Width           =   1500
      End
      Begin VB.OptionButton optnf 
         BackColor       =   &H00404040&
         Caption         =   "Reimpressão"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   0
         Top             =   435
         Value           =   -1  'True
         Width           =   1305
      End
      Begin MSMask.MaskEdBox msknumnf 
         Height          =   315
         Left            =   1095
         TabIndex        =   2
         Top             =   825
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   10724259
         ForeColor       =   4210752
         MaxLength       =   6
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
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "Processo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FA9923&
         Height          =   195
         Left            =   60
         TabIndex        =   16
         Top             =   45
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Serie :"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   2220
         TabIndex        =   7
         Top             =   945
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lbl1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Nota Fiscal :"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   945
         Width           =   885
      End
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   285
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin CentroDeDistribuicao.chameleonButton cmdOk 
      Height          =   510
      Left            =   2445
      TabIndex        =   13
      Top             =   3045
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Processar"
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
      MICON           =   "Relreeminf.frx":0000
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
      Left            =   3870
      TabIndex        =   14
      Top             =   3045
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Configura"
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
      MICON           =   "Relreeminf.frx":001C
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
      Left            =   5295
      TabIndex        =   15
      Top             =   3045
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Retorna"
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
      MICON           =   "Relreeminf.frx":0038
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
Attribute VB_Name = "RelReemiNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ParaImpr As Boolean
Dim adoItensNf As New ADODB.Recordset
Dim adoCarimbo As New ADODB.Recordset
Dim rdorsExtra1 As New ADODB.Recordset
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
Dim wDescricao As String

Dim Wnatureza As String
Dim rdoTransportadora As New ADODB.Recordset

Private Sub cmdConf_Click()
    
  '  MDISup.cdlMDI.ShowPrinter

End Sub

Private Sub Form_Load()
    
    carregarPosicaoTela Me
    
    sql = "Select * from transportadora"
            
    rdoTransportadora.CursorLocation = adUseClient
    rdoTransportadora.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
     
    If Not rdoTransportadora.EOF Then
       wnome = rdoTransportadora("tr_nome")
       wendereco = rdoTransportadora("tr_endereco")
       wbairro = rdoTransportadora("tr_bairro")
       westado = rdoTransportadora("tr_estado")
    End If
    serie = "S2"
    rdoTransportadora.Close
    
End Sub

Private Sub cmdOK_Click()
    
     If Trim(msknumnf.Text) = "" Then
       MsgBox "Informe o numero da nota", vbCritical, "Atenção"
       
       Exit Sub
    End If
    
    
    
    Screen.MousePointer = 11
    Frame1.Visible = False
    ProgressBar1.Visible = True
    cmdRetorna.Caption = "&Cancelar"
    wlin = 99
    Me.Refresh
    
   ' Status ("Processando dados...")
    cmdOK.Visible = True
    cmdConf.Visible = True
    
    wLojaMCE85ouCD = "CD"
    If optLojaOrigem(0).Value = True Then
        lojaorigem = "CD"
    Else
        lojaorigem = "CMCE"
    End If
    
   
    
          ' ImpressoraPadrao = "NotaFis"
        '   For Each Impressora In Printers
        '       If Impressora.DeviceName = "NOTA FISCAL" Then
        '          Set Printer = Impressora
        '          Exit For
        '        End If
        '   Next
    
 '    For Each NomeImpressora In Printers
 '       If Trim(NomeImpressora.DeviceName) = UCase(GLB_ImpCotacao) Then
 '           ' Seta impressora no sistema
 '           Set Printer = NomeImpressora
 '           Exit For
 '       End If
 '   Next
    
    
  ' Dim wImpressora As String
  ' wImpressora = "nota fiscal"
  
 ' Dim X As Printer
 ' For Each X In Printers
 '   If X.Orientation = vbPRORPortrait Then
 '       ' Configurar impressora como padrão do sistema.
 '       Set Printer = X
 '       ' Encerrar a procura de impressora.
 '       Exit For
  '  End If
 ' Next
   
 
  
  Dim wImpressora As String
  wImpressora = "NOTA FISCAL"
   For Each nomeImpressora In Printers
       If UCase(nomeImpressora.DeviceName) = UCase(wImpressora) Then
          Set Printer = nomeImpressora
          
           Exit For
       End If
   Next
    
  'GLB_ImpCotacao = "NOTA FISCAL"
  '
  'For Each NomeImpressora In Printers
  '      If Trim(NomeImpressora.DeviceName) = UCase(GLB_ImpCotacao) Then
  '          ' Seta impressora no sistema
  '          Set Printer = NomeImpressora
  '          Exit For
  '      End If
  '  Next
    
    sql = "Select Count(Vc_notafiscal) As TotReg " _
        & "From CapanfVenda " _
        & "Where Vc_NotaFiscal= " & UCase(msknumnf.Text) & " " _
        & "And Vc_Serie='" & serie & "' " _
        & "And Vc_LojaOrigem in ('CD','CMCE') And Vc_TipoNota <>'C'"
                
        
        rdorsExtra1.CursorLocation = adUseClient
        rdorsExtra1.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    
            
        If rdorsExtra1.EOF Then
           MsgBox "Nota fiscal não encontrada", vbInformation, "Atenção"
           GoTo Finaliza
        Else
           If rdorsExtra1("TOTREG") <> 0 Then
              ProgressBar1.Max = rdorsExtra1("TOTREG")
           Else
              MsgBox "Nota fiscal não existe ou ja foi cancelada", vbInformation, "Atenção"
              GoTo Finaliza
           End If
        End If
 wSerieImpressao = serie
 DefineImpressora msknumnf.Text
                  
 If optnf(0).Value = True Then
   ' Status ("Imprimindo.....")
    'Call Imprimir
    wControlaQuebraDaPagina = 0
    
    wSerieImpressao = serie
    ImprimirNota msknumnf.Text
    Printer.EndDoc
 End If
 
 If optnf(1).Value = True Then
    If MsgBox("Confirma Cancelamento? ", vbYesNo + vbQuestion, "Deleção de Nota Fiscal") = vbYes Then
      ' Status ("Deletando ....")
       Call Delecao 'Cancelamento
    End If
 End If


Finaliza:
    flg = 0
    wlin = 99
    msknumnf.Mask = ""
    msknumnf.Text = ""
   ' TxtSerie.Text = ""
    Screen.MousePointer = 0
   ' Status ("Pronto")
    ProgressBar1.Value = 0
    cmdOK.Visible = True
    cmdConf.Visible = True
    Frame1.Visible = True
    ProgressBar1.Visible = True
    cmdRetorna.Caption = "&Retorna"
    msknumnf.SetFocus
    Me.Refresh
    
    rdorsExtra1.Close

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
  

    ADO_Cn_CD.BeginTrans
        sql = "Update CapanfVenda " _
            & "Set Vc_Tiponota= 'C' " _
            & "Where Vc_NotaFiscal = " & UCase(msknumnf.Text) & " " _
            & "And Vc_Serie = '" & serie & "' " _
            & "And Vc_LojaOrigem in ('CD','CMCE')"
    ADO_Cn_CD.Execute (sql)
    ADO_Cn_CD.CommitTrans
         
    
    ADO_Cn_CD.BeginTrans
    
        sql = "Update ItemnfVenda " _
            & "Set Vi_Tiponota= 'C' " _
            & "Where Vi_NotaFiscal = " & UCase(msknumnf.Text) & " " _
            & "And Vi_Serie = '" & serie & "' " _
            & "And Vi_LojaOrigem in ('CD','CMCE')"
    ADO_Cn_CD.Execute (sql)
    ADO_Cn_CD.CommitTrans
    
    
    sql = "Select Vc_LojaDestino,Vc_LojaVenda,Vc_CodigoOperacao, " _
        & "Vc_Notafiscal,Vc_Serie,Vc_LojaOrigem, " _
        & "Vi_Referencia,Vi_Quantidade " _
        & "From CapanfVenda,ItemnfVenda " _
        & "Where Vc_NotaFiscal=Vi_NotaFiscal " _
        & "And Vc_Serie=Vi_Serie " _
        & "And Vc_Lojaorigem=Vi_LojaOrigem " _
        & "And Vi_NotaFiscal = " & UCase(msknumnf.Text) & " " _
        & "And Vi_Serie = '" & serie & "' " _
        & "And Vi_LojaOrigem in ('CD','CMCE')"
    
    adoItensNf.CursorLocation = adUseClient
    adoItensNf.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
        
    
    Do While Not adoItensNf.EOF
        If adoItensNf("Vc_LojaDestino") <> "CMC" Then
           If adoItensNf("vc_codigooperacao") > 499 Then
             ADO_Cn_CD.BeginTrans
              sql = "Update estoque " _
                  & "Set  " _
                  & "Es_estoque = Es_estoque + " & adoItensNf("Vi_quantidade") & " " _
                  & "Where (Es_Referencia= '" & UCase(adoItensNf("Vi_referencia")) & "') " _
                  & "and (Es_loja = '" & adoItensNf("Vc_LojaOrigem") & "') "
             ADO_Cn_CD.Execute (sql)
             ADO_Cn_CD.CommitTrans
           
             ADO_Cn_CD.BeginTrans
              sql = "Update estoque " _
                  & "Set  " _
                  & "Es_Transito = Es_Transito - " & adoItensNf("Vi_quantidade") & " " _
                  & "Where (Es_Referencia= '" & UCase(adoItensNf("Vi_referencia")) & "') " _
                  & "and (Es_Loja= '" & adoItensNf("VC_LojaDestino") & "') "
             ADO_Cn_CD.Execute (sql)
             ADO_Cn_CD.CommitTrans
                             
             ADO_Cn_CD.BeginTrans
              sql = "Delete Espelho " _
                  & "Where EP_NotaFiscal= " & adoItensNf("Vc_NotaFiscal") & " " _
                  & "And EP_Serie= '" & UCase(adoItensNf("Vc_Serie")) & "'  " _
                  & "And EP_LojaOrigem ='" & adoItensNf("Vc_Lojaorigem") & "'  "
             ADO_Cn_CD.Execute (sql)
             ADO_Cn_CD.CommitTrans
           End If
        Else
             ADO_Cn_CD.BeginTrans
              sql = "update estoque " _
                  & "set  " _
                  & "es_estoque = Es_estoque - " & adoItensNf("vi_quantidade") & " " _
                  & "where (es_referencia= '" & UCase(adoItensNf("vi_referencia")) & "') " _
                  & "and (es_loja= 'CMC') "
             ADO_Cn_CD.Execute (sql)
             ADO_Cn_CD.CommitTrans
                 
             ADO_Cn_CD.BeginTrans
              sql = "update estoque " _
                  & "set  " _
                  & "Es_Transito = Es_Transito + " & adoItensNf("vi_quantidade") & " " _
                  & "where (es_referencia= '" & UCase(adoItensNf("vi_referencia")) & "') " _
                  & "and (es_loja= 'CMC') "
             ADO_Cn_CD.Execute (sql)
             ADO_Cn_CD.CommitTrans
        End If
        adoItensNf.MoveNext
    Loop
    
    On Error Resume Next
       
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
  
End Sub

Private Sub optnf_Click(Index As Integer)
Select Case Index
Case 0
   cmdOK.Caption = "&Imprimir"
Case 1
   cmdOK.Caption = "&Cancelar"
End Select
msknumnf.SetFocus
End Sub

