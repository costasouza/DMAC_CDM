VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form RelReroman 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Reimpressão de Romaneio"
   ClientHeight    =   3750
   ClientLeft      =   3945
   ClientTop       =   4785
   ClientWidth     =   7020
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   315
      ScaleHeight     =   45
      ScaleWidth      =   6405
      TabIndex        =   7
      Top             =   2805
      Width           =   6405
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Romaneios :"
      ForeColor       =   &H000000C0&
      Height          =   1335
      Left            =   1260
      TabIndex        =   3
      Top             =   345
      Width           =   4650
      Begin VB.TextBox mskNumero 
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
         Index           =   1
         Left            =   2715
         TabIndex        =   1
         Text            =   " "
         Top             =   570
         Width           =   1125
      End
      Begin VB.TextBox mskNumero 
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
         Index           =   0
         Left            =   990
         TabIndex        =   0
         Text            =   " "
         Top             =   570
         Width           =   1125
      End
      Begin VB.Label lblRomaneios 
         BackColor       =   &H00404040&
         Caption         =   "Romaneios :"
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
         Left            =   75
         TabIndex        =   6
         Top             =   60
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Até"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   2340
         TabIndex        =   5
         Top             =   675
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "De"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   660
         TabIndex        =   4
         Top             =   675
         Width           =   210
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   2325
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   645
      Visible         =   0   'False
      Width           =   480
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   315
      TabIndex        =   8
      Top             =   2205
      Visible         =   0   'False
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin CentroDeDistribuicao.chameleonButton cmdOk 
      Height          =   510
      Left            =   2475
      TabIndex        =   9
      Top             =   2970
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
      MICON           =   "RelReroman.frx":0000
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
      Left            =   3900
      TabIndex        =   10
      Top             =   2970
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
      MICON           =   "RelReroman.frx":001C
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
      Left            =   5325
      TabIndex        =   11
      Top             =   2970
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
      MICON           =   "RelReroman.frx":0038
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
Attribute VB_Name = "RelReroman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim wvar As String
Dim wvar1 As String
Dim wRodape As String

Dim wVal As Currency
Dim wVal1 As Currency
Dim wVal2 As Currency
Dim wUnit As Currency
Dim wAliq As Currency

Dim wpag As Long
Dim wlin As Long
Dim flg As Long

Dim CodBarras As String
Dim qtde As String * 4
Dim Valor As String * 8

Dim ParaImpr As Boolean

Dim rsCodBarras As New ADODB.Recordset

Dim Impressora As Printer


'Flag para identificar qdo botado inicia pode ser habilitado
'Compara no change de cada campo
'Todos os flags deverao ter 1 para habilitar
Dim VerFlag(1) As Long

Private Sub cmdConfIG_Click()
    
'    MDISup.cdlMDI.ShowPrinter

End Sub

Private Sub Form_Load()
    
    carregarPosicaoTela Me
    
    Erase VerFlag
    
End Sub

Private Sub cmdOK_Click()
    
    Dim rdoRomaneios As New ADODB.Recordset
    
    Screen.MousePointer = 11
    
   ' For Each Impressora In Printers
   '    If Impressora.DeviceName = "ROMANEIO" Then
   '       Set Printer = Impressora
   '        Exit For
   '    End If
   ' Next
    
   
  '  For Each NomeImpressora In Printers
  '      If Trim(NomeImpressora.DeviceName) = UCase(GLB_ImpCotacao) Then
  '          ' Seta impressora no sistema
  '          Set Printer = NomeImpressora
  '          Exit For
  '      End If
  '  Next
    
    
   Dim wImpressora As String
   wImpressora = "ROMANEIO"
   For Each Impressora In Printers
       If UCase(Impressora.DeviceName) = UCase(wImpressora) Then
          Set Printer = Impressora
           Exit For
       End If
   Next
    
    sql = "Select Distinct RO_NumeroRomaneio " _
        & "From Romaneio Where RO_lojaorigem <> 'CD2' and RO_NumeroRomaneio Between " & Val(mskNumero(0).Text) _
        & " and " & Val(mskNumero(1).Text)
    
    rdoRomaneios.CursorLocation = adUseClient
    rdoRomaneios.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    
    If rdoRomaneios.EOF Then
        rdoRomaneios.Close
        
        Screen.MousePointer = 0
        
        MsgBox "Nenhum registro foi encontrado.", vbInformation, "Informação"
    Else
        Frame1.Visible = False
        Label1.Visible = True
        ProgressBar1.Value = 0
        ProgressBar1.Visible = True
        Picture1.Visible = True
        cmdRet.Caption = "&Cancelar"
        cmdOK.Visible = False
        cmdConfig.Visible = False
        
        
        Do
            ProgressBar1.Value = 0
            flg = 0
            Me.Refresh
            
            ImprimeRomaneio rdoRomaneios("RO_NumeroRomaneio")
            
            rdoRomaneios.MoveNext
            
            If Not rdoRomaneios.EOF Then
                Printer.NewPage
            End If
        Loop While Not rdoRomaneios.EOF
        
        rdoRomaneios.Close
        
        Printer.EndDoc
        
        
        Screen.MousePointer = 0
        
     '   Status ("Pronto.")
        
        cmdOK.Visible = True
        cmdConfig.Visible = True
        cmdOK.Enabled = False
        Frame1.Visible = True
        Label1.Visible = True
        ProgressBar1.Visible = False
        Picture1.Visible = False
        cmdRet.Caption = "&Retorna"
        
    End If
    
    Me.Refresh
    
    Call CmpFlags(VerFlag)
    
End Sub

Sub ImprimeRomaneio(ByVal numeroromaneio As Long)

    
    wlin = 99
    
    Me.Refresh
    
   ' Status ("Processando dados...")
        
    sql = "Select Count(ro_numeroromaneio) As TotReg " _
        & "From romaneio " _
        & "where ro_lojaorigem <> 'CD2' and ro_numeroromaneio = " & numeroromaneio
        
     
    adorsExtra2.CursorLocation = adUseClient
    adorsExtra2.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    
            
    ProgressBar1.Max = adorsExtra2("totreg")
        
    'Monta consulta
  '  Status ("Imprimindo...")
        
    sql = "Select *,pr_referencia,pr_descricao,pr_aliquotaipi," _
        & "pr_precovenda1,pr_precocusto1 " _
        & "From romaneio,produto " _
        & "where ro_lojaorigem <> 'CD2' and ro_referencia = pr_referencia " _
        & "and ro_numeroromaneio = " & numeroromaneio & " " _
        & "order by pr_codigofornecedor, pr_descricao"
            
    adorsExtra1.CursorLocation = adUseClient
    adorsExtra1.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic

            
    'ROTINA PRINCIPAL
    wpag = 0
    wVal1 = 0
    
    Do While Not adorsExtra1.EOF
       flg = flg + 1
       ProgressBar1.Value = flg
       
       If wlin > 70 Then
          Call Cabecalho
       End If
       
      sql = ""
      sql = "Select prb_referencia, prb_codigobarras " _
            & " from produtobarras where prb_Referencia='" & adorsExtra1("ro_Referencia") & "'" _
            & " and prb_tipocodigo='B' "
      
       rsCodBarras.CursorLocation = adUseClient
       rsCodBarras.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    

     ' If Not rsCodBarras.EOF Then
     '       CodBarras = rsCodBarras("prb_codigobarras")
     ' Else
     '       CodBarras = ""
     ' End If
       CodBarras = ""
       
       If Not rsCodBarras.EOF Then
                
          Do While Not rsCodBarras.EOF
             CodBarras = CodBarras & Trim(rsCodBarras("prb_codigobarras")) & " "
             rsCodBarras.MoveNext
          
          Loop
       End If
        
      
      rsCodBarras.Close
        
       Printer.ScaleMode = vbMillimeters
      ' Printer.FontName = "draft 10cpi"
       Printer.FontName = "COURIER NEW"
       Printer.FontSize = 12
       Printer.ForeColor = "0"
       Printer.FontBold = False
        
        
       If adorsExtra1("pr_aliquotaipi") = 2 Or adorsExtra1("pr_aliquotaipi") = 4 Then
          wAliq = (adorsExtra1("pr_aliquotaipi") / 100) + 1
          wUnit = (adorsExtra1("pr_precovenda1") / wAliq)
       Else
          wUnit = adorsExtra1("pr_precocusto1")
       End If
                
       wVal = wUnit * adorsExtra1("ro_quantidadeenviada")
       wVal1 = wVal1 + wVal
       
       wvar = right$(Space(7) & adorsExtra1("ro_referencia"), 7) & "/" & CodBarras
       
       wvar1 = (left$(Format(adorsExtra1("pr_descricao"), ">") & Space(32), 32)) & Space(18) _
               & right$(Space(12) & Format(wVal, "####,###0.00"), 12) & " " & right$(Space(4) & adorsExtra1("ro_quantidadeenviada"), 4) _
               & Space(2) & "______"
       
       Printer.Print
       Printer.Print wvar
       Printer.Print wvar1
       'Debug.Print
       wlin = wlin + 2
       
       adorsExtra1.MoveNext
       
    Loop
    
    Call TotGeral
    
Finaliza:
    adorsExtra1.Close
    adorsExtra2.Close

End Sub



Private Sub Cabecalho()
    
    wlin = 7
    
    If wpag > 0 Then
       Printer.NewPage
    End If
    
    wpag = 1
    
  
   
       Printer.ScaleMode = vbMillimeters
     '  Printer.FontName = "draft 10cpi"
       Printer.FontName = "COURIER NEW"
       Printer.FontSize = 12
       Printer.ForeColor = "0"
       Printer.FontBold = False
   
   
    
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.Print "RELATÓRIO DE ROMANEIO DE TRANSFERÊNCIA CD" & Space(22) & "PAGINA:" & Space(4) & right$(Space(6) & Format(wpag), 6)
    Printer.Print
    Printer.Print "NR.ROMANEIO:" & Space(3) & right$(Space(6) & Format(adorsExtra1("ro_numeroromaneio")), 6) & Space(5) & adorsExtra1("ro_lojaorigem") & Space(5) & "- LOJA " & adorsExtra1("ro_lojadestino") & Space(6) & "IMPR:" & Space(2) & Now
    Printer.Print
    Printer.Print "REFERENCIA/CODIGO DE BARRAS"
    Printer.Print "DESCRICAO                                                   " & " VALOR  " & "QTDE" & Space(1) & "AJUSTE"
    Printer.Print
    Printer.CurrentY = Printer.CurrentY + 3
    Printer.CurrentX = 0
    Dim i As Long
    For i = 1 To 8
        Printer.Print "----------";
    Next
    Printer.Print
    'Debug.Print

End Sub

Private Sub cmdRet_Click()
   
    If cmdRet.Caption = "&Retorna" Then
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

Private Sub mskNumero_change(Index As Integer)
   
    If Val(mskNumero(Index).Text) > 0 Then
        VerFlag(Index) = 1
    Else
        VerFlag(Index) = 0
    End If
    
    Call CmpFlags(VerFlag)

End Sub

Sub CmpFlags(ByRef CmpVar)
   
    Dim cnt As Integer
    Dim OK As Boolean
    
    OK = True
    
    For cnt = 0 To UBound(CmpVar)
        If CmpVar(cnt) = 0 Then
           OK = False
           Exit For
        End If
    Next cnt
    
    If OK = True Then
       cmdOK.Enabled = True
       cmdOK.Default = True
    Else
       cmdOK.Enabled = False
       cmdOK.Default = False
    End If

End Sub

Sub TotGeral()

    If wlin > 60 Then
      Call Cabecalho
    End If
    
    If wVal1 > 0 Then
       wvar = Space(28) & " TOTAL GERAL    " & right$(Space(25) & Format(wVal1, "###,###,###0.00"), 15) & "    "
       Printer.Print
       Printer.Print wvar
    End If

End Sub

