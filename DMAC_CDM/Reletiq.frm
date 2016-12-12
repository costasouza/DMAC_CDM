VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form RelEtiq 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Relatório de Etiquetas"
   ClientHeight    =   4170
   ClientLeft      =   945
   ClientTop       =   1365
   ClientWidth     =   6930
   ControlBox      =   0   'False
   FillStyle       =   7  'Diagonal Cross
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   300
      TabIndex        =   5
      Top             =   2850
      Visible         =   0   'False
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1605
      Left            =   1695
      TabIndex        =   7
      Top             =   375
      Visible         =   0   'False
      Width           =   3600
      Begin VB.TextBox mskNf 
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
         Height          =   300
         Left            =   1770
         TabIndex        =   2
         Text            =   " "
         Top             =   705
         Width           =   1065
      End
      Begin VB.TextBox mskFornec 
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
         Height          =   300
         Left            =   1785
         TabIndex        =   1
         Top             =   270
         Width           =   1065
      End
      Begin VB.TextBox mskserie 
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
         Height          =   300
         Left            =   1770
         TabIndex        =   3
         Top             =   1155
         Width           =   1065
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         Caption         =   "Nota Fiscal"
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   780
         TabIndex        =   12
         Top             =   795
         Width           =   795
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "Fornecedor"
         ForeColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   825
         TabIndex        =   11
         Top             =   345
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         Caption         =   "Série"
         ForeColor       =   &H00E0E0E0&
         Height          =   180
         Left            =   1230
         TabIndex        =   10
         Top             =   1245
         Width           =   465
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   300
      ScaleHeight     =   45
      ScaleWidth      =   6405
      TabIndex        =   4
      Top             =   3450
      Width           =   6405
   End
   Begin VB.CheckBox chkNaoCodBarra 
      BackColor       =   &H00505050&
      Caption         =   "Não Imprime Código de Barras"
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   2340
      TabIndex        =   0
      Top             =   2295
      Width           =   2445
   End
   Begin CentroDeDistribuicao.chameleonButton cmdOK 
      Height          =   510
      Left            =   3900
      TabIndex        =   6
      Top             =   3600
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Imprime"
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
      MICON           =   "Reletiq.frx":0000
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
      TabIndex        =   8
      Top             =   3600
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
      MICON           =   "Reletiq.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1125
      Left            =   1725
      TabIndex        =   13
      Top             =   630
      Visible         =   0   'False
      Width           =   3555
      Begin VB.TextBox txtqtde 
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
         Height          =   300
         Left            =   1620
         TabIndex        =   15
         Top             =   690
         Width           =   1065
      End
      Begin VB.TextBox txtrefer 
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
         Height          =   300
         Left            =   1620
         TabIndex        =   14
         Text            =   " "
         Top             =   270
         Width           =   1065
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404040&
         Caption         =   "Quantidade"
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   630
         TabIndex        =   17
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404040&
         Caption         =   "Referência"
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   630
         TabIndex        =   16
         Top             =   330
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   1740
      TabIndex        =   18
      Top             =   1125
      Width           =   3555
      Begin VB.OptionButton optopcao 
         BackColor       =   &H00404040&
         Caption         =   "Por Referência"
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Index           =   1
         Left            =   2010
         TabIndex        =   20
         Top             =   270
         Width           =   1395
      End
      Begin VB.OptionButton optopcao 
         BackColor       =   &H00404040&
         Caption         =   "Por Nota Fiscal"
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Index           =   0
         Left            =   300
         TabIndex        =   19
         Top             =   270
         Width           =   1395
      End
   End
   Begin CentroDeDistribuicao.chameleonButton cmdvolta 
      Height          =   510
      Left            =   2490
      TabIndex        =   9
      Top             =   3600
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Voltar"
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
      MICON           =   "Reletiq.frx":0038
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
Attribute VB_Name = "RelEtiq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim ParaImpr As Boolean

Dim wMat1(4) As String
Dim wMat2(4) As String
Dim wMat3(4) As String

Dim wCont1 As Long
Dim wCont2 As Long
Dim wCont3 As Long
Dim wcont4 As Long
Dim I As Long
Dim i1 As Long
Dim flg As Long
Dim wAux As Long

Dim SalvaY As Double
Dim SalvaX As Double
Dim Linha As Long
Dim X As Integer
Dim restoDiv As Integer
Dim QtdeRef As Integer
Dim QtdeItem As Integer
Dim restoDiv2 As Integer

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
Dim bc_pattern$

Private Const EspacoVertical = 12.655
Private Const InicioPagina = 12.6
Private Const MargemEsquerda = 9
Private Const EspacoHorizontal = 47.625

'Flag para identificar qdo botado inicia pode ser habilitado
'Compara no change de cada campo
'Todos os flags deverao ter 1 para habilitar
Dim VerFlag(3) As Long

Dim wPrimeiraPagina As Boolean

Dim Impressora As Printer

Private Sub cmdOK_Click()

'For Each Impressora In Printers
'       If Impressora.DeviceName = "ETIQUETA" Then
'          Set Printer = Impressora
'           Exit For
'       End If
'Next

'Dim wImpressora As String
'   wImpressora = "ETIQUETA"
'   For Each Impressora In Printers
'       If UCase(Impressora.DeviceName) = UCase(wImpressora) Then
'          Set Printer = Impressora
'           Exit For
'       End If
'   Next


    If chkNaoCodBarra.Value = 0 Then
        ImprimeCodigoBarras
    Else
        ImprimeSemCodigoBarras
    End If

End Sub

Private Sub cmdRet_Click()
   
    If cmdRet.Caption = "&Retorna" Then
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
    
    Unload Me

End Sub

Private Sub cmdvolta_Click()
    Frame2.Visible = True
    Frame1.Visible = False
    Frame3.Visible = False
    mskFornec.Text = ""
    mskNf.Text = ""
    mskserie.Text = ""
    optopcao(0).Value = False
    optopcao(1).Value = False
    chkNaoCodBarra.Value = False
    ProgressBar1.Visible = False
End Sub

Private Sub mskFornec_Change()
   
    If mskFornec.Text <> " " Then
       VerFlag(0) = 1
    Else
       VerFlag(0) = 0
    End If

End Sub

Private Sub mskNF_Change()
   
    If mskNf.Text <> " " Then
       VerFlag(1) = 1
    Else
       VerFlag(1) = 0
    End If

End Sub

Private Sub mskSerie_Change()
    
    If mskserie.Text <> " " Then
       VerFlag(2) = 1
    Else
       VerFlag(2) = 0
    End If
    
    Call CmpFlags(VerFlag)

End Sub

Sub CmpFlags(ByRef CmpVar)
           
    cmdOK.Enabled = True
    cmdOK.Default = True

End Sub

Private Sub Form_Load()
    carregarPosicaoTela Me
End Sub

Private Sub optopcao_Click(Index As Integer)

    If optopcao(0).Value = True Then
       Frame1.Visible = True
       Frame3.Visible = False
       Frame2.Visible = False
       mskFornec.SetFocus
       ProgressBar1.Visible = False
    Else
       If optopcao(1).Value = True Then
          Frame1.Visible = False
          Frame2.Visible = False
          Frame3.Visible = True
          txtrefer.SetFocus
          ProgressBar1.Visible = True
       End If
    End If
    
End Sub

Private Sub EmiteReferencia()
    
   ' Status "Imprimindo..."
        
    sql = "select pr_referencia,pr_descricao " _
        & "from Produto " _
        & "where pr_referencia= '" & Trim(txtrefer.Text) & "' "
        
    rs.CursorLocation = adUseClient
    rs.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
    If rs.EOF Then
       MsgBox "Produto não cadastrado ", vbCritical, "Atenção"
       rs.Close
       Exit Sub
    Else
        
       flg = 0
       wCont1 = 1
       wCont3 = 0
       wcont4 = Val(txtQtde.Text)
       ProgressBar1.Max = Val(txtQtde.Text)
       ParaImpr = False
       Printer.ScaleMode = vbMillimeters
       Printer.FontName = "Arial"
       Printer.FontSize = 8
       Printer.ForeColor = "0"
       Printer.FontBold = False
       Linha = 0
       
       SalvaY = InicioPagina
       
       For i1 = 0 To wcont4 - 1
           flg = flg + 1
           ProgressBar1.Value = flg
           DoEvents
           If ParaImpr Then
              Printer.Print "***** INTERROMPIDO PELO USUÁRIO *****"
              Printer.EndDoc
              Exit Sub
           End If
            
           If wCont1 > 4 Then
              wCont1 = 1
              wCont3 = 0
              
              If Linha = 20 Then
                 Linha = 0
                 Printer.NewPage
              End If
              
              
              'SalvaY = SalvaY + (Linha * EspacoVertical)
              SalvaY = SalvaY + (Linha * EspacoVertical)
              
              If wPrimeiraPagina = True Then
                 SalvaY = 1
                 wPrimeiraPagina = False
              Else
                 SalvaY = SalvaY + (Linha * EspacoVertical)
              End If
              
              For I = 1 To 4
                  SalvaX = MargemEsquerda + ((I - 1) * EspacoHorizontal)
                  
                  ImprimeBarcode wMat1(I), SalvaX, SalvaY
                  
                  Printer.Print wMat2(I)
                  Printer.CurrentX = SalvaX
                  Printer.Print wMat3(I)
              Next I
              
              For I = 1 To 4
                  wMat1(I) = Space(20)
                  wMat2(I) = Space(20)
                  wMat3(I) = Space(20)
              Next I
              
              Linha = Linha + 1
              
           '   SalvaY = InicioPagina
               SalvaY = InicioPagina
               
           If wPrimeiraPagina = True Then
               SalvaY = 1
               wPrimeiraPagina = False
            Else
               SalvaY = InicioPagina
            End If
           
           End If
           
           wMat1(wCont1) = left$(rs("pr_referencia"), 7)
           wMat2(wCont1) = left$(rs("pr_descricao"), 16)
           wMat3(wCont1) = Mid$(rs("pr_descricao"), 17)
           
           wCont3 = 1
           wCont1 = wCont1 + 1
       Next
    End If
        
    If wCont3 = 1 Then
       
    
       ' SalvaY = SalvaY + (Linha * EspacoVertical)
         SalvaY = SalvaY + (Linha * EspacoVertical)
        
         If wPrimeiraPagina = True Then
            SalvaY = 1
            wPrimeiraPagina = False
         Else
            SalvaY = SalvaY + (Linha * EspacoVertical)
         End If
        
        For I = 1 To 4
            SalvaX = MargemEsquerda + ((I - 1) * EspacoHorizontal)
            ImprimeBarcode wMat1(I), SalvaX, SalvaY
            Printer.Print wMat2(I)
            Printer.CurrentX = SalvaX
            Printer.Print wMat3(I)
        Next I
    End If
    
    Printer.EndDoc
        
    rs.Close

End Sub

Sub ImprimeCodigoBarras()
  
    wPrimeiraPagina = False
    Screen.MousePointer = 11
    Frame1.Visible = False
    Label1.Visible = True
    ProgressBar1.Visible = True
    cmdRet.Caption = "&Cancelar"
    cmdOK.Visible = False
 
    Me.Refresh
    
   ' Status "Processando Etiquetas..."
        
    If optopcao(0).Value = True Then
       
       sql = "Select Count(ci_referencia) As TotReg " _
        & "From itemnfcompra " _
        & "where (ci_notafiscal = " & mskNf.Text & ") " _
        & "and (ci_fornecedor = " & Val(Trim(mskFornec.Text)) & ") " _
        & "and (ci_serie= '" & mskserie.Text & "')  "
        
        
       rs.CursorLocation = adUseClient
       rs.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
       If rs.EOF Then
          MsgBox "NAO FOI POSSIVEL ABRIR TABELA 1"
          GoTo Finaliza
       Else
          If rs("totreg") <> 0 Then
             ProgressBar1.Max = rs("totreg")
          Else
             MsgBox "NOTA FISCAL NAO CADASTRADA"
             GoTo Finaliza
             rs.Close
          End If
       End If
    
       ' Status "Imprimindo..."
                
        rs.Close
        
        sql = "select Itemnfcompra.*,pr_emiteetiqueta,pr_referencia,pr_descricao " _
            & "from Itemnfcompra,Produto " _
            & "where ( ci_referencia = pr_referencia) " _
            & "and   ( ci_notafiscal = " & mskNf.Text & " ) " _
            & "and   ( ci_fornecedor = " & Trim(mskFornec.Text) & " ) " _
            & "and   ( ci_serie = '" & mskserie.Text & "' ) " _
            & "order by ci_referencia"
        
        rs.CursorLocation = adUseClient
        rs.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
        If rs.EOF Then
           MsgBox "NAO FOI POSSÍVEL ABRIR TABELA 2"
           GoTo Finaliza
        End If
        
        'ROTINA PRINCIPAL
        flg = 0
        ParaImpr = False
        
        Printer.ScaleMode = vbMillimeters
        Printer.FontName = "Times New Roman"
        Printer.FontSize = 8
        Printer.ForeColor = "0"
        Printer.FontBold = False
        
        Linha = 0
         
       ' SalvaY = InicioPagina
        
        If wPrimeiraPagina = True Then
           SalvaY = 1
           wPrimeiraPagina = False
        Else
           SalvaY = InicioPagina
        End If
        
        
        Do While Not rs.EOF
            flg = flg + 1
            ProgressBar1.Value = flg
            DoEvents
            If ParaImpr Then
               Printer.Print "***** INTERROMPIDO PELO USUÁRIO *****"
               Printer.EndDoc
               GoTo Finaliza
            End If
            
            If rs("pr_emiteetiqueta") = "N" Or rs("pr_emiteetiqueta") = "B" Then
               GoTo Continua
            ElseIf rs("PR_EMITEETIQUETA") = "S" Or rs("PR_EMITEETIQUETA") = "T" Then
               wCont1 = wCont1 + 1
               wCont2 = wCont2 + 1
               If wCont1 > 4 Then
                  wCont1 = 1
                  wCont3 = 0
                  
                  If Linha = 20 Then
                     Linha = 0
                     Printer.NewPage
                  End If
                  
                  If wPrimeiraPagina = True Then
                     SalvaY = 1
                     wPrimeiraPagina = False
                  Else
                     SalvaY = SalvaY + (Linha * EspacoVertical)
                  End If
                    
                  For I = 1 To 4
                      SalvaX = MargemEsquerda + ((I - 1) * EspacoHorizontal)
                        
                      ImprimeBarcode wMat1(I), SalvaX, SalvaY
                        
                      Printer.Print wMat2(I)
                      Printer.CurrentX = SalvaX
                      Printer.Print wMat3(I)
                  Next I
                  
                  For I = 1 To 4
                      wMat1(I) = 0
                      wMat2(I) = 0
                      wMat3(I) = 0
                  Next I
                    
                  Linha = Linha + 1
                    
                 If wPrimeiraPagina = True Then
               SalvaY = 1
               wPrimeiraPagina = False
            Else
               SalvaY = InicioPagina
            End If
                  
               End If
               
               Do While rs("ci_quantidade") >= wCont2
                  If wCont1 > 4 Then
                     wCont1 = 1
                     wCont3 = 0
                     
                     If Linha = 20 Then
                        Linha = 0
                        Printer.NewPage
                     End If
                         
                     If wPrimeiraPagina = True Then
                        SalvaY = 1
                        wPrimeiraPagina = False
                     Else
                        SalvaY = SalvaY + (Linha * EspacoVertical)
                     End If
                          
                     For I = 1 To 4
                         SalvaX = MargemEsquerda + ((I - 1) * EspacoHorizontal)
                              
                         ImprimeBarcode wMat1(I), SalvaX, SalvaY
                              
                         Printer.Print wMat2(I)
                         Printer.CurrentX = SalvaX
                         Printer.Print wMat3(I)
                     Next I
                      
                     For I = 1 To 4
                         wMat1(I) = Space(20)
                         wMat2(I) = Space(20)
                         wMat3(I) = Space(20)
                     Next I
                    
                     Linha = Linha + 1
                      
                     SalvaY = InicioPagina
                  End If
                  
                  wMat1(wCont1) = left$(rs("pr_referencia"), 7)
                  wMat2(wCont1) = left$(rs("pr_descricao"), 16)
                  wMat3(wCont1) = Mid$(rs("pr_descricao"), 17)
                  
                  wCont3 = 1
                  wCont1 = wCont1 + 1
                  wCont2 = wCont2 + 1
               Loop
               
               wCont2 = 0
               wCont1 = wCont1 - 1
           End If
Continua:
            
           rs.MoveNext
        Loop
        
        If wCont3 = 1 Then
            If Linha = 20 Then
               Linha = 0
               Printer.NewPage
            End If
            
            If wPrimeiraPagina = True Then
               SalvaY = 1
               wPrimeiraPagina = False
            Else
               SalvaY = SalvaY + (Linha * EspacoVertical)
            End If
            
            For I = 1 To 4
                SalvaX = MargemEsquerda + ((I - 1) * EspacoHorizontal)
                  
                'cesar
                ImprimeBarcode wMat1(I), SalvaX, SalvaY
                  
                Printer.Print wMat2(I)
                Printer.CurrentX = SalvaX
                Printer.Print wMat3(I)
            Next I
        End If
        
        Printer.EndDoc
        
        rs.Close
    Else
        Call EmiteReferencia
    End If
        
Finaliza:
       flg = 0
       wAux = 0
       wCont1 = 0
       wCont2 = 0
       wCont3 = 0
       
       For I = 0 To 3
           wMat1(I) = ""
           wMat2(I) = ""
           wMat3(I) = ""
       Next
       
       chkNaoCodBarra.Value = False
       
       I = 0
       mskserie.Text = " "
       mskNf.Text = "      "
       mskFornec.Text = "   "
       txtrefer.Text = ""
       txtQtde.Text = ""
       Screen.MousePointer = 0
      ' Status "Pronto."
       ProgressBar1.Value = 0
       cmdOK.Visible = True
      
       Label1.Visible = True
       ProgressBar1.Visible = True
       cmdRet.Caption = "&Retorna"
       Frame1.Visible = False
       Frame3.Visible = False
       Frame2.Visible = True
       optopcao(0).Value = False
       optopcao(1).Value = False
       Me.Refresh


End Sub

Sub ImprimeSemCodigoBarras()

    
    Screen.MousePointer = 11
    Frame1.Visible = False
    Label1.Visible = True
    ProgressBar1.Visible = True
    cmdRet.Caption = "&Cancelar"
    cmdOK.Visible = False
    Me.Refresh
    
   ' Status "Processando Etiquetas..."
        
    If optopcao(0).Value = True Then
        EmiteEtiquetaSemCodigoBarrasNF
    Else
        Call EmiteEtiquetaSemCodigoBarra
    End If
        
Finaliza:
       flg = 0
       wAux = 0
       wCont1 = 0
       wCont2 = 0
       wCont3 = 0
       
       For I = 0 To 3
           wMat1(I) = ""
           wMat2(I) = ""
           wMat3(I) = ""
       Next
       
       I = 0
       mskserie.Text = " "
       mskNf.Text = "      "
       mskFornec.Text = "   "
       txtrefer.Text = ""
       txtQtde.Text = ""
       Screen.MousePointer = 0
       ProgressBar1.Value = 0
       cmdOK.Visible = True
       Label1.Visible = True
       ProgressBar1.Visible = True
       cmdRet.Caption = "&Retorna"
       Frame1.Visible = False
       Frame3.Visible = False
       Frame2.Visible = True
       optopcao(0).Value = False
       optopcao(1).Value = False
       Me.Refresh

End Sub
Sub EmiteEtiquetaSemCodigoBarra()
     
    sql = "select pr_referencia,pr_descricao " _
        & "from Produto " _
        & "where pr_referencia= '" & Trim(txtrefer.Text) & "' "
        
    rs.CursorLocation = adUseClient
    rs.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
    If rs.EOF Then
       MsgBox "Produto não cadastrado ", vbCritical, "Atenção"
       Exit Sub
    Else
        
       flg = 0
       wCont1 = 1
       wCont3 = 0
       wcont4 = Val(txtQtde.Text)
       ProgressBar1.Max = Val(txtQtde.Text)
       ParaImpr = False
       Printer.ScaleMode = vbMillimeters
       Printer.FontName = "Arial"
       Printer.FontSize = 8
       Printer.ForeColor = "0"
       Printer.FontBold = False
       Linha = 0
       
      X = 1
              
   
        Printer.Print
        Printer.Print
        Printer.Print
  
                
       SalvaY = InicioPagina
       
       restoDiv = wcont4 Mod 4
       
       For i1 = 0 To wcont4 - 1
           flg = flg + 1
           ProgressBar1.Value = flg
           DoEvents
           If ParaImpr Then
              Printer.Print "***** INTERROMPIDO PELO USUÁRIO *****"
              Printer.EndDoc
              Exit Sub
           End If
            
           If wCont1 > 4 Then
              wCont1 = 1
              wCont3 = 0
              
              If Linha = 20 Then
                 Linha = 0
                 Printer.NewPage
                       
                 X = 1
              '   For x = 1 To 4
              '      Printer.Print
              '   Next x
                 For X = 1 To 4
                    Printer.Print
                 Next X
                 
                  Printer.Print
                  Printer.Print
                  Printer.Print
              End If
              
             
                SalvaY = SalvaY + (Linha * EspacoVertical)
              
                Printer.ScaleMode = vbMillimeters
                Printer.FontName = "Arial"
                Printer.FontSize = 8
                Printer.ForeColor = "0"
                Printer.FontBold = False

              
              I = 1
              
                X = 1
                             
                Printer.Print Tab(8); wMat1(X); Tab(18); wMat2(X); Tab(45); wMat1(X + 1); Tab(55); wMat2(X + 1); Tab(83); wMat1(X + 2); Tab(93); wMat2(X + 2); Tab(120); wMat1(X + 3); Tab(130); wMat2(X + 3)
                Printer.Print Tab(8); wMat3(X); Tab(45); wMat3(X + 1); Tab(83); wMat3(X + 2); Tab(120); wMat3(X + 3)
                Printer.Print
                Printer.Print
                
            
                        
              For I = 1 To 4
                  wMat1(I) = Space(15)
                  wMat2(I) = Space(15)
                  wMat3(I) = Space(15)
              Next I
              
              Linha = Linha + 1
              
              SalvaY = InicioPagina
           End If
           
           wMat1(wCont1) = left$(rs("pr_referencia"), 7)
           wMat2(wCont1) = left$(rs("pr_descricao"), 16)
           wMat3(wCont1) = Mid$(rs("pr_descricao"), 17)
           
           wCont3 = 1
           wCont1 = wCont1 + 1
       Next
    End If
        
    If wCont3 = 1 Then
                X = 1
                
                If restoDiv = 1 Then
                    Printer.Print Tab(8); wMat1(X); Tab(18); wMat2(X)
                    Printer.Print Tab(8); wMat3(X)
                ElseIf restoDiv = 2 Then
                    Printer.Print Tab(8); wMat1(X); Tab(18); wMat2(X); Tab(45); wMat1(X + 1); Tab(55); wMat2(X + 1)
                    Printer.Print Tab(8); wMat3(X); Tab(45); wMat3(X + 1)
                ElseIf restoDiv = 3 Then
                    Printer.Print Tab(8); wMat1(X); Tab(18); wMat2(X); Tab(45); wMat1(X + 1); Tab(55); wMat2(X + 1); Tab(83); wMat1(X + 2); Tab(93); wMat2(X + 2)
                    Printer.Print Tab(8); wMat3(X); Tab(45); wMat3(X + 1); Tab(83); wMat3(X + 2)
                Else
                    Printer.Print Tab(8); wMat1(X); Tab(18); wMat2(X); Tab(45); wMat1(X + 1); Tab(55); wMat2(X + 1); Tab(83); wMat1(X + 2); Tab(93); wMat2(X + 2); Tab(120); wMat1(X + 3); Tab(130); wMat2(X + 3)
                    Printer.Print Tab(8); wMat3(X); Tab(45); wMat3(X + 1); Tab(83); wMat3(X + 2); Tab(120); wMat3(X + 3)
                End If
    End If
    
    chkNaoCodBarra.Value = False
    
    Printer.EndDoc
        
    rs.Close

End Sub
Private Sub EmiteEtiquetaSemCodigoBarrasNF()
 
      
    X = 1
    For X = 1 To 2
        sql = ""
        sql = "select Itemnfcompra.*,pr_emiteetiqueta,pr_referencia,pr_descricao " _
            & "from Itemnfcompra,Produto " _
            & "where ( ci_referencia = pr_referencia) " _
            & "and   ( ci_notafiscal = " & mskNf.Text & " ) " _
            & "and   ( ci_fornecedor = " & Trim(mskFornec.Text) & " ) " _
            & "and   ( ci_serie = '" & mskserie.Text & "' ) " _
            & "order by ci_referencia"
            
        If rs.State = 1 Then
            rs.Close
        End If
        
        
        rs.CursorLocation = adUseClient
        rs.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
         If X = 1 Then
            Do While Not rs.EOF
                QtdeRef = QtdeRef + rs("ci_quantidade")
                rs.MoveNext
            Loop
         End If
     Next X
        
        QtdeItem = rs("ci_QUANTIDADE")
        restoDiv2 = QtdeRef Mod 4
        
        
        If rs.EOF Then
           MsgBox "NAO FOI POSSÍVEL ABRIR TABELA 2"
        Else
        
                
       flg = 0
       wCont1 = 1
       wCont3 = 0
       wcont4 = QtdeRef
       'ProgressBar1.Max = Val(txtqtde.Text)
       'ProgressBar1.Max = QtdeRef
       ParaImpr = False
       Printer.ScaleMode = vbMillimeters
       Printer.FontName = "Arial"
       Printer.FontSize = 8
       Printer.ForeColor = "0"
       Printer.FontBold = False
       Linha = 0
                   
      X = 1
      
      For X = 2 To 4
        Printer.Print
        Printer.Print
        Printer.Print
   
      Next X
                

       SalvaY = InicioPagina
       
      ' restoDiv = wcont4 Mod 4
       
       For i1 = 0 To wcont4 - 1
           flg = flg + 1
           ProgressBar1.Value = flg
           DoEvents
           If ParaImpr Then
              Printer.Print "***** INTERROMPIDO PELO USUÁRIO *****"
              Printer.EndDoc
              Exit Sub
           End If
            
           If wCont1 > 4 Then
              wCont1 = 1
              wCont3 = 0
              
              If Linha = 20 Then
                 Linha = 0
                 Printer.NewPage
                       
                 X = 1
               
                 For X = 1 To 4
                    Printer.Print
                 Next X
            
                Printer.Print
                Printer.Print
                Printer.Print
            
              End If
              
              SalvaY = SalvaY + (Linha * EspacoVertical)
              
                Printer.FontName = "Arial"
                Printer.FontSize = 8
                Printer.ForeColor = "0"
                Printer.FontBold = False
              
                I = 1
                X = 1
                             
                Printer.Print Tab(8); wMat1(X); Tab(18); wMat2(X); Tab(45); wMat1(X + 1); Tab(55); wMat2(X + 1); Tab(83); wMat1(X + 2); Tab(93); wMat2(X + 2); Tab(120); wMat1(X + 3); Tab(130); wMat2(X + 3)
                Printer.Print Tab(8); wMat3(X); Tab(45); wMat3(X + 1); Tab(83); wMat3(X + 2); Tab(120); wMat3(X + 3)
                
                Printer.Print
                Printer.Print
                
                      
              For I = 1 To 4
                  wMat1(I) = Space(15)
                  wMat2(I) = Space(15)
                  wMat3(I) = Space(15)
              Next I
              
              Linha = Linha + 1
              
              SalvaY = InicioPagina
           End If
           
           
           'Do While QtdeItem <= 0
           QtdeItem = QtdeItem - 1
                      
                wMat1(wCont1) = left$(rs("pr_referencia"), 7)
                wMat2(wCont1) = left$(rs("pr_descricao"), 16)
                wMat3(wCont1) = Mid$(rs("pr_descricao"), 17)
           
           If QtdeItem = 0 Then
                rs.MoveNext
                If Not rs.EOF Then
                    QtdeItem = rs("ci_QUANTIDADE")
                End If
           End If
                       
           wCont3 = 1
           wCont1 = wCont1 + 1
       Next i1
    End If
    rs.Close
        
    If wCont3 = 1 Then
                X = 1
                
            
            If restoDiv = 1 Then
                    Printer.Print Tab(8); wMat1(X); Tab(18); wMat2(X)
                    Printer.Print Tab(8); wMat3(X)
                ElseIf restoDiv = 2 Then
                    Printer.Print Tab(8); wMat1(X); Tab(18); wMat2(X); Tab(45); wMat1(X + 1); Tab(55); wMat2(X + 1)
                    Printer.Print Tab(8); wMat3(X); Tab(45); wMat3(X + 1)
                ElseIf restoDiv = 3 Then
                    Printer.Print Tab(8); wMat1(X); Tab(18); wMat2(X); Tab(45); wMat1(X + 1); Tab(55); wMat2(X + 1); Tab(83); wMat1(X + 2); Tab(93); wMat2(X + 2)
                    Printer.Print Tab(8); wMat3(X); Tab(45); wMat3(X + 1); Tab(83); wMat3(X + 2)
                Else
                    Printer.Print Tab(8); wMat1(X); Tab(18); wMat2(X); Tab(45); wMat1(X + 1); Tab(55); wMat2(X + 1); Tab(83); wMat1(X + 2); Tab(93); wMat2(X + 2); Tab(120); wMat1(X + 3); Tab(130); wMat2(X + 3)
                    Printer.Print Tab(8); wMat3(X); Tab(45); wMat3(X + 1); Tab(83); wMat3(X + 2); Tab(120); wMat3(X + 3)
                End If
    
    End If
    
    chkNaoCodBarra.Value = False
    
    Printer.EndDoc
        
    

End Sub

Public Sub ImprimeSemBarcode(ByVal bc_string As String, Xref As Double, Yref As Double)
  
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


        RefPixelX = Printer.ScaleX(Xref, vbMillimeters, vbPixels)
        RefPixelY = Printer.ScaleY(Yref, vbMillimeters, vbPixels)
        
        'dimensions
        .ScaleMode = vbPixels
        

        
        dw = 7      'CInt(.ScaleHeight / 40)                    'space between bars
        'If dw < 1 Then dw = 1
        th = .TextHeight(bc_string)                     'text height
        tw = .TextWidth(bc_string)                      'text width
        
        new_string = Chr$(1) & bc_string & Chr$(2)      'add pre-amble, post-amble
        
        Y1 = .CurrentY '.ScaleTop
        Y2 = .CurrentY + 90 '.ScaleTop + .ScaleHeight - 1.5 * th

        
        xpos = RefPixelX

        Printer.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
        xpos = xpos + dw
        
        .CurrentX = RefPixelX
        .CurrentY = Y2 + 0.25 * th
        
        Printer.FontName = "Arial"
        Printer.FontSize = 6
        
        Printer.Print bc_string & "  ";
        
        .ScaleMode = SalvaEscala
    End With

End Sub



