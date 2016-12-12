VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form ManRoman 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Manutenção de Romaneio"
   ClientHeight    =   7350
   ClientLeft      =   735
   ClientTop       =   2685
   ClientWidth     =   15240
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtReferencia 
      BackColor       =   &H00A3A3A3&
      Height          =   345
      Left            =   990
      TabIndex        =   20
      Top             =   6270
      Width           =   1095
   End
   Begin VB.TextBox txtQuantidade 
      BackColor       =   &H00A3A3A3&
      Height          =   345
      Left            =   3270
      TabIndex        =   5
      Top             =   6270
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   30
      Left            =   30
      ScaleHeight     =   30
      ScaleWidth      =   15030
      TabIndex        =   15
      Top             =   6675
      Width           =   15030
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   30
      TabIndex        =   10
      Top             =   60
      Width           =   15030
      Begin MSMask.MaskEdBox mskDataFinal 
         Height          =   330
         Left            =   6420
         TabIndex        =   9
         Top             =   270
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   10724259
         ForeColor       =   4210752
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDataInicial 
         Height          =   330
         Left            =   4710
         TabIndex        =   8
         Top             =   240
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   10724259
         ForeColor       =   4210752
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00505050&
         BorderStyle     =   0  'None
         Caption         =   "Opções"
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   45
         TabIndex        =   13
         Top             =   60
         Width           =   3795
         Begin VB.OptionButton optPesquisa 
            BackColor       =   &H00505050&
            Caption         =   "Consulta"
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Index           =   1
            Left            =   1500
            TabIndex        =   6
            Top             =   330
            Width           =   1305
         End
         Begin VB.OptionButton optPesquisa 
            BackColor       =   &H00505050&
            Caption         =   "Manutenção"
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   0
            Top             =   330
            Value           =   -1  'True
            Width           =   1515
         End
         Begin VB.Label lblOpcoes 
            BackColor       =   &H00505050&
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
            Left            =   105
            TabIndex        =   16
            Top             =   15
            Width           =   705
         End
      End
      Begin VB.TextBox txtRomaneio 
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
         Left            =   4890
         TabIndex        =   7
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label LblPeriodo 
         BackColor       =   &H00404040&
         Caption         =   "Período"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   4050
         TabIndex        =   18
         Top             =   405
         Width           =   780
      End
      Begin VB.Label Lbla 
         BackColor       =   &H00404040&
         Caption         =   "a"
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   6150
         TabIndex        =   14
         Top             =   390
         Width           =   255
      End
      Begin VB.Label LblRomaneio 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Romaneio"
         ForeColor       =   &H00E0E0E0&
         Height          =   180
         Left            =   4065
         TabIndex        =   11
         Top             =   375
         Width           =   720
      End
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdManutencaoRomaneio 
      Height          =   5340
      Left            =   30
      TabIndex        =   17
      Top             =   855
      Width           =   15030
      _cx             =   26511
      _cy             =   9419
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   3158064
      ForeColor       =   -2147483643
      BackColorFixed  =   0
      ForeColorFixed  =   16423203
      BackColorSel    =   16423203
      ForeColorSel    =   -2147483634
      BackColorBkg    =   5263440
      BackColorAlternate=   3947580
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   23
      Cols            =   8
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Manroman.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   5
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      ComboSearch     =   0
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   16777215
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin CentroDeDistribuicao.chameleonButton cmdPesquisa 
      Height          =   510
      Index           =   1
      Left            =   10875
      TabIndex        =   2
      Top             =   6750
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Pesquisa"
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
      MICON           =   "Manroman.frx":0148
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdNovaPesquisa 
      Height          =   510
      Index           =   2
      Left            =   12285
      TabIndex        =   3
      Top             =   6750
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Nova Pesquisa"
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
      MICON           =   "Manroman.frx":0164
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
      Index           =   3
      Left            =   13680
      TabIndex        =   4
      Top             =   6750
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Retornar"
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
      MICON           =   "Manroman.frx":0180
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdImprimir 
      Height          =   510
      Index           =   0
      Left            =   9465
      TabIndex        =   1
      Top             =   6750
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Imprimir"
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
      MICON           =   "Manroman.frx":019C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label LblReferencia 
      BackColor       =   &H00505050&
      Caption         =   "Referencia"
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Left            =   105
      TabIndex        =   19
      Top             =   6375
      Width           =   795
   End
   Begin VB.Label lblQuantidade 
      BackColor       =   &H00505050&
      Caption         =   "Quantidade"
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   2310
      TabIndex        =   12
      Top             =   6375
      Width           =   885
   End
End
Attribute VB_Name = "ManRoman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    Dim sql As String
    
    Dim vWhere As String
    Dim vWhere2 As String
    
    Dim vSequencia As Integer
         
Private Sub Form_Load()
    
    '--- Determina Posicionamento ---
    'ManRoman.top = (Screen.Height - FrmCriaRomaneio.Height) / 2
    'ManRoman.left = (Screen.Width - FrmCriaRomaneio.Width) / 2

   carregarPosicaoTamanhoTela Me

'--- Mesclando linhas e Colunas ---
    grdManutencaoRomaneio.MergeRow(0) = True
    grdManutencaoRomaneio.MergeRow(1) = True
    
    grdManutencaoRomaneio.MergeCol(0) = True
    grdManutencaoRomaneio.MergeCol(1) = True
    grdManutencaoRomaneio.MergeCol(2) = True
    grdManutencaoRomaneio.MergeCol(3) = True
    grdManutencaoRomaneio.MergeCol(4) = True
    grdManutencaoRomaneio.MergeCol(5) = True
    grdManutencaoRomaneio.MergeCol(6) = True
    
    optPesquisa(0).Value = False
    cmdPesquisa(1).Enabled = False
    txtReferencia.Text = ""
    
    'JanelaTOP Me  'sair se clicar no retorna
    
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
    
    'Como vai Ficar a Impressão
       grdManutencaoRomaneio.BackColor = vbWhite
       grdManutencaoRomaneio.BackColorAlternate = vbWhite
       grdManutencaoRomaneio.BackColorBkg = vbWhite
       grdManutencaoRomaneio.BackColorFixed = vbWhite
       grdManutencaoRomaneio.BackColorFrozen = vbWhite
       grdManutencaoRomaneio.ForeColorFixed = vbBlack
       grdManutencaoRomaneio.ForeColor = &H80000006
      
       grdManutencaoRomaneio.ColWidth(0) = 900
       grdManutencaoRomaneio.ColWidth(1) = 900
       grdManutencaoRomaneio.ColWidth(2) = 1150
       grdManutencaoRomaneio.ColWidth(3) = 1000
       grdManutencaoRomaneio.ColWidth(4) = 1250
       grdManutencaoRomaneio.ColWidth(5) = 4400
       grdManutencaoRomaneio.ColWidth(6) = 1150
       
       grdManutencaoRomaneio.FontSize = 11
       
       
    grdManutencaoRomaneio.PrintGrid "Manutenção de Romaneio", True, 1, 300, 500
    
    'Cores Grd Original
       grdManutencaoRomaneio.BackColor = &H303030
       grdManutencaoRomaneio.BackColorAlternate = &H3C3C3C
       grdManutencaoRomaneio.BackColorBkg = &H505050
       grdManutencaoRomaneio.BackColorFixed = &H0
       grdManutencaoRomaneio.BackColorFrozen = vbWhite
       grdManutencaoRomaneio.ForeColorFixed = &HFA9923
       grdManutencaoRomaneio.ForeColor = &H80000005
       
       grdManutencaoRomaneio.ColWidth(0) = 1200
       grdManutencaoRomaneio.ColWidth(1) = 1200
       grdManutencaoRomaneio.ColWidth(2) = 1200
       grdManutencaoRomaneio.ColWidth(3) = 1200
       grdManutencaoRomaneio.ColWidth(4) = 1125
       grdManutencaoRomaneio.ColWidth(5) = 7605
       grdManutencaoRomaneio.ColWidth(6) = 1200
       
       grdManutencaoRomaneio.FontSize = 8
    
End Sub
  
Private Sub cmdNovaPesquisa_Click(Index As Integer) 'limpa campos

    optPesquisa(0).Value = False
    optPesquisa(1).Value = False
    grdManutencaoRomaneio.Rows = 2
    txtRomaneio.Text = ""
    mskDataInicial = "__/__/____"
    mskDataFinal = "__/__/____"
    txtQuantidade.Text = ""
    txtReferencia.Text = ""
    
End Sub

Private Sub cmdRetorna_Click(Index As Integer) 'retorna para tela inicial

     'frmControleCD.lblNomeTelas.Caption = ""
     Unload Me
     
End Sub

Private Sub cmdPesquisa_Click(Index As Integer)

    CarregaGrdManutencaoRomaneio
    
End Sub

Private Sub CarregaGrdManutencaoRomaneio()

    vWhere = ""
    vWhere2 = ""
    
       
        If optPesquisa(0).Value = True Then
           If txtRomaneio.Text = "" Then
              mensagemCampoObrigatorio "Romaneio"
              txtRomaneio.SetFocus
           Exit Sub
           Else
            vWhere = " and ro_numeroRomaneio = " & txtRomaneio.Text
           End If
        End If
        
        
        If optPesquisa(1).Value = True Then
           If mskDataInicial = "__/__/____" Or mskDataFinal = "__/__/____" Then
           mensagemCampoObrigatorio "Data"
           mskDataInicial.SetFocus
           Exit Sub
           Else
            vWhere2 = " and ro_DataSolicitacao between '" & Format(mskDataInicial.Text, "mm/dd/yyyy") _
            & "' and '" & Format(mskDataFinal.Text, "mm/dd/yyyy") & "'"
           End If
        End If
        
        
    sql = " select ro_lojaOrigem, ro_lojaDestino, ro_quantidadePedida, ro_quantidadeEnviada, ro_referencia, " _
            & " pr_referencia, pr_descricao , ro_numeroRomaneio, ro_DataSolicitacao, ro_Sequencia,lo_regiao " _
            & " from Romaneio, Produto, loja " _
            & " where ro_referencia = pr_referencia and lo_loja = ro_lojaDestino " _
            & vWhere & vWhere2 & " order by lo_Regiao,ro_numeroRomaneio"


    rs.CursorLocation = adUseServer
    rs.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockBatchOptimistic
      
    grdManutencaoRomaneio.Rows = 2

    Do While Not rs.EOF
    
        grdManutencaoRomaneio.AddItem rs("ro_lojaOrigem") & Chr(9) & rs("ro_lojaDestino") & Chr(9) & _
                            rs("ro_quantidadePedida") & Chr(9) & rs("ro_quantidadeEnviada") & Chr(9) & _
                            rs("pr_referencia") & Chr(9) & rs("pr_descricao") & Chr(9) & _
                            rs("ro_numeroRomaneio") & Chr(9) & rs("ro_sequencia")
    
        
        rs.MoveNext
    Loop
    
    rs.Close
    
End Sub

Private Sub grdManutencaoRomaneio_Click()

    txtReferencia.Text = grdManutencaoRomaneio.TextMatrix(grdManutencaoRomaneio.row, 4)
    
End Sub

Private Sub grdManutencaoRomaneio_KeyUp(KeyCode As Integer, Shift As Integer)

    txtReferencia.Text = grdManutencaoRomaneio.TextMatrix(grdManutencaoRomaneio.row, 4)
    
    
End Sub

Private Sub optPesquisa_Click(Index As Integer)

    If optPesquisa(0).Value = True Then
       
        LblPeriodo.Visible = False
        mskDataInicial.Visible = False
        Lbla.Visible = False
        mskDataFinal.Visible = False
        
        LblRomaneio.Visible = True
        txtRomaneio.Visible = True
        cmdPesquisa(1).Enabled = True
        
        lblReferencia.Visible = True
        txtReferencia.Visible = True
        lblQuantidade.Visible = True
        txtQuantidade.Visible = True
        
    End If
    
    If optPesquisa(1).Value = True Then
        
        LblRomaneio.Visible = False
        txtRomaneio.Visible = False
        
        LblPeriodo.Visible = True
        mskDataInicial.Visible = True
        Lbla.Visible = True
        mskDataFinal.Visible = True
        cmdPesquisa(1).Enabled = True
        
        lblReferencia.Visible = False
        txtReferencia.Visible = False
        lblQuantidade.Visible = False
        txtQuantidade.Visible = False
         
        
    End If
End Sub

Private Sub txtQuantidade_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        atualizaQuantidadePedida
        txtQuantidade.Text = ""
    
    Else
        If KeyAscii >= 58 And KeyAscii <= 126 Then
            MsgBox "No Campo Quantidade só valores numérico", vbInformation, "Atenção"
        Exit Sub
        End If
    End If
End Sub

Private Sub atualizaQuantidadePedida()
    
    vSequencia = grdManutencaoRomaneio.TextMatrix(grdManutencaoRomaneio.row, 7)
    
    If Val(txtQuantidade.Text) > Val(grdManutencaoRomaneio.TextMatrix(grdManutencaoRomaneio.row, 2)) Or _
       Val(txtQuantidade.Text) < 0 Then
       MsgBox "Quantidade Invalida", vbInformation, "Atenção"
       Exit Sub
    End If
    
    If Val(txtQuantidade.Text) = 0 Then
       MsgBox " Quantidade igual a zero a referencia será excluida do Romaneio", vbInformation, "Atenção"
        sql = "delete romaneio where ro_sequencia = " & vSequencia
       
        ADO_Cn_CDLocal.Execute sql
    Else
 
        sql = "update Romaneio set ro_quantidadeEnviada = " & txtQuantidade & _
             " where ro_Sequencia = " & vSequencia

        ADO_Cn_CDLocal.Execute sql
        
   End If
   CarregaGrdManutencaoRomaneio
End Sub






