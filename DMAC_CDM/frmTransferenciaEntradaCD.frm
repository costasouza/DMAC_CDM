VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmTransferenciaEntradaCD 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Transferência de Entrada do CD"
   ClientHeight    =   7470
   ClientLeft      =   1425
   ClientTop       =   1950
   ClientWidth     =   15195
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   15195
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "&H00C0C0C0&"
      Height          =   6090
      Left            =   11895
      TabIndex        =   11
      Top             =   120
      Width           =   3225
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   45
      ScaleHeight     =   45
      ScaleWidth      =   15120
      TabIndex        =   7
      Top             =   6780
      Width           =   15120
   End
   Begin VB.TextBox txtNotaFiscal 
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
      Left            =   105
      TabIndex        =   3
      Top             =   6330
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.TextBox txtSerie 
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
      Left            =   1185
      TabIndex        =   2
      Top             =   6330
      Visible         =   0   'False
      Width           =   990
   End
   Begin CentroDeDistribuicao.chameleonButton cmdOK 
      Height          =   510
      Left            =   12300
      TabIndex        =   5
      Top             =   6900
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
      MICON           =   "frmTransferenciaEntradaCD.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdRetornar 
      Height          =   510
      Left            =   13725
      TabIndex        =   6
      Top             =   6900
      Width           =   1410
      _ExtentX        =   2487
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
      MICON           =   "frmTransferenciaEntradaCD.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdTransfAbertas 
      Height          =   2910
      Left            =   75
      TabIndex        =   8
      Top             =   450
      Width           =   6900
      _cx             =   12171
      _cy             =   5133
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
      ForeColor       =   12632256
      BackColorFixed  =   0
      ForeColorFixed  =   16423203
      BackColorSel    =   16423203
      ForeColorSel    =   8388608
      BackColorBkg    =   5263440
      BackColorAlternate=   3947580
      GridColor       =   5263440
      GridColorFixed  =   8421504
      TreeColor       =   3947580
      FloodColor      =   5263440
      SheetBorder     =   3947580
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmTransferenciaEntradaCD.frx":0038
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   5
      MergeCompare    =   0
      AutoResize      =   -1  'True
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
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   5263440
      ForeColorFrozen =   4210752
      WallPaperAlignment=   9
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdTransfProcessadas 
      Height          =   2910
      Left            =   7035
      TabIndex        =   9
      Top             =   450
      Width           =   4770
      _cx             =   8414
      _cy             =   5133
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
      ForeColor       =   12632256
      BackColorFixed  =   0
      ForeColorFixed  =   16423203
      BackColorSel    =   16423203
      ForeColorSel    =   8388608
      BackColorBkg    =   5263440
      BackColorAlternate=   3947580
      GridColor       =   5263440
      GridColorFixed  =   8421504
      TreeColor       =   3947580
      FloodColor      =   5263440
      SheetBorder     =   3947580
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmTransferenciaEntradaCD.frx":00FC
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   5
      MergeCompare    =   0
      AutoResize      =   -1  'True
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
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   5263440
      ForeColorFrozen =   4210752
      WallPaperAlignment=   9
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdDados 
      Height          =   2775
      Left            =   75
      TabIndex        =   10
      Top             =   3420
      Width           =   11730
      _cx             =   20690
      _cy             =   4895
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
      ForeColor       =   12632256
      BackColorFixed  =   0
      ForeColorFixed  =   16423203
      BackColorSel    =   16423203
      ForeColorSel    =   8388608
      BackColorBkg    =   5263440
      BackColorAlternate=   3947580
      GridColor       =   5263440
      GridColorFixed  =   8421504
      TreeColor       =   3947580
      FloodColor      =   5263440
      SheetBorder     =   3947580
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmTransferenciaEntradaCD.frx":0186
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   5
      MergeCompare    =   0
      AutoResize      =   -1  'True
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
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   5263440
      ForeColorFrozen =   4210752
      WallPaperAlignment=   9
   End
   Begin VB.Label lblDataNota 
      BackColor       =   &H00505050&
      Caption         =   "Data"
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Left            =   2370
      TabIndex        =   4
      Top             =   6360
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00505050&
      Caption         =   "Transferências Processadas"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   7035
      TabIndex        =   1
      Top             =   120
      Width           =   2010
   End
   Begin VB.Label lblTrasnfAberto 
      AutoSize        =   -1  'True
      BackColor       =   &H00505050&
      Caption         =   "Transferência em Aberto"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   1740
   End
End
Attribute VB_Name = "frmTransferenciaEntradaCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim Conexao As New rdoConnection
Dim adoDados As New ADODB.Recordset
Dim adoDadosItens As New ADODB.Recordset
Dim rsMovEstoque As New ADODB.Recordset
Dim rsDataControle As New ADODB.Recordset
Dim ControleLoja As New ADODB.Recordset
Dim WBANCO As String
Dim sql As String
Dim ConverteQtde As Integer
Dim SomaTransferenciaEntrada As Integer
Dim CalculaCheckListEntrada As Long
Dim CalculaEstoqueFinal As Integer
Dim CalculaEstoque As Integer
Dim Loja As String
Dim wObservacao As String
Dim I As Integer


Private Sub cmdOK_Click()
    Dim LojaDestino As String
    
    If grdDados.Rows > 1 Then
        Screen.MousePointer = 11
        ProcessaNotaTransf NomeUsuario, SenhaUsuario
        cmdRetornar.SetFocus
        Call Limpa
        Screen.MousePointer = 0

    Else
        MsgBox "Não Exite Nota Para Processar", vbCritical, "Atenção"
    End If
    
End Sub


Private Sub cmdRetornar_Click()
   frmControleCD.lblNomeTelas.Caption = ""
   Unload Me

End Sub
 
 Private Sub Form_Load()

    On Error Resume Next
    
    frmTransferenciaEntradaCD.top = (Screen.Height - frmTransferenciaEntradaCD.Height) / 2
   frmTransferenciaEntradaCD.left = (Screen.Width - frmTransferenciaEntradaCD.Width) / 2
   
    carregarPosicaoTamanhoTela Me
    'JanelaTOP Me
    
    grdDados.ColWidth(6) = 0
    
    Screen.MousePointer = 11
    PreencheGrideNotaAberta

 Screen.MousePointer = 0

End Sub

Function AchaLoja() As String
    
    AchaLoja = "CD','MC85','CMC"
   
End Function

Private Sub grdTransfAbertas_click()
    Dim L As Integer
'
' perguntar se a linha do grid está em azul, se estiver proceder como nas processadas
'
    Screen.MousePointer = 11
    If grdTransfAbertas.CellBackColor = &HFFFF80 Then    'Azul
        PreenheGrideItens grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 1), grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 2), grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 0)
        For L = 1 To grdTransfProcessadas.Rows - 1
            grdTransfProcessadas.row = L
            If grdTransfProcessadas.CellBackColor = &HC0FFFF Then
                grdTransfProcessadas.col = 0
                grdTransfProcessadas.ColSel = 3
                grdTransfProcessadas.FillStyle = flexFillRepeat
                grdTransfProcessadas.CellBackColor = &H80000005  'BRANCO
                grdTransfProcessadas.FillStyle = flexFillSingle
            End If
        Next L
        For L = 1 To grdTransfAbertas.Rows - 1
            grdTransfAbertas.row = L
            If grdTransfAbertas.CellBackColor = &HC0FFFF Then
                grdTransfAbertas.col = 0
                grdTransfAbertas.ColSel = 4
                grdTransfAbertas.FillStyle = flexFillRepeat
                grdTransfAbertas.CellBackColor = &H80000005  'BRANCO
                grdTransfAbertas.FillStyle = flexFillSingle
            End If
        Next L
        cmdOK.Enabled = False
    Else
        PreenheGrideItens grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 1), grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 2), grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 0)
        PintaGrideNotaAbertas grdTransfAbertas.row
        If grdTransfAbertas.Rows > 1 Then
            txtNotaFiscal.Text = grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 1)
            txtSerie.Text = grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 2)
            lblDataNota.Caption = grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 4)
            Loja = grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 0)
        End If
        cmdOK.Enabled = True
    End If
    Screen.MousePointer = 0
    
End Sub

Private Sub grdTransfProcessadas_Click()

    If grdTransfProcessadas.Rows > 2 Then
        Screen.MousePointer = 11
        PreenheGrideItens grdTransfProcessadas.TextMatrix(grdTransfProcessadas.row, 1), grdTransfProcessadas.TextMatrix(grdTransfProcessadas.row, 2), grdTransfProcessadas.TextMatrix(grdTransfProcessadas.row, 0)
        PintaGrideNotaProcessadas grdTransfProcessadas.row
        cmdOK.Enabled = False
        Screen.MousePointer = 0
    End If

End Sub

Private Sub grdTransfProcessadas_DblClick()
     
    If SelecionaCapaNfVenda(adoDados, Format(DateAdd("d", -15, Date), "mm/dd/yyyy"), Date) = True Then
        grdTransfProcessadas.Rows = 1
        Do While Not adoDados.EOF
            wObservacao = IIf(IsNull(Mid(adoDados("VC_Observacao"), 1, 2)), "0", Mid(adoDados("VC_Observacao"), 1, 2))
            If wObservacao <> "0" Then
                PreencheGrideNotaProcessada
            End If
            adoDados.MoveNext
        Loop
    End If
    
End Sub

Private Sub txtNotaFiscal_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
        If IsNumeric(txtNotaFiscal.Text) = False Then
            txtNotaFiscal.SetFocus
            txtNotaFiscal.SelStart = 0
            txtNotaFiscal.SelLength = Len(txtNotaFiscal.Text)
        Else
            txtSerie.SetFocus
        End If
    End If
    
End Sub

Private Sub txtSerie_LostFocus()
    txtSerie.Text = UCase(txtSerie.Text)
End Sub

Sub Limpa()
    
    txtNotaFiscal.Text = ""
    txtSerie.Text = ""
    lblDataNota.Caption = ""
    lblDataNota.Visible = False
    grdDados.Rows = 1
    grdDados.Rows = 2
    cmdOK.Enabled = False
    
End Sub

Function MovimentoEstoqueCentral(ByVal Loja As String, ByVal data As String, ByVal Referencia As String)

    Dim adoVerEstoque As New ADODB.Recordset
    Dim adoVerMovimentacaoEstoque As New ADODB.Recordset

    sql = "Select ES_Estoque from Estoque " _
        & "where ES_Referencia = '" & Referencia & "' " _
        & "and ES_Loja='" & Trim(Loja) & "' "
        
        
    adoVerEstoque.CursorLocation = adUseClient
    adoVerEstoque.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
  
    If Not adoVerEstoque.EOF Then
        sql = ""
        sql = "Select ME_Referencia from MovimentacaoEstoque " _
            & "where ME_Referencia='" & Referencia & "' " _
            & "and ME_Loja='" & Loja & "' " _
            & "and ME_DataMovimento='" & Format(data, "mm/dd/yyyy") & "'"
            
            
        adoVerMovimentacaoEstoque.CursorLocation = adUseClient
        adoVerMovimentacaoEstoque.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    
        
        If adoVerMovimentacaoEstoque.EOF Then
            ADO_Cn_CD.BeginTrans
                sql = "Insert into MovimentacaoEstoque (ME_DataMovimento, ME_Loja, ME_Referencia, " _
                & "ME_EstoqueInicial, ME_Venda, ME_TransferenciaSaida, ME_RemessaSaida, ME_AjusteSaida, " _
                & "ME_DevolucaoCompras, ME_EntradaCompras, ME_TransferenciaEntrada, ME_AjusteEntrada, " _
                & "ME_DevolucaoVenda, ME_EstoqueFinal, ME_ChekListEntrada, ME_ChekListSaida, " _
                & "ME_Situacao, ME_MovimentoOK) Values ('" & Format(data, "mm/dd/yyyy") & "', " _
                & "'" & Loja & "', '" & grdDados.TextMatrix(I, 1) & "'," & adoVerEstoque("Es_Estoque") & "," _
                & "0,0,0,0,0,0," & grdDados.TextMatrix(I, 3) & ",0,0," & adoVerEstoque("Es_Estoque") + grdDados.TextMatrix(I, 3) & ",0,0,0 ,'S')"
  
                 
             
            ADO_Cn_CD.Execute (sql)
            ADO_Cn_CD.CommitTrans
            
           
        Else
           ADO_Cn_CD.BeginTrans
            sql = "Update MovimentacaoEstoque set ME_TransferenciaEntrada = ME_TransferenciaEntrada + " & grdDados.TextMatrix(I, 3) & ", " _
                & "ME_EstoqueFinal=ME_EstoqueFinal + " & grdDados.TextMatrix(I, 3) & " " _
                & "where ME_DataMovimento= '" & Format(data, "mm/dd/yyyy") & "' and ME_Loja= '" & Loja & "' and " _
                & "ME_Referencia='" & Referencia & "' "
           ADO_Cn_CD.Execute (sql)
            ADO_Cn_CD.CommitTrans
        
        End If
    End If

End Function



Function SelecionaCapaNfVenda(ByRef adoVar As ADODB.Recordset, ByVal DataInicio As String, ByVal DataFim As String) As Boolean

    sql = ""
    sql = "Select VC_DataEmissao,VC_TotalNota,VC_NotaFiscal,VC_Serie,VC_LojaOrigem,VC_Observacao,VC_LojaDestino from CapaNfVenda " _
        & "where VC_DataEmissao between '" & Format(DataInicio, "yyyy/mm/dd") & "' and '" & Format(DataFim, "yyyy/mm/dd") & "' " _
        & "and VC_TipoNota='T' and VC_LojaDestino in ('" & Trim(AchaLoja) & "') " _
        & "order by VC_LojaOrigem,VC_NotaFiscal"

    
    adoVar.CursorLocation = adUseClient
    adoVar.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    
    If Not adoVar.EOF Then
       SelecionaCapaNfVenda = True
    Else
       SelecionaCapaNfVenda = False
    End If

End Function

Function SelecionaItemNfVenda(ByRef adoVar As ADODB.Recordset, ByVal NotaFiscal As Double, ByVal serie As String, ByVal LojaOrigem As String) As Boolean
    
    sql = ""
    sql = "Select VI_NotaFiscal,VI_Serie,vi_referencia,vi_quantidade, " _
    & "vi_precounitario,Vi_numeroitem,vi_valormercadoria,VI_TipoAtualizaTransito,pr_descricao " _
    & "from itemnfvenda,produto " _
    & "Where  VI_LojaOrigem = '" & LojaOrigem & "' " _
    & "and vI_notafiscal=" & NotaFiscal & " " _
    & "and vi_serie='" & serie & "' " _
    & "and vi_referencia=pr_referencia and vi_tiponota = 'T' order by Vi_numeroitem"
    
        adoVar.CursorLocation = adUseClient
    adoVar.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    
    If Not adoVar.EOF Then
        SelecionaItemNfVenda = True
    Else
        SelecionaItemNfVenda = False
    End If

End Function

Function PreencheGrideNotaAberta()

    sql = ""
    sql = "Select CS_DataEstoque from ControleSup"
    
    rsDataControle.CursorLocation = adUseClient
    rsDataControle.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
       
'Format(DateAdd("d", -15, Date)
   ' If SelecionaCapaNfVenda(adoDados, rsDataControle("CS_DataEstoque"), Date) = True Then
    If SelecionaCapaNfVenda(adoDados, Format(DateAdd("d", -30, Date), "yyyy/mm/dd"), Format(Date, "yyyy/mm/dd")) = True Then
        grdTransfAbertas.Rows = 1
        grdTransfProcessadas.Rows = 1
        Do While Not adoDados.EOF
            wObservacao = IIf(IsNull(Mid(adoDados("VC_Observacao"), 1, 2)), "0", Mid(adoDados("VC_Observacao"), 1, 2))
            If wObservacao = "0" Then
                grdTransfAbertas.AddItem Trim(adoDados("VC_LojaOrigem")) & Chr(9) _
                    & adoDados("VC_NotaFiscal") & Chr(9) _
                    & adoDados("VC_Serie") & Chr(9) _
                    & UCase(adoDados("VC_LojaDestino")) & Chr(9) _
                    & Format(adoDados("VC_DataEmissao"), "dd/mm/yyyy") & Chr(9) _
                    & Format(adoDados("VC_TotalNota"), "##,###,###0.00")
            Else
              '  PreenchegrideNotaProcessada
            End If
            adoDados.MoveNext
        Loop
    End If
    
End Function

Function PreencheGrideNotaProcessada()

    grdTransfProcessadas.AddItem Trim(adoDados("VC_LojaOrigem")) & Chr(9) _
        & adoDados("VC_NotaFiscal") & Chr(9) _
        & adoDados("VC_Serie") & Chr(9) _
        & adoDados("VC_Observacao")
    
End Function

Function PreenheGrideItens(ByVal NotaFiscal As Double, ByVal serie As String, ByVal LojaOrigem As String)

    If SelecionaItemNfVenda(adoDadosItens, NotaFiscal, serie, LojaOrigem) = True Then
        grdDados.Rows = 1
        Do While Not adoDadosItens.EOF
            grdDados.AddItem adoDadosItens("Vi_numeroItem") & Chr(vbKeyTab) & adoDadosItens("Vi_Referencia") _
                & Chr(vbKeyTab) & adoDadosItens("Pr_Descricao") _
                & Chr(vbKeyTab) & adoDadosItens("Vi_Quantidade") _
                & Chr(vbKeyTab) & Format(adoDadosItens("VI_PrecoUnitario"), "##,###,###0.00") _
                & Chr(vbKeyTab) & Format(adoDadosItens("VI_ValorMercadoria"), "##,###,###0.00") _
                & Chr(vbKeyTab) & IIf(IsNull(adoDadosItens("VI_TipoAtualizaTransito")), 0, adoDadosItens("VI_TipoAtualizaTransito"))
            adoDadosItens.MoveNext
        Loop
    End If

End Function

Function PintaGrideNotaAbertas(ByVal Linha As Integer)
    Dim L As Integer
    
    For L = 1 To grdTransfProcessadas.Rows - 1
        grdTransfProcessadas.row = L
        If grdTransfProcessadas.CellBackColor = &HC0FFFF Then
            grdTransfProcessadas.col = 0
            grdTransfProcessadas.ColSel = 3
            grdTransfProcessadas.FillStyle = flexFillRepeat
            grdTransfProcessadas.CellBackColor = &H80000005  'BRANCO
            grdTransfProcessadas.FillStyle = flexFillSingle
        End If
    Next L
    If grdTransfAbertas.CellBackColor = &HC0FFFF Then
        grdTransfAbertas.col = 0
        grdTransfAbertas.ColSel = 5
        grdTransfAbertas.FillStyle = flexFillRepeat
        grdTransfAbertas.CellBackColor = &H80000005  'BRANCO
        grdTransfAbertas.FillStyle = flexFillSingle
        grdDados.Rows = 1
    Else
        For L = 1 To grdTransfAbertas.Rows - 1
            grdTransfAbertas.row = L
            If grdTransfAbertas.CellBackColor = &HC0FFFF Then
                grdTransfAbertas.col = 0
                grdTransfAbertas.ColSel = 5
                grdTransfAbertas.FillStyle = flexFillRepeat
                grdTransfAbertas.CellBackColor = &H80000005  'BRANCO
                grdTransfAbertas.FillStyle = flexFillSingle
            End If
        Next L
        
        grdTransfAbertas.row = Linha
        grdTransfAbertas.col = 0
        grdTransfAbertas.ColSel = 5
        grdTransfAbertas.FillStyle = flexFillRepeat
        grdTransfAbertas.CellBackColor = &HC0FFFF  'AMARELOGRID
        grdTransfAbertas.FillStyle = flexFillSingle
   End If
   
End Function

Sub PintaGrideNotaProcessadas(ByVal Linha As Integer)
    Dim L As Integer
    
    For L = 1 To grdTransfAbertas.Rows - 1
        grdTransfAbertas.row = L
        If grdTransfAbertas.CellBackColor = &HC0FFFF Then
            grdTransfAbertas.col = 0
            grdTransfAbertas.ColSel = 4
            grdTransfAbertas.FillStyle = flexFillRepeat
            grdTransfAbertas.CellBackColor = &H80000005  'BRANCO
            grdTransfAbertas.FillStyle = flexFillSingle
        End If
    Next L
    If grdTransfProcessadas.CellBackColor = &HC0FFFF Then
        grdTransfProcessadas.col = 0
        grdTransfProcessadas.ColSel = 3
        grdTransfProcessadas.FillStyle = flexFillRepeat
        grdTransfProcessadas.CellBackColor = &H80000005  'BRANCO
        grdTransfProcessadas.FillStyle = flexFillSingle
        grdDados.Rows = 1
    Else
        For L = 1 To grdTransfProcessadas.Rows - 1
            grdTransfProcessadas.row = L
            If grdTransfProcessadas.CellBackColor = &HC0FFFF Then
                grdTransfProcessadas.col = 0
                grdTransfProcessadas.ColSel = 3
                grdTransfProcessadas.FillStyle = flexFillRepeat
                grdTransfProcessadas.CellBackColor = &H80000005  'BRANCO
                grdTransfProcessadas.FillStyle = flexFillSingle
            End If
        Next L
        
        grdTransfProcessadas.row = Linha
        grdTransfProcessadas.col = 0
        grdTransfProcessadas.ColSel = 3
        grdTransfProcessadas.FillStyle = flexFillRepeat
        grdTransfProcessadas.CellBackColor = &HC0FFFF  'AMARELOGRID
        grdTransfProcessadas.FillStyle = flexFillSingle
   End If
   
End Sub

Function ProcessaNotaTransf(ByVal Usuario As String, ByVal Senha As String)
    
    Dim LojaDestino As String
    Dim wObservacao As String
    
    For I = 1 To grdDados.Rows - 1
        Screen.MousePointer = 11
        LojaDestino = grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 3)
            
        If grdDados.TextMatrix(I, 6) = 0 Then
            
            MovimentoEstoqueCentral LojaDestino, Date, grdDados.TextMatrix(I, 1)
            
            '
            '-----------------Atualiza Estoque e Transito Central
            '
            ADO_Cn_CD.BeginTrans
            sql = "Update Estoque set es_estoque=es_estoque + " & grdDados.TextMatrix(I, 3) & "," _
                & "es_transito=es_transito - " & grdDados.TextMatrix(I, 3) & "" _
                & " where es_referencia= '" & grdDados.TextMatrix(I, 1) & "' " _
                & " and es_loja= '" & LojaDestino & "'"
           ADO_Cn_CD.Execute (sql)
            ADO_Cn_CD.CommitTrans
            
           
            
            grdDados.TextMatrix(I, 6) = 1
            
            '
            '-----------------Atualiza Tipo de Atualizacao do item
            '
            sql = ""
            
            ADO_Cn_CD.BeginTrans
                sql = "Update ItemNfVenda set VI_TipoAtualizaTransito=1 " _
                    & "Where VI_NotaFiscal = " & Val(txtNotaFiscal.Text) & " and " _
                    & "VI_Serie = '" & txtSerie.Text & "' and VI_LojaOrigem = '" & Loja & "' and " _
                    & "VI_DataEmissao = '" & Format(lblDataNota.Caption, "mm/dd/yyyy") & "'" _
                    & "and VI_Referencia='" & grdDados.TextMatrix(I, 1) & "'"
            ADO_Cn_CD.Execute (sql)
            ADO_Cn_CD.CommitTrans
            
            
               
        
        
        End If
            
    
            sql = ""
            ADO_Cn_CD.BeginTrans
                sql = "Update ItemNfVenda set VI_TipoAtualizaTransito=2 " _
                    & "Where VI_NotaFiscal = " & Val(txtNotaFiscal.Text) & " and " _
                    & "VI_Serie = '" & txtSerie.Text & "' and VI_LojaOrigem = '" & Loja & "' and " _
                    & "VI_DataEmissao = '" & Format(lblDataNota.Caption, "mm/dd/yyyy") & "'" _
                    & "and VI_Referencia = '" & grdDados.TextMatrix(I, 1) & "'"
            ADO_Cn_CD.Execute (sql)
            ADO_Cn_CD.CommitTrans
            
            
               
        
    Next
    '
    '----------------------Atualizando Capa Central
    '
    wObservacao = "OK - " & UCase(Usuario) & " - " & Format(Date, "dd/mm/yy") & " - " & Format(time, "hh:mm:ss")
        
        
    ADO_Cn_CD.BeginTrans
        sql = "Update CapaNFVenda Set VC_Observacao = '" & wObservacao & "', " _
            & "VC_TipoAtualizaTransito= 2, VC_DataTransferenciaEntrada='" & Format(Date, "mm/dd/yyyy") & "' " _
            & "Where VC_NotaFiscal = " & Val(txtNotaFiscal.Text) & " and " _
            & "VC_Serie = '" & txtSerie.Text & "' and VC_LojaOrigem = '" & Loja & "' and " _
            & "VC_DataEmissao = '" & Format(lblDataNota.Caption, "mm/dd/yyyy") & "'"
    ADO_Cn_CD.Execute (sql)
    ADO_Cn_CD.CommitTrans
            
       
       
    MsgBox "Nota Processada com sucesso", vbInformation, "Sucesso"
    grdTransfAbertas.col = 0
    grdTransfAbertas.ColSel = 5
    grdTransfAbertas.FillStyle = flexFillRepeat
    grdTransfAbertas.CellBackColor = &HFFFF80     'Azul
    grdTransfAbertas.FillStyle = flexFillSingle
       
End Function



