VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmConsRegiao 
   Appearance      =   0  'Flat
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Consulta de Frete por Região"
   ClientHeight    =   7800
   ClientLeft      =   1185
   ClientTop       =   2310
   ClientWidth     =   17235
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   17235
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbAno 
      Height          =   315
      Left            =   10350
      TabIndex        =   9
      Top             =   9960
      Width           =   1080
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   44
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   14880
      TabIndex        =   8
      Top             =   6810
      Width           =   14880
   End
   Begin VB.PictureBox PtbSistema 
      Height          =   555
      Left            =   960
      ScaleHeight     =   495
      ScaleWidth      =   8850
      TabIndex        =   6
      Top             =   9720
      Width           =   8910
      Begin VB.CommandButton cmdConsultar1 
         Caption         =   "Pesquisa"
         Height          =   495
         Left            =   6450
         TabIndex        =   1
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdSair1 
         Caption         =   "Sair"
         Height          =   495
         Left            =   7650
         TabIndex        =   7
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.Frame fraPesquisa 
      BackColor       =   &H00404040&
      Caption         =   "Pesquisa"
      ForeColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   240
      TabIndex        =   4
      Top             =   10680
      Width           =   14865
      Begin VB.ComboBox cmbLojas 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1215
         TabIndex        =   0
         Top             =   285
         Width           =   1230
      End
      Begin VB.Label lblloja 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Loja:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   825
         TabIndex        =   5
         Top             =   375
         Width           =   345
      End
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdVendaRegiao 
      Height          =   6165
      Left            =   7740
      TabIndex        =   3
      Top             =   450
      Width           =   7290
      _cx             =   12859
      _cy             =   10874
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      ForeColorSel    =   -2147483634
      BackColorBkg    =   5263440
      BackColorAlternate=   3947580
      GridColor       =   5263440
      GridColorFixed  =   5263440
      TreeColor       =   -2147483632
      FloodColor      =   5263440
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmConsRegiao.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
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
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdRegiao 
      Height          =   6165
      Left            =   150
      TabIndex        =   2
      Top             =   450
      Width           =   7395
      _cx             =   13044
      _cy             =   10874
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      ForeColorSel    =   -2147483634
      BackColorBkg    =   5263440
      BackColorAlternate=   3947580
      GridColor       =   5263440
      GridColorFixed  =   5263440
      TreeColor       =   -2147483632
      FloodColor      =   5263440
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmConsRegiao.frx":0071
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
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
      Editable        =   1
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
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin CentroDeDistribuicao.chameleonButton cmdRetorna 
      Height          =   510
      Left            =   13620
      TabIndex        =   11
      Top             =   6945
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
      MICON           =   "frmConsRegiao.frx":00E5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblPesquisa 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Pesquisa"
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
      Left            =   150
      TabIndex        =   12
      Top             =   150
      Width           =   855
   End
   Begin VB.Label lblAno 
      AutoSize        =   -1  'True
      Caption         =   "Ano:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   9960
      TabIndex        =   10
      Top             =   10050
      Width           =   330
   End
End
Attribute VB_Name = "frmConsRegiao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rdoConnBatch As New rdoConnection
Dim rsValor As New ADODB.Recordset
Dim rsDados As New ADODB.Recordset
Dim rdoConn As New ADODB.Connection
Dim rdoConnReg As New ADODB.Connection
Dim rsRegiao As New ADODB.Recordset
Dim rsCodigoRegiao As New ADODB.Recordset
Dim rdoConnCod As New ADODB.Connection
Dim sql As String
Dim SQL2 As String
Dim zona As String
Dim cor As String
Dim cor1 As String
Dim cor2 As String
Dim rsLoja As New ADODB.Recordset
Dim rdopintagrid As New ADODB.Recordset
Dim rdoConnLoja As New ADODB.Connection
Dim CodigoRegiao As String
Dim rsReg As New ADODB.Recordset
Dim Registro As Integer
Dim wcontrolacor As Integer
Dim wcoluna As Integer
Dim registroZona As Integer
Dim Valor As Double
Dim AC_Vendas As Double
Dim AC_Quantidade As Double
Dim ValorZona As Double
Dim LojaOrigem As String
Dim rdoConnRegiao As New ADODB.Connection
Dim rsNomeReg As New ADODB.Recordset
Dim Ano As String
Dim rsAno As New ADODB.Recordset
Dim Reg As String
Dim qnt As String
Dim VLR As String
Dim AnoReg As String
Dim anox As String
Dim ano_qnt As String
Dim ano_vlr As String
Dim rsCodRegiao As New ADODB.Recordset
Dim rsTotValor As New ADODB.Recordset
Dim CodRegiao As String
Dim i As Integer
Dim rsGrid As New ADODB.Recordset

Private Sub cmdConsultar_click()

End Sub

Private Sub CarregaOutrosAnos()
    Dim Lojas As String
    Dim TotFrete As Double
    
    
    If cmbLojas.Text = "Todas" Then
        Lojas = "RG_Loja not in ('Todas','Centro') "
    ElseIf cmbLojas.Text = "Centro" Then
        Lojas = "RG_Loja in ('28','316','48','85') "
    Else
        Lojas = "RG_Loja = '" & cmbLojas.Text & "'"
    End If
    
    grdRegiao.Rows = 1
    TotFrete = 0
    
     
'---------------------------------------------------------------------------------------------------------------------------
    'ricardo
        sql = ""
        sql = "Select RV_Codigoregiao,RV_Nomeregiao,sum(Rv_ValorFrete) as Rv_ValorFrete" _
            & " from RegiaoFrete  Where RV_CodigoRegiao = RV_CodigoRegiao" _
            & " group by RV_CodigoRegiao,RV_NomeRegiao,Rv_ValorFrete "
        
    
    rsDados.CursorLocation = adUseClient
    rsDados.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    
    Do While Not rsDados.EOF
     
        TotFrete = TotFrete + rsDados("Rv_ValorFrete")
        grdRegiao.AddItem Trim(rsDados("rv_nomeregiao")) & vbTab & rsDados("rv_codigoregiao") & vbTab & rsDados("Rv_ValorFrete")
    
        rsDados.MoveNext
    Loop
    grdRegiao.AddItem " "
    'grdRegiao.AddItem " T    O    T    A    L" & vbTab & " " & vbTab & Format(AC_Quantidade, "#,##0.00") & vbTab & Format(AC_Vendas, "#,##0.00")
    'grdRegiao.AddItem " T    O    T    A    L" & vbTab & " " & vbTab & Format(TotFrete, "#,##0.00")
    

    rsDados.Close
    
End Sub

Private Sub cmdSair_Click()

    Unload Me

End Sub

Private Sub cmdConsulta_Click()
  
End Sub

Private Sub cmdRetorna_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 11
    Call CarregaOutrosAnos
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
'    frmConsRegiao.top = (Screen.Height - frmConsRegiao.Height) / 2
'    frmConsRegiao.left = (Screen.Width - frmConsRegiao.Width) / 2
    
      carregarPosicaoTamanhoTela Me
    

'----------------------------------------------------------------------------------------------------------------------------------
'ricardo
        
         
''    SQL = ""
''    SQL = "select CTS_lOJA from ControleSistema"
''
''        cmbLojas.CursorLocation = adUseClient
''        cmbLojas.Open SQL, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
''
''    Do While Not rsLoja.EOF
''
''        cmbLojas.AddItem Trim(rsLoja("CTS_lOJA"))
''        rsLoja.MoveNext
''    Loop
''    rsLoja.Close


'-----------------------------------------------------------------------------------------------------------------------------------
'O CERTO

    sql = ""
    sql = "select * from loja where lo_situacao = 'A' and lo_loja not in ('181','182','183','184','185','314','CMC','conso','CD')"

    rsLoja.CursorLocation = adUseClient
    rsLoja.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic

    Do While Not rsLoja.EOF
        cmbLojas.AddItem Trim(rsLoja("Lo_Loja"))
        rsLoja.MoveNext
    Loop
    cmbLojas.AddItem "Centro"
    cmbLojas.AddItem "Todas"
    rsLoja.Close


'------------------------------------------------------------------------------------------------------------------------------------


End Sub

Private Sub grdRegiao_AfterEdit(ByVal row As Long, ByVal col As Long)

   'ricardo
     If col = 2 Then
    
        desc = Trim(grdRegiao.TextMatrix(grdRegiao.row, 2))
        
        If MsgBox("Confirma a Alteração do valor frete?", vbYesNo + vbQuestion + vbDefaultButton2, "Alterar valor frete") = vbYes Then
            gravaDesc
        Else
            grdRegiao.TextMatrix(grdRegiao.row, 2) = descAnte
        End If
        
    End If


End Sub
Private Sub gravaDesc()
    'ricardo
    Dim sql As String
    Dim codigo As String
    
    codigo = Trim(grdRegiao.TextMatrix(grdRegiao.row, 2))
   
    
    sql = "Update RegiaoFrete set rv_valorFrete = '" & codigo & "' " _
           & " where rv_CodigoRegiao = '" & Trim(grdRegiao.TextMatrix(grdRegiao.row, 1)) & "' "
    
     ADO_Cn_CDLocal.Execute (sql)
     
    sql = "Update RegiaoLocalFrete set Rg_valorFrete = '" & codigo & "' " _
          & " where rg_Regiao = '" & Trim(grdRegiao.TextMatrix(grdRegiao.row, 1)) & "' "
          
    
    ADO_Cn_CDLocal.Execute (sql)
    Call CarregaOutrosAnos
    Call grdRegiao_Click
    
End Sub

Private Sub grdRegiao_Click()

    Screen.MousePointer = 11

    grdVendaRegiao.Rows = 1
    
    CodRegiao = Trim(grdRegiao.TextMatrix(grdRegiao.row, 1))
    AC_Quantidade = 0
    AC_Vendas = 0
    
'-------------------------------------------------------------------------------------------------------------
    'RICARDO
    
    'sql = "SELECT RegiaoFrete.*,RegiaoLocalFrete.* From RegiaoFrete, RegiaoLocalFrete "
          '& "Where RV_CodigoRegiao = RG_Regiao and RG_Regiao='" & CodRegiao & "' group by RegiaoFrete.*,RegiaoLocalFrete.*  Order by RG_ValorFrete Desc"
    sql = ""
    sql = "SELECT  RG_Regiao,RG_CEPRua,RG_ValorFrete  From RegiaoFrete , RegiaoLocalFrete  Where RV_CodigoRegiao = RG_Regiao and RG_Regiao='" & CodRegiao & "' group by RG_Regiao,RG_CEPRua,RG_ValorFrete"


    rsGrid.CursorLocation = adUseClient
    rsGrid.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    Do While Not rsGrid.EOF
       
       grdVendaRegiao.AddItem Trim(rsGrid("RG_Regiao") & vbTab & Trim(rsGrid("RG_CEPRua") & vbTab & Format(rsGrid("RG_ValorFrete"), "###,###,##0.00")))
       rsGrid.MoveNext
    Loop
    'grdVendaRegiao.AddItem " "
    'grdVendaRegiao.AddItem " T    O    T    A    L" & vbTab & AC_Quantidade & vbTab & Format(AC_Vendas, "###,###,##0.00")

    rsGrid.Close
    
    Screen.MousePointer = 0
    
End Sub

Sub PintaGrid(ByRef NomeGrid)
    If wcontrolacor = 1 Then
       cor = cor1
       wcontrolacor = 2
    Else
       cor = cor2
       wcontrolacor = 1
    End If
     NomeGrid.row = NomeGrid.Rows - 1
     NomeGrid.col = 0
     
     NomeGrid.FillStyle = flexFillRepeat
     
     NomeGrid.FillStyle = flexFillSingle
   
   If Not wcoluna = 3 Then
          wcoluna = 2
   End If
End Sub

Sub limpaGrid(ByRef GradeUsu)
    
    GradeUsu.Rows = GradeUsu.FixedRows + 1
    GradeUsu.AddItem ""
    GradeUsu.RemoveItem GradeUsu.FixedRows

End Sub





