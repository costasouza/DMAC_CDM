VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmCriaSugestao 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Cria Sugestão"
   ClientHeight    =   7710
   ClientLeft      =   1260
   ClientTop       =   1875
   ClientWidth     =   15240
   DrawStyle       =   1  'Dash
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   150
      TabIndex        =   13
      Top             =   3975
      Width           =   9510
      Begin VB.TextBox txtReferencia 
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
         Left            =   150
         TabIndex        =   18
         ToolTipText     =   "Referência"
         Top             =   390
         Width           =   975
      End
      Begin VB.TextBox txtDescricao 
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
         Left            =   1200
         TabIndex        =   17
         ToolTipText     =   "Descrição"
         Top             =   390
         Width           =   4710
      End
      Begin VB.ComboBox cmbLojaOrigem 
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
         Left            =   5985
         TabIndex        =   16
         ToolTipText     =   "Loja Origem"
         Top             =   390
         Width           =   1095
      End
      Begin VB.ComboBox cmbLojaDestino 
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
         Left            =   7155
         TabIndex        =   15
         ToolTipText     =   "Loja Destino"
         Top             =   390
         Width           =   1095
      End
      Begin VB.TextBox txtQuantidade 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00A3A3A3&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
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
         Left            =   8325
         TabIndex        =   14
         ToolTipText     =   "Quantidade"
         Top             =   390
         Width           =   975
      End
      Begin VB.Label lblReferencia 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Referência"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   150
         TabIndex        =   23
         Top             =   150
         Width           =   780
      End
      Begin VB.Label lblDescricao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   1200
         TabIndex        =   22
         Top             =   150
         Width           =   720
      End
      Begin VB.Label lblLojaOrigem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loja Origem"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   5985
         TabIndex        =   21
         Top             =   150
         Width           =   840
      End
      Begin VB.Label lblLojaDestino 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loja Destino"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   7155
         TabIndex        =   20
         Top             =   150
         Width           =   885
      End
      Begin VB.Label lblQuantidade 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   8325
         TabIndex        =   19
         Top             =   150
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   150
      TabIndex        =   5
      Top             =   150
      Width           =   9510
      Begin VB.OptionButton optReferencia 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Referência"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   150
         TabIndex        =   12
         ToolTipText     =   "Referência"
         Top             =   150
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton optFornecedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Fornecedor"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   1395
         TabIndex        =   11
         ToolTipText     =   "Fornecedor"
         Top             =   150
         Width           =   1335
      End
      Begin VB.OptionButton optLinha 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Linha"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   2760
         TabIndex        =   10
         ToolTipText     =   "Linha"
         Top             =   150
         Width           =   855
      End
      Begin VB.OptionButton opt180 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "+180"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   3600
         TabIndex        =   9
         ToolTipText     =   "180"
         Top             =   150
         Width           =   855
      End
      Begin VB.OptionButton optEstoqueAbaixo 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Estoque Abaixo do Mínimo"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   6870
         TabIndex        =   8
         ToolTipText     =   "Estoque Abaixo do Mínimo Informado"
         Top             =   150
         Width           =   2655
      End
      Begin VB.OptionButton optEstoqueAcima 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Estoque Acima do Máximo"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   4440
         TabIndex        =   7
         ToolTipText     =   "Estoque Acima do Máximo Informado"
         Top             =   150
         Width           =   2655
      End
      Begin VB.TextBox txtPesquisa 
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
         Left            =   150
         TabIndex        =   6
         ToolTipText     =   "Pesquisa"
         Top             =   450
         Width           =   9195
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   44
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   14880
      TabIndex        =   0
      Top             =   6810
      Width           =   14880
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdPesquisa 
      Height          =   2535
      Left            =   150
      TabIndex        =   1
      Top             =   1275
      Width           =   9510
      _cx             =   16775
      _cy             =   4471
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
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   5263440
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
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmCriaSugestao.frx":0000
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
   Begin VSFlex7DAOCtl.VSFlexGrid grdSugestaoAberta 
      Height          =   1590
      Left            =   150
      TabIndex        =   2
      Top             =   5040
      Width           =   9510
      _cx             =   16775
      _cy             =   2805
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
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   5263440
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
      Rows            =   1
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmCriaSugestao.frx":0054
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
   Begin VSFlex7DAOCtl.VSFlexGrid grdEstoque 
      Height          =   6480
      Left            =   9840
      TabIndex        =   3
      Top             =   150
      Width           =   5175
      _cx             =   9128
      _cy             =   11430
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
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   5263440
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
      Rows            =   1
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmCriaSugestao.frx":011D
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
   Begin CentroDeDistribuicao.chameleonButton cmdRetorna 
      Height          =   510
      Left            =   13620
      TabIndex        =   4
      ToolTipText     =   "Retorna"
      Top             =   6945
      Width           =   1410
      _extentx        =   2487
      _extenty        =   900
      btype           =   14
      tx              =   "Retorna"
      enab            =   -1  'True
      font            =   "frmCriaSugestao.frx":020A
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   5263440
      bcolo           =   0
      fcol            =   16423203
      fcolo           =   16423203
      mcol            =   5263440
      mptr            =   1
      micon           =   "frmCriaSugestao.frx":0236
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdPesquisa 
      Height          =   510
      Left            =   12120
      TabIndex        =   24
      ToolTipText     =   "Retorna"
      Top             =   6945
      Width           =   1410
      _extentx        =   2487
      _extenty        =   900
      btype           =   14
      tx              =   "Retorna"
      enab            =   -1  'True
      font            =   "frmCriaSugestao.frx":0254
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   5263440
      bcolo           =   0
      fcol            =   16423203
      fcolo           =   16423203
      mcol            =   5263440
      mptr            =   1
      micon           =   "frmCriaSugestao.frx":0280
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
End
Attribute VB_Name = "frmCriaSugestao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdPesquisa_Click()
    pesquisar
End Sub

Private Sub cmdRetorna_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Dim adoLojas As New ADODB.Recordset
    Dim sql As String
    
    carregarPosicaoTamanhoTela Me
    
    sql = "sp_cdm_busca_sugestao 10,'','','','','','','','','',''"
    
    adoLojas.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not adoLojas.EOF Then
        cmbLojaDestino.Text = adoLojas("LO_loja")
        cmbLojaOrigem.Text = adoLojas("LO_loja")
        
        Do While Not adoLojas.EOF
        
            cmbLojaDestino.AddItem adoLojas("LO_loja")
            cmbLojaOrigem.AddItem adoLojas("LO_loja")
            adoLojas.MoveNext
        Loop
    End If
    
    adoLojas.Close
    cmdPesquisa.Caption = "Pesquisar"
End Sub

Private Sub grdPesquisa_DblClick()
    txtReferencia = grdPesquisa.TextMatrix(grdPesquisa.row, 0)
    txtDescricao = grdPesquisa.TextMatrix(grdPesquisa.row, 1)
    pesquisaEstoque
End Sub

Private Sub opt180_Click()
    txtPesquisa.SetFocus
End Sub

Private Sub optEstoqueAbaixo_Click()
    txtPesquisa.SetFocus
End Sub

Private Sub optEstoqueAcima_Click()
    txtPesquisa.SetFocus
End Sub

Private Sub optFornecedor_Click()
    txtPesquisa.MaxLength = 3
    txtPesquisa.SetFocus
End Sub

Private Sub optLinha_Click()
    txtPesquisa.SetFocus
End Sub

Private Sub optReferencia_Click()
    txtPesquisa.MaxLength = 7
    txtPesquisa.SetFocus
    
End Sub


Private Sub pesquisar()

    Dim sql As String
    Dim Data1 As Date
    Dim Data2 As Date
    Dim adoPesquisa As New ADODB.Recordset
    
    
    'Limpa Grid Pesquisa
    grdPesquisa.Rows = 1
    grdPesquisa.AddItem ""
    grdPesquisa.RemoveItem (1)
    
    If optReferencia.Value Then
    
        sql = "EXEC SP_CDM_Busca_Sugestao 1,'','" & txtPesquisa.Text & "','','','','','','','',''"
    
    ElseIf optFornecedor.Value Then
        
        sql = "EXEC SP_CDM_Busca_Sugestao 2,'" & txtPesquisa.Text & "','','','','','','','','',''"
    
    ElseIf optLinha.Value Then
        
        sql = "EXEC SP_CDM_Busca_Sugestao 3, '','','" & txtPesquisa.Text & "','','','','','','',''"
    
    ElseIf opt180.Value Then
    
        Data1 = "01/" & Mid(Date, 4, 7)
        Data2 = DateAdd("d", -1, Data1)
        Data1 = "01/" & Mid(DateAdd("m", -1, Data1), 4, 7)
        
        sql = "EXEC SP_CDM_Busca_Sugestao 4,'','','','" & Format(Data1, "yyyy/mm/dd") & "','" & Format(Data2, "yyyy/mm/dd") & "','','','','',''"
    
    ElseIf optEstoqueAbaixo.Value Then
        
        sql = "EXEC SP_CDM_Busca_Sugestao 5,'','','','','','','','','',''"
        
    ElseIf optEstoqueAcima.Value Then
    
        sql = "EXEC SP_CDM_Busca_Sugestao 6,'','','','','','','','','',''"
    
    End If
    
    adoPesquisa.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
     
    If Not adoPesquisa.EOF Then
        Do While Not adoPesquisa.EOF
            grdPesquisa.AddItem adoPesquisa("referencia") & Chr(9) & adoPesquisa("descricao")
            adoPesquisa.MoveNext
        Loop
    
        'Seleciona o primeiro Registro
        grdPesquisa.row = 1
        pesquisaEstoque
    End If
    
    adoPesquisa.Close
    
    
End Sub

Private Sub pesquisaEstoque()

    Dim sql As String
    Dim referencia As String
    Dim adoEstoque As New ADODB.Recordset
    Dim wOrganiza As Integer
    Dim organiza As Integer
    
    wOrganiza = 0
    organiza = 0
    
    'Limpa Grid Estoque
    grdEstoque.Rows = 1
    grdEstoque.AddItem ""
    grdEstoque.RemoveItem (1)
    
    referencia = Trim(grdPesquisa.TextMatrix(grdPesquisa.row, 0))
    
    sql = "exec sp_cdm_busca_sugestao 7, '','" & referencia & "','','','','','','','',''"
    
    adoEstoque.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not adoEstoque.EOF Then
        Do While Not adoEstoque.EOF
            
            grdEstoque.AddItem adoEstoque("es_loja") & Chr(9) & adoEstoque("es_estoque") & Chr(9) & adoEstoque("es_romaneio") & Chr(9) & adoEstoque("es_transito") & Chr(9) & adoEstoque("es_venda") & Chr(9) & adoEstoque("es_MinimoInformado") & Chr(9) & adoEstoque("es_maximoInformado") & Chr(9) & adoEstoque("es_mixProduto")
            adoEstoque.MoveNext
        
        Loop
    
    grdEstoque.row = 1
    
    For i = 1 To grdEstoque.Rows - 1 Step 1
       If Val(grdEstoque.TextMatrix(i, 1)) < Val(grdEstoque.TextMatrix(i, 5)) Then
          organiza = organiza + 1
          grdEstoque.TextMatrix(i, 8) = organiza
          grdEstoque.FillStyle = flexFillRepeat
          grdEstoque.row = i
          grdEstoque.RowSel = i
          grdEstoque.col = 0
          grdEstoque.ColSel = grdEstoque.Cols - 1
          grdEstoque.CellForeColor = &HFF&        '&H80000008
          grdEstoque.FillStyle = flexFillSingle
       End If
    Next i

    wOrganiza = organiza + 1
    
    For i = 1 To grdEstoque.Rows - 1 Step 1
       If Trim(grdEstoque.TextMatrix(i, 8)) = "" Then
          grdEstoque.TextMatrix(i, 8) = wOrganiza
          wOrganiza = wOrganiza + 1
       End If
    Next i
    
    grdEstoque.row = 1
    grdEstoque.col = 8
    grdEstoque.RowSel = grdEstoque.Rows - 1
    grdEstoque.Sort = 1
    grdEstoque.RowSel = 0
    
    End If
    
    adoEstoque.Close
End Sub

Private Sub SugestaoAbertas()
    
    Dim adoSugestaoAberta As New ADODB.Recordset
    Dim sql As String
    
    sql = "exec sp_cdm_Busca_Sugestao 8, '','','','','','','','','',''"
    adoSugestaoAberta.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not adoSugestaoAberta.EOF Then
        'Limpa Grid Estoque
        grdSugestaoAberta.Rows = 1
        grdSugestaoAberta.AddItem ""
        grdSugestaoAberta.RemoveItem (1)
    
        Do While Not adoSugestaoAberta.EOF
            
            grdSugestaoAberta.AddItem adoSugestaoAberta("st_numeroSugestao") & Chr(9) & _
            adoSugestaoAberta("st_referencia") & Chr(9) & _
            adoSugestaoAberta("pr_descricao") & Chr(9) & _
            adoSugestaoAberta("st_lojaOrigem") & Chr(9) & _
            adoSugestaoAberta("st_lojaDestino") & Chr(9) & _
            adoSugestaoAberta("st_quantidade")
            
            adoSugestaoAberta.MoveNext
        Loop
        
    End If
    adoSugestaoAberta.Close
End Sub



Private Sub SalvaSugestao()

    Dim sql As String
    Dim referencia As String
    Dim quantidade As String
    Dim LojaOrigem As String
    Dim LojaDestino As String
    
    referencia = txtReferencia.Text
    quantidade = txtQuantidade.Text
    LojaOrigem = cmbLojaOrigem.Text
    LojaDestino = cmbLojaDestino.Text
    
    sql = "Exec TransfLoja '" & LojaOrigem & "', '" & LojaDestino & "', '" & referencia & "', " & quantidade & ", 'A', '" & GLB_USU_Nome & "'"
    ADO_Cn_CDLocal.Execute (sql)
    
    txtReferencia.Text = ""
    txtDescricao.Text = ""
    txtQuantidade.Text = ""
    SugestaoAbertas

End Sub

Private Sub txtQuantidade_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If cmbLojaOrigem.Text <> cmbLojaDestino.Text Then
            SalvaSugestao
        Else
            MsgBox "As lojas de origem e destino devem ser distintas.", vbExclamation, "Atenção!"
            Exit Sub
        End If
    End If

End Sub
