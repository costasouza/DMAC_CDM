VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmManutencaoSugestao 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Manutenção Sugestão"
   ClientHeight    =   8145
   ClientLeft      =   3960
   ClientTop       =   2325
   ClientWidth     =   15510
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   15510
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSugestao 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   510
      Left            =   120
      TabIndex        =   30
      Top             =   780
      Width           =   1335
      Begin VB.TextBox txtNroSugestao 
         BackColor       =   &H00C0C0C0&
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
         Left            =   120
         TabIndex        =   31
         ToolTipText     =   "Numero da Sugestao"
         Top             =   90
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   165
      TabIndex        =   25
      Top             =   6150
      Visible         =   0   'False
      Width           =   14880
      Begin VB.TextBox txtDescricao 
         BackColor       =   &H00C0C0C0&
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
         Left            =   120
         TabIndex        =   28
         ToolTipText     =   "Descricao"
         Top             =   90
         Width           =   10710
      End
      Begin VB.TextBox txtQuantidade 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
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
         Left            =   10905
         TabIndex        =   27
         ToolTipText     =   "Quantidade"
         Top             =   90
         Width           =   975
      End
      Begin VB.CheckBox chkMarcaTodos 
         BackColor       =   &H00404040&
         Caption         =   "Marcar Todos"
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   13215
         TabIndex        =   26
         ToolTipText     =   "Marcar Todos"
         Top             =   90
         Width           =   1575
      End
   End
   Begin VB.Frame fraStatus 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   150
      TabIndex        =   16
      Top             =   150
      Width           =   14880
      Begin VB.Frame fraPesquisa 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   195
         Left            =   8400
         TabIndex        =   22
         Top             =   150
         Width           =   3285
         Begin VB.OptionButton optFornecedor 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "Fornecedor"
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            Left            =   2000
            TabIndex        =   24
            ToolTipText     =   "Pesquisar por Fornecedor"
            Top             =   0
            Width           =   1665
         End
         Begin VB.OptionButton optTodos 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "Todos"
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            Left            =   0
            TabIndex        =   23
            ToolTipText     =   "Pesquisar Todos os Fornecedores"
            Top             =   0
            Value           =   -1  'True
            Width           =   1470
         End
      End
      Begin VB.OptionButton optProcessar 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "A Processar"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   4200
         TabIndex        =   21
         ToolTipText     =   "A Processar"
         Top             =   150
         Width           =   1900
      End
      Begin VB.OptionButton optProcessado 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Processados"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   6300
         TabIndex        =   20
         ToolTipText     =   "Processados"
         Top             =   150
         Width           =   1900
      End
      Begin VB.OptionButton optPeriodo 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Por Período"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   2200
         TabIndex        =   19
         ToolTipText     =   "Pesquisar por Período"
         Top             =   150
         Width           =   1900
      End
      Begin VB.OptionButton optSugestao 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Número da Sugestão"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   150
         TabIndex        =   18
         ToolTipText     =   "Pesquisar por Número da Sugestão"
         Top             =   150
         Value           =   -1  'True
         Width           =   1900
      End
      Begin VB.TextBox txtFornecedor 
         BackColor       =   &H00C0C0C0&
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
         Left            =   11730
         TabIndex        =   17
         ToolTipText     =   "Codigo do fornecedor"
         Top             =   90
         Width           =   735
      End
   End
   Begin VB.Frame fraProcessos 
      BackColor       =   &H00505050&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   195
      Left            =   3960
      TabIndex        =   9
      Top             =   240
      Width           =   2895
   End
   Begin VB.Frame FraPeriodo 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   150
      TabIndex        =   1
      Top             =   780
      Visible         =   0   'False
      Width           =   14880
      Begin VB.ComboBox CmbLojaDestino 
         BackColor       =   &H00C0C0C0&
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
         Left            =   6120
         TabIndex        =   4
         Text            =   "Combo2"
         ToolTipText     =   "Loja Destino"
         Top             =   90
         Width           =   1095
      End
      Begin VB.ComboBox CmbLojaOrigem 
         BackColor       =   &H00C0C0C0&
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
         Left            =   4080
         TabIndex        =   3
         Text            =   "Combo1"
         ToolTipText     =   "Loja Origem"
         Top             =   90
         Width           =   1095
      End
      Begin MSMask.MaskEdBox MskDataFim 
         Height          =   315
         Left            =   2040
         TabIndex        =   2
         ToolTipText     =   "Data de Termino"
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12632256
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskDataInicio 
         Height          =   315
         Left            =   465
         TabIndex        =   29
         ToolTipText     =   "Data de Inicio"
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12632256
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
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         Caption         =   "Destino"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   5505
         TabIndex        =   8
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Origem"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   3510
         TabIndex        =   7
         Top             =   150
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "à"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   1800
         TabIndex        =   6
         Top             =   150
         Width           =   90
      End
      Begin VB.Label lblDe 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "De"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   150
         Width           =   210
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
   Begin VSFlex7DAOCtl.VSFlexGrid grdSugestao 
      Height          =   4620
      Left            =   150
      TabIndex        =   10
      Top             =   1395
      Width           =   14880
      _cx             =   26247
      _cy             =   8149
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmManutencaoSugestao.frx":0000
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
      BackColorFrozen =   5263440
      ForeColorFrozen =   4210752
      WallPaperAlignment=   9
   End
   Begin CentroDeDistribuicao.chameleonButton cmdRetorna 
      Height          =   510
      Left            =   13620
      TabIndex        =   11
      ToolTipText     =   "Retornar"
      Top             =   6945
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
      MICON           =   "frmManutencaoSugestao.frx":0136
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdPesquisa 
      Height          =   510
      Left            =   12180
      TabIndex        =   12
      ToolTipText     =   "Pesquisar"
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Pesquisar"
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
      MICON           =   "frmManutencaoSugestao.frx":0152
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdCancela 
      Height          =   510
      Left            =   10740
      TabIndex        =   13
      ToolTipText     =   "Cancelar"
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Cancelar"
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
      MICON           =   "frmManutencaoSugestao.frx":016E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdProcessa 
      Height          =   510
      Left            =   9300
      TabIndex        =   14
      ToolTipText     =   "Processar"
      Top             =   6945
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
      MICON           =   "frmManutencaoSugestao.frx":018A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdEnvia 
      Height          =   510
      Left            =   7860
      TabIndex        =   15
      ToolTipText     =   "Enviar para Loja"
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Enviar p/ Loja"
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
      MICON           =   "frmManutencaoSugestao.frx":01A6
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
Attribute VB_Name = "frmManutencaoSugestao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPesquisa_Click()
    pesquisar
End Sub

Private Sub chkMarcaTodos_Click()
    Dim i As Integer
    
    If chkMarcaTodos.Value = 1 Then
        
        For i = 1 To grdSugestao.Rows - 1
            grdSugestao.TextMatrix(i, 0) = 1
        Next i
    Else
        For i = 1 To grdSugestao.Rows - 1
            grdSugestao.TextMatrix(i, 0) = 0
        Next i
    
    End If
    
    
End Sub

Private Sub cmdCancela_Click()
    CancelaSugestao
    pesquisar
End Sub

Private Sub cmdEnvia_Click()
    EnviaSugestao
End Sub

Private Sub cmdProcessa_Click()
    ProcessaSugestao
    pesquisar
End Sub

Private Sub cmdRetorna_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Dim adoLojas As New ADODB.Recordset
    Dim sql As String
    
    carregarPosicaoTamanhoTela Me
    FraPeriodo.Visible = False
    fraSugestao.Visible = True
    MskDataInicio.Text = Format(DateAdd("d", -1, Date), "dd/mm/yyyy")
    MskDataFim.Text = Format(Date, "dd/mm/yyyy")
    sql = "sp_cdm_busca_sugestao 10,'','','','','','','','','',''"

    adoLojas.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not adoLojas.EOF Then
        CmbLojaDestino.Text = "Todas"
        CmbLojaOrigem.Text = "Todas"
        
        Do While Not adoLojas.EOF
        
            CmbLojaDestino.AddItem adoLojas("LO_loja")
            CmbLojaOrigem.AddItem adoLojas("LO_loja")
            adoLojas.MoveNext
            
        Loop
        
        CmbLojaOrigem.AddItem "Todas"
        CmbLojaDestino.AddItem "Todas"
    End If
    
    adoLojas.Close
    
End Sub


Private Sub grdSugestao_Click()
    If grdSugestao.row <> 0 Then
        txtDescricao.Text = grdSugestao.TextMatrix(grdSugestao.row, 5)
        txtQuantidade.Text = grdSugestao.TextMatrix(grdSugestao.row, 6)
    End If
End Sub


Private Sub grdSugestao_dblClick()
    
    Dim sql As String
    
    If Trim(grdSugestao.TextMatrix(grdSugestao.row, 8)) <> "B" Then
    
        If Trim(grdSugestao.TextMatrix(grdSugestao.row, 1)) <> 0 Then
            If MsgBox("Confirma O cancelamento desta sugestão ?" & Chr(vbKeyReturn) & Chr(vbKeyReturn), vbQuestion + vbYesNo, "Exclusão de Sugestão") = vbYes Then
                
                CancelaSugestao
                
                pesquisar
                
            End If
        End If
    End If
End Sub

Private Sub optFornecedor_Click()
    txtFornecedor.Enabled = True
End Sub

Private Sub optPeriodo_Click()

    MskDataInicio.Text = Format(DateAdd("d", -1, Date), "dd/mm/yyyy")
    MskDataFim.Text = Format(Date, "dd/mm/yyyy")
    FraPeriodo.Visible = True
    fraSugestao.Visible = False
    CmbLojaDestino.Text = "Todas"
    CmbLojaOrigem.Text = "Todas"

End Sub

Private Sub optProcessado_Click()
    FraPeriodo.Visible = False
    fraSugestao.Visible = False
End Sub

Private Sub optProcessar_Click()
    FraPeriodo.Visible = False
    fraSugestao.Visible = False
End Sub

Private Sub optSugestao_Click()

    FraPeriodo.Visible = False
    fraSugestao.Visible = True
    
End Sub

Private Sub pesquisar()
    
    Dim sql As String
    Dim Situacao As String
    Dim adoPesquisa As New ADODB.Recordset

    'Limpa Grid Pesquisa
    grdSugestao.Rows = 1
    grdSugestao.AddItem ""
    grdSugestao.RemoveItem (1)
    
    If optProcessado.Value Then
        Situacao = "'P'"
    Else
        Situacao = "'A'"
    End If
    
    
    If optSugestao.Value Then
        
        sql = "Exec sp_cdm_busca_sugestao 13,'','','','','','','','" & txtNroSugestao.Text & "','',''"
    
    Else
        
        If optPeriodo.Value And optTodos.Value = False Then
            
            sql = "Exec sp_cdm_busca_sugestao 11,'" & txtFornecedor.Text & "','','','" & Format(MskDataInicio.Text, "yyyy/mm/dd") & "','" & Format(MskDataFim.Text, "yyyy/mm/dd") & "','" & CmbLojaOrigem.Text & "'," & Situacao & ",'','" & CmbLojaDestino & "',''"
        Else
            
            sql = "Exec sp_cdm_busca_sugestao 12,'','','','" & Format(MskDataInicio.Text, "yyyy/mm/dd") & "','" & Format(MskDataFim.Text, "yyyy/mm/dd") & "','" & CmbLojaOrigem.Text & "'," & Situacao & ",'','" & CmbLojaDestino & "',''"
        End If
        
    End If
     
     adoPesquisa.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
     
     If Not adoPesquisa.EOF Then
     
        
     
        Do While Not adoPesquisa.EOF
        
            grdSugestao.AddItem Chr(9) & Trim(adoPesquisa("st_numeroSugestao")) & Chr(9) _
                              & Trim(adoPesquisa("st_referencia")) & Chr(9) _
                              & Trim(adoPesquisa("st_lojaOrigem")) & Chr(9) _
                              & Trim(adoPesquisa("st_lojaDestino")) & Chr(9) _
                              & Trim(adoPesquisa("pr_descricao")) & Chr(9) _
                              & Trim(adoPesquisa("st_quantidade")) & Chr(9) _
                              & Trim(adoPesquisa("st_dataSugestao")) & Chr(9) _
                              & Trim(adoPesquisa("st_situacao")) & Chr(9) _
                              & Trim(adoPesquisa("st_NomeUsuario"))
            adoPesquisa.MoveNext
        Loop
        
    End If
    
    adoPesquisa.Close
    
End Sub

Private Sub optTodos_Click()
    txtFornecedor.Enabled = False
End Sub

Private Sub txtQuantidade_LostFocus()
    Dim sql As String
    
    sql = "Exec SP_CDM_Busca_Sugestao 14,'','" & grdSugestao.TextMatrix(grdSugestao.row, 2) & "','','','','" & _
    grdSugestao.TextMatrix(grdSugestao.row, 3) & "','','" & grdSugestao.TextMatrix(grdSugestao.row, 1) & _
    "', '" & grdSugestao.TextMatrix(grdSugestao.row, 4) & "', " & txtQuantidade.Text
    
    If GetAsyncKeyState(vbKeyTab) <> 0 Then
        If txtQuantidade.Text = "" Then
            txtQuantidade.Text = 0
        End If
        
        If txtQuantidade.Text = 0 Then
            MsgBox "Quantidade não pode ser igual a zero.", vbInformation, "Informação"
            txtQuantidade.SetFocus
        Else
        
            ADO_Cn_CDLocal.Execute sql
               
           
                grdSugestao.TextMatrix(grdSugestao.row, 6) = txtQuantidade.Text
                txtDescricao.Text = ""
                txtQuantidade.Text = ""
        End If
    End If
    
End Sub


Sub ProcessaSugestao()
    Dim i As Integer
    Dim sql As String
    
    For i = 1 To grdSugestao.Rows - 1
        
        
        If grdSugestao.TextMatrix(i, 0) <> 0 Then

            sql = "EXEC SP_CDM_Busca_Sugestao 16,'','" & grdSugestao.TextMatrix(i, 2) & "','','" & Format(grdSugestao.TextMatrix(i, 7), "yyyy/mm/dd") & "','','" & grdSugestao.TextMatrix(i, 3) & "','','','" & grdSugestao.TextMatrix(i, 4) & "',''"
            
            ADO_Cn_CDLocal.Execute sql
            
        
        End If
    
    Next i
    
    chkMarcaTodos.Value = 0
    chkMarcaTodos_Click

End Sub


Sub CancelaSugestao()
    Dim i As Integer
    Dim sql As String
    
    
    For i = 1 To grdSugestao.Rows - 1
        grdSugestao.TextMatrix(i, 0) = 1
        If grdSugestao.TextMatrix(i, 0) <> 0 Then

            
            sql = "EXEC SP_CDM_Busca_Sugestao 15,'','" & grdSugestao.TextMatrix(i, 2) & "','','" & Format(grdSugestao.TextMatrix(i, 7), "yyyy/mm/dd") & "','','" & grdSugestao.TextMatrix(i, 3) & "','','" & grdSugestao.TextMatrix(i, 1) & "','" & grdSugestao.TextMatrix(i, 4) & "',''"
            
            ADO_Cn_CDLocal.Execute sql
               
        End If
    
    Next i
    
    chkMarcaTodos.Value = 0
    chkMarcaTodos_Click
    
End Sub


Sub EnviaSugestao()
    Dim i As Integer
    Dim sql As String
    
    
    For i = 1 To grdSugestao.Rows - 1
        
        If grdSugestao.TextMatrix(i, 0) <> 0 Then
        
            
            sql = "EXEC SP_CDM_Busca_Sugestao 17,'','" & grdSugestao.TextMatrix(i, 2) & "','','" & grdSugestao.TextMatrix(i, 7) & "','','" & grdSugestao.TextMatrix(i, 3) & "','','" & grdSugestao.TextMatrix(i, 1) & "','" & grdSugestao.TextMatrix(i, 4) & "',''"
            
            ADO_Cn_CDLocal.Execute sql
        
        End If
    
    Next i
    
    chkMarcaTodos.Value = 0
    chkMarcaTodos_Click
    
    If Not GLB_modoOffline Then
        sql = "SP_Atualiza_Tarefas 2"
        
        ADO_Cn_CDLocal.Execute (sql)
    End If
End Sub
