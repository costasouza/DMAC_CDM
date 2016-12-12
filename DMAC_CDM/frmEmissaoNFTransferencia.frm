VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmEmissaoNFTransferencia 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Emissão de Nota Fiscal Transferencia"
   ClientHeight    =   7965
   ClientLeft      =   735
   ClientTop       =   2115
   ClientWidth     =   15300
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   15300
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameCancelamento 
      BackColor       =   &H00393939&
      BorderStyle     =   0  'None
      Caption         =   "?"
      Height          =   3030
      Left            =   4710
      TabIndex        =   17
      Top             =   1755
      Visible         =   0   'False
      Width           =   2730
      Begin VB.TextBox txtSerie 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1395
         MaxLength       =   3
         TabIndex        =   20
         ToolTipText     =   "Referência"
         Top             =   675
         Width           =   1170
      End
      Begin VB.TextBox txtNotaFiscal 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   150
         MaxLength       =   10
         TabIndex        =   19
         ToolTipText     =   "Código do Produto"
         Top             =   675
         Width           =   1170
      End
      Begin CentroDeDistribuicao.chameleonButton cmdCancelar 
         Height          =   330
         Left            =   150
         TabIndex        =   22
         Top             =   2535
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         BTYPE           =   14
         TX              =   "&Cancelar"
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
         MICON           =   "frmEmissaoNFTransferencia.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CentroDeDistribuicao.chameleonButton cmdSairCancelamento 
         Height          =   330
         Left            =   1395
         TabIndex        =   23
         Top             =   2535
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         BTYPE           =   14
         TX              =   "&Sair"
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
         MICON           =   "frmEmissaoNFTransferencia.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelar Nota"
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
         Height          =   285
         Left            =   150
         TabIndex        =   25
         Top             =   150
         Width           =   2070
      End
      Begin VB.Label lblinfoNota 
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         ForeColor       =   &H00C0C0C0&
         Height          =   990
         Left            =   150
         TabIndex        =   24
         Top             =   1170
         Width           =   2235
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Nota Fiscal"
         ForeColor       =   &H00C0C0C0&
         Height          =   210
         Left            =   150
         TabIndex        =   21
         Top             =   435
         Width           =   1470
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Serie"
         ForeColor       =   &H00C0C0C0&
         Height          =   300
         Left            =   1395
         TabIndex        =   18
         Top             =   435
         Width           =   1170
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   44
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   14880
      TabIndex        =   10
      Top             =   6810
      Width           =   14880
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdItens 
      Height          =   5520
      Left            =   4605
      TabIndex        =   6
      Top             =   1125
      Width           =   10425
      _cx             =   18389
      _cy             =   9737
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmEmissaoNFTransferencia.frx":0038
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
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame fraPesquisa 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   150
      TabIndex        =   2
      Top             =   150
      Width           =   14865
      Begin VB.Frame fraUltimaNf 
         Appearance      =   0  'Flat
         BackColor       =   &H00393939&
         BorderStyle     =   0  'None
         Caption         =   "Última Nota Fiscal"
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   13200
         TabIndex        =   7
         Top             =   105
         Width           =   1545
         Begin VB.Label Label1 
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Última Nota Fiscal"
            ForeColor       =   &H00FA9923&
            Height          =   195
            Left            =   150
            TabIndex        =   12
            Top             =   45
            Width           =   2655
         End
         Begin VB.Label lblNroNotaFiscal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00393939&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000F&
            Height          =   300
            Left            =   195
            TabIndex        =   8
            ToolTipText     =   "Dê um duplo-clique para atualiza"
            Top             =   240
            Width           =   1200
         End
      End
      Begin VB.CheckBox chkRefereForne 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Todas as Referência / Fornecedor"
         ForeColor       =   &H00C0C0C0&
         Height          =   240
         Left            =   2040
         TabIndex        =   4
         Top             =   420
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.TextBox txtPesquisa 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   120
         MaxLength       =   7
         TabIndex        =   3
         Top             =   375
         Width           =   1755
      End
      Begin VB.Label lblRefereForne 
         BackColor       =   &H00404040&
         Caption         =   "Fornecedor / Referência"
         ForeColor       =   &H00FA9923&
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   2070
      End
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdRomaneio 
      Height          =   5520
      Left            =   1305
      TabIndex        =   1
      Top             =   1125
      Width           =   3030
      _cx             =   5345
      _cy             =   9737
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
      FormatString    =   $"frmEmissaoNFTransferencia.frx":013A
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
      BackColorFrozen =   5263440
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      Begin VB.CheckBox chkTodosRomaneios 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Todos os Romaneio"
         ForeColor       =   &H8000000F&
         Height          =   195
         Left            =   30
         TabIndex        =   14
         Top             =   30
         Width           =   195
      End
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdLojas 
      Height          =   5520
      Left            =   150
      TabIndex        =   0
      Top             =   1125
      Width           =   1005
      _cx             =   1773
      _cy             =   9737
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
      FormatString    =   $"frmEmissaoNFTransferencia.frx":01B2
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
      BackColorFrozen =   5263440
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      Begin VB.CheckBox chkTodasLojas 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Todas as Lojas"
         ForeColor       =   &H8000000F&
         Height          =   195
         Left            =   30
         MaskColor       =   &H00000000&
         TabIndex        =   13
         Top             =   30
         Width           =   195
      End
   End
   Begin CentroDeDistribuicao.chameleonButton cmdRetornar 
      Height          =   510
      Left            =   13620
      TabIndex        =   9
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
      MICON           =   "frmEmissaoNFTransferencia.frx":01F6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdImprimirNF 
      Height          =   510
      Left            =   10740
      TabIndex        =   11
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Imprimir NF"
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
      MICON           =   "frmEmissaoNFTransferencia.frx":0212
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdLimpa 
      Height          =   510
      Left            =   12180
      TabIndex        =   15
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Limpar"
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
      MICON           =   "frmEmissaoNFTransferencia.frx":022E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdCancelarFRAME 
      Height          =   510
      Left            =   9300
      TabIndex        =   16
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Cancelar NF"
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
      MICON           =   "frmEmissaoNFTransferencia.frx":024A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdReImpressao 
      Height          =   510
      Left            =   6300
      TabIndex        =   26
      Top             =   6945
      Visible         =   0   'False
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Cancelar NF"
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
      MICON           =   "frmEmissaoNFTransferencia.frx":0266
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdTransfMerc 
      Height          =   510
      Index           =   1
      Left            =   7800
      TabIndex        =   27
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "CD"
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
      MICON           =   "frmEmissaoNFTransferencia.frx":0282
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
Attribute VB_Name = "frmEmissaoNFTransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim whereLoja As String
Dim whereRomaneio As String
Dim whereReferenciaOuFornecedor As String
Dim orderByPadrao As String
Dim orderByPadraoRomaneio As String

Dim vetorLojasSelecionadas() As String
Dim proximoValorVetor As String


Private Sub cmdTransfMerc_Click(Index As Integer)

    If cmdTransfMerc(1).Caption = "CD" Then
        cmdTransfMerc(1).Caption = "CMC"
    ElseIf cmdTransfMerc(1).Caption = "CMC" Then
        cmdTransfMerc(1).Caption = "CMCE"
    Else
        cmdTransfMerc(1).Caption = "CD"
    End If
    
End Sub

'///////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////// COMPONENTES //////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////

Private Sub Form_Load()
    carregaValoresPadraoFormulario
    
    Call CarregaLojas
    lblNroNotaFiscal.Caption = ultimaNotaFiscal
    
    ReDim Preserve vetorLojasSelecionadas(0) As String
    vetorLojasSelecionadas(0) = "0"
End Sub

Private Sub Form_Initialize()
    carregarPosicaoTamanhoTela Me
    carregarPosicaoFrame frameCancelamento
    montaColunaGridItens
End Sub

Private Sub chkTodasLojas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    marcaTodasCaixa grdLojas, chkTodasLojas, 0
    carregaGridRomaneio
    carregaGridItens
End Sub

Private Sub chkTodosRomaneios_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    marcaTodasCaixa grdRomaneio, chkTodosRomaneios, 0
    carregaGridItens
End Sub

Private Sub cmdImprimirNF_Click()
     'MousePointer = 11
    
    Dim numeroRomaneioImpressao As Integer
    Dim adoNotaFiscalImpressao As New ADODB.Recordset
    Dim sql As String
    
    Dim adoPrecoCusto As New ADODB.Recordset
    
    'ricardo 26/10/2016
    sql = "select distinct PR_PrecoCusto1 " _
        & " From romaneio, produto, Loja " _
        & " where ro_situacao= 'A' and ro_referencia = pr_referencia and lo_loja = ro_lojaDestino and ro_numeroRomaneio > 0 and  " _
        & " pr_referencia = '" & grdItens.TextMatrix(grdItens.row, 3) & "' "
        
             
    adoPrecoCusto.CursorLocation = adUseClient
    adoPrecoCusto.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
       
       
    If IsNull(adoPrecoCusto("PR_PrecoCusto1")) Or adoPrecoCusto("PR_PrecoCusto1") = "" Or adoPrecoCusto("PR_PrecoCusto1") = 0 Then
            MsgBox "O custo do Produto não pode ser zero(0) ou estar vazio!", vbInformation
        Exit Sub
    
    Else
    
        numeroRomaneioImpressao = gerarNumeroRomaneioImpressao
        gerarNotaFiscalImpressao numeroRomaneioImpressao
    
        cmdTransfMerc(1).Enabled = False
        LojaOrigem = cmdTransfMerc(1).Caption
        'serieImpressao = "NE"
        
        sql = "select nf as numeroNotaFiscal, serie as serie " & _
          "from nfcapa " & _
          "where LojaOrigem = '" & LojaOrigem & _
          "' and protocolo = " & numeroRomaneioImpressao


    With adoNotaFiscalImpressao
        .CursorLocation = adUseClient
        .Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic


        Do While Not .EOF
            NroNotaFiscal = adoNotaFiscalImpressao("numeroNotaFiscal")
            serieImpressao = adoNotaFiscalImpressao("serie")
            'If ImprimirNotaFiscal(adoNotaFiscalImpressao("numeroNotaFiscal")) = True Then
                      sql = "update nfcapa set situacaoEnvio = 'I' " & _
                      "where nf = " & NroNotaFiscal & " and " & _
                      "serie = '" & serieImpressao & "' and LojaOrigem = '" & LojaOrigem & _
                      "' and protocolo = " & numeroRomaneioImpressao
                ADO_Cn_CDLocal.Execute sql


                sql = "update romaneio set ro_situacao = 'P'," & vbNewLine & _
                    "ro_notaFiscal = '" & adoNotaFiscalImpressao("numeroNotaFiscal") & "'" & vbNewLine & _
                    "where " & whereRomaneio & " and " & whereLoja & " and ro_situacao='A' and ro_numeroRomaneio > 0"
                ADO_Cn_CDLocal.Execute sql


                Call carimboImposto(Str(NroNotaFiscal), serieImpressao, LojaOrigem)


                If Trim(serieImpressao) = "NE" Then
                    sql = "exec sp_vda_cria_nfe '" & LojaOrigem & "','" & adoNotaFiscalImpressao("numeroNotaFiscal") & "','" & serieImpressao & "','" + wImpressoraNota + "'"
                    ADO_Cn_CDLocal.Execute sql
                Else
                    ImprimeTransferencia00 CStr(NroNotaFiscal), serieImpressao, LojaOrigem
                End If

            'Else
                'MsgBox "Nota Fiscal " & adoNotaFiscalImpressao("numeroNotaFiscal") & " não Encontrada", vbCritical, "Erro"
            'End If
            'Printer.EndDoc
            .MoveNext
        Loop
    End With


    limpaGrid grdItens
    carregaGridItens
    carregaValoresPadraoFormulario


    frmStartaProcessos.Show vbModal


    lblNroNotaFiscal.Caption = ultimaNotaFiscal
    'MousePointer = 0
    'MsgBox "Impressão concluída", vbInformation, Me.Caption

    End If
    
adoPrecoCusto.Close
End Sub

Private Sub cmdRetornar_Click()
    Unload Me
End Sub

Private Sub lblDescricaoFornecedor_Click()

End Sub

Private Sub frameCadastraReferencia_DragDrop(Source As Control, x As Single, y As Single)

End Sub

Private Sub lblNroNotaFiscal_DblClick()
    MousePointer = 11
    lblNroNotaFiscal.Caption = ultimaNotaFiscal
    MousePointer = 0
End Sub

Private Sub txtNotaFiscal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtSerie.SetFocus
    KeyAscii = campoNumerico(KeyAscii)
End Sub

Private Sub txtNotaFiscal_LostFocus()
    If txtNotaFiscal.Text <> "" And txtSerie.Text <> "" Then
        carregaInfoNotaCancelamento txtNotaFiscal.Text, txtSerie.Text
    End If
End Sub

Private Sub carregaInfoNotaCancelamento(ByRef nf As String, ByRef serie As String)

    Dim adoCancelamento As New ADODB.Recordset
    Dim sql As String
    
    sql = "select nf as nf, " & vbNewLine & _
          "dataemi as data, " & vbNewLine & _
          "totalnota as total, " & vbNewLine & _
          "lojat as lojaDestino " & vbNewLine & _
          "from nfcapa " & vbNewLine & _
          "where serie = '" & serie & "' and nf = '" & nf & "'" & vbNewLine & _
          "AND DATAEMI = '" & Format(Date, "YYYY/MM/DD") & "'"
          
    With adoCancelamento
        .CursorLocation = adUseClient
        .Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
        If Not adoCancelamento.EOF Then
            lblinfoNota.Caption = "Nota Fiscal: " & adoCancelamento("NF") & vbNewLine & _
            "Serie: " & txtSerie.Text & vbNewLine & _
            "Data Emissão: " & adoCancelamento("data") & vbNewLine & _
            "Total Nota: " & adoCancelamento("total") & vbNewLine & _
            "Loja Destino: " & adoCancelamento("lojaDestino")
            cmdCancelar.Enabled = True
        Else
            cmdCancelar.Enabled = False
            lblinfoNota.Caption = "Nenhuma nota encontrada OU nota já cancelada no sistema"
        End If
        
        .Close
        
    End With
                  
          
End Sub

Private Sub txtPesquisa_LostFocus()
    carregaGridItens
End Sub

Private Sub grdLojas_Click()
    MousePointer = 11
    marcaCaixaGrid grdLojas, chkTodasLojas
    carregaGridRomaneio
    carregaGridItens
    MousePointer = 0
End Sub

Private Sub grdRomaneio_Click()
MousePointer = 11
    marcaCaixaGrid grdRomaneio, chkTodosRomaneios
    carregaGridItens
MousePointer = 0
End Sub

Private Sub txtPesquisa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If IsNumeric(txtPesquisa) Then
            MousePointer = 11
        
            carregaGridRomaneio
            carregaGridItens
        
            MousePointer = 0
        Else
            txtPesquisa = ""
        End If
    End If
End Sub

'///////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////// PROCEDIMENTOS ////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////

Private Sub lojasMarcada(ByRef grid)
    Dim i As Integer
    Dim posicaoVetor As Integer
    
    ReDim vetorLojasSelecionadas(grid.Rows - grid.FixedRows) As String
    For i = grid.FixedRows To grid.Rows - grid.FixedRows
        If grid.TextMatrix(i, 0) Then
            vetorLojasSelecionadas(posicaoVetor) = Trim(grid.TextMatrix(i, 1))
            posicaoVetor = posicaoVetor + 1
        End If
    Next i
End Sub

Private Sub carregarLojasMarcadas(ByRef grid)
    Dim i As Integer
    Dim posicaoVetor As Integer
    
    For i = grid.FixedRows To grid.Rows - grid.FixedRows
        For posicaoVetor = 0 To UBound(vetorLojasSelecionadas())
            If Trim(grid.TextMatrix(i, 1)) = vetorLojasSelecionadas(posicaoVetor) Then
                grid.TextMatrix(i, 0) = True
            End If
        Next posicaoVetor
    Next i
End Sub

Private Sub carregaValoresPadraoFormulario()
    txtPesquisa = ""
    lblNroNotaFiscal.Caption = ultimaNotaFiscal
    chkRefereForne = 0
    chkTodosRomaneios = 0
    marcaTodasCaixa grdLojas, False, 0
    limpaGrid grdRomaneio
    limpaGrid grdItens
    
    limparVariaveis
    cmdImprimirNF.Enabled = False
End Sub

Private Sub limparVariaveis()
    whereLoja = ""
    whereRomaneio = ""
    whereReferenciaOuFornecedor = ""
    orderByPadrao = " order by LO_Regiao "
    orderByPadraoRomaneio = " order by ro_lojaDestino  "
End Sub

Private Function gerarNumeroRomaneioImpressao() As Integer
    Dim sql As String
    Dim adoNumeroRomaneioImpressao As New ADODB.Recordset
    
    ADO_Cn_CDLocal.Execute "update controlecdm set cs_romaneioImpressao = cs_romaneioImpressao + 1"
    
    sql = "select top 1 cs_romaneioImpressao ultimoNumeroRomaneioImpressao from controlecdm"
    adoNumeroRomaneioImpressao.CursorLocation = adUseClient
    adoNumeroRomaneioImpressao.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not adoNumeroRomaneioImpressao.EOF Then
        gerarNumeroRomaneioImpressao = adoNumeroRomaneioImpressao("ultimoNumeroRomaneioImpressao")
    End If
    
    adoNumeroRomaneioImpressao.Close
End Function

Private Sub gerarNotaFiscalImpressao(numeroRomaneioImpressao As Integer)
    Dim sql As String
    Dim adoNFImpressao As New ADODB.Recordset
    
    sql = "update romaneio " & _
          "set ro_romaneioImpressao = " & numeroRomaneioImpressao & " " & _
          "where " & whereLoja & " and " & whereRomaneio & " and ro_situacao='A' and ro_numeroRomaneio > 0"
          
    If pesquisaPorReferenciaOuFornecedor Then
        sql = sql & " and " & whereReferenciaOuFornecedor
    End If
    ADO_Cn_CDLocal.Execute sql
          
    sql = "exec SP_Cria_NF_Transferencia_novo " & numeroRomaneioImpressao & ""
    ADO_Cn_CDLocal.Execute sql
    
End Sub

Private Function ultimaNotaFiscal() As Integer
    Dim sql As String
    Dim adoNumeroNota As New ADODB.Recordset
    sql = "select cts_numerone as numeroNE from controleSistema"
    adoNumeroNota.CursorLocation = adUseClient
    adoNumeroNota.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not adoNumeroNota.EOF Then
        ultimaNotaFiscal = adoNumeroNota("numeroNE")
    End If
    adoNumeroNota.Close
End Function

Private Sub montaColunaGridItens()
    grdItens.MergeRow(0) = True
    grdItens.MergeRow(1) = True
    grdItens.MergeCol(0) = True
    grdItens.MergeCol(1) = True
    grdItens.MergeCol(2) = True
    grdItens.MergeCol(3) = True
    grdItens.MergeCol(4) = True
    grdItens.MergeCol(5) = True
End Sub

Private Sub CarregaLojas()
    Dim adoLojas As New ADODB.Recordset
    Dim sql As String
    
    sql = "select lo_loja " & _
          "From Loja " & _
          "Where (lo_Regiao < 450 or lo_loja = 'CMC') and LO_Loja Not in('185') " & _
          orderByPadrao
          
    With adoLojas
        .CursorLocation = adUseClient
        .Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
        limpaGrid grdLojas
        Do While Not .EOF
            grdLojas.AddItem chkTodasLojas & Chr(9) & Trim(adoLojas("Lo_Loja"))
            .MoveNext
        Loop
         
        .Close
    End With
End Sub

Private Sub carregaRomaneio()
    Dim sql As String
    Dim adoNumeroNota As New ADODB.Recordset
    sql = "select cs_numeroNotaFiscal from controlecdm"
    adoNumeroNota.CursorLocation = adUseClient
    adoNumeroNota.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If adoNumeroNota.EOF Then
        lblNroNotaFiscal.Caption = "0000"
    Else
        lblNroNotaFiscal.Caption = adoNumeroNota("cs_numeroNotaFiscal")
    End If
End Sub

Private Sub marcaCaixaGrid(ByRef grid, check As CheckBox)
    Dim i As Integer

    grid.col = 0
    If grid.row >= grid.FixedRows Then
        If grid.Text = "True" Then
            grid.Text = False
            check.Value = False
        Else
            grid.Text = True
        End If
        'lojasMarcada grid.TextMatrix(grid.Row, 1), True
    End If
End Sub

Private Sub marcaTodasCaixa(ByRef grid, ativa As Boolean, coluna As Byte)
    Dim i As Integer
    If Not gridVazio(grid) Then
        For i = grid.FixedRows To grid.Rows - 1
            grid.TextMatrix(i, coluna) = ativa
        Next i
    End If
    grid.row = 0
End Sub

Private Function pesquisaPorLoja(ByRef grid) As Boolean
    
    Dim i As Integer
    pesquisaPorLoja = False
    whereLoja = "ro_lojaDestino in('',"
    
    For i = 1 To grid.Rows - 1
        If grid.TextMatrix(i, 0) = "True" Then
            whereLoja = whereLoja & "'" & RTrim(grid.TextMatrix(i, 1)) & "',"
            pesquisaPorLoja = True
        End If
    Next i
    
    whereLoja = left(whereLoja, (Len(whereLoja) - 1)) & ")"

End Function

Private Function pesquisaPorRomaneio() As Boolean
    Dim i As Integer
    pesquisaPorRomaneio = False
    
    whereRomaneio = "ro_numeroromaneio in ('',"
    For i = 1 To grdRomaneio.Rows - 1
        If grdRomaneio.TextMatrix(i, 0) = True Then
            whereRomaneio = whereRomaneio & "'" & RTrim(grdRomaneio.TextMatrix(i, 2)) & "',"
            pesquisaPorRomaneio = True
        End If
    Next i
    
    whereRomaneio = left(whereRomaneio, (Len(whereRomaneio) - 1)) & ")"
End Function

Private Sub carregaGridRomaneio()
    Dim adoRomaneio As New ADODB.Recordset
    Dim sql As String
    cmdTransfMerc(1).Enabled = False
    lojasMarcada grdRomaneio
    
    limpaGrid grdRomaneio
    sql = "select distinct RO_DataSolicitacao,ro_lojadestino,ro_numeroromaneio, LO_Regiao from romaneio, produto, loja" & _
          " where ro_situacao='A' and ro_referencia = pr_referencia and lo_loja = ro_lojaDestino and ro_numeroRomaneio > 0" & _
          " and ro_lojaOrigem = '" & cmdTransfMerc(1).Caption & "'"
          
    
    If pesquisaPorLoja(grdLojas) Then
        sql = sql & " and " & whereLoja
        If pesquisaPorReferenciaOuFornecedor Then
            sql = sql & " and " & whereReferenciaOuFornecedor
        End If
    Else
        Exit Sub
    End If
    
    sql = sql + orderByPadrao
    adoRomaneio.CursorLocation = adUseClient
    adoRomaneio.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    Do While Not adoRomaneio.EOF
        grdRomaneio.AddItem chkTodosRomaneios & Chr(9) & adoRomaneio("ro_lojadestino") & Chr(9) & _
        adoRomaneio("ro_numeroromaneio") & Chr(9) & adoRomaneio("RO_DataSolicitacao")
        adoRomaneio.MoveNext
    Loop
     
    carregarLojasMarcadas grdRomaneio
    
    adoRomaneio.Close
End Sub

Private Sub marcaCampoPesquisaAnterior()

End Sub

Private Function pesquisaPorReferenciaOuFornecedor() As Boolean
    pesquisaPorReferenciaOuFornecedor = False
    
    If Len(txtPesquisa) >= 1 And Len(txtPesquisa) <= 3 Then
        whereReferenciaOuFornecedor = "pr_codigoFornecedor = '" & txtPesquisa & "'"
        pesquisaPorReferenciaOuFornecedor = True
    ElseIf Len(txtPesquisa) = 7 Then
        whereReferenciaOuFornecedor = "ro_referencia = '" & txtPesquisa & "'"
        pesquisaPorReferenciaOuFornecedor = True
    End If
End Function

Private Sub carregaGridItens()
    Dim adoItens As New ADODB.Recordset
    Dim sql As String
    'Dim selectTabela As String
    cmdTransfMerc(1).Enabled = False
    limpaGrid grdItens
    sql = "select ro_sequencia, ro_lojaDestino, " & vbNewLine & _
    "ro_referencia, ro_quantidadePedida, pr_descricao, " & vbNewLine & _
    "ro_numeroRomaneio " & vbNewLine & _
    "from romaneio, Produto, loja" & vbNewLine
    
    sql = sql & " where ro_referencia = pr_referencia" & vbNewLine & _
    "and lo_loja = ro_lojaDestino " & vbNewLine & _
    "and ro_numeroRomaneio > 0" & vbNewLine
    If pesquisaPorLoja(grdLojas) And pesquisaPorRomaneio And pesquisaPorLoja(grdRomaneio) Then
        sql = sql & " and " & whereRomaneio & vbNewLine & " and " & whereLoja & vbNewLine & "and ro_situacao='A'"
        If pesquisaPorReferenciaOuFornecedor Then
            sql = sql & " and " & whereReferenciaOuFornecedor
        End If
    Else
        'MousePointer = 0
        cmdImprimirNF.Enabled = False
        Exit Sub
    End If
    
    sql = sql + orderByPadraoRomaneio + ", ro_numeroRomaneio"
    'MousePointer = 11
    adoItens.CursorLocation = adUseClient
    adoItens.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    Do While Not adoItens.EOF
        grdItens.AddItem adoItens("ro_lojaDestino") & Chr(9) & adoItens("ro_numeroRomaneio") & Chr(9) & _
                         "" & Chr(9) & adoItens("ro_referencia") & Chr(9) & _
                         adoItens("ro_quantidadePedida") & Chr(9) & adoItens("pr_descricao")
        adoItens.MoveNext
    Loop
    
    adoItens.Close
    cmdImprimirNF.Enabled = True
End Sub



Private Sub txtSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCancelar.SetFocus
End Sub

Private Sub txtSerie_LostFocus()
    If txtNotaFiscal.Text <> "" And txtSerie.Text <> "" Then
        carregaInfoNotaCancelamento txtNotaFiscal.Text, txtSerie.Text
    End If
End Sub

Private Sub cmdCancelar_Click()
    If cancelarNota(txtNotaFiscal.Text, txtSerie.Text) Then
        MsgBox "Nota " & txtNotaFiscal.Text & " cancelada com sucesso!", vbInformation, "Cancelamento"
        frameCancelamento.Visible = False
    End If
End Sub

Private Sub cmdCancelarFRAME_Click()
    exibirFrameCancelamento
End Sub

Private Sub exibirFrameCancelamento()
    frameCancelamento.Visible = True
    txtNotaFiscal.Text = ""
    txtSerie.Text = ""
    lblinfoNota.Caption = ""
    cmdCancelar.Enabled = False
    txtNotaFiscal.SetFocus
End Sub

Private Sub cmdLimpa_Click()
    carregaValoresPadraoFormulario
End Sub

Private Sub cmdSairCancelamento_Click()
    frameCancelamento.Visible = False
End Sub
