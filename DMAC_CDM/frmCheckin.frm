VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmCheckin 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Checkin de Mercadoria"
   ClientHeight    =   7470
   ClientLeft      =   3705
   ClientTop       =   1275
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7470
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Frame fraNota 
      BackColor       =   &H00393939&
      BorderStyle     =   0  'None
      Height          =   2640
      Left            =   1320
      TabIndex        =   40
      Top             =   3855
      Visible         =   0   'False
      Width           =   11235
      Begin VB.TextBox txtcnpj 
         BackColor       =   &H00393939&
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000F&
         Height          =   195
         Left            =   7320
         TabIndex        =   49
         Top             =   120
         Width           =   3735
      End
      Begin VSFlex7DAOCtl.VSFlexGrid grdXml 
         Height          =   2145
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   11055
         _cx             =   19500
         _cy             =   3784
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
         ForeColorSel    =   8388608
         BackColorBkg    =   3750201
         BackColorAlternate=   3947580
         GridColor       =   5263440
         GridColorFixed  =   8421504
         TreeColor       =   3947580
         FloodColor      =   5263440
         SheetBorder     =   3947580
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
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
         FormatString    =   $"frmCheckin.frx":0000
         ScrollTrack     =   -1  'True
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
         BackColorFrozen =   5263440
         ForeColorFrozen =   4210752
         WallPaperAlignment=   9
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   6720
         TabIndex        =   48
         Top             =   120
         Width           =   540
      End
      Begin VB.Label lblXmlRazao 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   3720
         TabIndex        =   47
         Top             =   120
         Width           =   2895
      End
      Begin VB.Label lblxmlFornecedor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fornecedor: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   2640
         TabIndex        =   46
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblSerieNota 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   2160
         TabIndex        =   45
         Top             =   120
         Width           =   345
      End
      Begin VB.Label lblxmlSerie 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Série:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1560
         TabIndex        =   44
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblNumNota 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   720
         TabIndex        =   43
         Top             =   120
         Width           =   645
      End
      Begin VB.Label lblxmlNota 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nota:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Frame frameCadastraReferencia 
      BackColor       =   &H00393939&
      BorderStyle     =   0  'None
      Height          =   4080
      Left            =   3000
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   8115
      Begin VB.TextBox txtFiltro 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1260
         TabIndex        =   34
         Top             =   900
         Width           =   6690
      End
      Begin VB.ComboBox cmbFornecedor 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1260
         TabIndex        =   33
         Top             =   510
         Width           =   6690
      End
      Begin VB.TextBox txtCodBarrasAdd 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   150
         MaxLength       =   20
         TabIndex        =   19
         ToolTipText     =   "Código do Produto"
         Top             =   3585
         Width           =   2265
      End
      Begin VB.TextBox txtReferenciaAdd 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2550
         MaxLength       =   7
         TabIndex        =   20
         ToolTipText     =   "Referência"
         Top             =   3585
         Width           =   2265
      End
      Begin CentroDeDistribuicao.chameleonButton cmdAddCodBarra 
         Height          =   330
         Left            =   5055
         TabIndex        =   21
         Top             =   3585
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         BTYPE           =   14
         TX              =   "&Gravar Código"
         ENAB            =   0   'False
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
         MICON           =   "frmCheckin.frx":00CD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CentroDeDistribuicao.chameleonButton cmdSairCadastroReferencia 
         Height          =   330
         Left            =   6540
         TabIndex        =   23
         Top             =   3585
         Width           =   1395
         _ExtentX        =   2461
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
         MICON           =   "frmCheckin.frx":00E9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VSFlex7DAOCtl.VSFlexGrid grdReferenciaSemBarras 
         Height          =   1905
         Left            =   150
         TabIndex        =   17
         Top             =   1305
         Width           =   7815
         _cx             =   13785
         _cy             =   3360
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
         ForeColorSel    =   8388608
         BackColorBkg    =   3750201
         BackColorAlternate=   3947580
         GridColor       =   5263440
         GridColorFixed  =   8421504
         TreeColor       =   3947580
         FloodColor      =   5263440
         SheetBorder     =   3947580
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
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
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmCheckin.frx":0105
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
         ForeColorFrozen =   4210752
         WallPaperAlignment=   9
      End
      Begin VB.Label lblFiltro 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   150
         TabIndex        =   39
         Top             =   960
         Width           =   435
      End
      Begin VB.Label Label7 
         BackColor       =   &H00393939&
         Caption         =   "Fornecedor"
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
         Height          =   255
         Left            =   150
         TabIndex        =   35
         Top             =   555
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00393939&
         Caption         =   "Adicionar referência a código de barras"
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
         TabIndex        =   25
         Top             =   150
         Width           =   2655
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Referência"
         ForeColor       =   &H00C0C0C0&
         Height          =   300
         Left            =   2550
         TabIndex        =   24
         Top             =   3315
         Width           =   855
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Código do Produto"
         ForeColor       =   &H00C0C0C0&
         Height          =   300
         Left            =   150
         TabIndex        =   22
         Top             =   3315
         Width           =   1605
      End
   End
   Begin VB.Frame frmFornecedor 
      BackColor       =   &H00393939&
      BorderStyle     =   0  'None
      Caption         =   "Checkin de Mercadoria"
      Height          =   1815
      Left            =   150
      TabIndex        =   36
      Top             =   1080
      Visible         =   0   'False
      Width           =   5775
      Begin VSFlex7DAOCtl.VSFlexGrid grdFornecedores 
         Height          =   1305
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   5535
         _cx             =   9763
         _cy             =   2302
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
         ForeColorSel    =   8388608
         BackColorBkg    =   3750201
         BackColorAlternate=   3947580
         GridColor       =   5263440
         GridColorFixed  =   8421504
         TreeColor       =   3947580
         FloodColor      =   5263440
         SheetBorder     =   3947580
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
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
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmCheckin.frx":017D
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
         ForeColorFrozen =   4210752
         WallPaperAlignment=   9
      End
      Begin VB.Label lblFornecedor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Escolha um Fornecedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   120
         Width           =   2070
      End
   End
   Begin VB.Timer tmeCodigoBarraSetFocus 
      Left            =   120
      Top             =   6960
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   44
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   14880
      TabIndex        =   18
      Top             =   6810
      Width           =   14880
   End
   Begin VB.Frame frameCodigoDeBarras 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1320
      Left            =   6150
      TabIndex        =   11
      Top             =   150
      Width           =   8895
      Begin VB.CheckBox chkModoSubtrair 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "&Modo Subtrair"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   7350
         TabIndex        =   27
         Top             =   120
         Width           =   1320
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         TabIndex        =   14
         Top             =   435
         Width           =   8640
         Begin VB.TextBox txtCodigobarras 
            Appearance      =   0  'Flat
            BackColor       =   &H00303030&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   630
            Left            =   180
            MaxLength       =   20
            TabIndex        =   29
            Text            =   "50*45645612312341564"
            ToolTipText     =   "Código de Barras"
            Top             =   120
            Width           =   8175
         End
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         Caption         =   "Códigos de Barras"
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   1290
      End
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdProduto 
      Height          =   4740
      Left            =   150
      TabIndex        =   15
      Top             =   1920
      Width           =   14880
      _cx             =   26247
      _cy             =   8361
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
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
      ForeColorSel    =   8388608
      BackColorBkg    =   5263440
      BackColorAlternate=   3947580
      GridColor       =   5263440
      GridColorFixed  =   -2147483632
      TreeColor       =   3947580
      FloodColor      =   5263440
      SheetBorder     =   3947580
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmCheckin.frx":01E8
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
      ForeColorFrozen =   4210752
      WallPaperAlignment=   9
   End
   Begin VB.Frame frameCapa 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   1320
      Left            =   150
      TabIndex        =   7
      Top             =   150
      Width           =   6015
      Begin VB.TextBox txtSerie 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   5040
         MaxLength       =   2
         TabIndex        =   2
         ToolTipText     =   "Série"
         Top             =   510
         Width           =   570
      End
      Begin VB.TextBox txtNotafiscal 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   1
         ToolTipText     =   "Nota Fiscal"
         Top             =   510
         Width           =   735
      End
      Begin VB.TextBox txtFornecedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   720
         MaxLength       =   15
         TabIndex        =   0
         ToolTipText     =   "Fornecedor"
         Top             =   510
         Width           =   1335
      End
      Begin VB.TextBox txtLoja 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   195
         MaxLength       =   3
         TabIndex        =   28
         ToolTipText     =   "Loja"
         Top             =   510
         Width           =   435
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   120
         TabIndex        =   13
         Top             =   435
         Width           =   5595
      End
      Begin VB.Label lblNota 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00303030&
         BackStyle       =   0  'Transparent
         Caption         =   "Nota Fiscal não Localizada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3165
         TabIndex        =   32
         Top             =   975
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblNomeFantasia 
         BackColor       =   &H00404040&
         Caption         =   "Fornecedor: TESTE STESTE"
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   150
         TabIndex        =   31
         Top             =   975
         Width           =   3330
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404040&
         Caption         =   "CNPJ"
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   720
         TabIndex        =   30
         Top             =   120
         Width           =   450
      End
      Begin VB.Label Label6 
         BackColor       =   &H00404040&
         Caption         =   "Série"
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   5040
         TabIndex        =   10
         Top             =   120
         Width           =   570
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404040&
         Caption         =   "NF"
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   3960
         TabIndex        =   9
         Top             =   120
         Width           =   450
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         Caption         =   "Loja"
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   150
         TabIndex        =   8
         Top             =   120
         Width           =   450
      End
   End
   Begin CentroDeDistribuicao.chameleonButton cmdComparar 
      Height          =   510
      Left            =   9300
      TabIndex        =   3
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "&Comparar NF"
      ENAB            =   0   'False
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
      MICON           =   "frmCheckin.frx":0262
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
      Left            =   13620
      TabIndex        =   6
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
      MICON           =   "frmCheckin.frx":027E
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
      TabIndex        =   5
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "&Limpa"
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
      MICON           =   "frmCheckin.frx":029A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdEncerra 
      Height          =   510
      Left            =   10740
      TabIndex        =   4
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "&Encerra"
      ENAB            =   0   'False
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
      MICON           =   "frmCheckin.frx":02B6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdProdutoTitulo 
      Height          =   300
      Left            =   150
      TabIndex        =   26
      Top             =   1665
      Width           =   14895
      _cx             =   26282
      _cy             =   529
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   0   'False
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
      ForeColorSel    =   8388608
      BackColorBkg    =   5263440
      BackColorAlternate=   3947580
      GridColor       =   5263440
      GridColorFixed  =   -2147483632
      TreeColor       =   3947580
      FloodColor      =   5263440
      SheetBorder     =   3947580
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmCheckin.frx":02D2
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
      ForeColorFrozen =   4210752
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "frmCheckin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim item As Integer
Dim wDescricao As String
Dim tempoRestante, tempoRestantePadrao As String

Dim edita As Boolean

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
x As Long
y As Long
End Type
Dim rsFornecedor  As New ADODB.Recordset
Dim posicao As Printer
Dim codigoFornecedor As String
Dim codFornecedorXML As String
Dim pos As Integer
Dim codBarra As String
Dim qnt As Integer

Private Sub cmbFornecedor_Click()
    carregaProdutoSemBarras
    txtReferenciaAdd.Text = ""
    cmdAddCodBarra.Enabled = False
    LblNomeFantasia.Caption = "Fornecedor: "
End Sub

Private Sub cmbFornecedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    carregaProdutoSemBarras
    txtReferenciaAdd.Text = ""
    cmdAddCodBarra.Enabled = False
End If

End Sub

Private Sub Form_Load()
   txtLoja.Text = LojaINI
   carregarPosicaoFrame frameCadastraReferencia
   carregarPosicaoFrame fraNota
   
   carregarPosicaoTamanhoTela Me
   'JanelaTOP Me
   tempoRestantePadrao = "00:00:05"
   txtCodigobarras.Text = ""
   LblNomeFantasia.Caption = ""
   tempoRestante = tempoRestantePadrao
   'txtFornecedor.SetFocus
   'frmControleCD.Enabled = False
   fraNota.Visible = False
   Montagrid 2
End Sub

Private Sub chkModoSubtrair_Click()
    If chkModoSubtrair Then
        If grdProduto.Rows = 0 Then
            chkModoSubtrair.Value = False
        End If
        txtCodigobarras.ForeColor = vbRed
        txtCodigobarras.Text = ""
    Else
        txtCodigobarras.ForeColor = vbWhite
        txtCodigobarras.Text = ""
    End If
    txtCodigobarras.SetFocus
End Sub

Private Sub cmdAddCodBarra_Click()
    
'       SQL = "Insert Into ProdutoBarras Values ('" & Trim(txtReferenciaAdd.Text) & "', '" & Trim(txtCodBarrasAdd.Text) & _
'       "', '" & txtFornecedor.Text & "', '1', '', 'B')" & vbNewLine & _
'       "update chekinMercadoria set chm_referencia = '" & txtReferenciaAdd & "' where chm_Fornecedor = " _
'       & txtFornecedor & " and chm_CodigoBarras = '" & txtCodBarrasAdd & "'"
       Screen.MousePointer = 11
       sql = "SP_inserir_Referencia_CodigoBarras_Chekin '" & _
       Trim(txtReferenciaAdd.Text) & "','" & Trim(txtCodBarrasAdd.Text) & "','" & codigoFornecedor & "'"
        
       ADO_Cn_CDLocal.Execute (sql)
       ativaCadastraReferencia False
       Screen.MousePointer = 0
                
End Sub


Private Sub cmdComparar_Click()
       Screen.MousePointer = 11
        
       limpaGrid grdProduto
       Dim i, j As Integer
       i = 1
       Montagrid 1
       
       sql = "exec SP_ComparaChekinNFCompras '" & Trim(txtLoja.Text) & "','" & txtNotaFiscal _
       & "','" & txtSerie & "','" & codigoFornecedor & "'"
    
       'ADO_Cn_CD.BeginTrans
           ADO_Cn_CDLocal.Execute (sql)
        'ADO_Cn_CD.CommitTrans
    
      ' Itens do Chekin
       
       sql = "Select pr_referencia, CHM_CodigoBarras, CHM_QuantidadeChekin, CHM_QuantidadeNF, pr_descricao, ci_codigoBarra, ci_descricaoFornecedor ," & _
       "CHM_ChekinOK from chekinmercadoria, produto,itemnfcompra where chm_loja = '" & txtLoja.Text & _
       "' and chm_fornecedor = '" & codigoFornecedor & "' and chm_notafiscal = '" & txtNotaFiscal.Text & _
       "' and chm_serie = '" & txtSerie.Text & "' and pr_referencia = chm_referencia " & _
       "  and ci_notafiscal = chm_notafiscal" & _
       "  and ci_serie = chm_serie" & _
       "  and ci_fornecedor = chm_fornecedor" & _
       "  and ci_referencia = chm_referencia" & _
       "  order by chm_item"
              
       adoDemeoChekinChekin2.CursorLocation = adUseClient
       adoDemeoChekinChekin2.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
       
       Do While Not adoDemeoChekinChekin2.EOF
                 
           grdProduto.AddItem adoDemeoChekinChekin2("pr_referencia") & Chr(9) & adoDemeoChekinChekin2("CHM_CodigoBarras") & Chr(9) & _
           adoDemeoChekinChekin2("pr_descricao") & Chr(9) & adoDemeoChekinChekin2("CHM_QuantidadeChekin") & _
           Chr(9) & adoDemeoChekinChekin2("chm_quantidadeNF") & Chr(9) _
           & adoDemeoChekinChekin2("CI_CodigoBarra") & Chr(9) _
           & adoDemeoChekinChekin2("CI_DescricaoFornecedor")
                       
           grdProduto.row = i
               For j = 0 To 6
                   grdProduto.col = j
                   grdProduto.CellForeColor = vbBlack
                 
                   If adoDemeoChekinChekin2("CHM_ChekinOK") = "N" Then
                       grdProduto.CellForeColor = vbRed
                       cmdEncerra.Enabled = False
                   End If
               Next j
           i = i + 1
                       
           adoDemeoChekinChekin2.MoveNext
        Loop
        'End If
        
        adoDemeoChekinChekin2.Close
        
        ' Carrega Itens sem Referencia
        
           sql = "Select * from chekinmercadoria, itemnfcompra where chm_loja = '" & txtLoja.Text & _
          "' and chm_fornecedor = '" & codigoFornecedor & "' and chm_notafiscal = '" & txtNotaFiscal.Text & _
          "' and chm_serie = '" & txtSerie.Text & "' and chm_referencia = ''" & _
          "  and ci_notafiscal = chm_notafiscal and ci_serie = chm_serie and ci_fornecedor = chm_fornecedor" & _
          " order by chm_item"
    
    adoDemeoChekinChekin2.CursorLocation = adUseClient
    adoDemeoChekinChekin2.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
    If Not adoDemeoChekinChekin2.EOF Then
    wDescricao = ""
        Do While Not adoDemeoChekinChekin2.EOF
              
              grdProduto.AddItem "" & Chr(9) & adoDemeoChekinChekin2("CHM_CodigoBarras") & Chr(9) & _
              adoDemeoChekinChekin2("pr_descricao") & Chr(9) & adoDemeoChekinChekin2("CHM_QuantidadeChekin") & _
              Chr(9) & adoDemeoChekinChekin2("chm_quantidadeNF") & Chr(9) & _
              adoDemeoChekinChekin2("CI_CodigoBarra") & Chr(9) & _
              adoDemeoChekinChekin2("CI_DescricaoFornecedor")
        
              grdProduto.row = i
                    For j = 0 To 6
                           grdProduto.col = j
                           grdProduto.CellForeColor = vbBlack
                 
                        If adoDemeoChekinChekin2("CHM_ChekinOK") = "N" Then
                           grdProduto.CellForeColor = vbRed
                        End If
                    Next j
                 i = i + 1
                       
              adoDemeoChekinChekin2.MoveNext
        Loop
    
    End If

    adoDemeoChekinChekin2.Close
    
    ' Itens Esquecidos
    
    sql = "select CI_Referencia,PR_CodigoBarra,PR_Descricao, 0, CI_Quantidade, CI_CodigoBarra, CI_DescricaoFornecedor" _
     & " From ItemNFCompra, produto" _
     & " Where PR_Referencia = CI_Referencia" _
     & " and CI_NotaFiscal = '" & txtNotaFiscal.Text & "'" _
     & " and CI_Serie = '" & txtSerie.Text & "'" _
     & " and CI_Fornecedor = '" & codigoFornecedor & "'" _
     & " and CI_referencia not in (select CHM_referencia" _
                                & " From ChekinMercadoria" _
                                & " where CHM_NotaFiscal = '" & txtNotaFiscal.Text & "'" _
                                & " and CHM_Serie = '" & txtSerie.Text & "'" _
                                & " and CHM_Fornecedor = '" & codigoFornecedor & "')"
    
    adoDemeoChekinChekin2.CursorLocation = adUseClient
    adoDemeoChekinChekin2.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not adoDemeoChekinChekin2.EOF Then
        Do While Not adoDemeoChekinChekin2.EOF
            
            grdProduto.AddItem adoDemeoChekinChekin2("ci_referencia") & Chr(9) _
            & adoDemeoChekinChekin2("pr_codigoBarra") & Chr(9) _
            & adoDemeoChekinChekin2("pr_descricao") & Chr(9) _
            & 0 & Chr(9) _
            & adoDemeoChekinChekin2("ci_quantidade") & Chr(9) _
            & adoDemeoChekinChekin2("ci_codigoBarra") & Chr(9) _
            & adoDemeoChekinChekin2("ci_DescricaoFornecedor")
            
            grdProduto.row = i
               For j = 0 To 6
                   grdProduto.col = j
                       grdProduto.CellForeColor = vbRed
                       cmdEncerra.Enabled = False
               Next j
           i = i + 1
        
            adoDemeoChekinChekin2.MoveNext
        Loop
        End If
        adoDemeoChekinChekin2.Close
        
        
        
    'Item a mais
    
    sql = "select * " _
        & " from ChekinMercadoria,produto " _
        & " where chm_notafiscal = '" & txtNotaFiscal.Text & "'" _
        & " and pr_referencia = chm_referencia" _
        & " and chm_serie = '" & txtSerie.Text & "'" _
        & " and chm_fornecedor = '" & codigoFornecedor & "'" _
        & " and chm_referencia not in (select ci_referencia from itemnfcompra where ci_notafiscal = '" & txtNotaFiscal.Text & "'" & " and ci_serie = '" & txtSerie.Text & "'" & "and ci_fornecedor = '" & codigoFornecedor & "')"
    
    adoDemeoChekinChekin2.CursorLocation = adUseClient
    adoDemeoChekinChekin2.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not adoDemeoChekinChekin2.EOF Then
        Do While Not adoDemeoChekinChekin2.EOF
            
            grdProduto.AddItem adoDemeoChekinChekin2("chm_referencia") & Chr(9) _
            & adoDemeoChekinChekin2("chm_codigoBarras") & Chr(9) _
            & adoDemeoChekinChekin2("pr_descricao") & Chr(9) _
            & adoDemeoChekinChekin2("chm_quantidadeChekin") & Chr(9) _
            & 0 & Chr(9) _
            & Chr(9) _
            & " "
            
            
            grdProduto.row = i
               For j = 0 To 6
                   grdProduto.col = j
                       grdProduto.CellForeColor = vbRed
                       cmdEncerra.Enabled = False
               Next j
           i = i + 1
        
            adoDemeoChekinChekin2.MoveNext
        Loop
        End If
        adoDemeoChekinChekin2.Close
    
    
    If verificaChekinOK Then
        cmdEncerra.Enabled = True
    Else
        cmdEncerra.Enabled = False
    End If
        
    txtCodigobarras.SetFocus
    Screen.MousePointer = 0
End Sub


Private Sub cmdEncerra_Click()
    
    If Not verificaChekinOK Then
         MsgBox "Há número de chekin diferentes da quantidade do item", vbExclamation, "Checkin Mercadoria"
         cmdEncerra.Enabled = False
    Else
         'ADO_Cn_CD.BeginTrans
         tmeCodigoBarraSetFocus.Enabled = False
         sql = "Update chekinmercadoria set chm_situacao = 'E' " _
         & " where chm_loja = '" & txtLoja.Text & _
         "' and chm_fornecedor = '" & Mid(codigoFornecedor, 1, 4) & "' and chm_notafiscal = '" & txtNotaFiscal.Text & _
         "' and chm_serie = '" & txtSerie.Text & "'"
         ADO_Cn_CDLocal.Execute (sql)
         'ADO_Cn_CD.CommitTrans
    
         MsgBox "Checkin encerrado com sucesso!", vbInformation, "Checkin Mercadoria"
         
        LimpaTela
        ativaCheckin False
        Montagrid 2
        
    End If

End Sub


Private Sub cmdLimpa_Click()
    tmeCodigoBarraSetFocus.Enabled = False
    'Limpa Grid Fornecedores
    grdFornecedores.Rows = 1
    grdFornecedores.AddItem ""
    grdFornecedores.RemoveItem (1)
    Montagrid 2
    'Limpa Label
    LblNomeFantasia.Caption = ""
    
    'Desativa frmFornecedores
    frmFornecedor.Visible = False
    
    If mensagemLimparCampos Then
        LimpaTela
        ativaCheckin False
    Else
        If frameCapa.Enabled Then
            If txtFornecedor.Enabled Then
                txtFornecedor.SetFocus
            End If
        ElseIf frameCodigoDeBarras.Enabled Then
            txtCodigobarras.SetFocus
        End If
    End If
    tmeCodigoBarraSetFocus.Enabled = True
End Sub

Private Sub cmdRetorna_Click()
    Unload Me
End Sub


Private Sub cmdSairCadastroReferencia_Click()
    ativaCadastraReferencia False
End Sub






Private Sub grdFornecedores_DblClick()
    
    grdFornecedores_KeyPress 13
    
End Sub

Private Sub grdFornecedores_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
    
        frmFornecedor.Visible = False

    ElseIf KeyAscii = 13 Or KeyAscii = 9 Then
        If grdFornecedores.row <> 0 Then
            txtFornecedor = grdFornecedores.TextMatrix(grdFornecedores.row, 1)
            LblNomeFantasia.Caption = "Fornecedor: " & grdFornecedores.TextMatrix(grdFornecedores.row, 2)
            codigoFornecedor = grdFornecedores.TextMatrix(grdFornecedores.row, 0)
            frmFornecedor.Visible = False
        End If
    End If
    
    frameCapa.Enabled = True
    txtNotaFiscal.SetFocus
    
    'Limpa Grid Fornecedores
    grdFornecedores.Rows = 1
    grdFornecedores.AddItem ""
    grdFornecedores.RemoveItem (1)
    
    

End Sub

Private Sub grdProduto_Click()
    tempoRestante = tempoRestantePadrao
End Sub

Private Sub grdProduto_DblClick()
    If grdProduto.row >= grdProduto.FixedRows And grdProduto.TextMatrix(grdProduto.row, 0) = "" Then
        CarregaFornecedor
        ativaCadastraReferencia True
        txtCodBarrasAdd.Text = Trim(grdProduto.TextMatrix(grdProduto.row, 1))
        
    End If
End Sub

Private Sub ativaCadastraReferencia(ativa As Boolean)

    If ativa Then
        frameCadastraReferencia.Visible = True
        carregaProdutoSemBarras
        ativaCheckin False
        frameCapa.Enabled = False
        cmdLimpa.Enabled = False
        
        grdProduto.col = 1
        If grdProduto = "" Then
            txtCodBarrasAdd = ""
            txtCodBarrasAdd.SetFocus
        Else
            txtCodBarrasAdd = grdProduto
            txtReferenciaAdd = ""
            'txtReferenciaAdd.SetFocus
        End If
        
    Else
        frameCadastraReferencia.Visible = False
        ativaCheckin True
        cmdLimpa.Enabled = True
        'txtCodigobarras.SetFocus
        Call CarregaGrid
    End If
    
End Sub

Private Sub grdProduto_keyup(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        grdProduto.col = 1
        If mensagemExluir("o item " & Trim(grdProduto)) Then
            sql = "delete chekinmercadoria where chm_loja = '" & txtLoja & "' and chm_fornecedor = '" & codigoFornecedor & _
            "' and chm_notaFiscal = '" & txtNotaFiscal & "' and chm_serie = '" & txtSerie & "' AND CHM_CODIGOBARRAS = '" & Trim(grdProduto) & "'"
            'ADO_Cn_CD.BeginTrans
            ADO_Cn_CDLocal.Execute (sql)
            'ADO_Cn_CD.CommitTrans
            CarregaGrid
        End If
    End If
    tempoRestante = tempoRestantePadrao
    'tmeCodigoBarraSetFocus.Enabled
    If KeyCode = 112 Then
        tmeCodigoBarraSetFocus.Enabled = False
        fraNota.Visible = True
        ConsultaXML txtNotaFiscal.Text, txtSerie.Text, codigoFornecedor
        grdXml.SetFocus
        txtCodigobarras.Enabled = False
    End If
End Sub

Private Sub grdProduto_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      If (grdProduto.MouseCol = 0 Or grdProduto.MouseCol = 4) Then
            If grdProduto.MouseRow >= grdProduto.FixedRows And grdProduto.MouseRow < grdProduto.Rows Then
                grdProduto.ToolTipText = grdProduto.TextMatrix(grdProduto.MouseRow, grdProduto.MouseCol)
            End If
        ElseIf grdProduto.MouseCol <> 0 Or grdProduto.MouseCol <> 4 Then
            grdProduto.ToolTipText = ""
        End If
End Sub

Private Sub grdReferenciaSemBarras_Click()
    txtReferenciaAdd.Text = Trim(grdReferenciaSemBarras.TextMatrix(grdReferenciaSemBarras.row, 0))
    If txtReferenciaAdd <> "" Then
        cmdAddCodBarra.Enabled = True
    End If
End Sub

Private Sub txtCFOP_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNumerico(KeyAscii)
    If KeyAscii = 13 Then
        If verificarCamposObrigCapa Then
            ativaCheckin True
        End If
    End If
End Sub




Private Sub grdXml_AfterEdit(ByVal row As Long, ByVal col As Long)

    Dim sql As String
    
    If vbKeyReturn Then
    
            sql = "update itemnfcompra set ci_codigoBarra = '" & grdXml.TextMatrix(grdXml.row, 4) _
            & "' where ci_notafiscal = '" & txtNotaFiscal.Text & "' and ci_serie = '" & txtSerie.Text _
            & "' and ci_item = " & grdXml.TextMatrix(grdXml.row, 0)
        
            ADO_Cn_CDLocal.Execute (sql)
            
            sql = "update itemnfcompra set ci_referencia = '" & grdXml.TextMatrix(grdXml.row, 1) _
            & "' where ci_notafiscal = '" & txtNotaFiscal.Text & "' and ci_serie = '" & txtSerie.Text _
            & "' and ci_item = " & grdXml.TextMatrix(grdXml.row, 0)
        
            ADO_Cn_CDLocal.Execute (sql)


    End If
    ConsultaXML txtNotaFiscal.Text, txtSerie.Text, codigoFornecedor

End Sub


Private Sub grdXml_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    fraNota.Visible = False
    txtCodigobarras.Enabled = True
    txtCodigobarras.SetFocus
    tmeCodigoBarraSetFocus.Enabled = True
End If

End Sub

Private Sub grdXml_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    
    deletaItemXML Trim(lblNumNota.Caption), grdXml.TextMatrix(grdXml.row, 0), grdXml.TextMatrix(grdXml.row, 5), Trim(lblSerieNota.Caption)
    ConsultaXML txtNotaFiscal.Text, txtSerie.Text, codigoFornecedor
End If
End Sub

Private Sub tmeCodigoBarraSetFocus_Timer()
    contagemRegressiva
End Sub



Private Sub txtCnpj_KeyPress(KeyAscii As Integer)
    Dim sql As String
    
    If KeyAscii = 13 Then
        If MsgBox("Confirma a alteração do CNPJ do Fornecedor?", vbYesNo + vbQuestion + vbDefaultButton2, "Alteração de Fornecedor") = vbYes Then
            sql = " update fornecedor set fo_cgc = '" & txtCNPJ.Text & "' where fo_codigoFornecedor = " & codFornecedorXML
            ADO_Cn_CDLocal.Execute (sql)
        End If
    Else
            
        If KeyAscii = 27 Then
                fraNota.Visible = False
                txtCodigobarras.Enabled = True
                tmeCodigoBarraSetFocus.Enabled = False
                txtCodigobarras.SetFocus
                tmeCodigoBarraSetFocus.Enabled = True
        End If
    End If
End Sub

Private Sub txtCodigobarras_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim adoDemeoChekin As New ADODB.Recordset

    If KeyCode = 13 Then
        
        If chkModoSubtrair.Value = False Then
        pos = InStr(txtCodigobarras.Text, "*")
        If pos > 0 Then
        codBarra = Mid(txtCodigobarras.Text, pos + 1, Len(txtCodigobarras.Text))
        qnt = Mid(txtCodigobarras.Text, 1, pos - 1)
        Else
        
        qnt = 1
        codBarra = txtCodigobarras.Text
        
        End If
        

            Dim sql As String
            
            sql = "Select * from produtobarras, produto where prb_codigobarras = '" & codBarra & _
            "' and pr_referencia = prb_referencia "
            adoDemeoChekin.CursorLocation = adUseClient '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            adoDemeoChekin.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
            
            If adoDemeoChekin.EOF Then
                If txtCodigobarras = "" Then
                    Exit Sub
                End If
            
                sql = "Select * from chekinMercadoria where chm_loja = '" & txtLoja.Text & "' and chm_Fornecedor = " _
                & codigoFornecedor & " and chm_CodigoBarras = '" & codBarra & "'"
                adoDemeoChekin5.CursorLocation = adUseClient
                adoDemeoChekin5.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
              
                If adoDemeoChekin5.EOF Then
                                  sql = "Insert into chekinmercadoria (chm_loja, chm_fornecedor, chm_notafiscal, chm_serie, " _
                    & " chm_item, chm_codigobarras, chm_referencia , CHM_QuantidadeNF, CHM_QuantidadeChekin, chm_tipopedido, chm_situacao) " _
                    & " values ('" & txtLoja.Text & "', " & codigoFornecedor & ", '" & txtNotaFiscal.Text & "" _
                    & "', '" & txtSerie.Text & "', '" & item & "', '" & codBarra & "" _
                    & "', '', 0, " & qnt & ", 0, '')"
                Else
                    sql = "Update chekinmercadoria set CHM_QuantidadeChekin = CHM_QuantidadeChekin +" & qnt & " ," _
                    & " chm_item = '" & item & "' " _
                    & " where chm_loja = '" & txtLoja.Text & _
                    "' and chm_fornecedor = '" & Mid(codigoFornecedor, 1, 4) & "' and chm_notafiscal = '" & txtNotaFiscal.Text & _
                    "' and chm_serie = '" & txtSerie.Text & "' and chm_CodigoBarras = '" & codBarra & "'"
                End If
                adoDemeoChekin5.Close
            
                'ADO_Cn_CDLocal.BeginTrans
                ADO_Cn_CDLocal.Execute (sql)
                'ADO_Cn_CDLocal.CommitTrans
                
                txtCodigobarras.SetFocus
                txtCodigobarras.Text = ""
                
                Call CarregaGrid
                
                adoDemeoChekin.Close
            
                Exit Sub
            End If
        
            sql = "Select max(chm_item) as item from chekinmercadoria where chm_loja = '" & txtLoja.Text & _
            "' and chm_fornecedor = '" & Mid(codigoFornecedor, 1, 4) & "' and chm_notafiscal = '" & txtNotaFiscal.Text & _
            "' and chm_serie = '" & txtSerie.Text & "'"
            adoDemeoChekinChekin2.CursorLocation = adUseClient
            adoDemeoChekinChekin2.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
            If IsNull(adoDemeoChekinChekin2("item")) Then
                item = 1
            Else
                item = adoDemeoChekinChekin2("item") + 1
            End If
    
            adoDemeoChekinChekin2.Close
        
            sql = "Select count(*) numeroItem from chekinmercadoria where chm_loja = '" & txtLoja.Text & _
            "' and chm_fornecedor = '" & Mid(codigoFornecedor, 1, 4) & "' and chm_notafiscal = '" & txtNotaFiscal.Text & _
            "' and chm_serie = '" & txtSerie.Text & "' and chm_referencia = '" & adoDemeoChekin("pr_referencia") & "'"
            adoDemeoChekinChekin2.CursorLocation = adUseClient
            adoDemeoChekinChekin2.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
            If adoDemeoChekinChekin2("numeroItem") = 0 Then
                ADO_Cn_CDLocal.BeginTrans
                   sql = "Insert into chekinmercadoria (chm_loja, chm_fornecedor, chm_notafiscal, chm_serie, " _
                       & " chm_item, chm_codigobarras, chm_referencia , CHM_QuantidadeNF, CHM_QuantidadeChekin, chm_tipopedido, chm_situacao) " _
                       & " values ('" & txtLoja.Text & "', " & Mid(codigoFornecedor, 1, 4) & ", '" & txtNotaFiscal.Text & "" _
                       & "', '" & txtSerie.Text & "', '" & item & "', '" & codBarra & "" _
                       & "', '" & adoDemeoChekin("pr_referencia") & "', 0, " & qnt & ", 0, 'A')"
                ADO_Cn_CDLocal.Execute (sql)
                ADO_Cn_CDLocal.CommitTrans
            Else
                 ADO_Cn_CDLocal.BeginTrans
                   sql = "Update chekinmercadoria set CHM_QuantidadeChekin = CHM_QuantidadeChekin + " & qnt & ", " _
                   & " chm_item = '" & item & "' " _
                       & " where chm_loja = '" & txtLoja.Text & _
                     "' and chm_fornecedor = '" & Mid(codigoFornecedor, 1, 4) & "' and chm_notafiscal = '" & txtNotaFiscal.Text & _
                     "' and chm_serie = '" & txtSerie.Text & "' and chm_referencia = '" & adoDemeoChekin("pr_referencia") & "'"
                ADO_Cn_CDLocal.Execute (sql)
                ADO_Cn_CDLocal.CommitTrans
            End If
        
            adoDemeoChekinChekin2.Close
            adoDemeoChekin.Close
        Else
    
            sql = "Select chm_quantidadechekin from chekinmercadoria where chm_loja = '" & txtLoja & _
            "' and chm_fornecedor = '" & codigoFornecedor & "' and chm_notafiscal = '" & txtNotaFiscal & _
            "' and chm_serie = '" & txtSerie & "' and chm_CODIGOBARRAS = '" & txtCodigobarras & "' " & _
            "and chm_quantidadeChekin > 0"
            adoDemeoChekinChekin2.CursorLocation = adUseClient
            adoDemeoChekinChekin2.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
            
            If adoDemeoChekinChekin2("chm_quantidadechekin") = 1 Then
              
            Else
                ADO_Cn_CDLocal.BeginTrans
                sql = "Update chekinmercadoria set CHM_QuantidadeChekin = CHM_QuantidadeChekin - 1 " _
                & " where chm_loja = '" & txtLoja.Text & _
                "' and chm_fornecedor = '" & Mid(codigoFornecedor, 1, 4) & "' and chm_notafiscal = '" & txtNotaFiscal.Text & _
                "' and chm_serie = '" & txtSerie.Text & "' and chm_CODIGOBARRAS = '" & txtCodigobarras & "'"
                ADO_Cn_CDLocal.Execute (sql)
                ADO_Cn_CDLocal.CommitTrans
            End If
            
            adoDemeoChekinChekin2.Close
        
        End If
        
        CarregaGrid
        cmdEncerra.Enabled = False
    
        campoSelecionadoComCaracter txtCodigobarras
        txtCodigobarras.SetFocus
        
    End If
    
End Sub


Private Sub txtCodBarrasAdd_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNumerico(KeyAscii)
End Sub

Private Sub txtNumeroPedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 10 Then
        txtCodigobarras.SetFocus
    End If
End Sub


Private Sub CarregaGrid()
Dim MSG As String
Dim j, i As Integer
    MSG = ""
    
    Montagrid 2
    
   sql = "Select pr_referencia,CHM_CodigoBarras,CHM_QuantidadeChekin,CHM_QuantidadeNF,pr_descricao" & _
    " from chekinmercadoria, produto where chm_loja = '" & txtLoja.Text & _
    "' and chm_fornecedor = '" & Mid(codigoFornecedor, 1, 4) & "' and chm_notafiscal = '" & txtNotaFiscal.Text & _
    "' and chm_serie = '" & txtSerie.Text & "' and pr_referencia = chm_referencia order by chm_item desc"
              
    adoDemeoChekinChekin2.CursorLocation = adUseClient
    adoDemeoChekinChekin2.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    limpaGrid grdProduto
        
    If Not adoDemeoChekinChekin2.EOF Then

        Do While Not adoDemeoChekinChekin2.EOF
              
              'If left$(adoDemeoChekinChekin2("pr_referencia"), 3) <> codigofornecedor Then
                'msg = "(Referência não corresponde há este fornecedor)"
              'Else
                MSG = ""
              'End If
              
              grdProduto.AddItem adoDemeoChekinChekin2("pr_referencia") & Chr(9) & adoDemeoChekinChekin2("CHM_CodigoBarras") & Chr(9) & _
              adoDemeoChekinChekin2("CHM_QuantidadeChekin") & Chr(9) & adoDemeoChekinChekin2("CHM_QuantidadeNF") & _
              Chr(9) & adoDemeoChekinChekin2("pr_descricao")
                       
             grdProduto.row = i
               For j = 0 To 4
                   grdProduto.col = j
                   grdProduto.CellForeColor = &HE0E0E0
                 
                   If MSG <> "" Then
                       grdProduto.CellForeColor = vbRed
                   End If
               Next j
           i = i + 1
                       
                       
                      
                       
              adoDemeoChekinChekin2.MoveNext
        Loop
    
    End If

    adoDemeoChekinChekin2.Close
    
     sql = "Select * from chekinmercadoria where chm_loja = '" & txtLoja.Text & _
              "' and chm_fornecedor = '" & Mid(codigoFornecedor, 1, 4) & "' and chm_notafiscal = '" & txtNotaFiscal.Text & _
              "' and chm_serie = '" & txtSerie.Text & "' and chm_referencia = ''"
    
    adoDemeoChekinChekin2.CursorLocation = adUseClient
    adoDemeoChekinChekin2.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
    If Not adoDemeoChekinChekin2.EOF Then
    wDescricao = ""
        Do While Not adoDemeoChekinChekin2.EOF
              
              grdProduto.AddItem "" & Chr(9) & adoDemeoChekinChekin2("CHM_CodigoBarras") & Chr(9) & _
              adoDemeoChekinChekin2("CHM_QuantidadeChekin") & Chr(9) & adoDemeoChekinChekin2("CHM_QuantidadeNF") & _
             Chr(9) & wDescricao
                       
              adoDemeoChekinChekin2.MoveNext
        Loop
    
    End If

    adoDemeoChekinChekin2.Close
    
'    if txtCodigobarras
    grdProduto.row = grdProduto.Rows - 1
    
    cmdComparar.Enabled = Not itemSemReferencia
    
End Sub


Private Sub LimpaTela()
    'txtLoja.Text = ""
    txtFornecedor.Text = ""
    txtNotaFiscal.Text = ""
    txtSerie.Text = ""
    'txtCFOP.Text = ""
    txtCodigobarras.Text = ""
    cmdEncerra.Enabled = False
    cmdComparar.Enabled = False
   'limpaGrid grdProduto
   
    grdProduto.Rows = 1
    grdProduto.AddItem ""
    grdProduto.RemoveItem (1)
   
    LblNomeFantasia.Caption = ""
    lblNota.Caption = ""
End Sub

Private Sub carregaProdutoSemBarras()
    
    Dim sql As String
    Dim adoProdutoSemBarras As New ADODB.Recordset

    ' Produtos sem Barra
    
    sql = "Select PR_Referencia,PR_Descricao,PR_CodigoFornecedor from produto where PR_codigoFornecedor = " & Mid(cmbFornecedor.Text, 1, 4) _
                     & " and pr_referencia not in " _
                     & "(select PRB_referencia from produtobarras where prb_codigoFornecedor = " & Mid(cmbFornecedor.Text, 1, 4) _
                     & " and PRB_TipoCodigo = 'B')" _
                     & " and ( pr_descricao like '%" & txtFiltro.Text & "%'" _
                     & " or pr_referencia like '%" & txtFiltro.Text & "%'" _
                     & " or pr_complemento like '%" & txtFiltro.Text & "%'" _
                     & " or pr_CodigoProdutoNoFornecedor like '%" & txtFiltro.Text & "%')" _
                     & " order by PR_Descricao"
                                               
    adoProdutoSemBarras.CursorLocation = adUseClient
    adoProdutoSemBarras.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    limpaGrid grdReferenciaSemBarras
                    
    If Not adoProdutoSemBarras.EOF Then
        Do While Not adoProdutoSemBarras.EOF
                           
            grdReferenciaSemBarras.AddItem adoProdutoSemBarras("pr_referencia") & Chr(9) & " " & Chr(9) & adoProdutoSemBarras("pr_descricao")
                           
            adoProdutoSemBarras.MoveNext
        Loop
    
    End If
    adoProdutoSemBarras.Close
    
    ' Produtos com Barra
    
    sql = "Select PR_Referencia,PR_Descricao,PR_CodigoFornecedor, prb_codigoBarras from produto ,produtobarras where PR_codigoFornecedor = " & Mid(cmbFornecedor.Text, 1, 4) _
                     & " and pr_referencia = prb_referencia " _
                     & " and PRB_TipoCodigo = 'B'" _
                     & " and ( pr_descricao like '%" & txtFiltro.Text & "%'" _
                     & " or pr_referencia like '%" & txtFiltro.Text & "%'" _
                     & " or pr_complemento like '%" & txtFiltro.Text & "%'" _
                     & " or pr_CodigoProdutoNoFornecedor like '%" & txtFiltro.Text & "%')" _
                     & " order by PR_Descricao"
                                               
    adoProdutoSemBarras.CursorLocation = adUseClient
    adoProdutoSemBarras.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
                    
    If Not adoProdutoSemBarras.EOF Then
        Do While Not adoProdutoSemBarras.EOF
                           
            grdReferenciaSemBarras.AddItem adoProdutoSemBarras("pr_referencia") & Chr(9) & adoProdutoSemBarras("prb_CodigoBarras") & Chr(9) & adoProdutoSemBarras("pr_descricao")
                           
        adoProdutoSemBarras.MoveNext
        Loop
    
    End If
    adoProdutoSemBarras.Close
    
End Sub


Private Function verificarCamposObrigCapa() As Boolean

    verificarCamposObrigCapa = False

    If Not verificarCampoObrig(txtLoja) Then
        Exit Function
    ElseIf Not verificarCampoObrig(txtFornecedor) Then
        Exit Function
    ElseIf Not verificarCampoObrig(txtNotaFiscal) Then
        Exit Function
    ElseIf Not verificarCampoObrig(txtSerie) Then
        Exit Function
    Else
        verificarCamposObrigCapa = True
    End If
End Function


'////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////   CAMPOS DE TEXTO   /////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////

Private Sub ativaCheckin(ativa As Boolean)
    If ativa Then
        frameCapa.Enabled = False
        frameCodigoDeBarras.Enabled = True
        cmdComparar.Enabled = True
        'cmdEncerra.Enabled = True
        grdProduto.Enabled = True
        grdProdutoTitulo.Enabled = True
        'txtCodigobarras.SetFocus
        tmeCodigoBarraSetFocus.Interval = 1000
        tmeCodigoBarraSetFocus.Enabled = True
        CarregaGrid
    Else
        chkModoSubtrair.Value = False
        frameCapa.Enabled = True
        frameCodigoDeBarras.Enabled = False
        cmdComparar.Enabled = False
        grdProduto.Enabled = False
        grdProdutoTitulo.Enabled = False
        tmeCodigoBarraSetFocus.Enabled = False
        'txtFornecedor.SetFocus
    End If
End Sub

Private Sub txtCodBarrasAdd_GotFocus()
    campoSelecionadoComCaracter txtCodBarrasAdd
End Sub

Private Sub txtCodigobarras_KeyPress(KeyAscii As Integer)
If KeyAscii = 42 And chkModoSubtrair.Value = False Then
    KeyAscii = 42
Else
  '  KeyAscii = campoNumerico(KeyAscii)
End If
    
End Sub


Private Sub txtCodigobarras_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        tmeCodigoBarraSetFocus.Enabled = False
        fraNota.Visible = True
        ConsultaXML txtNotaFiscal.Text, txtSerie.Text, codigoFornecedor
        grdXml.SetFocus
        txtCodigobarras.Enabled = False
    End If
End Sub

Private Sub txtFiltro_LostFocus()
txtReferenciaAdd.Text = ""
cmdAddCodBarra.Enabled = False
carregaProdutoSemBarras
End Sub

Private Sub txtFornecedor_GotFocus()
    campoSelecionadoComCaracter txtFornecedor
End Sub

Private Sub txtFornecedor_KeyPress(KeyAscii As Integer)
    
    
    If KeyAscii = 13 Then
        txtNotafiscal_GotFocus
    Else
        KeyAscii = campoNumerico(KeyAscii)
        carregaChekin KeyAscii, txtNotaFiscal
    End If
    
End Sub

Private Sub txtNotafiscal_GotFocus()
    Dim quantidade As Integer
    
    campoSelecionadoComCaracter txtNotaFiscal
    
    ' Pesquisa Quantos Fornecedores existem com o mesmo CNPJ
    sql = "select count (*) as quantidade from fornecedor where fo_cgc like '%" & txtFornecedor & "%'"
    
    rsFornecedor.CursorLocation = adUseClient
    rsFornecedor.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    quantidade = rsFornecedor(quantidade)
    
    rsFornecedor.Close
    

    If quantidade = 1 Then

        codigoFornecedor = ""
        sql = "select fo_nomeFantasia,fo_codigofornecedor, fo_cgc from fornecedor where fo_cgc like '%" & txtFornecedor & "%'"
        rsFornecedor.CursorLocation = adUseClient
        rsFornecedor.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
        LblNomeFantasia.Caption = "Fornecedor: " & rsFornecedor("fo_nomeFantasia")
        codigoFornecedor = rsFornecedor("fo_codigofornecedor")
        txtFornecedor.Text = rsFornecedor("fo_cgc")
        
        rsFornecedor.Close
    
    ElseIf quantidade > 1 And LblNomeFantasia.Caption = "" Then
    
    
    frameCapa.Enabled = False
    
    ' Exibe Pop-up
        frmFornecedor.Visible = True
    
    ' Pesquisa CNPJ
        sql = " select fo_nomeFantasia,fo_codigofornecedor,fo_cgc from fornecedor where fo_cgc like '%" & txtFornecedor & "%'"
        rsFornecedor.CursorLocation = adUseClient
        rsFornecedor.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    
    
    'Preenche o Grid
        Do While Not rsFornecedor.EOF
            grdFornecedores.AddItem rsFornecedor("fo_codigoFornecedor") & Chr(9) & _
            rsFornecedor("fO_CGC") & Chr(9) & _
            rsFornecedor("fo_NomeFantasia")
            
            rsFornecedor.MoveNext
        Loop
        
    'Encerra o Cursor
        rsFornecedor.Close
        txtFornecedor.Text = ""
    
    ElseIf quantidade = 0 Then

        MsgBox "Fornecedor Não Cadastrado!", vbInformation, "Atenção!"
        txtFornecedor.SetFocus
        
    End If
End Sub

Private Sub txtNotaFiscal_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNumerico(KeyAscii)
    carregaChekin KeyAscii, txtNotaFiscal
End Sub

Private Sub txtReferenciaAdd_KeyPress(KeyAscii As Integer)
    proximoCampoEnter KeyAscii
    KeyAscii = campoNumerico(KeyAscii)
End Sub

Private Sub txtReferenciaAdd_GotFocus()
    campoSelecionadoComCaracter txtReferenciaAdd
End Sub

Private Sub txtLoja_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If txtLoja.Text <> "" Then
            txtLoja.Locked = True
        End If
        txtFornecedor.SetFocus
    Else
        txtLoja.SetFocus
    End If
End Sub

Private Sub txtSerie_GotFocus()
    campoSelecionadoComCaracter txtSerie
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
 txtSerie.Text = UCase$(txtSerie.Text)
 End If
  KeyAscii = campoNormal(KeyAscii)
  carregaChekin KeyAscii, txtNotaFiscal
 
End Sub

Public Function verificarChekinAberto() As Boolean
    Dim adoChekinSituacao As New ADODB.Recordset
    
    sql = "Select COUNT(*) situacao from chekinmercadoria, produto where chm_loja = '" & txtLoja & _
    "' and chm_fornecedor = '" & codigoFornecedor & "' and chm_notafiscal = '" & txtNotaFiscal & _
    "' and chm_serie = '" & txtSerie & "' and pr_referencia = chm_referencia"
    adoChekinSituacao.CursorLocation = adUseClient
    adoChekinSituacao.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If adoChekinSituacao("situacao") = 0 Then
        verificarChekinAberto = True
    Else
        adoChekinSituacao.Close
        
        sql = "select count(*) situacao from chekinmercadoria where " & _
        "chm_loja = '" & txtLoja & "' and chm_fornecedor = '" & codigoFornecedor & _
        "' and chm_notafiscal = '" & txtNotaFiscal & "' " & _
        "and chm_serie = '" & txtSerie & "' and chm_situacao <> 'E'"
               
        adoChekinSituacao.CursorLocation = adUseClient
        adoChekinSituacao.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
        If adoChekinSituacao("situacao") = 0 Then
            verificarChekinAberto = False
        Else
            verificarChekinAberto = True
        End If
    End If
    
    adoChekinSituacao.Close
End Function


Public Sub carregaChekin(KeyAscii As Integer, numeroNota As String)

Dim sql As String

If KeyAscii = 13 Then
    If verificarCamposObrigCapa Then
    
        lblNota.Visible = False
    
        sql = "select * from capanfcompra where cc_notafiscal=" & txtNotaFiscal & " and cc_fornecedor in ( select fo_codigoFornecedor from fornecedor where fo_cgc = '" & Trim(txtFornecedor.Text) & "') and cc_serie='" & txtSerie & "'"
        rsFornecedor.CursorLocation = adUseClient
        rsFornecedor.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
        If rsFornecedor.EOF Then
            MsgBox "Nota não Encontrada!", vbInformation, "Atenção!"
            lblNota.Caption = "Nota Não Encontrada!"
            lblNota.Visible = True
        Else
            txtLoja.Text = rsFornecedor("cc_loja")
            If rsFornecedor("cc_fornecedor") <> codigoFornecedor Then
            
            
            atualizafornecedor codigoFornecedor, rsFornecedor("cc_fornecedor"), rsFornecedor("cc_notafiscal"), rsFornecedor("cc_serie")
            
            End If
        
        End If
    
    
    
        If verificarChekinAberto Then
            ativaCheckin True
        Else
            If rsFornecedor("cc_situacao") = "D" Then
                If MsgBox("Já foi realizada a contagem na Nota Fiscal " & numeroNota & "Deseja Reabrir o Checkin?", vbYesNo + vbQuestion + vbDefaultButton2, "Chekin Concluindo") = vbYes Then
                    reabreChekin
                End If
            Else
                MsgBox "Já foi realizada a contagem na Nota Fiscal " & numeroNota, vbInformation, "Atenção!"
                LimpaTela
            End If
    
        End If
    End If




rsFornecedor.Close
End If




End Sub

Public Function itemSemReferencia() As Boolean
    Dim adoItemSemReferencia As New ADODB.Recordset
    sql = "select count(*) itemSemReferencia from chekinmercadoria where chm_loja = '" & txtLoja & _
    "' and chm_fornecedor = '" & codigoFornecedor & "' and chm_notaFiscal = '" & txtNotaFiscal & _
    "' and chm_serie = '" & txtSerie & "' and chm_referencia = ''"
    adoItemSemReferencia.CursorLocation = adUseClient
    adoItemSemReferencia.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If adoItemSemReferencia("itemSemReferencia") = 0 Then
        itemSemReferencia = False
    Else
        itemSemReferencia = True
    End If
    
    adoItemSemReferencia.Close
End Function

Public Function contagemRegressiva()
    If tempoRestante <> "00:00:00" Then
        tempoRestante = DateAdd("s", -1, tempoRestante)
    Else
        txtCodigobarras.SetFocus
    End If
End Function

Public Function verificaChekinOK() As Boolean
    Screen.MousePointer = 11
    Dim chekinOK As New ADODB.Recordset
    
    sql = "select count(*) quantChekinOK from chekinmercadoria where chm_loja = '" & txtLoja & _
    "' and chm_fornecedor = '" & codigoFornecedor & "' and chm_notaFiscal = '" & txtNotaFiscal & _
    "' and chm_serie = '" & txtSerie & "' and chm_chekinOK <> 'S'"
    chekinOK.CursorLocation = adUseClient
    chekinOK.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If chekinOK("quantChekinOK") = 0 Then
        verificaChekinOK = True
    Else
        verificaChekinOK = False
    End If
    
    chekinOK.Close
    Screen.MousePointer = 0
End Function

Public Function CarregaFornecedor()

cmbFornecedor.Clear
sql = "select FO_CodigoFornecedor, FO_RazaoSocial from  fornecedor"
rsFornecedor.CursorLocation = adUseClient
rsFornecedor.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
If Not rsFornecedor.EOF Then
Do While Not rsFornecedor.EOF
cmbFornecedor.AddItem Format(rsFornecedor("FO_CodigoFornecedor"), "0000") & " - " & rsFornecedor("FO_RazaoSocial")
        If rsFornecedor("FO_CodigoFornecedor") = codigoFornecedor Then
        cmbFornecedor.ListIndex = cmbFornecedor.ListCount - 1
        
        End If

rsFornecedor.MoveNext
Loop
rsFornecedor.Close
End If


End Function


Private Sub Montagrid(tipogrid As Integer)

    ' Limpa o Grid
    
    grdProduto.Rows = 1
    grdProduto.AddItem ""
    grdProduto.RemoveItem (1)
    
    If tipogrid = 1 Then
    
        ' Adiciona/ Remove Coluna
        grdProdutoTitulo.Cols = 7
        grdProduto.Cols = 7
    
        grdProdutoTitulo.TextMatrix(0, 0) = "Referência"
        grdProdutoTitulo.ColWidth(0) = 1395
        grdProduto.ColWidth(0) = 1395
    
        grdProdutoTitulo.TextMatrix(0, 1) = "Cód. de Barras"
        grdProdutoTitulo.ColWidth(1) = 2070
        grdProduto.ColWidth(1) = 2070
    
        grdProdutoTitulo.TextMatrix(0, 2) = "Descrição"
        grdProdutoTitulo.ColWidth(2) = 3000
        grdProduto.ColWidth(2) = 3000
    
        grdProdutoTitulo.TextMatrix(0, 3) = "Qtd. Checkin"
        grdProdutoTitulo.ColWidth(3) = 1300
        grdProduto.ColWidth(3) = 1300
    
        grdProdutoTitulo.TextMatrix(0, 4) = "Qtd. NF"
        grdProdutoTitulo.ColWidth(4) = 1300
        grdProduto.ColWidth(4) = 1300
    
        grdProdutoTitulo.TextMatrix(0, 5) = "Cód.Barras NF"
        grdProdutoTitulo.ColWidth(5) = 2070
        grdProduto.ColWidth(5) = 2070
    
        grdProdutoTitulo.TextMatrix(0, 6) = "Descrição NF"
        grdProdutoTitulo.ColWidth(6) = 3000
        grdProduto.ColWidth(6) = 3000
    Else
    
        grdProdutoTitulo.Cols = 5
        grdProduto.Cols = 5
    
        grdProdutoTitulo.TextMatrix(0, 0) = "Referência"
        grdProdutoTitulo.ColWidth(0) = 1395
        grdProduto.ColWidth(0) = 1395
    
        grdProdutoTitulo.TextMatrix(0, 1) = "Cód. de Barras"
        grdProdutoTitulo.ColWidth(1) = 2070
        grdProduto.ColWidth(1) = 2070
    
        grdProdutoTitulo.TextMatrix(0, 2) = "Qtd. Checkin"
        grdProdutoTitulo.ColWidth(2) = 1605
        grdProduto.ColWidth(2) = 1605
    
        grdProdutoTitulo.TextMatrix(0, 3) = "Qtd. NF"
        grdProdutoTitulo.ColWidth(3) = 1185
        grdProduto.ColWidth(3) = 1185
        
        grdProdutoTitulo.TextMatrix(0, 4) = "Descrição"
        grdProdutoTitulo.ColWidth(4) = 2000
        grdProduto.ColWidth(4) = 2000
    
    End If

End Sub

Private Sub atualizafornecedor(fornecedorNovo As String, fornecedorAntigo As String, nota As String, serie As String)

    Dim sql As String
    ' Verifica se já há uma nota com esta especificação
    sql = "delete from capanfcompra " & vbNewLine & _
          "where cc_notafiscal = " & nota & " " & vbNewLine & _
          "and cc_fornecedor = " & fornecedorNovo & " " & vbNewLine & _
          "and cc_serie = '" & serie & "'"
    ADO_Cn_CDLocal.Execute (sql)
    
    sql = "delete from itemnfcompra " & vbNewLine & _
          "where ci_notafiscal = " & nota & " " & vbNewLine & _
          "and ci_fornecedor = " & fornecedorNovo & " " & vbNewLine & _
          "and ci_serie = '" & serie & "'"
    ADO_Cn_CDLocal.Execute (sql)
    
    ' Atualiza Capa
    sql = "update capanfcompra" & vbNewLine & _
          " set cc_fornecedor = " & fornecedorNovo & " " & vbNewLine & _
          " where cc_notafiscal = '" & nota & "' " & vbNewLine & _
          " and cc_loja = '" & txtLoja.Text & "' " & vbNewLine & _
          " and cc_serie = '" & serie & "'"
          
    ADO_Cn_CDLocal.Execute (sql)
    
    ' Atualiza Itens
    sql = "update itemnfcompra " & vbNewLine & _
          "set ci_fornecedor = " & fornecedorNovo & " " & vbNewLine & _
          "where ci_notafiscal = '" & nota & "' " & vbNewLine & _
          " and ci_loja = '" & txtLoja.Text & "' " & vbNewLine & _
          "and ci_serie = '" & serie & "'"
    ADO_Cn_CDLocal.Execute (sql)
    
    ' Atualiza Vencimentos
    sql = "update vencimentosFornecedor " & vbNewLine & _
          "set vf_fornecedor = " & fornecedorNovo & " " & vbNewLine & _
          "where vf_notafiscal = '" & nota & "' " & vbNewLine & _
          "and vf_serie = '" & serie & "'"
    ADO_Cn_CDLocal.Execute (sql)

End Sub

Private Sub ConsultaXML(nota As String, serie As String, fornecedor As String)

    Dim sql As String
    Dim rsNota As New ADODB.Recordset
    Dim rsFornecedor As New ADODB.Recordset
    
    'Limpa Grid Fornecedores
    grdXml.Rows = 1
    grdXml.AddItem ""
    grdXml.RemoveItem (1)
    
    sql = " select ci_item, ci_referencia, ci_DescricaoFornecedor,ci_produtoFornecedor,ci_fornecedor ,ci_codigoBarra from itemnfcompra, capanfcompra where ci_notafiscal = '" & nota & "' and ci_serie ='" & serie & "'" _
    & " and ci_notafiscal = cc_notafiscal and ci_serie = cc_serie and ci_fornecedor = cc_fornecedor and cc_situacao = 'D'"
    
    rsNota.CursorLocation = adUseClient
    rsNota.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not rsNota.EOF Then
        lblNumNota.Caption = nota
        lblSerieNota.Caption = serie
        sql = "select fo_codigoFornecedor, fo_nomeFantasia, fo_cgc from fornecedor where fo_codigoFornecedor = " & rsNota("ci_fornecedor")
        rsFornecedor.CursorLocation = adUseClient
        rsFornecedor.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        If Not rsFornecedor.EOF Then
            lblXmlRazao.Caption = rsFornecedor("fo_codigoFornecedor") & " - " & rsFornecedor("fo_nomeFantasia")
            txtCNPJ.Text = rsFornecedor("fo_cgc")
            codFornecedorXML = rsFornecedor("fo_codigoFornecedor")
            txtCNPJ.SetFocus
        End If
        rsFornecedor.Close
        Do While Not rsNota.EOF
            grdXml.AddItem rsNota("ci_item") & Chr(9) _
            & rsNota("ci_referencia") & Chr(9) _
            & rsNota("ci_descricaoFornecedor") & Chr(9) _
            & rsNota("ci_produtoFornecedor") & Chr(9) _
            & rsNota("ci_codigoBarra") & Chr(9) _
            & rsNota("ci_fornecedor")
            
            rsNota.MoveNext
        Loop
            grdXml.row = 1
    End If
    rsNota.Close

End Sub

Private Sub reabreChekin()
    
    Dim sql As String
    Dim rsNota As New ADODB.Recordset
    
    
    sql = "select cc_situacao from capanfcompra where cc_notafiscal = " & txtNotaFiscal.Text & " and cc_serie = '" & txtSerie.Text & "'"
    
    rsNota.CursorLocation = adUseClient
    rsNota.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not rsNota.EOF Then
        If rsNota("cc_situacao") = "D" Then
            sql = "update ChekinMercadoria set chm_situacao = 'A' where chm_fornecedor = " & codigoFornecedor & " and chm_notaFiscal = " & txtNotaFiscal.Text & " and chm_serie = '" & txtSerie.Text & "'"
            ADO_Cn_CDLocal.Execute (sql)
        End If
    End If
    
    rsNota.Close
End Sub

Private Sub deletaItemXML(NotaFiscal, item, fornecedor, serie)
    
    Dim sql As String
    
    If MsgBox("Confirma a exclusão do item?", vbYesNo + vbQuestion + vbDefaultButton2, "Excluir item XML") = vbYes Then

        sql = "delete itemnfcompra where ci_notafiscal = " & NotaFiscal & " and ci_serie = '" & serie & "' and ci_fornecedor = " & fornecedor & " and ci_item = " & item
        ADO_Cn_CDLocal.Execute (sql)
    
    End If
    
End Sub

