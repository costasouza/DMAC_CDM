VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7.ocx"
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7d.ocx"
Begin VB.Form frmEntradaNFCompras 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Entrada de Nota Fiscal Eletrônica Compras"
   ClientHeight    =   7920
   ClientLeft      =   2400
   ClientTop       =   990
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7914.708
   ScaleMode       =   0  'User
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmVencimento 
      Appearance      =   0  'Flat
      BackColor       =   &H00393939&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   4290
      Left            =   3075
      TabIndex        =   84
      Top             =   2715
      Visible         =   0   'False
      Width           =   3135
      Begin VB.TextBox txtValor 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1470
         MaxLength       =   8
         TabIndex        =   71
         ToolTipText     =   "Código do Produto"
         Top             =   825
         Width           =   1200
      End
      Begin VSFlex7DAOCtl.VSFlexGrid grdVencimentos 
         Height          =   2355
         Left            =   150
         TabIndex        =   72
         Top             =   1305
         Width           =   2835
         _cx             =   5001
         _cy             =   4154
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
         BackColor       =   3487029
         ForeColor       =   12632256
         BackColorFixed  =   0
         ForeColorFixed  =   16423203
         BackColorSel    =   16423203
         ForeColorSel    =   8388608
         BackColorBkg    =   3750201
         BackColorAlternate=   3750201
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
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEntradaNFCompras.frx":0000
         ScrollTrack     =   0   'False
         ScrollBars      =   1
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
      Begin CentroDeDistribuicao.chameleonButton cmdOK 
         Height          =   345
         Left            =   870
         TabIndex        =   73
         Top             =   3780
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
         BTYPE           =   14
         TX              =   "E&ncerra Nota"
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
         MICON           =   "frmEntradaNFCompras.frx":006B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSMask.MaskEdBox mskDataVencimento 
         Height          =   315
         Left            =   150
         TabIndex        =   70
         ToolTipText     =   "Data Emissão"
         Top             =   825
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483648
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
         ForeColor       =   &H00C0C0C0&
         Height          =   300
         Left            =   1470
         TabIndex        =   87
         Top             =   540
         Width           =   690
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         ForeColor       =   &H00C0C0C0&
         Height          =   300
         Left            =   150
         TabIndex        =   86
         Top             =   540
         Width           =   690
      End
      Begin VB.Label Label33 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimentos"
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
         TabIndex        =   85
         Top             =   150
         Width           =   2655
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   44
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   14880
      TabIndex        =   80
      Top             =   6815
      Width           =   14880
   End
   Begin VB.Frame frameItens 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   67
      Top             =   5824
      Width           =   14910
      Begin VB.TextBox txtPedido 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   13455
         MaxLength       =   10
         TabIndex        =   81
         Tag             =   "Refêrencia / Código de Barras / Código Produto Fornecedor"
         ToolTipText     =   "Número Pedido"
         Top             =   375
         Width           =   1320
      End
      Begin VB.TextBox txtPercentualDesconto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   8970
         MaxLength       =   6
         TabIndex        =   45
         ToolTipText     =   "Porcentagem Desconto"
         Top             =   370
         Width           =   825
      End
      Begin VB.TextBox txtAliquotaIPI 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   8240
         MaxLength       =   6
         TabIndex        =   44
         ToolTipText     =   "Porcentagem IPI"
         Top             =   370
         Width           =   630
      End
      Begin VB.TextBox txtPrecoUnitario 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   5420
         MaxLength       =   8
         TabIndex        =   41
         ToolTipText     =   "Preço Unitário"
         Top             =   370
         Width           =   1260
      End
      Begin VB.TextBox txtQuantidade 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   4660
         MaxLength       =   4
         TabIndex        =   40
         ToolTipText     =   "Quantidade"
         Top             =   370
         Width           =   660
      End
      Begin VB.TextBox txtReferencia 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   120
         MaxLength       =   20
         TabIndex        =   39
         Tag             =   "Refêrencia / Código de Barras / Código Produto Fornecedor"
         Top             =   370
         Width           =   4440
      End
      Begin VB.TextBox txtCFOP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   6780
         MaxLength       =   7
         TabIndex        =   42
         ToolTipText     =   "CFOP"
         Top             =   370
         Width           =   630
      End
      Begin VB.TextBox txtPorcentagemICMS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   7510
         MaxLength       =   6
         TabIndex        =   43
         ToolTipText     =   "Porcentagem ICMS"
         Top             =   370
         Width           =   630
      End
      Begin CentroDeDistribuicao.chameleonButton cmdLimparCamposItemNFCompra 
         Height          =   315
         Left            =   9895
         TabIndex        =   79
         Top             =   360
         Visible         =   0   'False
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         BTYPE           =   14
         TX              =   "L&impar Campos"
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
         MICON           =   "frmEntradaNFCompras.frx":0087
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número Pedido"
         ForeColor       =   &H8000000F&
         Height          =   195
         Left            =   13455
         TabIndex        =   82
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Desc."
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   8970
         TabIndex        =   78
         Top             =   120
         Width           =   585
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% IPI"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   8240
         TabIndex        =   77
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço Unitário"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   5420
         TabIndex        =   76
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde."
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   4660
         TabIndex        =   75
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Refêrencia / Código de Barras / Código Produto Fornecedor"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   120
         TabIndex        =   74
         ToolTipText     =   "Código Produto"
         Top             =   120
         Width           =   4275
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% ICMS"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   7510
         TabIndex        =   69
         Top             =   120
         Width           =   555
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CFOP"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   6780
         TabIndex        =   68
         Top             =   120
         Width           =   420
      End
   End
   Begin VB.Frame frameCadastraReferencia 
      BackColor       =   &H00393939&
      BorderStyle     =   0  'None
      Height          =   3285
      Left            =   6675
      TabIndex        =   62
      Top             =   2520
      Visible         =   0   'False
      Width           =   8055
      Begin VB.Frame Frame1 
         BackColor       =   &H00393939&
         BorderStyle     =   0  'None
         Height          =   810
         Index           =   0
         Left            =   165
         TabIndex        =   64
         Top             =   2460
         Width           =   7725
         Begin VB.TextBox txtAddCodigoProduto 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   105
            MaxLength       =   20
            TabIndex        =   57
            ToolTipText     =   "Código do Produto"
            Top             =   270
            Width           =   2265
         End
         Begin VB.TextBox txtAddReferencia 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   2460
            MaxLength       =   7
            TabIndex        =   58
            ToolTipText     =   "Refêrencia"
            Top             =   270
            Width           =   2265
         End
         Begin CentroDeDistribuicao.chameleonButton cmdAddReferencia 
            Height          =   345
            Left            =   4875
            TabIndex        =   59
            Top             =   255
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   609
            BTYPE           =   14
            TX              =   "&Gravar Código"
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
            MICON           =   "frmEntradaNFCompras.frx":00A3
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
            Height          =   345
            Left            =   6315
            TabIndex        =   60
            Top             =   255
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   609
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
            MICON           =   "frmEntradaNFCompras.frx":00BF
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "Refêrencia"
            ForeColor       =   &H00C0C0C0&
            Height          =   300
            Left            =   2460
            TabIndex        =   66
            Top             =   -15
            Width           =   855
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Código do Produto"
            ForeColor       =   &H00C0C0C0&
            Height          =   300
            Left            =   120
            TabIndex        =   65
            Top             =   -30
            Width           =   1605
         End
      End
      Begin VSFlex7DAOCtl.VSFlexGrid grdReferenciaSemBarras 
         Height          =   1850
         Left            =   150
         TabIndex        =   63
         Top             =   480
         Width           =   7695
         _cx             =   13573
         _cy             =   3263
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
         BackColorAlternate=   3158064
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
         FormatString    =   $"frmEntradaNFCompras.frx":00DB
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
      Begin VB.Label Label32 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
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
         TabIndex        =   83
         Top             =   150
         Width           =   2655
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grdPrincipal 
      Height          =   3345
      Left            =   120
      TabIndex        =   61
      Top             =   2325
      Width           =   14910
      _cx             =   26300
      _cy             =   5900
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmEntradaNFCompras.frx":0153
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
      ShowComboButton =   0   'False
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   0
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   5263440
      ForeColorFrozen =   4210752
      WallPaperAlignment=   9
   End
   Begin VB.Frame frmEntrada 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2070
      Left            =   120
      TabIndex        =   56
      Top             =   150
      Width           =   14910
      Begin VB.TextBox txtTotalDesconto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   9000
         MaxLength       =   8
         TabIndex        =   88
         ToolTipText     =   "IPI"
         Top             =   1605
         Width           =   1400
      End
      Begin VB.TextBox txtCodigoFornecedor 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   3675
         MaxLength       =   4
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Código Fornecedor"
         Top             =   350
         Width           =   735
      End
      Begin VB.ComboBox cmbTipoEntrada 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   9480
         TabIndex        =   7
         ToolTipText     =   "Tipo de Entrada"
         Top             =   350
         Width           =   1740
      End
      Begin VB.TextBox txtDataEntrada 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   11300
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Data Entrada"
         Top             =   350
         Width           =   1110
      End
      Begin VB.TextBox txtCGC 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1600
         MaxLength       =   15
         TabIndex        =   2
         ToolTipText     =   "CNPJ"
         Top             =   350
         Width           =   2000
      End
      Begin VB.TextBox txtNotaFiscal 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   7520
         MaxLength       =   8
         TabIndex        =   5
         ToolTipText     =   "Nota Fiscal"
         Top             =   350
         Width           =   1400
      End
      Begin VB.TextBox txtSerie 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   9000
         MaxLength       =   2
         TabIndex        =   6
         ToolTipText     =   "Série"
         Top             =   350
         Width           =   400
      End
      Begin VB.TextBox txtNomeFantasia 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   4500
         MaxLength       =   60
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Fornecedor"
         Top             =   350
         Width           =   2940
      End
      Begin VB.ComboBox cmbLocal 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Loja"
         Top             =   350
         Width           =   1400
      End
      Begin VB.TextBox txtTotalCalculado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   13425
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Total Calculado"
         Top             =   1610
         Width           =   1400
      End
      Begin VB.TextBox txtBateNota 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   11960
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Bate Nota"
         Top             =   1610
         Width           =   1400
      End
      Begin VB.TextBox txtValorTotalNota 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   10480
         MaxLength       =   8
         TabIndex        =   21
         ToolTipText     =   "Total Nota Fiscal"
         Top             =   1610
         Width           =   1400
      End
      Begin VB.TextBox txtValorICMSSubsTrib 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   6040
         MaxLength       =   8
         TabIndex        =   15
         ToolTipText     =   "ICMS Subst."
         Top             =   980
         Width           =   1400
      End
      Begin VB.TextBox txtValorICMS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   4560
         MaxLength       =   8
         TabIndex        =   14
         ToolTipText     =   "ICMS"
         Top             =   980
         Width           =   1400
      End
      Begin VB.TextBox txtBaseICMS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1600
         MaxLength       =   8
         TabIndex        =   12
         ToolTipText     =   "Base ICMS"
         Top             =   980
         Width           =   1400
      End
      Begin VB.TextBox txtFrete 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   7520
         MaxLength       =   8
         TabIndex        =   16
         ToolTipText     =   "Frete"
         Top             =   980
         Width           =   1400
      End
      Begin VB.TextBox txtValorIPI 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   9000
         MaxLength       =   8
         TabIndex        =   17
         ToolTipText     =   "IPI"
         Top             =   980
         Width           =   1400
      End
      Begin VB.TextBox txtDespesas 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   11960
         MaxLength       =   8
         TabIndex        =   19
         ToolTipText     =   "Despesas"
         Top             =   980
         Width           =   1400
      End
      Begin VB.TextBox txtOutros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   13425
         MaxLength       =   8
         TabIndex        =   20
         ToolTipText     =   "Outros"
         Top             =   980
         Width           =   1400
      End
      Begin VB.TextBox txtEmbalagem 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   10480
         MaxLength       =   8
         TabIndex        =   18
         ToolTipText     =   "Embalagem"
         Top             =   980
         Width           =   1400
      End
      Begin VB.TextBox txtValorMercadorias 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   120
         MaxLength       =   8
         TabIndex        =   11
         ToolTipText     =   "Valor Mercadorias"
         Top             =   980
         Width           =   1400
      End
      Begin VB.TextBox txtaliquotaicms 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   3080
         MaxLength       =   8
         TabIndex        =   13
         ToolTipText     =   "Aliq. ICMS"
         Top             =   980
         Width           =   1400
      End
      Begin MSMask.MaskEdBox mskDataEmissao 
         Height          =   315
         Left            =   12490
         TabIndex        =   9
         ToolTipText     =   "Data Emissão"
         Top             =   350
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483648
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDataRecebimento 
         Height          =   315
         Left            =   13695
         TabIndex        =   10
         ToolTipText     =   "Data Recebimento"
         Top             =   345
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483648
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Desconto"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   8985
         TabIndex        =   89
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Entrada"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   9480
         TabIndex        =   50
         Top             =   120
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Entrada"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   11300
         TabIndex        =   38
         Top             =   120
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Emissão"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   12490
         TabIndex        =   37
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   1600
         TabIndex        =   55
         Top             =   120
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nota Fiscal"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   7520
         TabIndex        =   52
         Top             =   120
         Width           =   795
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Série"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   9000
         TabIndex        =   51
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo / Fornecedor"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   3680
         TabIndex        =   53
         Top             =   120
         Width           =   1470
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bate Nota"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   11960
         TabIndex        =   28
         Top             =   1380
         Width           =   720
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Calculado"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   13425
         TabIndex        =   27
         Top             =   1380
         Width           =   1110
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Nota Fiscal"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   10480
         TabIndex        =   26
         Top             =   1380
         Width           =   1200
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ICMS Subst."
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   6040
         TabIndex        =   33
         Top             =   750
         Width           =   885
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ICMS"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   4560
         TabIndex        =   32
         Top             =   750
         Width           =   270
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base ICMS"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   1600
         TabIndex        =   30
         Top             =   750
         Width           =   795
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frete"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   7520
         TabIndex        =   34
         Top             =   750
         Width           =   360
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IPI"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   9000
         TabIndex        =   35
         Top             =   750
         Width           =   195
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Despesas"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   11960
         TabIndex        =   24
         Top             =   750
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Outros"
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Left            =   13425
         TabIndex        =   0
         Top             =   750
         Width           =   465
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Embalagem"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   10480
         TabIndex        =   25
         Top             =   750
         Width           =   825
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Mercadorias"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   750
         Width           =   1275
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Receb."
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   13695
         TabIndex        =   36
         Top             =   120
         Width           =   915
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aliq. ICMS"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   3080
         TabIndex        =   31
         Top             =   750
         Width           =   735
      End
   End
   Begin CentroDeDistribuicao.chameleonButton cmdLimpar 
      Height          =   510
      Left            =   9300
      TabIndex        =   47
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "&Limpar"
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
      MICON           =   "frmEntradaNFCompras.frx":028E
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
      TabIndex        =   49
      Top             =   6950
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
      MICON           =   "frmEntradaNFCompras.frx":02AA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdEncerraNota 
      Height          =   510
      Left            =   10740
      TabIndex        =   46
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "&Encerra Nota"
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
      MICON           =   "frmEntradaNFCompras.frx":02C6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdExcluiNota 
      Height          =   510
      Left            =   12180
      TabIndex        =   48
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "E&xclui Nota"
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
      MICON           =   "frmEntradaNFCompras.frx":02E2
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
Attribute VB_Name = "frmEntradaNFCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Editado por: Felip3FL (felip3.fl@gmail.com)
'Última atualização: 25/06/2013
'Versão: 1.5.84
'Última atualização: 22/09/2015


Option Explicit

Dim porcentagemFrete, valorFrete As Double
Dim Conta, IPI, ICMS, desconto As Double
Dim statusChekin As Boolean
Dim rdoProduto As New ADODB.Recordset
Dim rdoChekin As New ADODB.Recordset

Public Sub montaComboLoja(comboLojas As ComboBox)
'On Error GoTo TrataErro
    Dim ado_loja As New ADODB.Recordset
    Dim ado_loja2 As New ADODB.Recordset
    Dim lojasWhere As String
    Dim i As Byte

    With ado_loja
        sql = "select cts_loja as lojas From ControleSistema"
        .CursorLocation = adUseClient
        .Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
            sql = "select LO_Loja AS LOJAS from loja where LO_Situacao = 'A' AND lo_loja not in ('CONSO')"
            ado_loja2.CursorLocation = adUseClient
            ado_loja2.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
            
            lojasWhere = "("
            
            Do While Not ado_loja2.EOF
            
                If Trim(ado_loja("lojas")) = Trim(ado_loja2("lojas")) Then
                    i = ado_loja2.AbsolutePosition
                End If
            
                comboLojas.AddItem Trim(ado_loja2("lojas"))
                lojasWhere = lojasWhere & "'" & Trim(ado_loja2("lojas")) & "'" & ","
                ado_loja2.MoveNext
                
            Loop
            
            lojasWhere = left(lojasWhere, (Len(lojasWhere) - 1)) & ")"
            ado_loja2.Close
            
        .Close
    End With

    comboLojas.ListIndex = i - 1

End Sub

Private Sub atualizaLojaCapaNFCompra(nf As String, serie As String, cnpjFornecedor As String, lojaNova As String)

    sql = "update capanfcompra " & vbNewLine & _
          "set cc_loja = '" & lojaNova & "' " & vbNewLine & _
          "where cc_notafiscal = '" & nf & "' " & vbNewLine & _
          "and cc_serie = '" & serie & "' " & vbNewLine & _
          "and cc_fornecedor in (select fo_codigofornecedor from fornecedor where fo_cgc = '" & cnpjFornecedor & "')" & vbNewLine
    
    sql = sql & "update itemnfcompra " & vbNewLine & _
          "set ci_loja = '" & lojaNova & "' " & vbNewLine & _
          "where ci_notafiscal = '" & nf & "' " & vbNewLine & _
          "and ci_serie = '" & serie & "' " & vbNewLine & _
          "and ci_fornecedor in (select fo_codigofornecedor from fornecedor where fo_cgc = '" & cnpjFornecedor & "')" & vbNewLine
    
    ADO_Cn_CDLocal.Execute sql
    
End Sub

Private Sub Form_Load()

    statusChekin = False

    txtDataEntrada.Text = Format(CStr(Date), "dd/mm/yyyy")
   ' MontaComboEntrada
    cmbTipoEntrada.Text = ""
    montaComboLoja cmbLocal
    'cmbLocal.ListIndex = 0
    
    carregarPosicaoTamanhoTela Me
    'JanelaTOP frmControleCD
    frmControleCD.Enabled = False
    
    carregarPosicaoFrame frameCadastraReferencia
    carregarPosicaoFrame frmVencimento
    'frmVencimento.top = (tamanhoTelaY / 2) - (frmVencimento.Height / 2)
    'frmVencimento.left = (tamanhoTelaX / 2) - (frmVencimento.Width / 2)
    
    'txtCodigoFornecedor = "45"
    'txtNotaFiscal = "446426"
    'txtSerie = "NE"
    
    'grdVencimentos.AddItem "teste"
    
End Sub

Private Sub cmdAddReferencia_Click()
    abilitaFrameCadastraReferencia False
    
    Dim sql As String
    
    sql = "update itemNFcompra set ci_referencia = '" & txtReferencia & "' where ci_notafiscal = '" & txtNotaFiscal _
    & "" & "' and ci_serie = '" & txtSerie & "' and ci_fornecedor in(select fo_codigoFornecedor from fornecedor where fo_cgc ='" & txtCGC.Text _
    & "') and ci_loja = '" & cmbLocal & "' and ci_codigoProduto = '" & txtAddCodigoProduto & "'"
    
    ADO_Cn_CDLocal.Execute (sql)
    
End Sub

Private Sub cmdExcluiNota_Click()
    If mensagemExluir("Nota Fiscal Eletrônica") Then
        sql = "SP_deletaCapaItemNFCompra '" & txtNotaFiscal & "','" & txtSerie & _
        "','" & txtCodigoFornecedor & "','" & cmbLocal & "'"
        ADO_Cn_CDLocal.Execute sql
        limparTodosCampos True
    End If
End Sub

Private Sub cmdRetorna_Click()
    Unload Me
End Sub


Private Sub cmdSairCadastroReferencia_Click()
    abilitaFrameCadastraReferencia False
End Sub

Private Sub cmdSairVendimento_Click()
    ativaVencimento False
End Sub

Sub MontaComboEntrada()
    
    Dim adoCombos As New ADODB.Recordset
    Dim sql As String
    
    sql = "select tpc_codigo, tpc_descricao from tipopedidocompra"
    adoCombos.CursorLocation = adUseClient
    adoCombos.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    Do While Not adoCombos.EOF
        cmbTipoEntrada.AddItem Format(adoCombos("tpc_codigo"), "00") & " - " _
        & Trim(adoCombos("tpc_descricao"))
        adoCombos.MoveNext
    Loop
    cmbTipoEntrada.ListIndex = 0
    
    adoCombos.Close
    
End Sub


'===========================================================================================================
'========================================= CAMPOS OBRIGATORIOS DO FORMULARIO ===============================
'===========================================================================================================

Private Function verificarCampoObrigItem() As Boolean

    verificarCampoObrigItem = False
    
    If txtReferencia.Text = "" Then
        txtReferencia.SetFocus
        mensagemCampoObrigatorio "Refêrencia"
        
    ElseIf txtPrecoUnitario.Text = "" Then
        txtPrecoUnitario.SetFocus
        mensagemCampoObrigatorio "Preço Unitário"
        
    Else
        verificarCampoObrigItem = True
    End If
    
End Function

Private Function verificaCamposObrigatorioCapa() As Boolean

    verificaCamposObrigatorioCapa = False
        
    If txtCGC = "" Then
        txtCGC.SetFocus
        mensagemCampoObrigatorio "CGC"
        
    ElseIf txtCodigoFornecedor.Text = "" Then
        txtCodigoFornecedor.SetFocus
        mensagemCampoObrigatorio "Código Fornecedor"
        
    ElseIf txtNotaFiscal.Text = "" Then
        txtNotaFiscal.SetFocus
        mensagemCampoObrigatorio "Nota Fiscal"
        
    ElseIf txtSerie = "" Then
        txtSerie.SetFocus
        mensagemCampoObrigatorio "Serie"
        
    'ElseIf txtDataEntrada = "" Then
        'txtDataEntrada.SetFocus
        'mensagemCampoObrigatorio "Data Entrada"
        
    ElseIf mskDataEmissao.ClipText = "" Or Len(mskDataEmissao.ClipText) < 8 Then
        'mskDataEmissao.Text = Date
        'mskDataEmissao.SetFocus
        'mensagemCampoObrigatorio "Data emissão"
        
    ElseIf mskDataRecebimento.ClipText = "" Or Len(mskDataRecebimento.ClipText) < 8 Then
        mskDataRecebimento = "__/___/____"
        mskDataRecebimento.SetFocus
        mensagemCampoObrigatorio "Data Recebimento"
        
    ElseIf txtValorMercadorias = "" Then
        txtValorMercadorias.SetFocus
        mensagemCampoObrigatorio "Valor Mercadoria"
        
    ElseIf txtValorTotalNota = "" Then
        txtValorTotalNota.SetFocus
        mensagemCampoObrigatorio "Total Nota Fiscal"

    Else
        verificaCamposObrigatorioCapa = True
    End If
    
    
    If txtBaseICMS = "" Then
    autoPreencherZero txtBaseICMS: End If
    If txtTotalDesconto = "" Then
    autoPreencherZero txtTotalDesconto: End If
    If txtaliquotaicms = "" Then
    autoPreencherZero txtaliquotaicms: End If
    If txtValorICMS = "" Then
    autoPreencherZero txtValorICMS: End If
    If txtValorICMSSubsTrib = "" Then
    autoPreencherZero txtValorICMSSubsTrib: End If
    If txtFrete = "" Then
    autoPreencherZero txtFrete: End If
    If txtValorIPI = "" Then
    autoPreencherZero txtValorIPI: End If
    If txtEmbalagem = "" Then
    autoPreencherZero txtEmbalagem: End If
    If txtDespesas = "" Then
    autoPreencherZero txtDespesas: End If
    If txtOutros = "" Then
    autoPreencherZero txtOutros: End If

End Function

Private Sub bloquearCamposItemCompra(bloquear As Boolean)

    If bloquear Then
        limparTodosCampos False
        txtQuantidade.Enabled = False
        txtAliquotaIPI.Enabled = False
        txtPercentualDesconto.Enabled = False
        txtCFOP.Enabled = False
        txtPorcentagemICMS.Enabled = False
    Else
        txtQuantidade.Enabled = True
        txtAliquotaIPI.Enabled = True
        txtPercentualDesconto.Enabled = True
        txtCFOP.Enabled = True
        txtPorcentagemICMS.Enabled = True
    End If
    
End Sub


'===========================================================================================================
'===================================   CARREGAR DADOS DO BANCO DE DADOS   ==================================
'===========================================================================================================


Private Function carregarDescricaoProduto(referencia As String) As String
    Dim rdoDescricao As New ADODB.Recordset
    Dim sql As String
    
    sql = "select pr_descricao from produto where pr_referencia = '" _
    & referencia & "'"
    
    rdoDescricao.CursorLocation = adUseClient
    rdoDescricao.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If rdoDescricao.BOF And rdoDescricao.EOF Then
        carregarDescricaoProduto = " (Não foi possível encontra o produto) "
    Else
        carregarDescricaoProduto = rdoDescricao("pr_descricao")
    End If
    
    rdoDescricao.Close
End Function



Private Function carregarDadosGrid()
    Dim itemSemPedido As Integer
    
    If txtPedido.Text <> "" Then
    
        sql = "select pc_autorizacaoCompras from capapedido where pc_numeroPedido = " & txtPedido.Text
        rdoProduto.CursorLocation = adUseClient
        rdoProduto.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
        If rdoProduto.EOF Then
            MsgBox "Pedido Não Encontrado.", vbCritical, "Pedido Não Encontrado"
            rdoProduto.Close
            Exit Function
        Else
            If rdoProduto("pc_autorizacaoCompras") <> "A" Then
                MsgBox "Pedido não Autorizado.", vbCritical, "Pedido Não Autorizado"
                Exit Function
            End If
        End If
        rdoProduto.Close
    
    End If
    
    sql = "exec SP_AtualizaNumeroPedido '" & txtNotaFiscal & "','" & txtSerie & _
    "','" & txtCodigoFornecedor & "','" & cmbLocal & "','" & txtPedido & "'"
    ADO_Cn_CDLocal.Execute sql
    
    sql = " select TPC_Descricao from capaPedido, tipoPedidoCompra where pc_tipoPedido = tpc_codigo and pc_NumeroPedido = '" & txtPedido.Text & "'"
    
    rdoProduto.CursorLocation = adUseClient
    rdoProduto.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
     
    If Not rdoProduto.EOF Then
        If cmbTipoEntrada.Text = "" Or cmbTipoEntrada.Text = Trim(rdoProduto("TPC_Descricao")) Then
            cmbTipoEntrada.Text = Trim(rdoProduto("TPC_Descricao"))
        Else
            MsgBox "Não é possível inserir Tipos de Entrada diferentes. Itens Anteriores:" & cmbTipoEntrada.Text & " Item Atual: " & Trim(rdoProduto("TPC_Descricao")), vbInformation, "Aviso"
            Exit Function
        End If
    End If
    rdoProduto.Close
    
    limpaGrid grdPrincipal
    txtPedido.Text = ""
    
    sql = "select CI_Referencia, CI_codigoProduto, CI_quantidade, CI_PrecoUnitario, " _
    & "CI_CFOP, CI_aliqICMS, CI_aliquotaIPI, CI_PercentualDesconto, CI_NossoPedido " _
    & "from ItemNFCompra where ci_notaFiscal = '" & txtNotaFiscal _
    & "' and ci_serie = '" & txtSerie & "' and ci_fornecedor in (select fo_codigoFornecedor from fornecedor where fo_cgc = '" & txtCGC.Text _
    & "')  order by CI_ITEM"
    
    rdoProduto.CursorLocation = adUseClient
    rdoProduto.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
            
    Do While Not rdoProduto.EOF
    
        grdPrincipal.AddItem rdoProduto("CI_codigoProduto") & Chr(9) & rdoProduto("CI_Referencia") _
        & Chr(9) & carregarDescricaoProduto(rdoProduto("CI_Referencia")) & Chr(9) _
        & rdoProduto("CI_quantidade") & Chr(9) & Format(rdoProduto("CI_PrecoUnitario"), "##,#0.00") _
        & Chr(9) & rdoProduto("CI_CFOP") & Chr(9) & rdoProduto("CI_aliqICMS") & Chr(9) _
        & rdoProduto("CI_aliquotaIPI") & Chr(9) & rdoProduto("CI_PercentualDesconto") _
        & Chr(9) & rdoProduto("CI_NossoPedido")
        
        If rdoProduto("CI_NossoPedido") = "0" Then
            itemSemPedido = itemSemPedido + 1
        End If
        
        rdoProduto.MoveNext
    Loop
    
    txtBateNota = calculaDiferenca(formatParaCalculo(txtValorTotalNota), formatParaCalculo(txtTotalCalculado), formatParaCalculo(txtTotalDesconto))
    formataCampoDinheiro txtTotalCalculado
    formataCampoDinheiro txtBateNota
    
    If itemSemPedido > 0 Or mskDataRecebimento.Text = "01/01/1900" Then
        cmdEncerraNota.Enabled = False
    Else
        cmdEncerraNota.Enabled = True
    End If

    rdoProduto.Close
    
End Function

Private Function gravaDadosItemNFCompra() As Boolean
    
    If verificarCampoObrigItem Then
    
        Dim rdoVerificarCodigo As New ADODB.Recordset
        Dim codigo As String
    
        'SQL = "select * from ItemNFCompra where ci_notaFiscal = '" & txtNotaFiscal _
        & "' and ci_serie = '" & txtSerie & "' and ci_fornecedor = '" & txtCodigoFornecedor _
        & "' and ci_loja = '" & Trim(cmbLocal) & "' and ci_codigoProduto = '" & txtReferencia & "'"
        
        'rdoProduto.CursorLocation = adUseClient
        'rdoProduto.Open SQL, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
        'If rdoProduto.BOF And rdoProduto.EOF Then
        
            sql = "select pr_referencia, pr_codigoprodutonofornecedor " _
            & "from produto where pr_referencia='" & txtReferencia.Text & "' or " _
            & "pr_codigoprodutonofornecedor='" & txtReferencia.Text & "' Union " _
            & "select prb_referencia,prb_codigobarras from produtobarras " _
            & "where prb_codigobarras='" & txtReferencia.Text & "' and prb_tipocodigo='B'"
            
            rdoVerificarCodigo.CursorLocation = adUseClient
            rdoVerificarCodigo.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
            If rdoVerificarCodigo.BOF And rdoVerificarCodigo.EOF Then
                codigo = ""
            Else
                codigo = rdoVerificarCodigo("pr_referencia")
            End If
            
            txtTotalCalculado.Text = formatParaCalculo(txtTotalCalculado) + calculoNota(porcentagemFrete, formatParaCalculo(txtPrecoUnitario.Text), txtQuantidade, _
            formatParaCalculo(txtAliquotaIPI.Text), formatParaCalculo(txtPorcentagemICMS.Text), formatParaCalculo(txtPercentualDesconto.Text))
            
            sql = "exec SP_GravaItemNFCompra '" & Trim(cmbLocal) & "','" & txtNotaFiscal & "','" & txtSerie _
            & "','" & txtCodigoFornecedor & "','" & codigo & "','" & txtQuantidade & "','" _
            & formatParaGravar(txtPrecoUnitario) _
            & "','" & txtAliquotaIPI & "','" & txtPorcentagemICMS & "','" & txtPercentualDesconto & "','" & txtReferencia.Text & "','" _
            & Format(txtDataEntrada, "mm/dd/yyyy") & "','" & formatParaGravar(valorFrete) & "', '" _
            & formatParaGravar(IPI) & "', '" & formatParaGravar(ICMS) & "', '" & txtCFOP & "'"
            ADO_Cn_CDLocal.Execute sql
        
            gravaDadosItemNFCompra = True
            rdoVerificarCodigo.Close
            
        'Else
                'ADO_Cn_CDLocal.Execute "exec SP_atualizaItemNFCompra '" & txtNotaFiscal & "', '" & txtSerie _
                '& "', '" & txtReferencia & "', '" & txtQuantidade + rdoProduto("CI_quantidade") & "'"
                'gravaDadosItemNFCompra = True
    
        'End If
        
        carregarDadosGrid
        validaChekin
        atualizaDadosCapaNFCompra
        abilitaCamposItemNFCompra True
        limparTodosCampos False
        txtReferencia.SetFocus
    
    End If
    
End Function

Private Function calculoNota(ByVal porcentagemFrete As Double, precoUnitario As Double, _
ByVal quantidade As Double, porcentagemIPI As Double, porcentagemICMS As Double, desconto As Double) As Double

    valorFrete = calculaFrete(porcentagemFrete, precoUnitario, quantidade)
    ICMS = (((precoUnitario * quantidade) + valorFrete) * porcentagemICMS) / 100
    IPI = ((((precoUnitario * quantidade) + _
    ((precoUnitario * quantidade) * porcentagemFrete) / 100) * porcentagemIPI) / 100)
    desconto = (desconto * (precoUnitario * quantidade)) / 100
    Conta = (precoUnitario * quantidade) + IPI + valorFrete - desconto
    
    'txtTotalCalculado.Text = Conta
    calculoNota = Conta

End Function

Private Function atualizaDadosCapaNFCompra()



    sql = "SP_atualizaCapaNFCompra '" & txtNotaFiscal & "','" & txtSerie & "','" & txtCodigoFornecedor _
    & "','" & cmbLocal & "','" & txtTotalCalculado & "', '" & "D" & "'"
    
    If mskDataEmissao.Text = "__/__/____" Or mskDataEmissao.Text = "" Then
        mskDataEmissao.Text = Date
    End If
    
    sql = "exec SP_atualizaCapaNFCompra '" & txtNotaFiscal & "','" & txtSerie & "','" & txtCodigoFornecedor & "','" & Trim(cmbLocal) & "','" & txtTotalCalculado & "','" & "E" & "', '" & Format(mskDataEmissao.Text, "yyyy/mm/dd") & "', '" & Format(mskDataRecebimento.Text, "yyyy/mm/dd") & "', '" & Format(txtDataEntrada.Text, "yyyy/mm/dd") & "'"
    
    ADO_Cn_CDLocal.Execute sql



End Function


'===========================================================================================================
'===========================================================================================================
'===========================================================================================================


Private Sub limparTodosCampos(limpezaCompleta As Boolean)
    
    If limpezaCompleta Then
    
        grdPrincipal.Rows = 1
        txtCGC.Text = ""
        txtCodigoFornecedor.Text = ""
        txtNomeFantasia.Text = ""
        txtNotaFiscal.Text = ""
        txtSerie.Text = ""
        'txtDataEntrada.Text = ""
        mskDataEmissao = "__/__/____"
        mskDataRecebimento = "__/__/____"
        txtValorMercadorias.Text = ""
        txtBaseICMS.Text = ""
        txtTotalDesconto.Text = ""
        txtaliquotaicms.Text = ""
        txtValorICMS.Text = ""
        txtValorICMSSubsTrib.Text = ""
        txtFrete.Text = ""
        txtValorIPI.Text = ""
        txtEmbalagem.Text = ""
        txtDespesas.Text = ""
        txtOutros.Text = ""
        txtValorTotalNota.Text = ""
        txtBateNota.Text = ""
        txtTotalCalculado.Text = ""
        txtPedido.Text = ""
        cmbTipoEntrada.Text = ""
        
        abilitaCapa True
    End If
    
    txtReferencia = "": txtQuantidade = "": txtPrecoUnitario = "": txtCFOP = "": txtPorcentagemICMS = ""
    txtAliquotaIPI = "": txtPercentualDesconto.Text = ""
    
End Sub

Private Sub autoPreencherZero(campo As TextBox)
    If campo.Text = "" Then
        campo.Text = "0,00"
    Else
        formataCampoDinheiro campo
    End If
End Sub

Private Sub formataCampoDinheiro(campoDinheiro As TextBox)
    campoDinheiro.Text = Format(campoDinheiro.Text, "##,#0.00")
End Sub

Private Function formatCampoDinheiro(campoDinheiro As String) As String
    formatCampoDinheiro = Format(campoDinheiro, "##,#0.00")
End Function

Private Sub abilitaCamposItemNFCompra(abilita As Boolean)

    If abilita Then
        txtPrecoUnitario.Enabled = True
        txtAliquotaIPI.Enabled = True
        txtPercentualDesconto.Enabled = True
        txtCFOP.Enabled = True
        txtPorcentagemICMS.Enabled = True
        cmdLimparCamposItemNFCompra.Visible = False
    Else
        txtPrecoUnitario.Enabled = False
        txtAliquotaIPI.Enabled = False
        txtPercentualDesconto.Enabled = False
        txtCFOP.Enabled = False
        txtPorcentagemICMS.Enabled = False
        cmdLimparCamposItemNFCompra.Visible = True
    End If
    
End Sub

Private Sub Form_LostFocus()
    'frmEntradaNFCompras.Show
End Sub






Private Sub grdPrincipal_DblClick()

    'Dim LastRow As Long               ' Ultima linha em que se editou
    'Dim LastCol As Long               ' ultima coluna em que se editou
    
    If grdPrincipal.Rows > 0 Then
        grdPrincipal.col = 1
        If Trim(grdPrincipal) = "" Then
            grdPrincipal.col = 0
            txtAddCodigoProduto.Text = Trim(grdPrincipal.Text)
            abilitaFrameCadastraReferencia True
        End If
    End If
    
    'frmCadastraReferencia.Show
    
    'LastRow = grdPrincipal.Row
    'LastCol = grdPrincipal.Col
    
    'Select Case LastCol
    'Case Else
    '    txtCadastraReferencia.Move grdPrincipal.CellLeft - Screen.TwipsPerPixelX, grdPrincipal.CellTop + 1000 - Screen.TwipsPerPixelY, grdPrincipal.CellWidth + Screen.TwipsPerPixelX * 2, grdPrincipal.CellHeight + Screen.TwipsPerPixelY * 2
    '    txtCadastraReferencia.Text = grdPrincipal.Text
    '    If Len(grdPrincipal.Text) = 0 Then
    '        If LastRow > 1 Then
    '            txtCadastraReferencia.Text = grdPrincipal.TextMatrix(LastRow - 1, LastCol)
    '        End If
    '    End If
    'End Select
    
End Sub


Private Sub grdPrincipal_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If txtSerie <> "NE" Then
            grdPrincipal.col = 0
            If mensagemExluir("o item " & Trim(grdPrincipal)) Then
                sql = "SP_deletaItemNFCompra '" & txtNotaFiscal & _
                "','" & txtSerie & "','" & txtCodigoFornecedor & _
                "','" & cmbLocal & "','" & grdPrincipal & "'"
                ADO_Cn_CDLocal.Execute (sql)
    
                txtTotalCalculado.Text = formatParaCalculo(txtTotalCalculado) - _
                calculoNota(porcentagemFrete, formatParaCalculo(grdPrincipal.TextMatrix(grdPrincipal.row, 4)), _
                grdPrincipal.TextMatrix(grdPrincipal.row, 3), _
                formatParaCalculo(grdPrincipal.TextMatrix(grdPrincipal.row, 7)), _
                formatParaCalculo(grdPrincipal.TextMatrix(grdPrincipal.row, 6)), _
                formatParaCalculo(grdPrincipal.TextMatrix(grdPrincipal.row, 8)))
    
                carregarDadosGrid
                validaChekin
                atualizaDadosCapaNFCompra
            End If
        Else
            MsgBox "Não é permitido deletar um item da serie NE", vbInformation, "Aviso"
        End If
    End If
End Sub

Private Sub grdReferenciaSemBarras_DblClick()
    grdReferenciaSemBarras.col = 1
    If grdReferenciaSemBarras > grdReferenciaSemBarras.FixedRows Then
        txtAddReferencia.Text = grdReferenciaSemBarras
    End If
End Sub

Private Sub mskDataVencimento_KeyPress(KeyAscii As Integer)
    If Len(mskDataVencimento.ClipText) = 7 Then
        'SendKeys "{TAB}"
    End If
    
    If KeyAscii = 27 Then
        frmVencimento.Visible = False
        Exit Sub
    End If
End Sub

Private Sub txtAddCodigoProduto_GotFocus()
    campoSelecionadoComCaracter txtAddCodigoProduto
End Sub

Private Sub txtAddCodigoProduto_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNormal(KeyAscii)
End Sub

Private Sub txtAddReferencia_GotFocus()
    campoSelecionadoComCaracter txtAddCodigoProduto
End Sub

Private Sub txtAddReferencia_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNumerico(KeyAscii)
End Sub

Private Sub txtBateNota_Change()
    If formatParaCalculo(txtBateNota) >= -0.99 And formatParaCalculo(txtBateNota) <= 0.99 Then
        txtBateNota.ForeColor = vbBlue
    Else
        txtBateNota.ForeColor = vbRed
    End If
End Sub

Private Sub txtCGC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        verificaCampoCarregaCapa
    End If
    KeyAscii = campoNumerico(KeyAscii)
End Sub

'===========================================================================================================
'==================================   LOST FOCUS DE TODOS OS CAMPOS   ======================================
'===========================================================================================================

'==================================   LOST FOCUS ITEM NF COMPRA

Private Sub txtCFOP_LostFocus()
    If txtCFOP = "" And txtReferencia <> "" Then
        txtCFOP = "0"
    End If
End Sub


Private Sub txtCodigoFornecedor_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNumerico(KeyAscii)
End Sub

Private Sub txtCodigoFornecedor_LostFocus()
        If txtCodigoFornecedor <> "" Then
            txtCGC = buscaCodigoFornecedor(Val(txtCodigoFornecedor))
            txtNomeFantasia = buscaNomeFornecedor(Val(txtCodigoFornecedor))
            txtNotaFiscal.SetFocus
        End If
        verificaCampoCarregaCapa
End Sub

Private Sub txtNomeFantasia_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNormal(KeyAscii)
End Sub

Private Sub txtNomeFantasia_LostFocus()
    verificaCampoCarregaCapa
End Sub

Private Sub txtNotaFiscal_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNumerico(KeyAscii)
End Sub

Private Sub txtNotaFiscal_LostFocus()
    verificaCampoCarregaCapa
End Sub

Private Sub txtPedido_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNumerico(KeyAscii)
    If KeyAscii = 13 Then
        carregarDadosGrid
       ' validaChekin
    End If
End Sub

Private Sub txtPedido_LostFocus()
    
    'Dim rdoItemPedido As rdoResultset
    'Dim sql As String
    
    'sql = "select PI_numeroPedido from itempedido where pi_referencia = '4771012' and pi_saldoPedido >= 5  and pi_situacao = 'A'"
    'Set ItemPedido = ADO_Cn_CDLocal.OpenResultset(sql, Options:=rdExecDirect)
    
    'ItemPedido
    
    'select * from itempedido where pi_referencia = '4771012' and pi_saldoPedido >= 5  and pi_situacao = 'A'
    
    carregarDadosGrid
    'validaChekin
    
End Sub

Private Sub txtPorcentagemICMS_LostFocus()
    If txtPorcentagemICMS = "" And txtReferencia <> "" Then
        txtPorcentagemICMS = "0"
    End If
End Sub

Private Sub txtPrecoUnitario_LostFocus()
    If txtPrecoUnitario <> "" Then
        formataCampoDinheiro txtPrecoUnitario
    End If
End Sub

Private Sub txtPercentualDesconto_LostFocus()
    If txtPercentualDesconto = "" And txtReferencia <> "" Then
        txtPercentualDesconto.Text = "0"
    End If
End Sub

Private Sub txtAliquotaIPI_LostFocus()
    If txtAliquotaIPI = "" And txtReferencia <> "" Then
        txtAliquotaIPI = "0"
    End If
End Sub

Private Sub txtQuantidade_LostFocus()
    If txtQuantidade = "" Or txtQuantidade = "0" And txtReferencia <> "" Then
        txtQuantidade.Text = "1"
    End If
End Sub

'==================================   LOST FOCUS CAPA NF COMPRA

Private Sub txtDespesas_LostFocus()
    autoPreencherZero txtDespesas
End Sub

Private Sub txtEmbalagem_LostFocus()
    autoPreencherZero txtEmbalagem
End Sub

Private Sub txtOutros_LostFocus()
    autoPreencherZero txtOutros
End Sub


Private Sub txtBaseICMS_LostFocus()
    autoPreencherZero txtBaseICMS
End Sub

Private Sub txtCGC_LostFocus()
    If validaCampoCNPJ(txtCGC) Then
        txtCodigoFornecedor = buscaCodigoFornecedor(Val(txtCGC.Text))
        txtNomeFantasia = buscaNomeFornecedor(Val(txtCGC.Text))
        txtNotaFiscal.SetFocus
    Else
 
            'txtCodigoFornecedor.SetFocus
    End If
    verificaCampoCarregaCapa
End Sub



'Private Sub verificarReferenciaNoItemNFCompra()
'
'   Dim i As Integer
'   i = 1
'
'   Do While i < grdPrincipal.Rows
'        grdPrincipal.Col = 0
'        grdPrincipal.Row = i
'        If Trim(grdPrincipal) = Trim(txtReferencia) Then
'            abilitaCamposItemNFCompra False
'            grdPrincipal.Col = 4: txtPrecoUnitario = grdPrincipal
'            grdPrincipal.Col = 5: txtCFOP = grdPrincipal
'            grdPrincipal.Col = 6: txtPorcentagemICMS = grdPrincipal
'            grdPrincipal.Col = 7: txtAliquotaIPI = grdPrincipal
'            grdPrincipal.Col = 8: txtPercentualDesconto = grdPrincipal
'            Exit Sub
'        End If
'        i = i + 1
'   Loop
'
'   grdPrincipal.Col = 0
'   grdPrincipal.Row = 0
'
'End Sub



Private Sub txtTeste_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
        'txtTeste.Text = KeyAscii
    Else
        KeyAscii = 0
    End If
    'grdPrincipal.Text = txtTeste.Text
    
    If KeyAscii = 8 Then
        MsgBox "OK"
    End If
    'grdPrincipal.Text = txtTeste.Text
    
End Sub

Private Sub txtTeste_LostFocus()
    'txtTeste.Visible = False
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNormal(KeyAscii)
End Sub

Private Sub txtTotalDesconto_LostFocus()
    autoPreencherZero txtTotalDesconto
End Sub

Private Sub txtValorICMS_LostFocus()
    autoPreencherZero txtValorICMS
End Sub

Private Sub txtFrete_LostFocus()
    autoPreencherZero txtFrete
End Sub

Private Sub txtValorICMSSubsTrib_LostFocus()
    autoPreencherZero txtValorICMSSubsTrib
End Sub

Private Sub txtValorIPI_LostFocus()
    autoPreencherZero txtValorIPI
End Sub

Private Sub txtaliquotaicms_LostFocus()
    autoPreencherZero txtaliquotaicms
End Sub

Private Sub txtValorMercadorias_LostFocus()
    txtValorMercadorias = formatCampoDinheiro(txtValorMercadorias)
End Sub

Private Sub txtValorTotalNota_LostFocus()
    If txtValorTotalNota = "0" Or txtValorTotalNota = "0,00" Then
       txtValorTotalNota.Text = ""
    Else
       formataCampoDinheiro txtValorTotalNota
    End If
End Sub


'===========================================================================================================
'========================================   TODOS OS BOTOES DO FORMULARIO ==================================
'===========================================================================================================


Private Sub cmdGravaCapa()
    
    Screen.MousePointer = 11

    If mskDataEmissao.ClipText = "" Then mskDataEmissao.Text = Date

    If Not verificaExisteCapa Then
        If verificaCamposObrigatorioCapa Then
        
            porcentagemFrete = calculaPorcentagem(formatParaCalculo(txtFrete), formatParaCalculo(txtValorMercadorias))
        
            sql = "exec SP_GravaCapaNFCompra '" & txtNotaFiscal _
            & "', '" & txtSerie & "', '" & txtCodigoFornecedor & "', '" & RTrim(cmbLocal) & "', '" _
            & Format(mskDataEmissao.Text, "mm/dd/yyyy") _
            & "', '" & Format(mskDataRecebimento.Text, "mm/dd/yyyy") & "', '" & Format(txtDataEntrada, "mm/dd/yyyy") & "', '" _
            & formatParaGravar(txtValorMercadorias) & "', '" & formatParaGravar(txtEmbalagem) _
            & "', '" & formatParaGravar(txtFrete) & "', '" & formatParaGravar(txtOutros) & "', '" _
            & formatParaGravar(txtDespesas) & "', '" & formatParaGravar(txtaliquotaicms) _
            & "', '" & formatParaGravar(txtBaseICMS) & "', '" & formatParaGravar(txtValorICMS) _
            & "','" & formatParaGravar(txtValorIPI) & "', '" & formatParaGravar(txtValorICMSSubsTrib) _
            & "', '" & formatParaGravar(txtValorTotalNota) & "', '" & left$(cmbTipoEntrada, 2) & "', '" _
            & formatParaGravar(porcentagemFrete) & "', '" & "D" & "'"
            
            ADO_Cn_CDLocal.Execute sql
            carregarDadosGrid
            'pnlItens.Enabled = True
            If validaChekin Then
            
            abilitaCapa False
            End If
        End If
    End If
    
    Screen.MousePointer = 0
End Sub

Private Sub cmdLimparCamposItemNFCompra_Click()
    limparTodosCampos False
    abilitaCamposItemNFCompra True
    txtReferencia.SetFocus
End Sub

Private Sub cmdLimpar_Click()
    If mensagemLimparCampos Then
        limparTodosCampos True
    End If
End Sub

Private Sub cmdEncerraNota_Click()
    ativaVencimento True
End Sub

Public Sub EncerraNota()
    Screen.MousePointer = 11
    
    Dim rdoitem As New ADODB.Recordset
    Dim rdoPedido As New ADODB.Recordset

    sql = "select count(*) item from ItemNFCompra where ci_notaFiscal = '" _
    & txtNotaFiscal & "' and ci_serie = '" & txtSerie & "' and ci_fornecedor in (select fo_Codigofornecedor from fornecedor where fo_cgc ='" _
    & txtCGC.Text & "') "
    rdoitem.CursorLocation = adUseClient
    rdoitem.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not rdoitem.EOF Then
    
        sql = "select count(ci_nossoPedido) nossoPedido from itemNFcompra where ci_notafiscal = '" _
        & txtNotaFiscal & "' and ci_serie = '" & txtSerie & "' and ci_fornecedor in (select fo_Codigofornecedor from fornecedor where fo_cgc = '" _
        & txtCGC.Text & "') and ci_nossoPedido = 0"
        rdoPedido.CursorLocation = adUseClient
        rdoPedido.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
        If rdoPedido("nossoPedido") = "0" Then
            sql = "exec SP_atualizaCapaNFCompra '" & txtNotaFiscal & "','" & txtSerie & "','" & txtCodigoFornecedor _
            & "','" & Trim(cmbLocal) & "','" & txtTotalCalculado & "','" & "E" & "', '" & Format(mskDataEmissao.Text, "yyyy/mm/dd") _
            & "', '" & Format(mskDataRecebimento.Text, "yyyy/mm/dd") & "', '" & Format(txtDataEntrada.Text, "yyyy/mm/dd") & "'" _
            & vbNewLine & _
            "update ItemNFCompra set ci_situacao = 'E' where ci_notaFiscal = '" & txtNotaFiscal & "' and ci_serie = '" & txtSerie & "' and ci_fornecedor = '" & txtCodigoFornecedor & "'" _
            & vbNewLine & _
            "update VencimentosFornecedor SET vf_situacao = 'E' where vf_notaFiscal = '" & txtNotaFiscal & "' and vf_serie = '" & txtSerie & "' and vf_fornecedor = '" & txtCodigoFornecedor & "'"
            ADO_Cn_CDLocal.Execute sql
            
            limparTodosCampos True
            abilitaCapa True
        ElseIf rdoPedido("nossoPedido") = "1" Then
            MsgBox "Há 1 item que ainda não possui um número de pedido", vbInformation, "Item sem número pedido"
            txtPedido.SetFocus
        Else
            MsgBox "Há " & rdoPedido("nossoPedido") & " itens que ainda não possui um número de pedido", vbInformation, "Item sem número pedido"
            txtPedido.SetFocus
        End If
        
        rdoPedido.Close
        
    Else
        MsgBox "Você precisa adicionar pelo mesmo 1 item", vbInformation, "Não há Itens"
    End If
    
    rdoitem.Close
    Screen.MousePointer = 0
    
    frmStartaProcessos.Show vbModal
    
End Sub


'===========================================================================================================
'========================================   CAMPOS PARA SER PREENCHIDOS  ==================================
'===========================================================================================================

'==================================   ITEM NF COMPRA


Private Sub txtPercentualDesconto_KeyPress(KeyAscii As Integer)
    gravaDadosEnterItemNF KeyAscii
    KeyAscii = campoNumericoVirgula(KeyAscii)
End Sub

Private Sub txtCFOP_KeyPress(KeyAscii As Integer)
    gravaDadosEnterItemNF KeyAscii
    KeyAscii = campoNumericoVirgula(KeyAscii)
End Sub

Private Sub txtPorcentagemICMS_KeyPress(KeyAscii As Integer)
    gravaDadosEnterItemNF KeyAscii
    KeyAscii = campoNumerico(KeyAscii)
End Sub

Private Sub txtPrecoUnitario_KeyPress(KeyAscii As Integer)
    gravaDadosEnterItemNF KeyAscii
    KeyAscii = campoNumericoVirgula(KeyAscii)
End Sub

Private Sub txtQuantidade_KeyPress(KeyAscii As Integer)
    gravaDadosEnterItemNF KeyAscii
    KeyAscii = campoNumerico(KeyAscii)
End Sub

Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNumerico(KeyAscii)
    gravaDadosEnterItemNF KeyAscii
End Sub

Private Sub txtAliquotaIPI_KeyPress(KeyAscii As Integer)
    gravaDadosEnterItemNF KeyAscii
    KeyAscii = campoNumericoVirgula(KeyAscii)
End Sub

Private Sub gravaDadosEnterItemNF(KeyAscii As Integer)
    If KeyAscii = 13 Then
        gravaDadosItemNFCompra
    End If
End Sub

'==================================   ITEM NF COMPRA

Private Sub txtBaseICMS_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNumericoVirgula(KeyAscii)
    If KeyAscii = 13 Then
        cmdGravaCapa
    End If
End Sub

Private Sub txtValorMercadorias_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNumericoVirgula(KeyAscii)
    If KeyAscii = 13 Then
        cmdGravaCapa
    End If
End Sub

Private Sub txtAliquotaICMS_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNumericoVirgula(KeyAscii)
    If KeyAscii = 13 Then
        cmdGravaCapa
    End If
End Sub

Private Sub txtValorICMS_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNumericoVirgula(KeyAscii)
    If KeyAscii = 13 Then
        cmdGravaCapa
    End If
End Sub

Private Sub txtValorICMSSubsTrib_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNumericoVirgula(KeyAscii)
    If KeyAscii = 13 Then
        cmdGravaCapa
    End If
End Sub

Private Sub txtFrete_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNumericoVirgula(KeyAscii)
    If KeyAscii = 13 Then
        cmdGravaCapa
    End If
End Sub

Private Sub txtValorIPI_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNumericoVirgula(KeyAscii)
    If KeyAscii = 13 Then
        cmdGravaCapa
    End If
End Sub

Private Sub txtEmbalagem_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNumericoVirgula(KeyAscii)
    If KeyAscii = 13 Then
        cmdGravaCapa
    End If
End Sub

Private Sub txtDespesas_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNumericoVirgula(KeyAscii)
    If KeyAscii = 13 Then
        cmdGravaCapa
    End If
End Sub

Private Sub txtOutros_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNumericoVirgula(KeyAscii)
    If KeyAscii = 13 Then
        cmdGravaCapa
    End If
End Sub

Private Sub txtValorTotalNota_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNumericoVirgula(KeyAscii)
    If KeyAscii = 13 Then
        cmdGravaCapa
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function verificaCampoCarregaCapa()
    If txtCGC <> "" And txtCodigoFornecedor <> "" And txtSerie <> "" And txtNotaFiscal <> "" Then
        Call verificaExisteCapa
    End If
End Function

Private Sub txtSerie_LostFocus()
txtSerie.Text = UCase(txtSerie.Text)
    verificaCampoCarregaCapa
End Sub
Private Function validaChekin() As Boolean
validaChekin = statusChekin

sql = "select CHM_chekinOK from  ChekinMercadoria WHERE  CHM_NotaFiscal=   '" & txtNotaFiscal & "' AND " _
    & " CHM_Loja=  '" & Trim(cmbLocal) & "' and  CHM_Serie='" & txtSerie & "' and chm_fornecedor in (select fo_codigoFornecedor from fornecedor where fo_cgc = '" & txtCGC.Text & "') " _
    & " group by  CHM_chekinOK"
    rdoChekin.CursorLocation = adUseClient
    rdoChekin.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
  
     If rdoChekin.EOF Then
           Chekin
    ElseIf rdoChekin.RecordCount = 1 Then
        If rdoChekin("CHM_chekinOK") <> "S" Then
            Chekin
        End If
    ElseIf rdoChekin.RecordCount > 1 Then
          Chekin
     End If
     rdoChekin.Close
     validaChekin = statusChekin
End Function
Private Function verificaExisteCapa() As Boolean
Screen.MousePointer = 11

Dim rdoConsultarCapaD As New ADODB.Recordset
Dim rdoConsultaItens As New ADODB.Recordset
Dim Registro As Integer
        
        sql = "select * from CapaNFCompra where cc_notaFiscal = '" & txtNotaFiscal & "' and cc_serie = '" _
        & txtSerie & "' and cc_fornecedor in (select fo_codigofornecedor from fornecedor where fo_cgc = '" & txtCGC.Text & "')"
        rdoConsultarCapaD.CursorLocation = adUseClient
        rdoConsultarCapaD.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic

        If Not rdoConsultarCapaD.BOF And Not rdoConsultarCapaD.EOF Then
            If Trim(rdoConsultarCapaD("cc_notafiscal")) = txtNotaFiscal And _
            rdoConsultarCapaD("cc_serie") = txtSerie And _
            rdoConsultarCapaD("cc_situacao") = "D" Then
            mskDataEmissao.Text = Format(rdoConsultarCapaD("cc_dataEmissao"), "dd/mm/yyyy")
                
                atualizaLojaCapaNFCompra txtNotaFiscal, txtSerie, txtCGC.Text, cmbLocal
                
                If Format(rdoConsultarCapaD("cc_dataRecebimento"), "dd/mm/yyyy") <> "01/01/1900" Then
                    mskDataRecebimento = Format(rdoConsultarCapaD("cc_dataRecebimento"), "dd/mm/yyyy")
                End If
                
                txtValorMercadorias = formatCampoDinheiro(rdoConsultarCapaD("cc_ValorMercadorias"))
                txtBaseICMS = formatCampoDinheiro(rdoConsultarCapaD("cc_BaseICMS"))
                txtaliquotaicms = formatCampoDinheiro(rdoConsultarCapaD("cc_AliquotaICMS"))
                txtValorICMS = formatCampoDinheiro(rdoConsultarCapaD("cc_ValorICMS"))
                txtValorICMSSubsTrib = formatCampoDinheiro(rdoConsultarCapaD("cc_ValorICMSSubsTrib"))
                txtFrete = formatCampoDinheiro(rdoConsultarCapaD("cc_Frete"))
                txtValorIPI = formatCampoDinheiro(rdoConsultarCapaD("cc_ValorIPI"))
                txtEmbalagem = formatCampoDinheiro(rdoConsultarCapaD("cc_Embalagem"))
                txtDespesas = formatCampoDinheiro(rdoConsultarCapaD("cc_Despesas"))
                txtOutros = formatCampoDinheiro(rdoConsultarCapaD("cc_Outros"))
                txtValorTotalNota = formatCampoDinheiro(rdoConsultarCapaD("cc_ValorTotalNota"))
                txtTotalCalculado = formatCampoDinheiro(rdoConsultarCapaD("cc_ValorCalculado"))
                txtTotalDesconto = formatCampoDinheiro(rdoConsultarCapaD("CC_Desconto"))
                mskDataRecebimento.Enabled = True
                'cmbLocal.Enabled = True
                'cmbLocal.Clear
                'cmbLocal.AddItem rdoConsultarCapaD("CC_Loja")
                'cmbLocal.ListIndex = 0
                
  '-----------------------------------------------------------------------------------------------------------
         ' 2) Pesquisa Tipo Entrada
         
                sql = "select top 1 * from itemnfcompra, Capapedido,TipoPedidoCompra where ci_nossoPedido = pc_numeropedido and pc_tipoPedido = tpc_codigo and ci_nossoPedido <> 0 and" _
                & " ci_notaFiscal = '" & txtNotaFiscal & "' and ci_serie = '" _
                & txtSerie & "' and ci_fornecedor in (select fo_codigofornecedor from fornecedor where fo_cgc = '" & txtCGC.Text & "') "
                rdoConsultaItens.CursorLocation = adUseClient
                rdoConsultaItens.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
               
                If Not rdoConsultaItens.EOF Then
                    cmbTipoEntrada.Text = Trim(rdoConsultaItens("tpc_descricao"))
                End If
              
                rdoConsultaItens.Close

                verificaExisteCapa = True
                carregarDadosGrid
                If validaChekin Then
                    abilitaCapa False
                
                    If txtNomeFantasia = "" Then
                        txtNomeFantasia = buscaNomeFornecedor(txtCGC)
                    End If
                
                ElseIf rdoConsultarCapaD("cc_situacao") <> "D" Then
                    MsgBox "Nota Fiscal já cadastrada", vbInformation, "NFe"
                    txtCGC.SetFocus
                    limparTodosCampos True
                    verificaExisteCapa = True
                End If
            Else
                verificaExisteCapa = False
            End If
End If
    rdoConsultarCapaD.Close

Screen.MousePointer = 0
End Function

Private Sub abilitaCapa(abilita As Boolean)
    If abilita Then
       ' frameItens.Enabled = False
        frmEntrada.Enabled = True
        cmdExcluiNota.Enabled = False
        cmdEncerraNota.Enabled = False
        cmdLimpar.Enabled = True
        grdPrincipal.Enabled = False
        txtCGC.Enabled = True
        txtCodigoFornecedor.Enabled = True
        txtNomeFantasia.Enabled = True
        txtNotaFiscal.Enabled = True
        txtSerie.Enabled = True
        cmbTipoEntrada.Enabled = True
        txtDataEntrada.Enabled = True
        mskDataEmissao.Enabled = True
        txtValorMercadorias.Enabled = True
        txtBaseICMS.Enabled = True
        txtTotalDesconto.Enabled = True
        txtaliquotaicms.Enabled = True
        txtValorICMS.Enabled = True
        txtValorICMSSubsTrib.Enabled = True
        txtFrete.Enabled = True
        txtValorIPI.Enabled = True
        txtEmbalagem.Enabled = True
        txtDespesas.Enabled = True
        'txtOutras.Enabled = True
        txtValorTotalNota.Enabled = True
        txtBateNota.Enabled = True
        txtTotalCalculado.Enabled = True
        txtCGC.SetFocus
    Else
        frameItens.Enabled = True
        'frmEntrada.Enabled = False
        txtCGC.Enabled = False
        txtCodigoFornecedor.Enabled = False
        txtNomeFantasia.Enabled = False
        txtNotaFiscal.Enabled = False
        txtSerie.Enabled = False
        cmbTipoEntrada.Enabled = False
        txtDataEntrada.Enabled = False
        mskDataEmissao.Enabled = False
        txtValorMercadorias.Enabled = False
        txtBaseICMS.Enabled = False
        txtTotalDesconto.Enabled = False
        txtaliquotaicms.Enabled = False
        txtValorICMS.Enabled = False
        txtValorICMSSubsTrib.Enabled = False
        txtFrete.Enabled = False
        txtValorIPI.Enabled = False
        txtEmbalagem.Enabled = False
        txtDespesas.Enabled = False
        'txtOutras.Enabled = False
        txtValorTotalNota.Enabled = False
        txtBateNota.Enabled = False
        txtTotalCalculado.Enabled = False
        cmdExcluiNota.Enabled = True
        cmdEncerraNota.Enabled = True
        cmdLimpar.Enabled = True
        grdPrincipal.Enabled = True
        txtReferencia.SetFocus
    End If
End Sub




'////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////   FRM VENCIMENTO   ////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub grdVencimentos_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
        grdVencimentos.col = 0
        If mensagemExluir("Parcela " & grdVencimentos) Then
            sql = "SP_deletaVencimentosFornecedor '" & txtNotaFiscal & _
            "','" & txtSerie & "','" & txtCodigoFornecedor & _
            "','" & cmbLocal & "','" & grdVencimentos & "'"
            ADO_Cn_CDLocal.Execute sql
            carregaGridVencimentos
        End If
    End If
End Sub

Private Sub GravaVenctoFornecedor()
Screen.MousePointer = 11
    Dim sql As String
        
    sql = "SP_inserir_Vencimento_Fornecedor '" & txtNotaFiscal & "','" & txtSerie & _
    "','" & txtCodigoFornecedor & "','" & Format(mskDataVencimento, "yyyy/mm/dd") & "','2','" & _
    ConverteVirgula(Format(txtValor, "0.00")) & "'"
    ADO_Cn_CDLocal.Execute sql
    
Screen.MousePointer = 0
End Sub

Private Sub ativaVencimento(ativa As Boolean)
    If ativa Then
        frmVencimento.Visible = True
        mskDataVencimento.SetFocus
        carregaGridVencimentos
        grdPrincipal.Enabled = False
     '   frameItens.Enabled = False
        cmdLimpar.Enabled = False
        cmdExcluiNota.Enabled = False
        cmdEncerraNota.Enabled = False
    Else
        frmVencimento.Visible = False
        EncerraNota
        limpaVencimento
        limparTodosCampos True
    End If
End Sub

Private Sub limpaVencimento()
    mskDataVencimento = "__/__/____"
    txtValor.Text = ""
    grdVencimentos.Rows = 1
End Sub


Private Sub cmdOK_Click()
        If grdVencimentos.Rows > 1 Then
            ativaVencimento False
        Else
            MsgBox "Nenhum Vencimento foi especificado.", vbExclamation, "Atenção"
        End If
End Sub


Private Sub carregaGridVencimentos()
    Screen.MousePointer = 11
    
    Dim adoVenctoForne As New ADODB.Recordset
    Dim Valor As Long
    grdVencimentos.Rows = 1
    
    'If MDISup.ActiveForm.Name = "frmConsInformativo" Then
    '    Valor = Val(MDISup.ActiveForm.txtCodFornecedor.Text)
    'Else
    '    Valor = Val(MDISup.ActiveForm.txtNomeFantasia.Text)
    'End If
    
    
    sql = "Select VF_Parcela, VF_DataVencimento, VF_TipoPagamento,VF_ValorParcela from VencimentosFornecedor Where " _
    & "VF_NotaFiscal = " & txtNotaFiscal & " and " _
    & "VF_Serie = '" & txtSerie & "'" _
    & " and vf_fornecedor = " & txtCodigoFornecedor.Text _
    & " order by VF_Parcela"
    
        
        adoVenctoForne.CursorLocation = adUseClient
        adoVenctoForne.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not adoVenctoForne.EOF Then
        'DisableClick = True
        
        'optPagto(adoVenctoForne("VF_TipoPagamento")).Value = True
        
        'DisableClick = False
        
        Do While Not adoVenctoForne.EOF
            grdVencimentos.AddItem adoVenctoForne("VF_Parcela") & Chr(9) & adoVenctoForne("VF_DataVencimento") _
            & Chr(9) & Format(adoVenctoForne("VF_ValorParcela"), "0.00")
            
            adoVenctoForne.MoveNext
        Loop
        
        'grdVencimentos.RemoveItem 1
    End If
    
    adoVenctoForne.Close
    
    Screen.MousePointer = 0
End Sub


Private Sub txtValor_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNumericoVirgula(KeyAscii)
    
    If KeyAscii = 27 Then
        frmVencimento.Visible = False
        Exit Sub
    End If
    
    If KeyAscii = vbKeyReturn Then
        If Trim(txtValor.Text) = "" Or Trim(txtValor.Text) = "'" Then
            MsgBox "Informe o valor da parcela!"
            Exit Sub
        Else
            'mskDataVencimento.SetFocus
        End If

        KeyAscii = 0
        
        If InStr(mskDataVencimento.Text, "_") = 0 Then
            If IsDate(mskDataVencimento.Text) Then
            
'                grdVencimentos.Row = grdVencimentos.Rows - 1
'                grdVencimentos.Col = 0
'
'                ''IF
'
'                SQL = "Insert Into VencimentosFornecedor (" _
'                & "VF_NotaFiscal,VF_Serie,VF_Fornecedor,VF_DataVencimento," _
'                & "VF_Parcela,VF_ValorParcela, VF_TipoPagamento) values (" _
'                & txtNotafiscal & ",'" _
'                & txtSerie & "'," _
'                & txtCodigoFornecedor & ",'" _
'                & Format(mskDataVencimento, "yyyy/mm/dd") & "'," _
'                & Trim(grdVencimentos) + 1 _
'                & ", " & ConverteVirgula(Format(txtValor, "0.00")) _
'                & ", 2" & ")"
                
                'ADO_Cn_CDLocal.Execute SQL
                
                GravaVenctoFornecedor
                
                mskDataVencimento.Text = "__/__/____"
                mskDataVencimento.SetFocus
                txtValor.Text = ""
                carregaGridVencimentos
                Exit Sub
            End If
        End If
        
        MsgBox "Data inválida.", vbExclamation, "Atenção"
    End If

End Sub

Private Sub abilitaFrameCadastraReferencia(ativa As Boolean)

    If ativa Then
        frameCadastraReferencia.Visible = True
        carregaGridSemReferencia
    Else
        frameCadastraReferencia.Visible = False
        txtAddCodigoProduto = ""
        txtAddReferencia = ""
    End If
    
End Sub

Private Sub carregaGridSemReferencia()
    Dim rdoReferenciaSemBarras As New ADODB.Recordset
    
    limpaGrid grdReferenciaSemBarras
    
    sql = "Select PR_Referencia,PR_Descricao from produto where PR_codigoFornecedor = " & txtCodigoFornecedor & _
    " and pr_referencia not in (select PRB_referencia from produtobarras where prb_codigoFornecedor = " & txtCodigoFornecedor & _
    " and PRB_TipoCodigo = 'B') order by PR_Descricao"
    
    rdoReferenciaSemBarras.CursorLocation = adUseClient
    rdoReferenciaSemBarras.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    Do While Not rdoReferenciaSemBarras.EOF
        grdReferenciaSemBarras.AddItem "" & Chr(9) & rdoReferenciaSemBarras("PR_Referencia") & Chr(9) & _
        rdoReferenciaSemBarras("PR_Descricao")
        rdoReferenciaSemBarras.MoveNext
    Loop
    
End Sub
Private Sub Chekin()

 If MsgBox("ATENÇÃO! Não foi realizado o chekin dessa nota fiscal." & vbNewLine & _
                "Deseja continuar?", vbExclamation + vbYesNo, "Chekin não encontrado") = vbYes Then
                If MsgBox("ATENÇÃO! Chekin será ignorado e marcado como feito pelo sistema. Esse procedimento não poderá ser desfeito." & vbNewLine & _
                "Deseja continuar?", vbExclamation + vbYesNo, "Chekin não encontrado") = vbYes Then
                       sql = "exec SP_ComparaChekinNFCompras '" & Trim(cmbLocal) & "','" & txtNotaFiscal _
                            & "','" & txtSerie & "','" & txtCodigoFornecedor & "'"
                        ADO_Cn_CDLocal.BeginTrans
                        ADO_Cn_CDLocal.Execute sql
                        ADO_Cn_CDLocal.CommitTrans
                        
                        sql = "update ChekinMercadoria  set  CHM_chekinOK='S' WHERE  CHM_NotaFiscal=   '" & txtNotaFiscal & "' AND " _
                        & "CHM_Fornecedor in (select fo_codigoFornecedor from fornecedor where fo_cgc = '" & txtCGC.Text & "')  and  CHM_Loja=  '" & Trim(cmbLocal) & "' and  CHM_Serie='" & txtSerie & "' "
                        ADO_Cn_CDLocal.BeginTrans
                        ADO_Cn_CDLocal.Execute sql
                        ADO_Cn_CDLocal.CommitTrans
                      
                        statusChekin = True
                    
                        
                Else
                      limparTodosCampos True
                       statusChekin = False
                     
    End If
            Else
            limparTodosCampos True
            statusChekin = False
       
            
            End If
End Sub


