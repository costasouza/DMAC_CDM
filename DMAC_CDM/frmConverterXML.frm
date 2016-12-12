VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmConverterXML 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Converter XML"
   ClientHeight    =   10980
   ClientLeft      =   135
   ClientTop       =   525
   ClientWidth     =   16680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10980
   ScaleWidth      =   16680
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbTipoEntrada 
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
      Left            =   8670
      TabIndex        =   103
      Top             =   135
      Width           =   2805
   End
   Begin VB.TextBox txtDataRecebimento 
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
      Height          =   285
      Left            =   5970
      TabIndex        =   102
      Top             =   135
      Width           =   1440
   End
   Begin VB.TextBox txtFornecedor 
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
      Height          =   285
      Left            =   1125
      TabIndex        =   0
      Top             =   135
      Width           =   3810
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   60
      ScaleHeight     =   45
      ScaleWidth      =   15240
      TabIndex        =   95
      Top             =   10410
      Width           =   15240
   End
   Begin VB.Frame fraCaminhoXML 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Caminho XML"
      ForeColor       =   &H00FF8080&
      Height          =   855
      Left            =   105
      TabIndex        =   93
      Top             =   510
      Width           =   12780
      Begin VB.TextBox Text1 
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
         Height          =   285
         Left            =   45
         TabIndex        =   94
         Top             =   405
         Width           =   6735
      End
      Begin CentroDeDistribuicao.chameleonButton cmdBuscaXML 
         Height          =   300
         Left            =   6795
         TabIndex        =   101
         Top             =   390
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         BTYPE           =   14
         TX              =   "..."
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
         MICON           =   "frmConverterXML.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "Caminho XML"
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
         Left            =   45
         TabIndex        =   107
         Top             =   60
         Width           =   1260
      End
   End
   Begin VB.Frame fraTransp 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Transp"
      ForeColor       =   &H00FF8080&
      Height          =   1935
      Left            =   11610
      TabIndex        =   83
      Top             =   5100
      Width           =   3645
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   62
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   91
         Top             =   135
         Width           =   1530
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   63
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   90
         Top             =   360
         Width           =   1530
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   64
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   89
         Top             =   585
         Width           =   1530
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   65
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   88
         Top             =   825
         Width           =   1530
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   66
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   1065
         Width           =   1530
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   67
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   86
         Top             =   1305
         Width           =   1530
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   68
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   85
         Top             =   1545
         Width           =   1530
      End
      Begin VSFlex7DAOCtl.VSFlexGrid VSFlexGrid5 
         Height          =   1695
         Left            =   765
         TabIndex        =   84
         Top             =   150
         Width           =   900
         _cx             =   1587
         _cy             =   2990
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   0
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
         BackColor       =   16568794
         ForeColor       =   -2147483640
         BackColorFixed  =   0
         ForeColorFixed  =   16423203
         BackColorSel    =   16568794
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16568794
         BackColorAlternate=   16568794
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   7
         Cols            =   1
         FixedRows       =   7
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmConverterXML.frx":001C
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
         BackColorFrozen =   16568794
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Label Label11 
         BackColor       =   &H00404040&
         Caption         =   "Transp"
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
         Left            =   15
         TabIndex        =   117
         Top             =   15
         Width           =   690
      End
   End
   Begin VB.Frame fraCobr 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Cobr"
      ForeColor       =   &H00FF8080&
      Height          =   2370
      Left            =   11610
      TabIndex        =   81
      Top             =   7080
      Width           =   3645
      Begin VSFlex7DAOCtl.VSFlexGrid grdCobr 
         Height          =   975
         Left            =   90
         TabIndex        =   82
         Top             =   300
         Width           =   3465
         _cx             =   6112
         _cy             =   1720
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
         FormatString    =   $"frmConverterXML.frx":00A3
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
      Begin VSFlex7DAOCtl.VSFlexGrid grdCobrFat 
         Height          =   945
         Left            =   90
         TabIndex        =   97
         Top             =   1335
         Width           =   3465
         _cx             =   6112
         _cy             =   1667
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
         FormatString    =   $"frmConverterXML.frx":0105
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
      Begin VB.Label Label12 
         BackColor       =   &H00404040&
         Caption         =   "Cobr"
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
         Left            =   45
         TabIndex        =   118
         Top             =   30
         Width           =   435
      End
   End
   Begin VB.Frame fraTotal 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Total"
      ForeColor       =   &H00FF8080&
      Height          =   4350
      Left            =   9255
      TabIndex        =   65
      Top             =   5100
      Width           =   2295
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   48
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   80
         Top             =   675
         Width           =   1095
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   49
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   885
         Width           =   1095
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   50
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   78
         Top             =   1125
         Width           =   1095
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   51
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   77
         Top             =   1365
         Width           =   1095
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   52
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   76
         Top             =   1605
         Width           =   1095
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   53
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   1845
         Width           =   1095
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   54
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   2085
         Width           =   1095
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   55
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   2340
         Width           =   1095
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   56
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   2565
         Width           =   1095
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   57
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   2805
         Width           =   1095
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   58
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   3045
         Width           =   1095
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   59
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   3270
         Width           =   1095
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   60
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   3510
         Width           =   1095
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   61
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   3765
         Width           =   1095
      End
      Begin VSFlex7DAOCtl.VSFlexGrid VSFlexGrid4 
         Height          =   3360
         Left            =   180
         TabIndex        =   66
         Top             =   690
         Width           =   885
         _cx             =   1561
         _cy             =   5927
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   0
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
         BackColor       =   16568794
         ForeColor       =   -2147483640
         BackColorFixed  =   0
         ForeColorFixed  =   16423203
         BackColorSel    =   16568794
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16568794
         BackColorAlternate=   16568794
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   14
         Cols            =   1
         FixedRows       =   14
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmConverterXML.frx":0167
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
         BackColorFrozen =   16568794
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Label Label10 
         BackColor       =   &H00404040&
         Caption         =   "Total"
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
         Left            =   75
         TabIndex        =   116
         Top             =   165
         Width           =   435
      End
   End
   Begin VB.Frame fraDetImp 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Det/Imposto"
      ForeColor       =   &H00FF8080&
      Height          =   1545
      Left            =   9255
      TabIndex        =   63
      Top             =   3495
      Width           =   6000
      Begin VSFlex7DAOCtl.VSFlexGrid grdImposto 
         Height          =   1200
         Left            =   675
         TabIndex        =   64
         Top             =   255
         Width           =   5205
         _cx             =   9181
         _cy             =   2117
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
         BackColorSel    =   5263440
         ForeColorSel    =   8388608
         BackColorBkg    =   5263440
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
         Cols            =   26
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmConverterXML.frx":01EB
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
      Begin VB.Label Label9 
         BackColor       =   &H00404040&
         Caption         =   "Det/Imp"
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
         Left            =   30
         TabIndex        =   115
         Top             =   15
         Width           =   1260
      End
   End
   Begin VB.Frame fraDet 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Det"
      ForeColor       =   &H00FF8080&
      Height          =   1545
      Left            =   9270
      TabIndex        =   61
      Top             =   1890
      Width           =   5985
      Begin VSFlex7DAOCtl.VSFlexGrid grdDet 
         Height          =   1215
         Left            =   675
         TabIndex        =   62
         Top             =   240
         Width           =   5190
         _cx             =   9155
         _cy             =   2143
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
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmConverterXML.frx":04B7
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
      Begin VB.Label Label8 
         BackColor       =   &H00404040&
         Caption         =   "Det"
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
         Left            =   75
         TabIndex        =   114
         Top             =   165
         Width           =   435
      End
   End
   Begin VB.Frame fraEnderDest 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Ender Dest"
      ForeColor       =   &H00FF8080&
      Height          =   3225
      Left            =   4560
      TabIndex        =   49
      Top             =   7035
      Width           =   4605
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   37
         Left            =   825
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   495
         Width           =   3705
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   38
         Left            =   825
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   720
         Width           =   3705
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   39
         Left            =   825
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   945
         Width           =   3705
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   40
         Left            =   825
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   1200
         Width           =   3705
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   41
         Left            =   825
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   1425
         Width           =   3705
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   42
         Left            =   825
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   1650
         Width           =   3705
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   43
         Left            =   825
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   1905
         Width           =   3705
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   44
         Left            =   825
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   2160
         Width           =   3705
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   45
         Left            =   825
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   2385
         Width           =   3705
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   270
         Index           =   46
         Left            =   825
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   2640
         Width           =   3705
      End
      Begin VSFlex7DAOCtl.VSFlexGrid VSFlexGrid3 
         Height          =   2400
         Left            =   90
         TabIndex        =   50
         Top             =   495
         Width           =   750
         _cx             =   1323
         _cy             =   4233
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   0
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
         BackColor       =   16568794
         ForeColor       =   -2147483640
         BackColorFixed  =   0
         ForeColorFixed  =   16423203
         BackColorSel    =   16568794
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16568794
         BackColorAlternate=   16568794
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   1
         FixedRows       =   10
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmConverterXML.frx":0618
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
         BackColorFrozen =   16568794
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404040&
         Caption         =   "Ender Dest"
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
         Left            =   90
         TabIndex        =   113
         Top             =   165
         Width           =   1260
      End
   End
   Begin VB.Frame fraDest 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Dest"
      ForeColor       =   &H00FF8080&
      Height          =   1575
      Left            =   4560
      TabIndex        =   44
      Top             =   5310
      Width           =   4620
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   35
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   495
         Width           =   3705
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   36
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   720
         Width           =   3705
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   47
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   945
         Width           =   3705
      End
      Begin VSFlex7DAOCtl.VSFlexGrid VSFlexGrid2 
         Height          =   705
         Left            =   105
         TabIndex        =   45
         Top             =   495
         Width           =   750
         _cx             =   1323
         _cy             =   1244
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   0
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
         BackColor       =   16568794
         ForeColor       =   -2147483640
         BackColorFixed  =   0
         ForeColorFixed  =   16423203
         BackColorSel    =   16568794
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16568794
         BackColorAlternate=   16568794
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   1
         FixedRows       =   3
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmConverterXML.frx":067E
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
         BackColorFrozen =   16568794
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Label Label6 
         BackColor       =   &H00404040&
         Caption         =   "Dest"
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
         Left            =   105
         TabIndex        =   112
         Top             =   195
         Width           =   1260
      End
   End
   Begin VB.Frame fraEnderEmit 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Ender Emit"
      ForeColor       =   &H00FF8080&
      Height          =   3315
      Left            =   4545
      TabIndex        =   28
      Top             =   1890
      Width           =   4635
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   21
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   450
         Width           =   3705
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   22
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   705
         Width           =   3705
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   23
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   900
         Width           =   3705
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   24
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   1125
         Width           =   3705
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   25
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   1350
         Width           =   3705
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   26
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1605
         Width           =   3705
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   27
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1845
         Width           =   3705
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   28
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   2085
         Width           =   3705
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   29
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   2340
         Width           =   3705
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   30
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   2565
         Width           =   3705
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   31
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   2805
         Width           =   3705
      End
      Begin VSFlex7DAOCtl.VSFlexGrid VSFlexGrid1 
         Height          =   2640
         Left            =   105
         TabIndex        =   29
         Top             =   420
         Width           =   750
         _cx             =   1323
         _cy             =   4657
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   0
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
         BackColor       =   16568794
         ForeColor       =   -2147483640
         BackColorFixed  =   0
         ForeColorFixed  =   16423203
         BackColorSel    =   16568794
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16568794
         BackColorAlternate=   16568794
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   11
         Cols            =   1
         FixedRows       =   11
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmConverterXML.frx":06B8
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
         BackColorFrozen =   16568794
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404040&
         Caption         =   "Ender Emit"
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
         Left            =   90
         TabIndex        =   111
         Top             =   135
         Width           =   1260
      End
   End
   Begin VB.Frame fraEmit 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Emit"
      ForeColor       =   &H00FF8080&
      Height          =   2850
      Left            =   105
      TabIndex        =   23
      Top             =   7395
      Width           =   4395
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   240
         Index           =   71
         Left            =   795
         Locked          =   -1  'True
         TabIndex        =   96
         Top             =   2160
         Width           =   3540
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   32
         Left            =   795
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   1455
         Width           =   3540
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   33
         Left            =   795
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   1680
         Width           =   3540
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   34
         Left            =   795
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   1935
         Width           =   3540
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   18
         Left            =   795
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   720
         Width           =   3540
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   19
         Left            =   795
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   975
         Width           =   3540
      End
      Begin VB.TextBox txtValor 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Index           =   20
         Left            =   795
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1200
         Width           =   3540
      End
      Begin VSFlex7DAOCtl.VSFlexGrid grdEmit 
         Height          =   1680
         Left            =   60
         TabIndex        =   24
         Top             =   720
         Width           =   750
         _cx             =   1323
         _cy             =   2963
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   0
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
         BackColor       =   16568794
         ForeColor       =   -2147483640
         BackColorFixed  =   0
         ForeColorFixed  =   16423203
         BackColorSel    =   16568794
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16568794
         BackColorAlternate=   16568794
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   7
         Cols            =   1
         FixedRows       =   7
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmConverterXML.frx":0724
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
         BackColorFrozen =   16568794
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404040&
         Caption         =   "Emit"
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
         Left            =   75
         TabIndex        =   110
         Top             =   210
         Width           =   1260
      End
   End
   Begin VB.Frame fraIde 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Ide"
      ForeColor       =   &H00FF8080&
      Height          =   5430
      Left            =   120
      TabIndex        =   3
      Top             =   1890
      Width           =   4380
      Begin VB.TextBox txtValor 
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
         Height          =   285
         Index           =   0
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   585
         Width           =   3465
      End
      Begin VB.TextBox txtValor 
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
         Height          =   285
         Index           =   1
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   825
         Width           =   3465
      End
      Begin VB.TextBox txtValor 
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
         Height          =   285
         Index           =   2
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1065
         Width           =   3465
      End
      Begin VB.TextBox txtValor 
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
         Height          =   285
         Index           =   3
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1305
         Width           =   3465
      End
      Begin VB.TextBox txtValor 
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
         Height          =   285
         Index           =   4
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1545
         Width           =   3465
      End
      Begin VB.TextBox txtValor 
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
         Height          =   285
         Index           =   5
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1785
         Width           =   3465
      End
      Begin VB.TextBox txtValor 
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
         Height          =   285
         Index           =   6
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2025
         Width           =   3465
      End
      Begin VB.TextBox txtValor 
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
         Height          =   285
         Index           =   7
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2265
         Width           =   3465
      End
      Begin VB.TextBox txtValor 
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
         Height          =   285
         Index           =   8
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2475
         Width           =   3465
      End
      Begin VB.TextBox txtValor 
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
         Height          =   285
         Index           =   9
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2730
         Width           =   3465
      End
      Begin VB.TextBox txtValor 
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
         Height          =   285
         Index           =   10
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2970
         Width           =   3465
      End
      Begin VB.TextBox txtValor 
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
         Height          =   285
         Index           =   11
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   3210
         Width           =   3465
      End
      Begin VB.TextBox txtValor 
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
         Height          =   285
         Index           =   12
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   3450
         Width           =   3465
      End
      Begin VB.TextBox txtValor 
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
         Height          =   285
         Index           =   13
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   3690
         Width           =   3465
      End
      Begin VB.TextBox txtValor 
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
         Height          =   285
         Index           =   14
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   3930
         Width           =   3465
      End
      Begin VB.TextBox txtValor 
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
         Height          =   285
         Index           =   15
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   4170
         Width           =   3465
      End
      Begin VB.TextBox txtValor 
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
         Height          =   285
         Index           =   16
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   4425
         Width           =   3465
      End
      Begin VB.TextBox txtValor 
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
         Height          =   285
         Index           =   17
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   4650
         Width           =   3465
      End
      Begin VSFlex7DAOCtl.VSFlexGrid grdIde 
         Height          =   4320
         Left            =   60
         TabIndex        =   22
         Top             =   585
         Width           =   750
         _cx             =   1323
         _cy             =   7620
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   0
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
         BackColor       =   16568794
         ForeColor       =   -2147483640
         BackColorFixed  =   0
         ForeColorFixed  =   16423203
         BackColorSel    =   16568794
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16568794
         BackColorAlternate=   16568794
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   18
         Cols            =   1
         FixedRows       =   18
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmConverterXML.frx":0774
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
         BackColorFrozen =   16568794
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         Caption         =   "Ide"
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
         Left            =   60
         TabIndex        =   109
         Top             =   195
         Width           =   1260
      End
   End
   Begin VB.TextBox txtValor 
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
      Height          =   285
      Index           =   70
      Left            =   570
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1455
      Width           =   6315
   End
   Begin VB.TextBox txtValor 
      BackColor       =   &H00A3A3A3&
      Height          =   720
      Index           =   69
      Left            =   10140
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   9510
      Width           =   5115
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdInfAd 
      Height          =   240
      Left            =   9240
      TabIndex        =   92
      Top             =   9510
      Width           =   885
      _cx             =   1561
      _cy             =   423
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   0
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
      BackColor       =   16568794
      ForeColor       =   -2147483640
      BackColorFixed  =   0
      ForeColorFixed  =   16423203
      BackColorSel    =   16568794
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16568794
      BackColorAlternate=   16568794
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
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
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmConverterXML.frx":0819
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
      BackColorFrozen =   16568794
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin CentroDeDistribuicao.chameleonButton cmdRetornar 
      Height          =   510
      Left            =   13875
      TabIndex        =   98
      Top             =   10545
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
      MICON           =   "frmConverterXML.frx":084E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdXML 
      Height          =   510
      Left            =   12435
      TabIndex        =   99
      Top             =   10545
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "XML"
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
      MICON           =   "frmConverterXML.frx":086A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdCarregaXML 
      Height          =   510
      Left            =   11010
      TabIndex        =   100
      Top             =   10545
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Carrega XML"
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
      MICON           =   "frmConverterXML.frx":0886
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
      BackColor       =   &H00505050&
      Caption         =   "NFE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   120
      TabIndex        =   108
      Top             =   1470
      Width           =   405
   End
   Begin VB.Label lblTipoEntrada 
      BackColor       =   &H00505050&
      Caption         =   "Tipo Entrada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   7500
      TabIndex        =   106
      Top             =   195
      Width           =   1110
   End
   Begin VB.Label lblDataRecebimento 
      BackColor       =   &H00505050&
      Caption         =   "Data Rec."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   5025
      TabIndex        =   105
      Top             =   195
      Width           =   885
   End
   Begin VB.Label lblFornecedor 
      BackColor       =   &H00505050&
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
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   105
      TabIndex        =   104
      Top             =   195
      Width           =   1005
   End
End
Attribute VB_Name = "frmConverterXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
    Dim I As Integer
    Dim oXML As DOMDocument30
    Dim ndeAux As IXMLDOMNode
    Dim strXML As String
    Dim strTeste As String
    Dim Cont As Integer
    Dim qtde As Integer
    Dim cnpj As String
    Dim vet(0 To 70) As String
    Dim Valor As Double
     Dim valor2 As Double
     Dim referencia As String
    Dim verReferencia As Boolean
    Dim valorMerc As Double
    Dim totalCalculado As Double
    Dim vlrMercadoria As Double
    Dim vlrIcms As Double
    Dim vlrIPI As Double
    Dim vlrDespesa As Double
    Dim vlrJuros As Double
    Dim outrasDespesas As Double
    Dim embalagem As Double
    Dim frete As Double
    Dim vlrICMSST As Double
    Dim totalNota As Double

Private Sub cmdRetornar_Click()
    Unload Me
End Sub

Private Sub cmdBuscaXML_Click()
    
    Unload Me
    telaChamou = "frmConverterXML"
End Sub

Private Sub cmdCarregaXML_Click()
   If InStr(Text1.Text, "c:\teste xml\") = 0 Then
       MsgBox "Caminho inválido!" & vbCrLf & "Selecione um arquivo XML do diretório: c:\teste xml\", vbCritical, "ATENÇÃO"
       Exit Sub
    End If
    If txtFornecedor.Text = "" Then
        MsgBox "Informe o fornecedor!", vbCritical, "ATENÇÃO"
        txtFornecedor.SetFocus
        Exit Sub
    End If
    
    If txtDataRecebimento.Text = "" Then
        MsgBox "Informe a data de recebimento!", vbCritical, "ATENÇÃO"
        txtDataRecebimento.SetFocus
        Exit Sub
    End If
    
    telaChamou = ""
    
    Call converterXML
   
End Sub

Private Sub converterXML()

    Set oXML = New DOMDocument30
       
    strXML = Me.Text1.Text

    If oXML.Load(strXML) Then
        For Each ndeAux In oXML.selectNodes("/nfeProc")
            strTeste = ndeAux.XML
            

            Call limpaTela
            'emit------------------------------------------------------------------------------------------
            
           
           If ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/emit").FirstChild.BaseName = "CPF" Then
                
                If Trim(cnpj) <> Trim(ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/emit/CNPJ").nodeTypedValue) Then
                    MsgBox "O fornecedor informado difere do fornecedor da nota!", vbCritical, "ATENÇÃO"
                    txtValor(18).Text = ""
                    txtFornecedor.SetFocus
                    Exit Sub
                End If
                txtValor(18).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/emit/CPF").nodeTypedValue
           Else
                

                If Trim(cnpj) <> Trim(ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/emit/CNPJ").nodeTypedValue) Then
                    MsgBox "O fornecedor informado difere do fornecedor da nota!", vbCritical, "ATENÇÃO"
                    txtValor(18).Text = ""
                    txtFornecedor.SetFocus
                    Exit Sub
                End If
                txtValor(18).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/emit/CNPJ").nodeTypedValue
           End If
           If InStr(strTeste, "<xNome") <> 0 Then
                txtValor(19).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/emit/xNome").nodeTypedValue
                            
            End If
           
            If InStr(strTeste, "<xFant") <> 0 Then
                txtValor(20).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/emit/xFant").nodeTypedValue
            End If
           
            
          'ide--------------------------------------------------------------------------------------
          Dim qtdChild As Integer
          Dim no As String
          I = 0
            qtdChild = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/ide").childNodes.Length
          Do While I < qtdChild
            no = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/ide").childNodes(I).BaseName
            If no = "cUF" Then
                txtValor(0).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/ide/" & no).nodeTypedValue
            End If
             If no = "cNF" Then
                txtValor(1).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/ide/" & no).nodeTypedValue
            End If
             If no = "natOp" Then
                txtValor(2).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/ide/" & no).nodeTypedValue
            End If
             If no = "indPag" Then
                txtValor(3).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/ide/" & no).nodeTypedValue
            End If
             If no = "mod" Then
                txtValor(4).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/ide/" & no).nodeTypedValue
            End If
             If no = "serie" Then
                txtValor(5).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/ide/" & no).nodeTypedValue
            End If
             If no = "nNF" Then
                txtValor(6).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/ide/" & no).nodeTypedValue
            End If
             If no = "dEmi" Then
                txtValor(7).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/ide/" & no).nodeTypedValue
            End If
            If no = "dSaiEnt" Then
                txtValor(8).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/ide/" & no).nodeTypedValue
            End If
            If no = "tpNF" Then
                txtValor(9).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/ide/" & no).nodeTypedValue
            End If
            If no = "cMunFG" Then
                txtValor(10).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/ide/" & no).nodeTypedValue
            End If
            If no = "tpImp" Then
                txtValor(11).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/ide/" & no).nodeTypedValue
            End If
            If no = "tpEmis" Then
                txtValor(12).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/ide/" & no).nodeTypedValue
            End If
            If no = "cDV" Then
                txtValor(13).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/ide/" & no).nodeTypedValue
            End If
            If no = "tpAmb" Then
                txtValor(14).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/ide/" & no).nodeTypedValue
            End If
            If no = "finNFe" Then
                txtValor(15).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/ide/" & no).nodeTypedValue
            End If
            If no = "procEmi" Then
                txtValor(16).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/ide/" & no).nodeTypedValue
            End If
            If no = "verProc" Then
                txtValor(17).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/ide/" & no).nodeTypedValue
            End If
            I = I + 1
          Loop
           
            
            'emit/enderEmit--------------------------------------------------------------------------------------
             I = 0
            qtdChild = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/emit/enderEmit").childNodes.Length
          Do While I < qtdChild
            no = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/emit/enderEmit").childNodes(I).BaseName
           If no = "xLgr" Then
                txtValor(21).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/emit/enderEmit/" & no).nodeTypedValue
            End If
           If no = "nro" Then
                txtValor(22).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/emit/enderEmit/" & no).nodeTypedValue
           End If
           If no = "xCpl" Then
                txtValor(23).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/emit/enderEmit/" & no).nodeTypedValue
           End If
           If no = "xBairro" Then
                txtValor(24).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/emit/enderEmit/" & no).nodeTypedValue
           End If
           If no = "cMun" Then
                txtValor(25).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/emit/enderEmit/" & no).nodeTypedValue
            End If
          If no = "xMun" Then
                txtValor(26).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/emit/enderEmit/" & no).nodeTypedValue
          End If
           If no = "UF" Then
                txtValor(27).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/emit/enderEmit/" & no).nodeTypedValue
          End If
          If no = "CEP" Then
                txtValor(28).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/emit/enderEmit/" & no).nodeTypedValue
          End If
          If no = "cPais" Then
                txtValor(29).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/emit/enderEmit/" & no).nodeTypedValue
          End If
          If no = "xPais" Then
                txtValor(30).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/emit/enderEmit/" & no).nodeTypedValue
          End If
          If no = "fone" Then
                txtValor(31).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/emit/enderEmit/" & no).nodeTypedValue
          End If
          I = I + 1
          Loop
            
           
    'emit------------------------------------------------------------------------------------------
           Cont = 0
           qtdChild = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/emit").childNodes.Length
          Do While Cont < qtdChild
            no = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/emit").childNodes(Cont).BaseName
           If no = "IE" Then
                txtValor(32).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/emit/" & no).nodeTypedValue
            End If
           If no = "IM" Then
                txtValor(33).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/emit/" & no).nodeTypedValue
            End If
           If no = "CNAE" Then
                txtValor(34).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/emit/" & no).nodeTypedValue
            End If
           If no = "CRT" Then
                txtValor(71).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/emit/" & no).nodeTypedValue
            End If
         Cont = Cont + 1
         Loop
            
    'dest------------------------------------------------------------------------------------------
         I = 0
           qtdChild = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/dest").childNodes.Length
          Do While I < qtdChild
            no = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/dest").childNodes(I).BaseName
           If no = "CNPJ" Then
                txtValor(35).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/dest/" & no).nodeTypedValue
            End If
           If no = "xNome" Then
                txtValor(36).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/dest/" & no).nodeTypedValue
            End If
            
            If no = "IE" Then
                txtValor(47).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/dest/" & no).nodeTypedValue
            End If
          I = I + 1
          Loop
           'dest/enderDest------------------------------------------------------------------------------------------
        I = 0
           qtdChild = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/dest/enderDest").childNodes.Length
          Do While I < qtdChild
            no = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/dest/enderDest").childNodes(I).BaseName
            If no = "xLgr" Then
                txtValor(37).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/dest/enderDest/" & no).nodeTypedValue
            End If
            If no = "nro" Then
                txtValor(38).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/dest/enderDest/" & no).nodeTypedValue
            End If
            If no = "xBairro" Then
                txtValor(39).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/dest/enderDest/" & no).nodeTypedValue
            End If
            If no = "cMun" Then
                txtValor(40).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/dest/enderDest/" & no).nodeTypedValue
            End If
            If no = "xMun" Then
                txtValor(41).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/dest/enderDest/" & no).nodeTypedValue
            End If
            If no = "UF" Then
                txtValor(42).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/dest/enderDest/" & no).nodeTypedValue
            End If
            If no = "CEP" Then
                txtValor(43).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/dest/enderDest/" & no).nodeTypedValue
            End If
             If no = "cPais" Then
                txtValor(44).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/dest/enderDest/" & no).nodeTypedValue
             End If
                  If no = "xPais" Then
                txtValor(45).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/dest/enderDest/" & no).nodeTypedValue
            End If
            If no = "fone" Then
                txtValor(46).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/dest/enderDest/" & no).nodeTypedValue
            End If
            I = I + 1
            Loop
            
            
            
            'ver qtde de produtos-----
            
            qtde = 1
            Cont = 0
            Do While InStr(strTeste, "<det nItem=" & """" & qtde & """") <> 0
                Cont = Cont + 1
                qtde = qtde + 1
            Loop
          
           qtde = Cont
          'prod------------------------
          Cont = 0
          I = 0
          Dim contAux As Integer
          contAux = 0

          Do While Cont < qtde
          
          Do While contAux <= 29 'limpando o vetor
            vet(contAux) = ""
            contAux = contAux + 1
          Loop
          I = 0
           qtdChild = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/prod").childNodes.Length
          Do While I < qtdChild
          no = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/prod").childNodes(I).BaseName
          'grid Det
          If no = "cProd" Then
              vet(0) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/prod/" & no).nodeTypedValue
          End If
          If no = "xProd" Then
              vet(1) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/prod/" & no).nodeTypedValue
          End If
          If no = "NCM" Then
              vet(2) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/prod/" & no).nodeTypedValue
          End If
          If no = "CFOP" Then
              vet(3) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/prod/" & no).nodeTypedValue
          End If
          If no = "uCom" Then
              vet(4) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/prod/" & no).nodeTypedValue
          End If
          If no = "qCom" Then
              vet(5) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/prod/" & no).nodeTypedValue
          End If
          If no = "vUnCom" Then
              vet(6) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/prod/" & no).nodeTypedValue
          End If
          If no = "vProd" Then
              vet(7) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/prod/" & no).nodeTypedValue
          End If
          If no = "uTrib" Then
              vet(8) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/prod/" & no).nodeTypedValue
          End If
          If no = "qTrib" Then
              vet(9) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/prod/" & no).nodeTypedValue
          End If
          If no = "vUnTrib" Then
              vet(10) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/prod/" & no).nodeTypedValue
          End If
          If no = "indTot" Then
              vet(11) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/prod/" & no).nodeTypedValue
          End If
          If no = "cEAN" Then
              vet(32) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/prod/" & no).nodeTypedValue
          End If
        
                        
           I = I + 1
         Loop
         Cont = Cont + 1
         grdDet.AddItem vet(0) & Chr(9) & _
                        vet(1) & Chr(9) & _
                        vet(2) & Chr(9) & _
                        vet(3) & Chr(9) & _
                        vet(4) & Chr(9) & _
                        vet(5) & Chr(9) & _
                        vet(6) & Chr(9) & _
                        vet(7) & Chr(9) & _
                        vet(8) & Chr(9) & _
                        vet(9) & Chr(9) & _
                        vet(10) & Chr(9) & _
                        vet(11) & Chr(9) & _
                        vet(32)
        Loop
        
          
                
                
    'grid det impostos
          Cont = 0
          I = 0

          contAux = 0
          Dim verNo As String
          Do While Cont < qtde
          
          Do While contAux <= 29 'limpando o vetor
            vet(contAux) = ""
            contAux = contAux + 1
          Loop
          I = 0
          
          If InStr(strTeste, "ICMS10") <> 0 Then
             verNo = "ICMS10"
          ElseIf InStr(strTeste, "ICMS00") <> 0 Then
            verNo = "ICMS00"
          End If
          
          qtdChild = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/ICMS/" & verNo).childNodes.Length
          Do While I < qtdChild
          qtdChild = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/ICMS/" & verNo).childNodes.Length
          no = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/ICMS/" & verNo).childNodes(I).BaseName
          
          If no = "orig" Then
              vet(11) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/ICMS/" & verNo & "/" & no).nodeTypedValue
          End If
          
          If no = "CST" Then
              vet(12) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/ICMS/" & verNo & "/" & no).nodeTypedValue
          End If
          
          If no = "modBC" Then
              vet(13) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/ICMS/" & verNo & "/" & no).nodeTypedValue
          End If
          
          If no = "vBC" Then
              vet(14) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/ICMS/" & verNo & "/" & no).nodeTypedValue
          End If
          
          If no = "pICMS" Then
              vet(15) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/ICMS/" & verNo & "/" & no).nodeTypedValue
          End If
          
          If no = "vICMS" Then
              vet(16) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/ICMS/" & verNo & "/" & no).nodeTypedValue
          End If
          
          If no = "modBCST" Then
              vet(17) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/ICMS/" & verNo & "/" & no).nodeTypedValue
          End If
          If no = "pMVAST" Then
              vet(18) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/ICMS/" & verNo & "/" & no).nodeTypedValue
          End If
          If no = "vBCST" Then
              vet(19) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/ICMS/" & verNo & "/" & no).nodeTypedValue
          End If
          If no = "pICMSST" Then
              vet(20) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/ICMS/" & verNo & "/" & no).nodeTypedValue
          End If
          If no = "vICMSST" Then
              vet(21) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/ICMS/" & verNo & "/" & no).nodeTypedValue
          End If
          
          I = I + 1
         Loop
         
         'definindo a quantidade de childNodes do próximo nó
         I = 0
         qtdChild = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/PIS/PISAliq").childNodes.Length
         Do While I < qtdChild
          no = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/PIS/PISAliq").childNodes(I).BaseName
          
          If no = "CST" Then
              vet(22) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/PIS/PISAliq/" & no).nodeTypedValue
          End If
          If no = "vBC" Then
              vet(23) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/PIS/PISAliq/" & no).nodeTypedValue
          End If
           If no = "pPIS" Then
              vet(24) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/PIS/PISAliq/" & no).nodeTypedValue
          End If
          If no = "vPIS" Then
              vet(25) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/PIS/PISAliq/" & no).nodeTypedValue
          End If
          I = I + 1
          Loop
          
          I = 0
          qtdChild = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/COFINS/COFINSAliq").childNodes.Length
          Do While I < qtdChild
          no = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/COFINS/COFINSAliq").childNodes(I).BaseName
          
          If no = "CST" Then
              vet(26) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/COFINS/COFINSAliq/" & no).nodeTypedValue
          End If
          If no = "vBC" Then
              vet(27) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/COFINS/COFINSAliq/" & no).nodeTypedValue
          End If
          If no = "pCOFINS" Then
              vet(28) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/COFINS/COFINSAliq/" & no).nodeTypedValue
          End If
          If no = "vCOFINS" Then
              vet(29) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/COFINS/COFINSAliq/" & no).nodeTypedValue
          End If
          I = I + 1
          Loop
         
          If InStr(strTeste, "cEnq") <> 0 Then
              vet(30) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/IPI/cEnq").nodeTypedValue
          End If
          If InStr(strTeste, "IPINT") <> 0 And InStr(strTeste, "CST") <> 0 Then
              vet(31) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/IPI/IPINT/CST").nodeTypedValue
          End If
          If InStr(strTeste, "IPI") <> 0 And InStr(strTeste, "qSelo") <> 0 Then
              vet(32) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/IPI/qSelo").nodeTypedValue
          End If
          
          If InStr(strTeste, "IPITrib") <> 0 Then
              vet(33) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/IPI/IPITrib/CST").nodeTypedValue
          End If
          If InStr(strTeste, "IPITrib") <> 0 Then
              vet(34) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/IPI/IPITrib/vBC").nodeTypedValue
          End If
          If InStr(strTeste, "IPITrib") <> 0 And InStr(strTeste, "pIPI") <> 0 Then
              vet(35) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/IPI/IPITrib/pIPI").nodeTypedValue
          End If
          If InStr(strTeste, "IPITrib") <> 0 And InStr(strTeste, "vIPI") <> 0 Then
              vet(36) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/det[ " & Cont & "]/imposto/IPI/IPITrib/vIPI").nodeTypedValue
          End If
          
           
         Cont = Cont + 1
                
                        
               grdImposto.AddItem vet(11) & Chr(9) & _
                        vet(12) & Chr(9) & _
                        vet(13) & Chr(9) & _
                        vet(14) & Chr(9) & _
                        vet(15) & Chr(9) & _
                        vet(16) & Chr(9) & _
                        vet(17) & Chr(9) & _
                        vet(18) & Chr(9) & _
                        vet(19) & Chr(9) & _
                        vet(20) & Chr(9) & _
                        vet(21) & Chr(9) & _
                        vet(22) & Chr(9) & _
                        vet(23) & Chr(9) & _
                        vet(24) & Chr(9) & _
                        vet(25) & Chr(9) & _
                        vet(26) & Chr(9) & _
                        vet(27) & Chr(9) & _
                        vet(28) & Chr(9) & _
                        vet(29) & Chr(9) & _
                        vet(30) & Chr(9) & _
                        vet(31) & Chr(9) & _
                        vet(32) & Chr(9) & _
                        vet(33) & Chr(9) & _
                        vet(34) & Chr(9) & _
                        vet(35) & Chr(9) & vet(36)
                      
           
             Loop
                
        'Total -------------------------------------------------------------------------------------------------
            I = 0
           qtdChild = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/total/ICMSTot").childNodes.Length
          Do While I < qtdChild
            no = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/total/ICMSTot").childNodes(I).BaseName
                If no = "vBC" Then
                    txtValor(48).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/total/ICMSTot/" & no).nodeTypedValue
                End If
                If no = "vICMS" Then
                    txtValor(49).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/total/ICMSTot/" & no).nodeTypedValue
                End If
                If no = "vBCST" Then
                    txtValor(50).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/total/ICMSTot/" & no).nodeTypedValue
                End If
                If no = "vST" Then
                    txtValor(51).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/total/ICMSTot/" & no).nodeTypedValue
                End If
                If no = "vProd" Then
                    txtValor(52).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/total/ICMSTot/" & no).nodeTypedValue
                End If
                If no = "vFrete" Then
                    txtValor(53).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/total/ICMSTot/" & no).nodeTypedValue
                End If
                If no = "vSeg" Then
                    txtValor(54).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/total/ICMSTot/" & no).nodeTypedValue
                End If
                If no = "vDesc" Then
                    txtValor(55).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/total/ICMSTot/" & no).nodeTypedValue
                End If
                If no = "vII" Then
                    txtValor(56).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/total/ICMSTot/" & no).nodeTypedValue
                End If
                If no = "vIPI" Then
                    txtValor(57).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/total/ICMSTot/" & no).nodeTypedValue
                End If
                If no = "vPIS" Then
                    txtValor(58).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/total/ICMSTot/" & no).nodeTypedValue
                End If
                If no = "vCOFINS" Then
                    txtValor(59).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/total/ICMSTot/" & no).nodeTypedValue
                End If
                If no = "vOutro" Then
                    txtValor(60).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/total/ICMSTot/" & no).nodeTypedValue
                End If
                If no = "vNF" Then
                    txtValor(61).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/total/ICMSTot/" & no).nodeTypedValue
                End If
             I = I + 1
             Loop
                'Transp------------------------------------------------------------------------------------------
         
                If InStr(strTeste, "<modFrete") <> 0 Then
                    txtValor(62).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/transp/modFrete").nodeTypedValue
                End If
                If InStr(strTeste, "<qVol") <> 0 Then
                    txtValor(63).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/transp/vol/qVol").nodeTypedValue
                End If
                If InStr(strTeste, "<esp") <> 0 Then
                    txtValor(64).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/transp/vol/esp").nodeTypedValue
                End If
                If InStr(strTeste, "<marca") <> 0 Then
                    txtValor(65).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/transp/vol/marca").nodeTypedValue
                End If
                If InStr(strTeste, "<nVol") <> 0 Then
                    txtValor(66).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/transp/vol/nVol").nodeTypedValue
                End If
                If InStr(strTeste, "<pesoL") <> 0 Then
                    txtValor(67).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/transp/vol/pesoL").nodeTypedValue
                End If
                If InStr(strTeste, "<pesoB") <> 0 Then
                    txtValor(68).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/transp/vol/pesoB").nodeTypedValue
                End If
                
        'Cobr------------------------------------------------------------------------------------------
        contAux = 0
        Do While contAux <= 29 'limpando o vetor
            vet(contAux) = ""
            contAux = contAux + 1
          Loop
          
          Cont = 1
          contAux = 0
         Do While InStr(Cont, strTeste, "<dup") <> 0
             Cont = InStr(Cont, strTeste, "<dup") + 1
             contAux = contAux + 1
         Loop
                
       
        Cont = 0
        Do While Cont < contAux
                
                 I = 0
           qtdChild = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/cobr/dup[ " & Cont & "]").childNodes.Length
          Do While I < qtdChild
            no = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/cobr/dup[ " & Cont & "]").childNodes(I).BaseName
            If no = "nDup" Then
              vet(0) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/cobr/dup[ " & Cont & "]/" & no).nodeTypedValue
            End If
            If no = "dVenc" Then
              vet(1) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/cobr/dup[ " & Cont & "]/" & no).nodeTypedValue
            End If
            If no = "vDup" Then
              vet(2) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/cobr/dup[ " & Cont & "]/" & no).nodeTypedValue
            End If
            I = I + 1
         Loop
         
         grdCobr.AddItem vet(0) & Chr(9) & _
                        vet(1) & Chr(9) & _
                        vet(2)
                        
                
         Cont = Cont + 1
         Loop
         
         
        
        
          If InStr(strTeste, "<fat") <> 0 Then
          
          Cont = 1
          contAux = 0
         Do While InStr(Cont, strTeste, "<fat") <> 0
             Cont = InStr(Cont, strTeste, "<fat") + 1
             contAux = contAux + 1
         Loop
         
         
            Cont = 0
        Do While Cont < contAux
            I = 0
           qtdChild = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/cobr/fat[ " & Cont & "]").childNodes.Length
          Do While I < qtdChild
            no = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/cobr/fat[ " & Cont & "]").childNodes(I).BaseName
            
            If no = "nFat" Then
              vet(3) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/cobr/fat[ " & Cont & "]/" & no).nodeTypedValue
            End If
         
            If no = "vOrig" Then
              vet(4) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/cobr/fat[ " & Cont & "]/" & no).nodeTypedValue
            End If
            If no = "vLiq" Then
              vet(5) = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/cobr/fat[ " & Cont & "]/" & no).nodeTypedValue
            End If
          I = I + 1
          Loop
          grdCobrFat.AddItem vet(3) & Chr(9) & _
                        vet(4) & Chr(9) & _
                        vet(5)
                        
                
         Cont = Cont + 1
         Loop
          End If
        
         
        'infAdic
            If InStr(strTeste, "<infCpl") <> 0 Then
                    txtValor(69).Text = ndeAux.selectSingleNode("/nfeProc/NFe/infNFe/infAdic/infCpl").nodeTypedValue
            End If
            
         'txtValor(70).Text = Mid(strTeste, (InStr(strTeste, "Id=") + 4), ((InStr(120, strTeste, "versao=") - 2) - (InStr(strTeste, "Id=") + 4)))
            
       txtValor(70).Text = ndeAux.selectSingleNode("/nfeProc/protNFe/infProt/chNFe").nodeTypedValue
            
        Next
          Call carregaBanco
        Else
            MsgBox "XML não encontrado!", vbCritical, "ATENÇÃO"
        End If
        
   
End Sub

Private Sub cmdXML_Click()
    frmXml.Show 1
End Sub



Public Sub inibeFrame()
   fraIde.Visible = False
    fraEmit.Visible = False
    fraEnderEmit.Visible = False
    fraDest.Visible = False
    fraEnderDest.Visible = False
    fraDet.Visible = False
    fraDetImp.Visible = False
    fraTotal.Visible = False
    fraTransp.Visible = False
    fraCobr.Visible = False
    Label2.Visible = False
    txtValor(70).Visible = False
    grdInfAd.Visible = False
    txtValor(69).Visible = False
    
    frmConverterXML.WindowState = 0
   ' frmConverterXML.Width = 7740
  '  frmConverterXML.Height = 4410
    
     frmConverterXML.Width = 11600
    frmConverterXML.Height = 4850
    
    
    fraCaminhoXML.Width = 11400
    Me.Picture1.Width = 11450
    Picture1.top = 4000
    
    Me.cmdCarregaXML.left = 7235
    cmdCarregaXML.top = 4200
    
     cmdXML.left = 8660
     cmdXML.top = 4200
    
      Me.cmdRetornar.left = 10080
    cmdRetornar.top = 4200
    
  '  left = (Screen.Width - Width) / 2
  '  top = (Screen.Height - Height) / 2
  
   frmConverterXML.top = 5700
    frmConverterXML.left = 90
    
    fraCaminhoXML.top = 900
End Sub
Public Sub mostraFrame()
   fraIde.Visible = True
    fraEmit.Visible = True
    fraEnderEmit.Visible = True
    fraDest.Visible = True
    fraEnderDest.Visible = True
    fraDet.Visible = True
    fraDetImp.Visible = True
    fraTotal.Visible = True
    fraTransp.Visible = True
    fraCobr.Visible = True
    Label2.Visible = True
    txtValor(70).Visible = True
    grdInfAd.Visible = True
    txtValor(69).Visible = True
    
    
    Me.Picture1.Width = 15240
    Picture1.top = 10335
    Me.fraCaminhoXML.Width = 15165
    
    Me.cmdRetornar.left = 13890
    cmdRetornar.top = 10560
    
     cmdXML.left = 12500
     cmdXML.top = 10560
    
    cmdCarregaXML.left = 11120
    cmdCarregaXML.top = 10560
    
    frmConverterXML.WindowState = 2
    
    fraCaminhoXML.top = 510
End Sub

Private Sub limpaTela()
   ' Text1.Text = ""
    Cont = 0
    Do While Cont <= 70
        txtValor(Cont) = ""
        Cont = Cont + 1
    Loop
    Me.grdDet.Rows = 1
    Me.grdImposto.Rows = 1
    Me.grdCobr.Rows = 1
    Me.grdCobrFat.Rows = 1
    Text1.SetFocus
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
    telaChamou = ""
    Call inibeFrame
    
    Call carregaTipoEntrada
    
    If telaChamou = "frmCadastraCodProdutoNoFornecedor" Then
        Call carregaBanco
    End If
    
End Sub


Private Sub txtDataRecebimento_Change()
     'Colocando em formato de data no msk
     If Len(txtDataRecebimento.Text) = 2 Then
       txtDataRecebimento.Text = txtDataRecebimento.Text & "/"
       txtDataRecebimento.SelStart = 3
    ElseIf Len(txtDataRecebimento.Text) = 5 Then
       txtDataRecebimento.Text = txtDataRecebimento.Text & "/"
       txtDataRecebimento.SelStart = 6
    ElseIf Len(txtDataRecebimento.Text) = 10 Then
        'mskDataTermino.SetFocus
    End If
End Sub

Private Sub txtFornecedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
      If txtFornecedor.Text <> "" And IsNumeric(txtFornecedor.Text) = True Then
       Call buscaFornecedor
      End If
    End If
End Sub

Private Sub carregaBanco()
    SQL = "Select * from capanfcompra where cc_notafiscal = '" & txtValor(6).Text & _
          "' and cc_serie = 'NE' and cc_loja = 'CD'"
    adoXML.CursorLocation = adUseClient
    adoXML.Open SQL, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
        
    If Not adoXML.EOF Then
        MsgBox "Nota fiscal já cadastrada!", vbCritical, "ATENÇÃO"
        adoXML.Close
        Call limpaTela
        Exit Sub
    End If
    
    adoXML.Close
    
    Call verificaCodigoProdutoFornecedor
    If verReferencia = True Then
    
    ADO_Cn_CD.BeginTrans
        SQL = "Insert into capanfcompra (cc_notafiscal, cc_serie, cc_fornecedor, cc_loja, cc_codigooperacao, " _
                & " cc_naturezaoperacao, cc_condicaopagamento, cc_dataemissao, cc_datarecebimento, cc_dataentrada, " _
                & " cc_valormercadorias, cc_embalagem, cc_desconto, cc_frete, cc_outros, cc_despesas, cc_juros, " _
                & " cc_aliquotaicms, cc_baseicms, cc_valoricms, cc_valoricmssubstrib, cc_valoripi, cc_valortotalnota, " _
                & " cc_valorcalculado, cc_situacao, cc_notaconsignacao, cc_serieconsignacao, cc_critica, " _
                & " cc_observacao, cc_verbapropaganda, cc_verbapaga, cc_tipofrete, cc_tipoentrada, cc_acaoentrada, " _
                & " cc_cfopnf, cc_cfopentrada, cc_categoriapedido) " _
                & " values ('" & txtValor(6).Text & "', 'NE', " & Mid(txtFornecedor.Text, 1, 4) & ", 'CD', " _
                & " 0, 0, 0, '" & txtValor(7).Text & "', '" & Format(txtDataRecebimento.Text, "yyyy/mm/dd") & "', " _
                & " '" & Format(Date, "yyyy/mm/dd") & "', '" & txtValor(52).Text & "', 0, '" & txtValor(55).Text & _
                "', '" & txtValor(53).Text & "', '" & txtValor(60).Text & "', 0, 0, 0, '" & txtValor(48).Text & _
                "', '" & txtValor(49).Text & "', '" & txtValor(51).Text & " ', '" & txtValor(57).Text & _
                "', '" & txtValor(61).Text & "', 0, 'D', 0, 0, ' ', ' ', 0, 0, 0, '" & Mid(cmbTipoEntrada.Text, 1, 3) & "', 0, 0, 0, 0)"
 
    ADO_Cn_CD.Execute (SQL)
    ADO_Cn_CD.CommitTrans
   
   Cont = 1
   Do While Cont <= grdDet.Rows - 1
  
   Valor = ConvertePonto(grdImposto.TextMatrix(Cont, 14))
   valor2 = ConvertePonto(grdImposto.TextMatrix(Cont, 18))
   Valor = Valor + valor2

   ADO_Cn_CD.BeginTrans
  
   SQL = "Insert into itemnfcompra (ci_notafiscal, ci_serie, ci_fornecedor, ci_dataentrada, ci_item, ci_referencia, " _
        & " ci_quantidade, ci_precounitario, ci_percentualdesconto, ci_dataliberacao, ci_datatransferencia, " _
        & " ci_nossopedido, ci_custoliquido, ci_custoanterior, ci_novocusto, ci_precovendadata, ci_novoprecovenda, " _
        & " ci_aliquotaipi, ci_estoqueanterior, ci_situacao, ci_observacao, ci_horamanutencao, ci_tipoatualizacompras, " _
        & " ci_aliqicms, ci_valoricms, ci_valoripi, ci_customedioliquido, ci_reducaoicms, ci_valorpiscofins, " _
        & " ci_valorestoqueanterior, ci_valorestoqueentrada, ci_estoqueconso, ci_loja, ci_preconossopedido, " _
        & " ci_criticapreco, ci_calculacritica, ci_cfop, ci_cfopentrada, ci_st, ci_filialmu, ci_isentoicms, " _
        & " ci_outroicms, ci_baseicms, ci_produtofornecedor, ci_descricaofornecedor, ci_codigobarra, " _
        & " ci_classificacaofiscal) values ('" & txtValor(6).Text & "', 'NE', " & Mid(txtFornecedor.Text, 1, 4) & "," _
        & " '" & txtValor(7).Text & "', '" & grdDet.Rows - 1 & "', '" & vet(Cont) & "', " _
        & " '" & Mid(grdDet.TextMatrix(Cont, 5), 1, 1) & "', '" & grdDet.TextMatrix(Cont, 7) & "', 0, 0, 0, 0, 0, " _
        & " 0, 0, 0, 0, '" & grdImposto.TextMatrix(Cont, 24) & "', 0, 'D', ' ', 0, 0, " _
        & " '" & grdImposto.TextMatrix(Cont, 4) & "', '" & grdImposto.TextMatrix(Cont, 5) & _
        "', '" & grdImposto.TextMatrix(Cont, 25) & "', 0, '" & Mid(grdImposto.TextMatrix(Cont, 1), 1, 1) & _
        "', '" & ConverteVirgula(Valor) & "', 0, 0, 0, 'CD', 0, ' ', ' ', '" & grdDet.TextMatrix(Cont, 3) & "', " _
        & " '" & grdDet.TextMatrix(Cont, 3) & _
        "', ' ', ' ', 0, 0, 0, " _
        & " '" & grdDet.TextMatrix(Cont, 0) & "', '" & Format(grdDet.TextMatrix(Cont, 1), "0") & _
        "', '" & grdDet.TextMatrix(Cont, 12) & "', 0)"
    
   ADO_Cn_CD.Execute (SQL)
   ADO_Cn_CD.CommitTrans
   Cont = Cont + 1
   Loop
   
   MsgBox "XML carregado com sucesso!", vbInformation, "XML"
   Call carregaFrmEntradaNfCompras
   Call limpaTela
     ' telaChamou = "frmConverterXML"
  ' frmEntradaNFCompras.Show
   Unload Me
   Else
   Unload frmConverterXML
   Unload frmListarArquivos
    frmCadastraCodProdutoNoFornecedor.Show
    
   End If
   
End Sub

Private Sub verificaCodigoProdutoFornecedor()
    
    Cont = 0
    verReferencia = True
    
    'LIMPANDO O VETOR PARA ARMAZENAR AS REFERENCIAS
    Do While Cont < 70
        vet(Cont) = ""
        Cont = Cont + 1
    Loop
    
    'VERIFICANDO SE O CÓDIGO DO PRODUTO DO FORNECEDOR ESTÁ CADASTRADO
    Cont = 1
    frmCadastraCodProdutoNoFornecedor.grdProduto.Rows = 1
    Do While Cont <= grdDet.Rows - 1

        SQL = "select pi_referencia, pr_codigoprodutonofornecedor, pr_descricao from itempedido, produto " _
            & " where pr_referencia = pi_referencia and " _
            & " pr_codigoprodutonofornecedor = '" & Format(grdDet.TextMatrix(Cont, 0), "0") & "'"
            
        adoXML.CursorLocation = adUseClient
        adoXML.Open SQL, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
        
        If adoXML.EOF Then
            verReferencia = False
            frmCadastraCodProdutoNoFornecedor.grdProduto.AddItem grdDet.TextMatrix(Cont, 1) & Chr(9) & _
                                                                Format(grdDet.TextMatrix(Cont, 0), "0")
            
        Else
            vet(Cont) = Trim(adoXML("pi_referencia"))
        End If
        adoXML.Close
        Cont = Cont + 1
    Loop
    
End Sub

Private Sub carregaFrmEntradaNfCompras()
    valorMerc = 0
    totalCalculado = 0
    frmEntradaNFCompras.txtLoja.Text = "CD"
    frmEntradaNFCompras.txtCnpj.Text = txtValor(18).Text
    frmEntradaNFCompras.txtFornecedor.Text = Mid(txtFornecedor.Text, 7)
    frmEntradaNFCompras.txtNf.Text = txtValor(6).Text
    frmEntradaNFCompras.txtSerie.Text = "NE"
    frmEntradaNFCompras.txtDataEmissao.Text = Format(txtValor(7).Text, "dd/mm/yyyy")
    frmEntradaNFCompras.txtDataEntrada.Text = Format(Date, "dd/mm/yyyy")
    totalNota = Format(ConvertePonto(txtValor(61).Text), "###,###,##0.00")
    frmEntradaNFCompras.txtTotalNotaFiscal.Text = Format(totalNota, "###,###,##0.00")
    frmEntradaNFCompras.txtDataRecebimento.Text = Format(txtDataRecebimento.Text, "dd/mm/yyyy")
    frmEntradaNFCompras.txtTipoEntrada.Text = cmbTipoEntrada.Text
    frmEntradaNFCompras.txtCfop.Text = grdDet.TextMatrix(1, 3)
    SQL = "Select * from cfopentradasaida where cfo_codigo = '" & Trim(grdDet.TextMatrix(1, 3)) & "'"
        adoXML.CursorLocation = adUseClient
        adoXML.Open SQL, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    
    If Not adoXML.EOF Then
        frmEntradaNFCompras.txtDescricaoOperacao.Text = adoXML("cfo_descricaooperacao")
    End If
    adoXML.Close
        
    Cont = 1
    
   

    Do While Cont <= grdDet.Rows - 1
        SQL = "Select pr_referencia, pr_descricao, chm_quantidade from produto, chekinmercadoria " _
            & " where pr_codigoprodutonofornecedor = '" & Format(grdDet.TextMatrix(Cont, 0), "0") & _
            "' and pr_referencia = chm_referencia and chm_notafiscal = '" & txtValor(6).Text & _
            "' and chm_serie = 'NE' and chm_loja = 'CD' and chm_fornecedor = " & Mid(txtFornecedor.Text, 1, 4) & ""
        adoXML.CursorLocation = adUseClient
        adoXML.Open SQL, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
        If Not adoXML.EOF Then
           Do While Not adoXML.EOF
            frmEntradaNFCompras.grdProduto.AddItem adoXML("pr_referencia") & Chr(9) & _
                            adoXML("pr_descricao") & Chr(9) & _
                            grdDet.TextMatrix(Cont, 2) & Chr(9) & _
                            Mid(grdImposto.TextMatrix(Cont, 1), 1, 1) & Chr(9) & _
                            grdDet.TextMatrix(Cont, 3) & Chr(9) & _
                            "PC" & Chr(9) & _
                            Mid(grdDet.TextMatrix(Cont, 5), 1, InStr(1, grdDet.TextMatrix(Cont, 5), ".") - 1) & Chr(9) & _
                            adoXML("chm_quantidade") & Chr(9) & _
                            Format(ConvertePonto(grdDet.TextMatrix(Cont, 6)), "###,###,##0.00") & Chr(9) & _
                            Format(ConvertePonto(grdDet.TextMatrix(Cont, 7)), "###,###,##0.00")
            valorMerc = valorMerc + Format(ConvertePonto(grdDet.TextMatrix(Cont, 6)), "###,###,##0.000") * frmEntradaNFCompras.grdProduto.TextMatrix(Cont, 6)
            
            If frmEntradaNFCompras.grdProduto.TextMatrix(frmEntradaNFCompras.grdProduto.Rows - 1, 6) <> _
               frmEntradaNFCompras.grdProduto.TextMatrix(frmEntradaNFCompras.grdProduto.Rows - 1, 7) Then
               
               frmEntradaNFCompras.grdProduto.FillStyle = flexFillRepeat
               frmEntradaNFCompras.grdProduto.Col = 0
               frmEntradaNFCompras.grdProduto.Row = frmEntradaNFCompras.grdProduto.Rows - 1
               frmEntradaNFCompras.grdProduto.ColSel = frmEntradaNFCompras.grdProduto.Cols - 1
                'grdDados.CellBackColor = &H3333FF
                frmEntradaNFCompras.grdProduto.CellBackColor = &H703BA
                frmEntradaNFCompras.grdProduto.ColSel = 0
                frmEntradaNFCompras.grdProduto.FillStyle = flexFillSingle
            End If
            adoXML.MoveNext
            Loop
        End If
        adoXML.Close
    Cont = Cont + 1
    Loop
    
    
    
    vlrMercadoria = Format(valorMerc, "###,###,##0.00")
    vlrIcms = Format(ConvertePonto(txtValor(49).Text), "###,###,##0.00")
    vlrIPI = Format(ConvertePonto(txtValor(57).Text), "###,###,##0.00")
    vlrDespesa = "0,00"
    vlrJuros = "0,00"
    outrasDespesas = Format(ConvertePonto(txtValor(60).Text), "###,###,##0.00")
    embalagem = "0,00"
    frete = Format(ConvertePonto(txtValor(53).Text), "###,###,##0.00")
    vlrICMSST = Format(ConvertePonto(txtValor(51).Text), "###,###,##0.00")
    
    frmEntradaNFCompras.txtValorMercadoria.Text = Format(vlrMercadoria, "###,###,##0.00")
    frmEntradaNFCompras.txtValorICMS.Text = Format(vlrIcms, "###,###,##0.00")
    frmEntradaNFCompras.txtBaseICMSST.Text = Format(ConvertePonto(txtValor(50).Text), "###,###,##0.00")
    frmEntradaNFCompras.txtValorICMSST.Text = Format(vlrICMSST, "###,###,##0.00")
    frmEntradaNFCompras.txtValorFrete.Text = Format(frete, "###,###,##0.00")
    frmEntradaNFCompras.txtValorIPI.Text = Format(vlrIPI, "###,###,##0.00")
    frmEntradaNFCompras.txtBaseICMS.Text = Format(ConvertePonto(txtValor(48).Text), "###,###,##0.00")
    frmEntradaNFCompras.txtValorEmbalagem.Text = Format(embalagem, "###,###,##0.00")
    frmEntradaNFCompras.txtValorDespesa.Text = Format(vlrDespesa, "###,###,##0.00")
    frmEntradaNFCompras.txtValorJuros.Text = Format(vlrJuros, "###,###,##0.00")
    frmEntradaNFCompras.txtOutrasDespesas.Text = Format(outrasDespesas, "###,###,##0.00")
    
    totalCalculado = vlrMercadoria + vlrICMSST + vlrIPI + vlrDespesa + vlrJuros + outrasDespesas + embalagem + frete
    
    frmEntradaNFCompras.txtTotalCalculado.Text = Format(totalCalculado, "###,###,##0.00")

    frmEntradaNFCompras.txtBateNota.Text = Format(totalNota - totalCalculado, "###,###,##0.00")
    
   
    
    'Carrega grid de pedidos em aberto
    Cont = 1
    Dim qtdeBaixar As Integer
    Dim qtdeTabela As Integer
    Dim atingiuQtde As Boolean
    atingiuQtde = False
    Do While Cont <= frmEntradaNFCompras.grdProduto.Rows - 1
        pedidoAberto frmEntradaNFCompras.grdProduto.TextMatrix(Cont, 0)
        qtdeBaixar = frmEntradaNFCompras.grdProduto.TextMatrix(Cont, 6)
        If Not adoXML.EOF Then
            Do While Not adoXML.EOF
                
                    qtdeBaixar = qtdeBaixar - adoXML("pi_saldopedido")
                    If qtdeBaixar > 0 Then
                        qtdeTabela = adoXML("pi_saldopedido")
                        atingiuQtde = False
                    ElseIf qtdeBaixar < 0 Then
                        qtdeTabela = adoXML("pi_saldopedido") + qtdeBaixar
                        atingiuQtde = False
                    Else
                        atingiuQtde = True
                    End If
                    
                    If atingiuQtde = False Then
                    ADO_Cn_CD.BeginTrans
                      SQL = "Insert into nfcpc (NCP_Fornecedor, NCP_NotaFiscal, NCP_Serie, NCP_Referencia, NCP_Item, " _
                       & " NCP_Qtde, NCP_NroPedido) values ('" & Mid(txtFornecedor.Text, 1, 4) & _
                       "', '" & txtValor(6).Text & "', 'NE', '" & frmEntradaNFCompras.grdProduto.TextMatrix(Cont, 0) & _
                       "', '" & Cont & "', '" & qtdeTabela & "', '" & adoXML("pc_numeropedido") & "')"
             
                    ADO_Cn_CD.Execute (SQL)
                    ADO_Cn_CD.CommitTrans
                    End If
                    
                
                frmEntradaNFCompras.grdPedidoAberto.AddItem adoXML("pc_numeropedido") & Chr(9) & _
                                                            adoXML("pc_dataentrega") & Chr(9) & _
                                                            adoXML("pi_saldopedido")
                                                            
                
                adoXML.MoveNext
            Loop
            ADO_Cn_CD.BeginTrans
            If qtdeBaixar > 0 Then
              SQL = "Update nfcpc set ncp_situacao = 'C' where ncp_notafiscal = '" & txtValor(6).Text & _
               "' and ncp_serie = 'ne' and ncp_referencia = '" & frmEntradaNFCompras.grdProduto.TextMatrix(Cont, 0) & "'"
             Else
              SQL = "Update nfcpc set ncp_situacao = 'O' where ncp_notafiscal = '" & txtValor(6).Text & _
               "' and ncp_serie = 'ne' and ncp_referencia = '" & frmEntradaNFCompras.grdProduto.TextMatrix(Cont, 0) & "'"
            End If
            ADO_Cn_CD.Execute (SQL)
            ADO_Cn_CD.CommitTrans
        End If
        adoXML.Close
        Cont = Cont + 1
    Loop
End Sub

Private Sub carregaTipoEntrada()
    SQL = "Select * from tipopedidocompra"
        adoXML.CursorLocation = adUseClient
        adoXML.Open SQL, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
        
    If Not adoXML.EOF Then
        Do While Not adoXML.EOF
            cmbTipoEntrada.AddItem Trim(Format(adoXML("tpc_codigo"), "00") & " - " & adoXML("tpc_descricao"))
            adoXML.MoveNext
        Loop
    End If
    
    cmbTipoEntrada.ListIndex = 0
    
    adoXML.Close
End Sub

Private Sub txtFornecedor_LostFocus()
    If txtFornecedor.Text <> "" And IsNumeric(txtFornecedor.Text) = True Then
        Call buscaFornecedor
    End If
    If IsNumeric(Mid(txtFornecedor.Text, 1, 4)) = False And txtFornecedor.Text <> "" Then
        MsgBox "Fornecedor não cadastrado!", vbCritical, "ATENÇÃO"
        txtFornecedor.SetFocus
    End If
End Sub

Private Sub buscaFornecedor()
    SQL = "Select * from fornecedor where fo_codigofornecedor = '" & txtFornecedor.Text & "'"
        adoXML.CursorLocation = adUseClient
        adoXML.Open SQL, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
        
        If Not adoXML.EOF Then
            txtFornecedor.Text = Format(txtFornecedor.Text, "0000") & " - " & adoXML("fo_razaosocial")
            cnpj = Len(adoXML("fo_cgc"))
            If cnpj = 15 Then
                cnpj = Mid(adoXML("fo_cgc"), 2)
            Else
                cnpj = adoXML("fo_cgc")
            End If
        Else
            MsgBox "Fornecedor não cadastrado!", vbCritical, "ATENÇÃO"
            txtFornecedor.Text = ""
            txtFornecedor.SetFocus
        End If
        adoXML.Close
End Sub

Private Sub VSFlexGrid6_Click()

End Sub
