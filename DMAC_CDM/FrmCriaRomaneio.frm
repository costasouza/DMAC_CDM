VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7d.ocx"
Begin VB.Form FrmCriaRomaneio 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Romaneio"
   ClientHeight    =   8040
   ClientLeft      =   3825
   ClientTop       =   1605
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8040
   ScaleMode       =   0  'User
   ScaleWidth      =   15225
   ShowInTaskbar   =   0   'False
   Begin CentroDeDistribuicao.chameleonButton cmdTransfMerc 
      Height          =   510
      Index           =   1
      Left            =   2850
      TabIndex        =   28
      Top             =   6975
      Visible         =   0   'False
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
      MICON           =   "FrmCriaRomaneio.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   6585
      TabIndex        =   20
      Top             =   150
      Width           =   8445
      Begin VB.CheckBox chkDeletatodos 
         BackColor       =   &H00404040&
         Caption         =   "Deletar todas as  Referencias do Romaneio"
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
         Height          =   255
         Left            =   2280
         MaskColor       =   &H00FF0000&
         TabIndex        =   27
         Top             =   480
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.TextBox txtRefereForne 
         BackColor       =   &H00A3A3A3&
         Height          =   330
         Left            =   120
         MaxLength       =   7
         TabIndex        =   22
         Top             =   435
         Visible         =   0   'False
         Width           =   2010
      End
      Begin VB.TextBox TxtFantasia 
         BackColor       =   &H00A3A3A3&
         Height          =   330
         Left            =   2235
         TabIndex        =   21
         Top             =   435
         Visible         =   0   'False
         Width           =   4125
      End
      Begin VB.Label lblrealestoque 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
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
         Left            =   6900
         TabIndex        =   23
         Top             =   405
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label LblRefereForne 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Referência/Fornecedor"
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
         Left            =   120
         TabIndex        =   26
         Top             =   135
         Visible         =   0   'False
         Width           =   2400
      End
      Begin VB.Label LblNomeFantasia 
         BackColor       =   &H00404040&
         Caption         =   "Nome Fantasia"
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
         Height          =   240
         Left            =   2220
         TabIndex        =   25
         Top             =   135
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.Label lblEstoquecd 
         BackColor       =   &H00404040&
         Caption         =   "Estoque Origem"
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
         Height          =   375
         Left            =   6600
         TabIndex        =   24
         Top             =   135
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.CheckBox chkTodosRomaneios 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1635
      TabIndex        =   9
      Top             =   1290
      Width           =   195
   End
   Begin VB.CheckBox chkTodasLojas 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   195
      TabIndex        =   8
      Top             =   1275
      Width           =   210
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   44
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   14880
      TabIndex        =   5
      Top             =   6810
      Width           =   14880
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdRomaneio 
      Height          =   5280
      Left            =   4665
      TabIndex        =   16
      Top             =   1245
      Width           =   10365
      _cx             =   18283
      _cy             =   9313
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
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   9
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmCriaRomaneio.frx":001C
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
      BackColorFrozen =   5263440
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      Begin VB.Frame frmeImprimir 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "frmeImprimir"
         Height          =   1365
         Left            =   2760
         TabIndex        =   17
         Top             =   1200
         Visible         =   0   'False
         Width           =   2775
         Begin VB.CheckBox chkimpeSemRomaerio 
            BackColor       =   &H00404040&
            Caption         =   "Romaneio Separação"
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
            Height          =   375
            Left            =   360
            TabIndex        =   19
            Top             =   720
            Width           =   2175
         End
         Begin VB.CheckBox chkImpeRomaneio 
            BackColor       =   &H00404040&
            Caption         =   "Romaneio"
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
            Height          =   375
            Left            =   360
            TabIndex        =   18
            Top             =   240
            Width           =   1935
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   6270
      Begin VB.OptionButton optImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Impressão Romaneio"
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
         Height          =   435
         Left            =   3360
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Cancelar Romaneio"
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
         Height          =   435
         Left            =   4815
         TabIndex        =   14
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optManutencao 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Manutenção Romaneio"
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
         Height          =   495
         Left            =   1700
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton opiCriar 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Criar Romaneio"
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
         Height          =   435
         Left            =   150
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin CentroDeDistribuicao.chameleonButton cmdImprimir 
      Height          =   510
      Index           =   0
      Left            =   8940
      TabIndex        =   2
      Top             =   6945
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
      MICON           =   "FrmCriaRomaneio.frx":0189
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdCriarRomaneio 
      Height          =   510
      Index           =   1
      Left            =   7380
      TabIndex        =   3
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Criar Romaneio"
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
      MICON           =   "FrmCriaRomaneio.frx":01A5
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
      Index           =   2
      Left            =   10500
      TabIndex        =   1
      Top             =   6945
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
      MICON           =   "FrmCriaRomaneio.frx":01C1
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
      Left            =   13620
      TabIndex        =   4
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
      MICON           =   "FrmCriaRomaneio.frx":01DD
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
      Index           =   1
      Left            =   12060
      TabIndex        =   6
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
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
      MICON           =   "FrmCriaRomaneio.frx":01F9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdImprimiRomaneio 
      Height          =   510
      Index           =   0
      Left            =   5835
      TabIndex        =   7
      Top             =   6945
      Visible         =   0   'False
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Imprimir Romaneio"
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
      MICON           =   "FrmCriaRomaneio.frx":0215
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   -1  'True
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grd2 
      Height          =   5280
      Left            =   1590
      TabIndex        =   10
      Top             =   1245
      Width           =   2910
      _cx             =   5133
      _cy             =   9313
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
      FormatString    =   $"FrmCriaRomaneio.frx":0231
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
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdLojas 
      Height          =   5280
      Left            =   150
      TabIndex        =   11
      Top             =   1230
      Width           =   1275
      _cx             =   2249
      _cy             =   9313
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
      FormatString    =   $"FrmCriaRomaneio.frx":02AE
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
   End
End
Attribute VB_Name = "FrmCriaRomaneio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim sql As String
Dim Largura As String
Dim NomeTxt As String
Dim lblLagura As String
Dim vWhere As String
Dim vWhere2 As String
Dim vwhere3 As String
Dim vWhere4 As String
Dim qtdantiga As Integer
Dim wlojaant As String
Dim wNumeroRomaneio As String
Dim whereLoja As String
Dim whereRomaneio As String
Dim whereLoja2 As String
Dim orderByPadrao As String
Dim novaquantidade As String
Dim contchk As Integer
Dim ManutencaoRomaneio As Boolean

Private Sub CmbLoja_LostFocus()

'    If CmbLoja = "99-Todas" Then
'        txtRefereForne.Enabled = False
'        TxtFantasia.Enabled = False
'    End If
   
   '-----LimpaGrid e values
    limpaGrid grd2
    chkTodasLojas.Value = 0
    chkTodosRomaneios.Value = 0
    limpaGrid grdLojas
    CarregaLojas
      
    cmdCriarRomaneio(1).Enabled = True
    cmdImprimir(0).Enabled = True
    cmdPesquisa(2).Enabled = True
    
    cmdImprimiRomaneio(0).Enabled = False
        
End Sub

Private Sub estoquecdreal(LojaOrigem As String, referencia As String)

  sql = "select ES_Estoque from Estoque" _
      & " Where  ES_Referencia = '" & referencia & "' and " _
      & " ES_Loja = '" & cmdTransfMerc(1).Caption & "'"

     
        rs4.CursorLocation = adUseServer
        rs4.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockBatchOptimistic
      
    If Not rs4.EOF Then
        lblrealestoque.Caption = rs4("ES_Estoque")
    End If
    
    rs4.Close
End Sub

Private Sub imprimirseparacao()

    Dim SQL1 As String
    Dim romaneio As String
    Dim romaneio1 As String
    Dim quanti As String
    
    pesquisaPorRomaneio
    
    pesquisaPorLoja grd2

    If txtRefereForne.Text = "" And txtRefereForne.Text = "" And grd2.Rows > 1 Then
        SQL1 = whereRomaneio & " And " & whereLoja
    Else
        SQL1 = "  ro_numeroromaneio between " & txtRefereForne.Text & " and " & TxtFantasia.Text
    End If
   
   
   grdRomaneio.Rows = 2
            
    sql = " select ro_sequencia, ro_lojaOrigem, ro_lojaDestino,ro_quantidadePedida, ro_quantidadeEnviada, ro_referencia, " _
        & " pr_referencia, pr_descricao, ro_numeroRomaneio, pr_codigoFornecedor, pr_codigobarra " _
        & " from romaneio, produto, loja " _
        & " where ro_situacao <> 'C' " _
        & " and ro_numeroRomaneio > 0 " _
        & " and ro_referencia = pr_referencia " _
        & " and lo_loja = ro_lojaDestino " _
        & " and ro_lojaOrigem = '" & cmdTransfMerc(1).Caption & "'" _
        & " and " & SQL1 & " order by ro_lojaOrigem,ro_numeroromaneio"
    
    
    
    rs.CursorLocation = adUseServer
    rs.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockBatchOptimistic
    
    Do While Not rs.EOF
          
        romaneio = rs("ro_numeroRomaneio")
            
        If romaneio1 = "" Then
            romaneio1 = rs("ro_numeroRomaneio")
        End If
            
        If romaneio1 = romaneio Then
            
            If chkimpeSemRomaerio.Value = 1 Then
                quanti = "________"
            Else
                quanti = rs("ro_quantidadeEnviada")
            End If
            
            grdRomaneio.AddItem rs("ro_lojaOrigem") & Chr(9) & _
                                rs("ro_lojaDestino") & Chr(9) & _
                                rs("ro_quantidadePedida") & Chr(9) & _
                                quanti & Chr(9) & _
                                rs("pr_referencia") & Chr(9) & _
                                rs("pr_codigobarra") & Chr(9) & _
                                rs("pr_descricao") & Chr(9) & _
                                rs("ro_numeroRomaneio")
                            
            romaneio1 = rs("ro_numeroRomaneio")
            rs.MoveNext
            
            
        Else
            romaneio1 = rs("ro_numeroRomaneio")
            
            If chkimpeSemRomaerio.Value = 1 Or contchk = 1 Then
                Imprimir
            Else
                Imprimir2
            End If
            
            limpaGrid grdRomaneio
        End If
        
        Loop
        
        If chkimpeSemRomaerio.Value = 1 Or contchk = 1 Then
            Imprimir
        Else
            Imprimir2
        End If
        
        limpaGrid grdRomaneio
        
            
        rs.Close
        sql = "update romaneio set ro_impressao='I' where " & SQL1
        ADO_Cn_CDLocal.Execute sql
            

End Sub

Private Sub chkImpeRomaneio_Click()
    If chkImpeRomaneio.Value = 1 Then
        Screen.MousePointer = 11
        imprimirseparacao
        frmeImprimir.Visible = False
        limpaGrid grd2
        Carregagrd2
        limpaGrid grdRomaneio
        Screen.MousePointer = 0
        chkImpeRomaneio.Value = 0
    End If
End Sub

Private Sub chkimpeSemRomaerio_Click()
    If chkimpeSemRomaerio.Value = 1 Then
        Screen.MousePointer = 11
        imprimirseparacao
        frmeImprimir.Visible = False
        limpaGrid grd2
        Carregagrd2
        limpaGrid grdRomaneio
        Screen.MousePointer = 0
        chkimpeSemRomaerio.Value = 0
    End If
End Sub

Private Sub cmdCriarRomaneio_Click(Index As Integer)
    
    Screen.MousePointer = 11
    'CarregaNumeroRomaneio
    CriaRomaneio
    cmdCriarRomaneio(1).Enabled = False
    limpaGrid grd2
    Carregagrd2
    limpaGrid grdRomaneio
    frmStartaProcessos.Show 1
    Screen.MousePointer = 0

End Sub

Private Sub cmdTransfMerc_Click(Index As Integer)
    
    If cmdTransfMerc(1).Caption = "CD" Then
        cmdTransfMerc(1).Caption = "CMC"
    ElseIf cmdTransfMerc(1).Caption = "CMC" Then
        cmdTransfMerc(1).Caption = "CMCE"
    Else
        cmdTransfMerc(1).Caption = "CD"
    End If

End Sub

Private Sub Form_Activate()
    Call opiCriar_MouseUp(0, 0, 0, 0)
End Sub

Private Sub Form_Load()
'ConectaADO
   NomeTxt = LblNomeFantasia.Caption
   lblLagura = LblNomeFantasia.Width
   Largura = TxtFantasia.Width

'--- Determina Posicionamento ---
    'FrmCriaRomaneio.top = (Screen.Height - FrmCriaRomaneio.Height) / 2
    'FrmCriaRomaneio.left = (Screen.Width - FrmCriaRomaneio.Width) / 2

   carregarPosicaoTamanhoTela Me

'--- Mesclando linhas e Colunas ---
    grdRomaneio.MergeRow(0) = True
    grdRomaneio.MergeRow(1) = True
    grdRomaneio.MergeCol(0) = True
    grdRomaneio.MergeCol(1) = True
    grdRomaneio.MergeCol(2) = True
    grdRomaneio.MergeCol(3) = True
    grdRomaneio.MergeCol(4) = True
    grdRomaneio.MergeCol(5) = True
    grdRomaneio.MergeCol(6) = True
    
    grdRomaneio.ColWidth(0) = 0
    grdRomaneio.ColWidth(1) = 700 - 75
    grdRomaneio.ColWidth(2) = 900 - 75
    grdRomaneio.ColWidth(3) = 900 - 75
    grdRomaneio.ColWidth(4) = 980 - 75
    grdRomaneio.ColWidth(5) = 1480 - 75
    grdRomaneio.ColWidth(6) = 4640 - 75
    grdRomaneio.ColWidth(7) = 1000 - 75
   
   
    'JanelaTOP Me  'sair se clicar no retorna

     ''''''grids novos
    CarregaLojas
    orderByPadrao = " order by ro_nu "
    
    'cmdImprimiRomaneio(0).Enabled = False
    cmdCriarRomaneio(1).Enabled = False
    cmdImprimir(0).Enabled = False
    cmdPesquisa(2).Enabled = False
    
    ManutencaoRomaneio = False

End Sub

Public Sub CarregaGrdRomaneio()

       
    sql = " select ro_sequencia,ro_lojaOrigem, ro_lojaDestino, ro_quantidadePedida, ro_quantidadeEnviada, ro_referencia, " _
        & " pr_referencia, pr_descricao , ro_numeroRomaneio, pr_codigoFornecedor, pr_codigoBarra " _
        & " from Romaneio, Produto, loja " _
        & " where ro_situacao='A' and ro_numeroromaneio > 0  and ro_referencia = pr_referencia  and  lo_loja = ro_lojaDestino  and ro_lojaOrigem = '" & cmdTransfMerc(1).Caption & "'" _
        & vWhere & vWhere2 & vwhere3 & " order by lo_loja"


    rs.CursorLocation = adUseServer
    rs.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockBatchOptimistic
      
    grdRomaneio.Rows = 2
    If Not rs.EOF Then
    Do While Not rs.EOF
    
        grdRomaneio.AddItem rs("ro_lojaOrigem") & Chr(9) & rs("ro_lojaDestino") & Chr(9) & _
                            rs("ro_quantidadePedida") & Chr(9) & rs("ro_quantidadeEnviada") & Chr(9) & _
                            rs("pr_referencia") & Chr(9) & rs("pr_codigoBarra") & Chr(9) & rs("pr_descricao") & Chr(9) & _
                            rs("ro_numeroRomaneio") & Chr(9) & rs("ro_sequencia")
    
        rs.MoveNext
    Loop
    Else
    MsgBox "Romaneio não encontrado !", vbInformation, "Atenção"
    End If
    rs.Close

End Sub

Public Sub CarregaNumeroRomaneio()

    Dim imprimirRomaneio As Boolean

    If MsgBox("Deseja Imprimir Romaneio para Separação?", vbYesNo) = vbYes Then
        imprimirRomaneio = True
    End If
    
    grdRomaneio.Rows = 2
    
    sql = " select ro_lojaOrigem, ro_lojaDestino, ro_quantidadePedida, ro_quantidadeEnviada, ro_referencia, " _
    & " pr_referencia, pr_descricao , ro_numeroRomaneio, pr_codigoFornecedor, ro_Sequencia, PR_CodigoBarra " _
    & " from Romaneio, Produto, loja " _
    & "where ro_lojaOrigem = '" & cmdTransfMerc(1).Caption & "'" & orderByPadrao

    rs.CursorLocation = adUseServer
    rs.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockBatchOptimistic

    wlojaant = "999"

    Do While Not rs.EOF
    
        If wlojaant <> rs("ro_lojaDestino") Then
        
            sql = "update controleCDM set cs_NumeroRomaneio = (cs_NumeroRomaneio + 1)"
            ADO_Cn_CDLocal.Execute sql
            
            sql = "select * from ControleCDM "
            rs2.CursorLocation = adUseServer
            rs2.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockBatchOptimistic
            
            wNumeroRomaneio = rs2("cs_NumeroRomaneio")
            
            If wlojaant <> "999" And imprimirRomaneio Then
                Imprimir
                grdRomaneio.Rows = 2
            End If
            
            wlojaant = rs("ro_LojaDestino")
            
            rs2.Close
    
        End If
    
        sql = "update Romaneio set ro_NumeroRomaneio = " & wNumeroRomaneio & "," & vbNewLine _
        & " ro_dataprocesso = '" & Format(Date, "YYYY/MM/DD") & "'," & vbNewLine _
        & " ro_conexao = '" & "I" & "'" & vbNewLine _
        & " Where ro_sequencia = " & rs("ro_Sequencia") _
        & " and ro_lojaOrigem = '" & cmdTransfMerc(1).Caption & "'"
        
        ADO_Cn_CDLocal.Execute sql
        
    
        grdRomaneio.AddItem rs("ro_lojaOrigem") & Chr(9) & _
        rs("ro_lojaDestino") & Chr(9) & _
        rs("ro_quantidadePedida") & Chr(9) & _
        "________" & Chr(9) & _
        rs("pr_referencia") & Chr(9) & _
        rs("PR_CodigoBarra") & Chr(9) & _
        rs("pr_descricao") & Chr(9) & _
        wNumeroRomaneio
        
         'If wlojaant <> rs("ro_lojaDestino") Then
 
         'End If
        
        rs.MoveNext
    Loop

    rs.Close

    If imprimirRomaneio Then
        Imprimir
        sql = "update romaneio set ro_impressao='I' where ro_NumeroRomaneio=" & grdRomaneio.TextMatrix(2, 7)
        ADO_Cn_CDLocal.Execute sql
    End If
    
End Sub



Private Sub frmeImprimir_DragDrop(Source As Control, x As Single, y As Single)
    frmeImprimir.Visible = False
End Sub

Private Sub grdRomaneio_AfterEdit(ByVal row As Long, ByVal col As Long)
    
    Dim novo As String
    Dim solicitada As String
    Dim origem As String
    Dim Loja As String
    Dim referencia As String
    Dim numero As String
    Dim sequencia As String

    If vbKeyReturn Then
    solicitada = grdRomaneio.TextMatrix(grdRomaneio.row, 2)
    novo = grdRomaneio.TextMatrix(grdRomaneio.row, 3)
    If CInt(novo) <= CInt(solicitada) Then
        Screen.MousePointer = 11
        
        origem = grdRomaneio.TextMatrix(grdRomaneio.row, 0)
        Loja = grdRomaneio.TextMatrix(grdRomaneio.row, 1)
        referencia = grdRomaneio.TextMatrix(grdRomaneio.row, 4)
        numero = grdRomaneio.TextMatrix(grdRomaneio.row, 7)
        sequencia = grdRomaneio.TextMatrix(grdRomaneio.row, 8)
        
        sql = "manutencaoromaneio('" & origem & "','" & Loja & "','" & referencia & "'," & novo & "," & qtdantiga & "," & numero & "," & sequencia & ")"
        ADO_Cn_CDLocal.Execute sql
        
        ManutencaoRomaneio = True
        CarregaGRD3
        grdRomaneio.col = col
                
        If row < grdRomaneio.Rows Then
            grdRomaneio.row = row
        Else
            grdRomaneio.row = row - 1
        End If
        
        grdRomaneio.SetFocus
        Call estoquecdreal(origem, referencia)
        Screen.MousePointer = 0
            
    Else
        
        MsgBox "Quantidade maior do que Solicitada", vbInformation, "Atenção!"
        
        If txtRefereForne.Text <> "" Then
            CarregaGrdRomaneio
        Else
            CarregaGRD3
        End If
        
    End If
    End If
    
End Sub

Private Sub grdRomaneio_Click()
    grdRomaneio.Editable = flexEDNone
    frmeImprimir.Visible = False
End Sub

Private Sub grdRomaneio_DblClick()
    grdRomaneio.Editable = flexEDNone
End Sub

Private Sub grdRomaneio_KeyPress(KeyAscii As Integer)
    
    If KeyAscii >= 48 And KeyAscii <= 57 Then

        If (optManutencao.Value = True) And grdRomaneio.col = 3 And grdRomaneio.row >= 2 Then
            qtdantiga = CInt(grdRomaneio.TextMatrix(grdRomaneio.row, 3))
            grdRomaneio.TextMatrix(grdRomaneio.row, 3) = ""
            grdRomaneio.Editable = flexEDKbdMouse
        Else
            grdRomaneio.Editable = flexEDNone
        End If
           
    End If

    If optManutencao.Value = True Then
        Call estoquecdreal(cmdTransfMerc(1).Caption, grdRomaneio.TextMatrix(grdRomaneio.row, 4))
    End If
   
End Sub

Private Sub grdRomaneio_KeyPressEdit(ByVal row As Long, ByVal col As Long, KeyAscii As Integer)

If KeyAscii = 13 Then
    
    
    If txtRefereForne.Text <> "" Then
        CarregaGrdRomaneio
    Else
        CarregaGRD3
    End If
    
    grdRomaneio.col = col
                
    If row < grdRomaneio.Rows Then
        grdRomaneio.row = row
    Else
        grdRomaneio.row = row - 1
    End If
        grdRomaneio.SetFocus
End If

End Sub


Private Sub grdRomaneio_KeyUp(KeyCode As Integer, Shift As Integer)
    

    Dim Mensagem As String

    If chkDeletatodos.Value = 1 Then
        Mensagem = "Confirma cancelamento do Romaneio?"
    Else
        Mensagem = "Confirma cancelamento do Item?"
    End If

    If KeyCode = 46 And optCancelar.Value = True Then
        If MsgBox(Mensagem, vbYesNo + vbQuestion + vbDefaultButton2, "Cancela Romaneio") = vbYes Then
           
           If chkDeletatodos.Value <> 1 Then
                ' cancela 1
                cancelaRomaneio grdRomaneio.TextMatrix(grdRomaneio.row, 7), grdRomaneio.TextMatrix(grdRomaneio.row, 4), grdRomaneio.TextMatrix(grdRomaneio.row, 1), grdRomaneio.TextMatrix(grdRomaneio.row, 3)
            Else
                ' cancela todos
                
                For i = 2 To grdRomaneio.Rows - 1
                   cancelaRomaneio grdRomaneio.TextMatrix(2, 7), grdRomaneio.TextMatrix(2, 4), grdRomaneio.TextMatrix(2, 1), grdRomaneio.TextMatrix(2, 3)
                Next i
                      
            End If
        
        End If
        chkDeletatodos.Value = 0
    End If
End Sub

Private Sub grdRomaneio_SelChange()
If optManutencao.Value = True Then
   Call estoquecdreal("CD", grdRomaneio.TextMatrix(grdRomaneio.row, 4))
End If
   
End Sub


Private Sub opiCriar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If opiCriar.Value = True Then

        LblNomeFantasia.Caption = NomeTxt
        'LblNomeFantasia.Width = lblLagura
        TxtFantasia.Width = Largura
        TxtFantasia.Visible = True
        txtRefereForne.Visible = True
        LblNomeFantasia.Visible = True
        LblRefereForne.Visible = True
        LblRefereForne.Caption = "Referencia/Fornecedor"
        cmdCriarRomaneio(1).Enabled = False
        txtRefereForne.Text = ""
        TxtFantasia.Text = ""
        txtRefereForne.SetFocus
        grd2.Enabled = True
        grdRomaneio.Enabled = True
        grdLojas.Enabled = True
        limpaGrid grdLojas
        CarregaLojas
        limpaGrid grd2
        limpaGrid grdRomaneio
        grdRomaneio.SelectionMode = flexSelectionByRow
        cmdCriarRomaneio(1).Enabled = True
        cmdImprimir(0).Enabled = False
        lblEstoquecd.Visible = False
        lblrealestoque.Visible = False
        grd2.ColHidden(3) = True
        frmeImprimir.Visible = False
        cmdImprimiRomaneio(0).Enabled = False
        cmdCriarRomaneio(1).Enabled = False
        cmdImprimir(0).Enabled = False
        cmdPesquisa(2).Enabled = False
        cmdPesquisa(2).Caption = "Pesquisa"
        chkTodasLojas.Value = 0
        chkTodosRomaneios.Value = 0
        marcaTodasCaixa grdLojas, chkTodasLojas, 0
        chkDeletatodos.Visible = False
End If

    If ManutencaoRomaneio = True Then
        ManutencaoRomaneio = False
        frmStartaProcessos.Show 1
    End If

End Sub


Private Sub optCancelar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If optCancelar.Value = True Then
        
        TxtFantasia.Visible = False
        lblrealestoque.Visible = False
        txtRefereForne.Visible = True
        LblNomeFantasia.Visible = False
        lblEstoquecd.Visible = False
        LblRefereForne.Visible = True
        LblRefereForne.Caption = "Numero Romaneio"
        txtRefereForne.Text = ""
        grdLojas.Enabled = True
        grd2.Enabled = True
        chkTodasLojas.Enabled = True
        chkTodosRomaneios.Enabled = True
        grdRomaneio.SelectionMode = flexSelectionByRow
        cmdPesquisa(2).Caption = "Pesquisa"
        cmdCriarRomaneio(1).Enabled = False
        cmdImprimir(0).Enabled = False
        limpaGrid grd2
        grdLojas.Rows = 1
        CarregaLojas
        limpaGrid grdRomaneio
        grd2.ColHidden(3) = True
        frmeImprimir.Visible = False
        cmdImprimiRomaneio(0).Enabled = False
        cmdCriarRomaneio(1).Enabled = False
        cmdImprimir(0).Enabled = False
        cmdPesquisa(2).Enabled = False
        cmdPesquisa(2).Caption = "Pesquisa"
        chkTodasLojas.Value = 0
        chkTodosRomaneios.Value = 0
        chkDeletatodos.Visible = True
        marcaTodasCaixa grdLojas, chkTodasLojas, 0
        
End If

    If ManutencaoRomaneio = True Then
        ManutencaoRomaneio = False
        frmStartaProcessos.Show 1
    End If

End Sub

Private Sub optImprimir_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If optImprimir.Value = True Then
    
    
        LblNomeFantasia.Caption = "Nº Fim de Romaneio"
        'LblNomeFantasia.Width = LblRefereForne.Width
        TxtFantasia.Width = txtRefereForne.Width
        TxtFantasia.Visible = True
        lblrealestoque.Visible = False
        txtRefereForne.Visible = True
        LblNomeFantasia.Visible = True
        lblEstoquecd.Visible = False
        LblRefereForne.Visible = True
        LblRefereForne.Caption = "Nº Inicio de Romaneio "
        txtRefereForne.Text = ""
        grdLojas.Enabled = True
        grd2.Enabled = True
        chkTodasLojas.Enabled = True
        chkTodosRomaneios.Enabled = True
        grdRomaneio.SelectionMode = flexSelectionByRow
        cmdPesquisa(2).Caption = "Pesquisa"
        cmdCriarRomaneio(1).Enabled = False
        cmdImprimir(0).Enabled = False
        limpaGrid grd2
        grdLojas.Rows = 1
        CarregaLojas
        limpaGrid grdRomaneio
        grd2.ColHidden(3) = False
        frmeImprimir.Visible = False
        cmdImprimiRomaneio(0).Enabled = False
        cmdCriarRomaneio(1).Enabled = False
        cmdImprimir(0).Enabled = False
        cmdPesquisa(2).Enabled = False
        cmdPesquisa(2).Caption = "Pesquisa"
        chkTodasLojas.Value = 0
        chkTodosRomaneios.Value = 0
        marcaTodasCaixa grdLojas, chkTodasLojas, 0
        chkDeletatodos.Visible = False
End If

    If ManutencaoRomaneio = True Then
        ManutencaoRomaneio = False
        frmStartaProcessos.Show 1
    End If

End Sub

Private Function manutencaoestoque(LojaOrigem As String, LojaDestino As String, referencia As String, novo As Integer, antigo As Integer, numeroromaneio As Integer, sequencia As Integer)

Dim quantidade As Integer
quantidade = antigo - novo

    sql = "Update Estoque set" _
    & " ES_Estoque = (ES_Estoque + " & quantidade & ")" _
    & " Where  ES_Referencia = '" & referencia & "' and " _
    & " ES_Loja = '" & LojaOrigem & "'"
    
    ADO_Cn_CDLocal.Execute sql
    
    
    sql = "Update Estoque set " _
    & " ES_Romaneio = (ES_Romaneio - " & quantidade & ")" _
    & " Where  ES_Referencia =  '" & referencia & "' and " _
    & " ES_Loja = '" & LojaDestino & "'"

    ADO_Cn_CDLocal.Execute sql
     

If novo <> 0 Then
    sql = "Update Romaneio set  RO_QuantidadeEnviada= " & novo & "" _
        & " where  RO_LojaOrigem = '" & LojaOrigem & "'  and RO_LojaDestino ='" & LojaDestino & "'  and " _
        & " RO_Referencia = '" & referencia & "'  and RO_NumeroRomaneio = " & numeroromaneio & "" _
        & " And ro_sequencia = " & sequencia & ""
        
        ADO_Cn_CDLocal.Execute sql
         
Else
    sql = "update  Romaneio set ro_situacao='C' where  RO_LojaOrigem = '" & LojaOrigem & "'   and RO_LojaDestino ='" & LojaDestino & "'  and " _
        & " RO_Referencia = '" & referencia & "'  and RO_NumeroRomaneio = " & numeroromaneio & "" _
        & " and ro_sequencia=" & sequencia & ""
    
    ADO_Cn_CDLocal.Execute sql
End If
End Function

Private Sub optManutencao_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If optManutencao.Value = True Then

            TxtFantasia.Visible = False
            lblrealestoque.Visible = True
            lblrealestoque.top = TxtFantasia.top
            lblrealestoque.left = TxtFantasia.left
            txtRefereForne.Visible = True
            LblNomeFantasia.Visible = False
            lblEstoquecd.Visible = True
            lblEstoquecd.top = LblNomeFantasia.top
            lblEstoquecd.left = LblNomeFantasia.left
            LblRefereForne.Visible = True
            cmdCriarRomaneio(1).Enabled = True
            LblRefereForne.Caption = "Numero Romaneio"
            txtRefereForne.Text = ""
            lblrealestoque.Caption = ""
            grdLojas.Enabled = True
            frmeImprimir.Visible = False
            grd2.Enabled = True
            chkTodasLojas.Enabled = True
            chkTodosRomaneios.Enabled = True
            grdRomaneio.SelectionMode = flexSelectionFree
            cmdPesquisa(2).Caption = "Pesquisa"
            cmdCriarRomaneio(1).Enabled = False
            cmdImprimir(0).Enabled = False
            cmdImprimiRomaneio(0).Enabled = False
            cmdCriarRomaneio(1).Enabled = False
            cmdImprimir(0).Enabled = False
            cmdPesquisa(2).Enabled = False
            cmdPesquisa(2).Caption = "Pesquisa"
            grd2.ColHidden(3) = True
            limpaGrid grd2
            grdLojas.Rows = 1
            CarregaLojas
            limpaGrid grdRomaneio
            chkTodasLojas.Value = 0
            chkTodosRomaneios.Value = 0
            marcaTodasCaixa grdLojas, chkTodasLojas, 0
            chkDeletatodos.Visible = False
    
End If
End Sub

Private Sub TxtFantasia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdPesquisa(2).SetFocus
End If

End Sub

Private Sub txtRefereForne_Click()

    If txtRefereForne.Text <> "" Then
 '       limpaGrid grdLojas
 '       CarregaLojas
        limpaGrid grdRomaneio
        'TxtFantasia.Text = ""
        cmdCriarRomaneio(1).Enabled = False
          cmdImprimiRomaneio(0).Enabled = False
    cmdCriarRomaneio(1).Enabled = False
    cmdImprimir(0).Enabled = False
     cmdPesquisa(2).Enabled = False
     cmdPesquisa(2).Caption = "Pesquisa"
End If
End Sub

Private Sub txtRefereForne_GotFocus()

    limpaGrid grdRomaneio
    limpaGrid grd2
    marcaTodasCaixa grdLojas, False, 0
    chkTodasLojas.Value = False
    chkTodosRomaneios.Value = False
    If optCancelar.Value Then
        chkDeletatodos.Value = 0
        cmdPesquisa(2).Enabled = False
    End If
    
End Sub

Private Sub txtRefereForne_LostFocus()


    If optManutencao.Value = True Or optCancelar.Value = True Then
        cmdPesquisa(2).Enabled = True
        cmdPesquisa(2).SetFocus

    ElseIf optImprimir.Value = True Then
        
        cmdPesquisa(2).Enabled = True
        TxtFantasia.SetFocus
    
    ElseIf opiCriar.Value = True Then
        
        grdLojas.Enabled = True
        grd2.Enabled = True
        chkTodasLojas.Enabled = True
        chkTodosRomaneios.Enabled = True
        grdLojas.SetFocus
        grdLojas.row = 1
    ElseIf optCancelar.Value Then
        chkDeletatodos.Value = 0
        cmdPesquisa(2).Enabled = False
    End If

    If Len(txtRefereForne.Text) = 3 And opiCriar.Value = True Then
        TxtFantasia.Text = buscaNomeFornecedor(Mid(txtRefereForne.Text, 1, 3))
    ElseIf Len(txtRefereForne.Text) = 7 And opiCriar.Value = True Then
        sql = "Select pr_descricao from produto  where pr_referencia ='" & txtRefereForne.Text & "'"
        
        rs.CursorLocation = adUseServer
        rs.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
        If Not rs.EOF Then
            Do While Not rs.EOF
            
                TxtFantasia.Text = rs("pr_descricao")
                rs.MoveNext
            Loop
            rs.Close
         Else
            TxtFantasia.Text = ""
         End If
         
    End If


End Sub

'--- Evento dos Botoes
Private Sub cmdRetorna_Click(Index As Integer)
    
    Call opiCriar_MouseUp(0, 0, 0, 0)
    
    frmControleCD.lblNomeTelas.Caption = ""
    Unload Me
     
End Sub
Private Sub cmdPesquisa_Click(Index As Integer)
Dim origem As String
Dim Loja As String
Dim referencia As String
Dim numero As Integer
Dim quantidade As Integer
Dim sequencia As Integer
Dim deleta As String
cmdTransfMerc(1).Enabled = False

   contchk = 0
    vWhere = ""
    vWhere2 = ""
    vwhere3 = ""
    '-----LimpaGrid e values

    Screen.MousePointer = 11
If optImprimir.Value = True And cmdPesquisa(2).Caption = "Pesquisa" Then
      
    vWhere2 = " and ro_numeroromaneio between " & txtRefereForne.Text & " and " & TxtFantasia.Text
    If txtRefereForne.Text <> "" And TxtFantasia.Text <> "" Then
 
        limpaGrid grdRomaneio
        chkTodasLojas.Value = 0
        chkTodosRomaneios.Value = 0
        limpaGrid grdLojas
        CarregaLojas
        CarregaGrdRomaneio
     
    End If

End If
      
      

If optManutencao.Value = True Or optCancelar.Value = True And cmdPesquisa(2).Caption = "Pesquisa" Then

    Screen.MousePointer = 11
    vWhere2 = " and ro_numeroromaneio=" & txtRefereForne.Text
 
    If txtRefereForne.Text <> "" Then
 
'       limpaGrid grd2
'        chkTodasLojas.Value = 0
'       chkTodosRomaneios.Value = 0
'        limpaGrid grdLojas
'        CarregaLojas
'        CarregaGrdRomaneio
        CarregaGRD3
     
    End If
    
    Screen.MousePointer = 0
    '
    If grdRomaneio.Rows > 2 Then
        
        grdRomaneio.row = 2
        grdRomaneio.SetFocus
    
    End If

End If


If cmdPesquisa(2).Caption = "Deletar" Then
numero = CInt(grdRomaneio.TextMatrix(2, 7))
For i = 2 To grdRomaneio.Rows - 1
numero = CInt(grdRomaneio.TextMatrix(i, 7))
    If MsgBox("Deseja cancelar o Romaneio " & grdRomaneio.TextMatrix(i, 7) & "?", vbQuestion + vbYesNo, "Cancelamento de Romaneio") = vbYes Then
     Screen.MousePointer = 11
                Do While numero = grdRomaneio.TextMatrix(i, 7)
                        origem = grdRomaneio.TextMatrix(i, 0)
                        Loja = grdRomaneio.TextMatrix(i, 1)
                        quantidade = CInt(grdRomaneio.TextMatrix(i, 3))
                        referencia = grdRomaneio.TextMatrix(i, 4)
                        numero = CInt(grdRomaneio.TextMatrix(i, 7))
                        sequencia = CInt(grdRomaneio.TextMatrix(i, 8))
                        sql = "manutencaoromaneio('" & origem & "','" & Loja & "','" & referencia & "',0," & quantidade & "," & numero & "," & sequencia & ")"
                        ADO_Cn_CDLocal.Execute sql
            If i < grdRomaneio.Rows - 1 Then
                i = i + 1
            Else
                i = i + 1
                Exit Do
            End If
                
                Loop
                
            Screen.MousePointer = 0
            
    Else
    Do While numero = grdRomaneio.TextMatrix(i, 7)
                       numero = CInt(grdRomaneio.TextMatrix(i, 7))
            If i < grdRomaneio.Rows - 1 Then
                i = i + 1
            Else
                i = i + 1
                Exit Do
            End If
         
          Loop
          End If
    
     i = i - 1
 Next i
 
frmStartaProcessos.Show 1

cmdPesquisa(2).Caption = "Pesquisar"
cmdPesquisa(2).Enabled = False
If txtRefereForne.Text <> "" Then
    CarregaGrdRomaneio
 
Else

    CarregaGRD3
    Carregagrd2
    Call optCancelar_MouseUp(0, 0, 0, 0)
    
End If
End If


If grdRomaneio.Rows <> 2 And optCancelar.Value = True Then
  ' cmdPesquisa(2).Caption = "Deletar"
Else
    cmdPesquisa(2).Caption = "Pesquisa"
End If

If grdRomaneio.Rows <> 2 And optImprimir.Value = True And optImprimir.Value = True Then
   cmdImprimir(0).Enabled = True
End If


    Screen.MousePointer = 0


End Sub


Private Sub cmdImprimir_Click(Index As Integer)

    If grdRomaneio.Rows > 2 Then
        frmeImprimir.Visible = True
    End If
    
End Sub
Public Sub Imprimir2()

    grdRomaneio.BackColor = vbWhite
    grdRomaneio.BackColorAlternate = vbWhite
    grdRomaneio.BackColorBkg = vbWhite
    grdRomaneio.BackColorFixed = vbWhite
    grdRomaneio.BackColorFrozen = vbWhite
    grdRomaneio.ForeColorFixed = vbBlack
    grdRomaneio.ForeColor = &H80000006
    
    grdRomaneio.ColWidth(0) = 0
    grdRomaneio.ColWidth(1) = 700
    grdRomaneio.ColWidth(2) = 900
    grdRomaneio.ColWidth(3) = 900
    grdRomaneio.ColWidth(4) = 980
    grdRomaneio.ColWidth(5) = 1480
    grdRomaneio.ColWidth(6) = 4640
    grdRomaneio.ColWidth(7) = 1000
    
    grdRomaneio.FontSize = 9
    
    grdRomaneio.PrintGrid "Romaneio " & Now, False, 1, 300, 500
    
    grdRomaneio.BackColor = &H303030
    grdRomaneio.BackColorAlternate = &H3C3C3C
    grdRomaneio.BackColorBkg = &H505050
    grdRomaneio.BackColorFixed = &H0
    grdRomaneio.BackColorFrozen = vbWhite
    grdRomaneio.ForeColorFixed = &HFA9923
    grdRomaneio.ForeColor = &H80000005
    
    grdRomaneio.ColWidth(0) = 0
    grdRomaneio.ColWidth(1) = 700
    grdRomaneio.ColWidth(2) = 900
    grdRomaneio.ColWidth(3) = 900
    grdRomaneio.ColWidth(4) = 980
    grdRomaneio.ColWidth(5) = 1480
    grdRomaneio.ColWidth(6) = 4640
    grdRomaneio.ColWidth(7) = 1000
    
    grdRomaneio.FontSize = 8
           
End Sub


Private Sub cmdLimpa_Click(Index As Integer)

    
    txtRefereForne.Text = ""
    TxtFantasia.Text = ""
    grdRomaneio.Rows = 2
    
    txtRefereForne.Enabled = True
    TxtFantasia.Enabled = True
    'CmbLoja.SetFocus
    
   
    txtRefereForne.Enabled = True
    TxtFantasia.Enabled = True
   
    chkTodasLojas.Value = 0
    chkTodosRomaneios.Value = 0
  
    cmdImprimiRomaneio(0).Enabled = False
    cmdCriarRomaneio(1).Enabled = False
    cmdImprimir(0).Enabled = False
    cmdPesquisa(2).Enabled = False
    cmdPesquisa(2).Caption = "Pesquisa"
    limpaGrid grdLojas
    CarregaLojas
      
    limpaGrid grd2
    limpaGrid grdRomaneio
    chkTodasLojas.Value = 0
    
End Sub

Private Sub txtRomaneio_Click()
    'cmbLoja.Enabled = False
    txtRefereForne.Enabled = False
    TxtFantasia.Enabled = False
End Sub

Private Sub txtRomaneio_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        vWhere = ""
        
        vWhere = " and ro_numeroRomaneio = " ' & 'txtRomaneio.Text
    
            
        sql = " select ro_lojaOrigem, ro_lojaDestino, ro_quantidadePedida, ro_quantidadeEnviada, ro_referencia, " _
                & " pr_referencia, pr_descricao , pr_codigoBarras ,ro_numeroRomaneio, ro_DataSolicitacao, ro_Sequencia,lo_regiao " _
                & " from Romaneio, Produto, loja " _
                & " where ro_referencia = pr_referencia and lo_loja = ro_lojaDestino and ro_lojaOrigem = '" & cmdTransfMerc(1).Caption & "'" _
                & vWhere & " order by lo_Regiao,ro_numeroRomaneio"
              
        rs.CursorLocation = adUseServer
        rs.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockBatchOptimistic
          
        grdRomaneio.Rows = 2
    
        Do While Not rs.EOF
        
            grdRomaneio.AddItem rs("ro_lojaOrigem") & Chr(9) & rs("ro_lojaDestino") & Chr(9) & _
                                rs("ro_quantidadePedida") & Chr(9) & rs("ro_quantidadeEnviada") & Chr(9) & _
                                rs("pr_referencia") & Chr(9) & rs("pr_codigoBarras") & rs("pr_descricao") & Chr(9) & _
                                rs("ro_numeroRomaneio")
                               
            rs.MoveNext
        Loop
        
        rs.Close
    End If
End Sub


''''''''''''''''''''''''''''''''''''''''''''GRIDS NOVOS

Private Sub CarregaLojas()  ' grdLojas
  
    Dim rsLojas As New ADODB.Recordset
    Dim sql As String
    
    sql = "select lo_loja From Loja Where (lo_Regiao < 450 or lo_loja = 'CMC') and LO_Loja Not in('185') order by lo_Regiao"
    
    rsLojas.CursorLocation = adUseClient
    rsLojas.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    limpaGrid grdLojas
    
    Do While Not rsLojas.EOF
        grdLojas.AddItem chkTodasLojas & Chr(9) & rsLojas("Lo_Loja")
        rsLojas.MoveNext
    Loop
     
    rsLojas.Close
   
End Sub

Private Sub Carregagrd2() 'GRD2

    Dim rsRomaneio As New ADODB.Recordset
    Dim statusImprecao As String
    Dim sql As String
  
    
    limpaGrid grd2
   cmdTransfMerc(1).Enabled = False
   
    sql = "select distinct ro_impressao,ro_lojadestino,ro_numeroromaneio,lo_regiao from romaneio, produto, loja" & _
          " where ro_situacao='A' and ro_referencia = pr_referencia and lo_loja =  ro_lojaDestino and ro_lojaOrigem = '" & cmdTransfMerc(1).Caption & "'"
   
    If pesquisaPorLoja(grdLojas) Then
        sql = sql & " and " & whereLoja
    Else
        Exit Sub
    End If
    
 
    If opiCriar.Value = True Then
            If Len(txtRefereForne) < 4 And Len(txtRefereForne) > 1 Then
                vWhere2 = " and pr_CodigoFornecedor = '" & txtRefereForne.Text & "' "
            
            ElseIf Len(txtRefereForne) = 7 Then
                 vWhere2 = " and ro_referencia = '" & txtRefereForne.Text & "' "
            Else
                    vWhere2 = ""
            End If
    
    sql = sql & " and ro_numeroRomaneio = 0"
    ElseIf optManutencao.Value = True Or optCancelar.Value = True Or optImprimir.Value = True Then
     sql = sql & " and ro_numeroRomaneio <> 0 "
     vWhere2 = ""
    End If
    
    
    sql = sql & vWhere2 & " "
    
    sql = sql + " order by ro_numeroromaneio"
    rsRomaneio.CursorLocation = adUseClient
    rsRomaneio.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    
    Do While Not rsRomaneio.EOF
        
        
        If rsRomaneio("ro_impressao") = "N" Then
            statusImprecao = "Não"
        ElseIf rsRomaneio("ro_impressao") = "I" Then
            statusImprecao = "Sim"
        End If
    grd2.AddItem chkTodosRomaneios & Chr(9) & Replace(rsRomaneio("ro_lojadestino"), " ", "") & Chr(9) & rsRomaneio("ro_numeroromaneio") & Chr(9) & statusImprecao
        
        rsRomaneio.MoveNext
    Loop
    
        
    rsRomaneio.Close
    
End Sub

Private Sub CarregaGRD3() 'GRDRomaneio

    Dim rsRomaneio As New ADODB.Recordset
    Dim sql As String
    ' Limpa Grid
    limpaGrid grdRomaneio
    
    ' Consulta Lojas Selecionadas
    pesquisaPorLoja grd2
    pesquisaPorRomaneio
    
        If optManutencao.Value Or optCancelar.Value Then
            If txtRefereForne.Text <> "" Then
                ' Select 1
                sql = "select ro_lojaOrigem, ro_lojaDestino, ro_quantidadePedida, ro_quantidadeEnviada, ro_referencia, " _
                    & " pr_referencia, PR_CodigoBarra, pr_descricao , ro_numeroRomaneio, ro_sequencia " _
                    & " from Romaneio, Produto" _
                    & " where ro_referencia = pr_referencia" _
                    & " and ro_situacao = 'A'" _
                    & " and ro_Numeroromaneio = " & txtRefereForne.Text _
                    & " and ro_lojaOrigem = '" & cmdTransfMerc(1).Caption & "'" _
                    & " order by ro_numeroRomaneio, ro_lojaDestino asc, ro_referencia"
            Else
                If grd2.Rows > 1 Then
                    sql = "select ro_lojaOrigem, ro_lojaDestino, ro_quantidadePedida, ro_quantidadeEnviada, ro_referencia, " _
                        & " pr_referencia, PR_CodigoBarra, pr_descricao , ro_numeroRomaneio, ro_sequencia " _
                        & " from Romaneio, Produto" _
                        & " where ro_referencia = pr_referencia" _
                        & " and ro_situacao = 'A'" _
                        & " and " & whereLoja _
                        & " and " & whereRomaneio _
                        & " and ro_lojaOrigem = '" & cmdTransfMerc(1).Caption & "'" _
                        & " order by ro_numeroRomaneio, ro_lojaDestino asc, ro_referencia"
                End If

            End If
    
        ElseIf opiCriar.Value Then
        
                        
            ' Select 2
            sql = "select ro_lojaOrigem, ro_lojaDestino, ro_quantidadePedida, ro_quantidadeEnviada, ro_referencia, " _
                & " pr_referencia, PR_CodigoBarra, pr_descricao , ro_numeroRomaneio, ro_sequencia " _
                & " from Romaneio, Produto, loja" _
                & " where ro_referencia = pr_referencia" _
                & " and ro_lojaDestino = lo_loja" _
                & " and ro_situacao = 'A'" _
                & " and " & whereLoja _
                & " and " & whereRomaneio _
                & " and ro_lojaOrigem = '" & cmdTransfMerc(1).Caption & "'" _
                & " and ro_referencia like '" & txtRefereForne.Text & "%'" _
                & " and ro_numeroRomaneio = 0" _
                & " order by ro_numeroRomaneio, ro_lojaDestino asc, ro_referencia"
            
        ElseIf optImprimir.Value Then
        
            ' Select Impressao
                
            If grd2.Rows > 1 Then
            
                If txtRefereForne.Text = "" And TxtFantasia.Text = "" Then
                    
                    sql = "select ro_lojaOrigem, ro_lojaDestino, ro_quantidadePedida, ro_quantidadeEnviada, ro_referencia, " _
                        & " pr_referencia, PR_CodigoBarra, pr_descricao , ro_numeroRomaneio, ro_sequencia " _
                        & " from Romaneio, Produto" _
                        & " where ro_referencia = pr_referencia" _
                        & " and ro_situacao <> 'C'" _
                        & " and " & whereLoja _
                        & " and " & whereRomaneio _
                        & " and ro_lojaOrigem = '" & cmdTransfMerc(1).Caption & "'" _
                        & " and ro_numeroRomaneio <> 0" _
                        & " order by ro_numeroRomaneio, ro_lojaDestino asc, ro_referencia"
                Else
                
                    sql = "select ro_lojaOrigem, ro_lojaDestino, ro_quantidadePedida, ro_quantidadeEnviada, ro_referencia, " _
                        & " pr_referencia, PR_CodigoBarra, pr_descricao , ro_numeroRomaneio, ro_sequencia " _
                        & " from Romaneio, Produto" _
                        & " where ro_referencia = pr_referencia" _
                        & " and ro_situacao <> 'C'" _
                        & " and ro_lojaOrigem = '" & cmdTransfMerc(1).Caption & "'" _
                        & " and ro_numeroRomaneio between '" & Format(txtRefereForne.Text, "yyyy/mm/dd") & "' and '" & Format(TxtFantasia.Text, "yyyy/mm/dd") & "'" _
                        & " order by ro_numeroRomaneio, ro_lojaDestino asc, ro_referencia"
                
                End If
            
            End If
    
        End If
    
        rsRomaneio.CursorLocation = adUseClient
        rsRomaneio.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    
        Do While Not rsRomaneio.EOF And opiCriar.Value = True
            grdRomaneio.AddItem rsRomaneio("ro_lojaOrigem") & Chr(9) & rsRomaneio("ro_lojaDestino") & Chr(9) & _
                                rsRomaneio("ro_quantidadePedida") & Chr(9) & "________" & Chr(9) & _
                                rsRomaneio("pr_referencia") & Chr(9) & rsRomaneio("PR_CodigoBarra") & Chr(9) & _
                                rsRomaneio("pr_descricao") & Chr(9) & _
                                rsRomaneio("ro_numeroRomaneio") & Chr(9) & _
                                rsRomaneio("ro_sequencia")
            rsRomaneio.MoveNext
        Loop
    
    
        Do While Not rsRomaneio.EOF And (optManutencao.Value = True Or optImprimir.Value = True Or optCancelar.Value = True)
            grdRomaneio.AddItem rsRomaneio("ro_lojaOrigem") & Chr(9) & rsRomaneio("ro_lojaDestino") & Chr(9) & _
                                rsRomaneio("ro_quantidadePedida") & Chr(9) & rsRomaneio("ro_quantidadeEnviada") & Chr(9) & _
                                rsRomaneio("pr_referencia") & Chr(9) & rsRomaneio("PR_CodigoBarra") & Chr(9) & _
                                rsRomaneio("pr_descricao") & Chr(9) & _
                                rsRomaneio("ro_numeroRomaneio") & Chr(9) & _
                                rsRomaneio("ro_sequencia")
            rsRomaneio.MoveNext
        Loop
    
        rsRomaneio.Close
        
    
End Sub
Private Function pesquisaPorLoja(ByRef grid2) As Boolean
    
    Dim i As Integer
    Dim a As Integer
    a = 0
    pesquisaPorLoja = False
    whereLoja = "ro_lojaDestino in("
    
     For i = 1 To grid2.Rows - 1
        If grid2.TextMatrix(i, 0) Then
            a = a + 1
            whereLoja = whereLoja & "'" & RTrim(grid2.TextMatrix(i, 1)) & "',"
            pesquisaPorLoja = True
        End If
    Next i
    
    If a <> 0 Then
        whereLoja = left(whereLoja, (Len(whereLoja) - 1)) & ")"
    Else
        whereLoja = whereLoja & " ' ' ) "
    End If
End Function

Private Function pesquisaPorRomaneio() As Boolean
    
    Dim i As Integer
    Dim a As Integer
    a = 0
    pesquisaPorRomaneio = False
    
    whereRomaneio = "ro_numeroromaneio in ("
    For i = 1 To grd2.Rows - 1
        If grd2.TextMatrix(i, 0) = True Then
        a = a + 1
            whereRomaneio = whereRomaneio & "'" & RTrim(grd2.TextMatrix(i, 2)) & "',"
            pesquisaPorRomaneio = True
        End If
    Next i
    If a <> 0 Then
    
        whereRomaneio = left(whereRomaneio, (Len(whereRomaneio) - 1)) & ")"
    Else
        whereRomaneio = whereRomaneio & " '')"
    End If

End Function

Private Sub grdLojas_Click()
    limpaGrid grdRomaneio
    marcaCaixaGrid grdLojas, chkTodasLojas
    Carregagrd2
    grdRomaneio.Enabled = True
    cmdImprimiRomaneio(0).Enabled = False
    cmdCriarRomaneio(1).Enabled = False
    cmdImprimir(0).Enabled = False
    cmdPesquisa(2).Enabled = False
    cmdPesquisa(2).Caption = "Pesquisa"
    
    If optManutencao.Value = True Or optCancelar.Value = True Or optImprimir.Value = True Then
        txtRefereForne.Text = ""
        TxtFantasia.Text = ""
    End If
    
End Sub

Private Sub grd2_Click()
    Screen.MousePointer = 11
    limpaGrid grdRomaneio
    marcaCaixaGrid grd2, chkTodosRomaneios
    CarregaGRD3
    Screen.MousePointer = 0
    
    If grdRomaneio.Rows > 2 And opiCriar.Value = True Then
        cmdCriarRomaneio(1).Enabled = True
    End If
 
    If grdRomaneio.Rows <> 2 And optImprimir.Value = True Then
        cmdImprimir(0).Enabled = True
    Else
        cmdImprimir(0).Enabled = False
    End If
    
End Sub


Private Sub chkTodosRomaneios_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  
    marcaTodasCaixa grd2, chkTodosRomaneios, 0
    CarregaGRD3
    Carregagrd2
    If optImprimir.Value = True Then
      cmdImprimir(0).Enabled = True
    End If
    
    If grdRomaneio.Rows > 2 And opiCriar.Value = True Then
        cmdCriarRomaneio(1).Enabled = True
    End If
    If optCancelar.Value = True And grdRomaneio.Rows > 2 Then
        cmdPesquisa(2).Enabled = True
    '    cmdPesquisa(2).Caption = "Deletar"
    End If
    If grdRomaneio.Rows <> 2 And optImprimir.Value = True Then
        cmdImprimir(0).Enabled = True
    Else
        cmdImprimir(0).Enabled = False
    End If
    
End Sub

Private Sub chkTodasLojas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If chkTodasLojas.Value = 0 Then
        chkTodosRomaneios.Value = 0
    End If
    
    marcaTodasCaixa grdLojas, chkTodasLojas, 0
    Carregagrd2
     
    grdRomaneio.Enabled = True
    cmdImprimiRomaneio(0).Enabled = False
    cmdCriarRomaneio(1).Enabled = False
    cmdImprimir(0).Enabled = False
    cmdPesquisa(2).Enabled = False
    cmdPesquisa(2).Caption = "Pesquisa"
     
    If optManutencao.Value = True Or optCancelar.Value = True Or optImprimir.Value = True Then
    txtRefereForne.Text = ""
    TxtFantasia.Text = ""
    End If
         
     
End Sub

Private Sub marcaCaixaGrid(ByRef grid, check As CheckBox)
    Dim i As Integer

    grid.col = 0
    If grid.row >= grid.FixedRows Then
        If grid.Text Then
            grid.Text = False
            check.Value = False
        Else
            grid.Text = True
           
        End If
    End If
End Sub
Private Sub marcaTodasCaixa(ByRef grid, ativa As Boolean, coluna As Byte)
    Dim i As Integer
    
    If Not gridVazio(grid) Then
        For i = grid.FixedRows To grid.Rows - 1
            grid.TextMatrix(i, coluna) = ativa
        Next i
    End If
    
End Sub

Private Sub cmdImprimiRomaneio_Click(Index As Integer)

    limpaGrid grdRomaneio
  
    Dim rsRomaneio As New ADODB.Recordset
    Dim sql As String
    Dim numero As String

    
    sql = " select ro_lojaOrigem, ro_lojaDestino, ro_quantidadePedida, ro_quantidadeEnviada, ro_referencia, " _
            & " pr_referencia, pr_descricao , ro_numeroRomaneio, ro_sequencia " _
            & " from Romaneio, Produto, loja " _
            

    sql = sql & " where ro_referencia = pr_referencia and lo_loja = ro_lojaDestino" _
              & " and ro_lojaOrigem = '" & cmdTransfMerc(1).Caption & "'" _
    
    
    If pesquisaPorLoja(grdLojas) And pesquisaPorRomaneio And pesquisaPorLoja(grd2) Then
        sql = sql & " and " & whereRomaneio & " and " & whereLoja & " and ro_situacao='A'"
    Else
    Exit Sub
    End If
    
    sql = sql + orderByPadrao
     
    rsRomaneio.CursorLocation = adUseClient
    rsRomaneio.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    numero = rsRomaneio("ro_numeroRomaneio")
    
    grdRomaneio.Redraw = False
        Do While Not rsRomaneio.EOF
                             
            If numero <> rsRomaneio("ro_numeroRomaneio") Then
                Imprimir
                limpaGrid grdRomaneio
                numero = rsRomaneio("ro_numeroRomaneio")
            End If
            grdRomaneio.AddItem rsRomaneio("ro_lojaOrigem") & Chr(9) & rsRomaneio("ro_lojaDestino") & Chr(9) & _
                                rsRomaneio("ro_quantidadePedida") & Chr(9) & rsRomaneio("ro_quantidadeEnviada") & Chr(9) & _
                                rsRomaneio("pr_referencia") & Chr(9) & rsRomaneio("pr_descricao") & Chr(9) & _
                                rsRomaneio("ro_numeroRomaneio") & Chr(9) & rsRomaneio("ro_sequencia")
           
            rsRomaneio.MoveNext
        Loop
     
    Imprimir
    grdRomaneio.Redraw = True
    rsRomaneio.Close
    
    
                   ' grdromaneio2
End Sub


Public Sub Imprimir()

    'sql = "select * from Romaneio where RO_RomaneioImpressao = "
    
    'rs2.CursorLocation = adUseServer
    'rs2.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockBatchOptimistic
    
    grdRomaneio.BackColor = vbWhite
    grdRomaneio.BackColorAlternate = vbWhite
    grdRomaneio.BackColorBkg = vbWhite
    grdRomaneio.BackColorFixed = vbWhite
    grdRomaneio.BackColorFrozen = vbWhite
    grdRomaneio.ForeColorFixed = vbBlack
    grdRomaneio.ForeColor = &H80000006
    
    grdRomaneio.ColWidth(0) = 0
    grdRomaneio.ColWidth(1) = 700
    grdRomaneio.ColWidth(2) = 900
    grdRomaneio.ColWidth(3) = 900
    grdRomaneio.ColWidth(4) = 980
    grdRomaneio.ColWidth(5) = 1480
    grdRomaneio.ColWidth(6) = 4640
    grdRomaneio.ColWidth(7) = 1000
    
    grdRomaneio.FontSize = 9
    
    grdRomaneio.PrintGrid "Romaneio de Separação " & Now, False, 1, 300, 500
    'grdRomaneio.PrintGrid "Romaneio de Separação " & Date, False, 1, 300, 500
    
    
    grdRomaneio.BackColor = grd2.BackColor
    grdRomaneio.BackColorAlternate = grd2.BackColorAlternate
    grdRomaneio.BackColorBkg = grd2.BackColorBkg
    grdRomaneio.BackColorFixed = grd2.BackColorFixed
    grdRomaneio.BackColorFrozen = grd2.BackColorFrozen
    grdRomaneio.ForeColorFixed = grd2.ForeColorFixed
    grdRomaneio.ForeColor = grd2.ForeColor
    
    
    grdRomaneio.ColWidth(0) = 0
    grdRomaneio.ColWidth(1) = 700 - 75
    grdRomaneio.ColWidth(2) = 900 - 75
    grdRomaneio.ColWidth(3) = 900 - 75
    grdRomaneio.ColWidth(4) = 980 - 75
    grdRomaneio.ColWidth(5) = 1480 - 75
    grdRomaneio.ColWidth(6) = 4640 - 75
    grdRomaneio.ColWidth(7) = 1000 - 75
    
    grdRomaneio.FontSize = 8
           
End Sub


Private Sub cancelaRomaneio(romaneio As String, referencia As String, Loja As String, quantidadeAnterior As String)

Dim sql As String

        ' Atualiza Romaneio loja Destino
        sql = "update estoque set es_romaneio = es_romaneio - " & quantidadeAnterior & "where es_loja = '" & Loja & "' and es_referencia = '" & referencia & "'"
        ADO_Cn_CDLocal.Execute (sql)
        
        ' Atualiza Estoque Loja Origem
        sql = "update estoque set es_estoque = es_estoque + " & quantidadeAnterior & "where es_loja = '" & cmdTransfMerc(1).Caption & "' and es_referencia = '" & referencia & "'"
        ADO_Cn_CDLocal.Execute (sql)
        
        If referencia = "" Then
        
            ' Apaga Romaneio de todas as referências
            sql = "update romaneio set ro_situacao = 'C' where ro_numeroRomaneio = '" & romaneio & "' and ro_lojaDestino = '" & Loja & "'"
            ADO_Cn_CDLocal.Execute (sql)
            
        Else
            
            ' Apaga Romaneio de uma única Referencia
            sql = "update romaneio set ro_situacao = 'C' where ro_numeroRomaneio = '" & romaneio & "' and ro_referencia  = '" & referencia & "' and ro_lojaDestino = '" & Loja & "'"
            ADO_Cn_CDLocal.Execute (sql)
        
        End If
        
        CarregaGRD3
        

End Sub


Private Sub CriaRomaneio()
    
    Dim sql As String
    Dim imprimirRomaneio As Boolean
    Dim sequencias() As String
    Dim Romaneios() As String
    Dim i As Integer
    Dim wWhere As String
    Dim Tamanho As String
    Dim rsRomaneio As New ADODB.Recordset
    Dim a As Integer
    
    i = 0
    Tamanho = grdRomaneio.Rows
    ReDim sequencias(Tamanho)
    ReDim Romaneios(Tamanho)
    
    If MsgBox("Deseja Imprimir Romaneio para Separação?", vbYesNo) = vbYes Then
        imprimirRomaneio = True
    End If
   
   ' Armazena as Sequências
    If imprimirRomaneio = True Then
        For i = 0 To Tamanho - 1
            If i > 1 Then
                sequencias(i) = grdRomaneio.TextMatrix(i, 8)
            Else
                sequencias(i) = 0
            End If
           Next i
    End If
    
    ' Cria Romaneios
    Do While grdRomaneio.Rows > 2
        sql = "Exec SP_CDM_Cria_Romaneio '" & grdRomaneio.TextMatrix(2, 1) & "'"
        ADO_Cn_CDLocal.Execute (sql)
        opiCriar.Value = True
        CarregaGRD3
    Loop
    
    ' Where
    wWhere = ""
    
    For i = 0 To Tamanho - 1
        If i > 1 Then
            wWhere = sequencias(i) & " , " & wWhere
        Else
            wWhere = sequencias(i)
        End If
      Next i
    
        
    ' Seleciona Romaneios

    sql = "select distinct (ro_numeroRomaneio) as romaneio from romaneio where ro_sequencia in ( " & wWhere & ") and ro_lojaOrigem = '" & cmdTransfMerc(1).Caption & "'"
    rsRomaneio.CursorLocation = adUseClient
    rsRomaneio.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    i = 0
    
    Do While Not rsRomaneio.EOF
        Romaneios(i) = rsRomaneio("romaneio")
        rsRomaneio.MoveNext
        i = i + 1
    Loop
    rsRomaneio.Close

    
    ' Imprime
    If imprimirRomaneio Then
        
        For a = 0 To i - 1
            ' Limpa Grid
            grdRomaneio.Rows = 2
            
            ' Busca
            sql = "select * from romaneio, produto where ro_numeroRomaneio in ( " & Romaneios(a) & ")" _
                & " and ro_referencia = pr_referencia" _
                & " and ro_lojaOrigem = '" & cmdTransfMerc(1).Caption & "'" _
                & " order by ro_lojaDestino, ro_referencia"
     
            rsRomaneio.CursorLocation = adUseClient
            rsRomaneio.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
            ' CarregaGrid
            Do While Not rsRomaneio.EOF
                grdRomaneio.AddItem rsRomaneio("ro_lojaOrigem") & Chr(9) & rsRomaneio("ro_lojaDestino") & Chr(9) & _
                                    rsRomaneio("ro_quantidadePedida") & Chr(9) & "________" & Chr(9) & _
                                    rsRomaneio("pr_referencia") & Chr(9) & rsRomaneio("PR_CodigoBarra") & Chr(9) & _
                                    rsRomaneio("pr_descricao") & Chr(9) & _
                                    rsRomaneio("ro_numeroRomaneio") & Chr(9) & _
                                    rsRomaneio("ro_sequencia")
                                    rsRomaneio.MoveNext
            Loop
            
            'Imprime
            Imprimir
            
            sql = "update romaneio set ro_impressao='I' where ro_NumeroRomaneio=" & grdRomaneio.TextMatrix(2, 7)
            ADO_Cn_CDLocal.Execute sql
            rsRomaneio.Close
          Next a
    End If
  
End Sub

