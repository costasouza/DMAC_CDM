VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form frmTransfMerc 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Transferência de Mercadoria"
   ClientHeight    =   8025
   ClientLeft      =   2310
   ClientTop       =   2490
   ClientWidth     =   15300
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form9"
   LockControls    =   -1  'True
   ScaleHeight     =   8025
   ScaleWidth      =   15300
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmItem 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Item "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1845
      TabIndex        =   15
      Top             =   7065
      Visible         =   0   'False
      Width           =   14910
      Begin VB.Label Label5 
         BackColor       =   &H00404040&
         Caption         =   "Referência"
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
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblReferencia 
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
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00404040&
         Caption         =   "Descrição"
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
         Left            =   1560
         TabIndex        =   21
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lblDescicao 
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
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   1560
         TabIndex        =   20
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404040&
         Caption         =   "Estoque CD."
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
         Left            =   7560
         TabIndex        =   19
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblestoquecd 
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
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   7560
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackColor       =   &H00404040&
         Caption         =   "Total de Sugestão"
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
         Left            =   8760
         TabIndex        =   17
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label lbltotalsu 
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
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   8760
         TabIndex        =   16
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   44
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   14880
      TabIndex        =   12
      Top             =   6810
      Width           =   14880
   End
   Begin VB.TextBox txtSugestao 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   675
      TabIndex        =   10
      Text            =   "Sugestão"
      Top             =   1545
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Frame fraPesquisa 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Pesquisa "
      ForeColor       =   &H000000C0&
      Height          =   840
      Left            =   150
      TabIndex        =   9
      Top             =   150
      Width           =   14880
      Begin VB.TextBox txtNomeFornecedor 
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
         Left            =   2235
         TabIndex        =   4
         Top             =   360
         Width           =   5325
      End
      Begin VB.ComboBox cmbComprador 
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
         Left            =   7620
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   2235
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
         Height          =   315
         Left            =   150
         TabIndex        =   3
         Top             =   360
         Width           =   2010
      End
      Begin VB.Label lblNome 
         BackColor       =   &H00404040&
         Caption         =   "Nome"
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
         Left            =   2235
         TabIndex        =   14
         Top             =   150
         Width           =   1710
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         Caption         =   "Comprador"
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
         Left            =   7620
         TabIndex        =   13
         Top             =   150
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblPesquisa 
         BackColor       =   &H00404040&
         Caption         =   "Fornecedor/Referência"
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
         Height          =   210
         Left            =   150
         TabIndex        =   11
         Top             =   150
         Width           =   2010
      End
   End
   Begin CentroDeDistribuicao.chameleonButton cmdTransfMerc 
      Height          =   510
      Index           =   0
      Left            =   13620
      TabIndex        =   0
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Retorna"
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
      MICON           =   "TransfMerc.frx":0000
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
      Index           =   3
      Left            =   12180
      TabIndex        =   1
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
      MICON           =   "TransfMerc.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid grdSugestao 
      Height          =   975
      Left            =   150
      TabIndex        =   2
      Top             =   5670
      Width           =   7575
      _cx             =   13361
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
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"TransfMerc.frx":0038
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
      Editable        =   2
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   0
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   5263440
      ForeColorFrozen =   4210752
      WallPaperAlignment=   9
   End
   Begin CentroDeDistribuicao.chameleonButton cmdTransfMerc 
      Height          =   510
      Index           =   4
      Left            =   10740
      TabIndex        =   6
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
      MICON           =   "TransfMerc.frx":0102
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid grdTende 
      Height          =   4350
      Left            =   150
      TabIndex        =   8
      ToolTipText     =   "Informe sua Sugestão"
      Top             =   1155
      Width           =   14910
      _cx             =   26300
      _cy             =   7673
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   20
      Cols            =   19
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"TransfMerc.frx":011E
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
      ComboSearch     =   0
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   5263440
      ForeColorFrozen =   4210752
      WallPaperAlignment=   9
   End
   Begin VSFlex7LCtl.VSFlexGrid grdItens 
      Height          =   975
      Left            =   150
      TabIndex        =   7
      Top             =   5400
      Width           =   7575
      _cx             =   13361
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
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"TransfMerc.frx":029D
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
      ComboSearch     =   0
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   5263440
      ForeColorFrozen =   4210752
      WallPaperAlignment=   9
   End
   Begin CentroDeDistribuicao.chameleonButton cmdTransfMerc 
      Height          =   510
      Index           =   1
      Left            =   9240
      TabIndex        =   25
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
      MICON           =   "TransfMerc.frx":0325
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblControle 
      BackStyle       =   0  'Transparent
      Caption         =   "ControleTela"
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   6960
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmTransfMerc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

Dim wWhere As String
Dim wwhere2 As String
Dim sql As String
Dim wQuantidade As String
Dim wQtdgrid As Integer
Dim wdestino As String
Dim resultado As Integer
Dim Linhas As Long
Dim ColunaTende As Integer
Dim LinhaTende As Integer
Dim ColunaItens As Integer
Dim LinhaItens As Integer
Dim ColunaSugestao As Integer
Dim LinhaSugestao As Integer
Dim WTotalSuge As Integer
Dim wtotalestoquereal As Integer
Dim forRefe As String
Dim Total As Long
Dim focu As Integer
Dim WcontrolRomaneio As Long
Dim quantidadeAnterior As String


Private Sub cmdTransfMerc_Click(Index As Integer)  ' Botões

    Select Case Index
        Case 0
     
            frmControleCD.lblNomeTelas.Caption = ""
            Unload Me
        Case 1
            
            If cmdTransfMerc(1).Caption = "CD" Then
                cmdTransfMerc(1).Caption = "CMC"
            ElseIf cmdTransfMerc(1).Caption = "CMC" Then
                cmdTransfMerc(1).Caption = "CMCE"
            Else
                cmdTransfMerc(1).Caption = "CD"
            End If
            
        Case 3
            tela2
            NovaPesquisa
            limpaGrid
                 
        Case 4
            tela2
            limpaGrid
            Pesquisa
            cmdTransfMerc(4).Enabled = False
            
           'tela2
    End Select

End Sub



Private Sub Form_Load() ' no carregamento do Formulario, determina o Tamanho
    
    focu = 0
    Dim MesesPassados As Variant
    Screen.MousePointer = 11

    'frmTransfMerc.top = (Screen.Height - frmTransfMerc.Height) / 2
    'frmTransfMerc.left = (Screen.Width - frmTransfMerc.Width) / 2

   'Me.top = posicaoTelaY
   'Me.left = posicaoTelaX
   
    WTotalSuge = 0
    'fraComparativo.Visible = True
    grdItens.top = grdTende.top
    grdItens.Height = grdTende.Height
    grdTende.left = (grdSugestao.Width + grdSugestao.left) + 150
    grdTende.Width = 7120
    
    carregarPosicaoTamanhoTela Me
    MesesPassados = MesesTendencia
    
    grdTende.MergeRow(0) = True
    grdTende.MergeRow(1) = True
    
    grdTende.MergeCol(0) = True
    grdTende.MergeCol(1) = True
    grdTende.MergeCol(2) = True
    grdTende.MergeCol(3) = True
    grdTende.MergeCol(4) = True
    grdTende.MergeCol(5) = True
    grdTende.MergeCol(6) = True
    grdTende.MergeCol(7) = True
    grdTende.MergeCol(8) = True
    grdTende.MergeCol(9) = True
    grdTende.MergeCol(10) = True
    grdTende.MergeCol(11) = True
    grdTende.MergeCol(12) = True
    grdTende.MergeCol(13) = True
    grdTende.MergeCol(14) = True
    grdTende.MergeCol(15) = True
    grdTende.MergeCol(16) = True
    grdTende.MergeCol(17) = True
    
    grdTende.TextMatrix(0, 0) = "Loja"
    grdTende.TextMatrix(0, 1) = "Mix"
    grdTende.TextMatrix(0, 2) = "Sugestão"
    grdTende.TextMatrix(0, 3) = "Roman."
    grdTende.TextMatrix(0, 4) = "Trânsito"
    grdTende.TextMatrix(0, 5) = "Min."
    grdTende.TextMatrix(0, 6) = "Max."
    grdTende.TextMatrix(0, 7) = MesesPassados(0)
    grdTende.TextMatrix(0, 8) = MesesPassados(0)
    grdTende.TextMatrix(0, 9) = MesesPassados(1)
    grdTende.TextMatrix(0, 10) = MesesPassados(1)
    grdTende.TextMatrix(0, 11) = MesesPassados(2)
    grdTende.TextMatrix(0, 12) = MesesPassados(2)
    grdTende.TextMatrix(0, 13) = MesesPassados(3)
    grdTende.TextMatrix(0, 14) = MesesPassados(3)
    grdTende.TextMatrix(0, 15) = MesesPassados(4)
    grdTende.TextMatrix(0, 16) = MesesPassados(4)
    grdTende.TextMatrix(0, 17) = "MinMax"

    grdTende.TextMatrix(1, 0) = "Loja"
    grdTende.TextMatrix(1, 1) = "Mix"
    grdTende.TextMatrix(1, 2) = "Sugestão"
    grdTende.TextMatrix(1, 3) = "Roman."
    grdTende.TextMatrix(1, 4) = "Trânsito"
    grdTende.TextMatrix(1, 5) = "Min."
    grdTende.TextMatrix(1, 6) = "Max."
    grdTende.TextMatrix(1, 7) = "Estoque"
    grdTende.TextMatrix(1, 8) = "Venda"
    grdTende.TextMatrix(1, 9) = "Estoque"
    grdTende.TextMatrix(1, 10) = "Venda"
    grdTende.TextMatrix(1, 11) = "Estoque"
    grdTende.TextMatrix(1, 12) = "Venda"
    grdTende.TextMatrix(1, 13) = "Estoque"
    grdTende.TextMatrix(1, 14) = "Venda"
    grdTende.TextMatrix(1, 15) = "Estoque"
    grdTende.TextMatrix(1, 16) = "Venda"
    grdTende.TextMatrix(1, 17) = "MinMax"
   
   
    grdTende.ColHidden(9) = True
    grdTende.ColHidden(10) = True
    grdTende.ColHidden(11) = True
    grdTende.ColHidden(12) = True
    grdTende.ColHidden(13) = True
    grdTende.ColHidden(14) = True
    grdTende.ColHidden(15) = True
    grdTende.ColHidden(16) = True
    grdTende.ColHidden(17) = True

   
    grdSugestao.Rows = 1
    CarregaComboComprador

    Screen.MousePointer = 0
    
    ControleTela
End Sub

Private Sub CarregaComboComprador()

           
    sql = "Select *  From Comprador " _
  
    rs.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic


    
      Do While Not rs.EOF               ' Fazer enquanto não chegar até o fim
      
       cmbComprador.AddItem Format(rs("co_codigoComprador"), "00") & "-" & rs("co_nome")

rs.MoveNext
 
      Loop
 
    cmbComprador.AddItem "99-Todos"
     cmbComprador.ListIndex = 0
    
    rs.Close
 
  
End Sub
Public Sub Pesquisa() ' Faço um select de acordo com o Codigo do Fornecedor

cmdTransfMerc(1).Enabled = False
forRefe = txtFornecedor.Text
Total = Len(forRefe)

Screen.MousePointer = 11
 wWhere = " "
 wwhere2 = " "
 
  If Total = 3 Then
      wWhere = " and pr_codigoFornecedor = " & txtFornecedor.Text
      
ElseIf Total = 7 Then
        wWhere = " and pr_referencia =   '" & txtFornecedor.Text & "'"
    
  End If
  
'  If Mid(cmbComprador.Text, 1, 2) <> "99" Then
'       wwhere2 = " and pr_Comprador = " & Mid(cmbComprador.Text, 1, 2)
'
'  End If
  
        ' seleciona e carrega grdItens
        sql = " select pr_referencia, pr_descricao,pr_CodigoFornecedor,pr_Comprador,es_estoque " _
            & "from Produto,estoque,Comprador " _
            & "where pr_comprador = co_codigoComprador and pr_referencia= es_referencia and es_loja = '" & Trim(cmdTransfMerc(1).Caption) & "' and es_estoque > 0 " & wWhere & wwhere2

   
        'abre a tabela selecionada
        rs.CursorLocation = adUseServer
        rs.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
       
       grdItens.Rows = 1
       Linhas = 1
       If Not rs.EOF Then
        Do While Not rs.EOF
        'carrega as colunas com os respectivos valores
        grdItens.AddItem rs("pr_referencia") & Chr(9) & rs("pr_descricao") & Chr(9) & rs("es_estoque")
                    grdItens.Cell(flexcpChecked, Linhas, 3) = flexUnchecked
        rs.MoveNext
        Linhas = Linhas + 1
        Loop
        grdItens.SetFocus
        grdItens.col = 0
        grdItens.row = 1
        
Else
    MsgBox "Não Há Referências no Estoque", vbInformation
    txtFornecedor.SetFocus
   
End If
   rs.Close
  Linhas = 0
  Screen.MousePointer = 0
End Sub
Public Sub limpaGrid()
    
    'Limpa grdTende
    grdTende.Rows = 2
    
    'Limpa grdItens
    grdItens.Rows = 1
    
    'Limpa grdSugestao
    grdSugestao.Rows = 1

    cmbComprador.ListIndex = 0
    

End Sub

Public Sub carregaGridTende()

Dim rs As New ADODB.Recordset

    sql = " select  es_loja, es_mixproduto, es_romaneio, es_transito, es_estoqueMinimo,es_estoqueMaximo, " _
           & "es_Estoque, es_Venda, es_Estoque1, es_Venda1, es_Estoque2, es_Venda2, es_Estoque3, es_Venda3, es_Estoque4, es_Venda4, es_Estoque5, es_Venda5" _
           & " From Estoque,Produto,loja " _
           & " where pr_referencia = es_referencia and es_referencia = '" & grdItens.TextMatrix(grdItens.row, 0) & "' and lo_loja = es_loja and lo_situacao = 'A' and lo_loja not in ('conso','cd','mc85','mc85s','185','182','183','cmce','cmcs','mc85e')" _
           & wWhere & " order by lo_regiao"
        
       
      
    'abre um recordset para a tabela selecionada
    rs.CursorLocation = adUseServer
    rs.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
       
    grdTende.Rows = 2

    Do While Not rs.EOF
        'carrega o GridTende e as colunas com os respectivos valores
        grdTende.AddItem rs("es_loja") & Chr(9) & rs("es_mixproduto") & Chr(9) _
                       & Chr(9) & rs("es_romaneio") & Chr(9) & rs("es_transito") _
                       & Chr(9) & rs("es_estoqueMinimo") & Chr(9) & rs("es_estoqueMaximo") _
                       & Chr(9) & rs("es_Estoque") & Chr(9) & rs("es_Venda") _
                       & Chr(9) & rs("es_Estoque1") & Chr(9) & rs("es_Venda1") _
                       & Chr(9) & rs("es_Estoque2") & Chr(9) & rs("es_Venda2") _
                       & Chr(9) & rs("es_Estoque3") & Chr(9) & rs("es_Venda3") _
                       & Chr(9) & rs("es_Estoque4") & Chr(9) & rs("es_Venda4") _
                       & Chr(9) & rs("es_Estoque5") & Chr(9) & rs("es_Venda5") _

      
       rs.MoveNext
        
                '*****
  grdTende.col = 1
  grdTende.row = grdTende.Rows - 1
          If grdTende = "N" Then
             grdTende.CellForeColor = &HFF&                         'cor vermelha 'red'
                
          End If
  
  grdTende.col = 5
  grdTende.row = grdTende.Rows - 1
          If grdTende <> "" Then
             grdTende.CellForeColor = &HFA9923
        
          End If

  grdTende.col = 6
  grdTende.row = grdTende.Rows - 1
          If grdTende <> "" Then
             grdTende.CellForeColor = &HFA9923
        
          End If
 
   
   Loop
   grdTende.AddItem Chr(9) & Chr(9)
    rs.Close
    
       grdTende.left = (grdSugestao.Width + grdSugestao.left) + 150
    grdTende.Width = 7145
  'grdTende.Row = LinhaTende

End Sub
Public Sub NovaPesquisa()   ' limpa campos
    txtFornecedor.Text = ""
    txtNomeFornecedor.Text = ""
    txtFornecedor.SetFocus
    cmdTransfMerc(4).Enabled = True
    
End Sub



Private Sub grdItens_EnterCell() ' um click no GridItens, carrega o GridTende
    'If Linhas = 0 Then
LinhaItens = grdItens.row
lblEstoquecd.Caption = grdItens.TextMatrix(grdItens.row, 2)
Screen.MousePointer = 11
carregaGridTende
carregaSugestao
Screen.MousePointer = 0
'End If


End Sub


Private Sub grdItens_KeyPress(KeyAscii As Integer)
      
      
If KeyAscii = 13 Then
    tela1
    
End If


End Sub

Private Sub grdSugestao_AfterEdit(ByVal row As Long, ByVal col As Long)
 
 Dim quantidadeAnterior As Integer
 Dim Diferenca As Integer
 Dim sql As String
 Dim rs1 As New ADODB.Recordset
 quantidadeAnterior = 0
 
 If vbKeyReturn Then
 
    sql = "select ro_quantidadeEnviada from romaneio where ro_sequencia = '" & grdSugestao.TextMatrix(grdSugestao.row, 5) & "'"
    rs1.CursorLocation = adUseServer
    rs1.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not rs1.EOF Then
        
        quantidadeAnterior = rs1("ro_QuantidadeEnviada")
        
    End If
    rs1.Close
    
    Diferenca = quantidadeAnterior - CInt(grdSugestao.TextMatrix(grdSugestao.row, 2))
    
    ModoSubtrair Diferenca, Trim(grdSugestao.TextMatrix(grdSugestao.row, 0)), Trim(grdSugestao.TextMatrix(grdSugestao.row, 3)), Trim(grdSugestao.TextMatrix(grdSugestao.row, 5)), quantidadeAnterior
    
    carregaSugestao
    carregaGridTende

 End If
End Sub

Private Sub grdSugestao_GotFocus()
focu = 1
tela2
End Sub


Private Sub grdSugestao_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 46 Then
    
        DeletaSugestao Trim(grdSugestao.TextMatrix(grdSugestao.row, 0)), Trim(lblControle.Caption), Trim(grdSugestao.TextMatrix(grdSugestao.row, 3)), Trim(grdSugestao.TextMatrix(grdSugestao.row, 2))
    
    End If

End Sub

Private Sub grdSugestao_LostFocus()
    tela2
End Sub

Private Sub grdTende_AfterEdit(ByVal row As Long, ByVal col As Long)
    Dim QuantidadeAtual As Integer
    
    If vbKeyReturn Then
        If grdTende.TextMatrix(grdTende.row, 2) <> "" Then
            QuantidadeAtual = grdTende.TextMatrix(grdTende.row, 2)
        
            'Carrega a Procedure
            sql = "Exec SP_Transfere_Mercadoria '" & grdItens.TextMatrix(LinhaItens, 0) & "','" & Trim(cmdTransfMerc(1).Caption) & "','" & grdTende.TextMatrix(grdTende.row, 0) & "','" & QuantidadeAtual & "'," & WcontrolRomaneio & ", '" & lblControle.Caption & "'"
            ADO_Cn_CDLocal.Execute sql
        
            resultado = Val(grdItens.TextMatrix(grdItens.row, 2)) - Val(QuantidadeAtual)
            grdItens.TextMatrix(grdItens.row, 2) = resultado
    
            grdTende.col = 2
            grdTende = ""
            wQtdgrid = grdTende.TextMatrix(grdTende.row, 3)
            grdTende.TextMatrix(grdTende.row, 3) = (QuantidadeAtual + wQtdgrid)
    
            WTotalSuge = WTotalSuge + wQuantidade
            lbltotalsu.Caption = WTotalSuge
            wtotalestoquereal = lblEstoquecd.Caption
            wtotalestoquereal = wtotalestoquereal - QuantidadeAtual
            lblEstoquecd.Caption = wtotalestoquereal
            wQuantidade = ""
            carregaSugestao
            carregaGridTende
        End If
    End If
            
End Sub

Private Sub grdTende_DblClick()
    If grdTende.TextMatrix(grdTende.row, 1) = "N" Then
        grdTende.TextMatrix(grdTende.row, 1) = "S"
        
        'Muda Cor
        grdTende.FillStyle = flexFillRepeat
        grdTende.row = grdTende.row
        grdTende.RowSel = grdTende.row
        grdTende.col = 1
        grdTende.ColSel = grdTende.col
        grdTende.CellFontBold = True
        grdTende.CellForeColor = &HC0C0C0
        grdTende.FillStyle = flexFillSingle
    Else
        grdTende.TextMatrix(grdTende.row, 1) = "N"
        
        ' Muda Cor
        grdTende.FillStyle = flexFillRepeat
        grdTende.row = grdTende.row
        grdTende.RowSel = grdTende.row
        grdTende.col = 1
        grdTende.ColSel = grdTende.col
        grdTende.CellFontBold = True
        grdTende.CellForeColor = &HFF&
        grdTende.FillStyle = flexFillSingle
            
    End If
End Sub

Private Sub grdTende_EnterCell()  'Habilita e Desabilita coluna

    LinhaTende = grdTende.row

    If grdTende.col = 2 Then

     If grdTende.TextMatrix(grdTende.row, 1) = "S" Then
      grdTende.Editable = flexEDKbdMouse   ' habilita
     Else

       grdTende.Editable = flexEDNone      ' Desabilita
            
            

     End If

   Else
      grdTende.Editable = flexEDNone
   End If
 
End Sub


Private Sub grdTende_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
    tela2
    End If
    If KeyAscii = 13 And grdTende.col <> 2 Then
    tela2
    End If
    If (grdTende.row + 1) = grdTende.Rows And grdTende.TextMatrix(grdTende.row, 1) = "N" Then
     tela2
   End If
   
End Sub

Private Sub grdTende_KeyPressEdit(ByVal row As Long, ByVal col As Long, KeyAscii As Integer)   ' Busca Procedure (sql)
 
     If KeyAscii >= 48 And KeyAscii <= 57 Then                          ' No campo Sugestão só numerico

         wQuantidade = wQuantidade & Chr(KeyAscii)

     End If

    
    If (grdTende.row + 1) = grdTende.Rows Then
        tela2
    End If

    If (grdTende.row + 1) = grdTende.Rows And wQuantidade = "" Then
        tela2
    End If
    
    End Sub


Private Sub grdTende_LostFocus()
'tela2
End Sub




Private Sub tela1()
                    '------grdSugestao---------
                    'grdSugestao.Rows = 1
                    '--------grdItens---------
grdItens.Cell(flexcpChecked, grdItens.row, 3) = flexChecked
                    grdItens.Visible = False
                    '-----------frame----
                    frmItem.Visible = True
                    frmItem.left = fraPesquisa.left
                    frmItem.top = fraPesquisa.top
                    frmItem.Width = fraPesquisa.Width
                    frmItem.Height = fraPesquisa.Height
                    fraPesquisa.Visible = False
                    '------grdTende
                  
                    'grdTende.Height = 4935
                    grdTende.left = 150
                    'grdTende.top = 720
                    grdTende.Width = 14900
                    grdTende.col = 0
                    grdTende.ColHidden(9) = False
                    grdTende.ColHidden(10) = False
                    grdTende.ColHidden(11) = False
                    grdTende.ColHidden(12) = False
                    grdTende.ColHidden(13) = False
                    grdTende.ColHidden(14) = False
                    grdTende.ColHidden(15) = False
                    grdTende.ColHidden(16) = False
                    grdTende.ColHidden(17) = False

                    
                    
                    grdTende.row = 2
                    grdTende.col = 2
                    '------Label
                    lblDescicao.Caption = grdItens.TextMatrix(grdItens.row, 1)
                    lblReferencia.Caption = grdItens.TextMatrix(grdItens.row, 0)
                    lblEstoquecd.Caption = grdItens.TextMatrix(grdItens.row, 2)
                    lbltotalsu.Caption = 0
                    WTotalSuge = 0
                    cmdTransfMerc(0).Enabled = False
                    cmdTransfMerc(3).Enabled = False
                    cmdTransfMerc(4).Enabled = False
                    
                    
                    
End Sub



Private Sub tela2()
                    
                    frmItem.Visible = False
                    grdItens.Visible = True
                    fraPesquisa.Visible = True
                    If focu = 1 Then
                    focu = 0
                    Else
                    grdItens.SetFocus
                    focu = 0
                    End If
                    grdItens.top = grdTende.top
                    grdItens.Height = grdTende.Height
                    grdTende.left = (grdSugestao.Width + grdSugestao.left) + 150
                    grdTende.Width = 7120
                    grdTende.ColHidden(9) = True
                    grdTende.ColHidden(10) = True
                    grdTende.ColHidden(11) = True
                    grdTende.ColHidden(12) = True
                    grdTende.ColHidden(13) = True
                    grdTende.ColHidden(14) = True
                    grdTende.ColHidden(15) = True
                    grdTende.ColHidden(16) = True
                    grdTende.ColHidden(17) = True
                    cmdTransfMerc(0).Enabled = True
                    cmdTransfMerc(3).Enabled = True
                    cmdTransfMerc(4).Enabled = True
                    carregaSugestao
End Sub




Private Sub carregagrdSugestao()
grdSugestao.Rows = 1

sql = "select Ro_lojadestino,Ro_referencia, ro_sequencia,pr_descricao,ro_quantidadepedida,ro_numeroromaneio,ro_datasolicitacao" _
  & " from romaneio,Produto where ro_situacao='A' and pr_referencia=Ro_referencia " _
  & "and RO_RomaneioImpressao='" & WcontrolRomaneio & "' order by ro_datasolicitacao desc"
         rs1.CursorLocation = adUseServer
            rs1.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
Do While Not rs1.EOF
      grdSugestao.AddItem rs1("Ro_lojadestino") & Chr(9) & _
                           rs1("ro_datasolicitacao") & Chr(9) & _
                         rs1("ro_quantidadepedida") & Chr(9) & _
                         rs1("Ro_referencia") & Chr(9) & _
                         rs1("pr_descricao") & Chr(9) & _
                         rs1("ro_sequencia")

     rs1.MoveNext
     
                         
Loop
rs1.Close


End Sub

Private Sub txtNomeFornecedor_GotFocus()

forRefe = txtFornecedor.Text
Total = Len(forRefe)

    If Total = "2" Or Total = "3" Then
        sql = "Select fo_razaosocial from fornecedor where fo_codigoFornecedor = '" & forRefe & "'"
        
        rs.CursorLocation = adUseServer
        rs.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
          If Not rs.EOF Then
       Do While Not rs.EOF
   
       txtNomeFornecedor.Text = rs("fo_razaosocial")
          
 rs.MoveNext
      
       Loop
Else
MsgBox "Não Há este Fornecedor", vbInformation
        txtNomeFornecedor = ""
        txtFornecedor.SetFocus

End If
rs.Close
ElseIf Total = 7 Then
               sql = "Select pr_descricao from produto  where pr_referencia ='" & forRefe & "'"
        
        rs.CursorLocation = adUseServer
        rs.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
          If Not rs.EOF Then
       Do While Not rs.EOF
   
       txtNomeFornecedor.Text = rs("pr_descricao")
          
 rs.MoveNext
         Loop
Else
        MsgBox "Não Há está Referência", vbInformation
        txtNomeFornecedor = ""
        txtFornecedor.SetFocus
    End If
    
    
         rs.Close
Else
    MsgBox "Fornecedor/Referência não Cadastrada! "
  End If

End Sub

Private Sub ModoSubtrair(quantidade As Integer, LojaDestino As String, referencia As String, sequencia As String, quantidadeAnterior As Integer)
    
    Dim sql As String
    
   If quantidade <> 0 Then
   
        ' Atualiza Romaneio loja Destino
        sql = "update estoque set es_romaneio = es_romaneio - " & quantidade & "where es_loja = '" & LojaDestino & "' and es_referencia = '" & referencia & "'"
        ADO_Cn_CDLocal.Execute (sql)
        
        ' Atualiza Estoque Loja Origem
        sql = "update estoque set es_estoque = es_estoque + " & quantidade & "where es_loja = '" & cmdTransfMerc(1).Caption & "' and es_referencia = '" & referencia & "'"
        ADO_Cn_CDLocal.Execute (sql)

        ' Atualiza Quantidade Romaneio
        sql = "update romaneio set ro_quantidadePedida = ro_quantidadePedida - " & quantidade & ", ro_quantidadeEnviada = ro_quantidadeEnviada - " & quantidade & _
        " where ro_sequencia = '" & sequencia & "' and ro_referencia  = '" & referencia & "'"
        
        ADO_Cn_CDLocal.Execute (sql)
    Else
        
        ' Atualiza Romaneio loja Destino
        sql = "update estoque set es_romaneio = es_romaneio - " & quantidadeAnterior & "where es_loja = '" & LojaDestino & "' and es_referencia = '" & referencia & "'"
        ADO_Cn_CDLocal.Execute (sql)
        
        ' Atualiza Estoque Loja Origem
        sql = "update estoque set es_estoque = es_estoque + " & quantidadeAnterior & "where es_loja = '" & cmdTransfMerc(1).Caption & "' and es_referencia = '" & referencia & "'"
        ADO_Cn_CDLocal.Execute (sql)
        
        
        ' Apaga Romaneio Zerado
        sql = "delete romaneio where ro_sequencia = '" & sequencia & "' and ro_referencia  = '" & referencia & "'"
        ADO_Cn_CDLocal.Execute (sql)
    End If

    
End Sub

Private Sub ControleTela()

    Dim sql As String
    Dim Controle As Integer
    Dim rsControle As New ADODB.Recordset
    
    sql = "update controleCDM set cs_romaneioImpressao = cs_romaneioImpressao + 1"
    ADO_Cn_CDLocal.Execute (sql)
    
    sql = "select cs_romaneioImpressao from controleCDM"
    
    rsControle.CursorLocation = adUseServer
    rsControle.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not rsControle.EOF Then
        
        lblControle.Caption = Trim(rsControle("cs_romaneioImpressao"))
        
    End If
    
    rsControle.Close
    
    
End Sub

Private Sub carregaSugestao()

    Dim sql As String
    Dim rsSugestao As New ADODB.Recordset
    
    ' Limpa Grid
    grdSugestao.Rows = 1
    
    'Consulta Sugestões
    sql = "select ro_lojaDestino, ro_dataSolicitacao, ro_quantidadeEnviada, ro_referencia, pr_descricao " _
        & " from romaneio, produto " _
        & " where pr_referencia = ro_referencia" _
        & " and ro_romaneioImpressao = '" & Trim(lblControle.Caption) & "'"
        
    rsSugestao.CursorLocation = adUseServer
    rsSugestao.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    Do While Not rsSugestao.EOF
        
        grdSugestao.AddItem rsSugestao("ro_lojaDestino") & Chr(9) _
        & Format(rsSugestao("ro_dataSolicitacao"), "dd/mm/yyyy") & Chr(9) _
        & rsSugestao("ro_quantidadeEnviada") & Chr(9) _
        & rsSugestao("ro_referencia") & Chr(9) _
        & rsSugestao("pr_descricao")
            
        rsSugestao.MoveNext
    Loop
    
    rsSugestao.Close
    
End Sub

Private Sub DeletaSugestao(Loja As String, Tipo As String, referencia As String, quantidade As String)
    
    Dim sql As String
    
    ' Volta Estoque do CD
    sql = "update estoque set es_estoque = es_estoque +" & quantidade & "where es_loja = '" & cmdTransfMerc(1).Caption & "' and es_referencia = '" & referencia & "'"
    ADO_Cn_CDLocal.Execute (sql)
    
    'Volta Romaneio
    sql = "update estoque set es_romaneio = es_romaneio - " & quantidade & "where es_loja = '" & Loja & " and es_referencia = '" & referencia & "'"
    ADO_Cn_CDLocal.Execute (sql)
    
    ' Deleta Romaneio
    sql = "delete from romaneio where ro_lojaDestino = '" & Loja & "' and ro_tipo = '" & lblControle.Caption & "' and ro_referencia = '" & referencia & "'"
    ADO_Cn_CDLocal.Execute (sql)
    
    carregaSugestao
End Sub
