VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmAgendaDeRecebimento 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Agenda de Recebimento"
   ClientHeight    =   7860
   ClientLeft      =   2880
   ClientTop       =   2505
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   15420
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   44
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   14880
      TabIndex        =   29
      Top             =   6810
      Width           =   14880
   End
   Begin VB.Frame fraConsultadeAgenda 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Consulta de Agenda"
      ForeColor       =   &H80000008&
      Height          =   6555
      Left            =   7875
      TabIndex        =   12
      Top             =   150
      Width           =   7155
      Begin VB.TextBox txtPesqData 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   690
         TabIndex        =   10
         Top             =   450
         Width           =   1185
      End
      Begin VSFlex7DAOCtl.VSFlexGrid grdAgendaRecebimento 
         Height          =   5145
         Left            =   150
         TabIndex        =   13
         Top             =   870
         Width           =   6855
         _cx             =   12091
         _cy             =   9075
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
         BackColorBkg    =   4210752
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAgendaDeRecebimento.frx":0000
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
      Begin CentroDeDistribuicao.chameleonButton CmdMenor 
         Height          =   285
         Left            =   2040
         TabIndex        =   21
         Top             =   450
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   503
         BTYPE           =   14
         TX              =   "<<"
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
         MICON           =   "frmAgendaDeRecebimento.frx":00A7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CentroDeDistribuicao.chameleonButton CmdMaior 
         Height          =   285
         Left            =   2460
         TabIndex        =   22
         Top             =   450
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   503
         BTYPE           =   14
         TX              =   ">>"
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
         MICON           =   "frmAgendaDeRecebimento.frx":00C3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CentroDeDistribuicao.chameleonButton chameleonButton1 
         Height          =   285
         Left            =   2970
         TabIndex        =   23
         Top             =   450
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   503
         BTYPE           =   14
         TX              =   "&Agenda Completa"
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
         MICON           =   "frmAgendaDeRecebimento.frx":00DF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblTotalGeralAgendamento 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblTotalGeralAgendamento"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   150
         TabIndex        =   27
         Top             =   6120
         Width           =   1875
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   150
         TabIndex        =   26
         Top             =   480
         Width           =   390
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Consulta de Agenda"
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
   End
   Begin VB.Frame fraCadastro 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Cadastro de Agenda de Recebimento"
      ForeColor       =   &H80000008&
      Height          =   6555
      Left            =   150
      TabIndex        =   11
      Top             =   150
      Width           =   7575
      Begin VB.Frame frmVencimento 
         Appearance      =   0  'Flat
         BackColor       =   &H00393939&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         ForeColor       =   &H80000008&
         Height          =   2265
         Left            =   150
         TabIndex        =   30
         Top             =   2520
         Width           =   7260
         Begin VSFlex7DAOCtl.VSFlexGrid grdNotas 
            Height          =   1305
            Left            =   150
            TabIndex        =   31
            Top             =   465
            Width           =   6945
            _cx             =   12250
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
            FormatString    =   $"frmAgendaDeRecebimento.frx":00FB
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
         Begin VB.Label Label2 
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Cadastro de Agenda de Recebimento"
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
            Height          =   200
            Left            =   150
            TabIndex        =   33
            Top             =   150
            Width           =   3375
         End
         Begin VB.Label lblTotalAgendamento 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblTotalAgendamento"
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            Left            =   150
            TabIndex        =   32
            Top             =   1905
            Width           =   1500
         End
      End
      Begin VB.TextBox txtObservacao 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   1140
         Left            =   150
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   5250
         Width           =   7275
      End
      Begin VB.ComboBox cmbHora 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1125
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   960
         Width           =   795
      End
      Begin VSFlex7LCtl.VSFlexGrid grdAgendaComplemento 
         Height          =   270
         Left            =   150
         TabIndex        =   16
         Top             =   1455
         Width           =   7275
         _cx             =   12832
         _cy             =   476
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
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAgendaDeRecebimento.frx":0184
         ScrollTrack     =   0   'False
         ScrollBars      =   0
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
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   5263440
         ForeColorFrozen =   4210752
         WallPaperAlignment=   9
      End
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   325
         Left            =   1920
         TabIndex        =   2
         Top             =   960
         Width           =   600
      End
      Begin VB.TextBox txtTelefone 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   325
         Left            =   6045
         MaxLength       =   30
         TabIndex        =   9
         Top             =   1740
         Width           =   1380
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   325
         Left            =   2910
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1740
         Width           =   3150
      End
      Begin VB.TextBox txtTransportadora 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   325
         Left            =   150
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1740
         Width           =   2775
      End
      Begin VB.TextBox txtVolume 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   325
         Left            =   6435
         MaxLength       =   50
         TabIndex        =   6
         Top             =   960
         Width           =   990
      End
      Begin VB.TextBox txtValor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   325
         Left            =   5145
         MaxLength       =   50
         TabIndex        =   5
         Top             =   960
         Width           =   1305
      End
      Begin VB.TextBox txtNotaFiscal 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   325
         Left            =   4485
         MaxLength       =   50
         TabIndex        =   4
         Top             =   960
         Width           =   675
      End
      Begin VB.TextBox txtFornecedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   325
         Left            =   2505
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   960
         Width           =   1995
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   325
         Left            =   150
         TabIndex        =   0
         Top             =   960
         Width           =   975
      End
      Begin VSFlex7DAOCtl.VSFlexGrid grdAgenda 
         Height          =   555
         Left            =   150
         TabIndex        =   15
         Top             =   450
         Width           =   7275
         _cx             =   12832
         _cy             =   979
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
         Rows            =   2
         Cols            =   7
         FixedRows       =   2
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAgendaDeRecebimento.frx":01F4
         ScrollTrack     =   0   'False
         ScrollBars      =   0
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
      Begin VSFlex7LCtl.VSFlexGrid VSFlexGrid1 
         Height          =   465
         Left            =   150
         TabIndex        =   17
         Top             =   4965
         Width           =   7275
         _cx             =   12832
         _cy             =   820
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
         SelectionMode   =   0
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
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmAgendaDeRecebimento.frx":0301
         ScrollTrack     =   0   'False
         ScrollBars      =   0
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
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   5263440
         ForeColorFrozen =   4210752
         WallPaperAlignment=   9
      End
      Begin VB.Label lblSequencia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblSequencia"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   150
         TabIndex        =   28
         Top             =   2145
         Width           =   915
      End
      Begin VB.Label Label33 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastro de Agenda de Recebimento"
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
         Height          =   200
         Left            =   150
         TabIndex        =   24
         Top             =   150
         Width           =   3375
      End
   End
   Begin CentroDeDistribuicao.chameleonButton cmdSair 
      Height          =   510
      Left            =   13620
      TabIndex        =   19
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
      MICON           =   "frmAgendaDeRecebimento.frx":0336
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdLimpar 
      Height          =   510
      Left            =   10740
      TabIndex        =   20
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
      MICON           =   "frmAgendaDeRecebimento.frx":0352
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmbGravar 
      Height          =   510
      Left            =   12180
      TabIndex        =   18
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "&Gravar"
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
      MICON           =   "frmAgendaDeRecebimento.frx":036E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image imgLogo 
      Height          =   420
      Left            =   885
      Top             =   6960
      Visible         =   0   'False
      Width           =   420
   End
End
Attribute VB_Name = "frmAgendaDeRecebimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rdoAgenda As New ADODB.Recordset
Dim wHora As Integer
Dim rdoForn As rdoResultset
Dim DiaAnterior As String
Dim nota As String
Dim Valor As String
Dim volume As String
Dim cont As Integer
Dim aux As String
Dim seq As Long
Dim Total As Double
Dim ProximoDia As String
Dim wWhere As String
Private Sub cmbGravar_Click()
 
    If txtData.Text = "" Then
        MsgBox "Preencha o campo data!", vbCritical, "ATENÇÃO"
        txtData.SetFocus
        Exit Sub
    End If
   ' If txtHora.Text = "" Then
   '     MsgBox "Preencha o campo hora!", vbCritical, "ATENÇÃO"
   '     txtHora.SetFocus
   '     Exit Sub
   ' End If
   ' If txtCodigo.Text = "" Then
   '     MsgBox "Preencha o campo código!", vbCritical, "ATENÇÃO"
   '     txtCodigo.SetFocus
   '     Exit Sub
   ' End If
    If txtFornecedor.Text = "" Then
        MsgBox "Preencha o campo fornecedor!", vbCritical, "ATENÇÃO"
        txtFornecedor.SetFocus
        Exit Sub
    End If
  '  If txtNotaFiscal.Text = "" Then
  '      MsgBox "Preencha o campo nota fiscal!", vbCritical, "ATENÇÃO"
  '      txtNotaFiscal.SetFocus
  '      Exit Sub
  '  End If
  '  If txtValor.Text = "" Then
  '      MsgBox "Preencha o campo valor!", vbCritical, "ATENÇÃO"
  '      txtValor.SetFocus
  '      Exit Sub
  '  End If
  '  If txtVolume.Text = "" Then
  '      MsgBox "Preencha o campo volume!", vbCritical, "ATENÇÃO"
  '      txtVolume.SetFocus
  '      Exit Sub
  '  End If
  If grdNotas.Rows = 1 Then
    MsgBox "Insira no mínimo uma nota fiscal!", vbCritical, "ATENÇÃO"
    txtNotaFiscal.SetFocus
    Exit Sub
  End If
    If txtTransportadora.Text = "" Then
        MsgBox "Preencha o campo transportadora!", vbCritical, "ATENÇÃO"
        txtTransportadora.SetFocus
        Exit Sub
    End If
    If txtNome.Text = "" Then
        MsgBox "Preencha o campo nome!", vbCritical, "ATENÇÃO"
        txtNome.SetFocus
        Exit Sub
    End If
    If txtTelefone.Text = "" Then
        MsgBox "Preencha o campo telefone!", vbCritical, "ATENÇÃO"
        txtTelefone.SetFocus
        Exit Sub
    End If
   ' If txtConfirmacao.Text = "" Then
   '     MsgBox "Preencha o campo confirmação!", vbCritical, "ATENÇÃO"
   '     txtConfirmacao.SetFocus
   '     Exit Sub
   ' End If
   
   ' If txtData.Text < Date Then
    If Format(txtData.Text, "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
        MsgBox "Data inferior a data atual!", vbCritical, "ATENÇÃO"
        txtData.SetFocus
        Exit Sub
    End If
    
    If cmbGravar.Caption = "Gravar" Then
    
     sql = "Select * from agendaderecebimento where ar_data = '" & Format(txtData, "yyyy/mm/dd") & "' and ar_hora = '" & cmbHora.Text & "' order by ar_hora"
     'Set rdoAgenda = ADO_Cn_CDLocal.OpenResultset(SQL)
     rdoAgenda.CursorLocation = adUseClient
     rdoAgenda.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
     
     If Not rdoAgenda.EOF Then
        If MsgBox("Já existe agendamento para esta data e horário!" & vbCrLf & "Deseja acrescentar uma nota?", vbQuestion + vbYesNo, "Agenda de Recebimento") = vbNo Then
            'MsgBox "Já existe agendamento para esta data e horário!", vbCritical, "ATENÇÃO"
            txtData.SelStart = 0
            txtData.SelLength = Len(txtData.Text)
            rdoAgenda.Close
        Exit Sub
        Else
            If lblSequencia.Caption <> "" Then
                sql = "Delete agendaderecebimento where ar_data = '" & Format(txtData.Text, "yyyy/mm/dd") & _
                      "' and ar_hora = '" & cmbHora.Text & "'"
                ADO_Cn_CDLocal.Execute (sql)
            End If
           
        End If
     End If
    rdoAgenda.Close
   
    cont = 1
    
    Do While cont <= grdNotas.Rows - 1
        If cont = 1 Then
            sql = "Insert into agendaderecebimento (ar_data, ar_hora, ar_fornecedor, ar_notafiscal, ar_valor, ar_volume, ar_transportadora, ar_nome, ar_telefone,ar_DataCadastro, ar_observacao) " _
                & "  values ('" & Format(txtData.Text, "yyyy/mm/dd") & "', '" & cmbHora.Text & "', '" & txtCodigo.Text & " " & txtFornecedor.Text & "', '" _
                & grdNotas.TextMatrix(cont, 0) & "', '" & ConverteVirgula(grdNotas.TextMatrix(cont, 1)) & "', '" & grdNotas.TextMatrix(cont, 2) & "', '" & txtTransportadora.Text & "', '" & txtNome.Text & "', '" _
                & txtTelefone.Text & "','" & Format(Date, "yyyy/mm/dd") & "', '" & txtObservacao.Text & "')"
            ADO_Cn_CDLocal.Execute (sql)
            
            sql = "Select max(ar_sequencia) as sequencia from agendaderecebimento"
                'Set rdoAgenda = ADO_Cn_CDLocal.OpenResultset(SQL)
                rdoAgenda.CursorLocation = adUseClient
                rdoAgenda.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
            seq = rdoAgenda("sequencia")
            rdoAgenda.Close
        End If
        If cont > 1 And cont <= 10 Then
            sql = "Update agendaderecebimento set ar_notafiscal" & cont - 1 & " = '" & grdNotas.TextMatrix(cont, 0) & _
                    "', ar_valor" & cont - 1 & " = '" & ConverteVirgula(grdNotas.TextMatrix(cont, 1)) & _
                    "', ar_volume" & cont - 1 & " = '" & grdNotas.TextMatrix(cont, 2) & _
                    "' where ar_sequencia = " & seq & ""
            ADO_Cn_CDLocal.Execute (sql)
        End If
        cont = cont + 1
    Loop
    MsgBox "Registro gravado com sucesso!", vbInformation, "Agenda de Recebimento"
    
    
    Else
        cont = 1
        If lblSequencia.Caption = "" Then
            MsgBox "Nenhuma nota foi selecionada!", vbCritical, "ATENÇÃO"
            Exit Sub
        End If
       
        If grdNotas.row = 1 Then
        sql = "Update agendaderecebimento set ar_data = '" & Format(txtData.Text, "yyyy/mm/dd") & _
                "', ar_hora = '" & cmbHora.Text & "', ar_fornecedor = '" & txtCodigo.Text & " " & txtFornecedor.Text & _
                "', ar_notafiscal = '" & txtNotaFiscal.Text & _
                "', ar_valor = '" & ConverteVirgula(txtValor.Text) & _
                "', ar_volume = '" & txtVolume.Text & "', ar_transportadora = '" & txtTransportadora.Text & _
                "', ar_nome = '" & txtNome.Text & "', ar_telefone = '" & txtTelefone.Text & _
                "', ar_observacao = '" & txtObservacao.Text & _
                "' where ar_sequencia = '" & lblSequencia.Caption & "'"
                ADO_Cn_CDLocal.Execute (sql)
        Else
        sql = "Update agendaderecebimento set ar_data = '" & Format(txtData.Text, "yyyy/mm/dd") & _
                "', ar_hora = '" & cmbHora.Text & "', ar_fornecedor = '" & txtCodigo.Text & " " & txtFornecedor.Text & _
                "', ar_notafiscal" & Trim(grdNotas.row - 1) & " = '" & txtNotaFiscal.Text & _
                "', ar_valor" & grdNotas.row - 1 & " = '" & ConverteVirgula(txtValor.Text) & _
                "', ar_volume" & grdNotas.row - 1 & " = '" & txtVolume.Text & "', ar_transportadora = '" & txtTransportadora.Text & _
                "', ar_nome = '" & txtNome.Text & "', ar_telefone = '" & txtTelefone.Text & _
                "', ar_observacao = '" & txtObservacao.Text & _
                "' where ar_sequencia = '" & lblSequencia.Caption & "'"
                ADO_Cn_CDLocal.Execute (sql)
        End If
        MsgBox "Registro alterado com sucesso!", vbInformation, "Agenda de Recebimento"
    End If
    
    Call LimpaTela
End Sub

Private Sub cmdEncerraNota_Click()

End Sub

Private Sub cmdLimpar_Click()
    Call LimpaTela
End Sub

Private Sub CmdMaior_Click()
If grdAgendaRecebimento <> "" Then
   LimpaGrid grdAgendaRecebimento
End If
    If txtPesqData.Text = "" Then
        txtPesqData.Text = Format(Date, "dd/mm/yyyy")
    End If
    ProximoDia = DateAdd("d", 1, txtPesqData.Text)
    ProximoDia = Format(ProximoDia, "dd/mm/yyyy")
    
    If ProximoDia = txtPesqData.Text Then
       ProximoDia = DateAdd("d", 1, ProximoDia)
       txtPesqData.Text = Format(ProximoDia, "dd/mm/yyyy")
    Else
       txtPesqData.Text = ProximoDia
    End If
    
    PesquisaAgenda
    
    
    
    Call LimpaCampos

End Sub

Private Sub CmdMenor_Click()

If grdAgendaRecebimento <> "" Then
   LimpaGrid grdAgendaRecebimento
End If
    
    DiaAnterior = DateAdd("d", -1, txtPesqData.Text)
    DiaAnterior = Format(DiaAnterior, "dd/mm/yyyy")
    
    If DiaAnterior = txtPesqData.Text Then
       DiaAnterior = DateAdd("d", -1, DiaAnterior)
       txtPesqData.Text = Format(DiaAnterior, "dd/mm/yyyy")
    Else
       txtPesqData.Text = DiaAnterior
    End If
   
   PesquisaAgenda
    
    Call LimpaCampos
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
sql = "Select * from agendaderecebimento order by ar_data, ar_hora"
        'Set rdoAgenda = ADO_Cn_CDLocal.OpenResultset(SQL)
        rdoAgenda.CursorLocation = adUseClient
        rdoAgenda.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
        If rdoAgenda.EOF Then
         '   MsgBox "Nenhum registro encontrado!", vbCritical, "ATENÇÃO"
            rdoAgenda.Close
        Else
            Call carregaGridAgenda
            
            wWhere = " "
            Call calculaTotal
            If IsNull(rdoAgenda("total")) = False Then
            lblTotalGeralAgendamento.Caption = "Total Geral Agendamento: R$ " & Format(rdoAgenda("total"), "###,###,###,#0.00")
         Else
            lblTotalGeralAgendamento.Caption = "Total Geral Agendamento: R$ 0,00"
         End If
         rdoAgenda.Close
        End If
End Sub

Private Sub Form_Load()

   carregarPosicaoTamanhoTela Me
    
   grdAgenda.MergeRow(0) = True
   grdAgenda.MergeRow(1) = True
   grdAgenda.MergeCol(0) = True
   grdAgenda.MergeCol(1) = True
   grdAgenda.MergeCol(2) = True
   grdAgenda.MergeCol(3) = True
   grdAgenda.MergeCol(4) = True
   grdAgenda.MergeCol(5) = True
   'grdAgenda.MergeCol(6) = True
   'grdAgenda.MergeCol(7) = True
   'grdAgenda.MergeCol(8) = True
   'grdAgenda.MergeCol(9) = True
   'grdAgenda.MergeCol(10) = True
   grdAgendaComplemento.MergeCol(0) = True
   grdAgendaComplemento.MergeCol(1) = True
   grdAgendaComplemento.MergeCol(2) = True
  ' grdAgendaComplemento.MergeCol(3) = True
   'grdAgendaComplemento.MergeCol(4) = True
   
   wHora = 7
     
   Do Until wHora > 17
      wHora = wHora + 1
      cmbHora.AddItem Format(wHora, "00") & "00"
      cmbHora.AddItem Format(wHora, "00") & "30"
  Loop
      cmbHora.AddItem "1900"
      cmbHora.ListIndex = 0
      
  txtData.Text = Format(Date, "dd/mm/yyyy")
  txtPesqData.Text = Format(Date, "dd/mm/yyyy")
     
   
   sql = "Select * from agendaderecebimento where ar_data = '" & Format(Date, "yyyy/mm/dd") & "' order by ar_hora"
   'Set rdoAgenda = ADO_Cn_CDLocal.OpenResultset(SQL)
   rdoAgenda.CursorLocation = adUseClient
   rdoAgenda.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
   
   Call carregaGridAgenda
   

   wWhere = "where ar_data = '" & Format((Date), "yyyy/mm/dd") & "' "
   Call calculaTotal
   If IsNull(rdoAgenda("total")) = False Then
   lblTotalGeralAgendamento.Caption = "Total Geral Agendamento: R$ " & Format(rdoAgenda("total"), "###,###,###,#0.00")
   Else
   lblTotalGeralAgendamento.Caption = "Total Geral Agendamento: R$ 0,00"
   End If
   rdoAgenda.Close
   
    lblSequencia.Caption = ""
    lblTotalAgendamento.Caption = ""
   ' lblTotalGeralAgendamento.Caption = ""
End Sub

Private Sub grdAgendaRecebimento_Click()
    Call carregaNotas
End Sub

Private Sub grdAgendaRecebimento_DblClick()
   If grdAgendaRecebimento.Rows > 1 Then
    If Trim(grdAgendaRecebimento.TextMatrix(grdAgendaRecebimento.row, 0)) <> "" And Trim(grdAgendaRecebimento.TextMatrix(grdAgendaRecebimento.row, 0)) <> "Data" Then
     If MsgBox("Deseja excluir o registro?", vbQuestion + vbYesNo, "Agenda de Recebimento") = vbYes Then
        sql = "Delete agendaderecebimento where ar_data = '" & Format(grdAgendaRecebimento.TextMatrix(grdAgendaRecebimento.row, 0), "yyyy/mm/dd") & "' and ar_hora = '" & grdAgendaRecebimento.TextMatrix(grdAgendaRecebimento.row, 1) & "' "
        ADO_Cn_CDLocal.Execute (sql)
        
        sql = "Select * from agendaderecebimento where ar_data = '" & Format(txtPesqData.Text, "yyyy/mm/dd") & "'"
        'Set rdoAgenda = ADO_Cn_CDLocal.OpenResultset(SQL)
        rdoAgenda.CursorLocation = adUseClient
        rdoAgenda.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        LimpaGrid grdAgendaRecebimento
            
        Call carregaGridAgenda
        
        
        txtData.Text = Format(Date, "dd/mm/yyyy")
        txtPesqData.Text = Format(Date, "dd/mm/yyyy")
        
        Call LimpaTela
        
     End If
    End If
   End If
End Sub

Private Sub grdAgendaRecebimento_EnterCell()
' If Trim(grdAgendaRecebimento.TextMatrix(grdAgendaRecebimento.Row, 7)) <> "" And Trim(grdAgendaRecebimento.TextMatrix(grdAgendaRecebimento.Row, 7)) <> "Nome" Then
    txtPesqData.Text = grdAgendaRecebimento.TextMatrix(grdAgendaRecebimento.row, 0)
' End If
    Call carregaNotas
    
    cmbGravar.Caption = "Editar"
End Sub

Private Sub grdNotas_DblClick()
    If grdNotas.Rows > 1 Then
       If Trim(grdNotas.TextMatrix(grdNotas.row, 0)) <> "Nota Fiscal" Then
          If grdNotas.TextMatrix(grdNotas.row, 3) <> "" Then
            If grdNotas.Rows - 1 = 1 Then
            If MsgBox("Deseja excluir o agendamento?", vbQuestion + vbYesNo, "Agenda de Recebimento") = vbYes Then
                sql = "Delete agendaderecebimento where ar_sequencia = '" & grdNotas.TextMatrix(grdNotas.row, 3) & "' "
                ADO_Cn_CDLocal.Execute (sql)
                MsgBox "Agendamento excluido com sucesso!", vbInformation, "Agenda de Recebimento"
                Call LimpaTela
            End If
            Else
            If MsgBox("Deseja excluir a nota?", vbQuestion + vbYesNo, "Agenda de Recebimento") = vbYes Then
                Call verificaNotas
                    sql = "Update agendaderecebimento set ar_notafiscal" & aux & " = 0, " _
                    & " ar_valor" & aux & " = 0, ar_volume" & aux & " = 0 " _
                    & " where ar_sequencia = '" & grdNotas.TextMatrix(grdNotas.row, 3) & "'"
                    ADO_Cn_CDLocal.Execute (sql)
                
                MsgBox "Nota excluida com sucesso!", vbInformation, "Agenda de Recebimento"
                Call carregaNotas
                txtNotaFiscal.Text = ""
                txtValor.Text = ""
                txtVolume.Text = ""
            End If
            End If
            wWhere = "where ar_data = '" & Format(txtData.Text, "yyyy/mm/dd") & "' "
            Call calculaTotal
            lblTotalGeralAgendamento.Caption = "Total Agendamento: R$ " & Format(rdoAgenda("total"), "###,###,###,#0.00")
            rdoAgenda.Close
          End If
      End If
    End If
End Sub

Private Sub grdNotas_EnterCell()
    If grdNotas.TextMatrix(grdNotas.row, 0) = "Nota Fiscal" Then
        Exit Sub
    End If
    
    txtNotaFiscal.Text = Trim(grdNotas.TextMatrix(grdNotas.row, 0))
    txtValor.Text = Trim(grdNotas.TextMatrix(grdNotas.row, 1))
    txtVolume.Text = Trim(grdNotas.TextMatrix(grdNotas.row, 2))
    
    If Trim(grdNotas.TextMatrix(grdNotas.row, 3)) <> "" Then
    lblSequencia.Caption = Trim(grdNotas.TextMatrix(grdNotas.row, 3))
    End If
End Sub

Private Sub txtCodigo_GotFocus()
txtCodigo.SelStart = 0
txtCodigo.SelLength = Len(txtCodigo.Text)
           
End Sub

Private Sub txtCodigo_LostFocus()
    
    
    If Trim(txtCodigo.Text) <> "" Then
        
        If IsNumeric(txtCodigo.Text) = False Then
           txtCodigo.SelStart = 0
           txtCodigo.SelLength = Len(txtCodigo.Text)
           txtCodigo.SetFocus
           Exit Sub
        End If
        
        txtFornecedor.Locked = True
        sql = "Select * from fornecedor where fo_codigofornecedor = '" & txtCodigo.Text & "'"
        'Set rdoAgenda = ADO_Cn_CDLocal.OpenResultset(SQL)
        rdoAgenda.CursorLocation = adUseClient
        rdoAgenda.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
        If Not rdoAgenda.EOF Then
            txtFornecedor = Trim(rdoAgenda("fo_nomefantasia"))
        End If
        
        rdoAgenda.Close
    Else
         
        txtFornecedor.Locked = False
    End If
End Sub

Private Sub txtData_Change()
    'Colocando em formato de data no msk
     If Len(txtData.Text) = 2 Then
       txtData.Text = txtData.Text & "/"
       txtData.SelStart = 3
    ElseIf Len(txtData.Text) = 5 Then
       txtData.Text = txtData.Text & "/"
       txtData.SelStart = 6
    ElseIf Len(txtData) = 10 Then
        'mskDataTermino.SetFocus
    End If
End Sub

Private Sub txtData_GotFocus()
txtData.SelStart = 0
txtData.SelLength = Len(txtData.Text)
End Sub

Private Sub TxtData_LostFocus()
   sql = "Select * from agendaderecebimento where ar_data = '" & Format(txtData, "yyyy/mm/dd") & "' order by ar_hora"
   rdoAgenda.CursorLocation = adUseClient
   rdoAgenda.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
   'Set rdoAgenda = ADO_Cn_CDLocal.OpenResultset(SQL)
   
   Call carregaGridAgenda
   
   wWhere = "where ar_data = '" & Format((txtData.Text), "yyyy/mm/dd") & "'"
    Call calculaTotal
     If IsNull(rdoAgenda("total")) = False Then
    lblTotalGeralAgendamento.Caption = "Total Geral Agendamento: R$ " & Format(rdoAgenda("total"), "###,###,###,#0.00")
    Else
    lblTotalGeralAgendamento.Caption = "Total Geral Agendamento: R$ 0,00"
    End If
    rdoAgenda.Close
End Sub

Private Sub txtFornecedor_GotFocus()
txtFornecedor.SelStart = 0
txtFornecedor.SelLength = Len(txtFornecedor.Text)
End Sub


Private Sub txtFornecedor_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 192 Then
         MsgBox "Caracter inválido", vbCritical
         txtFornecedor.Text = ""
      End If
End Sub

Private Sub txtNome_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 192 Then
         MsgBox "Caracter inválido", vbCritical
         txtNome.Text = ""
      End If
End Sub

Private Sub txtNotaFiscal_KeyDown(KeyCode As Integer, Shift As Integer)
  
      If KeyCode = 192 Then
         MsgBox "Caracter inválido", vbCritical
         txtNotaFiscal.Text = ""
      End If
      
      If KeyCode = 13 Then
        txtValor.SetFocus
      End If
     
End Sub



'Private Sub txtHora_Change()

'End Sub

Private Sub txtPesqData_Change()
    'Colocando em formato de data no msk
     If Len(txtPesqData.Text) = 2 Then
       txtPesqData.Text = txtPesqData.Text & "/"
       txtPesqData.SelStart = 3
    ElseIf Len(txtPesqData.Text) = 5 Then
       txtPesqData.Text = txtPesqData.Text & "/"
       txtPesqData.SelStart = 6
    ElseIf Len(txtPesqData) = 10 Then
        'mskDataTermino.SetFocus
    End If
End Sub

Private Sub txtPesqData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        sql = "Select * from agendaderecebimento where ar_data = '" & Format(txtPesqData.Text, "yyyy/mm/dd") & "'"
        'Set rdoAgenda = ADO_Cn_CDLocal.OpenResultset(SQL)
        rdoAgenda.CursorLocation = adUseClient
        rdoAgenda.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
        If rdoAgenda.EOF Then
            MsgBox "Nenhum registro encontrado!", vbCritical, "ATENÇÃO"
            rdoAgenda.Close
        Else
            Call carregaGridAgenda
        End If
    End If
End Sub
Private Sub PesquisaAgenda()
 sql = "Select * from agendaderecebimento where ar_data = '" & Format(txtPesqData.Text, "yyyy/mm/dd") & "'"
        'Set rdoAgenda = ADO_Cn_CDLocal.OpenResultset(SQL)
        rdoAgenda.CursorLocation = adUseClient
        rdoAgenda.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
        If rdoAgenda.EOF Then
         '   MsgBox "Nenhum registro encontrado!", vbCritical, "ATENÇÃO"
            rdoAgenda.Close
        Else
            Call carregaGridAgenda
        End If
End Sub


Private Sub carregaGridAgenda()
    grdAgendaRecebimento.Rows = 1
    If Not rdoAgenda.EOF Then
            Do While Not rdoAgenda.EOF
                   grdAgendaRecebimento.AddItem Trim(rdoAgenda("ar_data")) & Chr(9) & _
                                                Trim(rdoAgenda("ar_hora")) & Chr(9) & _
                                                Trim(rdoAgenda("ar_fornecedor")) & Chr(9) & _
                                                Trim(rdoAgenda("ar_observacao")) & Chr(9) & _
                                                Format(rdoAgenda("ar_DataCadastro"), "dd/mm/yyyy")

            rdoAgenda.MoveNext
            Loop
        End If
    rdoAgenda.Close
    
    
End Sub





Private Sub txtTelefone_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 192 Then
         MsgBox "Caracter inválido", vbCritical
         txtTelefone.Text = ""
      End If
End Sub


Private Sub txtTransportadora_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 192 Then
         MsgBox "Caracter inválido", vbCritical
         txtTransportadora.Text = ""
      End If
End Sub

Private Sub txtValor_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 192 Then
         MsgBox "Caracter inválido", vbCritical
         txtValor.Text = ""
      End If
      
      If KeyCode = 13 Then
        txtVolume.SetFocus
      End If
     
End Sub

Private Sub txtvalor_LostFocus()
    txtValor.Text = Format(txtValor.Text, "0.00")
End Sub



Private Sub txtVolume_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 192 Then
         MsgBox "Caracter inválido", vbCritical
         txtVolume.Text = ""
      End If
     
     If KeyCode = 13 Then
        
       ' lblSequencia.Caption = ""
        cmbGravar.Caption = "Gravar"
        If (grdNotas.Rows) <= 10 Then
            If txtNotaFiscal.Text = "" Then
                MsgBox "Preencha o campo nota fiscal!", vbCritical, "ATENÇÃO"
                txtNotaFiscal.SetFocus
                Exit Sub
            End If
            If txtValor.Text = "" Then
                MsgBox "Preencha o campo valor!", vbCritical, "ATENÇÃO"
                txtValor.SetFocus
                Exit Sub
            End If
            If txtVolume.Text = "" Then
                MsgBox "Preencha o campo volume!", vbCritical, "ATENÇÃO"
                txtVolume.SetFocus
                Exit Sub
            End If
            
          '  If lblSequencia.Caption = "" Then
            grdNotas.AddItem txtNotaFiscal.Text & Chr(9) & _
                         txtValor.Text & Chr(9) & _
                         txtVolume.Text
          '  End If
        Else
            MsgBox "Não é permitida a inserção de mais de 10 notas!", vbCritical, "ATENÇÃO"
        End If
        txtNotaFiscal.Text = ""
        txtValor.Text = ""
        txtVolume.Text = ""
        txtNotaFiscal.SetFocus
     End If
End Sub

Public Sub carregaNotas()
    If grdAgendaRecebimento.TextMatrix(grdAgendaRecebimento.row, 0) = "Data" Then
        grdAgendaRecebimento.row = 1
    End If
    sql = "Select * from agendaderecebimento where ar_data = '" & Format(grdAgendaRecebimento.TextMatrix(grdAgendaRecebimento.row, 0), "yyyy/mm/dd") & "' and ar_hora = '" & grdAgendaRecebimento.TextMatrix(grdAgendaRecebimento.row, 1) & "' "
    'Set rdoAgenda = ADO_Cn_CDLocal.OpenResultset(SQL)
    rdoAgenda.CursorLocation = adUseClient
    rdoAgenda.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    grdNotas.Rows = 1
    
    If Not rdoAgenda.EOF Then
        txtData.Text = rdoAgenda("ar_data")
        cmbHora.Text = Trim(rdoAgenda("ar_hora"))
        txtCodigo.Text = Trim(Mid(rdoAgenda("ar_fornecedor"), 1, InStr(rdoAgenda("ar_fornecedor"), " ")))
        txtFornecedor.Text = Trim(Mid(rdoAgenda("ar_fornecedor"), InStr(rdoAgenda("ar_fornecedor"), " ")))
        txtTransportadora.Text = Trim(rdoAgenda("ar_transportadora"))
        txtNome.Text = Trim(rdoAgenda("ar_nome"))
        txtTelefone.Text = Trim(rdoAgenda("ar_telefone"))
        txtObservacao.Text = Trim(rdoAgenda("ar_observacao"))
        lblSequencia.Caption = rdoAgenda("ar_sequencia")
        Do While Not rdoAgenda.EOF
            If rdoAgenda("ar_notafiscal") <> 0 Then
                 aux = ""
                 grdNotas.AddItem rdoAgenda("ar_notafiscal") & Chr(9) & _
                     Format(rdoAgenda("ar_valor"), "0.00") & Chr(9) & _
                     rdoAgenda("ar_volume") & Chr(9) & _
                     rdoAgenda("ar_sequencia")
                 aux = 0
            End If
            If rdoAgenda("ar_notafiscal1") <> 0 Then
                aux = 1
            End If
            If rdoAgenda("ar_notafiscal2") <> 0 Then
                aux = 2
            End If
            If rdoAgenda("ar_notafiscal3") <> 0 Then
                aux = 3
            End If
            If rdoAgenda("ar_notafiscal4") <> 0 Then
                aux = 4
            End If
            If rdoAgenda("ar_notafiscal5") <> 0 Then
                aux = 5
            End If
            If rdoAgenda("ar_notafiscal6") <> 0 Then
                aux = 6
            End If
            If rdoAgenda("ar_notafiscal7") <> 0 Then
                aux = 7
            End If
            If rdoAgenda("ar_notafiscal8") <> 0 Then
                aux = 8
            End If
            If rdoAgenda("ar_notafiscal8") <> 0 Then
                aux = 9
            End If
            
            
            cont = 1
            If aux <> "" Then
            Do While cont <= aux
              If rdoAgenda("ar_notafiscal" & cont) <> 0 Then
                grdNotas.AddItem rdoAgenda("ar_notafiscal" & cont) & Chr(9) & _
                     Format(rdoAgenda("ar_valor" & cont), "0.00") & Chr(9) & _
                     rdoAgenda("ar_volume" & cont) & Chr(9) & _
                     rdoAgenda("ar_sequencia")
               End If
                cont = cont + 1
            Loop
            End If
         rdoAgenda.MoveNext
        Loop
       
    End If
    
    rdoAgenda.Close
    
    wWhere = "where ar_data = '" & Format(grdAgendaRecebimento.TextMatrix(grdAgendaRecebimento.row, 0), "yyyy/mm/dd") & "' and ar_hora = '" & grdAgendaRecebimento.TextMatrix(grdAgendaRecebimento.row, 1) & "'"
    Call calculaTotal
    lblTotalAgendamento.Caption = "Total Agendamento: R$ " & Format(rdoAgenda("total"), "###,###,###,#0.00")
    rdoAgenda.Close
    
    
End Sub

Private Sub LimpaTela()
    LimpaGrid grdAgendaRecebimento
    
    If txtData.Text = Date Then
      sql = "Select * from agendaderecebimento where ar_data = '" & Format(Date, "yyyy/mm/dd") & "' order by ar_hora"
        'Set rdoAgenda = ADO_Cn_CDLocal.OpenResultset(SQL)
        rdoAgenda.CursorLocation = adUseClient
        rdoAgenda.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
        Call carregaGridAgenda
        
        wWhere = "where ar_data = '" & Format((Date), "yyyy/mm/dd") & "'"
        Call calculaTotal
         If IsNull(rdoAgenda("total")) = False Then
            lblTotalGeralAgendamento.Caption = "Total Geral Agendamento: R$ " & Format(rdoAgenda("total"), "###,###,###,#0.00")
         Else
            lblTotalGeralAgendamento.Caption = "Total Geral Agendamento: R$ 0,00"
         End If
         rdoAgenda.Close
    End If
    
    txtData.Text = Format(Date, "dd/mm/yyyy")
    txtPesqData.Text = Format(Date, "dd/mm/yyyy")
    'txtData.Text = ""
    txtPesqData.Text = Date
    cmbHora.ListIndex = 0
    txtCodigo.Text = ""
    txtFornecedor.Text = ""
    txtFornecedor.Locked = True
    txtNotaFiscal.Text = ""
    txtValor.Text = ""
    txtVolume.Text = ""
    txtTransportadora.Text = ""
    txtNome.Text = ""
    txtTelefone.Text = ""
    'txtConfirmacao.Text = ""
    
    grdNotas.Rows = 1
    txtObservacao.Text = ""
    lblSequencia.Caption = ""
    cmbGravar.Caption = "Gravar"
    lblTotalAgendamento.Caption = ""
    txtData.SetFocus
End Sub

Private Sub calculaTotal()
    sql = "Select sum(ar_valor) + sum(ar_valor1) + sum(ar_valor2) + sum(ar_valor3) + sum(ar_valor4) " _
              & " + sum(ar_valor5) + sum(ar_valor6) + sum(ar_valor7) + sum(ar_valor8) + sum(ar_valor9) as total " _
              & " from agendaderecebimento " _
              & " " & wWhere & " "
    'Set rdoAgenda = ADO_Cn_CDLocal.OpenResultset(SQL)
    rdoAgenda.CursorLocation = adUseClient
    rdoAgenda.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
End Sub

Private Sub LimpaCampos()
    txtData.Text = Format(Date, "dd/mm/yyyy")
    'txtData.Text = ""
    cmbHora.ListIndex = 0
    txtCodigo.Text = ""
    txtFornecedor.Text = ""
    txtFornecedor.Locked = True
    txtNotaFiscal.Text = ""
    txtValor.Text = ""
    txtVolume.Text = ""
    txtTransportadora.Text = ""
    txtNome.Text = ""
    txtTelefone.Text = ""
    'txtConfirmacao.Text = ""
    
    grdNotas.Rows = 1
    txtObservacao.Text = ""
    lblSequencia.Caption = ""
    cmbGravar.Caption = "Gravar"
    lblTotalAgendamento.Caption = ""
    
    If txtPesqData.Text = "" Then
        txtPesqData.Text = Format((Date), "dd/mm/yyyy")
    End If
    
    wWhere = "where ar_data = '" & Format((txtPesqData.Text), "yyyy/mm/dd") & "'"
        Call calculaTotal
        If IsNull(rdoAgenda("total")) = False Then
            lblTotalGeralAgendamento.Caption = "Total Geral Agendamento: R$ " & Format(rdoAgenda("total"), "###,###,###,#0.00")
         Else
            lblTotalGeralAgendamento.Caption = "Total Geral Agendamento: R$ 0,00"
         End If
         rdoAgenda.Close
End Sub

Private Sub verificaNotas()
    sql = "Select * from agendaderecebimento where ar_sequencia = '" & grdNotas.TextMatrix(grdNotas.row, 3) & "'"
        'Set rdoAgenda = ADO_Cn_CDLocal.OpenResultset(SQL)
        rdoAgenda.CursorLocation = adUseClient
        rdoAgenda.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If rdoAgenda("ar_notafiscal") = grdNotas.TextMatrix(grdNotas.row, 0) _
        And Format(rdoAgenda("ar_valor"), "###,###,###,#0.00") = grdNotas.TextMatrix(grdNotas.row, 1) _
        And rdoAgenda("ar_volume") = grdNotas.TextMatrix(grdNotas.row, 2) Then
        
       aux = ""
    ElseIf rdoAgenda("ar_notafiscal1") = grdNotas.TextMatrix(grdNotas.row, 0) _
        And Format(rdoAgenda("ar_valor1"), "###,###,###,#0.00") = grdNotas.TextMatrix(grdNotas.row, 1) _
        And rdoAgenda("ar_volume1") = grdNotas.TextMatrix(grdNotas.row, 2) Then
       aux = 1
    ElseIf rdoAgenda("ar_notafiscal2") = grdNotas.TextMatrix(grdNotas.row, 0) _
        And Format(rdoAgenda("ar_valor2"), "###,###,###,#0.00") = grdNotas.TextMatrix(grdNotas.row, 1) _
        And rdoAgenda("ar_volume2") = grdNotas.TextMatrix(grdNotas.row, 2) Then
       aux = 2
    ElseIf rdoAgenda("ar_notafiscal3") = grdNotas.TextMatrix(grdNotas.row, 0) _
        And Format(rdoAgenda("ar_valor3"), "###,###,###,#0.00") = grdNotas.TextMatrix(grdNotas.row, 1) _
        And rdoAgenda("ar_volume3") = grdNotas.TextMatrix(grdNotas.row, 2) Then
        aux = 3
    ElseIf rdoAgenda("ar_notafiscal4") = grdNotas.TextMatrix(grdNotas.row, 0) _
        And Format(rdoAgenda("ar_valor4"), "###,###,###,#0.00") = grdNotas.TextMatrix(grdNotas.row, 1) _
        And rdoAgenda("ar_volume4") = grdNotas.TextMatrix(grdNotas.row, 2) Then
        aux = 4
    ElseIf rdoAgenda("ar_notafiscal5") = grdNotas.TextMatrix(grdNotas.row, 0) _
        And Format(rdoAgenda("ar_valor5"), "###,###,###,#0.00") = grdNotas.TextMatrix(grdNotas.row, 1) _
        And rdoAgenda("ar_volume5") = grdNotas.TextMatrix(grdNotas.row, 2) Then
        aux = 5
    ElseIf rdoAgenda("ar_notafiscal6") = grdNotas.TextMatrix(grdNotas.row, 0) _
        And Format(rdoAgenda("ar_valor6"), "###,###,###,#0.00") = grdNotas.TextMatrix(grdNotas.row, 1) _
        And rdoAgenda("ar_volume6") = grdNotas.TextMatrix(grdNotas.row, 2) Then
        aux = 6
    ElseIf rdoAgenda("ar_notafiscal7") = grdNotas.TextMatrix(grdNotas.row, 0) _
        And Format(rdoAgenda("ar_valor7"), "###,###,###,#0.00") = grdNotas.TextMatrix(grdNotas.row, 1) _
        And rdoAgenda("ar_volume7") = grdNotas.TextMatrix(grdNotas.row, 2) Then
        aux = 7
    ElseIf rdoAgenda("ar_notafiscal8") = grdNotas.TextMatrix(grdNotas.row, 0) _
        And Format(rdoAgenda("ar_valor8"), "###,###,###,#0.00") = grdNotas.TextMatrix(grdNotas.row, 1) _
        And rdoAgenda("ar_volume8") = grdNotas.TextMatrix(grdNotas.row, 2) Then
        aux = 8
    ElseIf rdoAgenda("ar_notafiscal9") = grdNotas.TextMatrix(grdNotas.row, 0) _
        And Format(rdoAgenda("ar_valor9"), "###,###,###,#0.00") = grdNotas.TextMatrix(grdNotas.row, 1) _
        And rdoAgenda("ar_volume9") = grdNotas.TextMatrix(grdNotas.row, 2) Then
        aux = 9
    End If
    
    
    rdoAgenda.Close
End Sub
