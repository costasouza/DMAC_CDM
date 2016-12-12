VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7d.ocx"
Begin VB.Form frmPermissaoSistema 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Permissão de Usuários"
   ClientHeight    =   7470
   ClientLeft      =   2460
   ClientTop       =   2775
   ClientWidth     =   15195
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7470
   ScaleWidth      =   15195
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameNovoUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00393939&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   4725
      Left            =   540
      TabIndex        =   9
      Top             =   1815
      Visible         =   0   'False
      Width           =   3615
      Begin VB.TextBox txtSenha 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   540
         PasswordChar    =   "*"
         TabIndex        =   15
         Top             =   1875
         Width           =   2505
      End
      Begin VB.TextBox txtSenha2 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   540
         PasswordChar    =   "*"
         TabIndex        =   16
         Top             =   2235
         Width           =   2505
      End
      Begin VB.OptionButton optGravaAdministrador 
         Appearance      =   0  'Flat
         BackColor       =   &H00393939&
         Caption         =   "Administrador"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   540
         TabIndex        =   18
         Top             =   3435
         Width           =   1335
      End
      Begin VB.OptionButton optGravaComum 
         Appearance      =   0  'Flat
         BackColor       =   &H00393939&
         Caption         =   "Usuário Padrão"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   540
         TabIndex        =   17
         Top             =   3195
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.TextBox txtNome 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   540
         MaxLength       =   10
         TabIndex        =   13
         Top             =   1035
         Width           =   2505
      End
      Begin CentroDeDistribuicao.chameleonButton cmdGravarUsuario 
         Height          =   315
         Left            =   540
         TabIndex        =   19
         Top             =   3915
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         BTYPE           =   14
         TX              =   "Gravar"
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
         MICON           =   "frmPermissaoSistema.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CentroDeDistribuicao.chameleonButton cmdSairGravaUsuario 
         Height          =   315
         Left            =   1860
         TabIndex        =   20
         Top             =   3915
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         BTYPE           =   14
         TX              =   "Sair"
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
         MICON           =   "frmPermissaoSistema.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de usuário"
         ForeColor       =   &H00C0C0C0&
         Height          =   300
         Left            =   540
         TabIndex        =   14
         Top             =   2835
         Width           =   3885
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Senha"
         ForeColor       =   &H00C0C0C0&
         Height          =   300
         Left            =   540
         TabIndex        =   12
         Top             =   1635
         Width           =   3885
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Novo Usuário"
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
         Left            =   540
         TabIndex        =   11
         Top             =   435
         Width           =   2655
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome"
         ForeColor       =   &H00C0C0C0&
         Height          =   300
         Left            =   540
         TabIndex        =   10
         Top             =   795
         Width           =   1605
      End
   End
   Begin VB.Frame frameOpcoesUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   6450
      Left            =   3690
      TabIndex        =   3
      Top             =   150
      Width           =   11370
      Begin VB.OptionButton optAdministrador 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Administrador"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   360
         TabIndex        =   22
         Top             =   465
         Width           =   1575
      End
      Begin VB.OptionButton optComum 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Usuário Padrão"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   2040
         TabIndex        =   21
         Top             =   465
         Width           =   1695
      End
      Begin VSFlex7DAOCtl.VSFlexGrid grdPermitido 
         Height          =   5085
         Left            =   150
         TabIndex        =   4
         Top             =   1200
         Width           =   5220
         _cx             =   9208
         _cy             =   8969
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
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPermissaoSistema.frx":0038
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
      Begin VSFlex7DAOCtl.VSFlexGrid grdNaoPermitido 
         Height          =   5085
         Left            =   5985
         TabIndex        =   5
         Top             =   1200
         Width           =   5220
         _cx             =   9208
         _cy             =   8969
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
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPermissaoSistema.frx":0087
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
      Begin CentroDeDistribuicao.chameleonButton cmdDeletar 
         Height          =   420
         Left            =   8025
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   741
         BTYPE           =   14
         TX              =   "Deletar Usuário"
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
         MICON           =   "frmPermissaoSistema.frx":00D6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CentroDeDistribuicao.chameleonButton cmdAlterarSenha 
         Height          =   420
         Left            =   9600
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   741
         BTYPE           =   14
         TX              =   "Alterar Senha"
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
         MICON           =   "frmPermissaoSistema.frx":00F2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CentroDeDistribuicao.chameleonButton cmdAddAutorizado 
         Height          =   270
         Left            =   5460
         TabIndex        =   26
         Top             =   3345
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   476
         BTYPE           =   14
         TX              =   "<<"
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
         MICON           =   "frmPermissaoSistema.frx":010E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CentroDeDistribuicao.chameleonButton cmdRemoveAutorizado 
         Height          =   270
         Left            =   5460
         TabIndex        =   27
         Top             =   3690
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   476
         BTYPE           =   14
         TX              =   ">>"
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
         MICON           =   "frmPermissaoSistema.frx":012A
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
         Caption         =   "Tipo de Usuário"
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
         TabIndex        =   23
         Top             =   150
         Width           =   2655
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         Caption         =   "Não Autorizado"
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
         Left            =   5985
         TabIndex        =   8
         Top             =   915
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         Caption         =   "Autorizado"
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
         TabIndex        =   7
         Top             =   915
         Width           =   2655
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   180
      ScaleHeight     =   45
      ScaleWidth      =   14895
      TabIndex        =   0
      Top             =   6765
      Width           =   14900
   End
   Begin CentroDeDistribuicao.chameleonButton cmdRetorna 
      Height          =   510
      Left            =   13660
      TabIndex        =   1
      Top             =   6885
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
      MICON           =   "frmPermissaoSistema.frx":0146
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdUsuario 
      Height          =   6435
      Left            =   150
      TabIndex        =   2
      Top             =   150
      Width           =   3375
      _cx             =   5953
      _cy             =   11351
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
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPermissaoSistema.frx":0162
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
   Begin CentroDeDistribuicao.chameleonButton cmdNovoUsuario 
      Height          =   510
      Left            =   12225
      TabIndex        =   6
      Top             =   6885
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Novo Usuário"
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
      MICON           =   "frmPermissaoSistema.frx":01B4
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
Attribute VB_Name = "frmPermissaoSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String
Dim codigoUsuario As String

Private Sub frameNovoUsuarioAtivado(ativa As Boolean)
    If ativa Then
        frameNovoUsuario.Visible = True
        txtNome.SetFocus
    Else
        frameNovoUsuario.Visible = False
        limpaCadastroUsuario
        carregarGridUsuario

        grdUsuario.row = grdUsuario.Rows - 1
'        grdUsuario_Click
    End If
End Sub

Private Sub cmdAddAutorizado_Click()
    permitiUsuario True
End Sub

Private Sub cmdAlterarSenha_Click()
    ADO_Cn_CDLocal.Execute ("delete GLB_UsuariosSistema where us_codigo = " & codigoUsuario)
End Sub

Private Sub cmdDeletar_Click()
    
    Dim adoUsuarios As New ADODB.Recordset
    
    sql = "select count(*) admin from GLB_UsuariosSistema where us_nivelAcesso = 'A'"
    adoUsuarios.CursorLocation = adUseClient
    adoUsuarios.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If adoUsuarios("admin") < 2 Then
        MsgBox "Você não pode deletar todos os administradores", vbInformation, "Administradores"
    Else
        adoUsuarios.Close
        sql = "select rtrim(us_nome) as us_nome from GLB_UsuariosSistema where us_codigo = '" & codigoUsuario & "'"
        adoUsuarios.CursorLocation = adUseClient
        adoUsuarios.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
            
        If MsgBox("Deseja deleta o usuário " & adoUsuarios("us_nome") & " ?", vbQuestion + vbYesNo, "Deleta Usuário") = 6 Then
            ADO_Cn_CDLocal.Execute ("delete GLB_UsuariosSistema where us_codigo = " & codigoUsuario)
            carregarGridUsuario
        End If
        
    End If
    
    adoUsuarios.Close
    
End Sub

Private Sub cmdRemoveAutorizado_Click()
    permitiUsuario False
End Sub

Private Sub Form_Load()

carregarPosicaoTamanhoTela Me
carregarPosicaoFrame frameNovoUsuario
'JanelaTOP Me
     
   carregarGridUsuario
       
End Sub

Private Sub cmbUsuarios_LostFocus()
    'carregarGridPermitido
End Sub

Private Sub cmdGravarUsuario_Click()

    Dim adoUsuarios As New ADODB.Recordset
         
    If txtNome <> "" And txtSenha <> "" Then
        If txtSenha = txtSenha2 Then
            
            sql = "select rtrim(us_nome) as us_nome from GLB_UsuariosSistema where us_nome = '" & txtNome & "'"
            adoUsuarios.CursorLocation = adUseClient
            adoUsuarios.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
                
            If adoUsuarios.EOF Then
                sql = "exec SP_GLB_Grava_Usuarios_Sistema '" & txtNome & "', '" & txtSenha & "', "
                If optGravaAdministrador Then
                    sql = sql & "'A'"
                Else
                    sql = sql & "'C'"
                End If
                ADO_Cn_CDLocal.Execute sql
                
                frameNovoUsuarioAtivado False

            Else
                MsgBox "Nome de usuário já existente", vbExclamation, "Nome"
                txtNome.SetFocus
            End If
        Else
            MsgBox "Senhas não batem", vbExclamation, "Senha"
            txtSenha.SetFocus
        End If
    Else
        MsgBox "Preencha todos os campos", vbExclamation, "Campo Vazio"
        txtNome.SetFocus
    End If
End Sub

Private Sub cmdNovoUsuario_Click()
     frameNovoUsuarioAtivado True
End Sub

Private Sub cmdRetorna_Click()
    Unload Me
End Sub

Private Sub cmdSairGravaUsuario_Click()
    frameNovoUsuarioAtivado False
End Sub


Private Sub carregarGridUsuario()
     Dim adoUsuarios As New ADODB.Recordset
     
     sql = "Select us_codigo,rtrim(us_nome) as us_nome from GLB_UsuariosSistema order by us_codigo"
        adoUsuarios.CursorLocation = adUseClient
        adoUsuarios.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
     grdUsuario.Rows = 1
     If Not adoUsuarios.EOF Then
        Do While Not adoUsuarios.EOF
            grdUsuario.AddItem Format(adoUsuarios("us_codigo"), "000") & Chr(9) & adoUsuarios("us_nome")
        adoUsuarios.MoveNext
       Loop
    End If
  
    adoUsuarios.Close
End Sub


Private Sub carregarGridPermitido(codigoUsuario As String, administrador As Boolean)
    
     Dim adoTelas As New ADODB.Recordset
           
     If administrador Then
        sql = "select msi_codigo, msi_descricao from glb_menusistema order by msi_codigo"
     Else
        sql = "select distinct msi_codigo, msi_descricao from GLB_PermissaoSistema, GLB_MenuSistema " & _
        "where msi_codigo = ps_nomeTela and ps_codigoUsuario = " & codigoUsuario & " order by msi_codigo"
     End If
     
     adoTelas.CursorLocation = adUseClient
     adoTelas.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
     grdPermitido.Rows = 1
     If Not adoTelas.EOF Then
        Do While Not adoTelas.EOF
            If Mid(adoTelas("msi_codigo"), 3, 2) > "00" Then
                grdPermitido.AddItem Format(adoTelas("msi_codigo"), "000000") & Chr(9) & "  " & adoTelas("msi_descricao")
            ElseIf Mid(adoTelas("msi_codigo"), 5, 2) > "00" Then
                grdPermitido.AddItem Format(adoTelas("msi_codigo"), "000000") & Chr(9) & " " & adoTelas("msi_descricao")
            Else
                grdPermitido.AddItem Format(adoTelas("msi_codigo"), "000000") & Chr(9) & "" & adoTelas("msi_descricao")
            End If
        adoTelas.MoveNext
       Loop
    Else
        grdPermitido.Rows = 1
    End If
    
    adoTelas.Close

End Sub

Private Sub carregarGridNaoPermitido(codigoUsuario As String, administrador As Boolean)
    
     Dim adoTelas As New ADODB.Recordset
           
     grdNaoPermitido.Rows = 1
     If administrador Then
        Exit Sub
     Else
        sql = "select msi_codigo, msi_descricao from GLB_MenuSistema where not exists " & _
        "(select ps_codigoUsuario from GLB_PermissaoSistema where ps_codigoUsuario = " & codigoUsuario & _
        " and msi_codigo = ltrim(rtrim(ps_nomeTela))) order by msi_codigo"
     End If
     
     adoTelas.CursorLocation = adUseClient
     adoTelas.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
     If Not adoTelas.EOF Then
        Do While Not adoTelas.EOF
            If Mid(adoTelas("msi_codigo"), 3, 2) > "00" Then
                grdNaoPermitido.AddItem Format(adoTelas("msi_codigo"), "000000") & Chr(9) & "  " & adoTelas("msi_descricao")
            ElseIf Mid(adoTelas("msi_codigo"), 5, 2) > "00" Then
                grdNaoPermitido.AddItem Format(adoTelas("msi_codigo"), "000000") & Chr(9) & " " & adoTelas("msi_descricao")
            Else
                grdNaoPermitido.AddItem Format(adoTelas("msi_codigo"), "000000") & Chr(9) & "" & adoTelas("msi_descricao")
            End If
        adoTelas.MoveNext
       Loop
    Else
        grdNaoPermitido.Rows = 1
    End If
    
    adoTelas.Close

End Sub


Private Sub Frame1_DragDrop(Source As Control, x As Single, y As Single)

End Sub

Private Sub grdPermitido_DblClick()
    permitiUsuario False
End Sub

Public Sub permitiUsuario(permiti As Boolean)
    If permiti Then
        autorizaUsuario codigoUsuario, Trim(grdNaoPermitido.TextMatrix(grdNaoPermitido.row, 0))
        atualizaTabelas
    Else
        If Not optAdministrador Then
            desautorizaUsuario codigoUsuario, Trim(grdPermitido.TextMatrix(grdPermitido.row, 0))
            atualizaTabelas
        End If
    End If
End Sub

Private Sub grdNaoPermitido_DblClick()
    permitiUsuario True
End Sub

Private Sub atualizaTabelas()
    carregarGridPermitido codigoUsuario, usuarioAdministrador
    carregarGridNaoPermitido codigoUsuario, usuarioAdministrador
End Sub

Private Sub carregarTipoUsuario(codigoUsuario As String)
     Dim adoUsuarios As New ADODB.Recordset
     
     sql = "Select us_nivelAcesso from GLB_UsuariosSistema where us_codigo = " & codigoUsuario
     adoUsuarios.CursorLocation = adUseClient
     adoUsuarios.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
     If Mid(adoUsuarios("us_nivelAcesso"), 1, 1) = "A" Then
        optAdministrador.Value = True
     Else
        optComum.Value = True
     End If
  
    adoUsuarios.Close
End Sub

Private Function usuarioAdministrador() As Boolean
    If optAdministrador.Value = True Then
        usuarioAdministrador = True
    Else
        usuarioAdministrador = False
    End If
End Function

Private Sub autorizaUsuario(codigoUsuario As String, codigoTela As String)
    ADO_Cn_CDLocal.Execute ("exec SP_GLB_Grava_Permissao_Sistema '" & codigoUsuario & "', '" & codigoTela & "'")
End Sub

Private Sub desautorizaUsuario(codigoUsuario As String, codigoTela As String)
    ADO_Cn_CDLocal.Execute ("exec SP_GLB_Excluir_Permissao_Sistema '" & codigoUsuario & "', '" & codigoTela & "'")
End Sub

Private Sub deletaUsuario()
    If MsgBox("Deseja deleta o usuário " & grdUsuario.TextMatrix(grdUsuario.row, 1) & "?", vbQuestion + vbYesNo, "Deleta usuário") = vbYes Then
        sql = "delete GLB_UsuariosSistema where us_codigo = " & grdUsuario.TextMatrix(grdUsuario.row, 0) & " and us_codigo <> 1"
        ADO_Cn_CDLocal.Execute sql
        carregarGridUsuario
    End If
End Sub

Private Sub grdUsuario_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        deletaUsuario
    End If
End Sub

Private Sub optAdministrador_Click()
    ADO_Cn_CDLocal.Execute ("exec SP_GLB_Altera_Nivel_Acesso_Usuarios '" & codigoUsuario & "', 'A'")
    atualizaTabelas
End Sub

Private Sub optComum_Click()
    ADO_Cn_CDLocal.Execute ("exec SP_GLB_Altera_Nivel_Acesso_Usuarios '" & codigoUsuario & "', 'C'")
    atualizaTabelas
End Sub

Private Sub txtNome_GotFocus()
    campoSelecionadoComCaracter txtNome
End Sub

Private Sub txtSenha_GotFocus()
    campoSelecionadoComCaracter txtSenha
End Sub

Private Sub txtSenha2_GotFocus()
    campoSelecionadoComCaracter txtSenha2
End Sub

Private Sub limpaCadastroUsuario()
    txtNome = ""
    txtSenha = ""
    txtSenha2 = ""
    optGravaComum.Value = True
End Sub

Private Sub grdUsuario_EnterCell()
    If Trim(grdUsuario.TextMatrix(grdUsuario.row, 0)) <> "" And Trim(grdUsuario.TextMatrix(grdUsuario.row, 0)) <> "Código" Then
        frameOpcoesUsuario.Enabled = True
        codigoUsuario = Trim(grdUsuario.TextMatrix(grdUsuario.row, 0))
        carregarTipoUsuario codigoUsuario
        atualizaTabelas
    End If
End Sub

Private Sub LimpaTela()

    

End Sub


