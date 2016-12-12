VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmTransportadora 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Transportadoras"
   ClientHeight    =   7470
   ClientLeft      =   0
   ClientTop       =   450
   ClientWidth     =   15675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   15675
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraMunicipio 
      BackColor       =   &H00393939&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   1575
      Left            =   1800
      TabIndex        =   41
      Top             =   4920
      Visible         =   0   'False
      Width           =   6735
      Begin VSFlex7DAOCtl.VSFlexGrid grdMunicipio 
         Height          =   1545
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Width           =   6735
         _cx             =   11880
         _cy             =   2725
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
         Rows            =   0
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmTransportadora.frx":0000
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
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   5263440
         ForeColorFrozen =   4210752
         WallPaperAlignment=   9
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   135
            Left            =   480
            TabIndex        =   43
            Top             =   1680
            Visible         =   0   'False
            Width           =   375
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   975
      Left            =   120
      TabIndex        =   37
      Top             =   5400
      Width           =   11415
      Begin VB.TextBox txtEmail 
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
         Left            =   6000
         TabIndex        =   14
         Top             =   360
         Width           =   5295
      End
      Begin VB.TextBox txtContato 
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
         Left            =   2520
         TabIndex        =   13
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtTelefone 
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
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   6000
         TabIndex        =   40
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblContato 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contato"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   2520
         TabIndex        =   39
         Top             =   120
         Width           =   555
      End
      Begin VB.Label lblTelefone 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   120
         Width           =   630
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2415
      Left            =   120
      TabIndex        =   29
      Top             =   2760
      Width           =   11415
      Begin VB.TextBox txtMunicipio 
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
         Left            =   1680
         TabIndex        =   10
         Top             =   1800
         Width           =   6735
      End
      Begin VB.TextBox txtComplemento 
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
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   6255
      End
      Begin VB.ComboBox cmbUF 
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
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtCep 
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
         Left            =   8520
         TabIndex        =   11
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox txtNumero 
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
         Left            =   9960
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtBairro 
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
         Left            =   6480
         TabIndex        =   8
         Top             =   1080
         Width           =   4815
      End
      Begin VB.TextBox txtEndereco 
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
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   9735
      End
      Begin VB.Label lblComplemento 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   840
         Width           =   960
      End
      Begin VB.Label lblMunicipio 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Municipio"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   1680
         TabIndex        =   35
         Top             =   1560
         Width           =   675
      End
      Begin VB.Label lblUF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UF"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   1560
         Width           =   210
      End
      Begin VB.Label lblCEP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CEP"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   8520
         TabIndex        =   33
         Top             =   1560
         Width           =   315
      End
      Begin VB.Label lblNumero 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   9960
         TabIndex        =   32
         Top             =   120
         Width           =   555
      End
      Begin VB.Label lblBairro 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   6480
         TabIndex        =   31
         Top             =   840
         Width           =   405
      End
      Begin VB.Label lblEndereco 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   690
      End
   End
   Begin VB.TextBox txtFiltro 
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
      Left            =   12000
      TabIndex        =   21
      Top             =   300
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   14910
      TabIndex        =   0
      Top             =   6690
      Width           =   14910
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdTransportadora 
      Height          =   5715
      Left            =   11640
      TabIndex        =   19
      Top             =   720
      Width           =   3255
      _cx             =   5741
      _cy             =   10081
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
      FormatString    =   $"frmTransportadora.frx":0029
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   120
      TabIndex        =   22
      Top             =   840
      Width           =   11415
      Begin VB.ComboBox cmbSituacao 
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
         Left            =   9000
         TabIndex        =   4
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtInscricao 
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
         Left            =   4920
         TabIndex        =   3
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox txtCNPJ 
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
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   4695
      End
      Begin VB.TextBox txtCodigo 
         BackColor       =   &H00A3A3A3&
         Enabled         =   0   'False
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
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtNome 
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
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   9975
      End
      Begin VB.Label lblSituacao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Situação"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   9000
         TabIndex        =   28
         Top             =   840
         Width           =   630
      End
      Begin VB.Label lblIE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inscrição Estadual"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   4920
         TabIndex        =   27
         Top             =   840
         Width           =   1305
      End
      Begin VB.Label lblCNPJ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ /CPF"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   825
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblNome 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Razão Social"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   1320
         TabIndex        =   23
         Top             =   120
         Width           =   945
      End
   End
   Begin CentroDeDistribuicao.chameleonButton cmdRetorna 
      Height          =   510
      Left            =   13560
      TabIndex        =   18
      Top             =   6840
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
      MICON           =   "frmTransportadora.frx":0073
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdNovo 
      Height          =   510
      Left            =   9240
      TabIndex        =   15
      Top             =   6840
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "&Novo"
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
      MICON           =   "frmTransportadora.frx":008F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdGrava 
      Height          =   510
      Left            =   10680
      TabIndex        =   16
      Top             =   6840
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "&Grava"
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
      MICON           =   "frmTransportadora.frx":00AB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdAtualiza 
      Height          =   510
      Left            =   12120
      TabIndex        =   17
      Top             =   6840
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "&Atualiza"
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
      MICON           =   "frmTransportadora.frx":00C7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblFiltro 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filtro"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   11520
      TabIndex        =   20
      Top             =   360
      Visible         =   0   'False
      Width           =   330
   End
End
Attribute VB_Name = "frmTransportadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub grdMunicipio_DBLClick()
    txtMunicipio.Text = grdMunicipio.TextMatrix(grdMunicipio.row, 0)
    fraMunicipio.Visible = False
End Sub

Private Sub grdMunicipio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grdMunicipio_DBLClick
    End If
    If KeyAscii = 27 Then
        fraMunicipio.Visible = False
    End If
End Sub

Private Sub grdMunicipio_LeaveCell()
    If grdMunicipio.row >= 0 Then
    
    txtMunicipio.Text = grdMunicipio.TextMatrix(grdMunicipio.row, 0)
    
    End If
End Sub

Private Sub txtMunicipio_KeyPress(KeyAscii As Integer)
    If Len(txtMunicipio.Text) > 3 Then
        fraMunicipio.Visible = True
        carregaComboMunicipio cmbUF.Text, Trim(txtMunicipio.Text)
        If grdMunicipio.row > 0 Then
            grdMunicipio.row = 0
        End If
    End If
    
End Sub

Private Sub txtMunicipio_LostFocus()
    ValidaCamposObrigatorios
End Sub

Private Sub cmbSituacao_lostFocus()
    ValidaCamposObrigatorios
End Sub


Private Sub cmbUF_LostFocus()
    ValidaCamposObrigatorios
End Sub

Private Sub cmdAtualiza_Click()
    
    If MsgBox("Confirma a Atualização do Cadastro?", vbYesNo + vbQuestion + vbDefaultButton2, "Atualização Transportadora") = vbYes Then
            
            atualizaTransportadora
            pesquisaTransportadora grdTransportadora.TextMatrix(1, 0)
            
    End If
        
End Sub

Private Sub cmdGrava_Click()
    Dim sql As String
    Dim rsTransportadora As New ADODB.Recordset
    
    sql = "select tr_cgc from transportadora where tr_cgc like '" & txtcnpj.Text & "'"
    rsTransportadora.CursorLocation = adUseClient
    rsTransportadora.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If rsTransportadora.EOF Then
            criaTransportadora
            pesquisaTodasTransportadora
            cmdNovo_Click
    Else
        MsgBox "CNPJ/CPF já Cadastrado ", vbExclamation, "Transportadora já cadastrada"
    End If
    
    rsTransportadora.Close
End Sub

Private Sub cmdNovo_Click()
    
    cmbSituacao.Enabled = False
    txtCodigo.Text = ""
    txtNome.Text = ""
    txtEndereco.Text = ""
    txtBairro.Text = ""
    txtCep.Text = ""
    txtTelefone.Text = ""
    txtContato.Text = ""
    txtcnpj.Text = ""
    txtNumero.Text = ""
    txtInscricao.Text = ""
    txtEmail.Text = ""
    txtComplemento.Text = ""
    cmbSituacao = "Ativo"
    txtMunicipio.Text = ""
    cmdAtualiza.Enabled = False
    cmdGrava.Enabled = False
    
End Sub

Private Sub cmdRetorna_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    carregarPosicaoTamanhoTela Me
    pesquisaTodasTransportadora
    cmbSituacao.AddItem ("Ativo")
    cmbSituacao.AddItem ("Não Ativo")
    cmbSituacao.Text = "Ativo"
    carregaComboEstado
    If grdTransportadora.Rows > 1 Then
        pesquisaTransportadora grdTransportadora.TextMatrix(1, 0)
    End If
    carregaComboMunicipio Trim(cmbUF.Text), ""
    
End Sub


Private Sub pesquisaTodasTransportadora()
    
    Dim sql As String
    Dim rsTransportadora As New ADODB.Recordset
    
    sql = "select tr_codigo,tr_nome from transportadora order by tr_codigo"
    
    rsTransportadora.CursorLocation = adUseClient
    rsTransportadora.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not rsTransportadora.EOF Then
        
        Do While Not rsTransportadora.EOF
        
            grdTransportadora.AddItem rsTransportadora("tr_codigo") & Chr(9) _
            & rsTransportadora("tr_nome")
        
            rsTransportadora.MoveNext
        Loop
        
    End If
    
    rsTransportadora.Close

End Sub


Private Sub atualizaTransportadora()
    Dim sql As String
    Dim Situacao As String
    
    If cmbSituacao.Text = "Ativo" Then
        Situacao = "A"
    Else
        Situacao = "N"
    End If
    
    sql = "exec sp_cdm_Atualiza_transportadora " & Trim(txtCodigo.Text) & ",'" _
    & Situacao & "', '" & Trim(txtNome.Text) & "', " _
    & "'" & Trim(txtEndereco.Text) & "', '" & Trim(txtNumero.Text) & "','" _
    & Trim(txtComplemento.Text) & "', '" & Trim(txtBairro.Text) & "', '" _
    & Trim(txtCep.Text) & "', '" & Trim(txtMunicipio.Text) & "', '" _
    & Trim(cmbUF.Text) & "', '" & txtTelefone.Text & "', '" _
    & Trim(txtContato.Text) & "','" & Trim(txtcnpj.Text) & "', '" _
    & Trim(txtInscricao.Text) & "', '" & Trim(txtEmail.Text) & "'"
    
    ADO_Cn_CDLocal.Execute (sql)
End Sub


Private Sub criaTransportadora()
    Dim sql As String
    
    sql = "Exec SP_CDM_Cria_transportadora '" & Trim(txtNome.Text) & "', " _
    & "'" & Trim(txtEndereco.Text) & "', '" & Trim(txtNumero.Text) & "','" _
    & Trim(txtComplemento.Text) & "', '" & Trim(txtBairro.Text) & "', '" _
    & Trim(txtCep.Text) & "', '" & Trim(txtMunicipio.Text) & "', '" _
    & Trim(cmbUF.Text) & "', '" & txtTelefone.Text & "', '" _
    & Trim(txtContato.Text) & "','" & Trim(txtcnpj.Text) & "', '" _
    & Trim(txtInscricao.Text) & "', '" & Trim(txtEmail.Text) & "'"
    
    ADO_Cn_CDLocal.Execute (sql)
End Sub

Private Sub deletaTransportadora()
    
    Dim sql As String
    
    sql = "SP_CDM_deleta_transportadora " & Trim(txtCodigo.Text)
    ADO_Cn_CDLocal.Execute (sql)

End Sub

Private Sub pesquisaTransportadora(codigo As String)

    Dim sql As String
    Dim rsTransportadora As New ADODB.Recordset
    
    sql = " select * from transportadora where tr_codigo = " & Trim(codigo)
    
    rsTransportadora.CursorLocation = adUseClient
    rsTransportadora.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not rsTransportadora.EOF Then
    
        txtCodigo.Text = rsTransportadora("tr_codigo")
        txtNome.Text = rsTransportadora("tr_nome")
        txtEndereco.Text = rsTransportadora("tr_endereco")
        txtNumero.Text = rsTransportadora("tr_numero")
        txtComplemento.Text = rsTransportadora("tr_complemento")
        txtBairro.Text = rsTransportadora("tr_bairro")
        txtCep.Text = rsTransportadora("tr_cep")
        txtMunicipio.Text = rsTransportadora("tr_cidade")
        cmbUF.Text = rsTransportadora("tr_estado")
        txtTelefone.Text = rsTransportadora("tr_telefone")
        txtContato.Text = rsTransportadora("tr_contato")
        txtcnpj.Text = rsTransportadora("tr_cgc")
        txtInscricao.Text = rsTransportadora("tr_inscricaoEstadual")
        txtEmail.Text = rsTransportadora("tr_email")
        
        If rsTransportadora("tr_status") = "A" Then
            cmbSituacao.Text = "Ativo"
        Else
            cmbSituacao.Text = "Não Ativo"
        End If
    
    End If
    
    rsTransportadora.Close
    cmdAtualiza.Enabled = True
    cmdGrava.Enabled = False
    
End Sub

Private Sub carregaComboEstado()

    Dim sql As String
    Dim rsEstado As New ADODB.Recordset
    
    sql = "select uf_estado from fin_estado order by uf_estado"
    
    rsEstado.CursorLocation = adUseClient
    rsEstado.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not rsEstado.EOF Then
        cmbUF.Text = Trim(rsEstado("uf_estado"))
        Do While Not rsEstado.EOF
            cmbUF.AddItem Trim(rsEstado("uf_estado"))
            rsEstado.MoveNext
        Loop
    End If
    rsEstado.Close


End Sub

Private Sub grdTransportadora_dblClick()
    If grdTransportadora.row <> 0 Then
        pesquisaTransportadora grdTransportadora.TextMatrix(grdTransportadora.row, 0)
    End If
End Sub

Private Sub carregaComboMunicipio(uf As String, filtro As String)
    
    Dim sql As String
    Dim rsMunicipio As New ADODB.Recordset
    
    'Limpa Grid Municipio
    grdMunicipio.Rows = 0
    grdMunicipio.AddItem ""
    grdMunicipio.RemoveItem (0)
    
    sql = "select top 5 mun_nome from fin_municipio where mun_uf = '" & uf & "' and mun_nome like '" & filtro & "%'order by mun_nome"
    
    rsMunicipio.CursorLocation = adUseClient
    rsMunicipio.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not rsMunicipio.EOF Then
        Do While Not rsMunicipio.EOF
            grdMunicipio.AddItem Trim(rsMunicipio("mun_nome"))
            rsMunicipio.MoveNext
        Loop
    End If
    
    rsMunicipio.Close
    
End Sub

Private Sub grdTransportadora_RowColChange()
    grdTransportadora_dblClick
End Sub

Private Sub ValidaCamposObrigatorios()

    If txtNome.Text <> "" And txtCodigo.Text = "" And txtcnpj.Text <> "" And txtInscricao.Text <> "" And txtEndereco.Text <> "" And txtNumero.Text <> "" And txtBairro.Text <> "" And txtCep.Text <> "" And txtMunicipio.Text <> "" Then
        cmdGrava.Enabled = True
    Else
        cmdGrava.Enabled = False
    End If
End Sub

Private Sub txtBairro_lostFocus()
    ValidaCamposObrigatorios
End Sub

Private Sub txtCep_LostFocus()
    ValidaCamposObrigatorios
End Sub

Private Sub txtCNPJ_LostFocus()
    ValidaCamposObrigatorios
End Sub

Private Sub txtComplemento_lostFocus()
    ValidaCamposObrigatorios
End Sub

Private Sub txtContato_lostFocus()
    ValidaCamposObrigatorios
End Sub

Private Sub txtEmail_lostFocus()
    ValidaCamposObrigatorios
End Sub

Private Sub txtEndereco_lostFocus()
    ValidaCamposObrigatorios
End Sub

Private Sub txtInscricao_lostFocus()
    ValidaCamposObrigatorios
End Sub

Private Sub txtNome_lostFocus()
    ValidaCamposObrigatorios
End Sub

Private Sub txtNumero_lostFocus()
    ValidaCamposObrigatorios
End Sub

Private Sub txtTelefone_lostFocus()
    ValidaCamposObrigatorios
End Sub
