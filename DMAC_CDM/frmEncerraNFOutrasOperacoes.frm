VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7d.ocx"
Begin VB.Form frmEncerraNFOutrasOperacoes 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Encerra Nota Fiscal Outras Operações"
   ClientHeight    =   7875
   ClientLeft      =   23370
   ClientTop       =   3255
   ClientWidth     =   15300
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   15300
   ShowInTaskbar   =   0   'False
   Begin VSFlex7DAOCtl.VSFlexGrid grdMunicipio2 
      Height          =   750
      Left            =   6900
      TabIndex        =   91
      Top             =   4215
      Visible         =   0   'False
      Width           =   3435
      _cx             =   6059
      _cy             =   1323
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
      BackColor       =   16777215
      ForeColor       =   4210752
      BackColorFixed  =   0
      ForeColorFixed  =   16777215
      BackColorSel    =   3421236
      ForeColorSel    =   16777215
      BackColorBkg    =   12632256
      BackColorAlternate=   16777215
      GridColor       =   14737632
      GridColorFixed  =   8421504
      TreeColor       =   8421504
      FloodColor      =   16777215
      SheetBorder     =   8421504
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   0
      GridLineWidth   =   0
      Rows            =   8
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmEncerraNFOutrasOperacoes.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
      BackColorFrozen =   -2147483633
      ForeColorFrozen =   4210752
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   2070
      Left            =   7395
      TabIndex        =   83
      Top             =   4530
      Width           =   7605
      Begin VB.TextBox txtCarimbo 
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
         Height          =   1590
         Left            =   120
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   30
         Top             =   360
         Width           =   7365
      End
      Begin VB.Label lblCarimbo 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Carimbo Nota Fiscal (Máximo de 150 caracteres)"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   120
         TabIndex        =   84
         Top             =   120
         Width           =   3420
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Pesquisa"
      ForeColor       =   &H00FF0000&
      Height          =   465
      Left            =   120
      TabIndex        =   81
      Top             =   120
      Width           =   2385
      Begin VB.TextBox txtPedido 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1095
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   0
         Top             =   75
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00404040&
         Caption         =   "Pedido:"
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
         TabIndex        =   82
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   43
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   14880
      TabIndex        =   45
      Top             =   6810
      Width           =   14880
   End
   Begin VB.Frame fraRemetente 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   2085
      Left            =   120
      TabIndex        =   44
      Top             =   2280
      Width           =   14880
      Begin VB.ComboBox cmdLojaDestino 
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
         Left            =   6510
         TabIndex        =   88
         Text            =   "CD"
         Top             =   525
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.OptionButton optLoja 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Loja"
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   3180
         TabIndex        =   87
         Top             =   480
         Width           =   810
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
         Left            =   4455
         ScrollBars      =   1  'Horizontal
         TabIndex        =   18
         Top             =   1620
         Width           =   2250
      End
      Begin VB.TextBox txtCodigoMunicipio 
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
         Left            =   8850
         ScrollBars      =   1  'Horizontal
         TabIndex        =   20
         Top             =   1620
         Width           =   1350
      End
      Begin VB.TextBox txtBairroDestinatario 
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
         Left            =   10275
         ScrollBars      =   1  'Horizontal
         TabIndex        =   21
         Top             =   1620
         Width           =   2010
      End
      Begin VB.TextBox txtNroEnd 
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
         Left            =   3555
         TabIndex        =   17
         Top             =   1620
         Width           =   825
      End
      Begin VB.TextBox txtCodigoDestinatario 
         Alignment       =   1  'Right Justify
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
         Left            =   4875
         MaxLength       =   7
         TabIndex        =   11
         Top             =   420
         Width           =   1050
      End
      Begin VB.OptionButton optFornecedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Fornecedor"
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   120
         TabIndex        =   73
         Top             =   480
         Width           =   1125
      End
      Begin VB.OptionButton optCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Cliente"
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   1290
         TabIndex        =   31
         Top             =   480
         Value           =   -1  'True
         Width           =   810
      End
      Begin VB.OptionButton optInformado 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Informado"
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   2115
         TabIndex        =   32
         Top             =   480
         Width           =   1005
      End
      Begin VB.TextBox txtDestinatario 
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
         Top             =   1005
         Width           =   8040
      End
      Begin VB.TextBox txtCNPJDestinatario 
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
         Left            =   8235
         TabIndex        =   13
         Top             =   1005
         Width           =   1980
      End
      Begin VB.TextBox txtInscricaoDestinatario 
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
         Left            =   10290
         TabIndex        =   14
         Top             =   1005
         Width           =   2000
      End
      Begin VB.TextBox txtEnderecoDestinatario 
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
         TabIndex        =   16
         Top             =   1620
         Width           =   3360
      End
      Begin VB.TextBox txtMunicipioDestinatario 
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
         Left            =   6780
         ScrollBars      =   1  'Horizontal
         TabIndex        =   19
         Top             =   1620
         Width           =   2000
      End
      Begin VB.TextBox txtCepDestinatario 
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
         Left            =   12360
         ScrollBars      =   1  'Horizontal
         TabIndex        =   15
         Top             =   1005
         Width           =   2430
      End
      Begin VB.TextBox txtUFDestinatario 
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
         Left            =   12360
         ScrollBars      =   1  'Horizontal
         TabIndex        =   22
         Top             =   1620
         Width           =   405
      End
      Begin VB.TextBox txtFoneFaxDestinatario 
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
         Left            =   12840
         ScrollBars      =   1  'Horizontal
         TabIndex        =   23
         Top             =   1620
         Width           =   1950
      End
      Begin VB.Label lblCodigoMunicipioDest 
         BackColor       =   &H00404040&
         Caption         =   "Código Município"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   8850
         TabIndex        =   79
         Top             =   1395
         Width           =   1440
      End
      Begin VB.Label lblComplementoDest 
         BackColor       =   &H00404040&
         Caption         =   "Complemento"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   4455
         TabIndex        =   78
         Top             =   1395
         Width           =   1410
      End
      Begin VB.Label lblMunicipioDest 
         BackColor       =   &H00404040&
         Caption         =   "Município"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   6780
         TabIndex        =   77
         Top             =   1395
         Width           =   1440
      End
      Begin VB.Label lblNroDest 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   3555
         TabIndex        =   75
         Top             =   1395
         Width           =   555
      End
      Begin VB.Label lblCodigoRemetente 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Código:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   4290
         TabIndex        =   74
         Top             =   480
         Width           =   540
      End
      Begin VB.Label lblFoneFaxDest 
         BackColor       =   &H00404040&
         Caption         =   "Fone/Fax"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   12825
         TabIndex        =   62
         Top             =   1395
         Width           =   780
      End
      Begin VB.Label lblCNPJDest 
         BackColor       =   &H00404040&
         Caption         =   "CNPJ "
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   8235
         TabIndex        =   61
         Top             =   780
         Width           =   780
      End
      Begin VB.Label lblInscricaoEstadualDest 
         BackColor       =   &H00404040&
         Caption         =   "Inscrição Estadual"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   10290
         TabIndex        =   60
         Top             =   780
         Width           =   1590
      End
      Begin VB.Label lblEnderecoDest 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço "
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   1395
         Width           =   780
      End
      Begin VB.Label lblBairroDest 
         BackColor       =   &H00404040&
         Caption         =   "Bairro"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   10275
         TabIndex        =   58
         Top             =   1395
         Width           =   780
      End
      Begin VB.Label lblUFDest 
         BackColor       =   &H00404040&
         Caption         =   "UF"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   12345
         TabIndex        =   57
         Top             =   1395
         Width           =   345
      End
      Begin VB.Label lblCepDest 
         BackColor       =   &H00404040&
         Caption         =   "CEP"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   12375
         TabIndex        =   56
         Top             =   780
         Width           =   780
      End
      Begin VB.Label lblDestinatario 
         BackColor       =   &H00404040&
         Caption         =   "Destinatário"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   780
         Width           =   1110
      End
      Begin VB.Label lblDestinatarioRemetente 
         BackColor       =   &H00404040&
         Caption         =   "Destinatário/Remetente"
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
         TabIndex        =   54
         Top             =   120
         Width           =   2085
      End
   End
   Begin VB.Frame fraEmitente 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1395
      Left            =   120
      TabIndex        =   43
      Top             =   720
      Width           =   14880
      Begin VB.TextBox txtBairroEmitente 
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
         Left            =   7950
         Locked          =   -1  'True
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   945
         Width           =   2970
      End
      Begin VB.ComboBox cmbLojaOrigem 
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
         TabIndex        =   1
         Text            =   "CD"
         Top             =   330
         Width           =   915
      End
      Begin VB.TextBox txtEmitente 
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
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   330
         Width           =   3720
      End
      Begin VB.TextBox txtCNPJEmitente 
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
         Left            =   4905
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   330
         Width           =   1890
      End
      Begin VB.TextBox txtInscricaoEmitente 
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
         Left            =   6870
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   330
         Width           =   1800
      End
      Begin VB.TextBox txtEnderecoEmitente 
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
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   945
         Width           =   4710
      End
      Begin VB.TextBox txtFoneFaxEmitente 
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
         Left            =   12810
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   945
         Width           =   1965
      End
      Begin VB.TextBox txtUFEmitente 
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
         Left            =   10995
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   945
         Width           =   345
      End
      Begin VB.TextBox txtCepEmitente 
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
         Left            =   11415
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   945
         Width           =   1320
      End
      Begin VB.TextBox txtMunicipioEmitente 
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
         Left            =   4905
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   945
         Width           =   2970
      End
      Begin VB.TextBox txtDataEmissao 
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
         Left            =   8745
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   330
         Width           =   1215
      End
      Begin VB.Label lblBairro 
         BackColor       =   &H00404040&
         Caption         =   "Bairro"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   7950
         TabIndex        =   90
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "Emitente"
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
         TabIndex        =   72
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblInscricaoEstadual 
         BackColor       =   &H00404040&
         Caption         =   "Inscrição Estadual"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   4905
         TabIndex        =   53
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lblEmissao 
         BackColor       =   &H00404040&
         Caption         =   "Emissão"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   6870
         TabIndex        =   52
         Top             =   120
         Width           =   630
      End
      Begin VB.Label lblEndereco 
         BackColor       =   &H00404040&
         Caption         =   "Endereço"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label lblMunicipio 
         BackColor       =   &H00404040&
         Caption         =   "Município"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   4905
         TabIndex        =   50
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label lblUf 
         BackColor       =   &H00404040&
         Caption         =   "UF"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   11025
         TabIndex        =   49
         Top             =   720
         Width           =   315
      End
      Begin VB.Label lblCep 
         BackColor       =   &H00404040&
         Caption         =   "CEP"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   11430
         TabIndex        =   48
         Top             =   720
         Width           =   435
      End
      Begin VB.Label lblFoneFax 
         BackColor       =   &H00404040&
         Caption         =   "Fone/Fax"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   12795
         TabIndex        =   47
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblCnpj 
         BackColor       =   &H00404040&
         Caption         =   "CNPJ"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   1110
         TabIndex        =   46
         Top             =   120
         Width           =   1080
      End
   End
   Begin VB.Frame fraTotalNF 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Totais da Nota Fiscal"
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
      ForeColor       =   &H000703BA&
      Height          =   1050
      Left            =   15315
      TabIndex        =   42
      Top             =   5310
      Visible         =   0   'False
      Width           =   14880
      Begin VB.TextBox txtTotalNF 
         Alignment       =   1  'Right Justify
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
         Left            =   6900
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   585
         Width           =   1275
      End
      Begin VB.TextBox txtValormercadoria 
         Alignment       =   1  'Right Justify
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
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   585
         Width           =   1290
      End
      Begin VB.TextBox txtValorICMSST 
         Alignment       =   1  'Right Justify
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
         Left            =   4125
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   585
         Width           =   1305
      End
      Begin VB.TextBox txtBaseCalculoST 
         Alignment       =   1  'Right Justify
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
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   585
         Width           =   1275
      End
      Begin VB.TextBox txtValorICMS 
         Alignment       =   1  'Right Justify
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
         Left            =   1425
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   585
         Width           =   1275
      End
      Begin VB.TextBox txtBaseICMS 
         Alignment       =   1  'Right Justify
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
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   585
         Width           =   1275
      End
      Begin VB.Label lblBaseCalcICMS 
         BackColor       =   &H00404040&
         Caption         =   "Base Calc ICMS"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   120
         TabIndex        =   71
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label lblValorICMS 
         BackColor       =   &H00404040&
         Caption         =   "Valor ICMS"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   1425
         TabIndex        =   70
         Top             =   360
         Width           =   840
      End
      Begin VB.Label lblBaseCalcST 
         BackColor       =   &H00404040&
         Caption         =   "Base Cálculo ST"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   2760
         TabIndex        =   69
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label lblValorICMSST 
         BackColor       =   &H00404040&
         Caption         =   "Valor ICMS ST"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   4125
         TabIndex        =   68
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label lblValorMercadoria 
         BackColor       =   &H00404040&
         Caption         =   "Valor Mercadoria"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   5535
         TabIndex        =   67
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lblTotalNF 
         BackColor       =   &H00404040&
         Caption         =   "Total Nota Fiscal"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   6900
         TabIndex        =   66
         Top             =   360
         Width           =   1305
      End
      Begin VB.Label lblTotaisNotasFiscais 
         BackColor       =   &H00404040&
         Caption         =   "Totais das notas fiscais"
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
         TabIndex        =   65
         Top             =   120
         Width           =   2040
      End
   End
   Begin VB.Frame FraCFOP 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   2070
      Left            =   120
      TabIndex        =   41
      Top             =   4530
      Width           =   7095
      Begin VB.TextBox txtQuantidade 
         Alignment       =   1  'Right Justify
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
         Left            =   105
         MaxLength       =   7
         TabIndex        =   27
         Top             =   1635
         Width           =   1515
      End
      Begin VB.TextBox txtPeso 
         Alignment       =   1  'Right Justify
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
         Left            =   1695
         MaxLength       =   7
         TabIndex        =   28
         Top             =   1635
         Width           =   1515
      End
      Begin VB.ComboBox cmdTipoFrete 
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
         Left            =   3285
         TabIndex        =   29
         Text            =   "Combo2"
         Top             =   1635
         Width           =   3700
      End
      Begin VB.ComboBox cmbCFOP 
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
         Left            =   3285
         TabIndex        =   25
         Text            =   "Combo2"
         Top             =   345
         Width           =   3700
      End
      Begin VB.ComboBox cmbTipoES 
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
         TabIndex        =   24
         Text            =   "Combo1"
         Top             =   345
         Width           =   3090
      End
      Begin MSMask.MaskEdBox mskChaveNFe 
         Height          =   315
         Left            =   105
         TabIndex        =   26
         ToolTipText     =   "Data Emissão"
         Top             =   990
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483648
         MaxLength       =   54
         Format          =   "dd/mm/yyyy"
         Mask            =   "####-####-####-####-####-####-####-####-####-####-####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label8 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Peso"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   1710
         TabIndex        =   86
         Top             =   1395
         Width           =   1605
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404040&
         Caption         =   "Quantidade"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   120
         TabIndex        =   85
         Top             =   1395
         Width           =   2430
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404040&
         Caption         =   "Tipo Frete"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   3285
         TabIndex        =   80
         Top             =   1395
         Width           =   3510
      End
      Begin VB.Label lbChaveNFE 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Chave de Acesso da Nota Devolvida"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   120
         TabIndex        =   76
         Top             =   750
         Width           =   2640
      End
      Begin VB.Label lblCFOP 
         BackColor       =   &H00404040&
         Caption         =   "Código da Operação Fiscal"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   3285
         TabIndex        =   64
         Top             =   120
         Width           =   3660
      End
      Begin VB.Label lblTipoES 
         BackColor       =   &H00404040&
         Caption         =   "Tipo de Entrada/Saida"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   120
         TabIndex        =   63
         Top             =   120
         Width           =   2310
      End
   End
   Begin CentroDeDistribuicao.chameleonButton cmdEncerraNF 
      Height          =   510
      Left            =   12180
      TabIndex        =   34
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Encerra NF"
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
      MICON           =   "frmEncerraNFOutrasOperacoes.frx":0050
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
      TabIndex        =   35
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
      MICON           =   "frmEncerraNFOutrasOperacoes.frx":006C
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
Attribute VB_Name = "frmEncerraNFOutrasOperacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim adoLoja As New ADODB.Recordset
Dim adoCliente As New ADODB.Recordset
Dim adoFornecedor As New ADODB.Recordset
Dim rsComplementoVenda As New ADODB.Recordset
Dim adoCFOP As New ADODB.Recordset
Dim wLojaOrigem As String
Dim TipoNota As String
Dim codCli As String
Dim nf As String

'ricardo
Dim wPreencheInicio As Boolean



Private Sub cmbCFOP_click()
 
    Dim cfop As String
    Dim rsChave As New ADODB.Recordset
    Dim chave As String
    Dim sql As String
    
    cfop = Mid(Trim(cmbCFOP.Text), 1, 4)
    
    If cfop = "6909" Or cfop = "6202" Or cfop = "6411" Or cfop = "6918" Or cfop = "6913" Or cfop = "2202" Or cfop = "2411" Or cfop = "5918" Or cfop = "5202" Then
        'mskChaveNFe.Enabled = True
        mskChaveNFe.Mask = ("####-####-####-####-####-####-####-####-####-####-####")
        sql = "select ChaveNfeDevolucao from nfcapa where numeroPed = " & wPedido
        
        rsChave.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
            If Not rsChave.EOF Then
                chave = Trim(rsChave("ChaveNfeDevolucao"))
                If chave <> "" Then mskChaveNFe.Mask = chave
            End If
        rsChave.Close
    Else
        mskChaveNFe.Mask = ("####-####-####-####-####-####-####-####-####-####-####")
        'mskChaveNFe.Enabled = False
    End If
    
    
            
    If cmbTipoES.ListIndex = 0 Then
        TipoNota = "EA"
    ElseIf cfop = "5152" Or cfop = "5409" Then
        TipoNota = "TA"
        If optloja.Value = False Then optLoja_Click
    Else
        TipoNota = "SA"
    End If
    
    
    
End Sub

Private Sub cmbLojaOrigem_Click()
    carregaEmitente cmbLojaOrigem.Text
End Sub

Private Sub cmbTipoES_Click()
    CarregaCFOP
End Sub


Private Sub cmdEncerraNF_Click()

    Screen.MousePointer = 11
    
    atualizaCliente
    If optloja.Value = True Or validaDados(codCli) Then
        AtualizaNota
        
        frmStartaProcessos.Show vbModal
        'frmNotaFiscalOutrasOperacoes.cmdRetorna_Click
        'frmNotaFiscalOutrasOperacoes.ExcluiNF
        'wPedido = PegaNumPedido
        'frmNotaFiscalOutrasOperacoes.txtPedido.Text = wPedido
        finalizaOutrasOperacoes = True
        Unload Me
    Else
        MsgBox "Não é póssivel encerrar. Existi erro(s) no Destinatario.", vbExclamation, "Encerrar Nota Fiscal"
    End If
    
    Screen.MousePointer = 0
    
End Sub

Private Sub cmdLojaDestino_Click()
    If optloja.Value = True Then
        CarregaLojaDestino Trim(cmdLojaDestino.Text)
    End If
End Sub

Private Sub cmdLojaDestino_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdLojaDestino_Click
    End If
End Sub

Private Sub cmdLojaDestino_LostFocus()
    cmdLojaDestino_Click
End Sub

Private Sub cmdRetorna_Click()

    Unload Me

End Sub

Private Sub Form_Activate()
    txtPedido.Text = wPedido
    carregaComboloja GLB_Loja
    cmbLojaOrigem_Click
    CarregaNota
    'carregaDestinatario
    carregaLojasComboTransferencia
    finalizaOutrasOperacoes = False
End Sub

Private Function validaDados(codigoCliente As String) As Boolean

    Dim adoValidaCliente As New ADODB.Recordset
    Dim sql As String
    
    If optloja.Value = True Then Exit Function
    
    If codigoCliente = Empty Then
        MsgBox "Não foi informado o codigo do Cliente!", vbExclamation, "Cliente"
        validaDados = False
        Exit Function
    End If
    
    sql = "exec SP_GLB_Valida_Cliente '" & codigoCliente & "'"
    ADO_Cn_CDLocal.Execute sql
    
    sql = "select campoErrado as campoErrado from temp_Fin_Cliente_Erro"
    adoValidaCliente.CursorLocation = adUseClient
    adoValidaCliente.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    validaDados = True
    
    lblCNPJDest.ForeColor = lblCNPJ.ForeColor
    lblInscricaoEstadualDest.ForeColor = lblCNPJ.ForeColor
    lblDestinatario.ForeColor = lblCNPJ.ForeColor
    lblEnderecoDest.ForeColor = lblCNPJ.ForeColor
    lblBairroDest.ForeColor = lblCNPJ.ForeColor
    lblMunicipioDest.ForeColor = lblCNPJ.ForeColor
    lblCepDest.ForeColor = lblCNPJ.ForeColor
    lblFoneFaxDest.ForeColor = lblCNPJ.ForeColor
    lblNroDest.ForeColor = lblCNPJ.ForeColor
    lblCodigoMunicipioDest.ForeColor = lblCNPJ.ForeColor
    lblCodigoMunicipioDest.ForeColor = lblCNPJ.ForeColor
    
    Do While Not adoValidaCliente.EOF
        Select Case adoValidaCliente("campoErrado")
        Case "CGC"
            lblCNPJDest.ForeColor = vbRed
            validaDados = False
            
        Case "InscricaoEstadual"
            lblInscricaoEstadualDest.ForeColor = vbRed
            validaDados = False
            
        Case "Razao"
            lblDestinatario.ForeColor = vbRed
            validaDados = False
            
        Case "Endereco"
            lblEnderecoDest.ForeColor = vbRed
            validaDados = False
            
        Case "Bairro"
            lblBairroDest.ForeColor = vbRed
            validaDados = False
            
        Case "Municipio"
            lblMunicipioDest.ForeColor = vbRed
            validaDados = False
            
        Case "CEP"
            lblCepDest.ForeColor = vbRed
            validaDados = False
            
        Case "Telefone"
            lblFoneFaxDest.ForeColor = vbRed
            validaDados = False
            
        Case "Numero"
            lblNroDest.ForeColor = vbRed
            validaDados = False
            
        Case "Mun_Codigo"
            lblCodigoMunicipioDest.ForeColor = vbRed
            validaDados = False
            
        End Select

        adoValidaCliente.MoveNext
    Loop
    
    adoValidaCliente.Close
    
End Function

Private Sub Form_Load()

    carregarPosicaoTamanhoTela Me
    
    'CarregaTipoES
    carregaTipoFrete
    
    txtBaseCalculoST.Text = "0,00"
    txtValorICMSST.Text = "0,00"
    mskChaveNFe.Mask = ("####-####-####-####-####-####-####-####-####-####-####")
    'mskChaveNFe.Enabled = False
    
    cmdLojaDestino.left = txtCodigoDestinatario.left
    cmdLojaDestino.top = txtCodigoDestinatario.top
    
    
End Sub

Private Sub carregaTipoFrete()
    cmdTipoFrete.AddItem "0 - Emitente"
    cmdTipoFrete.AddItem "1 - Destinatário/remetente"
    cmdTipoFrete.AddItem "2 - Terceiros"
    cmdTipoFrete.AddItem "9 - Sem frete"
    'cmdTipoFrete.ListIndex = 1
End Sub

Private Sub atualizaCliente()


    Dim cgc As String
    Dim razao As String
    Dim inscEstadual As String
    Dim CEP As String
    Dim endereco As String
    Dim bairro As String
    Dim numero As String
    Dim municipio As String
    Dim estado As String
    Dim telefone As String
    Dim tel As String
    Dim codigoMunicipio As String
    Dim complemento As String
    Dim adoDestinatario As New ADODB.Recordset


    If optloja.Value = True Then
        codCli = cmdLojaDestino.Text
        Exit Sub
    End If
    
    cgc = Trim(txtCNPJDestinatario.Text)
    razao = Trim(txtDestinatario.Text)
    inscEstadual = Trim(txtInscricaoDestinatario.Text)
    CEP = Trim(txtCepDestinatario.Text)
    endereco = Trim(txtEnderecoDestinatario.Text)
    numero = Trim(txtNroEnd.Text)
    municipio = Trim(txtMunicipioDestinatario.Text)
    estado = Trim(txtUFDestinatario.Text)
    bairro = Trim(txtBairroDestinatario.Text)
    tel = Trim(txtFoneFaxDestinatario.Text)
    codigoMunicipio = Trim(txtCodigoMunicipio.Text)
    complemento = Trim(txtComplemento.Text)

    sql = "select top 1 ce_codigoCliente from fin_cliente where ce_cgc = '" & cgc & "'"
    
    adoDestinatario.CursorLocation = adUseClient
    adoDestinatario.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not adoDestinatario.EOF Then
        codCli = adoDestinatario("ce_CodigoCliente")
    End If
    
    adoDestinatario.Close
    
    If codCli = "" Then
    
        codCli = PegaCliente
        
        sql = "SP_FIN_Grava_Cliente_loja '" & codCli & "', '" & razao & " ', '" & cgc & "', 'J', '0','N', '" & _
        inscEstadual & "', '" & CEP & "', '" & endereco & "', '" & numero & "', '" & _
        municipio & "', '" & codigoMunicipio & "', '" & estado & "','" & complemento & "' , '" & bairro & "', 0 , '" & tel & "','' ,'','',0,'',0, '" & _
        endereco & "', '" & numero & "', '', '" & CEP & "', '" & bairro & "', '" & municipio & "', '" & _
        estado & "', '','','',0,'" & wLoja & "'"
        
        ADO_Cn_CDLocal.Execute (sql)
    Else
    
        sql = "SP_FIN_Altera_Cliente " & codCli & ",'" & razao & "','" & cgc & "'," _
                            & "'J'" & ", " & "'0'" & "," _
                            & "'N'" & ", '" _
                            & inscEstadual & "','" & CEP & "','" _
                            & endereco & "','" & numero & "','" _
                            & municipio & "','" & codigoMunicipio & "','" _
                            & estado & "','" & complemento & "','" _
                            & bairro & "'," & "0" & ",'" _
                            & tel & "','" & tel & "','" _
                            & tel & "','" & "" & "','" _
                            & "" & "','" & "" & "','" _
                            & "" & "', '" & "" & "','" _
                            & "" & "','" & "" & "','" _
                            & "" & "','" & "" & "','" _
                            & "" & "','" & "" & "',0.00"
                        
        ADO_Cn_CDLocal.Execute (sql)
    
    End If
    
    If optFornecedor Or optInformado Then
    
        sql = "select top 1 ce_codigoCliente from fin_cliente where ce_cgc = '" & cgc & "'"
        
        adoDestinatario.CursorLocation = adUseClient
        adoDestinatario.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
    If Not adoDestinatario.EOF Then
        codCli = adoDestinatario("ce_CodigoCliente")
    End If
    
    adoDestinatario.Close
    
    
    ElseIf optCliente Then
    
        codCli = txtCodigoDestinatario.Text
    
    End If
    
End Sub

Private Sub AtualizaNota()

    Dim sql As String
    Dim codOper As String
    Dim serie As String
    Dim carimbo As String
    Dim Data As String
    Dim tipoFrete As String * 1
    Dim chaveNFeDev As String
    Dim resCarimboImpostos As New ADODB.Recordset

    
    Data = Format$(Now, "yyyy/m/d hh:mm:ss")
    
    codOper = Mid(cmbCFOP.Text, 1, 4)
    
    chaveNFeDev = Replace(Replace(mskChaveNFe.Text, "-", ""), "_", "")
    carimbo = Trim(txtCarimbo.Text)
 

    tipoFrete = Mid(cmdTipoFrete.Text, "1", "1")
    
    serie = "NE"
    
    If TipoNota = "TA" Then
            
        serie = SerieTransferencia(wLojaOrigem, cmdLojaDestino.Text, serie)
        
        If serie = "NE" And (nf = 0 Or nf = "") Then
            nf = ExtraiSeqNotaControleNE
        ElseIf serie <> "NE" And (nf = 0 Or nf = "") Then
            nf = ExtraiSeqNotaControle00
        End If
            
        sql = "update nfcapa set lojaT = '" & cmdLojaDestino.Text & "'," & _
              "codoper = '" & codOper & "', nf = ' " & nf & "', serie = '" & serie & "'," & _
              "CFOAUX = '" & codOper & "'," & _
              "tipoFrete = '" & tipoFrete & "'," & _
              "lojaorigem = '" & wLojaOrigem & "'," & _
              "tiponota = '" & TipoNota & "'," & _
              "volume = '" & txtQuantidade.Text & "'," & _
              "pesoLq = '" & txtPeso.Text & "'," & _
              "tm = '" & "0" & "'," & _
              "cliente = '" & "0" & "'," & _
              "chaveNFeDevolucao = '" & chaveNFeDev & "' , dataped = '" & Data & "', " & _
              "hora = '" & Data & " ' where numeroped = '" & wPedido & "'"
            ADO_Cn_CDLocal.Execute (sql)
            
    Else
    
        If nf = 0 Or nf = "" Then
            nf = ExtraiSeqNotaControleNE
        End If
    
        sql = "update nfcapa set cliente = '" & codCli & "'," & _
              "codoper = '" & codOper & "', nf = ' " & nf & "', serie = '" & serie & "'," & _
              "CFOAUX = '" & codOper & "'," & _
              "tipoFrete = '" & tipoFrete & "'," & _
              "lojaorigem = '" & wLojaOrigem & "'," & _
              "tiponota = '" & TipoNota & "'," & _
              "volume = '" & txtQuantidade.Text & "'," & _
              "pesoLq = '" & txtPeso.Text & "'," & _
              "tm = '" & "0" & "'," & _
              "chaveNFeDevolucao = '" & chaveNFeDev & "' , dataped = '" & Data & "', " & _
              "hora = '" & Data & " ' where numeroped = '" & wPedido & "'"
        ADO_Cn_CDLocal.Execute (sql)
    
    End If
    
    sql = "update NFItens set " & _
          "tiponota = '" & TipoNota & "'," & _
          "lojaorigem = '" & wLojaOrigem & "'," & _
          "CFOP = '" & codOper & "'," & _
          "situacaoProcesso = 'A'," & _
          "serie = '" & serie & "' where numeroped = '" & wPedido & "'"
    ADO_Cn_CDLocal.Execute (sql)
    
    sql = "Insert into CarimboNotaFiscal(CNF_NumeroPed,CNF_Loja,CNF_serie,CNF_NF,CNF_Sequencia,CNF_Carimbo,CNF_TipoCarimbo,CNF_SituacaoProcesso)" _
      & "Values ( " & NroPedido & ",'" & wLojaOrigem & "','" _
      & serie & "'," & "0" & "," & 1 & ",'" & txtCarimbo.Text & "','I','A')"
    ADO_Cn_CDLocal.Execute (sql)
    
    'NOVO 2016 FELIPE
    Call carimboImposto(nf, serie, wLojaOrigem)
    '''''''''''
    
    sql = "exec SP_Atualiza_Processos_Venda '" & NroPedido & "','" & nf & "','0','0'"
    ADO_Cn_CDLocal.Execute (sql)
    
    If serie = "NE" Then
        sql = "exec sp_vda_cria_nfe '" & wLojaOrigem & "','" & nf & "','" & serie & "','" & wImpressoraNota & "'"
        ADO_Cn_CDLocal.Execute (sql)
    Else
        Call ImprimeTransferencia00(nf, "CT", wLojaOrigem)
    End If
    
    'adoDestinatario.Close
End Sub

Public Function SerieTransferencia(LojaOrigem As String, LojaDestino As String, serieAtual As String)

    Dim sql As String
    Dim adoLoja As New ADODB.Recordset
    Dim CNPJlojaOrigem As String
    Dim CNPJlojaDestino As String
    
    sql = "select lo_cgc as cgc from loja where lo_loja IN ('" & LojaOrigem & "')"
          
    adoLoja.CursorLocation = adUseClient
    adoLoja.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        CNPJlojaOrigem = Trim(adoLoja("cgc"))
    adoLoja.Close
    
    sql = "select lo_cgc as cgc from loja where lo_loja IN ('" & LojaDestino & "')"
    adoLoja.CursorLocation = adUseClient
    adoLoja.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        CNPJlojaDestino = Trim(adoLoja("cgc"))
    adoLoja.Close
    
        If CNPJlojaOrigem = CNPJlojaDestino Then
            SerieTransferencia = "CT"
        Else
            SerieTransferencia = serieAtual
        End If
    
    
    

End Function

Public Function carregaComboloja(Loja As String)

    Dim sql As String
    Dim adoLoja As New ADODB.Recordset
    Dim cnpj As String * 14
    
    sql = "select lo_cgc as cgc from loja where lo_loja = '" & Loja & "'"
          
    adoLoja.CursorLocation = adUseClient
    adoLoja.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
        If Not adoLoja.EOF Then
            cnpj = adoLoja("cgc")
        Else
            MsgBox "Erro interno! Não foi possível encontrar o CNPJ dessa loja", vbCritical, "DMAC CDM"
            Unload Me
        End If
        
    adoLoja.Close
    
    sql = "select lo_loja as loja from loja" & vbNewLine & _
          "where LO_Regiao between 900 and 990 " & vbNewLine & _
          "and LO_Situacao = 'A'" & vbNewLine & _
          "AND LO_CGC like '%" & cnpj & "%' order by LO_Regiao"
          
    adoLoja.CursorLocation = adUseClient
    adoLoja.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
        If Not adoLoja.EOF Then
            cmbLojaOrigem.Clear
            Do While Not adoLoja.EOF
                cmbLojaOrigem.AddItem adoLoja("loja")
                adoLoja.MoveNext
            Loop
            cmbLojaOrigem.Enabled = True
            cmbLojaOrigem.ListIndex = 0
        Else
            MsgBox "Erro interno! Não foi possível encontrar as lojas com o CNPJ informado", vbCritical, "DMAC CDM"
            Unload Me
        End If
        
    adoLoja.Close
    
End Function

Public Function carregaEmitente(Loja As String)
    
    Dim sql As String
    Dim adoEmitente As New ADODB.Recordset
    Dim wEmitente As String
    Dim wCnpjEmi As String
    Dim wIEmi As String
    Dim wEmissao As String
    Dim wEnderecoEmi As String
    Dim wMunicipioEmi As String
    Dim wBairroEmi As String
    Dim wUFEmi As String
    Dim wCepEmi As String
    Dim wFoneFaxEmi As String
    
    sql = "select top 1 * from loja where lo_loja = '" & cmbLojaOrigem.Text & "'"
    adoEmitente.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not adoEmitente.EOF Then
    
        wEmitente = Trim(adoEmitente("lo_razao"))
        wCnpjEmi = Trim(adoEmitente("lo_cgc"))
        wIEmi = Trim(adoEmitente("lo_inscricaoEstadual"))
        wEmissao = Format(Date, "dd/mm/yyyy")
        wEnderecoEmi = Trim(adoEmitente("lo_endereco"))
        wMunicipioEmi = Trim(adoEmitente("lo_municipio"))
        wBairroEmi = Trim(adoEmitente("lo_bairro"))
        wUFEmi = Trim(adoEmitente("lo_uf"))
        wCepEmi = Trim(adoEmitente("lo_cep"))
        wFoneFaxEmi = Trim(adoEmitente("lo_telefone"))
        
    End If
    
    adoEmitente.Close
    
    txtEmitente.Text = wEmitente
    txtCNPJEmitente.Text = wCnpjEmi
    txtInscricaoEmitente = wIEmi
    txtDataEmissao = wEmissao
    txtEnderecoEmitente = wEnderecoEmi
    txtMunicipioEmitente = wMunicipioEmi
    txtBairroEmitente = wBairroEmi
    txtUFEmitente = wUFEmi
    txtCepEmitente = wCepEmi
    txtFoneFaxEmitente = wFoneFaxEmi
    
    wLojaOrigem = cmbLojaOrigem.Text
    
End Function


Public Function CarregaCliente(CodCliente As String)
    
    Dim adoCliente As New ADODB.Recordset
    
    Dim sql As String
    Dim wdestinatario As String
    Dim wcnpjDest As String
    Dim wiDest As String
    Dim wEnderecoDest As String
    Dim wNumero As String
    Dim wMunicipio As String
    Dim wBairroDest As String
    Dim wUfDest As String
    Dim wCepDest As String
    Dim wFone As String
    Dim wCodMunicipio As String
    Dim wComplemento As String
    
    If CodCliente = 0 Then Exit Function
    
    sql = "select ce_razao,ce_cgc,ce_inscricaoEstadual,ce_endereco,ce_codigoMunicipio,ce_municipio,ce_bairro,ce_estado,ce_cep,ce_telefone,ce_numero,ce_bairro,ce_codIGOMunicipio,ce_Complemento from fin_cliente where ce_codigoCliente='" & CodCliente & "'"
    adoCliente.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not adoCliente.EOF Then
        
        wdestinatario = adoCliente("ce_razao")
        wcnpjDest = adoCliente("ce_cgc")
        wiDest = adoCliente("ce_inscricaoEstadual")
        wEnderecoDest = adoCliente("ce_endereco")
        wCodMunicipio = adoCliente("ce_codigoMunicipio")
        wMunicipio = adoCliente("ce_municipio")
        wbairro = adoCliente("ce_bairro")
        wUfDest = adoCliente("ce_estado")
        wCepDest = adoCliente("ce_cep")
        wFone = adoCliente("ce_telefone")
        
        wNumero = adoCliente("ce_numero")
        wBairroDest = adoCliente("ce_bairro")
        wCodMunicipio = adoCliente("ce_codIGOMunicipio")
        wComplemento = adoCliente("ce_Complemento")
        
    Else
        MsgBox "Cliente não Encontrado.", vbCritical + vbOKOnly, "Atenção"
        Exit Function
    End If
    
    adoCliente.Close

    txtDestinatario.Text = wdestinatario
    'txtDestinatario.Enabled = False
    txtCNPJDestinatario.Text = wcnpjDest
    'txtCNPJDestinatario.Enabled = False
    txtInscricaoDestinatario.Text = wiDest
    'txtInscricaoDestinatario.Enabled = False
    txtEnderecoDestinatario.Text = wEnderecoDest
    'txtEnderecoDestinatario.Enabled = False
    txtMunicipioDestinatario.Text = wMunicipio
'    txtMunicipioDestinatario.Enabled = False
    txtBairroDestinatario.Text = wbairro
    txtUFDestinatario.Text = wUfDest
    'txtUFDestinatario.Enabled = False
    txtCepDestinatario.Text = wCepDest
    'txtCepDestinatario.Enabled = False
    txtFoneFaxDestinatario.Text = wFone
    'txtFoneFaxDestinatario.Enabled = False
    txtCodigoMunicipio.Text = wCodMunicipio
    'txtCodigoMunicipio.Enabled = False
    txtComplemento.Text = wComplemento
    txtNroEnd.Text = wNumero
    'txtComplemento.Enabled = False
    
    Call validaDados(CDbl(CodCliente))
    
End Function

Public Function CarregaLojaDestino(CodCliente As String)
    
    Dim adoCliente As New ADODB.Recordset
    
    Dim sql As String
    Dim wdestinatario As String
    Dim wcnpjDest As String
    Dim wiDest As String
    Dim wEnderecoDest As String
    Dim wNumero As String
    Dim wMunicipio As String
    Dim wBairroDest As String
    Dim wUfDest As String
    Dim wCepDest As String
    Dim wFone As String
    Dim wCodMunicipio As String
    Dim wComplemento As String
    
    sql = "select lo_razao,lo_cgc,lo_inscricaoEstadual, " & vbNewLine & _
          "lo_endereco,lo_codigoMunicipio,lo_municipio," & vbNewLine & _
          "lo_bairro,lo_uf,lo_cep," & vbNewLine & _
          "lo_telefone,lo_numero,lo_bairro," & vbNewLine & _
          "lo_codIGOMunicipio " & vbNewLine & _
          "from loja " & vbNewLine & _
          "where lo_loja='" & CodCliente & "'"
          
    adoCliente.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not adoCliente.EOF Then
        
        wdestinatario = adoCliente("lo_razao")
        wcnpjDest = adoCliente("lo_cgc")
        wiDest = adoCliente("lo_inscricaoEstadual")
        wEnderecoDest = adoCliente("lo_endereco")
        wCodMunicipio = adoCliente("lo_codigoMunicipio")
        wMunicipio = adoCliente("lo_municipio")
        wbairro = adoCliente("lo_bairro")
        wUfDest = adoCliente("lo_uf")
        wCepDest = adoCliente("lo_cep")
        wFone = adoCliente("lo_telefone")
        
        wNumero = adoCliente("lo_numero")
        wBairroDest = adoCliente("lo_bairro")
        wCodMunicipio = adoCliente("lo_codIGOMunicipio")
        'wComplemento = adoCliente("lo_Complemento")
        
    Else
        MsgBox "Cliente não Encontrado.", vbCritical + vbOKOnly, "Atenção"
        Exit Function
    End If
    
    adoCliente.Close

    txtDestinatario.Text = wdestinatario
    'txtDestinatario.Enabled = False
    txtCNPJDestinatario.Text = wcnpjDest
    'txtCNPJDestinatario.Enabled = False
    txtInscricaoDestinatario.Text = wiDest
    'txtInscricaoDestinatario.Enabled = False
    txtEnderecoDestinatario.Text = wEnderecoDest
    'txtEnderecoDestinatario.Enabled = False
    txtMunicipioDestinatario.Text = wMunicipio
'    txtMunicipioDestinatario.Enabled = False
    txtBairroDestinatario.Text = wbairro
    txtUFDestinatario.Text = wUfDest
    'txtUFDestinatario.Enabled = False
    txtCepDestinatario.Text = wCepDest
    'txtCepDestinatario.Enabled = False
    txtFoneFaxDestinatario.Text = wFone
    'txtFoneFaxDestinatario.Enabled = False
    txtCodigoMunicipio.Text = wCodMunicipio
    'txtCodigoMunicipio.Enabled = False
    txtComplemento.Text = wComplemento
    txtNroEnd.Text = wNumero
    'txtComplemento.Enabled = False
    
    If optloja.Value = False Then Call validaDados(CDbl(CodCliente))
    
End Function

Public Function CarregaFornecedor(codFornecedor As String)
    
    Dim adoFornecedor As New ADODB.Recordset
    
    Dim sql As String
    Dim wdestinatario As String
    Dim wcnpjDest As String
    Dim wiDest As String
    Dim wEnderecoDest As String
    Dim wNumero As String
    Dim wMunicipio As String
    Dim wBairroDest As String
    Dim wUfDest As String
    Dim wCepDest As String
    Dim wFone As String
    Dim wCodMunicipio As String
    Dim wCodigoMunicipio As String
    Dim wComplemento As String
    
    
    sql = "select * from fornecedor where fo_codigoFornecedor='" & Mid(codFornecedor, 1, 4) & "'"
    adoFornecedor.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Not adoFornecedor.EOF Then
        
        wdestinatario = adoFornecedor("fo_razaoSocial")
        wcnpjDest = adoFornecedor("fo_cgc")
        wiDest = adoFornecedor("fo_inscricaoEstadual")
        wEnderecoDest = adoFornecedor("fo_endereco")
        wMunicipio = adoFornecedor("fo_municipio")
        wUfDest = adoFornecedor("fo_estado")
        wCepDest = adoFornecedor("fo_cep")
        wFone = adoFornecedor("fo_telefone")
        wNumero = adoFornecedor("fo_numero")
        wBairroDest = adoFornecedor("fo_bairro")
        wCodMunicipio = adoFornecedor("fo_codIGOMunicipio")
        wCodigoMunicipio = adoFornecedor("fo_codIGOMunicipio")
        wComplemento = adoFornecedor("fo_Complemento")
    Else
        
        MsgBox "Fornecedor não encontrado.", vbCritical + vbOKOnly, "Atenção"
        Exit Function
        
    End If
    
    adoFornecedor.Close

    txtDestinatario.Text = wdestinatario
    'txtDestinatario.Enabled = False
    txtCNPJDestinatario.Text = wcnpjDest
    'txtCNPJDestinatario.Enabled = False
    txtInscricaoDestinatario.Text = wiDest
    'txtInscricaoDestinatario.Enabled = False
    txtEnderecoDestinatario.Text = wEnderecoDest
    'txtEnderecoDestinatario.Enabled = False
    txtMunicipioDestinatario.Text = wMunicipio
    'txtMunicipioDestinatario.Enabled = False
    txtUFDestinatario.Text = wUfDest
    'txtUFDestinatario.Enabled = False
    txtCepDestinatario.Text = wCepDest
    'txtCepDestinatario.Enabled = False
    txtFoneFaxDestinatario.Text = wFone
    'txtFoneFaxDestinatario.Enabled = False
    txtBairroDestinatario.Text = wBairroDest
    'txtBairroDestinatario.Enabled = False
    txtNroEnd.Text = wNumero
    'txtNroEnd.Enabled = False
    txtCodigoMunicipio.Text = wCodigoMunicipio
    'txtCodigoMunicipio.Enabled = False
    txtComplemento.Text = wComplemento
    'txtComplemento.Enabled = False
    
End Function



Private Sub Label2_Click()

End Sub

Private Sub optCliente_Click()

    cmdLojaDestino.Visible = False
    txtCodigoDestinatario.Text = ""
    txtCodigoDestinatario.MaxLength = 7
    limpaDestinatario
    lblCodigoRemetente.Visible = True
    txtCodigoDestinatario.Visible = True
    txtCodigoDestinatario.Enabled = True
    
End Sub

Private Sub optFornecedor_Click()
    
    cmdLojaDestino.Visible = False
    txtCodigoDestinatario.Text = ""
    txtCodigoDestinatario.MaxLength = 4

    limpaDestinatario
    lblCodigoRemetente.Visible = True
    txtCodigoDestinatario.Visible = True
    txtCodigoDestinatario.Enabled = True

End Sub

Private Sub optInformado_Click()
    
    cmdLojaDestino.Visible = False
    txtCodigoDestinatario.Text = ""
    txtCodigoDestinatario.MaxLength = 7

    limpaDestinatario
    lblCodigoRemetente.Visible = False
    txtCodigoDestinatario.Visible = False
    txtCodigoDestinatario.Enabled = False
    txtDestinatario.Enabled = True
    txtCNPJDestinatario.Enabled = True
    txtInscricaoDestinatario.Enabled = True
    txtEnderecoDestinatario.Enabled = True
    txtMunicipioDestinatario.Enabled = True
    txtUFDestinatario.Enabled = True
    txtCepDestinatario.Enabled = True
    txtFoneFaxDestinatario.Enabled = True
    txtCodigoMunicipio.Enabled = True
    txtComplemento.Enabled = True
    
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNumerico(KeyAscii)
End Sub

Private Sub optLoja_Click()
    cmdLojaDestino.Visible = True
    cmdLojaDestino_Click
End Sub

Private Sub carregaLojasComboTransferencia()
    Dim adoLoja As New ADODB.Recordset
    
    sql = "Select LO_Loja as loja, LO_Endereco, lo_cgc from Loja where " _
    & "LO_OrdemLoja <> 888 and LO_Loja not in('CONSO','CMCS','CMCE') Order By lo_regiao"
    
    adoLoja.CursorLocation = adUseClient
    adoLoja.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    Do While Not adoLoja.EOF
        cmdLojaDestino.AddItem adoLoja("loja")
        adoLoja.MoveNext
    Loop
    
    If txtCodigoDestinatario.Text <> Empty Then
        cmdLojaDestino.ListIndex = 0
    End If
    adoLoja.Close
    
End Sub

Private Sub txtCodigoDestinatario_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        Screen.MousePointer = 11
        
        
        If optFornecedor Then
            CarregaFornecedor Trim(txtCodigoDestinatario.Text)
        ElseIf optCliente And Val(txtCodigoDestinatario.Text) > 0 Then
            CarregaCliente Trim(txtCodigoDestinatario.Text)
        End If
        
        Screen.MousePointer = 0
        
    End If
    
End Sub

Private Sub CarregaCFOP()
    
    Dim wDentroForaEstado As String
    Dim codigoNota As String
    Dim tipoPesquisaES As String
    Dim indexNota As Byte
    Dim i As Byte
    
    codigoNota = cmbCFOP.Text
    
    If txtUFEmitente.Text = "" Then txtUFEmitente.Text = "SP"
    If txtUFDestinatario.Text <> txtUFEmitente.Text Then
        wDentroForaEstado = "F"
        lblCFOP.Caption = "Código Operação (Somente Interestadual)"
    Else
        wDentroForaEstado = "D"
        lblCFOP.Caption = "Código Operação (Somente Estadual)"
    End If
    
    If cmbTipoES.ListIndex = 0 Then
        tipoPesquisaES = "<"
        TipoNota = "EA"
    Else
        tipoPesquisaES = ">"
        TipoNota = "SA"
    End If
    
    cmbCFOP.Clear
    sql = "Select CFO_Codigo as codigo, rtrim(CFO_DescricaoOperacao) as DescricaoOperacao from CFOPEntradaSaida Where " & _
    "cfo_dentroforaestado = '" & wDentroForaEstado & "' and " & vbNewLine & _
    "cfo_codigo " & tipoPesquisaES & " = 5000 " & vbNewLine & _
    " order by cfo_CODIGO"
    adoCFOP.CursorLocation = adUseClient
    adoCFOP.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic

    If Not adoCFOP.EOF Then
        Do While Not adoCFOP.EOF
            i = i + 1
            cmbCFOP.AddItem adoCFOP("Codigo") & " - " & adoCFOP("DescricaoOperacao")
            
            If adoCFOP("Codigo") = codigoNota Then
                indexNota = i
            End If
            
            adoCFOP.MoveNext
        Loop
        cmbCFOP.ListIndex = indexNota - 1
    End If
    
    adoCFOP.Close
    cmbCFOP.Enabled = True
 
End Sub

Private Sub CarregaNota()

    Dim sql As String
    Dim adoNotas As New ADODB.Recordset
    
    sql = "select * from nfcapa where numeroped = '" & txtPedido.Text & "'"
    'sql = "select * from nfcapa where numeroped = '155'"
          
    adoNotas.CursorLocation = adUseClient
    adoNotas.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic

    If Not adoNotas.EOF Then
        
        txtCodigoDestinatario.Text = adoNotas("CLIENTE")
        If txtCodigoDestinatario.Text <> "0" Then
            txtCodigoDestinatario_KeyPress 13
            
        End If
        
        'CFOP
        nf = adoNotas("nf")
        cmbCFOP.Text = Val(adoNotas("codoper"))
        cmbTipoES.AddItem "100 - Entrada"
        cmbTipoES.AddItem "500 - Saída"
        If cmbCFOP.Text < 5000 Then
            cmbTipoES.ListIndex = 0
        Else
            cmbTipoES.ListIndex = 1
        End If
        'CarregaCFOP
        
        txtPeso.Text = adoNotas("pesoLq")
        txtQuantidade.Text = adoNotas("volume")
        
        cmdTipoFrete.ListIndex = adoNotas("TipoFrete")
    
    End If
    
    adoNotas.Close
    
    
'    sql = "select top 1 CNF_Carimbo as Carimbo" & vbNewLine & _
'          "from CarimboNotaFiscal " & vbNewLine & _
'          "where CNF_numeroped = '" & wPedido & "'"
          
          'ricardo
            sql = "select top 1 CNF_Carimbo as Carimbo" & vbNewLine & _
          "from CarimboNotaFiscal " & vbNewLine & _
          "where CNF_numeroped = '155'"
          
    adoCFOP.CursorLocation = adUseClient
    adoNotas.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        If Not adoNotas.EOF Then
            txtCarimbo.Text = adoNotas("Carimbo")
        End If
    adoNotas.Close

End Sub

Private Sub limpaDestinatario()

    txtDestinatario.Text = ""
    txtCNPJDestinatario.Text = ""
    txtInscricaoDestinatario.Text = ""
    txtEnderecoDestinatario.Text = ""
    txtNroEnd.Text = ""
    txtNroEnd.Text = ""
    txtBairroDestinatario.Text = ""
    txtUFDestinatario.Text = ""
    txtCepDestinatario.Text = ""
    txtFoneFaxDestinatario.Text = ""
    txtComplemento.Text = ""
    txtCodigoMunicipio.Text = ""

End Sub






Private Sub txtPeso_LostFocus()
    If txtPeso.Text = "" Then
        txtPeso.Text = "1"
    End If
End Sub

Private Sub txtQuantidade_LostFocus()
    If txtQuantidade.Text = "" Then
        txtQuantidade.Text = "1"
    End If
End Sub

Private Sub txtUFDestinatario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CarregaCFOP
    End If
End Sub

Private Sub txtUFDestinatario_LostFocus()
    txtUFDestinatario_KeyPress 13
End Sub

Private Sub carregaDestinatario()

    Dim sql As String
    Dim rsDestinatario As New ADODB.Recordset
    
    sql = "select cliente from nfcapa where numeroped = " & txtPedido.Text
    
    rsDestinatario.CursorLocation = adUseClient
    rsDestinatario.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    
    If Not rsDestinatario.EOF Then
'        optInformado.Value = True
'        txtCodigoDestinatario.Text = ""
'        txtDestinatario.Text = rsDestinatario("NOMCLI")
'        txtCNPJDestinatario.Text = rsDestinatario("cgcCli")
'        txtInscricaoDestinatario.Text = rsDestinatario("inscriCli")
'        txtCepDestinatario.Text = rsDestinatario("cepCli")
'        txtEnderecoDestinatario.Text = rsDestinatario("endCli")
'        txtMunicipioDestinatario.Text = rsDestinatario("municipioCli")
'        txtCodigoMunicipio.Text = rsDestinatario("codMunicipioCli")
'        txtBairroDestinatario.Text = rsDestinatario("bairroCli")
'        txtUFDestinatario.Text = rsDestinatario("ufCliente")
'        txtFoneFaxDestinatario.Text = rsDestinatario("foneCli")
        CarregaCliente rsDestinatario("cliente")
    End If
    
    rsDestinatario.Close

End Sub

Private Sub txtMunicipioDestinatario_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
    Exit Sub
    ElseIf KeyAscii = 27 Then
        'Call Limpar
       ' Unload Me
    Else
        grdMunicipio2.Visible = True
        PreencheGridMunicipioPesquisa
    End If
    wPreencheInicio = False
End Sub

Private Sub PreencheGridMunicipioPesquisa()
Dim Ln As Integer
Dim rdoExisteForne As New ADODB.Recordset
Dim sql As String
Dim rsConsMuni As New ADODB.Recordset

'RICARDO
sql = "select * from fin_municipio where Mun_Nome like '%" & txtMunicipioDestinatario.Text & "%'"

If wPreencheInicio = True Then
   Exit Sub
End If

    grdMunicipio2.Rows = 0

    With grdMunicipio2
        .ExtendLastCol = True
        .OutlineBar = flexOutlineBarComplete
        .MergeCells = flexMergeSpill
        .Editable = flexEDNone
    End With


Ln = 0

        If Len(txtMunicipioDestinatario.Text) > 0 Then

          rsConsMuni.CursorLocation = adUseClient
          rsConsMuni.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic


         Else

         Exit Sub

         End If

         Do While Not rsConsMuni.EOF

         With grdMunicipio2
                     .AddItem Trim(rsConsMuni("Mun_nome")) & Chr(9) & Trim(rsConsMuni("Mun_Codigo")) & Chr(9) _
                    & Trim(rsConsMuni("Mun_UF"))
                     .IsSubtotal(.Rows - 1) = True
                     .RowOutlineLevel(.Rows - 1) = 3
                     .Cell(flexcpFontBold, .Rows - 1, 0) = False
                     .Redraw = flexRDBuffered
                End With

            rsConsMuni.MoveNext
            Ln = Ln + 1
         Loop

            Ln = Ln - 1
            Do While Ln >= 0
                grdMunicipio2.IsCollapsed(Ln) = flexOutlineCollapsed
                Ln = Ln - 1
            Loop

            If Not rsConsMuni.EOF Then
                grdMunicipio2.Visible = False
            End If

            rsConsMuni.Close
            
            
     '-------------------------------------------------------------------------------------------
            

'''sql = "SP_FIN_Ler_Codigo_Municipio_Por_Parametro '" & txtMunicipioDestinatario.Text & "'"
'''
'''
'''If wPreencheInicio = True Then
'''   Exit Sub
'''End If
'''
'''    grdMunicipio2.Rows = 0
'''
'''    With grdMunicipio2
'''        .ExtendLastCol = True
'''        .OutlineBar = flexOutlineBarComplete
'''        .MergeCells = flexMergeSpill
'''        .Editable = flexEDNone
'''    End With
'''
'''
'''Ln = 0
'''
'''        If Len(txtMunicipioDestinatario.Text) > 0 Then
'''
'''         rsConsMuni.CursorLocation = adUseClient
'''         rsConsMuni.Open sql, ADO_Cn_CD, adOpenDynamic, adLockReadOnly
'''
'''         Else
'''
'''         Exit Sub
'''
'''         End If
'''
'''         Do While Not rsConsMuni.EOF
'''
'''                With grdMunicipio2
'''                     .AddItem Trim(rsConsMuni("Mun_nome")) & Chr(9) & Trim(rsConsMuni("Mun_Codigo")) & Chr(9) _
'''                    & Trim(rsConsMuni("Mun_UF"))
'''                     .IsSubtotal(.Rows - 1) = True
'''                     .RowOutlineLevel(.Rows - 1) = 3
'''                     .Cell(flexcpFontBold, .Rows - 1, 0) = False
'''                     .Redraw = flexRDBuffered
'''                End With
'''
'''            rsConsMuni.MoveNext
'''            Ln = Ln + 1
'''         Loop
'''
'''            Ln = Ln - 1
'''            Do While Ln >= 0
'''                grdMunicipio2.IsCollapsed(Ln) = flexOutlineCollapsed
'''                Ln = Ln - 1
'''            Loop
'''
'''            If Not rsConsMuni.EOF Then
'''                grdMunicipio2.Visible = False
'''            End If
'''
'''            rsConsMuni.Close

End Sub

Private Sub PreencheGridMunicipio()
Dim rdoExisteForne2 As New ADODB.Recordset
Dim Ln As Integer

'RICARDO
    grdMunicipio2.Rows = 0

    With grdMunicipio2
        .ExtendLastCol = True
        .OutlineBar = flexOutlineBarComplete
        .MergeCells = flexMergeSpill
        .Editable = flexEDNone
    End With

Ln = 0


 Set rdoExisteForne2 = ADO_Cn_CDLocal.OpenResultset("SP_FIN_Pesquisa_Municipio", Options:=rdExecDirect)

         Do While Not rdoExisteForne2.EOF

                With grdMunicipio2
                     .AddItem Trim(rdoExisteForne2("Mun_nome")) & Chr(9) & Trim(rdoExisteForne2("Mun_Codigo")) & Chr(9) _
                    & Trim(rdoExisteForne2("Mun_UF"))
                     .IsSubtotal(.Rows) = True
                     .RowOutlineLevel(.Rows) = 3
                     .Cell(flexcpFontBold, .Rows, 0) = False
                     .Redraw = flexRDBuffered
                End With

            rdoExisteForne2.MoveNext
            Ln = Ln + 1
         Loop
            Ln = Ln - 1
            Do While Ln > 0
                grdMunicipio2.IsCollapsed(Ln) = flexOutlineCollapsed
                Ln = Ln - 1
            Loop
        rdoExisteForne2.Close

End Sub

Private Sub grdMunicipio_KeyPress(KeyAscii As Integer)

'RICARDO
    If KeyAscii = 13 Or KeyAscii = 27 Then
        grdMunicipio2.Visible = False
    End If

End Sub

Private Sub grdMunicipio2_LostFocus()

'RICARDO
    If grdMunicipio2.row < 0 Then
     Exit Sub
    Else

        txtMunicipioDestinatario.Text = UCase(grdMunicipio2.TextMatrix(grdMunicipio2.row, 0))
        grdMunicipio2.Visible = False

    End If
End Sub
Private Sub grdMunicipio2_RowColChange()
   On Error GoTo SaidaRotina
'RICARDO
    txtUFDestinatario.Text = UCase(grdMunicipio2.TextMatrix(grdMunicipio2.row, 2))
    txtCodigoMunicipio.Text = grdMunicipio2.TextMatrix(grdMunicipio2.row, 1)

SaidaRotina:
grdMunicipio2.Visible = False
    Exit Sub
End Sub
