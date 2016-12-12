VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form FrmEntrega 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Entrega"
   ClientHeight    =   7920
   ClientLeft      =   600
   ClientTop       =   2235
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   120
      ScaleHeight     =   45
      ScaleWidth      =   11400
      TabIndex        =   24
      Top             =   6360
      Width           =   11400
   End
   Begin CentroDeDistribuicao.chameleonButton cmdRetornar 
      Height          =   495
      Left            =   10080
      TabIndex        =   23
      Top             =   6600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "Retornar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   5263440
      BCOLO           =   5263440
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmEntrega.frx":0000
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
      BackColor       =   &H00505050&
      Height          =   3975
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   6735
      Begin VB.TextBox txtContato 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   120
         MaxLength       =   15
         TabIndex        =   32
         ToolTipText     =   "CNPJ"
         Top             =   3480
         Width           =   6435
      End
      Begin VB.TextBox txtLocalizacao 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   120
         MaxLength       =   15
         TabIndex        =   30
         ToolTipText     =   "CNPJ"
         Top             =   2880
         Width           =   6435
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   120
         MaxLength       =   15
         TabIndex        =   26
         ToolTipText     =   "CNPJ"
         Top             =   1080
         Width           =   915
      End
      Begin VB.TextBox txtMunicipio 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   120
         MaxLength       =   15
         TabIndex        =   22
         ToolTipText     =   "CNPJ"
         Top             =   2280
         Width           =   6435
      End
      Begin VB.TextBox txtBairro 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   120
         MaxLength       =   15
         TabIndex        =   21
         ToolTipText     =   "CNPJ"
         Top             =   1680
         Width           =   6435
      End
      Begin VB.TextBox txtNumero 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   5880
         MaxLength       =   15
         TabIndex        =   18
         ToolTipText     =   "CNPJ"
         Top             =   1080
         Width           =   675
      End
      Begin VB.TextBox txtEndereco 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   16
         ToolTipText     =   "CNPJ"
         Top             =   1080
         Width           =   4755
      End
      Begin VB.TextBox txtDestinatario 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   120
         MaxLength       =   15
         TabIndex        =   14
         ToolTipText     =   "CNPJ"
         Top             =   480
         Width           =   6435
      End
      Begin VB.Label Label14 
         BackColor       =   &H00505050&
         Caption         =   "Contato:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label13 
         BackColor       =   &H00505050&
         Caption         =   "Localização:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label15 
         BackColor       =   &H00505050&
         Caption         =   "CEP"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label11 
         BackColor       =   &H00505050&
         Caption         =   "Municipio"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label10 
         BackColor       =   &H00505050&
         Caption         =   "Bairro"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label9 
         BackColor       =   &H00505050&
         Caption         =   "Numero"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5880
         TabIndex        =   17
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00505050&
         Caption         =   "Endereço"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   15
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00505050&
         Caption         =   "Destinatário"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00505050&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.TextBox txtVolume 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2880
         MaxLength       =   15
         TabIndex        =   36
         ToolTipText     =   "CNPJ"
         Top             =   1200
         Width           =   555
      End
      Begin VB.OptionButton optFretePago 
         BackColor       =   &H00505050&
         Caption         =   "Frete Pago"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5400
         TabIndex        =   35
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton optFreteaPagar 
         BackColor       =   &H00505050&
         Caption         =   "Frete a Pagar"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3720
         TabIndex        =   34
         Top             =   1200
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   28
         ToolTipText     =   "Loja"
         Top             =   480
         Width           =   4275
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   10
         ToolTipText     =   "CNPJ"
         Top             =   1200
         Width           =   1275
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   9
         ToolTipText     =   "CNPJ"
         Top             =   1200
         Width           =   435
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   120
         MaxLength       =   15
         TabIndex        =   8
         ToolTipText     =   "CNPJ"
         Top             =   1200
         Width           =   915
      End
      Begin VB.ComboBox cmbLoja 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Loja"
         Top             =   480
         Width           =   915
      End
      Begin MSMask.MaskEdBox mskDataEntrega 
         Height          =   315
         Left            =   5400
         TabIndex        =   5
         ToolTipText     =   "Data Emissão"
         Top             =   480
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
      Begin VB.Label lblVolume 
         BackColor       =   &H00505050&
         Caption         =   "Volume"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2880
         TabIndex        =   29
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00505050&
         Caption         =   "Valor Total Nota"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   11
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00505050&
         Caption         =   "Serie"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H00505050&
         Caption         =   "Nota"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label3 
         BackColor       =   &H00505050&
         Caption         =   "Data Entrega"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5400
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00505050&
         Caption         =   "Transportador"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00505050&
         Caption         =   "Loja"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdTransfAbertas 
      Height          =   5940
      Left            =   6960
      TabIndex        =   25
      Top             =   240
      Width           =   4530
      _cx             =   7990
      _cy             =   10477
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
      ForeColorFixed  =   16777215
      BackColorSel    =   16423203
      ForeColorSel    =   8388608
      BackColorBkg    =   5263440
      BackColorAlternate=   3947580
      GridColor       =   5263440
      GridColorFixed  =   -2147483632
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
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmEntrega.frx":001C
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
   End
End
Attribute VB_Name = "FrmEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdRetornar_Click()
    Unload Me
End Sub



Private Sub Form_Load()
    carregarPosicaoTamanhoTela Me
End Sub

