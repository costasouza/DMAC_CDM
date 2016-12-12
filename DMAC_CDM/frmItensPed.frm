VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmItensPedido 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Consulta Itens do Pedido / Saldos"
   ClientHeight    =   7470
   ClientLeft      =   150
   ClientTop       =   1230
   ClientWidth     =   15105
   LinkTopic       =   "Form13"
   ScaleHeight     =   7470
   ScaleWidth      =   15105
   ShowInTaskbar   =   0   'False
   Begin CentroDeDistribuicao.chameleonButton cmdImprime 
      Height          =   510
      Left            =   9240
      TabIndex        =   65
      Top             =   6840
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Imprime"
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
      MICON           =   "frmItensPed.frx":0000
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
      ForeColor       =   &H00404040&
      Height          =   2055
      Left            =   5760
      TabIndex        =   44
      Top             =   4440
      Width           =   9135
      Begin VB.OptionButton Option2 
         BackColor       =   &H00505050&
         Caption         =   "Observação do CD"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   48
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtObsAlmox 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   1290
         Left            =   240
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   47
         Top             =   480
         Width           =   8700
      End
      Begin VB.TextBox txtObsForne 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00000000&
         Height          =   1290
         Left            =   240
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   46
         Top             =   480
         Width           =   8580
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00505050&
         Caption         =   "Observação do Forncecedor"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   240
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdImprime1 
      Caption         =   "Imprime"
      Height          =   480
      Left            =   9825
      TabIndex        =   43
      Top             =   11220
      Width           =   1245
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdSugestao 
      Height          =   1935
      Left            =   75
      TabIndex        =   42
      Top             =   4560
      Width           =   5610
      _cx             =   9895
      _cy             =   3413
      _ConvInfo       =   1
      Appearance      =   1
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
      ForeColor       =   -2147483640
      BackColorFixed  =   0
      ForeColorFixed  =   16777215
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   5263440
      BackColorAlternate=   3158064
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   255
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000013&
      Height          =   75
      Left            =   120
      ScaleHeight     =   15
      ScaleWidth      =   14715
      TabIndex        =   41
      Top             =   6600
      Width           =   14777
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdItemPed 
      Height          =   1620
      Left            =   75
      TabIndex        =   19
      Top             =   1200
      Width           =   14805
      _cx             =   26114
      _cy             =   2857
      _ConvInfo       =   1
      Appearance      =   1
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
      ForeColor       =   -2147483640
      BackColorFixed  =   0
      ForeColorFixed  =   16777215
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   5263440
      BackColorAlternate=   3947580
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   255
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
      Cols            =   31
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      Height          =   1395
      Left            =   75
      TabIndex        =   26
      Top             =   3000
      Width           =   14820
      Begin VB.TextBox txtSemana 
         BackColor       =   &H00A3A3A3&
         Height          =   315
         Left            =   13320
         TabIndex        =   40
         Top             =   330
         Width           =   1245
      End
      Begin VB.TextBox txtDtPed 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   1575
         TabIndex        =   39
         Top             =   345
         Width           =   960
      End
      Begin VB.TextBox txtDespFin 
         BackColor       =   &H00A3A3A3&
         Height          =   315
         Left            =   12615
         TabIndex        =   38
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtDespFrete 
         BackColor       =   &H00A3A3A3&
         Height          =   315
         Left            =   11130
         TabIndex        =   37
         Top             =   960
         Width           =   1440
      End
      Begin VB.TextBox txtDespEmb 
         BackColor       =   &H00A3A3A3&
         Height          =   315
         Left            =   9660
         TabIndex        =   36
         Top             =   960
         Width           =   1425
      End
      Begin VB.TextBox txtCondPag 
         BackColor       =   &H00A3A3A3&
         Height          =   300
         Left            =   180
         TabIndex        =   35
         Top             =   960
         Width           =   6405
      End
      Begin VB.TextBox txtNaturezaOperacao 
         BackColor       =   &H00A3A3A3&
         Height          =   300
         Left            =   10230
         TabIndex        =   34
         Top             =   345
         Width           =   3045
      End
      Begin VB.TextBox txtCodComp 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   4905
         TabIndex        =   33
         Top             =   345
         Width           =   1665
      End
      Begin VB.TextBox txtFilial 
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
         Height          =   300
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   585
      End
      Begin VB.TextBox txtEmpresaC 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   2580
         TabIndex        =   32
         Top             =   345
         Width           =   2280
      End
      Begin VB.TextBox txtNumPed 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   825
         TabIndex        =   21
         Top             =   360
         Width           =   720
      End
      Begin VB.TextBox txtDesconto 
         BackColor       =   &H00A3A3A3&
         Height          =   300
         Left            =   8145
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   960
         Width           =   1470
      End
      Begin VB.TextBox txtTotPed 
         BackColor       =   &H00A3A3A3&
         Height          =   300
         Left            =   6630
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   960
         Width           =   1470
      End
      Begin VB.TextBox txtSubTot 
         BackColor       =   &H00A3A3A3&
         Height          =   285
         Left            =   5190
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.TextBox txtCodigoOperacao 
         BackColor       =   &H00A3A3A3&
         Height          =   300
         Left            =   9585
         TabIndex        =   28
         Top             =   345
         Width           =   600
      End
      Begin VB.TextBox txtFornecedor 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   6615
         TabIndex        =   27
         Top             =   345
         Width           =   2910
      End
      Begin VB.Label Label16 
         BackColor       =   &H00404040&
         Caption         =   "Financeira"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   12600
         TabIndex        =   64
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label10 
         BackColor       =   &H00404040&
         Caption         =   "Condição de Pagamento"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   63
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackColor       =   &H00404040&
         Caption         =   "Sub-Total"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5160
         TabIndex        =   62
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label12 
         BackColor       =   &H00404040&
         Caption         =   "Total"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6600
         TabIndex        =   61
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label13 
         BackColor       =   &H00404040&
         Caption         =   "Desconto"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8160
         TabIndex        =   60
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label14 
         BackColor       =   &H00404040&
         Caption         =   "Embalagem"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   9720
         TabIndex        =   59
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label15 
         BackColor       =   &H00404040&
         Caption         =   "Frete"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   11160
         TabIndex        =   58
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label9 
         BackColor       =   &H00404040&
         Caption         =   "Data Entrega"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   13320
         TabIndex        =   57
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H00404040&
         Caption         =   "Natureza Operação"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   10200
         TabIndex        =   56
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404040&
         Caption         =   "CFO"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   9600
         TabIndex        =   55
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label6 
         BackColor       =   &H00404040&
         Caption         =   "Fornecedor"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6600
         TabIndex        =   54
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404040&
         Caption         =   "Comprador"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4920
         TabIndex        =   53
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404040&
         Caption         =   "Empresa"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   52
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         Caption         =   "Data"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   51
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         Caption         =   "Pedido"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   50
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "Filial"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdProcurar1 
      Caption         =   "Pesquisa"
      Height          =   495
      Left            =   12360
      TabIndex        =   16
      Top             =   11205
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Height          =   1155
      Left            =   4410
      TabIndex        =   24
      Top             =   0
      Width           =   4995
      Begin VB.OptionButton optOrdenaDescricao 
         BackColor       =   &H00404040&
         Caption         =   "Ordenar por Descrição"
         ForeColor       =   &H00C0C0C0&
         Height          =   240
         Left            =   135
         TabIndex        =   10
         Top             =   855
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.OptionButton optOrdenaReferencia 
         BackColor       =   &H00404040&
         Caption         =   "Ordenar por Referência"
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Left            =   2370
         TabIndex        =   11
         Top             =   870
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.TextBox txtEmpresa 
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
         Left            =   135
         MaxLength       =   2
         TabIndex        =   8
         Top             =   450
         Width           =   660
      End
      Begin MSMask.MaskEdBox mskPesquisa 
         Height          =   315
         Left            =   870
         TabIndex        =   9
         Top             =   450
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   10724259
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label17 
         BackColor       =   &H00404040&
         Caption         =   "Label Semana Entrega"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   71
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label19 
         BackColor       =   &H00404040&
         Caption         =   "Pesquisa"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   70
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label18 
         BackColor       =   &H00404040&
         Caption         =   "Empresa"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblSemanaEntrega1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label Semana Entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1935
         TabIndex        =   25
         Top             =   225
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.Frame fraPedidos 
      BackColor       =   &H00404040&
      Caption         =   "Pedidos"
      ForeColor       =   &H00FFFFFF&
      Height          =   1155
      Left            =   9465
      TabIndex        =   23
      Top             =   0
      Width           =   5445
      Begin VB.OptionButton optPedido 
         BackColor       =   &H00404040&
         Caption         =   "Todos"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   3
         Left            =   4305
         TabIndex        =   15
         Top             =   525
         Width           =   795
      End
      Begin VB.OptionButton optPedido 
         BackColor       =   &H00404040&
         Caption         =   "Cancelados"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   2
         Left            =   2760
         TabIndex        =   14
         Top             =   525
         Width           =   1155
      End
      Begin VB.OptionButton optPedido 
         BackColor       =   &H00404040&
         Caption         =   "Baixados"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   1
         Left            =   1410
         TabIndex        =   13
         Top             =   525
         Width           =   1080
      End
      Begin VB.OptionButton optPedido 
         BackColor       =   &H00404040&
         Caption         =   "Abertos"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   12
         Top             =   525
         Value           =   -1  'True
         Width           =   1020
      End
   End
   Begin VB.CommandButton cmdNovaConsulta1 
      Caption         =   "Nova Pesquisa"
      Height          =   495
      Left            =   11100
      TabIndex        =   17
      Top             =   11205
      Width           =   1245
   End
   Begin VB.CommandButton cmdRetorna1 
      Caption         =   "Retorna"
      Height          =   495
      Left            =   13575
      TabIndex        =   18
      Top             =   11205
      Width           =   1200
   End
   Begin VB.Frame fraPesq 
      BackColor       =   &H00404040&
      Caption         =   "Pesquisa"
      ForeColor       =   &H00FFFFFF&
      Height          =   1155
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Width           =   4260
      Begin VB.OptionButton optCGCForne 
         BackColor       =   &H00404040&
         Caption         =   "CGC Forne"
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   2730
         TabIndex        =   7
         Top             =   525
         Width           =   1185
      End
      Begin VB.OptionButton optLin 
         BackColor       =   &H00404040&
         Caption         =   "Linha"
         ForeColor       =   &H00C0C0C0&
         Height          =   300
         Left            =   1470
         TabIndex        =   4
         Top             =   495
         Width           =   750
      End
      Begin VB.OptionButton optFilial 
         BackColor       =   &H00404040&
         Caption         =   "Recebimento"
         ForeColor       =   &H00C0C0C0&
         Height          =   300
         Left            =   2730
         TabIndex        =   6
         Top             =   810
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.OptionButton optDescricao 
         BackColor       =   &H00404040&
         Caption         =   "Descrição"
         ForeColor       =   &H00C0C0C0&
         Height          =   300
         Left            =   1440
         TabIndex        =   3
         Top             =   225
         Width           =   1020
      End
      Begin VB.OptionButton optReferencia 
         BackColor       =   &H00404040&
         Caption         =   "Referência"
         ForeColor       =   &H00C0C0C0&
         Height          =   300
         Left            =   150
         TabIndex        =   5
         Top             =   765
         Width           =   1095
      End
      Begin VB.OptionButton optComprador 
         BackColor       =   &H00404040&
         Caption         =   "Comprador"
         ForeColor       =   &H00C0C0C0&
         Height          =   300
         Left            =   150
         TabIndex        =   1
         Top             =   495
         Width           =   1065
      End
      Begin VB.OptionButton optSemana 
         BackColor       =   &H00404040&
         Caption         =   "Data Entrega"
         ForeColor       =   &H00C0C0C0&
         Height          =   300
         Left            =   2730
         TabIndex        =   2
         Top             =   225
         Width           =   1470
      End
      Begin VB.OptionButton optFor 
         BackColor       =   &H00404040&
         Caption         =   "Fornecedor"
         ForeColor       =   &H00C0C0C0&
         Height          =   300
         Left            =   150
         TabIndex        =   0
         Top             =   225
         Width           =   1110
      End
   End
   Begin CentroDeDistribuicao.chameleonButton cmdNovaConsulta 
      Height          =   510
      Left            =   10680
      TabIndex        =   66
      Top             =   6840
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
      MICON           =   "frmItensPed.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdProcurar 
      Height          =   510
      Left            =   12120
      TabIndex        =   67
      Top             =   6840
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
      MICON           =   "frmItensPed.frx":0038
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
      Left            =   13560
      TabIndex        =   68
      Top             =   6840
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
      MICON           =   "frmItensPed.frx":0054
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
Attribute VB_Name = "frmItensPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Pedido1 As New ADODB.Recordset


Dim wDBForne As String
Dim Texto As TextBox
Dim SitPed As String
Dim CFO As String
Dim Linha As Integer
Dim Pagina As Integer

Dim SqlNomeFornecedor As New ADODB.Recordset

Dim rdoItens As New ADODB.Recordset

Dim rdosugestao As New ADODB.Recordset

Dim rdoSubtotal As New ADODB.Recordset
Dim SemanaAtu As String

Dim wNomeFantasia As String
Dim wCodFornecedor As String
Dim wOrderby As String
Dim wPrecoLiquido As Double
Dim wTotaldaReferencia As Double
Dim SubTot As Double
Dim wDataEntregaInic As Date
Dim wDataEntregaFiM As Date

Dim wTotalPedido As Double

Private Sub cmdImprime_Click()


'============================================================================================
'====                                                                                    ====
'==== IMPORTANTE -->  ESTA TELA TEM QUE SER IGUAL NO SUPRIMENTOS E NO GERLOJA            ====
'====                 APÓS ALGUMA ALTERACAO, COPIE A MESMA PARA OS DOIS PROJETOS         ====
'====                                                                                    ====
'============================================================================================




   'Escondendo as colunas que não serão Impressas
   
   grdItemPed.ColHidden(8) = True
   'grdItemPed.ColHidden(9) = True
   grdItemPed.ColHidden(10) = True
   grdItemPed.ColHidden(11) = True
   grdItemPed.ColHidden(12) = True
   grdItemPed.ColHidden(13) = True
   grdItemPed.ColHidden(14) = True
   
   'Significado dos Parametros abaixo ==> grdItemPed.PrintGrid -> Nome do Gride
                                        ' "Itens de Pedido" - > "Nomedo Relatorio "
                                        ' False -> para não abrir a caixa de dialogo da impressora,
                                        ' 1 = retrato ou  2 = paisagem,
                                        ' 300 -> tamanho da margem esquerda e direita,
                                        ' 500 -> tamanho da margem do Topo e do Final da Pagina
                         
   grdItemPed.PrintGrid "Itens de Pedido", False, 1, 50, 500
   
   
   'Voltando ao normal as colunas após impressão
   
   grdItemPed.ColHidden(8) = False
   'grdItemPed.ColHidden(9) = False
   grdItemPed.ColHidden(10) = False
   grdItemPed.ColHidden(11) = False
   grdItemPed.ColHidden(12) = False
   grdItemPed.ColHidden(13) = False
   grdItemPed.ColHidden(14) = False

 
End Sub



Private Sub cmdNovaConsulta_Click()

'============================================================================================
'====                                                                                    ====
'==== IMPORTANTE -->  ESTA TELA TEM QUE SER IGUAL NO SUPRIMENTOS E NO GERLOJA            ====
'====                 APÓS ALGUMA ALTERACAO, COPIE A MESMA PARA OS DOIS PROJETOS         ====
'====                                                                                    ====
'============================================================================================

   
   fraPesq.Enabled = True
   fraPedidos.Enabled = True
   mskPesquisa.Width = 3900
   Limpa
  
   cmdProcurar.Enabled = True
   optOrdenaDescricao.Value = True
   optPedido(0).Value = True
   optOrdenaDescricao.Visible = False
   optOrdenaReferencia.Visible = False
   
End Sub

Sub Limpa()
grdItemPed.Rows = 1
txtEmpresa.Text = ""
 
mskPesquisa.Mask = ""
mskPesquisa.Text = ""
   
txtFilial.Text = ""
txtNumPed.Text = ""
txtDtPed.Text = ""
txtEmpresaC.Text = ""
txtCodComp.Text = ""
txtFornecedor.Text = ""
txtCodigoOperacao.Text = ""
txtNaturezaOperacao.Text = ""
txtSemana.Text = ""
txtCondPag.Text = ""
txtSubTot.Text = ""
txtTotPed.Text = ""
txtDesconto.Text = ""
txtDespEmb.Text = ""
txtDespFrete.Text = ""
txtDespFin.Text = ""
txtObsForne.Text = ""
txtObsAlmox.Text = ""

optFor.Value = True
optFor.SetFocus

End Sub

Private Sub cmdProcurar_Click()



'============================================================================================
'====                                                                                    ====
'==== IMPORTANTE -->  ESTA TELA TEM QUE SER IGUAL NO SUPRIMENTOS E NO GERLOJA            ====
'====                 APÓS ALGUMA ALTERACAO, COPIE A MESMA PARA OS DOIS PROJETOS         ====
'====                                                                                    ====
'============================================================================================


    
    Dim wcti As Long
    Dim wLin1 As String
    Dim wLin2 As String
    Dim wLin3 As String
    Dim wLin4 As String
    Dim witens As Long
    Dim wvlr As Long
    
    
    Screen.MousePointer = 11
'    status "Aguarde, efetuando pesquisa..."
    
   
    ClausulaWhere = ""
    wTotalPedido = 0
    
    If optFor.Value = True Then
      If Len(Trim(mskPesquisa.Text)) > 0 Then
       If Len(Trim(mskPesquisa.Text)) = 3 Then
          ClausulaWhere = " FO_CodigoFornecedor = " & Val(mskPesquisa.Text)
       End If
      End If
    End If
    
    If optCGCForne.Value = True Then
       If Len(Trim(mskPesquisa.Text)) > 0 Then
          ClausulaWhere = " FO_CGC = '" & Trim(mskPesquisa.Text) & "'"
       End If
    End If
    
    If ClausulaWhere <> "" Then
      sql = "SELECT  FO_CodigoFornecedor,FO_Nomefantasia from Fornecedor where " & ClausulaWhere
                 
            SqlNomeFornecedor.CursorLocation = adUseServer
            SqlNomeFornecedor.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockBatchOptimistic
       If Not SqlNomeFornecedor.EOF Then
         
       End If
    End If
       
    If mskPesquisa.Text <> "" Then
       If Val(txtEmpresa.Text) = 0 Then
          Screen.MousePointer = 0
          status "Pronto"
          MsgBox "Preencha a Empresa.", vbExclamation, "Atenção"
          txtEmpresa.SetFocus
          Exit Sub
       End If
       
       ClausulaWhere = ""
       SqlNomeFornecedor.Close
       
       If optLin.Value = True Then
          If Len(Trim(mskPesquisa.Text)) > 0 Then
             ClausulaWhere = "PR_Linhaproduto= " & mskPesquisa.Text & " AND PI_Filial <>'800' AND"
          End If
       End If
       
       If optFor.Value = True Then
          If Len(Trim(mskPesquisa.Text)) > 0 Then
             If Len(Trim(mskPesquisa.Text)) = 3 Then
                ClausulaWhere = ClausulaWhere & " PC_CodigoFornecedor = " & Val(mskPesquisa.Text) & " And SubString(PI_Referencia,1,3) = '" & mskPesquisa.Text & "' AND PI_Filial <>'800' AND"
             Else
                ClausulaWhere = ClausulaWhere & " PC_CodigoFornecedor = " & Val(mskPesquisa.Text) & " AND PI_Filial <>'800' AND"
             End If
          End If
       End If
       
            
'      wDataEntregaInic = DateAdd("d", -2, Format(mskPesquisa.Text, "yyyy/mm/dd"))
'      wDataEntregaFiM = DateAdd("d", 2, Format(mskPesquisa.Text, "yyyy/mm/dd"))
    
       If optSemana.Value = True Then
          If Len(Trim(mskPesquisa.Text)) > 0 Then
             wDataEntregaInic = DateAdd("d", -2, Format(mskPesquisa.Text, "yyyy/mm/dd"))
             wDataEntregaFiM = DateAdd("d", 2, Format(mskPesquisa.Text, "yyyy/mm/dd"))
             
             ClausulaWhere = ClausulaWhere & " PC_DataEntrega between '" & Format(wDataEntregaInic, "yyyy/mm/dd") & "' and '" & Format(wDataEntregaFiM, "yyyy/mm/dd") & "' AND PI_Filial <>'800' AND"
            
          End If
       End If
       
       If optComprador.Value = True Then
          If Len(Trim(mskPesquisa.Text)) > 0 Then
             'If mskPesquisa.Text = 2 Then
             '   ClausulaWhere = ClausulaWhere & " PC_CodigoComprador=  " & Val(mskPesquisa.Text) & " AND PI_Filial ='800' AND"
             'Else
                ClausulaWhere = ClausulaWhere & " PC_CodigoComprador=  " & Val(mskPesquisa.Text) & " AND PI_Filial <>'800' AND"
             'End If
          End If
       End If
       
       If optReferencia.Value = True Then
          If Len(Trim(mskPesquisa.Text)) > 0 Then
             ClausulaWhere = ClausulaWhere & " PI_Referencia = '" & Trim(mskPesquisa.Text) & "' AND PI_Filial <>'800' AND"
          End If
       End If
       
       If optDescricao.Value = True Then
          If Len(Trim(mskPesquisa.Text)) > 0 Then
             ClausulaWhere = ClausulaWhere & " PR_Descricao like '%" & mskPesquisa.Text & "%' AND PI_Filial <>'800' AND"
          End If
       End If
       
       If optFilial.Value = True Then
          If Len(Trim(mskPesquisa.Text)) > 0 Then
             ClausulaWhere = ClausulaWhere & " PC_DataEntrega <= " _
            & right(SemanaAtu, 4) & left(SemanaAtu, 2) & " and PI_Situacao = 'A' and PC_CodigoFornecedor = " & Val(mskPesquisa.Text) & " AND PI_Filial <>'800' AND "
          End If
       End If
       
       If optCGCForne.Value = True Then
          If Len(Trim(mskPesquisa.Text)) > 0 Then
             ClausulaWhere = ClausulaWhere & " PC_CodigoFornecedor = FO_CodigoFornecedor AND FO_CGC = '" & Trim(mskPesquisa.Text) & "' AND PI_Filial <>'800'   "
             wDBForne = ", Fornecedor"
          End If
       End If
       
       ClausulaWhere = left(ClausulaWhere, Len(ClausulaWhere) - 3)
          
       VerificaOption
           
      If optPedido(3).Value = True Then
          If optOrdenaDescricao.Value = True And optOrdenaDescricao.Visible = True Then
             wOrderby = " order by PC_Filial, PR_Descricao"
             SitPed = ""
           
          ElseIf optOrdenaReferencia.Value = True And optOrdenaReferencia.Visible = True Then
            wOrderby = "  order by PC_Filial, PI_Referencia"
            SitPed = ""
           
          ElseIf optSemana.Value = True Then
                 wOrderby = "  order by PC_DataEntrega"
                 SitPed = ""
           Else
           wOrderby = "  order by PC_Filial, PC_NumeroPedido"
           SitPed = ""
           End If
       Else
          
          
          If optOrdenaDescricao.Value = True And optOrdenaDescricao.Visible = True Then
            wOrderby = "  order by PC_Filial, PR_Descricao"
           
            
          ElseIf optOrdenaReferencia.Value = True And optOrdenaReferencia.Visible = True Then
            wOrderby = " order by PC_Filial, PI_Referencia"
             
          ElseIf optSemana.Value = True Then
                 wOrderby = "  order by PC_DataEntrega"
          Else
            
             wOrderby = " order by PC_Filial, PC_DataEntrega"
           
          End If
            
      End If
   
          sql = "SELECT PC_Filial, PC_Dataentrega,PC_CodigoFornecedor,FO_CodigoFornecedor,FO_NomeFantasia, NO_codigonatureza,NO_descricao ,CP_CodigoCondicao, CP_Descricao, " _
                             & " EM_CodigoEmpresa, EM_Descricao,  PC_NumeroPedido, PC_CodigoComprador, " _
                             & " CO_Nome, convert(varchar(3), PC_CodigoOperacao) + ' - ' + CF_Descricao as CFO, PC_DataPedido, " _
                             & " PI_Referencia, PI_SaldoPedido, PI_PrecoUnitario, PI_AliquotaIPI, PC_Desconto, PC_DespesaEmbalagem," _
                             & " PC_DespesaFrete, PC_DespesaFinanceira, PI_Situacao, PR_Descricao,PC_Empresa,PC_NaturezaOperacao, " _
                             & " PC_CondicaoPagamento,PC_TotalPedido, PC_ObservacaoForne,PC_ObservacaoAlmox  " _
                             & " From ITEMPEDIDO, PRODUTO, CAPAPEDIDO, CodigoOperacao, Comprador,Empresa,CondicaoPagto, " _
                             & " naturezaoperacao,fornecedor " _
                             & " WHERE " _
                             & " PC_NumeroPedido   = PI_NumeroPedido And " _
                             & " PC_CodigoOperacao = CF_CodigoOperacao and " _
                             & " NO_Codigooperacao = PC_CodigoOperacao And " _
                             & " no_codigonatureza = pc_naturezaoperacao and  " _
                             & " PC_CondicaoPagamento = CP_CodigoCondicao and " _
                             & " PC_Empresa = EM_CodigoEmpresa and " _
                             & " PC_CodigoComprador = CO_Codigocomprador  and " _
                             & " PC_CodigoFornecedor = FO_CodigoFornecedor and " _
                             & " PR_Referencia = PI_Referencia and CP_VendaCompra = 'C' " _
                             & " And PC_Empresa = " & Val(txtEmpresa.Text) & SitPed & " And " & ClausulaWhere & wOrderby
                             
          
      
            rdoItens.CursorLocation = adUseServer
            rdoItens.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockBatchOptimistic
  Do While Not rdoItens.EOF

      grdItemPed.AddItem rdoItens("PC_Filial") & Chr(9) _
             & rdoItens("PC_NumeroPedido") & Chr(9) _
             & Format(rdoItens("PI_Situacao")) & Chr(9) _
             & 0 & Chr(9) _
             & rdoItens("pi_referencia") & Chr(9) _
             & rdoItens("PR_Descricao") & Chr(9) _
             & Format(rdoItens("PI_SaldoPedido"), "#,##0.00") & Chr(9) _
             & Format(rdoItens("PI_PrecoUnitario"), "#,##0.00") & Chr(9) _
             & Format(rdoItens("PI_AliquotaIPI"), "#,##0.00") & Chr(9) _
             & rdoItens("PC_DataPedido") & Chr(9) _
             & Format(rdoItens("PC_Desconto"), "#,##0.00") & Chr(9) _
             & Format(rdoItens("PC_DespesaEmbalagem"), "#,##0.00") & Chr(9) _
             & Format(rdoItens("PC_DespesaFrete"), "#,##0.00") & Chr(9) _
             & Format(rdoItens("PC_DespesaFinanceira"), "#,##0.00") & Chr(9) _
             & Mid(rdoItens("cfo"), 1, 3) & Chr(9) _
             & rdoItens("PC_Empresa") & Chr(9) _
             & rdoItens("CO_Nome") & Chr(9) _
             & rdoItens("FO_NomeFantasia") & Chr(9) _
             & rdoItens("PC_NaturezaOperacao") & Chr(9) _
             & rdoItens("PC_CondicaoPagamento") & Chr(9) _
             & Format(SubTot, "#,##0.00") & Chr(9) _
             & Format(rdoItens("PC_TotalPedido"), "#,##0.00") & Chr(9) _
             & rdoItens("PC_ObservacaoForne") & Chr(9) & rdoItens("PC_ObservacaoAlmox") & Chr(9) & 0 & Chr(9) _
             & rdoItens("EM_Descricao") & Chr(9) & rdoItens("PC_CodigoComprador") & Chr(9) & rdoItens("FO_CodigoFornecedor") & Chr(9) & rdoItens("CP_Descricao") & Chr(9) & rdoItens("NO_descricao") & Chr(9) & rdoItens("PC_DataEntrega")
             
             
             wTotalPedido = wTotalPedido + (rdoItens("PI_SaldoPedido") * Format(rdoItens("PI_PrecoUnitario"), "#,##0.00"))
             
             rdoItens.MoveNext
      
      Loop
      
      If grdItemPed.Rows > 1 Then
         grdItemPed.AddItem ""
         grdItemPed.AddItem Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & _
                            "TOTAL ITEM" & _
                            Chr(9) & Chr(9) & Format(wTotalPedido, "#,##0.00")
                                 
         PintaGridItemPedido grdItemPed
      
      End If
      
       witens = grdItemPed.Rows - 1
       
       Screen.MousePointer = 0
  
    Else
       Screen.MousePointer = 0

       
       If mskPesquisa.Visible = True And mskPesquisa.Text = "" Then
          MsgBox "Favor preencher o campo Pesquisa.", vbExclamation, "Atenção"
          mskPesquisa.SetFocus
          Exit Sub
       End If
    End If
    
    cmdProcurar.Enabled = False
    fraPesq.Enabled = False
    fraPedidos.Enabled = False


End Sub

Sub VerificaOption()
   
   If optPedido(0).Value = True Then
      
       SitPed = " And PI_Situacao= 'A' "
   
   ElseIf optPedido(1).Value = True Then
      
       SitPed = " And PI_Situacao= 'B' "
   ElseIf optPedido(2).Value = True Then
      
      SitPed = " And PI_Situacao= 'C' "
   End If

End Sub

Private Sub cmdRetorna_Click()
    
    Unload Me

End Sub

Private Sub Form_Activate()
    carregarPosicaoTamanhoTela Me
    SemanaAtu = Date
End Sub
Private Sub Form_Load()
    
    'carregarPosicaoTamanhoTela Me

'============================================================================================
'====                                                                                    ====
'==== IMPORTANTE -->  ESTA TELA TEM QUE SER IGUAL NO SUPRIMENTOS E NO GERLOJA            ====
'====                 APÓS ALGUMA ALTERACAO, COPIE A MESMA PARA OS DOIS PROJETOS         ====
'====                                                                                    ====
'============================================================================================


'Skin1.LoadSkin App.Path & "\skin.skn"
'Skin1.ApplySkin Me.hwnd

txtObsForne.ZOrder

   frmItensPedido.top = (Screen.Height - frmItensPedido.Height) / 2
   frmItensPedido.left = (Screen.Width - frmItensPedido.Width) / 2
   
   txtObsAlmox.Height = txtObsForne.Height
   txtObsAlmox.left = txtObsForne.left
   txtObsAlmox.top = txtObsForne.top
   txtObsAlmox.Width = txtObsForne.Width
   
   Screen.MousePointer = 11
   
   GetAsyncKeyState vbKeyTab
   Screen.MousePointer = 0
 
End Sub

Private Sub grdItemPed_EnterCell()




'============================================================================================
'====                                                                                    ====
'==== IMPORTANTE -->  ESTA TELA TEM QUE SER IGUAL NO SUPRIMENTOS E NO GERLOJA            ====
'====                 APÓS ALGUMA ALTERACAO, COPIE A MESMA PARA OS DOIS PROJETOS         ====
'====                                                                                    ====
'============================================================================================


 
If Trim(grdItemPed.TextMatrix(grdItemPed.Row, 0)) <> "" And grdItemPed.Row <> 0 Then

txtFilial.Text = grdItemPed.TextMatrix(grdItemPed.Row, 0)
txtNumPed.Text = grdItemPed.TextMatrix(grdItemPed.Row, 1)
txtDtPed.Text = grdItemPed.TextMatrix(grdItemPed.Row, 9)
txtEmpresaC.Text = grdItemPed.TextMatrix(grdItemPed.Row, 15) & " - " & grdItemPed.TextMatrix(grdItemPed.Row, 25)
txtCodComp.Text = grdItemPed.TextMatrix(grdItemPed.Row, 26) & " - " & grdItemPed.TextMatrix(grdItemPed.Row, 16)
txtFornecedor.Text = grdItemPed.TextMatrix(grdItemPed.Row, 27) & " - " & grdItemPed.TextMatrix(grdItemPed.Row, 17)
txtCodigoOperacao.Text = grdItemPed.TextMatrix(grdItemPed.Row, 14)
txtNaturezaOperacao.Text = grdItemPed.TextMatrix(grdItemPed.Row, 18) & " - " & grdItemPed.TextMatrix(grdItemPed.Row, 29)
txtSemana.Text = Format(grdItemPed.TextMatrix(grdItemPed.Row, 30), "DD/MM/YYYY")
txtCondPag.Text = grdItemPed.TextMatrix(grdItemPed.Row, 19) & " - " & grdItemPed.TextMatrix(grdItemPed.Row, 28)
txtSubTot.Text = grdItemPed.TextMatrix(grdItemPed.Row, 20)
txtTotPed.Text = grdItemPed.TextMatrix(grdItemPed.Row, 21)
txtDesconto.Text = grdItemPed.TextMatrix(grdItemPed.Row, 10)
txtDespEmb.Text = grdItemPed.TextMatrix(grdItemPed.Row, 11)
txtDespFrete.Text = grdItemPed.TextMatrix(grdItemPed.Row, 12)
txtDespFin.Text = grdItemPed.TextMatrix(grdItemPed.Row, 13)
txtObsForne.Text = grdItemPed.TextMatrix(grdItemPed.Row, 22)
txtObsAlmox.Text = grdItemPed.TextMatrix(grdItemPed.Row, 23)


grdSugestao.Rows = 1
 sql = "SELECT  RQC_FilialEntrega, RQC_FilialDistribuicao, RQC_SugestaoCompra,RQC_QtdeAtendida,RQC_SugestaoInformada, " _
            & " RQC_SaldoSugestao, LO_Regiao " _
            & " From RequisicaoCompras, capapedido,Loja   " _
            & " WHERE " _
            & " PC_NumeroPedido = " & grdItemPed.TextMatrix(grdItemPed.Row, 1) & " and " _
            & " RQC_NumeroRequisicao = PC_Requisicao and " _
            & " RQC_Referencia = '" & grdItemPed.TextMatrix(grdItemPed.Row, 4) & "' and " _
            & " RQC_FilialDistribuicao = LO_Loja " _
            & " Order by LO_Regiao "
                             
          

            rdosugestao.CursorLocation = adUseServer
            rdosugestao.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockBatchOptimistic
If rdosugestao.EOF = False Then
     Do While Not rdosugestao.EOF
     grdSugestao.AddItem rdosugestao("RQC_FilialEntrega") & Chr(9) _
             & rdosugestao("RQC_FilialDistribuicao") & Chr(9) _
             & rdosugestao("RQC_SugestaoInformada") & Chr(9) _
             & rdosugestao("RQC_QtdeAtendida") & Chr(9) _
             & rdosugestao("RQC_SugestaoInformada") & Chr(9)
             
             rdosugestao.MoveNext
      
      Loop

End If
End If

End Sub



Private Sub optCGCForne_Click()



'============================================================================================
'====                                                                                    ====
'==== IMPORTANTE -->  ESTA TELA TEM QUE SER IGUAL NO SUPRIMENTOS E NO GERLOJA            ====
'====                 APÓS ALGUMA ALTERACAO, COPIE A MESMA PARA OS DOIS PROJETOS         ====
'====                                                                                    ====
'============================================================================================



   optOrdenaDescricao.Visible = True
   optOrdenaReferencia.Visible = True
   lblSemanaEntrega.Visible = False
   mskPesquisa.Width = 3900
   
   mskPesquisa.Mask = ""
   mskPesquisa.Text = ""
   lblPesquisa.Caption = "CNPJ do Fornecedor"
   fraPedidos.Enabled = True

End Sub

Private Sub optComprador_Click()
   
   optOrdenaDescricao.Visible = False
   optOrdenaReferencia.Visible = False
   mskPesquisa.Width = 3900
   lblSemanaEntrega.Visible = False
   
   mskPesquisa.Mask = ""
   mskPesquisa.Text = ""
   lblPesquisa.Caption = "Comprador"
   fraPedidos.Enabled = True

End Sub

Private Sub optDescricao_Click()
   
   optOrdenaDescricao.Visible = False
   optOrdenaReferencia.Visible = False
   lblSemanaEntrega.Visible = False
   mskPesquisa.Width = 3900
   
   mskPesquisa.Mask = ""
   mskPesquisa.Text = ""
   lblPesquisa.Caption = "Descrição"
   fraPedidos.Enabled = True

End Sub

Private Sub optFilial_Click()
   
   optOrdenaDescricao.Visible = False
   optOrdenaReferencia.Visible = False
   lblSemanaEntrega.left = 900
   lblSemanaEntrega.top = 855
   lblSemanaEntrega.Visible = True
   lblSemanaEntrega.Caption = "Semana : " & left(SemanaAtu, 2) & "/" & right(SemanaAtu, 4) & " - " & CalculaPeriodo(Date)
   mskPesquisa.Width = 3900
   
   mskPesquisa.Mask = ""
   mskPesquisa.Text = ""
   lblPesquisa.Caption = "Fornecedor"
   fraPedidos.Enabled = False
   
End Sub

Private Sub optFor_Click()


'============================================================================================
'====                                                                                    ====
'==== IMPORTANTE -->  ESTA TELA TEM QUE SER IGUAL NO SUPRIMENTOS E NO GERLOJA            ====
'====                 APÓS ALGUMA ALTERACAO, COPIE A MESMA PARA OS DOIS PROJETOS         ====
'====                                                                                    ====
'============================================================================================


   
   optOrdenaDescricao.Visible = True
   optOrdenaReferencia.Visible = True
  ' lblSemanaEntrega.Visible = False
   mskPesquisa.Width = 3900
   
   mskPesquisa.Mask = ""
   mskPesquisa.Text = ""
  ' lblPesquisa.Caption = "Fornecedor"
   fraPedidos.Enabled = True

End Sub

Private Sub Option1_Click()
txtObsAlmox.Visible = False
txtObsForne.Visible = True
End Sub

Private Sub Option2_Click()
txtObsAlmox.Visible = True
txtObsForne.Visible = False
End Sub

Private Sub optLin_Click()
   
   optOrdenaDescricao.Visible = False
   optOrdenaReferencia.Visible = False
   lblSemanaEntrega.Visible = False
   mskPesquisa.Width = 3900
   
   mskPesquisa.Mask = ""
   mskPesquisa.Text = ""
   lblPesquisa.Caption = "Linha"
   fraPedidos.Enabled = True

End Sub

Private Sub optOrdenaDescricao_LostFocus()
   
   On Error Resume Next
   
   If optOrdenaDescricao.Value = True And optOrdenaDescricao.Visible = True Then
      optPedido(0).SetFocus
   End If

End Sub

Private Sub optOrdenaReferencia_LostFocus()
   
   If optOrdenaReferencia.Value = True And optOrdenaReferencia.Visible = True Then
      optPedido(0).SetFocus
   End If

End Sub

Private Sub optPedido_LostFocus(Index As Integer)
   
   If optPedido(Index).Value = True Then
      cmdProcurar.Enabled = True
      cmdProcurar.SetFocus
   End If

End Sub

Private Sub Optreferencia_Click()
   
   mskPesquisa.Mask = ""
   mskPesquisa.Text = ""
   lblPesquisa.Caption = "Referência"
   fraPedidos.Enabled = True

End Sub

Private Sub optReferencia_GotFocus()
   
   optOrdenaDescricao.Visible = False
   optOrdenaReferencia.Visible = False
   lblSemanaEntrega.Visible = False
   mskPesquisa.Width = 3900
   GetAsyncKeyState vbKeyTab

End Sub

Private Sub optSemana_Click()
   
   mskPesquisa.Mask = "##/##/####"
   lblPesquisa.Caption = "Data Entrega"
   fraPedidos.Enabled = True

End Sub

Private Sub optSemana_GotFocus()
   
   optOrdenaDescricao.Visible = False
   optOrdenaReferencia.Visible = False
   lblSemanaEntrega.Visible = False
   mskPesquisa.Width = 3900
   GetAsyncKeyState vbKeyTab

End Sub

Private Sub TabStrip1_Click()
Dim I As Integer
I = TabStrip1.SelectedItem.Index
If I = 1 Then
   txtObsForne.ZOrder
Else
   txtObsAlmox.ZOrder
End If
End Sub

Private Sub SkinLabel8_DragDrop(Source As Control, x As Single, y As Single)

End Sub

Private Sub txtEmpresa_LostFocus()
   
   If optFor.Value = True Or optCGCForne.Value = True Then
      mskPesquisa.Width = 3900
      optOrdenaDescricao.Visible = True
      optOrdenaReferencia.Visible = True
   Else
      optOrdenaDescricao.Visible = False
      optOrdenaReferencia.Visible = False
      mskPesquisa.Width = 3900
   End If

End Sub

Sub PintaGridItemPedido(ByRef NomeGrid)
Dim cor As String
    cor = &HFDCEAC
    
    NomeGrid.FillStyle = flexFillRepeat
    NomeGrid.Row = NomeGrid.Rows - 1
    NomeGrid.RowSel = NomeGrid.Rows - 1
    NomeGrid.Col = 0
    NomeGrid.ColSel = NomeGrid.Cols - 1
'    NomeGrid.CellBackColor = Cor
    NomeGrid.CellBackColor = &HFDCEAC
    NomeGrid.FillStyle = flexFillSingle
    NomeGrid.Row = 0
    NomeGrid.Col = 0
   
    NomeGrid.Cell(flexcpFontBold, NomeGrid.Rows - 1, 5) = True
    NomeGrid.Cell(flexcpFontBold, NomeGrid.Rows - 1, 7) = True
       
    NomeGrid.Cell(flexcpForeColor, NomeGrid.Rows - 1, 5) = Azul
    NomeGrid.Cell(flexcpForeColor, NomeGrid.Rows - 1, 7) = Azul
    
End Sub
