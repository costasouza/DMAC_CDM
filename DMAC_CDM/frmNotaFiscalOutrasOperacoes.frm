VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7d.ocx"
Begin VB.Form frmNotaFiscalOutrasOperacoes 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Nota Fiscal Outras Operações"
   ClientHeight    =   7980
   ClientLeft      =   24420
   ClientTop       =   1650
   ClientWidth     =   15300
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   7980
   ScaleWidth      =   15300
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraImportar 
      BackColor       =   &H00393939&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   3030
      Left            =   15480
      TabIndex        =   89
      Top             =   255
      Visible         =   0   'False
      Width           =   2730
      Begin VB.TextBox txtSerienf 
         BackColor       =   &H00C0C0C0&
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
         Left            =   1395
         TabIndex        =   95
         Top             =   675
         Width           =   1170
      End
      Begin VB.TextBox txtCnpj 
         BackColor       =   &H00C0C0C0&
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
         Left            =   150
         TabIndex        =   96
         Top             =   1365
         Width           =   2415
      End
      Begin VB.TextBox txtNf 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   94
         Top             =   675
         Width           =   1170
      End
      Begin CentroDeDistribuicao.chameleonButton cmdImpotar 
         Height          =   330
         Left            =   150
         TabIndex        =   97
         Top             =   2535
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         BTYPE           =   14
         TX              =   "&Importar"
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
         MICON           =   "frmNotaFiscalOutrasOperacoes.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CentroDeDistribuicao.chameleonButton cdmCancel 
         Height          =   330
         Left            =   1395
         TabIndex        =   98
         Top             =   2535
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         BTYPE           =   14
         TX              =   "&Cancelar"
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
         MICON           =   "frmNotaFiscalOutrasOperacoes.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblDescFornecedor 
         BackStyle       =   0  'Transparent
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
         Height          =   600
         Left            =   150
         TabIndex        =   99
         Top             =   1815
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.Label lblFornec 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ do Fornecedor"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   150
         TabIndex        =   93
         Top             =   1125
         Width           =   1485
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Série"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   1395
         TabIndex        =   92
         Top             =   435
         Width           =   360
      End
      Begin VB.Label lblNF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nota Fiscal"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   150
         TabIndex        =   91
         Top             =   435
         Width           =   795
      End
      Begin VB.Label lblImportar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Importar Nota"
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
         TabIndex        =   90
         Top             =   150
         Width           =   1170
      End
   End
   Begin VB.Frame frameCancelamento 
      BackColor       =   &H00393939&
      BorderStyle     =   0  'None
      Caption         =   "?"
      Height          =   3030
      Left            =   15945
      TabIndex        =   76
      Top             =   3810
      Visible         =   0   'False
      Width           =   2730
      Begin VB.TextBox txtNotaFiscal 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   150
         MaxLength       =   10
         TabIndex        =   77
         ToolTipText     =   "Código do Produto"
         Top             =   675
         Width           =   1170
      End
      Begin VB.TextBox txtSerie 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1395
         MaxLength       =   3
         TabIndex        =   79
         ToolTipText     =   "Referência"
         Top             =   675
         Width           =   1170
      End
      Begin CentroDeDistribuicao.chameleonButton cmdCancelar 
         Height          =   330
         Left            =   150
         TabIndex        =   81
         Top             =   2535
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         BTYPE           =   14
         TX              =   "&Cancelar"
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
         MICON           =   "frmNotaFiscalOutrasOperacoes.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CentroDeDistribuicao.chameleonButton cmdSairCancelamento 
         Height          =   330
         Left            =   1395
         TabIndex        =   83
         Top             =   2535
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
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
         MICON           =   "frmNotaFiscalOutrasOperacoes.frx":0054
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
         Caption         =   "Serie"
         ForeColor       =   &H00C0C0C0&
         Height          =   300
         Left            =   1395
         TabIndex        =   84
         Top             =   435
         Width           =   1170
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Nota Fiscal"
         ForeColor       =   &H00C0C0C0&
         Height          =   210
         Left            =   150
         TabIndex        =   82
         Top             =   435
         Width           =   1470
      End
      Begin VB.Label lblinfoNota 
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         ForeColor       =   &H00C0C0C0&
         Height          =   990
         Left            =   150
         TabIndex        =   80
         Top             =   1170
         Width           =   2235
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelar Nota"
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
         Left            =   150
         TabIndex        =   78
         Top             =   150
         Width           =   2070
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Pesquisa"
      ForeColor       =   &H00FF0000&
      Height          =   465
      Left            =   120
      TabIndex        =   73
      Top             =   120
      Width           =   2385
      Begin VB.TextBox txtPedido 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1095
         MaxLength       =   7
         TabIndex        =   75
         Top             =   75
         Width           =   1215
      End
      Begin VB.Label Label3 
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
         TabIndex        =   74
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "lblQtd"
      ForeColor       =   &H00FF0000&
      Height          =   825
      Left            =   120
      TabIndex        =   54
      Top             =   5835
      Width           =   14880
      Begin VB.TextBox txtTotalNota 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   13020
         TabIndex        =   31
         Top             =   360
         Width           =   1755
      End
      Begin VB.TextBox txtTotalIPI 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   11730
         TabIndex        =   30
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtOutrasDesp 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   10440
         TabIndex        =   29
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtDesconto 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   9150
         TabIndex        =   28
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtValSeguro 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   7860
         TabIndex        =   27
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtValFrete 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   6570
         TabIndex        =   26
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtValTotalMerc 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   5280
         TabIndex        =   25
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtValICMSSub 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   3990
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtBaseCalcICMS 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtBaseCalcICMSST 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2700
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtValICMS 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1410
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lbBaseCalcICMS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base Calc. ICMS"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   120
         TabIndex        =   65
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label lbValICMS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Val. ICMS"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   1410
         TabIndex        =   64
         Top             =   120
         Width           =   705
      End
      Begin VB.Label lbBaseCalcICMSST 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base ICMS ST"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   2700
         TabIndex        =   63
         Top             =   120
         Width           =   1050
      End
      Begin VB.Label lbValICMSSubs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Val. ICMS Subs."
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   3990
         TabIndex        =   62
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label lbValTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Val. Total Merc. "
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   5280
         TabIndex        =   61
         Top             =   120
         Width           =   1170
      End
      Begin VB.Label lbVlrFrete 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Val. Frete"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   6570
         TabIndex        =   60
         Top             =   120
         Width           =   675
      End
      Begin VB.Label lbVlrSeguro 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Val. Seguro"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   7860
         TabIndex        =   59
         Top             =   120
         Width           =   825
      End
      Begin VB.Label lbDesconto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   9150
         TabIndex        =   58
         Top             =   120
         Width           =   690
      End
      Begin VB.Label lbOutrasDesp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Outras Desp."
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   10440
         TabIndex        =   57
         Top             =   120
         Width           =   930
      End
      Begin VB.Label lbTotalIPI 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Val. Total IPI"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   11730
         TabIndex        =   56
         Top             =   120
         Width           =   915
      End
      Begin VB.Label lbValTotalNota 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Val. Total Nota"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   13020
         TabIndex        =   55
         Top             =   120
         Width           =   1065
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   43
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   14880
      TabIndex        =   33
      Top             =   6810
      Width           =   14880
   End
   Begin VB.Frame frmNotaFiscalOutrasOperacoes 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "lblQtd"
      ForeColor       =   &H00FF0000&
      Height          =   1440
      Left            =   120
      TabIndex        =   20
      Top             =   2505
      Width           =   14880
      Begin VB.TextBox txtBaseIPI 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   5280
         TabIndex        =   14
         Text            =   "0,00"
         Top             =   975
         Width           =   1215
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Text            =   "0,00"
         Top             =   975
         Width           =   1215
      End
      Begin VB.TextBox txtBaseAliqICMSST 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   9150
         TabIndex        =   8
         Text            =   "0,00"
         Top             =   975
         Width           =   1215
      End
      Begin VB.TextBox txtAliquotaICMSST 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   10440
         TabIndex        =   9
         Text            =   "0,00"
         Top             =   975
         Width           =   1215
      End
      Begin VB.TextBox txtValICMSST 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   11730
         TabIndex        =   10
         Text            =   "0,00"
         Top             =   975
         Width           =   1215
      End
      Begin VB.ComboBox cmbCST 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   7860
         TabIndex        =   6
         Top             =   360
         Width           =   5085
      End
      Begin VB.TextBox txtAliquotaIPI 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   6570
         TabIndex        =   15
         Text            =   "0,00"
         Top             =   975
         Width           =   1215
      End
      Begin VB.TextBox txtNCM 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   13020
         TabIndex        =   17
         Top             =   975
         Width           =   1725
      End
      Begin VB.TextBox txtAliquotaICMS 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2700
         TabIndex        =   12
         Text            =   "0,00"
         Top             =   975
         Width           =   1215
      End
      Begin VB.TextBox txtValorICMS 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   3990
         TabIndex        =   13
         Text            =   "0,00"
         Top             =   975
         Width           =   1215
      End
      Begin VB.TextBox txtDescricao 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1410
         TabIndex        =   3
         Top             =   360
         Width           =   4380
      End
      Begin VB.TextBox txtReferencia 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         MaxLength       =   7
         TabIndex        =   2
         Text            =   "#######"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtBaseICMS 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1410
         TabIndex        =   11
         Text            =   "0,00"
         Top             =   975
         Width           =   1215
      End
      Begin VB.TextBox txtValorIPI 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   7860
         TabIndex        =   16
         Text            =   "0,00"
         Top             =   975
         Width           =   1215
      End
      Begin VB.TextBox txtPreco 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   6570
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtQtde 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
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
         Left            =   5865
         TabIndex        =   4
         Text            =   "1"
         Top             =   360
         Width           =   630
      End
      Begin CentroDeDistribuicao.chameleonButton cmdImpostoSISTEMA 
         Height          =   330
         Left            =   13020
         TabIndex        =   86
         Top             =   360
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   582
         BTYPE           =   14
         TX              =   "Carregar"
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
         MICON           =   "frmNotaFiscalOutrasOperacoes.frx":0070
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Impostos Sistema"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   13020
         TabIndex        =   87
         Top             =   120
         Width           =   1230
      End
      Begin VB.Label lbBaseIPI 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base IPI"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   5280
         TabIndex        =   71
         Top             =   750
         Width           =   600
      End
      Begin VB.Label lbdesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   120
         TabIndex        =   69
         Top             =   750
         Width           =   690
      End
      Begin VB.Label lbBaseAliquotaST 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base ICMS ST"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   9150
         TabIndex        =   68
         Top             =   750
         Width           =   1050
      End
      Begin VB.Label lbAliquotaICMSST 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aliq. ICMS ST"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   10440
         TabIndex        =   67
         Top             =   750
         Width           =   990
      End
      Begin VB.Label lbValICMSST 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor ICMS ST"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   11730
         TabIndex        =   66
         Top             =   750
         Width           =   1050
      End
      Begin VB.Label lbCST 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CST"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   7860
         TabIndex        =   52
         Top             =   120
         Width           =   315
      End
      Begin VB.Label lbNCM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NCM"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   13020
         TabIndex        =   51
         Top             =   750
         Width           =   360
      End
      Begin VB.Label lbAliqutaICMS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alíquota ICMS"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   2700
         TabIndex        =   50
         Top             =   750
         Width           =   1035
      End
      Begin VB.Label lbVlrICMS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor ICMS"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   3990
         TabIndex        =   49
         Top             =   750
         Width           =   795
      End
      Begin VB.Label lbDescricao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   1410
         TabIndex        =   48
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lbReferencia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Referência"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   120
         Width           =   780
      End
      Begin VB.Label lbBaseICMS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base ICMS"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   1410
         TabIndex        =   46
         Top             =   750
         Width           =   795
      End
      Begin VB.Label lbAliquotaIPI 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alíquota IPI"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   6570
         TabIndex        =   45
         Top             =   750
         Width           =   840
      End
      Begin VB.Label lbValorIPI 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor IPI"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   7860
         TabIndex        =   44
         Top             =   750
         Width           =   600
      End
      Begin VB.Label lblPrecoPesq 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Preço do Item"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   6570
         TabIndex        =   39
         Top             =   120
         Width           =   990
      End
      Begin VB.Label lblQtd 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Qtde"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   5865
         TabIndex        =   38
         Top             =   120
         Width           =   345
      End
   End
   Begin VB.TextBox txtPesquisar 
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
      TabIndex        =   0
      Top             =   930
      Width           =   5835
   End
   Begin VB.Frame FraPesquisa 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Pesquisa"
      ForeColor       =   &H00FF0000&
      Height          =   465
      Left            =   2640
      TabIndex        =   18
      Top             =   120
      Width           =   12345
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   270
         Left            =   1095
         TabIndex        =   40
         Top             =   120
         Width           =   11370
         Begin VB.CheckBox ckbIPI 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "Não Informar IPI"
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Left            =   9720
            TabIndex        =   72
            Top             =   15
            Value           =   1  'Checked
            Width           =   1800
         End
         Begin VB.CheckBox ckbInfoImpostos 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "Não Informar Impostos"
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Left            =   7680
            TabIndex        =   70
            Top             =   15
            Width           =   1935
         End
         Begin VB.CheckBox ckbEditarReferencia 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "Não Editar Referência e Descrição"
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Left            =   4800
            TabIndex        =   53
            Top             =   15
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.OptionButton optPrecoInformado 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "Preço Informado"
            ForeColor       =   &H00E0E0E0&
            Height          =   225
            Left            =   3135
            TabIndex        =   43
            Top             =   0
            Width           =   1605
         End
         Begin VB.OptionButton optPrecoCusto 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "Preço de Custo"
            ForeColor       =   &H00E0E0E0&
            Height          =   240
            Left            =   1710
            TabIndex        =   42
            Top             =   0
            Width           =   1695
         End
         Begin VB.OptionButton optPrecoVenda 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "Preço de Venda"
            ForeColor       =   &H00E0E0E0&
            Height          =   240
            Left            =   75
            TabIndex        =   41
            Top             =   0
            Value           =   -1  'True
            Width           =   1725
         End
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "Pesquisa"
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
         TabIndex        =   37
         Top             =   120
         Width           =   855
      End
   End
   Begin CentroDeDistribuicao.chameleonButton cmdEncerraNF 
      Height          =   510
      Left            =   10740
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
      MICON           =   "frmNotaFiscalOutrasOperacoes.frx":008C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdExcluiNF 
      Height          =   510
      Left            =   12180
      TabIndex        =   35
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Exclui NF"
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
      MICON           =   "frmNotaFiscalOutrasOperacoes.frx":00A8
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
      TabIndex        =   36
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
      MICON           =   "frmNotaFiscalOutrasOperacoes.frx":00C4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdProduto 
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   1380
      Width           =   14880
      _cx             =   26247
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
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmNotaFiscalOutrasOperacoes.frx":00E0
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
   Begin VSFlex7DAOCtl.VSFlexGrid grdItens 
      Height          =   1605
      Left            =   120
      TabIndex        =   19
      Top             =   4080
      Width           =   14880
      _cx             =   26247
      _cy             =   2831
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
      Cols            =   18
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmNotaFiscalOutrasOperacoes.frx":0232
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
      Editable        =   1
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
   Begin CentroDeDistribuicao.chameleonButton cmdCancelarFRAME 
      Height          =   510
      Left            =   9300
      TabIndex        =   85
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Cancelar NF"
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
      MICON           =   "frmNotaFiscalOutrasOperacoes.frx":0468
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdImportar 
      Height          =   510
      Left            =   7845
      TabIndex        =   88
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Importar Nota de Compra"
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
      MICON           =   "frmNotaFiscalOutrasOperacoes.frx":0484
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblPesquisa 
      AutoSize        =   -1  'True
      BackColor       =   &H00505050&
      Caption         =   "Referência/Fornecedor/Descrição/Fornecedor Descrição"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   120
      TabIndex        =   32
      Top             =   720
      Width           =   4110
   End
End
Attribute VB_Name = "frmNotaFiscalOutrasOperacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim adoProduto As New ADODB.Recordset
Dim adoTipoICMS As New ADODB.Recordset
Dim wLoja As String
Dim wTotal As Double
Dim i As Integer
Dim Data As String

' Variáveis do item
Dim wQuantidade As Integer
Dim wPreco As Double
Dim wReferencia As String
Dim wDescricao As String
Dim wValICMS As Double
Dim wAliqICMS As Double
Dim wBaseICMs As Double
Dim wValIPI As Double
Dim wAliqIPI As Double
Dim wBaseIPI As Double
Dim wNCM As String
Dim wValICMSST As Double
Dim wAliqICMSST As Double
Dim wBaseICMSST As Double
Dim wDesconto As Double
Dim wCST As String
Dim wReferenciaAntiga As String

Private Sub cdmCancel_Click()
    fraImportar.Visible = False
End Sub

Private Sub ckbInfoImpostos_Click()
    If ckbInfoImpostos.Value = 1 Then
        wCST = "60"
        cmbCST.Enabled = False
        cmbCST.Text = "60 - ICMS cobrado anteriormente por substituição tributária"
    Else
        wCST = "00"
        cmbCST.Text = "00 - Tributada integralmente "
        cmbCST.Enabled = True
    End If
    Impostos
End Sub

Private Sub ckbIPI_Click()
    
    If ckbIPI.Value = 0 Then
        txtBaseIPI.Enabled = True
        txtAliquotaIPI.Enabled = True
        txtValorIPI.Enabled = True
    Else
        txtBaseIPI.Enabled = False
        txtAliquotaIPI.Enabled = False
        txtValorIPI.Enabled = False
    End If

End Sub

Private Sub cmbCST_Click()

    wCST = Mid(cmbCST.Text, 1, 2)
    Impostos
    
End Sub

Private Sub cmbCST_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        wCST = Mid(cmbCST.Text, 1, 2)
        Impostos
      
    End If

End Sub


Private Sub cmdCancelarFRAME_Click()
    exibirFrameCancelamento
End Sub

Private Sub cmdEncerraNF_Click()

    If grdItens.Rows <= 1 Then
        
        MsgBox "Insira ao menos um item na nota.", vbCritical + vbOKOnly, "Atenção"
        Exit Sub
    
    Else
    
        criaCapa Data, i
        frmEncerraNFOutrasOperacoes.Show 1
    
    End If
End Sub
Private Sub limpaForm()
        
    txtPesquisar.Text = ""
        
    'Limpa Grid Pesquisa
    grdProduto.Rows = 1
    grdProduto.AddItem ""
    grdProduto.RemoveItem (1)
        
    'Zera os Campos, pega novo número do pedido
'    Form_Load
        
        'Limpa Grid Itens
        grdItens.Rows = 1
        grdItens.AddItem ""
        grdItens.RemoveItem (1)
   
        'Limpa Campos
        txtBaseCalcICMS.Text = "0,00"
        txtValICMS.Text = "0,00"
        txtBaseCalcICMSST.Text = "0,00"
        txtValICMSSub.Text = "0,00"
        txtValTotalMerc.Text = "0,00"
        txtValFrete.Text = "0,00"
        txtValSeguro.Text = "0,00"
        txtDesconto.Text = "0,00"
        txtOutrasDesp.Text = "0,00"
        txtTotalIPI.Text = "0,00"
        txtTotalNota.Text = "0,00"
        txtPedido.Text = ""
        txtReferencia.Text = ""
        lblDescFornecedor.Caption = ""
        
        'Habilita Pesquisa
        fraPesquisa.Enabled = True
        
        cmbCST.ListIndex = 0
        optPrecoVenda.Value = True
        ckbEditarReferencia.Value = 1
        ckbInfoImpostos.Value = 0
        ckbIPI.Value = 1
        
        LimpaInformacao
        
        'txtPedido.SetFocus
        
    
End Sub

Private Sub cmdExcluiNF_Click()
   
    If MsgBox("Confirma a Exclusão da Nota?", vbYesNo + vbQuestion + vbDefaultButton2, "Exclusão de Produto") = vbYes Then
        ExcluiNF
    End If
   
    txtPedido.Text = wPedido
    txtPesquisar.SetFocus
    
End Sub

Private Sub autoPreencherZero(campo As TextBox)
    If campo.Text = "" Then
        campo.Text = "0,00"
    Else
        formataCampoDinheiro campo
    End If
End Sub

Private Sub carregaTipoICMS()
    sql = "select tpi_codigo,tpi_descricao from tipoIcms order by tpi_codigo"
    adoTipoICMS.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic

    Do While Not adoTipoICMS.EOF
        cmbCST.AddItem Trim(adoTipoICMS("tpi_codigo")) & " - " & (adoTipoICMS("tpi_descricao"))
        adoTipoICMS.MoveNext
    Loop
    adoTipoICMS.Close
    cmbCST.ListIndex = 0
End Sub

Private Sub cmdImportar_Click()
    exibirFrameImportar
End Sub

Private Sub cmdImpostoSISTEMA_Click()
    
    Dim adoConsultaICMS As New ADODB.Recordset
    Dim adoConsultaProduto As New ADODB.Recordset
    Dim wChaveICMSitem As String
    Dim wChaveICMS As String
    
    If txtReferencia.Text <> "" Then
    
        sql = "select top 1 pr_cst as cst, PR_ICMSSaida as AliqICMS, pr_precoCusto1 as precoCusto, pr_precovenda1 as precovenda, " & vbNewLine _
        & "pr_icmssaida,pr_codigoreducaoicms,PR_substituicaotributaria,*" _
        & "from produto where pr_referencia = '" & txtReferencia.Text & "'"
    
        adoConsultaProduto.CursorLocation = adUseClient
        adoConsultaProduto.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
        sql = "select tpi_codigo as codigo from tipoIcms order by tpi_codigo"
        adoConsultaICMS.CursorLocation = adUseClient
        adoConsultaICMS.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
        Do While Not adoConsultaICMS.EOF
            
                If Val(adoConsultaICMS("codigo")) = Val(adoConsultaProduto("cst")) Then
                    cmbCST.ListIndex = adoConsultaICMS.AbsolutePosition - 1
                    cmbCST_Click
                End If
            
                adoConsultaICMS.MoveNext
        Loop
    
''''''''''    wChaveICMS = "91"
''''''''''    wChaveICMSitem = wChaveICMS
''''''''''    'wChaveICMSitem = wChaveICMSitem & Format(adoConsultaProduto("pr_icmssaida"), "####00") & adoConsultaProduto("pr_codigoreducaoicms") & wSubstituicaoTributaria
''''''''''
''''''''''                        If adoConsultaProduto("PR_substituicaotributaria") = "N" _
''''''''''                           And adoConsultaProduto("PR_codigoreducaoicms") > 0 Then
''''''''''                            wST20 = "S"
''''''''''                        End If
''''''''''
''''''''''                        If adoConsultaProduto("PR_substituicaotributaria") = "S" Then
''''''''''                            wSubstituicaoTributaria = 1
''''''''''                            wST60 = "S"
''''''''''                            wChaveICMSitem = wChaveICMSitem & "000" & wSubstituicaoTributaria
''''''''''                        Else
''''''''''                            wSubstituicaoTributaria = 0
''''''''''                            wChaveICMSitem = wChaveICMSitem & Format(adoConsultaProduto("pr_icmssaida"), "####00") & adoConsultaProduto("pr_codigoreducaoicms") & wSubstituicaoTributaria
''''''''''                        End If
''''''''''
''''''''''
''''''''''                        If AcharICMSInterEstadual(adoConsultaProduto("PR_Referencia"), wChaveICMSitem) = False Then
''''''''''
''''''''''                              If AcharICMSInterEstadual(adoConsultaProduto("PR_Referencia"), Mid(Trim(wChaveICMSitem), 1, 2) & "1200") = False Then
''''''''''                                    'EncerraVenda = False
''''''''''                                    adoConsultaProduto.Close
''''''''''                                    'Exit Function
''''''''''                              End If
''''''''''                        End If
''''''''''
''''''''''                        wCFOItem = wIE_Cfo
''''''''''                        GLB_AliquotaAplicadaICMS = wIE_icmsAplicado
''''''''''                        GLB_Tributacao = wIE_Tributacao
''''''''''                        GLB_CFOP = wIE_Cfo
''''''''''                        wAnexoIten = adoConsultaProduto("PR_CodigoReducaoICMS")
''''''''''
''''''''''                            GLB_ValorCalculadoICMS = Format((((txtPreco.Text - txtDesc.Text) * GLB_AliquotaAplicadaICMS) / 100), "0.00")
''''''''''                            GLB_TotalICMSCalculado = (GLB_TotalICMSCalculado + GLB_ValorCalculadoICMS)
''''''''''                            If GLB_TotalICMSCalculado > 0 Then
''''''''''                                If wIE_BasedeReducao = 0 Then
''''''''''                                    If GLB_AliquotaAplicadaICMS = 0 Then
''''''''''                                        GLB_BasedeCalculoICMS = 0
''''''''''                                    Else
''''''''''                                        GLB_BasedeCalculoICMS = (txtPreco.Text - txtDesc.Text)
''''''''''                                    End If
''''''''''                                Else
''''''''''                                    GLB_BasedeCalculoICMS = Format((txtPreco.Text - txtDesc.Text) - _
''''''''''                                    (((txtPreco.Text - txtDesc.Text) * wIE_BasedeReducao) / 100), "0.00")
''''''''''                                End If
''''''''''                                GLB_BaseTotalICMS = (GLB_BaseTotalICMS + GLB_BasedeCalculoICMS)
''''''''''                            End If
''''''''''
''''''''''                            WAnexoAux = ""
''''''''''                            'If RsItensNF("pr_codigoreducaoicms") <> 0 Then
''''''''''                               'WAnexoAux = WAnexoAux & "," & Format(RsItensNF("ITEM"), "0")
''''''''''                            'End If
''''''''''
''''''''''




        If txtBaseICMS.Enabled = True Then
            txtBaseICMS.Text = adoConsultaProduto("precovenda")
            If optPrecoCusto.Value = True Then txtBaseICMS.Text = adoConsultaProduto("precoCusto")
        End If
        If txtAliquotaICMS.Enabled = True Then txtAliquotaICMS.Text = adoConsultaProduto("AliqICMS")
        txtBaseICMS_LostFocus
        txtaliquotaicms_LostFocus
    
        adoConsultaICMS.Close
        adoConsultaProduto.Close
    End If
End Sub

Private Sub cmdImpotar_Click()
    importarNota Trim(txtNf.Text), Trim(txtSerienf.Text), Trim(txtCNPJ.Text)
End Sub

Private Sub Form_Activate()
    If finalizaOutrasOperacoes = True Then
        limpaForm
        wPedido = PegaNumPedido
        txtPedido.Text = wPedido
        'txtReferencia.SetFocus
        txtPesquisar.SetFocus
    End If
End Sub

Private Sub Form_Load()

    Dim sql As String
    Dim Descricao As String
    
    carregarPosicaoTamanhoTela Me
    carregarPosicaoFrame frameCancelamento
    carregarPosicaoFrame fraImportar
    carregaTipoICMS
    
    limpaForm
    wPedido = PegaNumPedido
    txtPedido.Text = wPedido
    ckbInfoImpostos_Click
    
End Sub

Private Sub importarNota(nota As String, serie As String, cnpj As String)

    Dim sql As String
    Dim rsNota As New ADODB.Recordset
    Dim codFornecedor As String
    Dim Data As Date
    
    Data = Format(Date, "yyyy/mm/dd")
    
    sql = "select * from capanfcompra where cc_notaFiscal = " & nota & " and cc_serie = '" & serie & "' and cc_fornecedor in (select fo_codigoFornecedor from fornecedor where fo_cgc = '" & cnpj & "')"
    rsNota.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    
    If Not rsNota.EOF Then
        codFornecedor = rsNota("cc_fornecedor")
        rsNota.Close
    Else
        MsgBox "Nota não Encontrada", vbCritical, "Nota não Encontrada"
        rsNota.Close
        Exit Sub
    End If
    
    
    
    sql = "Exec sp_cria_nfcapa_compra " & nota & ", '" & serie & "'," & codFornecedor & ", 'CD', 'SA', '" & Format(Data, "YYYY/MM/DD") & "'"
    ADO_Cn_CD.Execute sql
    
    
     
    ' select * from nfcapa
    'where NUMEROPED in (select max(NUMEROPED) from nfcapa where NfDevolucao = '3210'
    sql = "select TOP 1 numeroped, baseicms, vlricms, baseicmsst, vlrMercadoria, valfrete, desconto, valorOutros, totalIPI, totalNota, valorICMSST from nfcapa where numeroped in ( select max(numeroPed) from nfcapa where nfDevolucao = " & nota & " and SerieDevolucao = '" & serie & "')"
    rsNota.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    
    txtPedido.Text = rsNota("numeroped")
    txtPedido_KeyPress 13
    
'    If Not rsNota.EOF Then
'
'        txtPedido.Text = rsNota("numeroped")
'        wPedido = rsNota("numeroped")
'        txtBaseCalcICMS.Text = Format(rsNota("baseICMS"), "#0.00")
'        txtValICMS.Text = Format(rsNota("vlrICMS"), "#0.00")
'        txtBaseCalcICMSST.Text = Format(rsNota("baseICMSST"), "#0.00")
'        txtValTotalMerc.Text = Format(rsNota("vlrMercadoria"), "#0.00")
'        txtValFrete.Text = Format(rsNota("valFrete"), "#0.00")
'        txtDesconto.Text = Format(rsNota("Desconto"), "#0.00")
'        txtOutrasDesp.Text = Format(rsNota("valorOutros"), "#0.00")
'        txtTotalIPI.Text = Format(rsNota("totalIPI"), "#0.00")
'        txtTotalNota.Text = Format(rsNota("totalNota"), "#0.00")
'        txtValICMSSub.Text = Format(rsNota("valorICMSST"), "#0.00")
'
'    End If
    rsNota.Close
'
'    sql = "select referencia, pr_descricao, qtde, vlUnit, desconto, baseICMS, ICMS, valorICMS, baseIPI, AliqIPI," _
'     & "vlIPI, baseSub, sub, valorSub, CSTICMS, pr_classeFiscal from nfitens,produto where numeroped = " & txtPedido.Text & " and pr_referencia = referencia"
'
'    rsNota.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
'
'    If Not rsNota.EOF Then
'        Do While Not rsNota.EOF
'
'            carregaGridTtens Trim(rsNota("referencia")), Trim(rsNota("pr_descricao")), rsNota("qtde"), _
'                            Format(rsNota("vlUnit"), "#0.00"), Format(rsNota("desconto"), "#0.00"), Format(rsNota("valorICMS"), "#0.00"), _
'                            Format(rsNota("baseICMS"), "#0.00"), Format(rsNota("ICMS"), "#0.00"), Format(rsNota("vlIPI"), "#0.00"), _
'                            Format(rsNota("baseIPI"), "#0.00"), Format(rsNota("aliqIPI"), "#0.00"), Format(rsNota("valorSub"), "#0.00"), _
'                            Format(rsNota("baseSub"), "#0.00"), Format(rsNota("sub"), "#0.00"), Trim(rsNota("CSTICMS")), _
'                            Trim(rsNota("pr_classeFiscal")), ""
'        rsNota.MoveNext
'        Loop
'    End If
    
    'rsNota.Close
    
    fraImportar.Visible = False
    
End Sub

Private Sub grdItens_KeyPressEdit(ByVal row As Long, ByVal col As Long, KeyAscii As Integer)
    Dim referencia As String
    Dim qtde As Integer
    
    
    If KeyAscii = 13 Then
        If Not alteraQuantidade(grdItens.TextMatrix(grdItens.row, 2), grdItens.TextMatrix(grdItens.row, 0)) Then
            MsgBox "Quantidade de itens maior do que a enviada na nota."
            grdItens.TextMatrix(grdItens.row, 2) = grdItens.TextMatrix(grdItens.row, 17)
        End If
    End If
End Sub

Private Sub grdItens_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 46 Then
        
        If MsgBox("Confirma a Exclusão deste item da Nota?", vbYesNo + vbQuestion + vbDefaultButton2, "Exclusão de Produto") = vbYes Then
            
            DeletaGrid
            
        End If
        
    End If
    
End Sub

Private Sub grdProduto_EnterCell()
    InformacoesProduto
End Sub

Private Sub grdProduto_KeyPress(KeyAscii As Integer)

    InformacoesProduto
    
End Sub

Private Sub InformacoesProduto()
    
    Dim sql As String
     
    If grdProduto.row <> 0 Then
    
        wReferencia = Trim(grdProduto.TextMatrix(grdProduto.row, 0))
        wReferenciaAntiga = grdProduto.TextMatrix(grdProduto.row, 0)
        wDescricao = Trim(grdProduto.TextMatrix(grdProduto.row, 1))
        sql = "select * from Produto where pr_referencia = '" & wReferencia & "'"
        adoProduto.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
        If optPrecoCusto.Value = True Then
            wPreco = Format(grdProduto.TextMatrix(grdProduto.row, 4), "#0.00")
        Else
            wPreco = Format(grdProduto.TextMatrix(grdProduto.row, 3), "#0.00")
        End If
    
        wNCM = Trim(adoProduto("pr_ClasseFiscal"))
        wBaseICMs = Format(adoProduto("pr_precoVenda1"), "#0.00")
        wBaseIPI = 0
        wAliqIPI = 0
        wAliqICMS = 0
        wValIPI = 0
        wValICMS = 0
        wQuantidade = 1
    
        'Preenche os TextBox
    
        txtReferencia.Text = wReferencia
        txtDescricao.Text = Format(wDescricao, "#0.00")
        txtPreco.Text = wPreco
        formataCampoDinheiro txtPreco
        txtQtde.Text = "1"
        txtNCM.Text = wNCM
        txtBaseICMS.Text = wBaseICMs
        formataCampoDinheiro txtBaseICMS
        txtAliquotaICMS.Text = "0,00"
        txtAliquotaIPI.Text = "0,00"
        txtValorICMS.Text = "0,00"
        txtValorIPI.Text = "0,00"
        txtBaseIPI.Text = "0,00"
        adoProduto.Close
        txtQtde.SetFocus
        
        cmdImpostoSISTEMA_Click
    
    End If
    
    
End Sub
Private Sub LimpaInformacao()
    
    txtReferencia.Text = ""
    txtDescricao.Text = ""
    txtPreco.Text = "0,00"
    txtQtde.Text = "1"
    txtNCM.Text = ""
    txtBaseICMS.Text = "0,00"
    txtAliquotaICMS.Text = "0,00"
    txtAliquotaIPI.Text = "0,00"
    txtValorICMS.Text = "0,00"
    txtValorIPI.Text = "0,00"
    txtBaseIPI.Text = "0,00"
    txtDesc.Text = "0,00"
    txtBaseAliqICMSST.Text = "0,00"
    txtValICMSST.Text = "0,00"
    txtAliquotaICMSST.Text = "0,00"
    
    'Limpa Pesquisa
    txtPesquisar.Text = ""
        
    'Limpa Grid Pesquisa
    grdProduto.Rows = 1
    grdProduto.AddItem ""
    grdProduto.RemoveItem (1)
    
    'txtPesquisar.SetFocus
    
End Sub
Private Sub InsereGrid()
    
    If validaImposto Then
        Exit Sub
    End If
    
    If checarCampoVazio Then
        Exit Sub
    End If

    Call carregaGridTtens(wReferencia, Trim(wDescricao), txtQtde.Text, txtPreco.Text, txtDesconto.Text, _
                      txtValorICMS.Text, txtBaseICMS.Text, txtAliquotaICMS.Text, txtValorIPI.Text, _
                      txtBaseIPI.Text, txtAliquotaIPI.Text, txtValICMSST.Text, txtBaseAliqICMSST.Text, txtAliquotaICMSST.Text, _
                      wCST, wNCM, wReferenciaAntiga)
    
    CriaNota
    LimpaInformacao
    TotaisNota
    grdProduto.SetFocus
    
End Sub

Private Sub carregaGridTtens(referencia As String, Descricao As String, quantidade As String, _
                             Preco As String, Desconto As String, _
                             valorICMS As String, baseICMS As String, AliquotaICMS As String, _
                             valorIPI As String, baseIPI As String, _
                             aliquotaIPI As String, valorICMSST As String, baseICMSST As String, _
                             aliquotaICMSST As String, CST As String, NCM As String, _
                             ReferenciaAntiga As String)

    grdItens.AddItem referencia & Chr(9) & _
    Descricao & Chr(9) & _
    quantidade & Chr(9) & _
    formataVariavelDinheiro(Preco) & Chr(9) & _
    formataVariavelDinheiro(Desconto) & Chr(9) & _
    formataVariavelDinheiro(baseICMS) & Chr(9) & _
    formataVariavelDinheiro(AliquotaICMS) & Chr(9) & _
    formataVariavelDinheiro(valorICMS) & Chr(9) & _
    formataVariavelDinheiro(baseIPI) & Chr(9) & _
    formataVariavelDinheiro(aliquotaIPI) & Chr(9) & _
    formataVariavelDinheiro(valorIPI) & Chr(9) & _
    formataVariavelDinheiro(baseICMSST) & Chr(9) & _
    formataVariavelDinheiro(aliquotaICMSST) & Chr(9) & _
    formataVariavelDinheiro(valorICMSST) & Chr(9) & _
    CST & Chr(9) & _
    NCM & Chr(9) & _
    ReferenciaAntiga & Chr(9) & _
    quantidade
    
    
    wPreco = Format(Preco, "#0.00")
    wDesconto = Format(Desconto, "#0.00")
    
    wBaseICMs = Format(baseICMS, "#0.00")
    wAliqICMS = Format(AliquotaICMS, "#0.00")
    wValICMS = Format(valorICMS, "#0.00")
    
'    wBaseIPI = Format(baseIPI, "#0.00")
    wAliqIPI = Format(aliquotaIPI, "#0.00")
    wValIPI = Format(valorIPI, "#0.00")
    
'    wBaseICMSST = Format(baseICMSST, "#0.00")
'    wAliqICMSST = Format(aliquotaICMSST, "#0.00")
'    wValICMSST = Format(valorICMSST, "#0.00")
    
    wCST = CST
    wQuantidade = CInt(quantidade)
    

End Sub

Private Sub DeletaGrid()
    DeletaNota grdItens.TextMatrix(grdItens.row, 0)
    grdItens.RemoveItem grdItens.row
    TotaisNota

End Sub

Private Sub DeletaNota(referencia As String)
    
Dim sql As String
Dim adoDeletaNota As New ADODB.Recordset

    sql = "delete from nfitens where NUMEROPED = '" & wPedido & "' and REFERENCIA = '" & referencia & "'"
    ADO_Cn_CDLocal.Execute (sql)
    
End Sub

Private Sub TotaisNota()

    txtValTotalMerc.Text = Format(Totaliza(3, 2), "#0.00")
    txtBaseCalcICMS.Text = Format(Totaliza(5), "#0.00")
    txtDesconto.Text = Format(Totaliza(4), "#0.00")
    txtValICMS.Text = Format(Totaliza(7), "#0.00")
    txtTotalIPI.Text = Format(Totaliza(10), "#0.00")
    txtValICMSSub.Text = Format(Totaliza(13), "#0.00")
    txtBaseCalcICMSST.Text = Format(Totaliza(11), "#0.00")
    
    txtTotalNota.Text = Format(CDbl(txtValICMSSub.Text) + CDbl(txtValTotalMerc.Text) + CDbl(txtTotalIPI.Text) + CDbl(txtValFrete.Text) - CDbl(txtDesconto.Text), "#0.00")
    
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNumerico(KeyAscii)
End Sub

Private Sub txtAliquotaICMS_KeyPress(KeyAscii As Integer)
    
    KeyAscii = campoNumericoVirgula(KeyAscii)
    
    If KeyAscii = 13 Then
        txtaliquotaicms_LostFocus
        txtQtde_keypress KeyAscii
    End If
    
End Sub

Private Sub txtaliquotaicms_LostFocus()

    autoPreencherZero txtAliquotaICMS
    
    txtValorICMS.Text = (CDbl(txtAliquotaICMS.Text) * CDbl(txtBaseICMS.Text)) / 100
    autoPreencherZero txtValorICMS
    
End Sub

Private Sub txtAliquotaICMSST_KeyPress(KeyAscii As Integer)
        
    KeyAscii = campoNumericoVirgula(KeyAscii)
        
    If KeyAscii = 13 Then
        txtAliquotaICMSST_LostFocus
        txtQtde_keypress KeyAscii
    End If
    
End Sub

Private Sub txtAliquotaICMSST_LostFocus()
    
    autoPreencherZero txtAliquotaICMSST
    
    txtValICMSST.Text = (CDbl(txtBaseAliqICMSST.Text) * CDbl(txtBaseCalcICMSST)) / 100
    autoPreencherZero txtValICMSST

End Sub

Private Sub txtAliquotaIPI_KeyPress(KeyAscii As Integer)
    
    KeyAscii = campoNumericoVirgula(KeyAscii)
    
    If KeyAscii = 13 Then
        txtAliquotaIPI_LostFocus
        txtQtde_keypress KeyAscii
    End If
    
End Sub


Private Sub txtAliquotaIPI_LostFocus()

    autoPreencherZero txtAliquotaIPI
    
    txtValorIPI.Text = (CDbl(txtBaseIPI.Text) * CDbl(txtAliquotaIPI.Text)) / 100
    autoPreencherZero txtValorIPI
    
End Sub

Private Sub txtBaseAliqICMSST_KeyPress(KeyAscii As Integer)
    
    KeyAscii = campoNumericoVirgula(KeyAscii)
    
    If KeyAscii = 13 Then
        txtBaseAliqICMSST_LostFocus
        txtQtde_keypress KeyAscii
    End If
    
End Sub

Private Sub txtBaseAliqICMSST_LostFocus()
    autoPreencherZero txtBaseAliqICMSST
End Sub

Private Sub txtBaseICMS_KeyPress(KeyAscii As Integer)
    
    KeyAscii = campoNumericoVirgula(KeyAscii)
        
    If KeyAscii = 13 Then
        txtBaseICMS_LostFocus
        txtQtde_keypress KeyAscii
    End If
    
End Sub

Private Sub txtBaseICMS_LostFocus()
    autoPreencherZero txtBaseICMS
End Sub


Private Sub txtBaseIPI_KeyPress(KeyAscii As Integer)
    
    KeyAscii = campoNumericoVirgula(KeyAscii)
    
    If KeyAscii = 13 Then
        txtBaseIPI_LostFocus
        txtQtde_keypress KeyAscii
    End If
    
End Sub

Private Sub txtBaseIPI_LostFocus()
    autoPreencherZero txtBaseIPI
End Sub

Private Sub txtCnpj_KeyPress(KeyAscii As Integer)
    Dim sql As String
    Dim rsNota As New ADODB.Recordset
    
    If KeyAscii = 13 Then
        
        sql = "select fo_razaoSocial from fornecedor where fo_cgc like '%" & txtCNPJ.Text & "%'"
        rsNota.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockReadOnly
        
        If Not rsNota.EOF Then
            lblDescFornecedor.Caption = Trim(rsNota("fo_razaoSocial"))
            lblDescFornecedor.Visible = True
            cmdImpotar.Enabled = True
            cmdImpotar.SetFocus
            
        Else
            MsgBox "Fornecedor não Encontrado!", vbCritical, "Fornecedor não Encontrado"
        End If
        
        rsNota.Close
        
    End If
    
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNumericoVirgula(KeyAscii)
End Sub

Private Sub txtDesc_LostFocus()
    autoPreencherZero txtDesc
End Sub

'Atualiza Campos

Private Sub txtDescricao_LostFocus()

    wDescricao = Trim(txtDescricao.Text)
    
End Sub

Private Sub txtNCM_Change()
    
    If txtNCM.Text <> "" Then
    
        wNCM = txtNCM.Text
        
    End If
    
End Sub

Private Sub txtNCM_KeyPress(KeyAscii As Integer)

    KeyAscii = campoNumerico(KeyAscii)

End Sub

Private Sub txtNCM_lostfocus()

    txtQtde_keypress 13
    
End Sub

Private Sub txtPedido_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNumericoVirgula(KeyAscii)
    
    If KeyAscii = 13 Then
        If Val(txtPedido.Text) > 0 Then
            If MsgBox("Atenção! Você irá perder as informações novas dessa nota. Deseja continuar?", vbYesNo + vbQuestion + vbDefaultButton2, "Carregar nova nota") = vbYes Then
            
                Dim pedido As String
                
                pedido = txtPedido.Text
                limpaForm
                txtPedido.Text = pedido
                carregaNotaCriada pedido
                wPedido = pedido
                NroPedido = pedido
                
            End If
        Else
            MsgBox "Número de pedido inválido!", vbExclamation, "Abrir Pedido"
        End If
    End If
    
End Sub

Private Sub carregaNotaCriada(ByRef numeroped As String)

    Dim adoConsultaCAPA As New ADODB.Recordset
    Dim adoConsultaITENS As New ADODB.Recordset

    sql = "select top 1 tm from nfcapa where numeroped = " & numeroped
    adoConsultaCAPA.CursorLocation = adUseClient
    adoConsultaCAPA.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
        If adoConsultaCAPA.EOF Then
            MsgBox "Pedido " & txtPedido.Text & " não encontrado!", vbExclamation, "Abrir Pedido"
            txtPedido.Text = ""
        Else
            
            Select Case adoConsultaCAPA("tm")
            Case 100
                MsgBox "Nota Fiscal já autorizada! Não é póssivel modificar as informações desse pedido.", vbInformation, "Abrir Pedido"
            Case 200
                MsgBox "Não foi emitido a NF do pedido " & txtPedido.Text & ". " & vbNewLine & _
                       "Execute o DMAC NFe e tente novamente", vbInformation, "Abrir Pedido"
            Case Else
            
                sql = "select referencia, pr_descricao, vlunit, desconto, " & _
                      "BaseICMS, ICMSAplicado, VALORICMS,baseIPI, VLIPI, QTDE, valorSub, sub, ALIQIPI," & _
                      "baseSub, pr_classeFiscal,CSTICMS from nfitens, Produto where pr_referencia = referencia and numeroped = " & numeroped & " order by item"
                adoConsultaITENS.CursorLocation = adUseClient
                adoConsultaITENS.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
                
                grdItens.Rows = 1
                
                Do While Not adoConsultaITENS.EOF
                    carregaGridTtens adoConsultaITENS("referencia"), adoConsultaITENS("pr_descricao"), adoConsultaITENS("QTDE"), adoConsultaITENS("vlunit"), _
                                     adoConsultaITENS("desconto"), adoConsultaITENS("VALORICMS"), adoConsultaITENS("BaseICMS"), adoConsultaITENS("ICMSAplicado"), _
                                     adoConsultaITENS("VLIPI"), adoConsultaITENS("baseIPI"), adoConsultaITENS("ALIQIPI"), adoConsultaITENS("valorSub"), _
                                     adoConsultaITENS("baseSub"), adoConsultaITENS("SUB"), adoConsultaITENS("CSTICMS"), adoConsultaITENS("pr_classeFiscal"), _
                                     adoConsultaITENS("referencia")
                    TotaisNota
                    adoConsultaITENS.MoveNext
                Loop
                    
                adoConsultaITENS.Close
            
            End Select
            
                             
        End If
        
    adoConsultaCAPA.Close
    
    
End Sub

Private Sub txtPreco_KeyPress(KeyAscii As Integer)

    KeyAscii = campoNumericoVirgula(KeyAscii)

    If KeyAscii = 13 Then
        txtPreco_lostFocus
        txtQtde_keypress 13
    End If
    
End Sub

Private Sub txtPreco_lostFocus()
    
    autoPreencherZero txtPreco
    
    txtBaseICMS.Text = Format(wBaseICMs, "#0.00")
    
    If ckbIPI.Value = 0 Then
        txtBaseIPI.Text = Format(wBaseICMs, "#0.00")
    End If
    
End Sub

Private Sub txtQtde_keypress(KeyAscii As Integer)
    Dim i As Long
    
    KeyAscii = campoNumerico(KeyAscii)
    
    If KeyAscii = 13 And IsNumeric(txtQtde.Text) Then
        i = 1
        Do While grdItens.Rows <> i
        
            If Trim(grdItens.TextMatrix(i, 0)) = wReferencia Then
                MsgBox "Este item já foi inserido.", vbCritical + vbOKOnly, "Atenção"
                Exit Sub
            End If
        i = i + 1
        Loop
        
        InsereGrid
    
    End If
    
End Sub

Private Sub txtQtde_Keyup(KeyCode As Integer, Shift As Integer)
    
    If (KeyCode >= 48 And KeyCode <= 57) Or (KeyCode >= 96 And KeyCode <= 105) Then
    
        If txtBaseICMS.Enabled = True Then
            txtBaseICMS.Text = Format(Format(txtPreco.Text, "#0.00") * txtQtde.Text, "#0.00")
        End If
        
        If ckbIPI.Value = 1 And txtBaseIPI.Enabled = True Then
            txtBaseIPI.Text = Format(txtBaseICMS.Text, "#0.00")
        End If
        
    End If
    
   
End Sub



Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
       KeyAscii = campoNumerico(KeyAscii)
End Sub

Private Sub txtReferencia_lostfocus()

    wReferencia = Trim(txtReferencia.Text)
    
End Sub

Public Sub cmdRetorna_Click()

    Unload Me
    
End Sub

Private Sub Impostos()
    
    Dim sql As String
    Dim adoImposto As New ADODB.Recordset
    
        txtValorICMS.Enabled = False
        txtValorICMS.Text = "0,00"
        
        txtAliquotaICMS.Enabled = False
        txtAliquotaICMS.Text = "0,00"
        
        txtBaseICMS.Enabled = False
        txtBaseICMS.Text = "0,00"
        
        txtNCM.Enabled = False
        txtNCM.Text = "0,00"
        
        txtBaseAliqICMSST.Enabled = False
        txtBaseAliqICMSST.Text = "0,00"
        
        txtValICMSST.Enabled = False
        txtValICMSST.Text = "0,00"
        
        txtAliquotaICMSST.Enabled = False
        txtAliquotaICMSST.Text = "0,00"
       
        
    If ckbInfoImpostos.Value = 0 Then
    
        
        sql = "select * from nfe_estrutura where etr_rotulo = 'ICMS" & wCST & "'"
        adoImposto.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
        Do While Not adoImposto.EOF
        
            Select Case Trim(adoImposto("ETR_Campo"))
        
                Case "VBC"
                
                    txtBaseICMS.Enabled = True
                    If txtPreco.Text <> Empty Then txtBaseICMS.Text = (txtPreco.Text) * Val(txtQtde.Text)
                
                Case "PICMS"
                
                    txtAliquotaICMS.Enabled = True
            
                Case "VICMS"
                
                    txtValorICMS.Enabled = True
            
                Case "VBCST"
            
                    txtBaseAliqICMSST.Enabled = True
            
                Case "PICMSST"
                
                    txtAliquotaICMSST.Enabled = True
                
                Case "VICMSST"
                
                    txtValICMSST.Enabled = True
            
            End Select
            adoImposto.MoveNext
            
        Loop
        adoImposto.Close
    End If
    
End Sub


'Pesquisa Grid Produtos

Private Sub txtPesquisar_GotFocus()

     txtPesquisar.SelStart = 0
     txtPesquisar.SelLength = Len(txtPesquisar)
     txtPesquisar.SetFocus
     txtQtde.Text = ""
     
End Sub

Private Sub txtPesquisar_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        fraPesquisa.Enabled = False

        grdProduto.Rows = 1
        Screen.MousePointer = 11
        If txtPesquisar.Text <> "" Then
            
            If IsNumeric(Trim(txtPesquisar.Text)) = True And Len(Trim(txtPesquisar.Text)) = 3 Then
                PesquisarProduto 1, txtPesquisar.Text '1 = Pesquisa por fonecedor
                
            ElseIf IsNumeric(Trim(txtPesquisar.Text)) = True And Len(Trim(txtPesquisar.Text)) = 7 Then
                PesquisarProduto 2, txtPesquisar.Text '2 = Pesquisa por referencia
                
            ElseIf IsNumeric(Trim(txtPesquisar.Text)) = True And Len(Trim(txtPesquisar.Text)) > 3 Then
                PesquisarProduto 5, txtPesquisar.Text '2 = Pesquisa por codigo de barras
               
            ElseIf IsNumeric(Trim(txtPesquisar.Text)) = False Then
                If IsNumeric(Mid(txtPesquisar.Text, 1, 3)) = True And Trim(Mid(txtPesquisar.Text, 4, 1)) = "" Then
                    PesquisarProduto 4, txtPesquisar.Text '3 = Pesquisa por Fornecedor e Descrcao
                Else
                    PesquisarProduto 3, txtPesquisar.Text '3 = Pesquisa por descricao
                End If
                grdProduto.SetFocus
                grdProduto.row = 1
            Else
                txtPesquisar.SelStart = 0
                txtPesquisar.SelLength = Len(txtPesquisar.Text)
                txtPesquisar.SetFocus
                Exit Sub
            End If
            
        Else
           txtPesquisar.SelStart = 0
           txtPesquisar.SelLength = Len(txtPesquisar.Text)
           txtPesquisar.SetFocus
           Exit Sub
        End If
        Screen.MousePointer = 0
        
    End If
    
    If ckbEditarReferencia.Value = False Then
        
        txtReferencia.Enabled = True
        txtDescricao.Enabled = True
        
    Else
    
        txtReferencia.Enabled = False
        txtDescricao.Enabled = False
        
    End If
    
    If optPrecoInformado.Value Then
    
        txtPreco.Enabled = True
        txtPreco.Locked = False
    Else
        txtPreco.Enabled = False
        txtPreco.Locked = True
    End If
End Sub

Function PesquisarProduto(ByVal tipoPesquisa As Integer, ByVal pesquisar As String)
    
wLoja = "CD"
        
'
'--------------------------------------Pesquisa Por fornecedor (1)--------------------------
'
    If tipoPesquisa = 1 Then
        sql = ""
        sql = "Select PR_Referencia,PR_Descricao,PR_Precovenda1,pr_customedio1,PR_CodigoFornecedor,es_Estoque," _
            & "PR_ICMSSaida,PR_ICMSEntrada,PR_IcmPdv,Pr_IcmPdvEntrada," _
            & "PR_SubstituicaoTributaria,es_Estoque from Produto,estoque " _
            & "where es_Referencia=PR_Referencia and PR_CodigoFornecedor = " & Trim(pesquisar) _
            & " and es_Loja ='" & Trim(wLoja) & "' and PR_Situacao <> 'E'" _
            & " order by PR_CodigoFornecedor,PR_Descricao"
            
'
'--------------------------------------Pesquisa Por Referencia (2)--------------------------
'
    ElseIf tipoPesquisa = 2 Then
        sql = ""
        sql = "Select PR_Referencia,PR_Descricao,PR_Precovenda1,pr_customedio1,es_Estoque," _
            & "PR_ICMSSaida,PR_ICMSEntrada,PR_IcmPdv,Pr_IcmPdvEntrada," _
            & "PR_SubstituicaoTributaria,es_Estoque from Produto,estoque " _
            & "where es_Referencia=PR_Referencia and PR_Referencia = '" & Trim(pesquisar) _
            & "' and es_Loja ='" & Trim(wLoja) & "' and PR_Situacao <> 'E'"
            
'
'--------------------------------------Pesquisa Por Descricao (3)---------------------------
'
    ElseIf tipoPesquisa = 3 Then
        sql = ""
        sql = "Select PR_Referencia,PR_Descricao,PR_Precovenda1,pr_customedio1,es_Estoque," _
            & "PR_ICMSSaida,PR_ICMSEntrada,PR_IcmPdv,Pr_IcmPdvEntrada," _
            & "PR_SubstituicaoTributaria,es_Estoque from Produto,estoque " _
            & "where es_Referencia=PR_Referencia and PR_Descricao Like '" _
            & Trim(UCase(Trim(pesquisar))) _
            & "%' and es_Loja ='" & Trim(wLoja) & "' and PR_Situacao <> 'E'" _
            & " order by PR_Descricao "
            

'
'-------------------------------Pesquisa Por Fornecedor , Decricao (4)-----------------------
    ElseIf tipoPesquisa = 4 Then
        sql = ""
        sql = "Select PR_Referencia,PR_Descricao,PR_Precovenda1,pr_customedio1,PR_CodigoFornecedor " _
            & ",PR_ICMSSaida,PR_ICMSEntrada,PR_IcmPdv,Pr_IcmPdvEntrada,PR_SubstituicaoTributaria,es_Estoque from Produto,estoque " _
            & "where es_Referencia=PR_Referencia and PR_Descricao Like '" _
            & Trim(UCase(Mid(Trim(pesquisar), 4, Len(Trim(Trim(pesquisar)))))) _
            & "%' And PR_CodigoFornecedor = " & Mid(pesquisar, 1, 3) _
            & " and es_Loja ='" & Trim(wLoja) & "' and PR_Situacao <> 'E'" _
            & " order by PR_CodigoFornecedor,PR_Descricao "

    Else
      txtPesquisar.SelStart = 0
      txtPesquisar.SelLength = Len(txtPesquisar)
      txtPesquisar.SetFocus
      Exit Function
    End If
    
    adoProduto.CursorLocation = adUseClient
    adoProduto.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If adoProduto.EOF Then
       txtPesquisar.SelStart = 0
       txtPesquisar.SelLength = Len(txtPesquisar)
       txtPesquisar.SetFocus
       adoProduto.Close
       Exit Function
    End If
    If Not adoProduto.EOF Then
      
        grdProduto.Redraw = False
        grdProduto.Rows = 1
        
        Do While Not adoProduto.EOF
        
           grdProduto.AddItem adoProduto("PR_Referencia") & Chr(9) _
                & adoProduto("PR_Descricao") & Chr(9) _
                & adoProduto("es_Estoque") & Chr(9) _
                & Format(adoProduto("PR_PrecoVenda1"), "###,###,##0.00") & Chr(9) _
                & Format(adoProduto("pr_customedio1"), "###,###,##0.00") & Chr(9) _
                & Format(adoProduto("PR_ICMSSaida"), "###,###,##0.00") & Chr(9) _
                & Format(adoProduto("PR_ICMPDV"), "###,###,##0.00") & Chr(9) _
                & Format(adoProduto("PR_ICMSEntrada"), "###,###,##0.00") & Chr(9) _
                & Format(adoProduto("Pr_IcmPdvEntrada"), "###,###,##0.00") & Chr(9) _
                & Format(adoProduto("PR_SubstituicaoTributaria"), "###,###,##0.00")
                                              
           adoProduto.MoveNext
        
        Loop
      
        grdProduto.Redraw = True
        grdProduto.SetFocus
       
   End If
   adoProduto.Close
End Function

Private Sub txtSerienf_LostFocus()
    If txtSerienf.Text <> "" Then
        txtSerienf = UCase(txtSerienf)
    End If
End Sub

Private Sub txtValFrete_lostfocus()
    If txtValFrete <> "" Then
        TotaisNota
    End If
End Sub

Private Sub txtValICMSST_KeyPress(KeyAscii As Integer)
    
    KeyAscii = campoNumericoVirgula(KeyAscii)
    
    If KeyAscii = 13 Then
        txtValICMSST_LostFocus
        txtQtde_keypress KeyAscii
    End If
    
End Sub

Private Sub txtValICMSST_LostFocus()
    autoPreencherZero txtValICMSST
End Sub

Private Sub txtValorICMS_KeyPress(KeyAscii As Integer)
    
    KeyAscii = campoNumericoVirgula(KeyAscii)
    
    If KeyAscii = 13 Then
        txtValorICMS_LostFocus
        txtQtde_keypress KeyAscii
    End If
    
End Sub

Private Sub txtValorICMS_LostFocus()
    autoPreencherZero txtValorICMS
End Sub

Public Function Totaliza(colunaI As Integer, Optional ColunaII As Integer)

    Dim i As Integer
    
    Totaliza = 0
    i = 1
    
    Do While i <> grdItens.Rows
        If ColunaII = 0 Then
            Totaliza = Totaliza + CDbl(grdItens.TextMatrix(i, colunaI))
        Else
            Totaliza = Totaliza + (CDbl(grdItens.TextMatrix(i, colunaI)) * CDbl(grdItens.TextMatrix(i, ColunaII)))
        End If
        
        i = i + 1
    Loop
    
End Function



Public Function CriaNota()
    
    ' Campos NF Itens
    Dim ReferenciaAtual As String
    Dim ReferenciaAntiga As String
    Dim Descricao As String
    Dim sequencia As Integer
    Dim quantidade As Integer
    Dim ValorUnitario As Double
    Dim ValorTotalItens As Double
    Dim ValIPI As Double
    Dim ICMS As Double
    Dim Desconto As Double
    Dim valorICMS As Double
    Dim Loja As String
    Dim AliqIPI As Double
    Dim baseICMS As Double
    Dim cfop As String
    Dim CST As String
    Dim valorSub As Double
    Dim subst As Double
    Dim basesub As Double
    Dim NCM As String 'Ainda não está inserindo no NFItens --- 06/08/2015
    Dim adoCriaNota As New ADODB.Recordset
    Dim sql As String
    Dim baseIPI As Double
    
    Data = Format(Date, "yyyy/mm/dd")
    
    i = grdItens.Rows - 1
    
    
        ReferenciaAtual = grdItens.TextMatrix(i, 0)
        Descricao = grdItens.TextMatrix(i, 1)
        sequencia = i
        quantidade = grdItens.TextMatrix(i, 2)
        ValorUnitario = grdItens.TextMatrix(i, 3)
        ValorTotalItens = CDbl(grdItens.TextMatrix(i, 2)) * CDbl(grdItens.TextMatrix(i, 3))
        Desconto = grdItens.TextMatrix(i, 4)
        
        baseICMS = grdItens.TextMatrix(i, 5)
        ICMS = grdItens.TextMatrix(i, 6)
        valorICMS = grdItens.TextMatrix(i, 7)
        
        baseIPI = grdItens.TextMatrix(i, 8)
        AliqIPI = grdItens.TextMatrix(i, 9)
        ValIPI = grdItens.TextMatrix(i, 10)
        
        basesub = grdItens.TextMatrix(i, 11)
        subst = grdItens.TextMatrix(i, 12)
        valorSub = grdItens.TextMatrix(i, 13)
        
        CST = grdItens.TextMatrix(i, 14)
        NCM = grdItens.TextMatrix(i, 15)
        
        ReferenciaAntiga = grdItens.TextMatrix(i, 16)
        
        Loja = wLoja
        'sql = "select * from Produto where pr_referencia = '" & ReferenciaAtual & "'"
        'adoCriaNota.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        criaItens ReferenciaAtual, wPedido, sequencia, quantidade, ValorUnitario, ValorTotalItens, ValIPI, ICMS, Desconto, valorICMS, AliqIPI, baseICMS, CST, NCM, ReferenciaAntiga, Loja, cfop, Data, valorSub, subst, basesub, baseIPI
        cfop = " "
        'adoCriaNota.Close
        
        
      
End Function

Public Function criaItens(referencia As String, pedido As String, sequencia As Integer, quantidade As Integer, ValorUnitario As Double, ValorTotalItens As Double, ValIPI As Double, ICMS As Double, Desconto As Double, valorICMS As Double, AliqIPI As Double, baseICMS As Double, CST As String, NCM As String, ReferenciaAntiga As String, Loja As String, cfop As String, Data As String, valorSub As Double, subst As Double, basesub As Double, baseIPI As Double)

    Dim adoCriaItens As New ADODB.Recordset
    Dim sql As String
    
    sql = "insert into NFItens  (NUMEROPED, DATAEMI,            REFERENCIA, " & _
                                "QTDE,      VLUNIT,             VLTOTITEM, " & _
                                "ICMSAplicado,      ITEM,               VLIPI, " & _
                                "DESCONTO,  VALORICMS," & _
                                "NF,        SERIE,              LOJAORIGEM," & _
                                "ALIQIPI,   TIPONOTA,      " & _
                                "BASEICMS,  SituacaoEnvio,      CFOP, " & _
                                "CSTICMS,   SituacaoProcesso,   dataprocesso, " & _
                                "valorSub,  [sub],              baseSub, " & _
                                "baseIPI)"
                                
    sql = sql & " Values        (" & "'" & pedido & "', '" & Data & "', '" & ReferenciaAntiga & "', " & _
                                 quantidade & ", " & ConverteVirgula(ValorUnitario) & ", " & ConverteVirgula(ValorTotalItens) & ", " & ConverteVirgula(ICMS) & ", " & _
                                 sequencia & ", " & ConverteVirgula(ValIPI) & ", " & ConverteVirgula(Desconto) & ", " & ConverteVirgula(valorICMS) & ", 0, " & _
                                 "'0', " & "'" & Loja & "', " & ConverteVirgula(AliqIPI) & ", 'SA', " & ConverteVirgula(baseICMS) & ", 'A', '" & _
                                 cfop & "'" & ", '" & CST & "', 'A', '" & Data & "'," & ConverteVirgula(valorSub) & ", " & ConverteVirgula(subst) & ", " & ConverteVirgula(basesub) & ", " & ConverteVirgula(baseIPI) & ")"
    
    ADO_Cn_CDLocal.Execute (sql)

    
End Function

Public Function criaCapa(Data As String, qtdItem As Integer)

    Dim sql As String
    Dim adoConsultaCAPA As New ADODB.Recordset
    
    sql = "select count(*) as QTDENF from nfcapa where numeroped = '" & wPedido & "'"
    adoConsultaCAPA.CursorLocation = adUseClient
    adoConsultaCAPA.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If Val(adoConsultaCAPA("QTDENF")) = 0 Then
            
        sql = "insert into nfcapa (numeroPed, dataemi, vlrMercadoria, desconto, lojaOrigem, tipoNota, " & vbNewLine & _
        " totalNota, BaseICMS, AliqICMS, VlrICMS, TotalIPI, QtdItem, valorSub, valoricmsst, sub,baseSub,baseicmsst, pesolq, " & _
        " volume,tipofrete,cliente,vendedor,vendedorlojavenda,outraloja,lojavenda,ecf) " & _
        "values ( " & wPedido & ", '" & Data & "', " & ConverteVirgula(txtValTotalMerc.Text) & ", " & vbNewLine & _
        "" & ConverteVirgula(txtDesconto.Text) & ", " & vbNewLine & _
        "'" & wLoja & "', 'SA'," & ConverteVirgula(txtTotalNota.Text) & ", " & vbNewLine & _
        "" & ConverteVirgula(txtBaseCalcICMS.Text) & ", " & vbNewLine & _
        "" & ConverteVirgula(txtAliquotaICMS.Text) & ", " & ConverteVirgula(txtValICMS.Text) & ", " & vbNewLine & _
        "" & ConverteVirgula(txtTotalIPI.Text) & ", " & qtdItem & ", " & ConverteVirgula(txtValICMSSub.Text) & ", " & ConverteVirgula(txtValICMSSub.Text) & "," & vbNewLine & _
        "" & ConverteVirgula(txtAliquotaICMSST.Text) & ", " & ConverteVirgula(txtBaseCalcICMSST.Text) & ", " & ConverteVirgula(txtBaseCalcICMSST.Text) & "," & vbNewLine & _
        "1,1,1,0,999,999,'" & LojaOrigem & "','" & LojaOrigem & "','1')"
        
        ADO_Cn_CDLocal.Execute (sql)
        
    Else
    
        sql = "update nfcapa set vlrMercadoria = '" & ConverteVirgula(txtValTotalMerc.Text) & _
        "', desconto = '" & ConverteVirgula(txtDesconto.Text) & "', totalNota = '" & ConverteVirgula(txtTotalNota.Text) & _
        "', BaseICMS = '" & ConverteVirgula(txtBaseCalcICMS.Text) & "', " & _
        "AliqICMS = '" & ConverteVirgula(txtAliquotaICMS.Text) & "', VlrICMS = '" & ConverteVirgula(txtValICMS.Text) & _
        "', TotalIPI = '" & ConverteVirgula(txtTotalIPI.Text) & "', QtdItem = '" & qtdItem & "', " & _
        "valorSub = '" & ConverteVirgula(txtValICMSSub.Text) & "', sub = '" & ConverteVirgula(txtAliquotaICMSST.Text) & "', " & _
        "valoricmsst = '" & ConverteVirgula(txtValICMSSub.Text) & _
        "', baseicmsst = '" & ConverteVirgula(txtBaseCalcICMSST.Text) & "' where numeroped = '" & wPedido & "'"
        
        ADO_Cn_CDLocal.Execute (sql)
        
    End If
    
    adoConsultaCAPA.Close
    
End Function

Function ConverteVirgula(ByVal numero As String) As String
 
    Dim ret As String
    Dim Charlido As String
    Dim Maximo As Long
    Dim i As Long
    
    ret = "0"
    numero = IIf(IsNull(numero), 0, numero)
    Maximo = Len(numero)
    
    For i = 1 To Maximo
        Charlido = Mid(numero, i, 1)
        
        
        If IsNumeric(Charlido) Then
            ret = ret & Charlido
        ElseIf Charlido = "," And InStr(ret, ".") = 0 Then
            ret = ret & "."
        End If
    Next
    
    ConverteVirgula = ret
 
End Function

Public Function ExcluiNF()
    
    sql = "delete from nfitens where numeroped = '" & wPedido & "'"
    ADO_Cn_CDLocal.Execute (sql)
    
    sql = "delete from nfcapa where numeroped = '" & wPedido & "'"
    ADO_Cn_CDLocal.Execute (sql)
    
    sql = "delete from carimboNotaFiscal where cnf_numeroped = '" & wPedido & "'"
    ADO_Cn_CDLocal.Execute (sql)
    
    sql = "delete from movimentoCaixa where mc_pedido = '" & wPedido & "'"
    ADO_Cn_CDLocal.Execute (sql)
    
    limpaForm
    
    wPedido = PegaNumPedido
    
End Function


Public Function validaImposto()

    If txtValICMSST.Enabled And txtValICMSST.Text = "0,00" Then
        MsgBox "Inserir Valor do ICMS ST.", vbCritical + vbOKOnly, "Atenção"
        validaImposto = True
        Exit Function
    End If
    
    If txtAliquotaICMSST.Enabled And txtAliquotaICMSST.Text = "0,00" Then
        MsgBox "Inserir a Alíquota ICMS ST.", vbCritical + vbOKOnly, "Atenção"
        validaImposto = True
        Exit Function
    End If
    
    If txtBaseAliqICMSST.Enabled And txtBaseAliqICMSST.Text = "0,00" Then
        MsgBox "Inserir a Base ICMS ST.", vbCritical + vbOKOnly, "Atenção"
        validaImposto = True
        Exit Function
    End If
    
    If txtValorICMS.Enabled And txtValorICMS.Text = "0,00" Then
        MsgBox "Inserir o Valor do ICMS.", vbCritical + vbOKOnly, "Atenção"
        validaImposto = True
        Exit Function
    End If
    
    If txtAliquotaICMS.Enabled And txtAliquotaICMS.Text = "0,00" Then
        MsgBox "Inserir o Valor da Alíquota do ICMS", vbCritical + vbOKOnly, "Atenção"
        validaImposto = True
        Exit Function
    End If
    
    If txtBaseICMS.Enabled And txtBaseICMS.Text = "0,00" Then
        MsgBox "Inserir Base ICMS.", vbCritical + vbOKOnly, "Atenção"
        validaImposto = True
        Exit Function
    End If
    
    If txtValorIPI.Enabled And txtValorIPI.Text = "0,00" And txtBaseIPI.Text > 0 Then
        MsgBox "Inserir o Valor do IPI.", vbCritical + vbOKOnly, "Atenção"
        validaImposto = True
        Exit Function
    End If
    
    If txtAliquotaIPI.Enabled And txtAliquotaIPI.Text = "0,00" And txtBaseIPI.Text > 0 Then
        MsgBox "Inserir Alíquota IPI", vbCritical + vbOKOnly, "Atenção"
        validaImposto = True
        Exit Function
    End If
    
'    If txtBaseIPI.Enabled And txtBaseIPI.Text = "0,00" Then
'        MsgBox "Inserir Base IPI", vbCritical + vbOKOnly, "Atenção"
'        validaImposto = True
'        Exit Function
'    End If
    
End Function

Private Function checarCampoVazio()
    If txtQtde.Enabled And Val(txtQtde.Text) < 1 Then
        MsgBox "Quantidade Inválida!", vbCritical + vbOKOnly, "Atenção"
        checarCampoVazio = True
        Exit Function
    End If
    
    If txtPreco.Enabled And txtPreco.Text = "0,00" Then
        MsgBox "Inserir Preço do Item", vbCritical + vbOKOnly, "Atenção"
        checarCampoVazio = True
        Exit Function
    End If
End Function

Private Sub txtValorIPI_KeyPress(KeyAscii As Integer)
        
    KeyAscii = campoNumericoVirgula(KeyAscii)
    
    If KeyAscii = 13 Then
        txtValorIPI_LostFocus
        txtQtde_keypress KeyAscii
    End If
    
End Sub

Private Sub txtValorIPI_LostFocus()
    autoPreencherZero txtValorIPI
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCancelar.SetFocus
End Sub

Private Sub txtSerie_LostFocus()
    If txtNotaFiscal.Text <> "" And txtSerie.Text <> "" Then
        carregaInfoNotaCancelamento txtNotaFiscal.Text, txtSerie.Text
    End If
End Sub

Private Sub cmdCancelar_Click()
    If cancelarNota(txtNotaFiscal.Text, txtSerie.Text) Then
        MsgBox "Nota " & txtNotaFiscal.Text & " cancelada com sucesso!", vbInformation, "Cancelamento"
        frameCancelamento.Visible = False
    End If
End Sub

Private Sub exibirFrameCancelamento()
    frameCancelamento.Visible = True
    txtNotaFiscal.Text = ""
    txtSerie.Text = ""
    lblinfoNota.Caption = ""
    cmdCancelar.Enabled = False
    txtNotaFiscal.SetFocus
End Sub


Private Sub cmdSairCancelamento_Click()
    frameCancelamento.Visible = False
End Sub


Private Sub txtNotaFiscal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtSerie.SetFocus
    KeyAscii = campoNumerico(KeyAscii)
End Sub

Private Sub txtNotaFiscal_LostFocus()
    If txtNotaFiscal.Text <> "" And txtSerie.Text <> "" Then
        carregaInfoNotaCancelamento txtNotaFiscal.Text, txtSerie.Text
    End If
End Sub

Private Sub carregaInfoNotaCancelamento(ByRef nf As String, ByRef serie As String)

    Dim adoCancelamento As New ADODB.Recordset
    Dim sql As String
    
    sql = "select nf as nf, " & vbNewLine & _
          "dataemi as data, " & vbNewLine & _
          "totalnota as total, " & vbNewLine & _
          "codoper as cfop " & vbNewLine & _
          "from nfcapa " & vbNewLine & _
          "where serie = '" & serie & "' and nf = '" & nf & "'" & vbNewLine & _
          "and DATAEMI = '" & Format(Date, "YYYY/MM/DD") & "'"
          
    With adoCancelamento
        .CursorLocation = adUseClient
        .Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
        If Not adoCancelamento.EOF Then
            lblinfoNota.Caption = "Nota Fiscal: " & adoCancelamento("NF") & vbNewLine & _
            "Serie: " & txtSerie.Text & vbNewLine & _
            "Data Emissão: " & adoCancelamento("data") & vbNewLine & _
            "Total Nota: " & adoCancelamento("total") & vbNewLine & _
            "CFOP: " & adoCancelamento("cfop")
            cmdCancelar.Enabled = True
        Else
            cmdCancelar.Enabled = False
            lblinfoNota.Caption = "Nenhuma nota encontrada OU nota já cancelada no sistema"
        End If
        
        .Close
        
    End With
                  
          
End Sub

Private Sub exibirFrameImportar()
    fraImportar.Visible = True
    txtNf.Text = ""
    txtSerienf.Text = ""
    txtCNPJ.Text = ""
    lblDescFornecedor.Caption = ""
    cmdCancelar.Enabled = False
    txtNf.SetFocus
End Sub

Function alteraQuantidade(quantidadeNova As Integer, referencia As String)

    Dim quantidadeAntiga As Integer
    Dim sql As String
    Dim rsItem As New ADODB.Recordset
    
    alteraQuantidade = False
    
   
    If grdItens.TextMatrix(grdItens.row, 17) > quantidadeNova Then
        sql = "update nfitens set qtde = " & quantidadeNova & " where referencia = '" & referencia & "'"
        ADO_Cn_CD.Execute sql
        alteraQuantidade = True
    End If
    
End Function


